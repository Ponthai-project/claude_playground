# `git -C *` 許可札の再設計案（第2版・全面改訂）

作成: 高坂昌信
改訂理由: 第1版は「サブコマンド名」の防御（尾）は正しかったが、「`-C` の引数（パス部分）」がワイルドカードである限り、そこに git の大域オプションを注入できる（胴）という穴を見落としていた。信玄が公式ドキュメント（`https://code.claude.com/docs/en/permissions`）で裏取りした事実に基づき、全面的に設計し直す。

対象: `C:\Users\topge\.claude\settings.json`（**本ファイルへの書き込みは行っていない。読み取りのみ**）
参照した実績データ: `c:\Users\topge\OneDrive\ドキュメント\GitHub\.claude\settings.local.json`（読み取りのみ）
参照した独立防御層: `C:\Users\topge\.claude\hooks\block-dangerous.ps1`（読み取りのみ。冒頭60行を確認）
殿のご裁定: `-C` の対象は ccken 一つに固定せず、**`C:/Users/topge/OneDrive/ドキュメント/GitHub/` 配下を広く許す**。

---

## 0. 前提となる確定事実（推論と明確に区別する）

以下は信玄が公式ドキュメント原文で確認した**確定事実**である。本設計はすべてこれらの上に立つ。

| # | 事実 | 出典 |
|---|---|---|
| F1 | Bash 許可札の照合は**コマンド文字列全体に対するフルマッチ**であり、`*` は空白を含む任意長の文字列（空文字列も可）にマッチする。ワイルドカードは先頭・中間・末尾のどこにでも置ける | 公式ドキュメント原文（信玄の引用） |
| F2 | 末尾が「リテラル＋**空白**＋`*`」の形（例 `ls *`）のとき、そのリテラルの直後は「空白 **または 文字列終端**」のいずれでもよい。ゆえに `ls *` は `ls -la` にも、引数なしの `ls` にも一致し、`lsof` には一致しない。一方 `ls*`（空白なし）は `lsof` にも一致する | 公式ドキュメント原文（信玄の引用） |
| F3 | シェルの区切り演算子（`&& \|\| ; \| \|& &` および改行）はパーミッション評価側が認識しており、複合コマンドは**各サブコマンドごとに独立して**札と照合される。「末尾`*`が`&&`以降まで飲み込んで任意コマンドを実行させる」という懸念は成立しない | 公式ドキュメント原文（信玄の引用） |
| F4 | allow 札は「既知の安全な環境変数」以外の**先頭代入を跨いで一致しない**（例: `GIT_SSH_COMMAND=x git -C ... status` は `Bash(git -C ... status*)` に一致せず ask に落ちる＝安全側）。deny/ask 札は先頭代入を**跨いで一致する**（例: `Bash(rm*)` は `FOO=bar rm -rf tmp/` にも一致する＝安全側） | 公式ドキュメント原文（信玄の引用） |
| F5 | `timeout`/`time`/`nice`/`nohup`/`stdbuf`/`command`/`builtin`/`noglob`、フラグなし `xargs` はラッパーとして剥がされ、内側のコマンドで札照合される | 公式ドキュメント原文（信玄の引用） |
| F6 | settings.json の permissions は **deny → ask → allow** の順に評価され、最初に一致した札が結果を決める。これは permissions 機構そのものの仕様であり、PreToolUse フックの成否とは**独立**に成立する（フックが無くても deny は allow に勝つ） | 公式ドキュメント原文（信玄の引用） |
| F7 | フックが allow 札に確実に打ち勝てるのは **`exit 2`** のときのみ。JSON の `permissionDecision: "deny"` を `exit 0` で返す方式（＝現行 `block-dangerous.ps1` の実装、冒頭60行を確認済み。同スクリプト自身のコメントに「exit code は全経路 0 のまま」と明記）が allow 札に優先するとはドキュメントに明記されていない | `~/.claude/hooks/block-dangerous.ps1` 実機確認 + 公式ドキュメント |
| F8 | 本番の `& "..."` 形でのフック呼び出しにおいて、PowerShell スクリプト内の裸の `exit 2` は終値コード **1** に化ける（信玄が実測）。exit 2 を確実に返すには **`[System.Environment]::Exit(2)`** を使う必要がある | 信玄の実測 |
| F9 | git の大域オプション（`-C`・`-c`・`--exec-path`・`--config-env`・`--git-dir`・`--work-tree` 等）は git 自身の仕様として、サブコマンドより前であれば**互いに任意の順序で並べられる**。`git -C . --exec-path=/tmp/evil status` は `git --exec-path=/tmp/evil -C . status` と等価に解釈される | git 公式マニュアル（`git(1)` SYNOPSIS。本設計の前提として一般に知られた仕様。実機での動作確認はしていない＝**推論**） |
| F10 | `git remote -v` の `-v/--verbose` は git 自身の仕様として「remote と副コマンドの間に置く」ものであり、`git remote -v set-url origin <url>` は `-v` を無視して `set-url` が実行される | git 公式マニュアル（`git-remote(1)`）に基づく**推論**。実機未検証 |

F9・F10 は git 自体の仕様に基づく推論であり、Claude Code のパーミッション機構とは無関係にOS上で成立する動作である。実機検証は6節に記載。

---

## 1. 第1版のどこが「胴」で破れていたか（根本原因の再定式化）

第1版は `branch -D` のような**サブコマンドの末尾に付け足すフラグ**による deny 回避（尾の問題）は正しく塞いだ。しかし見落としていたのは、`-C` の**引数そのもの**がワイルドカードである以上、F1（ワイルドカードは空白を跨ぐ）と F9（git の大域オプションは任意の位置に置ける）が組み合わさり、次の形が常に成立してしまう点である。

```
git -C <安全に見えるパス> --exec-path=/tmp/evil status
```

この文字列は、パス部分をどれだけ工夫して固定しても、`git -C ` の直後から `status`（またはどの許可サブコマンドでも）の直前までを**丸ごと呑み込むワイルドカード1個**さえあれば、無条件に一致してしまう。これは「サブコマンド側の綴り」の問題ではなく「`-C` の引数を表すワイルドカードの構造」そのものの問題であり、**尾をいくら締めても胴が空いていれば無意味**という信玄の指摘は正しい。

### 1-1. 引用符を付けても閉じない理由（実際に手を動かして確認した論理的帰結）

殿の指示どおり「パスの前方をリテラルで固定し、引用符で囲む」形を検討した。

```
Bash(git -C "C:/Users/topge/OneDrive/ドキュメント/GitHub/*" status *)
```

一見、`*` の直後にリテラルの `"` が続くため、閉じ引用符が来るまでの間しかワイルドカードが呑み込めないように見える。だが**このパターンは純粋な文字列照合であり、シェルの引用符構造を理解しない**。攻撃者は次のように、閉じ引用符を「注入した大域オプションの値の側」に用意すればよい。

```
git -C "C:/Users/topge/OneDrive/ドキュメント/GitHub/x" --exec-path="/tmp/evil" status
```

この文字列を先頭から見ると：
- `git -C "C:/Users/topge/OneDrive/ドキュメント/GitHub/` は前方一致リテラルと一致
- ワイルドカードが `x" --exec-path="/tmp/evil` を丸ごと呑み込む（`"` も `*` の一致対象に含まれるため問題なく呑み込める。F1 の「空白を含む任意の文字列」に文字種の制限は無い）
- 残りの `" status` が末尾リテラル `" status` と一致

→ **一致する。** しかも実際にシェルが解釈する引数列は `["-C", "C:/.../GitHub/x", "--exec-path=/tmp/evil", "status"]` であり、`--exec-path` は git 自身にとって正規の大域オプションとして機能する（`-C` の引数の「中に」文字列として埋め込まれているわけではなく、シェルの引用符閉じにより独立した引数になっている）。**つまり、引用符を付けても防げない。**

逆に引用符を付けない形（`Bash(git -C C:/Users/topge/OneDrive/ドキュメント/GitHub/* status *)`）も、境界となるリテラルが単なる空白（` status`）であるだけで、同型の攻撃（`git -C C:/.../GitHub/x --exec-path=/tmp/evil status`）がそのまま成立する。**引用符の有無は本質を変えない。**

### 1-2. 結論（正直な申告）

> **`-C <wildcard> <subcommand>` という形の allow 札は、引用符の有無にかかわらず、大域オプション注入を許可パターンの文法だけでは構造的に防げない。** これは綴りの巧拙の問題ではなく、「ワイルドカードは境界の外側にある文字列との区別がつかない」という glob 照合の性質そのものに起因する。

この穴は **allow 側の文法では閉じられない**。したがって本設計では、**deny 札（F6 により allow に無条件優先する）と、フック（`exit 2` 化後）の二層で、この穴を「注入されうる危険な大域オプション」を名指しで塞ぐ**方式を採る（3節）。これは「機能すれば大きな利得、機能しなくても（allow 側の脆弱性が）今より悪化するわけではない」非対称に有利な設計であり、殿の指示の趣旨と一致する。

**なお、この二層でも防げない残存経路が存在する。** 未知の・将来の git 大域オプイション、またはここに列挙し忘れた大域オプションを使われた場合は、deny 側の列挙が漏れている限り通ってしまう。この列挙型の防御は原理的に「知っているものしか塞げない」という限界を持つことを正直に申告する（8節）。

---

## 2. 撤去する札

`"Bash(git -C *)"`（settings.json 111行目）を撤去する。理由は1節のとおりで変わらない。

---

## 3. 新しい allow 札（そのまま貼れる JSON）

設計方針：
- **`-C` の対象パスを `C:/Users/topge/OneDrive/ドキュメント/GitHub/` 配下に固定**する（殿のご裁定どおり、GitHub フォルダ全体を対象にする。ただし1節の理由により、これは「注入を防ぐ」ためではなく「無関係なディレクトリでの `-C` 濫用を減らす」ための限定であり、注入対策そのものは4節の deny が担う）。
- サブコマンドは**空白＋末尾`*`のF2境界**を用いて、`diff`→`difftool`、`show`→`show-ref`等の**兄弟サブコマンドへの巻き込みを排除**する。
- `remote -v` は**完全一致**（末尾ワイルドカード無し）とし、`remote -v set-url ...`（F10）を allow 側で構造的に排除する。
- 殿の実績（`settings.local.json`）にある**引用符あり・引用符なしの両綴り**を尊重し、各サブコマンドについて両形を用意する。
- F2 の「空白+`*`は文字列終端も許容する」という挙動を**推論ではなく信玄が引用した公式文言そのもの**として採用しつつ、実績上とくに重要な「引数なし呼び出し」（`status`／`push` の bare 形）は**保険として完全一致の別札も重ねて用意**する（万一 F2 の終端許容がこの入れ子パターンで期待通り働かなかった場合の保険。6節参照）。

```json
"Bash(git -C \"C:/Users/topge/OneDrive/ドキュメント/GitHub/*\" status *)",
"Bash(git -C C:/Users/topge/OneDrive/ドキュメント/GitHub/* status *)",
"Bash(git -C \"C:/Users/topge/OneDrive/ドキュメント/GitHub/*\" status)",
"Bash(git -C C:/Users/topge/OneDrive/ドキュメント/GitHub/* status)",
"Bash(git -C \"C:/Users/topge/OneDrive/ドキュメント/GitHub/*\" log *)",
"Bash(git -C C:/Users/topge/OneDrive/ドキュメント/GitHub/* log *)",
"Bash(git -C \"C:/Users/topge/OneDrive/ドキュメント/GitHub/*\" diff *)",
"Bash(git -C C:/Users/topge/OneDrive/ドキュメント/GitHub/* diff *)",
"Bash(git -C \"C:/Users/topge/OneDrive/ドキュメント/GitHub/*\" show *)",
"Bash(git -C C:/Users/topge/OneDrive/ドキュメント/GitHub/* show *)",
"Bash(git -C \"C:/Users/topge/OneDrive/ドキュメント/GitHub/*\" remote -v)",
"Bash(git -C C:/Users/topge/OneDrive/ドキュメント/GitHub/* remote -v)",
"Bash(git -C \"C:/Users/topge/OneDrive/ドキュメント/GitHub/*\" fetch *)",
"Bash(git -C C:/Users/topge/OneDrive/ドキュメント/GitHub/* fetch *)",
"Bash(git -C \"C:/Users/topge/OneDrive/ドキュメント/GitHub/*\" branch *)",
"Bash(git -C C:/Users/topge/OneDrive/ドキュメント/GitHub/* branch *)",
"Bash(git -C \"C:/Users/topge/OneDrive/ドキュメント/GitHub/*\" branch)",
"Bash(git -C C:/Users/topge/OneDrive/ドキュメント/GitHub/* branch)",
"Bash(git -C \"C:/Users/topge/OneDrive/ドキュメント/GitHub/*\" add *)",
"Bash(git -C C:/Users/topge/OneDrive/ドキュメント/GitHub/* add *)",
"Bash(git -C \"C:/Users/topge/OneDrive/ドキュメント/GitHub/*\" commit -m*)",
"Bash(git -C C:/Users/topge/OneDrive/ドキュメント/GitHub/* commit -m*)",
"Bash(git -C \"C:/Users/topge/OneDrive/ドキュメント/GitHub/*\" commit -am*)",
"Bash(git -C C:/Users/topge/OneDrive/ドキュメント/GitHub/* commit -am*)",
"Bash(git -C \"C:/Users/topge/OneDrive/ドキュメント/GitHub/*\" pull *)",
"Bash(git -C C:/Users/topge/OneDrive/ドキュメント/GitHub/* pull *)",
"Bash(git -C \"C:/Users/topge/OneDrive/ドキュメント/GitHub/*\" push *)",
"Bash(git -C C:/Users/topge/OneDrive/ドキュメント/GitHub/* push *)",
"Bash(git -C \"C:/Users/topge/OneDrive/ドキュメント/GitHub/*\" push)",
"Bash(git -C C:/Users/topge/OneDrive/ドキュメント/GitHub/* push)"
```

**28札**（12種のサブコマンド操作 × 引用符あり/なし の2系統。加えて `status`／`push` の bare 完全一致保険2種 × 2系統 = 4札）。

### 3-1. 各札グループの説明（1行ずつ）

| グループ | 何を許すか | なぜこの綴りか（空白の有無） | この札で通ってしまう危険な形 |
|---|---|---|---|
| `status *` / `status`(完全一致) | GitHub配下での `git status`（引数付き・無し双方） | `status *` は F2 の境界規則で `status` 単体にも一致するはずだが未検証のため完全一致札を保険で重ねた。空白は無害操作ゆえ厳密さより実績網羅を優先 | `-C` 引数部への大域オプション注入（1節）。deny側で対処（4節） |
| `log *` | `git log`（`--oneline -N` 等） | 同上。`log` に兄弟サブコマンドは無いため空白境界は厳密には不要だが統一のため付与 | 同上 |
| `diff *` | `git diff` | **空白必須。** `diff*`（空白無し）だと `difftool`（設定不要の直接任意コマンド実行）まで一致してしまうため、境界のために空白を入れた | `-C` 引数部への大域オプション注入。deny側で対処 |
| `show *` | `git show` | **空白必須。** `show*` は `show-branch`/`show-index`/`show-ref` を巻き込むため空白で除外 | 同上 |
| `remote -v`（完全一致・末尾`*`無し） | `git remote -v`（現在のリモート一覧の閲覧のみ） | **末尾ワイルドカードを一切付けない完全一致**。F10 により `-v` は「remoteと副コマンドの間の verbose フラグ」であり、`remote -v set-url origin <url>` のように副コマンドを後続させられる。完全一致にすることで、`-v` の後に何かを続けること自体を allow 側で構造的に禁じた | `-C` 引数部への注入（ただし本札は完全一致のため、注入された文字列の後に「ちょうど `remote -v` で終わる」という極めて限定された形でなければ一致しない。ここは相対的に安全側） |
| `fetch *` | `git fetch`（`--all`、`origin` 等） | 空白境界。`fetch` に危険な兄弟サブコマンド名は無い | **`--upload-pack=<cmd>` によるローカル実行や、任意URLへの fetch（`curl`/`wget` 禁止の抜け道になりうる）は本札だけでは排除できない。deny側で `--upload-pack` を明示的に塞ぐ（4節）。任意URL fetch自体の是非は5節・8節で申告** |
| `branch *` / `branch`(完全一致) | `git branch`（一覧・`-a`/`-v`/`-r`/`--show-current`/`--merged`/新規作成等、幅広く） | 第1版は `branch` を個別列挙して守ったが、`-D`/`-M` 等の破壊操作は**4節の deny 鏡写しで直接遮断**する方式に切り替えたため、allow 側は広く取れる（deny が先に評価されるため F6 により安全）。空白は境界のためだが `branch` に危険な兄弟サブコマンドは無い | `branch -D`/`-M` は deny 鏡写しに依存。deny の綴りに漏れがあれば通る（4節参照） |
| `add *` | `git add <path>` | 既存の非`-C`形 `Bash(git add *)`(103行)と同水準 | `-C` 引数部への注入 |
| `commit -m*` | `git commit -m "..."` | メッセージが可変長のため末尾ワイルドカード必須。`-m` の直後は自由文字列であり兄弟サブコマンド問題は無い | **メッセージ末尾に `--no-verify` を追記する回避が理論上可能**（非`-C`の既存 `Bash(git commit -m*)` にも既にある古い不備。4節で deny 鏡写しを追加） |
| `commit -am*` | `git commit -am "..."` | 同上 | 同上 |
| `pull *` | `git pull`（`--rebase origin main` 等） | 既存の非`-C`広域許可 `Bash(git pull*)`(108行) と同水準に揃えた（第1版は実績限定の完全一致にしていたが、殿の「広く許したい」というご裁定と、deny鏡写しによる防御強化を踏まえ広げた） | `-C` 引数部への注入。将来 `pull` が持ちうる危険フラグ（強制系は基本無い）は現状無し |
| `push *` / `push`(完全一致) | `git push`（`origin main` 等、および引数無し） | 同上。既存の非`-C`広域許可 `Bash(git push*)`(14行)と同水準 | `push -f`/`--force` は deny 鏡写しで遮断（4節）。**deny鏡写しの綴りに漏れがあれば、ここが force push の抜け道になる**ため4節の deny が本札の実効的な安全弁である |

---

## 4. deny 札の鏡写し（新規追加。命じられた8種＋大域オプション注入対策）

F6（deny は常に allow に優先し、これは permissions 機構そのものの保証でありフックの成否に依存しない）に基づき、以下を `permissions.deny` に追加する。**「機能すれば大きな利得、機能しなくても現状より悪化しない」非対称な賭けとして、命じられた8種に加え、1節で明らかになった「胴」の穴（大域オプション注入）を塞ぐ札も併せて追加する。**

### 4-1. 命じられた8種の鏡写し

```json
"Bash(git -C * push* -f*)",
"Bash(git -C * push* --force*)",
"Bash(git -C * reset* --hard*)",
"Bash(git -C * clean* -f*)",
"Bash(git -C * branch* -D*)",
"Bash(git -C * checkout* -- *)",
"Bash(git -C * checkout* -f*)",
"Bash(git -C * filter-branch*)",
"Bash(git -C * rm *)",
"Bash(git -C * mv *)",
"Bash(git -C * update-ref*)",
"Bash(git -C * config *)",
"Bash(git -C * difftool*)",
"Bash(git -C * remote * set-url*)"
```

**綴りの要点**：`push* -f*`のように**サブコマンド語のあとに空白を明示**したのは、`-f` や `--force` が「独立した引数トークン」として現れる場合だけを狙うためである（例：`push origin main -f` のように、危険フラグが `-C` の直後ではなく後方に付いても検知できるよう、サブコマンドと危険フラグの間に任意長ワイルドカードを許した。F1 によりこれは有効）。`remote * set-url*` は F10 の `remote -v set-url` バイパスと素の `remote set-url` の両方を、中間ワイルドカードで一括して拾う。

### 4-2. 追加提案（命じられてはいないが、1節の「胴」を塞ぐために不可欠と判断したもの）

```json
"Bash(git -C * --exec-path*)",
"Bash(git -C * --exec=*)",
"Bash(git -C * -c *)",
"Bash(git -C * --config-env*)",
"Bash(git -C * --git-dir*)",
"Bash(git -C * --work-tree*)",
"Bash(git -C * --upload-pack*)",
"Bash(git -C * --receive-pack*)",
"Bash(git -C * --namespace*)",
"Bash(git -C * mergetool*)",
"Bash(git -C * branch* -M*)",
"Bash(git -C * commit* --no-verify*)",
"Bash(git -C * commit* --no-gpg-sign*)",
"Bash(git -C * push* --no-verify*)",
"Bash(git -C * rebase* --exec*)"
```

これらは1節で明らかになった「`-C` の引数ワイルドカードに大域オプションを混ぜて注入する」経路を、**allow 側の文法では閉じられない**以上、deny 側で名指しに塞ぐものである。`-c *`（任意の `-c key=value`）を丸ごと deny にしたのは、`-c core.pager=` `-c alias.x=` 等の危険な組み合わせを個別に網羅するより、`-c` 自体を GitHub 配下の `-C` 呼び出しでは原則使わせない方が堅牢と判断したため（5節に影響を明記）。

**正直な限界**：この列挙は「信玄が把握している git の危険な大域オプション」に基づく。**未知・将来の大域オプション、または列挙漏れのオプションは防げない。** これは推論ではなく設計上の構造的限界として申告する（8節）。

### 4-3. false positive（誤検知）の注意点

`rm *`／`mv *`（前後に空白必須）は、**コミットメッセージや引数文字列の中に単語として `rm`／`mv` が現れた場合にも誤って一致しうる**（例：`commit -m "please rm old files"` は `git -C * rm *` の対象ではなく `commit -m*` に一致するため実害は無いが、`git -C * add "notes about rm command"` のように add の引数中に単独の ` rm ` が現れると deny が誤発火し得る）。deny は unconditional に allow へ優先するため、誤発火時は**確認プロンプトではなく完全ブロック**になる。これは「取りこぼしゼロ」を優先した意図的なトレードオフであり、頻発するようであれば境界をより厳密な語境界（サブコマンド位置限定）に絞り直す再設計が必要になる。現時点ではこの誤検知率は低いと判断し許容した（推論。実測なし）。

---

## 5. 殿の実績が全て通るかの確認（一件ずつ）

`settings.local.json`（プロジェクト側実績）に記録された `git -C` 付きコマンド11件を、新設計の allow 札と照合した。

| # | 実績（settings.local.json） | 一致する新allow札 | 判定 |
|---|---|---|---|
| 1 | `git -C "...ccken" log --oneline -10` | `"...GitHub/*" log *`（引用符あり） | **PASS** |
| 2 | `git -C C:/.../ccken commit -m ' *`（unquoted） | `git -C C:/.../GitHub/* commit -m*`（引用符なし） | **PASS** |
| 3 | `git -C "...ccken" pull --rebase origin main` | `"...GitHub/*" pull *` | **PASS** |
| 4 | `git -C "...ccken" push origin main` | `"...GitHub/*" push *` | **PASS** |
| 5 | `git -C "...ccken" log --oneline -3` | `"...GitHub/*" log *` | **PASS** |
| 6 | `git -C "...ccken" status`（引数無し） | `"...GitHub/*" status`（完全一致・保険札） | **PASS**（F2の境界規則が本パターンで機能すれば `status *` でも一致するが、保険札で担保） |
| 7 | `git -C "...ccken" status --short` | `"...GitHub/*" status *` | **PASS** |
| 8 | `git -C "...ccken" add 参考資料/` | `"...GitHub/*" add *` | **PASS** |
| 9 | `git -C "...ccken" push`（引数無し） | `"...GitHub/*" push`（完全一致・保険札） | **PASS**（同上、保険札で担保） |
| 10 | `git -C "...ccken" add "参考資料/ハーネスエンジニアリング総まとめ.md"` | `"...GitHub/*" add *` | **PASS** |
| 11 | `git -C "...ccken" add "参考資料/ハーネスエンジニアリング ビジュアル解説.html"` | `"...GitHub/*" add *` | **PASS** |

**11件全件 PASS。** うち2件（#6, #9）は F2 の「空白+末尾`*`は文字列終端も許容する」という挙動に本来は乗る想定だが、**この挙動自体が実機未検証**（信玄の引用した公式文言からの適用であり、当該入れ子パターンでの動作は未確認）であるため、独立した完全一致札を保険として重ねてあり、F2 の挙動如何によらず PASS することを保証している。

### 5-1. 非`-C`では許可されているのに `-C` 付きだと ask になる組合せ（全数列挙）

第1版は `branch`／`pull`／`push` を実績限定の完全一致にしたため多くの gap を生んでいたが、本設計は非`-C`広域許可（`branch*`／`pull*`／`push*`）と同水準まで `-C` 側も広げたため、**gap は大幅に縮小した**。それでも残るものを列挙する。

| 非`-C`で許可されている形 | `-C` 付きでの扱い | 理由 |
|---|---|---|
| `git branch --show-current` | ✅ 新設計でカバー（`branch *`） | — |
| `git branch -d <name>` | ✅ 新設計でカバー（`branch *`。`-d`小文字は安全な削除で `-D`のみdeny対象） | — |
| `git branch <new-branch-name>`（新規作成） | ✅ 新設計でカバー（`branch *`） | — |
| `git branch --merged` | ✅ 新設計でカバー（`branch *`） | — |
| `git merge*`（非`-C`, 13行目は該当なし。実際は15行目 `Bash(git merge*)`） | ❌ `-C` 付きは ask のまま（allow未追加） | 殿の実績に `-C merge` の使用例が無いため保守的に見送った。追加要否は6節で再度諮る |
| `git checkout ...`（`--`を伴わない通常のブランチ切替） | ❌ 元々**非`-C`でも許可されていない**（settings.json に `Bash(git checkout*)` 自体が無い） | gapではなく元からの仕様。`-C`有無で差は無い |
| `git rebase`／`git stash`／`git tag`／`git cherry-pick`／`git revert` | ❌ 元々非`-C`でも許可されていない | 同上、gapではない |

**結論：真の gap（＝非`-C`では通るのに`-C`だと ask になる）は「`git -C <dir> merge ...`」の1点のみに収束した。** 殿がGitHub配下のリポジトリで頻繁に merge を使われるなら、以下を追加することを提案する（未採用。殿・信玄のご判断を仰ぐ）。

```json
"Bash(git -C \"C:/Users/topge/OneDrive/ドキュメント/GitHub/*\" merge *)",
"Bash(git -C C:/Users/topge/OneDrive/ドキュメント/GitHub/* merge *)"
```

追加する場合は、対応する deny 鏡写し `"Bash(git -C * mergetool*)"`（4-2節に記載済み）とセットで運用すること。

---

## 6. できなくなること（訂正版：フック＋確認プロンプトの二段）

前版は「hookが唯一の防波堤」と記載した箇所があったが誤りである。F6により、**allow札に一致しなければ既定の ask（確認プロンプト）が独立した第2層として働く**。正しい認識は次の通り。

- `git -C <dir> reset --hard`／`clean -f`／`checkout --`／`push -f`／`--force`／`branch -D` 等は、**(1) settings.json の deny 鏡写し（4節、config層で無条件にallowへ優先）→ (2) 万一 deny 側の綴りに漏れがあれば、allowに一致しないため既定の ask → (3) さらにフック（`block-dangerous.ps1`、修正・exit 2化後）**という**三段**で守られる。前版の「hookのみ」という記述は誤りであり、正しくは「**deny札＋確認プロンプト＋フックの多段**」である。
- `git -C <dir> merge`／`rebase`／`stash`／`tag` 等、allow に無い操作は今後 ask になる（5-1節の gap 表のとおり、merge以外は元々非対象）。
- `git -C <dir> -c ...`（任意の git config 上書き）は4-2節の deny により**用途を問わず一律ブロック**される。殿がGitHub配下で正当な `-c` 利用（例：一時的な pager 変更）をされる場合は今後 deny で止まる。用途が判明次第、個別の安全な `-c` キーのみ許可する例外札の追加を検討されたい。

---

## 7. フック改修との同期（現況の反映）

引継ぎ書（`引継ぎ_パーミッション是正_2026-08-25.md`）および山県の改修により、以下は**現況ではフックの網に入っている**（前版で「hookの網の外」としていた記述を撤回する）。

- `branch -D` / `filter-branch` / `rm` / `mv` / `checkout -f` / `update-ref` / `config`
- `difftool`/`mergetool` の `-x`/`--extcmd=`
- `fetch`/`clone`/`pull`/`push` の `--upload-pack=`/`--receive-pack=`/`--exec=`
- `--exec-path=`、`--config-env=`
- `remote set-url`/`add`/`remove`/`rename`
- `--no-verify` の長綴り、`commit -nm`

ただし7節の**フックが allow に優先する保証は `exit 2`（正確には `[System.Environment]::Exit(2)`。裸の `exit 2` は本番の `&` 呼び出し形で終了コード1に化けることを信玄が実測済み＝F8）に依存する**。現行フックは `exit 0` 固定でJSON `deny` を返す方式であり、**F6（settings.jsonのdeny→ask→allowという配置レベルの保証）とは別物**である。したがって、**本設計（4節のdeny札）は、フックの exit 2 化が完了していなくても単独で有効**である（deny札はconfigレベルの機構でありフックに依存しない）。フック側の exit 2 化は独立した並行作業として進めることを推奨する。

---

## 8. 未確定・実機検証を要する点（推論と確認済みを明確に区別）

| # | 項目 | 状態 |
|---|---|---|
| 1 | F1〜F9のうちF1・F2・F3・F4・F5・F6・F7は**公式ドキュメント原文からの確認済み事実**（信玄の引用）。F9（git大域オプションの順序自由性）とF10（`remote -v`の位置仕様）は**git公式マニュアルに基づく推論**であり、実機（実際にこのコマンドをgitに投げて動作を確認）はしていない | **要実機検証** |
| 2 | 3節の「`status *`のような空白+末尾`*`が、パスワイルドカードを挟んだ入れ子構造でも文字列終端を正しく許容するか」（F2の適用範囲の外挿） | **要実機検証**（5節の保険札で影響は限定済み） |
| 3 | 4節の deny 札が実際に allow 札より先に評価され、`git -C "...GitHub/x" --exec-path=/tmp/evil status` のような注入コマンドを実際に拒否するか | **要実機検証（最優先）** |
| 4 | 4-3節で申告した `rm *`/`mv *` の誤検知が実運用でどの程度の頻度で起きるか | **未検証・運用しながら様子見が必要** |
| 5 | 1-1節で示した引用符バイパスの攻撃文字列（`git -C "...GitHub/x" --exec-path="/tmp/evil" status`）が、実際にこの環境（PowerShell経由のBashツール）でシェル引用符処理を経て意図通りgitに渡るか | **要実機検証**（Bashツールの引用符解釈がPOSIX的かWindows的かで挙動が変わりうる） |
| 6 | 5節「merge」追加提案は未採用（殿の承認待ち） | 承認保留 |

**実機検証は信玄が自ら行うこと**（グローバル運用ルール9）。とくに #3（deny が実際に注入コマンドを止めるか）は、この設計案の正否を左右する最重要項目であり、`.claude`への適用前に必ず確認されたい。

---

## 9. 緊急・スコープ外だが看過できない発見（今回の任務対象外。至急共有）

本任務中に settings.json を読み込んだ際、**`-C` とは無関係に、現行の非`-C`許可札そのものにF1・F2と同型の穴が複数見つかった**。本ファイルの編集対象外（settings.jsonへの書き込みは禁止されている）だが、実害度が高いため報告する。**いずれも実機未検証・文字列照合ロジックからの論理的導出（推論）である。**

1. **`"Bash(git diff*)"`（36行目、空白なし）は `git difftool` にも一致する。** `git difftool -x "<任意コマンド>" HEAD~1 HEAD` は設定不要の直接任意コマンド実行であり、**`-C` を使わずとも今すぐ通ってしまう**可能性がある。
2. **`"Bash(git merge*)"`（15行目、空白なし）は `git mergetool` にも一致する。** 同様に `-x` 経由の任意コマンド実行が**`-C` なしで**成立しうる。
3. **`"Bash(git remote -v*)"`（110行目、空白なし）は `git remote -v set-url origin <url>` にも一致する。** F10のとおり `-v` は verbose フラグであり、`-C` を使わずとも**リモートの差し替えが今すぐ通ってしまう**可能性がある。

いずれも「サブコマンド名 + 空白なし `*`」という綴りが、境界の無い接頭辞一致を生んでいることが原因であり、本設計の3節で採用した「空白+末尾`*`」の境界パターンに直せば解消できる（`"Bash(git diff *)"` 等）。**`-C` 問題とは独立した、既存設定の即時是正が望ましい事項**として、信玄より山県昌信への申し送りを推奨する。

---

## 10. 変更点まとめ（第1版→第2版）

| 項目 | 第1版 | 第2版 |
|---|---|---|
| allow 札数 | 17 | 28（うち4札は保険） |
| `-C` 対象パス | 無制限（`*`のみ） | `C:/Users/topge/OneDrive/ドキュメント/GitHub/` 配下に限定（ただし1節の理由により注入対策としては無効、濫用範囲の限定効果のみ） |
| `branch` | 個別完全一致で列挙（`-D`等を機械的に除外） | 広く許可 + deny鏡写し（`-D`/`-M`）で遮断する方式に転換 |
| `remote -v` | 末尾ワイルドカード付き（`remote -v*`） | **完全一致**に変更（F10の`-v set-url`バイパスを閉じるため） |
| `pull`/`push` | 実績限定の完全一致（狭い） | 非`-C`広域許可と同水準まで拡大（deny鏡写しで補償） |
| deny 鏡写し | 無し（存在しない設計課題として言及のみ） | 8命じられた種＋大域オプション注入対策など計29札を新設 |
| 「できなくなること」の記述 | 「hookが唯一の防波堤」（誤り） | 「deny札＋確認プロンプト＋フックの多段」に訂正 |
| フックとの関係の記述 | 「JSONのdenyがallowに優先する」という誤前提が3箇所 | F6（config層のdeny>allowはフックと独立）とF7/F8（フック自体がallowに勝つには`[System.Environment]::Exit(2)`が必須）に基づき全面訂正 |
| `-C`引数への大域オプション注入 | **未検討（見落とし）** | 1節で根本原因として特定。allow側では防げないことを明示し、deny側（4-2節）で対策 |
| スコープ外の付随発見 | 無し | 9節：`diff*`/`merge*`/`remote -v*`（いずれも非`-C`・空白なし）の既存の穴を発見・報告 |

---

## 附則：ファイルパス一覧

- 読み取ったファイル（変更なし）：
  - `C:\Users\topge\.claude\settings.json`
  - `c:\Users\topge\OneDrive\ドキュメント\GitHub\.claude\settings.local.json`
  - `C:\Users\topge\.claude\hooks\block-dangerous.ps1`（冒頭60行）
  - `c:\Users\topge\OneDrive\ドキュメント\GitHub\claude_playground\work\hook-utf8-fix\引継ぎ_パーミッション是正_2026-08-25.md`
- 改訂した成果物：
  - `c:\Users\topge\OneDrive\ドキュメント\GitHub\claude_playground\work\hook-utf8-fix\git-allow-rules-plan.md`（本ファイル）
