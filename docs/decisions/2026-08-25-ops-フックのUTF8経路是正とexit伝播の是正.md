---
date: 2026-08-25
type: ops
project: claude_playground
title: フックのUTF-8経路是正とexit伝播の是正、fail-safe方針の裁定
status: wip
decision: フックの終了処理を全て [System.Environment]::Exit(N) に統一し、fail-safe は ask のまま据え置いたうえで allow 札の絞り込みと同時適用を必須条件とする
tags: [security, permissions, hooks, powershell, encoding, cp932]
---

# フックのUTF-8経路是正とexit伝播の是正、fail-safe方針の裁定

## 1. 依頼

引継ぎ書 `work/hook-utf8-fix/引継ぎ_パーミッション是正_2026-08-25.md` を読み、残件1（修正版フックの検分とBOM合成）から順に続きをやること。作業場所は `claude_playground` 配下。`.claude` 配下は読むだけにし、書き込み（フックの差し替え・settings.json の変更）は必ず承認を取ってから。**前回のように推論で「直ったはず」と判断せず、実機で確かめてから報告すること。**

## 2. 前提・制約

- 本番の PreToolUse フック `~/.claude/hooks/block-dangerous.ps1` は現在も危険コマンドを一切検知できていない（数か月間、常時 fail-safe の ask に倒れ、それが黙殺されていた）。
- 前回の是正は三層構造の真ん中の一層（BOM）しか塞いでおらず、「25/25 全合格」は偽合格だった。ハーネスが PowerShell 同士のパイプで書き手と読み手のエンコーディングを揃えてしまい、本番（書き手＝Node.js の UTF-8）を再現していなかったため。
- 作業ディレクトリのパスに日本語「ドキュメント」を含む。「ト」の3バイト目 `0x88` が直後の `\`(0x5C) を食う CP932 の「ダメ文字」が事故の根。
- subagent は PowerShell ツールを持たず、Bash 経由の間接呼び出しはグローバル規則6が禁じるため、**実機検証は信玄が自ら行う**。

## 3. 検討した選択肢

### 3-1. deny をどう返すか（allow 札に勝つ手段）

| 案 | 長所 | 短所 | 採否 | 却下理由 |
|---|---|---|---|---|
| JSON の `permissionDecision: "deny"` のみ | 実装が単純。既存踏襲 | allow 札に優先すると公式ドキュメントのどこにも書かれていない | 却下 | `Bash(git -C *)` のような広い札に負ける公算が高い |
| `exit 2` を返す | ドキュメントが「allow ルールに優先する」と明記 | **本番の `&` 呼び出し形では終了コードが 1 に化ける（実測）** | 却下 | exit 1 は non-blocking error＝ツールがそのまま進む。deny が素通りする |
| `[System.Environment]::Exit(2)` | `-File` / `-Command '& "x.ps1"'` のいずれでも終了コード2が届く（実測） | プロセス即時終了ゆえ出力の取りこぼしが懸念 | **採用** | - |
| settings.json 側を `-File` 起動形に変更 | フック本体を触らずに済む | `"shell": "powershell"` の展開形を Claude Code が握っており確実性が無い。設定変更はセッション再起動を要する | 却下 | 起動形に依存しない解の方が堅い |

### 3-2. 判定不能（fail-safe）時の振る舞い（殿の裁定事項）

| 案 | 長所 | 短所 | 採否 | 却下理由 |
|---|---|---|---|---|
| A: ask のまま＋allow 札の絞り込みを同時適用必須 | 運用の重さが変わらない。札を絞れば ask が黙殺されても通るのは無害なコマンドだけ | フック単独適用では穴が開いたまま。同時適用の規律が要る | **採用（殿の裁定）** | - |
| B: 判定不能も `Exit(2)` で止める | 最も確実。故障が即座に露見する | フック恒常故障時に Bash/PowerShell が全面停止し、手作業で settings.json を直すまで何も通らない | 却下 | 運用停止のリスクが大きすぎる |
| C: 故障の重さで振り分け（入力不能は止め、評価中例外は ask） | 中庸 | 実装が複雑になり、境界の判断が増える | 却下 | A で穴が閉じるなら複雑さを買う理由がない |

### 3-3. allow 札 `Bash(git -C *)` の置き換え方

| 案 | 長所 | 短所 | 採否 | 却下理由 |
|---|---|---|---|---|
| 撤去する | 最も安全 | `git -C` を使うたび確認プロンプト。実用性が崩壊 | 却下 | 殿が「撤去でも現状維持でもなく限定」と裁定済み |
| 17札に分割（`Bash(git -C * status*)` 形） | サブコマンド単位で絞れる | **中間 `*` が空白を跨ぐため `git -C . --exec-path=/tmp/evil status` に一致し任意実行を自動許可**（公式ドキュメントで確定） | 却下 | 胴に無制限の注入口が開く |
| `-C` の引数を単一リポジトリの固定パスに固める | 注入口が構造的に消滅 | ccken 以外で `git -C` を使うたび札の追加が要る | 却下 | 殿が「GitHub フォルダ配下を広く許したい」と裁定 |
| `-C` 引数の前方をリテラルで固定＋引用符で囲む | 注入を難しくしつつ GitHub 配下を広くカバー | 引用符無し形では防ぎきれない可能性 | **採用（設計中）** | - |

## 4. 決定と決め手

1. **フックの終了処理を全て `[System.Environment]::Exit(N)` に統一する。** 決め手は、本番と同じ `&` 呼び出し形で `exit 2` が終了コード1に化けることを実機で計測し、`[System.Environment]::Exit(2)` なら全起動形で2が届くことを確認したため。
2. **fail-safe は ask のまま据え置き、allow 札の絞り込みとの同時適用を必須条件とする（案A）。** 決め手は、フック恒常故障時に作業が全面停止する事態を避けつつ、札を絞れば穴が閉じるという等価性を優先したため。**片方だけの適用はしない。**
3. **`git config` の書き込み形は ask ではなく deny とする。** 決め手は、`git config core.pager <任意コマンド>` が「設定に書き込むだけで、以後の無害な `git log` が任意コマンドを実行する」時間差の任意実行であり、危険が殿の目に見えにくすぎるため。

## 5. やったこと

- 成果物：`c:\Users\topge\OneDrive\ドキュメント\GitHub\claude_playground\work\hook-utf8-fix\build\block-dangerous.ps1`（BOM付き）、同 `build\test-hook-utf8.ps1`（検証ハーネス）、`build-bom.sh`（BOM合成手順）
- BOM合成は `printf '\357\273\277' > 出力先` → `cat 本文 >> 出力先` の2手。Write ツールは BOM を付けないため。
- 新規ハーネスを作成（旧ハーネスは偽合格を出すため流用せず）。`ProcessStartInfo` で子シェルを起こし `StandardInput.BaseStream` に UTF-8 生バイトで書く。PowerShell の `|` は使わない。
- **75ケース × （シェル2種 × 起動形2種）＝ 300実行。** 起動形の軸は、本件の発見を受けて第2版で追加した。
- フックの検知漏れを塞いだ：`--no-verify` 長綴り、`commit -nm`、`branch -D`、`filter-branch`、`rm`、`mv`、`checkout -f`、`update-ref`、`config` 書込形、`difftool`/`mergetool` の `-x`/`--extcmd=`、`fetch/clone/pull/push` の `--upload-pack=`/`--receive-pack=`/`--exec=`、`--exec-path=`、`--config-env=`、`remote set-url/add/remove/rename`、引用符付き `-C` パス、`-c` 設定キー名の大文字小文字非区別。

## 6. 見直し条件

- Claude Code がフックの起動形（`"shell": "powershell"` の展開）を変更した場合、`[System.Environment]::Exit()` の必要性を再評価する。ただし全起動形で動くため、変更されても壊れない側に倒してある。
- Bash 許可札のワイルドカード意味論（任意位置可・空白を跨ぐ）が変わった場合、札の設計を再検討する。

## 7. 未解決・次の一手

- **適用後の実地確認（次セッション最優先）**：本番で `git -C . reset --hard HEAD` が実際に `Permission denied` で止まることをこの目で見るまで完了としない。`settings.json` はセッション開始時に凍結されるため、今セッションでは札の実効性を確認できない。
- deny 札 `Bash(git -C * rm *)` / `mv *` / `config *` / `-c *` の誤検知の様子見。deny は覆せないため、頻発するならサブコマンド位置限定の綴りに絞り直す。
- 他フックへの水平展開（`check-html` / `inject-context` / `worklog-check` / `test-block-dangerous` の4本）。`inject-context` は現に出力が化けており、出口側の穴の実例。
- 外部送信の出口制限（`WebFetch` のドメイン無制限、`Bash(git push*)` のリモートURL無制限）。

## 試行錯誤ログ（dev のみ）

- **プローブ自身が調査対象の穴に落ちた。** `[System.Environment]::Exit()` が stdout を切り落とさないかを確かめるプローブを日本語コメント付き・BOM無しで書いたところ、stdout が空になった。stderr を見ると `$json` が null で、`$bytes = ...` が「3行目」と報告されていた。日本語コメント3行が CP932 誤読で潰れ、代入行ごと飲み込まれていた。**Exit が悪いのではなく、プローブが何も出力していなかった。** ASCII のみに書き直して再測したところ、130バイトのJSONが完全に届き、終了コードも2であった。
  - 教訓：BOM 無しで日本語を含む `.ps1` は 5.1 で必ず壊れる。使い捨てのプローブでも例外ではない。**検証道具は ASCII のみで書くか、BOM を付けるかのどちらかを徹底する。**
- **ハーネスの軸の見落としは再発した。** 前回は「書き手が誰か」（Node.js か PowerShell か）の軸を見落として偽合格を出した。今回は「起動形」（`-File` か `& "x.ps1"` か）の軸を見落としかけた。第1版は 49/49 全合格だったが、`-File` で回していたため exit 伝播不良を構造的に検出できなかった。
  - 教訓：**ハーネスが全合格したときこそ「本番と違う軸はないか」を疑う。**

## 8. 適用（2026-08-26、殿の承認を得て実施）

- バックアップ：`~/.claude/hooks/block-dangerous.ps1.bak-20260826`（7,222バイト）、`~/.claude/settings.json.bak-20260826`（11,121バイト）
- フック本体を差し替え（build 版とバイト単位で同一を `cmp` で確認、BOM を `od` で確認）
- allow から `Bash(git -C *)` を撤去し、GitHub 配下限定の30札を新設。deny に `git -C *` 形30札を追加
- 既存の穴3件を修繕：`Bash(git diff*)` → `Bash(git diff *)`、`Bash(git merge*)` → `Bash(git merge *)`、`Bash(git remote -v*)` → `Bash(git remote -v)`。deny に `git difftool*` / `git mergetool*` / `git remote * set-url*` を追加
- 適用後、`settings.json` が JSON として健全（allow 147件・deny 160件・`hooks` キー健在）であることを実測確認
- **本番パスのフックに対してハーネスを再実行し、4ターゲットすべて 72/72、RESULT: PASS を確認**
- 本番の監査ログに、差し替え後のフックが日本語の理由文を化けなく記録していることを確認（`git -C . branch -D`、`git -C "My Folder" reset --hard`、`core.hookspath` を現に捕捉）

## memory昇格候補

- 検証ハーネスが全合格したら、まず「本番と異なる軸が残っていないか」を疑う（書き手のエンコーディング、起動形、実行ユーザ等）。前回・今回と2度続けて同じ型の見落としが起きた。 → memory: `harness-suspect-hidden-axis` に昇格
