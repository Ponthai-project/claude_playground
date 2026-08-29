# fault/ 作業記録（山県昌景）

対象：skill `graphreco-html` の検品規約 C1〜C9 に対する故障注入・先行2本。

## 1. base.html（対照群）の是正2箇所

`sample_figures.html` を複製し、以下2箇所のみを変更した。他は一切変更していない（diff で確認済み）。

### 是正1（C7対応）：`<text>` への `class="fig-label"` 付与

対象：散布図（03節）内の `<text>` 6件（元ファイル 643〜648行）。

元の記述（各行とも `class` なし）：
```
<text x="405" y="284" font-size="13" fill="#2A2723" font-weight="bold">コスト→</text>
<text x="8" y="18" font-size="13" fill="#2A2723" font-weight="bold">性能↑</text>
<text x="95" y="252" font-size="13" fill="#2A2723" font-weight="bold">次点モデル</text>
<text x="240" y="115" font-size="14" fill="#C15F3C" font-weight="bold">Opus 5</text>
<text x="300" y="55" font-size="14" fill="#2A2723" font-weight="bold">Fable 5</text>
<text x="252" y="98" font-size="13" fill="#C15F3C" font-weight="bold">←半額</text>
```

是正後（6件それぞれに個別付与。CSS定義 `.fig-label{ font-size:max(12px,0.75rem); filter:none; }` は元から存在していたが
マークアップで一度も使われていなかったため、これで初めて適用される）：
```
<text class="fig-label" x="405" y="284" font-size="13" fill="#2A2723" font-weight="bold">コスト→</text>
<text class="fig-label" x="8" y="18" font-size="13" fill="#2A2723" font-weight="bold">性能↑</text>
<text class="fig-label" x="95" y="252" font-size="13" fill="#2A2723" font-weight="bold">次点モデル</text>
<text class="fig-label" x="240" y="115" font-size="14" fill="#C15F3C" font-weight="bold">Opus 5</text>
<text class="fig-label" x="300" y="55" font-size="14" fill="#2A2723" font-weight="bold">Fable 5</text>
<text class="fig-label" x="252" y="98" font-size="13" fill="#C15F3C" font-weight="bold">←半額</text>
```

（`<g class="fig-label">` で6件をまとめて包む方式は採らなかった。C7の合格条件が
「`<text` の出現数 ＝ `class="fig-label"` の出現数」であるため、まとめて包むと1対6になり不合格になるからである。
詳細は末尾「4. 気づいた規約側の疑義」参照。）

### 是正2（C8(b)対応）：メディアクエリ内の rotate 行を削除

対象：元ファイル 236行（`@media (max-width:700px)` ブロック内）。

元の記述：
```
  @media (max-width:700px){
    .flow-arrow-wrap{ flex-basis:36px; }
    .flow-arrow-wrap svg{ transform:rotate(90deg); }
  }
```

是正後（1行削除のみ）：
```
  @media (max-width:700px){
    .flow-arrow-wrap{ flex-basis:36px; }
  }
```

削除理由：`.flow-arrow-wrap`（元ファイル674行）は `data-fig-type="対比"` の `.fig-block`（04節）の内側にあり、
C8(b)「`.fig-block` 内に `rotate(` が掛かる行が0件」に抵触していたため。

### 是正3（追加・C5対応。信玄の検算による指示）：01節 `.sec-head` 内svgの削除

対象：`fault/base.html` の483行（元ファイル484行相当）。**`sample_figures.html` 側は無変更（触っていない）。**

元の記述（是正1・是正2適用後、是正3適用前のbase.html 483行）：
```
        <svg width="48" height="48" style="margin-left:auto; flex-shrink:0;" aria-hidden="true"><use href="#talking-person"/></svg>
```

是正後：この1行を丸ごと削除（前後の `</div>` と `</div>`〈.sec-head閉じ〉のみ残る）。

削除理由：この見本は「見出しの右端に48pxアイコンを1個貼るだけ」というC5が禁じる形式的寄生パターンと
サイズ・`margin-left:auto`・`.sec-head` 直下という3条件すべてで一致しており、対照群がC5を満たしていなかった
（＝見本に潜んでいた真の違反。C5が一度も当てられていなかったことによる見落とし。規約側の欠陥ではない）。
対照群がC5不合格のままでは、case-a注入によるC5検出可否の判定が成立しないため、信玄の指示により除去した。

### 差分確認

`diff sample_figures.html fault/base.html` で以下3ブロックのみが差分として現れることを確認済み（他の行は完全一致）：
是正1（`<text>` 6件への `class="fig-label"` 付与）／是正2（236行 `rotate(90deg)` 行の削除）／
是正3（483行、01節 `.sec-head` 内svgの削除）。

## 2. case-a.html に仕込んだ改変（C5検査への注入）

`base.html` を複製し、**616行目に1行追加**した。狙う検査：**C5「右上寄生の禁止」**。

（初版では617行目だったが、後述の是正3で01節の既存svgを1行削除した結果、行番号が1つ繰り上がった。）

追加箇所：03節（コスト対性能）の `.sec-head` 内、`</div>` の直後・`.sec-head` の閉じタグ直前。

```html
      <div class="sec-head">
        <span class="sec-num">03</span>
        <div>
          <h2 class="section-title">コスト対性能</h2>
          <p class="sec-sub">同等以上の性能を、より低いコストで。</p>
        </div>
        <svg width="48" height="48" style="margin-left:auto; flex-shrink:0;" aria-hidden="true"><use href="#magnifier"/></svg>
      </div>
```

`diff base.html case-a.html` で上記1行の追加のみであることを確認済み。他は一切変更していない。

### 実測件数（是正3・作り直し後）

- `margin-left:auto` 出現数：base.html＝**0件** ／ case-a.html＝**1件**
- base.html の `<text` 出現数：**6件**（`class="fig-label"` も**6件**で一致。是正1は保たれている）
- base.html の `rotate(` 出現数：**1件**（`.circled::before{ transform:rotate(-1deg); }`。ヒーロー/マーカー装飾用途で
  `.fig-block` の外にあるため、C8(b)の対象外。是正2は保たれている）
- 是正3の前後で他に差分が生じていないこと：`diff sample_figures.html fault/base.html` を実施し、
  是正1・是正2・是正3の3ブロックのみが差分として現れることを確認（後述「差分確認」参照）。

## 3. 作成した3ファイルのパス

- `c:\Users\topge\OneDrive\ドキュメント\GitHub\claude_playground\work\graphreco-tegaki\fault\base.html`
- `c:\Users\topge\OneDrive\ドキュメント\GitHub\claude_playground\work\graphreco-tegaki\fault\case-a.html`
- `c:\Users\topge\OneDrive\ドキュメント\GitHub\claude_playground\work\graphreco-tegaki\fault\NOTES-base.md`（本ファイル）

## 4. 気づいた規約側の疑義

### 疑義A：C7「件数一致」条件は「ラベル層の分離」という条項名の趣旨と食い違う可能性がある

C7の合格条件が「`<text` の出現数 ＝ `class="fig-label"` の出現数」という**単純な個数一致**である場合、
実装者が「揺らぎフィルタの掛かる `<g>` でテキスト層をまとめて包み、その `<g>` 自体に `fig-label` を付す」という、
CSSコメント（68〜69行）が推奨している設計思想（「枠線・背景だけを描く専用レイヤーに揺らぎを掛け、
テキストは通常フローに置いて歪ませない」という分離パターンをテキスト側に応用する書き方）を選んだ場合、
件数が1対Nになり不合格になる。条項名が「ラベル層の分離」を志向しているなら、
「`<text>` 個々にクラスを持つ」ことと「テキストをまとめた層が独立している」ことは別の性質であり、
後者を志向する書き方（`<g class="fig-label"><text>...</text><text>...</text></g>`）を規約が排除してよいのか、
一度確認されたい。今回は規約の文言（個数一致）に厳密に従い、6件それぞれへの個別付与を選んだ。

### 疑義B：C5が検出対象とする「右上寄生」パターンが、清書済みの対照群（base.html）に既に1件存在する

01節（ひとことで言うと）の `.sec-head` には、元ファイル（`sample_figures.html`）の時点から
`<svg width="48" height="48" style="margin-left:auto; flex-shrink:0;" aria-hidden="true"><use href="#talking-person"/></svg>`
（複製後の base.html でも483行目に同一のまま存在）が置かれている。これは今回の指示書が
case-a.html への注入内容として説明した「見出しの右端に48pxアイコンを1個貼るだけ」の形式と**寸分違わず一致**する
（サイズ48px・`margin-left:auto`・`.sec-head` 直下、という3条件がすべて同じ）。

つまり、C5が本当に「サイズ48px・margin-left:auto・sec-head直下のsvg」という形式的特徴だけで検出しているなら、
**対照群であるはずの base.html が既にC5で不合格になる**可能性がある。逆にC5がこの01節の既存パターンを
合格として扱っているなら、それは「アイコンが対話相手（talking-person）を表し、直後のquote-row（会話吹き出し3つ）と
意味的に結びついている」といった文脈判断をC5が行っている（＝単純な形式検出ではない）ことを意味する。

**今回はcase-a.htmlの改変を03節（元々sec-head内にsvgが無い節）に加えることで、01節の既存要素とは独立した
「新規1件の注入」として切り分けた。** ただし、内藤の検査でC5が「不合格件数1件」ではなく「2件」（01節＋03節）
を返してきた場合、それは規約の疑義ではなく01節が既に潜在違反だったことの発覚であり、対照群base.htmlの
前提（「C1〜C9すべて合格する清書版」）が崩れる。この場合は base.html 側の是正範囲を01節にも広げる必要が
出てくるため、殿・内藤への報告時に必ず切り分けること。

**【決着・信玄の検算により確定】** 疑義Bは的中していた。01節の該当svgは見本（`sample_figures.html`）に
潜んでいた真の違反であり、C5が一度も当てられていなかったことによる見落としと判明した（規約の欠陥ではない）。
信玄の指示により「是正3」として `fault/base.html` の483行から当該svgを削除した（`sample_figures.html` 本体は
無変更）。あわせて `case-a.html` を是正3適用後のbase.htmlから作り直し、C5注入を03節の `.sec-head` に入れ直した。
結果、`margin-left:auto` の出現数は base.html＝0件／case-a.html＝1件となり、対照群と注入群の切り分けが成立した
（実測値は「2. case-a.htmlに仕込んだ改変」節末尾を参照）。

### 【信玄が別途発見した事項】03節の散布図が `.fig-block` を持たず、C1のカウントが実態と食い違う

03節（コスト対性能、618〜650行）には本物の位置づけ図（軸2本＋矢頭・3モデルのプロット・破線矢印・ラベル6件）が
存在するが、`.fig-block` や `data-fig-type` を伴わず、`.scatter-wrap` という別クラスで置かれている。
これは実質的にT6（位置づけ）の図解である。

帰結：
- C1のA（`data-fig-type` の出現数）＝2（02節「比率」・04節「対比」）という現状カウントは実態を過小評価しており、
  **図解を持つ節は実際には3節**である。C1はこの03節を「図解を置かなかった節」に誤分類する。
- 図解節が実は3節なら、**C3「3種以上」の下限は本来この見本で課される側**だったことになる
  （現状マークアップ上の型は比率・対比の2種＋マークアップに現れない位置づけ1種）。

**この03節に `.fig-block`/`data-fig-type` ラッパーを追加する是正は行っていない。** 意味論（どの節が図解節か・
何種の型を使っているか）に踏み込む変更であり、C7/C8(b)向けの清書（表層のクラス付与・不要行削除）の範囲を
超えるため、信玄の指示どおり「対照群の既知状態」として本ノートに記録するに留めた。

---

## 5. 残り10本（case-b〜case-h・case-k、fault-c9/case-i・case-j-writeup）

すべて `fault\base.html` から複製した（`case-a.html` からは複製していない）。
各ファイルとも、狙った検査に効く改変のみを1ファイル1論点で入れ、diffで単一改変であることを確認済み
（case-kのみ、指示通り2箇所を意図的に同居させている）。

### case-b.html（狙い：C1／C2系のカウンタ、`class="fig-block"`残置＋`data-fig-type`欠落）

- 改変箇所：512行。02節の `<div class="fig-block" data-fig-type="比率">` から `data-fig-type="比率"` のみを削除し、
  `<div class="fig-block">` にした。
- 元→後：`<div class="fig-block" data-fig-type="比率">` → `<div class="fig-block">`
- diff確認：この1行のみが差分。

### case-c.html（狙い：C1の語彙チェック、7語彙外の値）

- 改変箇所：665行。04節の `data-fig-type="対比"` を `data-fig-type="比較"` に書き換えた（「比較」は規約の7語彙に含まれない表記揺れ）。
- 元→後：`<div class="fig-block" data-fig-type="対比">` → `<div class="fig-block" data-fig-type="比較">`
- diff確認：この1行のみが差分。

### case-d.html（狙い：C7の個数一致条件）

- 改変箇所：03節散布図の `<text>` 6件中2件（642行「性能↑」、644行「Opus 5」）から `class="fig-label"` を外した。
- 元→後（642行）：`<text class="fig-label" x="8" y="18" ...>性能↑</text>` → `<text x="8" y="18" ...>性能↑</text>`
- 元→後（644行）：`<text class="fig-label" x="240" y="115" ...>Opus 5</text>` → `<text x="240" y="115" ...>Opus 5</text>`
- diff確認：この2行のみが差分。結果、`<text` は6件のまま、`class="fig-label"` は4件に減り不一致となる。

### case-e.html（狙い：C6「個別ブロックの空洞化」）

- 改変箇所：02節fig-block内、518〜591行（Frontier-Bench〜OSWorld 2.0の5件のbench-row全体）を削除。
  fig-caption-row（magnifierアイコン＋キャプション文）のみ残した。
- 元→後：5件のbench-row（badge-circle・bar-track・bar-label-row一式）をまるごと削除、他は無変更。
- diff確認：削除範囲（518〜591行）のみが差分。
- **実測（信玄の指示どおり必須確認）**：この節はもともと `<use href=` を1件（magnifier）しか持たない
  （bench-rowはbadge-circle等がdivであり svg/use を含まない）ため、bench-row全削除でも
  当該ブロックの `<use href=` 数は**1件のまま変化しない**。ファイル全体の `<use href=` もbase.htmlと同じ**12件**。
  すなわちC6の合否判定に使う総数は指示どおり不変であり、「個別ブロックを空洞にしても総数で通ってしまうか」を
  文字通り再現できている。

### case-f.html（狙い：C1／C3、無宣言の3つ目の図解ブロック）

- 改変箇所：616〜650行（03節の `.scatter-wrap` 全体）を `<div class="fig-block" data-fig-type="比率">...</div>` で包んだ。
  開始タグを617行の直前に、終了タグを650行目（`.scatter-wrap`の閉じ`</div>`の直後）に追加。2行の純追加のみ。
- 元→後：`<div class="scatter-wrap">` の直前に `<div class="fig-block" data-fig-type="比率">` を追加／
  対応する `</div>` を末尾に追加。
- diff確認：追加した2行のみが差分。
- 結果：`data-fig-type` は3件（比率×2・対比×1）になり、図解ブロックは3つ・型は2種となる。

### case-g.html（狙い：C3「3種以上」下限）

- 改変箇所：665行。04節の `data-fig-type="対比"` を `data-fig-type="比率"` に書き換えた。
- 元→後：`<div class="fig-block" data-fig-type="対比">` → `<div class="fig-block" data-fig-type="比率">`
- diff確認：この1行のみが差分。
- 結果：図解ブロックは2つのまま、型は「比率」1種のみになる。

### case-h.html（狙い：C4／C8(b)、fig-block内の48px要素・rotate要素／既存の`.circled`装飾rotateの誤検出有無）

- 改変箇所：665行の直後（04節fig-block冒頭、`.flow-wrap`の直前）に2行追加。
  1. `<svg width="48" height="48" aria-hidden="true"><use href="#check-mark"/></svg>`（width="48"を持つ要素）
  2. `<svg width="10" height="10" aria-hidden="true"><rect transform="rotate(15)" width="1" height="1" fill="none"/></svg>`
     （`transform="rotate(15)"`を持つ要素。有効なSVGとして成立させるため`<svg>`で包んだ）
- 123行の `.circled::before{ transform: rotate(-1deg); }`（`.fig-block`外の装飾）は無変更のまま維持。
- diff確認：追加した2行のみが差分。
- **実測**：`rotate(` はbase.html の1件（123行の`.circled::before`のみ）から**2件**に増加（新規注入分＋既存の装飾1件）。
  既存の装飾側が誤検出されないか（＝`.fig-block`外の1件と、内側の1件を判別できるか）を見る材料になる。
  `<use href=` は12件→13件（check-markアイコンを1件追加したことによる副次的な増加。狙いに直接効く改変ではないが、
  fig-partアイコンとして自然に48px要素を実装した結果生じた）。

### case-k.html（狙い：C6、①分母を下げる②分子を下げる、の同時発生）

- **改変①**：665行。04節の `<div class="fig-block" data-fig-type="対比">` から `class="fig-block"` と
  `data-fig-type="対比"` を両方外し、`<div>` にした（中身の `.flow-wrap` 以下は無変更のまま残置）。
- **改変②**：513〜591行。02節fig-block内の514行（magnifierアイコン＝当該ブロック唯一の`<use href=`）と、
  518〜591行の5件のbench-row（bar-track等の主要な部品）を削除。fig-caption-rowはキャプション文のみ残した。
- diff確認：①・②の該当範囲以外に差分なし。
- **実測（信玄の指示どおり必須確認）**：`data-fig-type`＝**1件**（02の「比率」のみ。04は属性ごと除去）、
  `<use href=`＝**11件**（base.htmlの12件から、02のmagnifier1件を除去した分だけ減少）。

**【要報告】case-kの数値未達について**：指示書は「ファイル全体の`<use href=`を3〜4件まで落とす」ことを
「②02節ブロック内の部品を削り」で達成するよう求めていたが、02節の `.fig-block` は前述case-eの実測が示す通り
**そもそも `<use href=` を1件（magnifier）しか持たない**（bench-row側はdivのみでsvg/useを含まない）。
そのため「02節ブロック内の部品」をどれだけ削っても、そこから減らせる `<use href=` は最大1件であり、
本ファイルの実測値は11件（12→11）に留まる。3〜4件まで落とすには、04節の中身（talking-person・hand-arrow・flask、
計3件。改変①で外側の`.fig-block`ラッパーは外したが中身は指示通り残置しているため健在）や、
02節owl-row・09節・ヒーローのアイコン等、**「02節ブロック」の外側**にまで手を入れる必要があり、
それは「1ファイルにつき、狙った検査に効く改変のみを入れよ」「ついでの修正をするな」という大原則、および
「改変は指示された箇所に限定する」という統制と衝突する。今回は指示書の字義（②は02節ブロック内に限定）を
優先し、範囲外への削除は行わなかった。**目標値3〜4件と実際に達成可能な値（11件）に乖離がある**ため、
このまま検査に掛けてよいか、範囲を広げてよいか、信玄の判断を仰ぎたい。

### fault-c9/case-i.html（狙い：C9系、既存情報を削って部品を足す事故の再現）

`fault-c9` ディレクトリを新規作成し、`fault\base.html` から複製した。

- 改変箇所1（削除）：522行。Frontier-Bench行の `<span class="bench-vs">vs Opus 4.8</span>` を
  `<span class="bench-vs"></span>` にし、比較対象表記を削除。
- 改変箇所2（削除）：540行。ARC-AGI 3行の `<div class="badge-circle">×3</div>` を `<div class="badge-circle"></div>` にし、
  倍率バッジの表記を削除。この2件目は元のCSS（`.badge-circle`は92×92pxの固定枠＋`border-radius:50%`）により、
  文字を消しても枠自体は残るため、**バー群（`.bench-bars`）の左横に空白の円形の帯**が生じる形になる
  （指示書の「バー群の横に不自然な空白帯が生じる形」に対応）。
- 改変箇所3（追加）：515行に `<svg class="fig-part" width="40" height="40" aria-hidden="true"><use href="#check-mark"/></svg>`
  を追加し、削った情報の代わりに部品svgを1点加えた。
- diff確認：上記3行変化以外に差分なし。
- 実測：`<use href=`＝13件（12→13、追加分）。他の指標（`data-fig-type`=2、`<text`=6、`class="fig-label"`=6、
  `margin-left:auto`=0、`rotate(`=1）はbase.htmlと同一で、この改変が意図せず他の検査に影響していないことを確認。

### fault-c9/case-j-writeup.md（狙い：C9の書き出し、不当な理由の混入検知）

HTMLではなく検品記録の文書として作成。`fault\base.html` で「図解を置かなかった節」（`data-fig-type`を持たない8節：
01・03・05・06・07・08・09・10）それぞれについて、図解を要しない理由を1行ずつ記載した。
うち2行（09・10）に「本文で厚く説明したため」「文章量が十分にあるため」という**不当な理由**を、
残る6行（01・03・05・06・07・08）には正当な理由（並列列挙で順序・因果・量比を持たない等）を記載した。
どの行が不当かは文書内に示していない（指示通り）。

なお03の行については、**この節は実体としては散布図を備えており「図解が無い」という前提自体が怪しい**
（本ノート4節の既発見事項と同一の齟齬）ため、他5行のような「関係が無い」という言い切りではなく、
分類側の齟齬を正直に記す形にした。これは「本文が厚いから」という不当理由のパターンには当たらないが、
8行のうち1行だけ性質が異なる点は検査役（内藤）に申し添えておく。

### 実測件数一覧（全ファイル・Grep `-o` による机上ではなく実測値）

| ファイル | data-fig-type | `<use href=` | `<text` | `class="fig-label"` | `margin-left:auto` | `rotate(` |
|---|---|---|---|---|---|---|
| fault/base.html | 2 | 12 | 6 | 6 | 0 | 1 |
| fault/case-a.html | 2 | 13 | 6 | 6 | 1 | 1 |
| fault/case-b.html | 1 | 12 | 6 | 6 | 0 | 1 |
| fault/case-c.html | 2 | 12 | 6 | 6 | 0 | 1 |
| fault/case-d.html | 2 | 12 | 6 | 4 | 0 | 1 |
| fault/case-e.html | 2 | 12 | 6 | 6 | 0 | 1 |
| fault/case-f.html | 3 | 12 | 6 | 6 | 0 | 1 |
| fault/case-g.html | 2 | 12 | 6 | 6 | 0 | 1 |
| fault/case-h.html | 2 | 13 | 6 | 6 | 0 | 2 |
| fault/case-k.html | 1 | 11 | 6 | 6 | 0 | 1 |
| fault-c9/case-i.html | 2 | 13 | 6 | 6 | 0 | 1 |

## 6. 気づいた規約側・指示側の疑義（追加分）

### 疑義C：case-kの数値目標（3〜4件）は「02節ブロック内限定」という範囲制約と両立しない

詳細は上記「case-k.html」節に記載の通り。02節の `.fig-block` はbase.html の時点で`<use href=`を1件しか持たず、
そこから削れる分は最大1件（12→11）である。目標の3〜4件に到達するには04節本体（改変①で外側のラッパーだけ外し、
中身の3件は指示通り残置している）や02節owl-row・09節・ヒーローのアイコンなど「02節ブロック」の外側に
手を入れる必要があり、これは「1ファイルにつき、狙った検査に効く改変のみ」という大原則と衝突する。
指示の字義を優先し範囲外の削除は行わなかった。信玄の判断を仰ぎたい。

### 疑義D：case-eとcase-kで判明した「02節fig-blockは`<use href=`を実質1件しか持たない」という構造的な偏り

02節（比率）は5件のbenchmarkを表示する厚い節だが、視覚的な「部品」（svg/use）はキャプション行の
magnifierアイコン1点のみで、残りは全てdiv（badge-circle・bar-track等）で構成されている。
一方04節（対比）は3件の`<use href=`（talking-person・hand-arrow・flask）を持つ。
C6「`<use href=`の総数 ≧ `data-fig-type`の数×3」という指標は、**節ごとの部品密度の偏りを検出できない**
（02節が1件、04節が3件でも、ファイル全体で合算すれば基準を満たしてしまう）。
これはcase-e・case-kの実測が示す通りで、C6が想定する「型ごとに平均3部品」という基準は、
個別ブロックの空洞化を検出する目的には無力であることが今回の作業で定量的に裏付けられた。

## 7. 作成した全ファイルのパス（今回分）

- `fault\case-b.html` / `case-c.html` / `case-d.html` / `case-e.html` / `case-f.html` / `case-g.html` /
  `case-h.html` / `case-k.html`
- `fault-c9\case-i.html`
- `fault-c9\case-j-writeup.md`
- 本ファイル（`fault-NOTES.md`、追記のみ・末尾に追加）

---

## 8. case-l.html（信玄・裁定2による追加指示。C6の陽性対照＝C6が実際に落ちる条件を一度作る）

### 位置づけ・大原則の例外である旨

case-a〜case-k はすべて「1ファイル1改変（狙った検査に効く改変のみ）」を守ってきたが、
**case-l.html は意図的に複数箇所（7箇所）を同時に改変する。** 理由は、case-e・case-k が示した
「個別ブロックを空洞にしても／ラッパーごと消しても、C6（`<use href=`の総数 ≧ `data-fig-type`の数×3）は
合格し続ける」という結果だけでは、**「C6に穴がある」のか「C6はそもそも一度も落ちない死条項なのか」が
区別できない**ため（信玄・裁定2）。C6が実際に不合格になる条件を一度作り、C6が機能しうる条項であることを
確認するための**陽性対照（positive control）**として作成した。狙った検査はC6のみだが、
「`<use href=`を7件同時に削る」こと自体が改変の本体であるため、複数箇所改変を例外的に許容している。

### base.htmlから複製し、削除した7件（すべて図解ブロック=`.fig-block`の外側にある装飾アイコン）

`fault\base.html`（`<use href=`12件）を複製し、以下7件の装飾アイコンを削除した。

1. **496行**：02節 `.owl-row`（`.fig-block`は512〜592行であり、495〜498行の`.owl-row`は`.sec-head`より前・
   `.fig-block`の外）。元の記述：`<svg class="owl" width="52" height="52"><use href="#owl-doc"/></svg>`
2. **822行**：09節 `.owl-row`（09節に`.fig-block`は存在しない。09節全体が図解ブロック外）。
   元の記述：`<svg class="owl" width="52" height="52"><use href="#owl-doc"/></svg>`
3. **833行**：09節 `.safety-list` 1件目（09節は図解ブロック外）。
   元の記述：`<li><svg class="check-icon" width="20" height="18"><use href="#check-mark"/></svg> 自動行動監査で「これまでで最も整列」（most aligned to date）</li>`
   → `<li>自動行動監査で「これまでで最も整列」（most aligned to date）</li>`
4. **834行**：同上2件目。元：`<li><svg class="check-icon" width="20" height="18"><use href="#check-mark"/></svg> Claude's Constitution への遵守度が Opus 4.8／Sonnet 5／Fable 5 より高い</li>` → svgのみ除去。
5. **835行**：同上3件目。元：`<li><svg class="check-icon" width="20" height="18"><use href="#check-mark"/></svg> 最低の欺瞞行為率</li>` → svgのみ除去。
6. **836行**：同上4件目。元：`<li><svg class="check-icon" width="20" height="18"><use href="#check-mark"/></svg> 悪用への耐性が最高</li>` → svgのみ除去。
7. **837行**：同上5件目。元：`<li><svg class="check-icon" width="20" height="18"><use href="#check-mark"/></svg> サイバー分類器の介入頻度：Fable 5 比 <span class="circled">約85%削減</span></li>` → svgのみ除去。

いずれも`<li>`・`.bubble`・テキスト本体は残し、svg（アイコン）部分のみを除去した。
diff確認：上記7行の変化以外に差分なし。

### 図解ブロック内の4件が無傷であることの確認

base.htmlの`<use href=`12件のうち、`.fig-block`の内側にあるのは以下4件のみであり、**case-l.htmlでも
一切変更していない**ことを実測で確認した。

- 513行：02節fig-block内、`<svg class="fig-part" ...><use href="#magnifier"/></svg>`
- 667行：04節fig-block内、`<svg class="fig-part" ...><use href="#talking-person"/></svg>`
- 671行：04節fig-block内、`<svg class="fig-part" ...><use href="#hand-arrow"/></svg>`
- 673行：04節fig-block内、`<svg class="fig-part" ...><use href="#flask"/></svg>`

削除した7件はいずれも464行（ヒーローのowl。今回は意図的に残置し、図解ブロック外の装飾を1件だけ残す設計とした）
を除く、02節owl-row・09節owl-row・09節check-mark×5の合計7件であり、**図解ブロックの部品には一件も触れていない。**

### 実測件数

- `<use href=`：**5件**（12件 − 7件。内訳＝464行のヒーローowl1件＋fig-block内4件）
- `data-fig-type`：**2件**（base.htmlと同じ、無変更）
- C6の合否判定：閾値＝`data-fig-type`数×3＝2×3＝**6件**。実測は5件で**6件に届かず、不合格になるはず**。

### 裁定2の発見（記録）

02節fig-block内の`<use href=`は1件、04節fig-block内は3件で、base.html全体12件のうち
**図解ブロック内にあるのはこの4件のみ**（残り8件は全て図解と無関係な装飾：ヒーロー1・02節owl-row1・
09節owl-row1・09節check-mark5）。C6は`<use href=`を**ファイル全体で**数えており、**図解ブロックの内側に
限定していない。** そのため「図解の中身が空でも、文書のどこかに装飾アイコンが散らばっていれば合格する」
という状態が起こり得る（case-e・case-kで実証済み）。case-l.htmlはこれを逆方向から照らす一本であり、
**図解ブロックの部品を一つも減らさずに、図解と無関係な装飾だけを消してもC6を落とせる**ことを示す
（＝C6が数えている母集団と、C6が本来評価すべきはずの母集団＝「図解ブロック内の部品」が一致していない、
という欠陥の所在をより明確にする）。

### 裁定1の記録（case-kは現状のまま確定）

case-k.html（`data-fig-type=1`／`<use href=`=11）はC6の閾値1×3=3を11が上回るため合格するが、
「片方の図解をラッパーごと消し、もう片方を空洞にしてなお合格する」ことでmaskingを実証済みであり、
これ以上3〜4件に寄せるための範囲拡大（大原則違反）は不要と信玄より裁定された。case-k.htmlへの追加変更は行っていない。
