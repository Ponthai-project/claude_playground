---
name: powerpoint
description: PowerPointファイル（.pptx）の作成・編集・操作を行う。スライド追加、テキスト・図形・画像の挿入、書式設定など。
allowed-tools: Bash, Write, Read, Glob
---

# PowerPoint 操作スキル

`python-pptx` ライブラリを使って PowerPoint ファイルを操作する。
Pythonコマンドは常に `py` を使うこと（`python` ではなく）。

## 引数
$ARGUMENTS

## 基本方針
- ファイルパスが指定されていない場合は、カレントディレクトリまたは適切な場所に保存する
- 既存ファイルへの操作は必ず事前に Read または Glob で確認する
- 操作完了後はファイルパスを明示して報告する

## よく使うコードパターン

### 新規プレゼン作成
```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

prs = Presentation()
# スライドサイズ設定（16:9）
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)
```

### スライド追加
```python
# レイアウト: 0=タイトル, 1=タイトルとコンテンツ, 5=空白, 6=タイトルのみ
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)

# タイトル設定
title = slide.shapes.title
title.text = "スライドタイトル"

# コンテンツ設定
body = slide.placeholders[1]
tf = body.text_frame
tf.text = "本文テキスト"
```

### テキストボックス追加
```python
from pptx.util import Inches, Pt
txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(2))
tf = txBox.text_frame
tf.word_wrap = True
p = tf.add_paragraph()
p.text = "テキスト内容"
p.font.size = Pt(24)
p.font.bold = True
p.font.color.rgb = RGBColor(0x1a, 0x0a, 0x00)
p.alignment = PP_ALIGN.CENTER
```

### 図形追加
```python
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches
shape = slide.shapes.add_shape(
    1,  # MSO_SHAPE_TYPE.RECTANGLE
    Inches(1), Inches(1), Inches(4), Inches(2)
)
shape.fill.solid()
shape.fill.fore_color.rgb = RGBColor(0x8b, 0x3a, 0x00)
shape.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
```

### 表追加
```python
rows, cols = 3, 4
table = slide.shapes.add_table(rows, cols, Inches(1), Inches(2), Inches(10), Inches(3)).table
table.cell(0, 0).text = "ヘッダー1"
```

### 保存
```python
prs.save("output.pptx")
print("保存完了: output.pptx")
```

### 既存ファイルを開く
```python
prs = Presentation("existing.pptx")
for i, slide in enumerate(prs.slides):
    print(f"スライド {i+1}: 図形数={len(slide.shapes)}")
    for shape in slide.shapes:
        if shape.has_text_frame:
            print(f"  テキスト: {shape.text_frame.text[:50]}")
```

## 実行方法
Bash ツールで以下のように実行する：
```bash
py -c "
# ここにPythonコード
"
```

または一時ファイルに書き出して実行：
```bash
py path/to/script.py
```
