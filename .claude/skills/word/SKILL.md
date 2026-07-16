---
name: word
description: Wordファイル（.docx）の作成・編集・操作を行う。文章の追加、見出し・表・リスト・画像の挿入、書式設定など。
allowed-tools: Bash, Write, Read, Glob
---

# Word 操作スキル

`python-docx` ライブラリを使って Word ファイルを操作する。
Pythonコマンドは常に `py` を使うこと（`python` ではなく）。

## 引数
$ARGUMENTS

## 基本方針
- ファイルパスが指定されていない場合は、カレントディレクトリまたは適切な場所に保存する
- 既存ファイルへの操作は必ず事前にファイルの存在確認を行う
- 操作完了後はファイルパスと概要を明示して報告する

## よく使うコードパターン

### 新規文書作成
```python
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
import copy

doc = Document()
```

### ページ設定
```python
from docx.shared import Cm
section = doc.sections[0]
section.page_width = Cm(21)    # A4横幅
section.page_height = Cm(29.7) # A4縦幅
section.left_margin = Cm(2.5)
section.right_margin = Cm(2.5)
section.top_margin = Cm(2.5)
section.bottom_margin = Cm(2.5)
```

### 見出し・段落追加
```python
# 見出し（レベル0=タイトル、1〜9=見出し）
doc.add_heading("ドキュメントタイトル", level=0)
doc.add_heading("第1章 概要", level=1)
doc.add_heading("1.1 背景", level=2)

# 段落
p = doc.add_paragraph("本文テキストをここに記述します。")
p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY  # 両端揃え

# 段落の書式設定
run = p.add_run("太字テキスト")
run.bold = True
run.font.size = Pt(12)
run.font.color.rgb = RGBColor(0x1a, 0x0a, 0x00)
```

### 箇条書き・番号リスト
```python
# 箇条書き
doc.add_paragraph("項目1", style="List Bullet")
doc.add_paragraph("項目2", style="List Bullet")

# 番号付きリスト
doc.add_paragraph("手順1", style="List Number")
doc.add_paragraph("手順2", style="List Number")
```

### 表の追加
```python
table = doc.add_table(rows=3, cols=4)
table.style = "Table Grid"
table.alignment = WD_TABLE_ALIGNMENT.CENTER

# ヘッダー行
hdr_cells = table.rows[0].cells
headers = ["項目", "数量", "単価", "合計"]
for i, h in enumerate(headers):
    hdr_cells[i].text = h
    hdr_cells[i].paragraphs[0].runs[0].bold = True

# データ行
data = [["商品A", "10", "¥1,500", "¥15,000"],
        ["商品B", "5",  "¥3,000", "¥15,000"]]
for row_data in data:
    row_cells = table.add_row().cells
    for i, val in enumerate(row_data):
        row_cells[i].text = val
```

### 改ページ
```python
doc.add_page_break()
```

### 画像挿入
```python
doc.add_picture("image.png", width=Inches(4))
```

### 既存ファイルを読み込む
```python
doc = Document("existing.docx")
for para in doc.paragraphs:
    print(f"[{para.style.name}] {para.text}")
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            print(cell.text, end="\t")
        print()
```

### フォント・スタイルのカスタマイズ
```python
from docx.oxml import OxmlElement
# 日本語フォント設定
style = doc.styles["Normal"]
style.font.name = "Yu Gothic"
style.font.size = Pt(10.5)
```

### 保存
```python
doc.save("output.docx")
print("保存完了: output.docx")
```

## 実行方法
Bash ツールで以下のように実行する：
```bash
py -c "
# ここにPythonコード
"
```
