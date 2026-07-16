---
name: excel
description: Excelファイル（.xlsx）の作成・読み込み・編集を行う。セルへのデータ書き込み、書式設定、グラフ作成、数式挿入など。
allowed-tools: Bash, Write, Read, Glob
---

# Excel 操作スキル

`openpyxl` ライブラリを使って Excel ファイルを操作する。
Pythonコマンドは常に `py` を使うこと（`python` ではなく）。

## 引数
$ARGUMENTS

## 基本方針
- ファイルパスが指定されていない場合は、カレントディレクトリまたは適切な場所に保存する
- 既存ファイルへの操作は必ず事前にファイルの存在確認を行う
- 操作完了後はファイルパスと概要を明示して報告する

## よく使うコードパターン

### 新規ブック作成
```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles import GradientFill
from openpyxl.utils import get_column_letter

wb = Workbook()
ws = wb.active
ws.title = "シート1"
```

### セルへのデータ書き込み
```python
# 直接指定
ws["A1"] = "タイトル"
ws["B1"] = 123
ws["C1"] = "=SUM(B2:B10)"  # 数式

# 行・列番号で指定
ws.cell(row=2, column=1, value="データ1")

# 複数行を一括追加
data = [
    ["名前", "売上", "達成率"],
    ["田中", 1500000, 0.95],
    ["鈴木", 2300000, 1.12],
]
for row in data:
    ws.append(row)
```

### 書式設定
```python
# フォント
ws["A1"].font = Font(name="Yu Gothic", size=14, bold=True, color="1A0A00")

# 背景色
ws["A1"].fill = PatternFill(fill_type="solid", fgColor="3D1A00")
ws["A1"].font = Font(color="E8D5A3", bold=True)

# 配置
ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# 罫線
thin = Side(style="thin", color="8B3A00")
ws["A1"].border = Border(left=thin, right=thin, top=thin, bottom=thin)

# 列幅・行高
ws.column_dimensions["A"].width = 20
ws.row_dimensions[1].height = 30
```

### 範囲への一括書式適用
```python
from openpyxl.styles import NamedStyle
for row in ws.iter_rows(min_row=1, max_row=1, min_col=1, max_col=5):
    for cell in row:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="3D1A00")
```

### セルの結合
```python
ws.merge_cells("A1:E1")
ws["A1"] = "タイトル行"
ws["A1"].alignment = Alignment(horizontal="center")
```

### グラフ作成
```python
from openpyxl.chart import BarChart, Reference
chart = BarChart()
chart.title = "売上グラフ"
chart.y_axis.title = "売上"
chart.x_axis.title = "担当者"

data_ref = Reference(ws, min_col=2, min_row=1, max_row=ws.max_row)
cats_ref = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
chart.add_data(data_ref, titles_from_data=True)
chart.set_categories(cats_ref)
ws.add_chart(chart, "E5")
```

### 既存ファイルを読み込む
```python
from openpyxl import load_workbook
wb = load_workbook("existing.xlsx")
ws = wb.active
for row in ws.iter_rows(values_only=True):
    print(row)
```

### 保存
```python
wb.save("output.xlsx")
print("保存完了: output.xlsx")
```

## 実行方法
Bash ツールで以下のように実行する：
```bash
py -c "
# ここにPythonコード
"
```
