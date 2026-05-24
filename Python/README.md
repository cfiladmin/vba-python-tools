# Python 自動化ツール

## ツール1: CSV整形・クレンジング

**「毎回Excelで手直ししているCSVの作業を自動化します」**

### 解決する課題

- 文字化けするCSV（Shift-JIS → UTF-8変換）
- セルに混入した余分なスペース
- 重複行の手動チェック
- 空行の目視削除

### Before / After

```
【Before: 受け取った時のCSV】          【After: 整形後のCSV】
"  株式会社A  ","  田中 ","50000"      "株式会社A","田中太郎","50000"
"株式会社B","鈴木花子","30000"    →    "株式会社B","鈴木花子","30000"
"  株式会社A  ","  田中 ","50000"      "株式会社C","佐藤一郎","80000"
"","","","  "                          "株式会社D","山田次郎","20000"
"株式会社D","山田次郎","20000"
```

7行 → 4行（重複2件・空行1件を自動除去）

### 使い方

```bash
pip install pandas
python PY-001_CSV整形クレンジング.py 入力ファイル.csv 出力ファイル.csv
```

### カスタマイズ箇所（3行だけ）

```python
INPUT_ENCODING      = "shift_jis"   # 受け取ったCSVの文字コードに合わせる
OUTPUT_ENCODING     = "utf-8-sig"   # Excelで開くならutf-8-sig推奨
DROP_DUPLICATES_COL = None          # 特定列で重複判定する場合は列名を指定
```

### 動作確認済み環境

- Python 3.14.3 / pandas 3.0.1
- 入力：Shift-JIS CSV（7行・重複2件・空行1件含む）
- 出力：UTF-8（BOM付き）CSV 4行 ✅

---

## ツール2: Excel帳票自動生成

**「毎月100件の請求書をExcelで手作りしている作業を10秒にします」**

### 解決する課題

- テンプレートExcelへの手動コピペ
- 件数が多いほど増える転記ミスのリスク
- ファイル名を一件ずつつける手間

### 仕組み

```
template.xlsx（雛形）
    ＋
data.csv（宛先・金額リスト）
    ↓
document_001.xlsx
document_002.xlsx   ← CSVの件数分を自動生成
document_003.xlsx
     …
```

### Before / After

```
【Before】手作業で1件5分 × 100件 = 500分（8時間以上）
【After】 python PY-003_Excel自動生成.py  → 約10秒で100件完成
```

### 使い方

```bash
pip install pandas openpyxl
python PY-003_Excel自動生成.py
```

`config`ファイル不要。スクリプト冒頭の3行を変更するだけ：

```python
TEMPLATE_FILE = "template.xlsx"   # テンプレートのファイル名
DATA_CSV      = "data.csv"        # データCSVのファイル名
OUTPUT_DIR    = Path("output")    # 出力先フォルダ
```

### セルの対応設定（MAPPINGを変更するだけ）

```python
MAPPING = {
    "会社名": "B3",   # CSVの「会社名」列 → ExcelのB3セル
    "担当者": "B4",
    "金額":   "D10",
    "日付":   "F2",
}
```

### 動作確認済み環境

- Python 3.14.3 / openpyxl 3.1.x / pandas 3.0.1
- 入力：data.csv 3件 + template.xlsx
- 出力：document_001〜003.xlsx（全セル正常書き込み） ✅

---

## 共通：納品物セット

```
スクリプト本体（.py）
操作手順書（PDF）
操作動画（Loom 2〜5分）
```

> AI使用・最終確認は人間が行っています。
