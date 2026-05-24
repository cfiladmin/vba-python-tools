# SKILL.md — 納品フロー & ツールテンプレート集

## 納品フロー標準（全案件共通）

```
1. 要件ヒアリング（必ず書面/チャットで残す）
2. ダミーデータでプロトタイプ作成（Claude Code使用）
3. クライアント確認・修正
4. 本番データで動作確認（クライアント環境）
5. 操作手順書作成（Markdown → PDF変換）
6. Loom動画録画（2〜5分、画面操作を見せる）
7. 納品 → 評価依頼 → メンテプラン提案
```

---

## VBAテンプレート集

> ファイル本体は `VBA/` フォルダに格納。カスタマイズ箇所は各ファイルの `← 変更` コメントを参照。

| ID | ファイル | 用途 |
|----|----------|------|
| VBA-001 | `VBA/VBA-001_集計マクロ.bas` | Dictionaryで高速集計・別シートに出力 |
| VBA-002 | `VBA/VBA-002_CSV一括取込マクロ.bas` | フォルダ内全CSVを1シートに結合 |
| VBA-003 | `VBA/VBA-003_帳票PDF自動出力.bas` | 指定範囲をPDFとして日付付きで保存 |

### [VBA-001] Excel集計マクロ（基本型）

```vba
' 用途: 指定シートのデータを集計して別シートに出力
' カスタマイズ: INPUT_SHEET, OUTPUT_SHEET, COL_TARGET を変更

Option Explicit

Const INPUT_SHEET  As String = "データ"
Const OUTPUT_SHEET As String = "集計"
Const COL_TARGET   As Long = 3   ' 集計対象列番号

Sub CollectData()
    Dim wsIn  As Worksheet
    Dim wsOut As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dict As Object

    Set wsIn  = ThisWorkbook.Sheets(INPUT_SHEET)
    Set wsOut = ThisWorkbook.Sheets(OUTPUT_SHEET)
    Set dict  = CreateObject("Scripting.Dictionary")

    lastRow = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).Row

    ' データ読み込み
    For i = 2 To lastRow
        Dim key As String
        key = CStr(wsIn.Cells(i, COL_TARGET).Value)
        If dict.exists(key) Then
            dict(key) = dict(key) + wsIn.Cells(i, COL_TARGET + 1).Value
        Else
            dict.Add key, wsIn.Cells(i, COL_TARGET + 1).Value
        End If
    Next i

    ' 出力
    wsOut.Cells.ClearContents
    wsOut.Cells(1, 1).Value = "項目"
    wsOut.Cells(1, 2).Value = "合計"

    Dim row As Long
    row = 2
    Dim k As Variant
    For Each k In dict.Keys
        wsOut.Cells(row, 1).Value = k
        wsOut.Cells(row, 2).Value = dict(k)
        row = row + 1
    Next k

    MsgBox "集計完了！" & (row - 2) & "件", vbInformation
End Sub
```

**流用チェックリスト:**
- [ ] INPUT_SHEET / OUTPUT_SHEET のシート名を変更
- [ ] COL_TARGET の列番号を変更
- [ ] ヘッダー行（2行目スタート）を確認

---

### [VBA-002] CSV一括取込マクロ

```vba
' 用途: フォルダ内の全CSVを読み込んで1シートに結合
' カスタマイズ: FOLDER_PATH, DELIMITER を変更

Option Explicit

Const FOLDER_PATH As String = "C:\work\csv_files\"   ' ← 変更
Const DELIMITER   As String = ","

Sub ImportAllCSV()
    Dim wsOut    As Worksheet
    Dim filePath As String
    Dim fileNum  As Integer
    Dim lineData As String
    Dim cols()   As String
    Dim outRow   As Long
    Dim isFirst  As Boolean

    Set wsOut = ThisWorkbook.Sheets.Add
    wsOut.Name = "CSV結合_" & Format(Now, "MMDD_HHmm")
    outRow = 1
    isFirst = True

    filePath = Dir(FOLDER_PATH & "*.csv")
    Do While filePath <> ""
        fileNum = FreeFile
        Open FOLDER_PATH & filePath For Input As #fileNum

        Dim skipHeader As Boolean
        skipHeader = Not isFirst   ' 2ファイル目以降はヘッダーをスキップ

        Do While Not EOF(fileNum)
            Line Input #fileNum, lineData
            If skipHeader Then
                skipHeader = False
            Else
                cols = Split(lineData, DELIMITER)
                Dim c As Long
                For c = 0 To UBound(cols)
                    wsOut.Cells(outRow, c + 1).Value = cols(c)
                Next c
                outRow = outRow + 1
            End If
        Loop

        Close #fileNum
        isFirst = False
        filePath = Dir
    Loop

    MsgBox "取込完了！" & (outRow - 1) & "行", vbInformation
End Sub
```

---

### [VBA-003] 帳票PDF自動出力マクロ

```vba
' 用途: 指定範囲をPDFとして保存（請求書・納品書等）
' カスタマイズ: RANGE_ADDRESS, SAVE_FOLDER を変更

Option Explicit

Const RANGE_ADDRESS As String = "A1:H40"    ' ← 印刷範囲
Const SAVE_FOLDER   As String = "C:\work\pdf_output\"  ' ← 保存先

Sub ExportToPDF()
    Dim ws       As Worksheet
    Dim fileName As String
    Dim savePath As String

    Set ws = ActiveSheet

    ' ファイル名：シート名 + 今日の日付
    fileName = ws.Name & "_" & Format(Date, "YYYYMMDD") & ".pdf"
    savePath = SAVE_FOLDER & fileName

    ws.Range(RANGE_ADDRESS).ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=savePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False

    MsgBox "PDF出力完了: " & savePath, vbInformation
End Sub
```

---

## Pythonテンプレート集

> ファイル本体は `Python/` フォルダに格納。`← 変更` のコメント箇所だけ書き換えて流用する。

| ID | ファイル | 用途 |
|----|----------|------|
| PY-001 | `Python/PY-001_CSV整形クレンジング.py` | 文字コード変換・空白除去・重複削除 |
| PY-002 | `Python/PY-002_Webスクレイピング.py` | ページネーション対応・CSV保存 |
| PY-003 | `Python/PY-003_Excel自動生成.py` | テンプレートへのデータ流し込み |
| PY-004 | `Python/PY-004_フォルダ監視自動処理.py` | CSV追加を検知して自動整形 |

### [PY-001] CSV整形・クレンジングスクリプト

```python
# 用途: CSVの文字コード変換・空白除去・重複削除・列名統一
# 使い方: python clean_csv.py input.csv output.csv

import pandas as pd
import sys
from pathlib import Path

INPUT_ENCODING  = "shift_jis"   # 変更可: utf-8, cp932
OUTPUT_ENCODING = "utf-8"
DROP_DUPLICATES_COL = None      # 重複削除の基準列名（Noneで全列）


def clean_csv(input_path: str, output_path: str) -> None:
    df = pd.read_csv(input_path, encoding=INPUT_ENCODING, dtype=str)

    # 列名の空白除去
    df.columns = [c.strip() for c in df.columns]

    # 全セルの空白除去
    df = df.apply(lambda col: col.str.strip() if col.dtype == "object" else col)

    # 空行削除
    df.dropna(how="all", inplace=True)

    # 重複削除
    if DROP_DUPLICATES_COL:
        df.drop_duplicates(subset=[DROP_DUPLICATES_COL], inplace=True)
    else:
        df.drop_duplicates(inplace=True)

    df.to_csv(output_path, index=False, encoding=OUTPUT_ENCODING)
    print(f"完了: {len(df)}行 → {output_path}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("使い方: python clean_csv.py input.csv output.csv")
        sys.exit(1)
    clean_csv(sys.argv[1], sys.argv[2])
```

---

### [PY-002] Webスクレイピング基本テンプレート

```python
# 用途: 指定URLからデータを取得してCSV保存
# 前提: robots.txtで許可されたサイトのみ使用
# 使い方: python scrape.py

import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

TARGET_URL   = "https://example.com/list"   # ← 変更
OUTPUT_CSV   = "output.csv"
WAIT_SECONDS = 2   # サーバー負荷軽減（必ず守る）
HEADERS = {
    "User-Agent": "Mozilla/5.0 (research bot; contact: your@email.com)"
}


def scrape_page(url: str) -> list[dict]:
    resp = requests.get(url, headers=HEADERS, timeout=10)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    results = []
    # ↓ セレクターをサイトに合わせて変更
    for item in soup.select(".item-class"):
        results.append({
            "タイトル": item.select_one(".title").get_text(strip=True),
            "価格":     item.select_one(".price").get_text(strip=True),
            "URL":      item.select_one("a")["href"],
        })
    return results


def main():
    all_data = []
    page = 1

    while True:
        url = f"{TARGET_URL}?page={page}"
        print(f"取得中: {url}")
        data = scrape_page(url)
        if not data:
            break
        all_data.extend(data)
        page += 1
        time.sleep(WAIT_SECONDS)

    df = pd.DataFrame(all_data)
    df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    print(f"完了: {len(df)}件 → {OUTPUT_CSV}")


if __name__ == "__main__":
    main()
```

---

### [PY-003] Excel自動生成スクリプト

```python
# 用途: テンプレートExcelにデータを流し込んで複数ファイル生成
# 使い方: python gen_excel.py

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pathlib import Path
import pandas as pd

TEMPLATE_FILE = "template.xlsx"   # ← テンプレートExcelのパス
DATA_CSV      = "data.csv"        # ← 流し込むCSV
OUTPUT_DIR    = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

# CSV列名とExcelセルの対応 （列名: セル番地）
MAPPING = {
    "会社名": "B3",
    "担当者": "B4",
    "金額":   "D10",
    "日付":   "F2",
}


def fill_template(template_path: str, data: dict, output_path: str) -> None:
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    for col_name, cell_addr in MAPPING.items():
        if col_name in data:
            ws[cell_addr] = data[col_name]

    wb.save(output_path)


def main():
    df = pd.read_csv(DATA_CSV, dtype=str)

    for i, row in df.iterrows():
        output_file = OUTPUT_DIR / f"document_{i+1:03d}.xlsx"
        fill_template(TEMPLATE_FILE, row.to_dict(), str(output_file))
        print(f"生成: {output_file}")

    print(f"完了: {len(df)}件")


if __name__ == "__main__":
    main()
```

---

### [PY-004] フォルダ監視 & 自動処理スクリプト

```python
# 用途: 指定フォルダにファイルが追加されたら自動でCSV整形を実行
# 使い方: python watcher.py

import time
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import pandas as pd

WATCH_FOLDER  = Path("C:/work/input")    # ← 監視フォルダ
OUTPUT_FOLDER = Path("C:/work/output")   # ← 出力フォルダ
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)


class CSVHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        path = Path(event.src_path)
        if path.suffix.lower() == ".csv":
            print(f"検知: {path.name}")
            self.process(path)

    def process(self, path: Path) -> None:
        try:
            df = pd.read_csv(path, encoding="shift_jis", dtype=str)
            df.columns = [c.strip() for c in df.columns]
            df = df.apply(lambda col: col.str.strip() if col.dtype == "object" else col)
            df.dropna(how="all", inplace=True)

            out = OUTPUT_FOLDER / path.name
            df.to_csv(out, index=False, encoding="utf-8-sig")
            print(f"出力完了: {out}")
        except Exception as e:
            print(f"エラー: {e}")


if __name__ == "__main__":
    observer = Observer()
    observer.schedule(CSVHandler(), str(WATCH_FOLDER), recursive=False)
    observer.start()
    print(f"監視中: {WATCH_FOLDER}  (Ctrl+C で停止)")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
```

---

## テンプレート流用ログ

| 案件ID | 使用テンプレ | カスタマイズ内容 | 納品日 | 評価 |
|--------|-------------|-----------------|--------|------|
| (記入例) | VBA-001 | 集計列を売上→数量に変更 | - | - |

---

## 操作手順書テンプレート（Markdown）

```markdown
# ツール操作マニュアル

## 環境要件
- Excel 2016以上（VBAの場合）
- Python 3.10以上 / 必要ライブラリ: `pip install -r requirements.txt`

## セットアップ
1. ファイルを解凍してCドライブ直下に配置
2. `config.py` の設定値を変更（詳細は下記）

## 使い方
1. XXXを開く
2. YYYボタンをクリック
3. ZZZフォルダを選択
4. 実行 → 完了ダイアログを確認

## よくあるエラーと対処法
| エラー | 原因 | 対処 |
|--------|------|------|
| ファイルが見つからない | パス設定ミス | config.pyのPATHを確認 |
| 文字化け | エンコード不一致 | INPUT_ENCODINGをutf-8に変更 |

## サポート
納品後30日間、同様の不具合は無償対応します。
```
