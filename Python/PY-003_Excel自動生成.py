# [PY-003] Excel自動生成スクリプト
# 用途: テンプレートExcelにデータを流し込んで複数ファイル生成
# 使い方: python PY-003_Excel自動生成.py
# 必要ライブラリ: pip install openpyxl pandas

import openpyxl
from pathlib import Path
import pandas as pd

TEMPLATE_FILE = "template.xlsx"   # ← テンプレートExcelのパス
DATA_CSV      = "data.csv"        # ← 流し込むCSVのパス
OUTPUT_DIR    = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

# CSV列名とExcelセルの対応 （列名: セル番地）← 変更
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
