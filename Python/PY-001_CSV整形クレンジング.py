# [PY-001] CSV整形・クレンジングスクリプト
# 用途: CSVの文字コード変換・空白除去・重複削除・列名統一
# 使い方: python PY-001_CSV整形クレンジング.py input.csv output.csv
# 必要ライブラリ: pip install pandas

import pandas as pd
import sys

INPUT_ENCODING      = "shift_jis"   # 変更可: utf-8, cp932
OUTPUT_ENCODING     = "utf-8-sig"   # utf-8-sig = BOM付き（Excelで文字化けしない）
DROP_DUPLICATES_COL = None          # 重複削除の基準列名（Noneで全列）


def clean_csv(input_path: str, output_path: str) -> None:
    df = pd.read_csv(input_path, encoding=INPUT_ENCODING, dtype=str)

    # 列名の空白除去
    df.columns = [c.strip() for c in df.columns]

    # 全セルの空白除去（pandas 3.x対応: mapで文字列を個別処理）
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)

    # 空白文字列をNaNに変換してから空行削除（",,," のような行も除去）
    df.replace("", pd.NA, inplace=True)
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
        print("使い方: python PY-001_CSV整形クレンジング.py input.csv output.csv")
        sys.exit(1)
    clean_csv(sys.argv[1], sys.argv[2])
