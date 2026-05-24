# [PY-004] フォルダ監視 & 自動処理スクリプト
# 用途: 指定フォルダにCSVが追加されたら自動で整形して出力フォルダへ保存
# 使い方: python PY-004_フォルダ監視自動処理.py
# 必要ライブラリ: pip install watchdog pandas

import time
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import pandas as pd

WATCH_FOLDER  = Path("C:/work/input")    # ← 監視フォルダ（変更）
OUTPUT_FOLDER = Path("C:/work/output")   # ← 出力フォルダ（変更）
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
