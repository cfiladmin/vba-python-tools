# [PY-002] Webスクレイピング基本テンプレート
# 用途: 指定URLからデータを取得してCSV保存
# 前提: robots.txtで許可されたサイトのみ使用
# 使い方: python PY-002_Webスクレイピング.py
# 必要ライブラリ: pip install requests beautifulsoup4 pandas

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
