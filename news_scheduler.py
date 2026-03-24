#!/usr/bin/env python3
"""毎朝9時に自治体関連ニュース10件をコンソールに表示するスケジューラー"""

import xml.etree.ElementTree as ET
from datetime import datetime
from urllib.request import urlopen
from urllib.parse import quote

import schedule
import time


QUERY = "自治体"
NEWS_COUNT = 10
SCHEDULE_TIME = "09:00"


def fetch_news():
    """Google News RSSから自治体関連ニュースを取得して表示する"""
    encoded_query = quote(QUERY)
    url = f"https://news.google.com/rss/search?q={encoded_query}&hl=ja&gl=JP&ceid=JP:ja"

    print(f"\n{'=' * 60}")
    print(f"  自治体関連ニュース  ({datetime.now().strftime('%Y-%m-%d %H:%M')})")
    print(f"{'=' * 60}\n")

    try:
        with urlopen(url, timeout=10) as response:
            xml_data = response.read()

        root = ET.fromstring(xml_data)
        channel = root.find("channel")
        items = channel.findall("item") if channel is not None else []

        if not items:
            print("ニュースが見つかりませんでした。")
            return

        for i, item in enumerate(items[:NEWS_COUNT], start=1):
            title = item.findtext("title", "タイトルなし")
            link = item.findtext("link", "")
            pub_date = item.findtext("pubDate", "")
            source = item.findtext("source", "")

            # 日付の整形
            date_str = ""
            if pub_date:
                try:
                    from email.utils import parsedate_to_datetime
                    dt = parsedate_to_datetime(pub_date)
                    date_str = dt.strftime("%m/%d %H:%M")
                except Exception:
                    date_str = pub_date[:16]

            print(f"[{i:2d}] {title}")
            if source:
                print(f"      出典: {source}  {date_str}")
            if link:
                print(f"      {link}")
            print()

    except Exception as e:
        print(f"ニュースの取得中にエラーが発生しました: {e}")

    print(f"{'=' * 60}\n")


def main():
    print(f"スケジューラー起動: 毎日 {SCHEDULE_TIME} に自治体関連ニュースを表示します")
    print("終了するには Ctrl+C を押してください\n")

    schedule.every().day.at(SCHEDULE_TIME).do(fetch_news)

    # 起動時にすぐ1回表示したい場合は以下のコメントを外してください
    # fetch_news()

    while True:
        schedule.run_pending()
        time.sleep(30)


if __name__ == "__main__":
    main()
