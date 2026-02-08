"""
market_radar_api.py - 市場調査自動化スクリプト（Google Custom Search API版）

指定したキーワードに基づいて以下を自動収集し、Excelレポートとして出力します。
  - Googleサジェスト（ユーザーの悩み・ニーズ）
  - Google Custom Search API（信頼できる情報源からの検索結果）

使い方:
  1. 下記「設定エリア」のキーワードリストを編集する
  2. python market_radar_api.py を実行する
  3. 同じディレクトリに market_report_api_YYYYMMDD_HHMMSS.xlsx が出力される

必要ライブラリ:
  pip install google-api-python-client requests pandas openpyxl
"""

import time
import xml.etree.ElementTree as ET
from datetime import datetime

import pandas as pd
import requests
from googleapiclient.discovery import build


# ============================================================
# 設定エリア（ここを編集してください）
# ============================================================

# Google Custom Search API 認証情報
API_KEY = "AIzaSyAlFQTnE8N7f3uN9s0JrKseosfiIUPCddc"
CX_ID = "95d5c465d517c41bd"

# 調査したいキーワードをリスト形式で指定してください
KEYWORDS = [
    "訪問看護 大阪",
    "老人ホーム 紹介",
    "サ高住 経営",
    "空き家 サブリース",
    "介護 離職",
]

# Googleサジェストの取得件数（上位N件）
SUGGEST_LIMIT = 5

# Google Custom Search の取得件数（最大10件/リクエスト）
SEARCH_NUM = 10

# 各APIコール間の待機秒数（レートリミット対策）
SLEEP_SECONDS = 1


# ============================================================
# 関数定義
# ============================================================


def fetch_google_suggests(keyword: str, limit: int = SUGGEST_LIMIT) -> str:
    """
    Googleサジェスト（オートコンプリート）を取得する。

    Args:
        keyword: 検索キーワード
        limit: 取得するサジェスト数の上限

    Returns:
        サジェストワードをカンマ区切りで連結した文字列
    """
    url = "http://www.google.com/complete/search"
    params = {"output": "toolbar", "q": keyword}

    response = requests.get(url, params=params, timeout=10)
    response.raise_for_status()

    # XMLレスポンスをパースしてサジェストワードを抽出
    root = ET.fromstring(response.content)
    suggestions = []
    for suggestion_elem in root.iter("suggestion"):
        data = suggestion_elem.get("data")
        if data:
            suggestions.append(data)
        if len(suggestions) >= limit:
            break

    return ", ".join(suggestions) if suggestions else "（取得なし）"


def extract_date_from_item(item: dict) -> str:
    """
    検索結果アイテムの pagemap メタデータから日付情報を抽出する。
    metatags の og:updated_time, article:published_time, date などを探索する。

    Args:
        item: Google Custom Search API が返す検索結果1件分の辞書

    Returns:
        見つかった日付文字列。見つからなければ空文字列。
    """
    pagemap = item.get("pagemap", {})

    # metatags から日付系フィールドを探す
    metatags_list = pagemap.get("metatags", [])
    if metatags_list:
        metatags = metatags_list[0]
        # よく使われる日付メタタグを優先順に試す
        date_keys = [
            "article:published_time",
            "og:updated_time",
            "date",
            "publishdate",
            "dc.date",
            "sailthru.date",
            "last-modified",
        ]
        for key in date_keys:
            value = metatags.get(key, "")
            if value:
                # ISO形式の長い日付は日付部分だけ切り出す（例: 2024-01-15T10:00:00+09:00 → 2024-01-15）
                return value[:10] if len(value) >= 10 else value

    return ""


def search_google_cse(keyword: str) -> list[dict]:
    """
    Google Custom Search API でウェブ検索を実行する。

    Args:
        keyword: 検索キーワード

    Returns:
        検索結果のリスト。各要素は {"title", "url", "snippet", "date"} を持つ辞書。
    """
    # Custom Search API クライアントを構築
    service = build("customsearch", "v1", developerKey=API_KEY)

    # 検索を実行
    response = (
        service.cse()
        .list(q=keyword, cx=CX_ID, num=SEARCH_NUM)
        .execute()
    )

    items = response.get("items", [])
    results = []
    for item in items:
        results.append(
            {
                "title": item.get("title", ""),
                "url": item.get("link", ""),
                "snippet": item.get("snippet", "").replace("\n", " "),
                "date": extract_date_from_item(item),
            }
        )

    return results


def process_keyword(keyword: str) -> list[dict]:
    """
    1つのキーワードに対してサジェスト取得と検索を実行し、
    結合された行データのリストを返す。

    Args:
        keyword: 調査対象キーワード

    Returns:
        行データ（辞書）のリスト
    """
    rows = []

    # --- Step 1: Googleサジェスト取得（ニーズ・悩み） ---
    print(f"  [Step 1] Googleサジェストを取得中...")
    try:
        suggests_text = fetch_google_suggests(keyword)
        count = len(suggests_text.split(", ")) if suggests_text != "（取得なし）" else 0
        print(f"    -> {count}件のサジェストを取得しました")
    except Exception as e:
        print(f"    -> サジェスト取得に失敗しました: {type(e).__name__}: {e}")
        suggests_text = "（取得エラー）"

    time.sleep(SLEEP_SECONDS)  # レートリミット対策

    # --- Step 2: Google Custom Search API検索（解決策・競合情報） ---
    print(f"  [Step 2] Google Custom Search APIで検索中...")
    try:
        search_results = search_google_cse(keyword)
        if search_results:
            print(f"    -> {len(search_results)}件ヒットしました")
        else:
            print(f"    -> 0件ヒット：検索結果が返りませんでした")
    except Exception as e:
        print(f"    -> 検索に失敗しました: {type(e).__name__}: {e}")
        search_results = []

    # --- データ統合 ---
    if search_results:
        for i, result in enumerate(search_results):
            rows.append(
                {
                    "キーワード": keyword,
                    # サジェストは先頭行にのみ表示（Excel上で見やすくするため）
                    "サジェスト（悩み）": suggests_text if i == 0 else "",
                    "タイトル": result["title"],
                    "URL": result["url"],
                    "要約": result["snippet"],
                    "日付": result["date"],
                }
            )
    else:
        # 検索結果が0件でもキーワードとサジェストは記録する
        rows.append(
            {
                "キーワード": keyword,
                "サジェスト（悩み）": suggests_text,
                "タイトル": "（ヒットなし）",
                "URL": "",
                "要約": "",
                "日付": "",
            }
        )

    return rows


def save_to_excel(df: pd.DataFrame) -> str:
    """
    DataFrameをExcelファイルとして保存する。
    ファイル名に実行日時を付与する。

    Args:
        df: 出力するDataFrame

    Returns:
        保存したファイルのパス
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"market_report_api_{timestamp}.xlsx"

    # openpyxlエンジンを使用してExcel出力
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="調査結果")

        # 列幅を見やすく調整
        worksheet = writer.sheets["調査結果"]
        column_widths = {
            "A": 22,  # キーワード
            "B": 45,  # サジェスト（悩み）
            "C": 50,  # タイトル
            "D": 50,  # URL
            "E": 60,  # 要約
            "F": 14,  # 日付
        }
        for col_letter, width in column_widths.items():
            worksheet.column_dimensions[col_letter].width = width

    return filename


# ============================================================
# メイン処理
# ============================================================


def main():
    print("=" * 60)
    print(" Market Radar API - 市場調査自動化ツール")
    print(" (Google Custom Search API 版)")
    print("=" * 60)
    print(f" 対象キーワード数: {len(KEYWORDS)}")
    print(f" 実行日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    all_rows = []

    for i, keyword in enumerate(KEYWORDS, start=1):
        print(f"\n[{i}/{len(KEYWORDS)}] キーワード: 「{keyword}」を処理中...")

        try:
            rows = process_keyword(keyword)
            all_rows.extend(rows)
        except Exception as e:
            # 個別キーワードの処理失敗はスクリプト全体を止めない
            print(f"  -> 予期しないエラーが発生しました: {type(e).__name__}: {e}")
            all_rows.append(
                {
                    "キーワード": keyword,
                    "サジェスト（悩み）": "（エラー）",
                    "タイトル": f"処理エラー: {e}",
                    "URL": "",
                    "要約": "",
                    "日付": "",
                }
            )

        # レートリミット対策（最後のキーワード以外で待機）
        if i < len(KEYWORDS):
            print(f"  -> {SLEEP_SECONDS}秒待機中...")
            time.sleep(SLEEP_SECONDS)

    # --- DataFrame作成・Excel出力 ---
    print(f"\n{'=' * 60}")
    print(" データ集計・Excel出力中...")

    df = pd.DataFrame(
        all_rows,
        columns=["キーワード", "サジェスト（悩み）", "タイトル", "URL", "要約", "日付"],
    )
    filename = save_to_excel(df)

    print(f" レポートを保存しました: {filename}")
    print(f" 合計 {len(df)} 行のデータを出力しました")
    print("=" * 60)
    print(" 完了！")


if __name__ == "__main__":
    main()
