"""
market_radar.py - 市場調査自動化スクリプト

指定したキーワードに基づいて以下を自動収集し、Excelレポートとして出力します。
  - Googleサジェスト（ユーザーの悩み・ニーズ）
  - DuckDuckGo検索結果（解決策・競合情報）

使い方:
  1. 下記「設定エリア」のキーワードリストを編集する
  2. python market_radar.py を実行する
  3. 同じディレクトリに YYYYMMDD_HHMMSS_market_report.xlsx が出力される
"""

import time
import xml.etree.ElementTree as ET
from datetime import datetime

import pandas as pd
import requests
from duckduckgo_search import DDGS


# ============================================================
# 設定エリア（ここを編集してください）
# ============================================================

# 調査したいキーワードをリスト形式で指定してください
KEYWORDS = [
    "訪問看護 大阪",
    "老人ホーム 紹介",
    "サ高住 経営",
]

# Googleサジェストの取得件数（上位N件）
SUGGEST_LIMIT = 5

# DuckDuckGo検索の取得件数（上位N件）
SEARCH_LIMIT = 5

# DuckDuckGo検索の地域設定（日本）
SEARCH_REGION = "jp-jp"

# DuckDuckGo検索の期間制限（None=制限なし, "d"=1日, "w"=1週間, "m"=1ヶ月, "y"=1年）
# ※ 期間指定が厳しいと0件になることがあるため、デフォルトはNone（制限なし）
SEARCH_TIMELIMIT = None

# DuckDuckGo検索のバックエンド（"html"=スクレイピング方式, "api"=API方式）
# ※ apiモードはブロックされやすいため、htmlモードを推奨
SEARCH_BACKEND = "html"

# 各キーワード処理間の待機秒数（サーバー負荷軽減のため）
SLEEP_SECONDS = 5


# ============================================================
# 関数定義
# ============================================================


def fetch_google_suggests(keyword: str, limit: int = SUGGEST_LIMIT) -> list[str]:
    """
    Googleサジェスト（オートコンプリート）を取得する。

    Args:
        keyword: 検索キーワード
        limit: 取得するサジェスト数の上限

    Returns:
        サジェストワードのリスト
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

    return suggestions


def search_duckduckgo(
    keyword: str,
    limit: int = SEARCH_LIMIT,
    region: str = SEARCH_REGION,
    timelimit: str | None = SEARCH_TIMELIMIT,
    backend: str = SEARCH_BACKEND,
) -> list[dict]:
    """
    DuckDuckGoでウェブ検索を実行する。

    Args:
        keyword: 検索キーワード
        limit: 取得する検索結果数の上限
        region: 検索地域（例: "jp-jp"）
        timelimit: 期間制限（None=制限なし, "y"=過去1年 など）
        backend: 検索バックエンド（"html" or "api"）

    Returns:
        検索結果のリスト。各要素は {"title", "url", "snippet"} を持つ辞書。
    """
    results = []
    # 検索パラメータを構築（timelimitがNoneの場合は引数自体を渡さない）
    search_kwargs = {
        "keywords": keyword,
        "region": region,
        "max_results": limit,
        "backend": backend,
    }
    if timelimit is not None:
        search_kwargs["timelimit"] = timelimit

    print(f"    [DEBUG] 検索パラメータ: backend={backend}, region={region}, timelimit={timelimit}, max_results={limit}")

    ddgs = DDGS()
    raw_results = ddgs.text(**search_kwargs)

    for r in raw_results:
        results.append(
            {
                "title": r.get("title", ""),
                "url": r.get("href", ""),
                "snippet": r.get("body", ""),
            }
        )

    print(f"    [DEBUG] DuckDuckGo({backend}) から {len(results)}件 取得")
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
        suggests = fetch_google_suggests(keyword)
        if suggests:
            print(f"    -> {len(suggests)}件のサジェストを取得しました")
        else:
            print(f"    -> サジェストが見つかりませんでした")
    except Exception as e:
        print(f"    -> サジェスト取得に失敗しました: {e}")
        suggests = []

    # サジェストを改行区切りのテキストにまとめる
    suggests_text = "\n".join(suggests) if suggests else "（取得なし）"

    time.sleep(1)  # サジェスト取得後に少し待機

    # --- Step 2: DuckDuckGo検索（解決策・競合情報） ---
    print(f"  [Step 2] DuckDuckGoで検索中 (backend={SEARCH_BACKEND})...")
    try:
        search_results = search_duckduckgo(keyword)
        if search_results:
            print(f"    -> {len(search_results)}件ヒットしました")
        else:
            print(f"    -> 0件ヒット：backend={SEARCH_BACKEND} で結果が返りませんでした")
    except Exception as e:
        print(f"    -> 検索に失敗しました: {type(e).__name__}: {e}")
        search_results = []

    # --- データ統合 ---
    if search_results:
        for i, result in enumerate(search_results):
            rows.append(
                {
                    "キーワード": keyword,
                    "サジェスト（悩み・ニーズ）": suggests_text if i == 0 else "",
                    "検索順位": i + 1,
                    "タイトル": result["title"],
                    "URL": result["url"],
                    "スニペット（本文要約）": result["snippet"],
                }
            )
    else:
        # 検索結果が0件の場合、使用したバックエンド情報を含めて記録する
        rows.append(
            {
                "キーワード": keyword,
                "サジェスト（悩み・ニーズ）": suggests_text,
                "検索順位": "",
                "タイトル": f"（ヒットなし / backend={SEARCH_BACKEND}）",
                "URL": "",
                "スニペット（本文要約）": "",
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
    filename = f"{timestamp}_market_report.xlsx"

    # openpyxlエンジンを使用してExcel出力
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="調査結果")

        # 列幅を見やすく調整
        worksheet = writer.sheets["調査結果"]
        column_widths = {
            "A": 25,  # キーワード
            "B": 40,  # サジェスト
            "C": 10,  # 検索順位
            "D": 50,  # タイトル
            "E": 50,  # URL
            "F": 60,  # スニペット
        }
        for col_letter, width in column_widths.items():
            worksheet.column_dimensions[col_letter].width = width

    return filename


# ============================================================
# メイン処理
# ============================================================


def main():
    print("=" * 60)
    print(" Market Radar - 市場調査自動化ツール")
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
            print(f"  -> キーワード「{keyword}」の処理中に予期しないエラーが発生しました: {e}")
            all_rows.append(
                {
                    "キーワード": keyword,
                    "サジェスト（悩み・ニーズ）": "（エラー）",
                    "検索順位": "",
                    "タイトル": f"エラー: {e}",
                    "URL": "",
                    "スニペット（本文要約）": "",
                }
            )

        # 最後のキーワード以外は待機する（サーバー負荷軽減）
        if i < len(KEYWORDS):
            print(f"  -> {SLEEP_SECONDS}秒待機中...")
            time.sleep(SLEEP_SECONDS)

    # --- DataFrame作成・Excel出力 ---
    print(f"\n{'=' * 60}")
    print(" データ集計・Excel出力中...")

    df = pd.DataFrame(all_rows)
    filename = save_to_excel(df)

    print(f" レポートを保存しました: {filename}")
    print(f" 合計 {len(df)} 行のデータを出力しました")
    print("=" * 60)
    print(" 完了！")


if __name__ == "__main__":
    main()
