#!/usr/bin/env python3
"""
ホテル口コミ CSV データ分析スクリプト

Google Spreadsheet からエクスポートされた口コミ CSV を読み込み、
サイト別統計、評価分布、10pt 換算などを算出して JSON で出力する。

使い方:
  python3 analyze_reviews.py --csv data.csv --start-month 2026-02 --end-month 2026-03 --output analysis.json
"""

import argparse
import csv
import json
import re
import statistics
import sys
from collections import Counter, defaultdict
from datetime import datetime


# === 評価スケール変換ルール ===
SCALE_10_SITES = {"Booking.com", "Trip.com", "Agoda"}
SCALE_5_SITES = {"じゃらん", "楽天トラベル", "Google", "一休.com"}

# ヘッダー名の候補（柔軟なカラム検出用）
HEADER_PATTERNS = {
    "site": ["サイト名", "サイト", "site", "OTA", "予約サイト", "レビューサイト"],
    "rating": ["評価", "スコア", "rating", "score", "点数", "評点"],
    "date": ["投稿日", "日付", "date", "レビュー日", "投稿日時"],
    "comment_all": ["コメント(全体)", "コメント（全体）", "全体コメント", "コメント", "comment", "レビュー"],
    "comment_good": ["コメント(良い点)", "コメント（良い点）", "良い点", "ポジティブ", "good"],
    "comment_bad": ["コメント(改善点)", "コメント（改善点）", "改善点", "ネガティブ", "bad"],
    "translated_all": ["翻訳:全体", "翻訳（全体）", "翻訳", "translation"],
    "translated_good": ["翻訳:良い点", "翻訳（良い点）"],
    "translated_bad": ["翻訳:改善点", "翻訳（改善点）"],
    "reviewer": ["投稿者", "レビュアー", "reviewer", "名前"],
}


def find_column(headers, field_name):
    """ヘッダー名ベースで柔軟にカラムを検出"""
    patterns = HEADER_PATTERNS.get(field_name, [])
    for i, h in enumerate(headers):
        h_clean = h.strip()
        for pattern in patterns:
            if pattern.lower() in h_clean.lower():
                return i
    return None


def normalize_site_name(site):
    """サイト名の正規化"""
    site = site.strip()
    mappings = {
        "booking": "Booking.com",
        "booking.com": "Booking.com",
        "trip": "Trip.com",
        "trip.com": "Trip.com",
        "agoda": "Agoda",
        "じゃらん": "じゃらん",
        "jalan": "じゃらん",
        "楽天": "楽天トラベル",
        "楽天トラベル": "楽天トラベル",
        "rakuten": "楽天トラベル",
        "google": "Google",
        "一休": "一休.com",
        "一休.com": "一休.com",
        "ikyu": "一休.com",
    }
    for key, value in mappings.items():
        if key.lower() == site.lower() or key.lower() in site.lower():
            return value
    return site


def get_scale(site):
    """サイト名から評価スケールを判定"""
    if site in SCALE_10_SITES:
        return 10
    elif site in SCALE_5_SITES:
        return 5
    return None


def to_10pt(rating, site):
    """10pt 換算"""
    scale = get_scale(site)
    if scale == 5:
        return rating * 2
    return rating


def parse_date(date_str):
    """日付文字列をパース（複数フォーマット対応）"""
    date_str = date_str.strip()
    formats = [
        "%Y/%m/%d", "%Y-%m-%d", "%Y年%m月%d日",
        "%m/%d/%Y", "%d/%m/%Y",
        "%Y/%m/%d %H:%M", "%Y-%m-%d %H:%M:%S",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    # 部分マッチ: 年月だけ
    m = re.search(r"(\d{4})[/-](\d{1,2})", date_str)
    if m:
        return datetime(int(m.group(1)), int(m.group(2)), 1)
    return None


def get_judgment(avg_10pt):
    """10pt 換算平均から判定文字列を返す"""
    if avg_10pt >= 9.0:
        return "優秀"
    elif avg_10pt >= 8.0:
        return "良好"
    elif avg_10pt >= 7.0:
        return "概ね良好"
    else:
        return "要改善"


def main():
    parser = argparse.ArgumentParser(description="ホテル口コミ CSV 分析")
    parser.add_argument("--csv", required=True, help="CSV ファイルパス")
    parser.add_argument("--start-month", required=True, help="開始月 (YYYY-MM)")
    parser.add_argument("--end-month", required=True, help="終了月 (YYYY-MM)")
    parser.add_argument("--output", default="analysis.json", help="出力 JSON ファイル")
    args = parser.parse_args()

    # 期間パース
    start_year, start_month = map(int, args.start_month.split("-"))
    end_year, end_month = map(int, args.end_month.split("-"))

    # CSV 読み込み
    rows = []
    with open(args.csv, "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        all_rows = list(reader)

    # ヘッダー行を検出（「サイト」「評価」を含む行を探す）
    header_row_idx = None
    headers = []
    for i, row in enumerate(all_rows):
        row_text = " ".join(row).lower()
        if any(k in row_text for k in ["サイト", "評価", "site", "rating"]):
            header_row_idx = i
            headers = row
            break

    if header_row_idx is None:
        print("ERROR: ヘッダー行が見つかりません", file=sys.stderr)
        sys.exit(1)

    # カラムマッピング
    col_site = find_column(headers, "site")
    col_rating = find_column(headers, "rating")
    col_date = find_column(headers, "date")
    col_comment_all = find_column(headers, "comment_all")
    col_comment_good = find_column(headers, "comment_good")
    col_comment_bad = find_column(headers, "comment_bad")
    col_translated_all = find_column(headers, "translated_all")
    col_translated_good = find_column(headers, "translated_good")
    col_translated_bad = find_column(headers, "translated_bad")
    col_reviewer = find_column(headers, "reviewer")

    if col_site is None or col_rating is None:
        print(f"ERROR: 必須カラム（サイト名, 評価）が見つかりません。ヘッダー: {headers}", file=sys.stderr)
        sys.exit(1)

    print(f"ヘッダー行: {header_row_idx + 1}")
    print(f"カラムマッピング: site={col_site}, rating={col_rating}, date={col_date}")

    # カラムオフセット検証: 最初のデータ行でサイト列に "TRUE"/"FALSE" が入っていたら
    # ヘッダーとデータが1列ずれている（Google Sheetsのセル結合由来）ので補正する
    KNOWN_SITES_LOWER = {"booking.com", "trip.com", "agoda", "じゃらん", "楽天トラベル",
                         "楽天", "google", "一休.com", "一休", "jalan", "rakuten", "booking", "trip"}
    for check_row in all_rows[header_row_idx + 1:]:
        if len(check_row) <= max(col_site, col_rating):
            continue
        site_val = check_row[col_site].strip().lower()
        if not site_val:
            continue
        if site_val in ("true", "false") or site_val not in KNOWN_SITES_LOWER:
            # サイト列に想定外の値 → 1列右にずらして再チェック
            shifted_site = check_row[col_site + 1].strip().lower() if col_site + 1 < len(check_row) else ""
            if shifted_site in KNOWN_SITES_LOWER:
                print(f"カラムオフセット補正: 全カラムを+1シフト ('{check_row[col_site]}' は有効なサイト名でない)")
                col_site += 1
                col_rating = col_rating + 1 if col_rating is not None else None
                col_date = col_date + 1 if col_date is not None else None
                col_comment_all = col_comment_all + 1 if col_comment_all is not None else None
                col_comment_good = col_comment_good + 1 if col_comment_good is not None else None
                col_comment_bad = col_comment_bad + 1 if col_comment_bad is not None else None
                col_translated_all = col_translated_all + 1 if col_translated_all is not None else None
                col_translated_good = col_translated_good + 1 if col_translated_good is not None else None
                col_translated_bad = col_translated_bad + 1 if col_translated_bad is not None else None
                col_reviewer = col_reviewer + 1 if col_reviewer is not None else None
                print(f"補正後カラムマッピング: site={col_site}, rating={col_rating}, date={col_date}")
        break  # 最初の有効行だけチェック

    # データ行をパース
    reviews = []
    seen = set()  # 重複検出用

    for row in all_rows[header_row_idx + 1:]:
        if len(row) <= max(col_site, col_rating):
            continue

        site_raw = row[col_site].strip()
        rating_raw = row[col_rating].strip()

        if not site_raw or not rating_raw:
            continue

        # サイト名正規化
        site = normalize_site_name(site_raw)

        # 評価値パース
        try:
            rating = float(rating_raw)
        except ValueError:
            continue

        # 日付フィルタ
        date_str = row[col_date].strip() if col_date is not None and col_date < len(row) else ""
        date_obj = parse_date(date_str) if date_str else None

        if date_obj:
            in_range = (
                (date_obj.year > start_year or (date_obj.year == start_year and date_obj.month >= start_month))
                and
                (date_obj.year < end_year or (date_obj.year == end_year and date_obj.month <= end_month))
            )
            if not in_range:
                continue

        # コメント取得
        def safe_get(idx):
            return row[idx].strip() if idx is not None and idx < len(row) else ""

        comment_all = safe_get(col_comment_all)
        comment_good = safe_get(col_comment_good)
        comment_bad = safe_get(col_comment_bad)
        translated_all = safe_get(col_translated_all)
        translated_good = safe_get(col_translated_good)
        translated_bad = safe_get(col_translated_bad)
        reviewer = safe_get(col_reviewer)

        # 重複チェック（サイト+評価+日付+コメント先頭30文字）
        dedup_key = f"{site}|{rating}|{date_str[:10]}|{comment_all[:30]}"
        if dedup_key in seen:
            continue
        seen.add(dedup_key)

        reviews.append({
            "site": site,
            "rating": rating,
            "date": date_str,
            "comment": comment_all,
            "good": comment_good,
            "bad": comment_bad,
            "translated": translated_all,
            "translated_good": translated_good,
            "translated_bad": translated_bad,
            "reviewer": reviewer,
        })

    if not reviews:
        print("WARNING: 対象期間内のレビューが0件です", file=sys.stderr)

    print(f"対象レビュー数: {len(reviews)}件")

    # === サイト別統計 ===
    site_groups = defaultdict(list)
    for r in reviews:
        site_groups[r["site"]].append(r["rating"])

    site_stats = []
    all_10pt = []

    for site, ratings in sorted(site_groups.items()):
        count = len(ratings)
        native_avg = round(statistics.mean(ratings), 2)
        scale = get_scale(site)

        ratings_10pt = [to_10pt(r, site) for r in ratings]
        avg_10pt = round(statistics.mean(ratings_10pt), 2)
        median_10pt = round(statistics.median(ratings_10pt), 1)

        all_10pt.extend(ratings_10pt)

        site_stats.append({
            "site": site,
            "count": count,
            "native_avg": native_avg,
            "scale": f"/{scale}" if scale else "unknown",
            "avg_10pt": avg_10pt,
            "median_10pt": median_10pt,
            "judgment": get_judgment(avg_10pt),
        })

    # 10pt換算で降順ソート
    site_stats.sort(key=lambda x: x["avg_10pt"], reverse=True)

    # === 全体統計 ===
    overall_avg_10pt = round(statistics.mean(all_10pt), 2) if all_10pt else 0
    total_reviews = len(reviews)

    # === 評価分布（10pt換算） ===
    dist_counter = Counter()
    for val in all_10pt:
        score = int(round(val))
        dist_counter[score] += 1

    distribution = []
    for score in sorted(dist_counter.keys(), reverse=True):
        count = dist_counter[score]
        pct = round(count / total_reviews * 100, 1) if total_reviews > 0 else 0
        distribution.append({
            "score": score,
            "count": count,
            "pct": f"{pct}%",
        })

    # === High / Mid / Low ===
    high_count = sum(1 for v in all_10pt if v >= 8)
    mid_count = sum(1 for v in all_10pt if 5 <= v < 8)
    low_count = sum(1 for v in all_10pt if v < 5)

    high_rate = round(high_count / total_reviews * 100, 1) if total_reviews > 0 else 0
    mid_rate = round(mid_count / total_reviews * 100, 1) if total_reviews > 0 else 0
    low_rate = round(low_count / total_reviews * 100, 1) if total_reviews > 0 else 0

    # === コメント一覧 ===
    comments = []
    for r in reviews:
        comments.append({
            "site": r["site"],
            "rating": r["rating"],
            "rating_10pt": to_10pt(r["rating"], r["site"]),
            "date": r["date"],
            "comment": r["comment"],
            "good": r["good"],
            "bad": r["bad"],
            "translated": r["translated"],
            "translated_good": r["translated_good"],
            "translated_bad": r["translated_bad"],
            "reviewer": r["reviewer"],
        })

    # === 結果JSON ===
    result = {
        "total_reviews": total_reviews,
        "overall_avg_10pt": overall_avg_10pt,
        "high_count": high_count,
        "high_rate": high_rate,
        "mid_count": mid_count,
        "mid_rate": mid_rate,
        "low_count": low_count,
        "low_rate": low_rate,
        "site_stats": site_stats,
        "distribution": distribution,
        "comments": comments,
    }

    # 出力
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"\n=== 分析結果サマリー ===")
    print(f"総レビュー数: {total_reviews}件")
    print(f"全体平均(10pt換算): {overall_avg_10pt}")
    print(f"高評価率(8-10点): {high_rate}% ({high_count}件)")
    print(f"中評価率(5-7点): {mid_rate}% ({mid_count}件)")
    print(f"低評価率(1-4点): {low_rate}% ({low_count}件)")
    print(f"\nサイト別:")
    for s in site_stats:
        print(f"  {s['site']}: {s['count']}件, ネイティブ平均{s['native_avg']}{s['scale']}, 10pt={s['avg_10pt']} [{s['judgment']}]")
    print(f"\n結果を {args.output} に保存しました。")


if __name__ == "__main__":
    main()
