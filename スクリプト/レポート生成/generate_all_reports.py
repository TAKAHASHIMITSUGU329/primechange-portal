#!/usr/bin/env python3
"""Generate DOCX and PPTX reports for 19 hotels using analysis JSONs and JS templates."""

import json
import os
import re
import subprocess
import tempfile
import shutil

BASE = "/home/user/primechange-portal"
JSON_DIR = BASE + "/データ/分析結果JSON"
OUT_DIR = BASE + "/納品レポート/ホテル別レポート"
DOCX_TEMPLATE = BASE + "/.claude/skills/hotel-review-report/assets/docx_template.js"
PPTX_TEMPLATE = BASE + "/.claude/skills/hotel-review-report/assets/pptx_template.js"
NODE_MODULES = BASE + "/node_modules"
from datetime import date as _date
_today = _date.today()
ANALYSIS_PERIOD = f"2025年12月〜{_today.year}年{_today.month}月"
REPORT_DATE = f"{_today.year}年{_today.month}月{_today.day}日"

HOTELS = {
    "daiwa_osaki": "ダイワロイネットホテル東京大崎",
    "chisan": "チサンホテル浜松町",
    "hearton": "ハートンホテル東品川",
    "keyakigate": "ホテルケヤキゲート東京府中",
    "richmond_mejiro": "リッチモンドホテル東京目白",
    "keisei_kinshicho": "京成リッチモンドホテル東京錦糸町",
    "daiichi_ikebukuro": "第一イン池袋",
    "comfort_roppongi": "コンフォートイン六本木",
    "comfort_suites_tokyobay": "コンフォートスイーツ東京ベイ",
    "comfort_era_higashikanda": "コンフォートホテルERA東京東神田",
    "comfort_yokohama_kannai": "コンフォートホテル横浜関内",
    "comfort_narita": "コンフォートホテル成田",
    "apa_kamata": "アパホテル蒲田駅東",
    "apa_sagamihara": "アパホテル相模原橋本駅東",
    "court_shinyokohama": "コートホテル新横浜",
    "comment_yokohama": "ホテルコメント横浜関内",
    "kawasaki_nikko": "川崎日航ホテル",
    "henn_na_haneda": "変なホテル東京羽田",
    "comfort_hakata": "コンフォートホテル博多",
}

# ─────────────────────────────────────────────
# Comment analysis helpers
# ─────────────────────────────────────────────

STRENGTH_THEMES = [
    ("立地・アクセス",      ["立地", "アクセス", "駅", "便利", "近い", "交通"]),
    ("清潔感",             ["清潔", "きれい", "キレイ", "綺麗"]),
    ("スタッフ対応",        ["スタッフ", "従業員", "対応", "親切", "丁寧", "フロント"]),
    ("部屋・設備",         ["部屋", "設備", "快適", "アメニティ", "広い", "ベッド", "シャワー"]),
    ("朝食",               ["朝食", "朝ごはん", "食事", "breakfast"]),
    ("コスパ・リピート",    ["コスパ", "価格", "また来", "リピート", "満足", "お得"]),
]

WEAKNESS_MAP = [
    ("部屋狭小感",  ["狭い", "狭小", "手狭"]),
    ("清潔感",      ["臭い", "汚い", "汚れ", "カビ"]),
    ("通信環境",    ["Wi-Fi", "wifi", "WiFi", "インターネット", "通信"]),
    ("騒音問題",    ["うるさい", "騒音", "音", "壁薄"]),
    ("朝食品質",    ["朝食", "朝ごはん", "食事", "breakfast"]),
    ("エレベーター待ち", ["エレベーター", "エレベーター待ち"]),
    ("水回り",      ["水回り", "水", "シャワー", "お湯", "排水"]),
    ("設備老朽化",  ["古い", "老朽", "傷", "壊れ", "故障"]),
    ("駐車場・立地", ["駐車場", "駐車", "立地", "アクセス"]),
]

PRIORITY_THRESHOLDS = {"S": 3, "A": 2, "B": 1, "C": 0}


def get_all_positive_text(comments):
    texts = []
    for c in comments:
        for field in ("comment", "good", "translated", "translated_good"):
            val = c.get(field, "")
            if val:
                texts.append(val)
    return texts


def get_all_negative_text(comments):
    texts = []
    for c in comments:
        for field in ("bad", "translated_bad"):
            val = c.get(field, "")
            if val:
                texts.append(val)
    return texts


def analyze_strengths(comments):
    pos_texts = get_all_positive_text(comments)
    results = []
    for theme_name, keywords in STRENGTH_THEMES:
        matching = []
        for text in pos_texts:
            if any(kw.lower() in text.lower() for kw in keywords):
                matching.append(text)
        count = len(matching)
        quotes = []
        for t in matching[:5]:
            # Find a sentence with a keyword
            for kw in keywords:
                if kw.lower() in t.lower():
                    # Extract a short segment
                    idx = t.lower().find(kw.lower())
                    start = max(0, idx - 5)
                    end = min(len(t), idx + 25)
                    snippet = t[start:end].strip()
                    snippet = re.sub(r'[\r\n]+', ' ', snippet)
                    if len(snippet) > 2:
                        quotes.append(snippet[:30])
                    break
            if len(quotes) >= 2:
                break
        results.append({
            "theme": theme_name,
            "count": count,
            "quotes": quotes[:2],
        })
    # Sort by count descending, pick top 6
    results.sort(key=lambda x: x["count"], reverse=True)
    return results[:6]


def analyze_weaknesses(comments):
    neg_texts = get_all_negative_text(comments)
    results = []
    for cat_name, keywords in WEAKNESS_MAP:
        matching = []
        for text in neg_texts:
            if any(kw.lower() in text.lower() for kw in keywords):
                matching.append(text)
        count = len(matching)
        if count > 0:
            # Get example detail
            detail = ""
            for t in matching[:1]:
                detail = re.sub(r'[\r\n]+', ' ', t)[:50]
            results.append({"cat": cat_name, "count": count, "detail": detail})
    results.sort(key=lambda x: x["count"], reverse=True)
    # Assign priority
    for i, r in enumerate(results):
        if i == 0 and r["count"] >= PRIORITY_THRESHOLDS["S"]:
            r["priority"] = "S"
            r["color"] = "DC2626"
        elif i <= 1 and r["count"] >= PRIORITY_THRESHOLDS["A"]:
            r["priority"] = "A"
            r["color"] = "EA580C"
        elif r["count"] >= PRIORITY_THRESHOLDS["B"]:
            r["priority"] = "B"
            r["color"] = "D4A843"
        else:
            r["priority"] = "C"
            r["color"] = "94A3B8"
    # Ensure at least some weaknesses even if few negative comments
    if len(results) < 3:
        defaults = [
            {"cat": "口コミ返信率", "count": 0, "detail": "口コミへの返信対応を強化する余地あり", "priority": "B", "color": "D4A843"},
            {"cat": "サービス改善", "count": 0, "detail": "継続的なサービス品質向上が必要", "priority": "C", "color": "94A3B8"},
            {"cat": "設備メンテナンス", "count": 0, "detail": "定期的なメンテナンス体制の整備", "priority": "C", "color": "94A3B8"},
        ]
        for d in defaults:
            if len(results) >= 6:
                break
            if not any(r["cat"] == d["cat"] for r in results):
                results.append(d)
    return results[:9]


def get_representative_quote(comments, keywords):
    for c in comments:
        for field in ("translated", "translated_good", "comment", "good"):
            text = c.get(field, "")
            if text and any(kw.lower() in text.lower() for kw in keywords):
                # Return a cleaned short snippet
                clean = re.sub(r'[\r\n]+', ' ', text).strip()
                return clean[:25]
    return ""


def js_str(s):
    """Escape a string for JS double-quote context."""
    s = str(s)
    s = s.replace("\\", "\\\\")
    s = s.replace('"', '\\"')
    s = s.replace("\n", "\\n")
    s = s.replace("\r", "")
    return s


def format_site_data_js(site_stats):
    lines = []
    for s in site_stats:
        site = js_str(s["site"])
        count = str(s["count"])
        native = f'{s["native_avg"]:.2f}'
        scale = js_str(s["scale"])
        avg10 = f'{s["avg_10pt"]:.2f}'
        median = f'{s["median_10pt"]:.1f}'
        judgment = js_str(s["judgment"])
        lines.append(f'  ["{site}", "{count}", "{native}", "{scale}", "{avg10}", "{median}", "{judgment}"]')
    return "[\n" + ",\n".join(lines) + "\n]"


def format_distribution_data_js(distribution):
    lines = []
    for d in distribution:
        score = d["score"]
        count = d["count"]
        pct = js_str(d["pct"])
        lines.append(f'  [{score}, {count}, "{pct}"]')
    return "[\n" + ",\n".join(lines) + "\n]"


def format_strength_themes_js(strength_analysis):
    lines = []
    for s in strength_analysis:
        theme = js_str(s["theme"])
        count = str(s["count"])
        quotes_str = "".join(f'「{js_str(q)}」' for q in s["quotes"]) if s["quotes"] else "口コミより"
        lines.append(f'  ["{theme}", "{count}件", "{js_str(quotes_str)}"]')
    return "[\n" + ",\n".join(lines) + "\n]"


def format_weakness_priority_js(weaknesses):
    lines = []
    for w in weaknesses:
        pri = w["priority"]
        cat = js_str(w["cat"])
        detail = js_str(w.get("detail", ""))
        count = w["count"]
        lines.append(f'  ["{pri}", "{cat}", "{detail}", "影響度・{count}件"]')
    return "[\n" + ",\n".join(lines) + "\n]"


def generate_improvement_actions(weaknesses, strength_analysis):
    """Generate Phase 1/2/3 improvement actions based on top weaknesses."""
    top_issues = [w["cat"] for w in weaknesses[:4]]

    action_map = {
        "部屋狭小感": {
            "phase1": ("部屋利用満足度の最大化", ["収納スペースの整理・最適化", "荷物預かりサービスの案内強化", "部屋タイプ別の使い勝手改善ガイド作成", "コンパクトな部屋の魅力を訴求するOTA説明文改訂"]),
            "phase2_title": "部屋改装計画の検討",
            "phase2_items": ["上位グレード部屋への誘導施策導入", "スペース効率化家具への入替計画立案", "客室タイプ構成の見直し"],
        },
        "清潔感": {
            "phase1": ("清潔感向上の即時対応", ["客室清掃チェックリストの強化・再徹底", "気になる箇所の即時補修対応フロー整備", "消臭・換気プロトコルの見直し", "清潔感評価の定期モニタリング開始"]),
            "phase2_title": "設備清潔維持の中期施策",
            "phase2_items": ["定期的な深掃除スケジュールの策定", "カーペット・カーテン等の計画的な交換", "清掃品質向上のためのスタッフ研修"],
        },
        "通信環境": {
            "phase1": ("Wi-Fi品質の即時改善", ["Wi-Fi接続手順の案内改善（館内掲示・サイト明記）", "接続不具合時の対応フロー整備", "フロントスタッフへのWi-Fi対応スクリプト共有", "速度測定と問題箇所の特定"]),
            "phase2_title": "通信インフラ強化",
            "phase2_items": ["Wi-Fiルーター増設・アップグレード計画", "帯域幅の拡大検討", "5G対応設備への移行ロードマップ策定"],
        },
        "騒音問題": {
            "phase1": ("騒音対策の即時対応", ["防音対策済み部屋の積極的案内", "チェックイン時の防音耳栓提供", "騒音クレーム対応マニュアルの整備", "防音性能の高い部屋への優先アサイン"]),
            "phase2_title": "防音設備の強化",
            "phase2_items": ["防音カーテン・防音パネルの設置計画", "騒音の多い部屋の用途転換検討", "建物外部騒音対策の工事計画"],
        },
        "朝食品質": {
            "phase1": ("朝食品質の即時改善", ["朝食メニューへのフィードバック収集と改善", "食材品質の見直しと産地表示の充実", "混雑緩和のための時間帯分散対策", "朝食スタッフへのホスピタリティ研修"]),
            "phase2_title": "朝食体験の抜本的改善",
            "phase2_items": ["メニュー拡充と季節限定品の導入", "朝食会場レイアウト改善", "地産地消食材の導入による差別化"],
        },
        "エレベーター待ち": {
            "phase1": ("エレベーター待ち時間の改善", ["ピーク時間帯の案内と階段利用促進", "チェックアウト時間分散の促進", "エレベーター近くの待機スペース改善", "待ち時間の見える化掲示"]),
            "phase2_title": "エレベーター運用改善",
            "phase2_items": ["エレベーター制御プログラムの最適化", "増設可否の建物調査実施", "荷物専用エレベーターの時間帯設定"],
        },
        "水回り": {
            "phase1": ("水回り不具合の即時対応", ["水回り点検チェックリストの強化", "排水・給水トラブルの優先対応フロー", "浴室清潔度の定期巡回チェック強化", "不具合報告専用の内線番号周知"]),
            "phase2_title": "水回り設備の計画的更新",
            "phase2_items": ["老朽化した給排水設備の交換計画", "シャワーヘッド・蛇口等の節水型設備導入", "定期メンテナンス契約の見直し"],
        },
        "設備老朽化": {
            "phase1": ("設備老朽化への即時対応", ["緊急修繕リストの作成と優先対応", "老朽化設備の在庫把握と交換計画", "スタッフによる日常点検の徹底", "故障時のバックアップ体制整備"]),
            "phase2_title": "設備更新の計画的実施",
            "phase2_items": ["中長期設備更新ロードマップの策定", "高頻度使用備品の計画的交換サイクル確立", "省エネ設備への転換による維持費削減"],
        },
    }

    default_phase1 = ("サービス品質全般の向上", ["口コミへの返信率100%達成（48時間以内）", "チェックイン・チェックアウト対応の効率化", "フロントスタッフへのCS向上研修実施", "お客様満足度アンケートの配布と集計体制整備"])
    default_phase2_title = "顧客体験向上の短期施策"
    default_phase2_items = ["OTA写真・説明文の定期的な見直しと更新", "リピーター向け優待プログラムの検討", "多言語対応サービスの充実（英語・中国語）"]

    phase1_items = []
    phase2_items = []
    phase3_items = []

    used = set()
    for issue in top_issues:
        if issue in action_map and issue not in used:
            info = action_map[issue]
            title, bullets = info["phase1"]
            phase1_items.append({"title": f"({len(phase1_items)+1}) {title}", "bullets": bullets})
            used.add(issue)
            if len(phase1_items) >= 4:
                break

    if not phase1_items:
        phase1_items.append({"title": f"(1) {default_phase1[0]}", "bullets": list(default_phase1[1])})

    # Phase 1 extras
    phase1_items.append({"title": f"({len(phase1_items)+1}) 口コミ返信率100%の実現", "bullets": [
        "全OTAサイトの口コミへの返信ルール策定（48時間以内）",
        "ネガティブ口コミへの丁寧な対応テンプレート作成",
        "フロントマネージャーによる週次口コミモニタリング",
    ]})

    # Phase 2
    for issue in top_issues[:3]:
        if issue in action_map:
            info = action_map[issue]
            phase2_items.append({"title": info["phase2_title"], "items": info["phase2_items"]})

    if not phase2_items:
        phase2_items.append({"title": default_phase2_title, "items": default_phase2_items})

    phase2_items.append({"title": "OTA評価スコア向上施策", "items": [
        "高評価レビュー投稿の依頼タイミング最適化（チェックアウト時）",
        "満足度の高い宿泊者へのフォローアップメール送信",
        "じゃらん・楽天の評価向上に向けた特典プログラム検討",
    ]})

    # Phase 3: Based on top strengths to double down + top weaknesses for structural fix
    top_strength = strength_analysis[0]["theme"] if strength_analysis else "立地・アクセス"
    phase3_items = [
        {"title": f"「{top_strength}」の競争優位性強化", "items": [
            f"{top_strength}を訴求したブランディング戦略の策定",
            "SNS・口コミサイトでの強み積極的発信",
            "旅行代理店・法人営業への強み訴求資料作成",
        ]},
        {"title": "客室リノベーション計画", "items": [
            "客室設備の全面的な見直しと更新計画策定",
            "最新のスマートホテル機能の導入検討",
            "環境に配慮したサステナブル客室改装",
        ]},
        {"title": "デジタル接客・サービス高度化", "items": [
            "AIチャットボットによる24時間問合せ対応",
            "モバイルチェックイン/アウトシステムの導入",
            "顧客データ分析によるパーソナライズサービスの実現",
        ]},
    ]

    return phase1_items, phase2_items, phase3_items


def build_kpi_targets(data):
    avg = data["overall_avg_10pt"]
    high_rate = data["high_rate"]
    low_rate = data["low_rate"]
    total = data["total_reviews"]

    target_avg = min(9.5, round(avg + 0.3, 1))
    target_high = min(95, round(high_rate + 5))
    target_low = max(0, round(low_rate - 1, 1))

    rows = [
        ["全体平均(10pt換算)", f"{avg:.2f}点", f"{target_avg}点以上", "2026年10月"],
        ["高評価率（8-10点）", f"{high_rate:.1f}%", f"{target_high}%以上", "2026年10月"],
        ["低評価率（1-4点）", f"{low_rate:.1f}%", f"{target_low}%以下", "2026年10月"],
        ["口コミ返信率", "未計測", "100%（48h以内）", "2026年7月"],
    ]

    # Add top site targets
    for s in data.get("site_stats", [])[:2]:
        site = s["site"]
        current = s["native_avg"]
        scale = s["scale"]
        target_v = round(current + 0.2, 1)
        if scale == "/5":
            rows.append([f"{site}評価", f"{current:.2f}/5点", f"{target_v}/5点以上", "2026年10月"])
        else:
            rows.append([f"{site}評価", f"{current:.2f}/10点", f"{target_v}/10点以上", "2026年10月"])

    return rows


def build_executive_summary(hotel_name, data, strength_analysis, weaknesses):
    total = data["total_reviews"]
    avg = data["overall_avg_10pt"]
    high_rate = data["high_rate"]
    low_rate = data["low_rate"]

    top_strengths = [s["theme"] for s in strength_analysis[:3] if s["count"] > 0]
    top_weaknesses = [w["cat"] for w in weaknesses[:3] if w["count"] > 0]

    intro = f"{ANALYSIS_PERIOD}に各OTAサイト・口コミサイトに投稿された{total}件のレビューを包括的に分析しました。以下が主要な発見事項です。"

    eval1 = f"{hotel_name}の総合評価は10点換算で{avg:.2f}点となり、高評価率は{high_rate:.1f}%です。"
    if avg >= 8.5:
        eval1 += f"全体的に高い評価を獲得しており、特に{top_strengths[0] if top_strengths else 'サービス全般'}が顧客に高く評価されています。"
    elif avg >= 7.5:
        eval1 += f"概ね良好な評価を受けており、{top_strengths[0] if top_strengths else 'サービス全般'}への好評が多く見られます。"
    else:
        eval1 += f"改善の余地があり、重点的な品質向上が求められます。"

    eval2 = f"一方、"
    if top_weaknesses:
        eval2 += f"{top_weaknesses[0]}など改善すべき課題も確認されました。本レポートの改善施策を実施することで、さらなる評価向上が期待できます。"
    else:
        eval2 += "特段の問題は見られませんが、継続的な品質向上と口コミ管理の強化により、さらに高い評価を目指すことが重要です。"

    strength_str = ""
    if top_strengths:
        parts = [f"{s}（{strength_analysis[i]['count']}件で言及）" for i, s in enumerate(top_strengths)]
        strength_str = "、".join(parts)
    else:
        strength_str = "サービス全般への好評価"

    weakness_str = ""
    if top_weaknesses:
        wparts = []
        for w in weaknesses[:3]:
            if w["count"] > 0:
                wparts.append(f"{w['cat']}（{w['count']}件）")
        weakness_str = "、".join(wparts) if wparts else "改善余地あり"
    else:
        weakness_str = "継続的な品質管理が必要"

    opportunity = "リピート意向の高い顧客が多く、口コミ管理と体験価値の向上で高評価率をさらに伸ばせるポテンシャルがあります。"

    return intro, eval1, eval2, strength_str, weakness_str, opportunity



# ─────────────────────────────────────────────
# JS generation
# ─────────────────────────────────────────────

def generate_docx_js(hotel_name, data, strength_analysis, weaknesses):
    with open(DOCX_TEMPLATE, "r", encoding="utf-8") as f:
        template = f.read()

    total = data["total_reviews"]
    avg = data["overall_avg_10pt"]
    high_count = data["high_count"]
    high_rate = data["high_rate"]
    mid_count = data["mid_count"]
    mid_rate = data["mid_rate"]
    low_count = data["low_count"]
    low_rate = data["low_rate"]

    site_names = [s["site"] for s in data.get("site_stats", [])]
    sites_text = " / ".join(site_names) if site_names else "各OTAサイト"
    target_sites_str = f"対象サイト：{sites_text}"

    low_color = "C0392B" if low_count > 0 else "27AE60"

    intro, eval1, eval2, strength_str, weakness_str, opportunity = build_executive_summary(
        hotel_name, data, strength_analysis, weaknesses
    )

    # KPI cards
    kpi_low_color = "C0392B" if low_count > 0 else "27AE60"
    kpi_low_bg = "FDEDEC" if low_count > 0 else "E8F5E9"
    kpi_cards_js = f"""[
  {{ label: "全体平均(10pt換算)", value: "{avg:.2f}", color: "27AE60", bgColor: "E8F5E9" }},
  {{ label: "高評価率(8-10点)", value: "{high_rate:.1f}%", color: "27AE60", bgColor: "E8F5E9" }},
  {{ label: "低評価率(1-4点)", value: "{low_rate:.1f}%", color: "{kpi_low_color}", bgColor: "{kpi_low_bg}" }},
  {{ label: "レビュー総数", value: "{total}件", color: "1B3A5C", bgColor: "D5E8F0" }},
]"""

    dist_data_js = format_distribution_data_js(data.get("distribution", []))
    dist_counts = [d["count"] for d in data.get("distribution", [])]
    dist_max = max(dist_counts) if dist_counts else 1

    strength_js = format_strength_themes_js(strength_analysis)
    weakness_js = format_weakness_priority_js(weaknesses)

    phase1_items, phase2_items, phase3_items = generate_improvement_actions(weaknesses, strength_analysis)

    def format_phase_items_js(items):
        lines = []
        for item in items:
            title = js_str(item["title"])
            bullets_js = ",\n      ".join([f'"{js_str(b)}"' for b in item["bullets"]])
            lines.append(f"""  {{
    title: "{title}",
    bullets: [
      {bullets_js},
    ],
  }}""")
        return "[\n" + ",\n".join(lines) + "\n]"

    kpi_targets = build_kpi_targets(data)
    kpi_target_lines = []
    for row in kpi_targets:
        kpi_target_lines.append(f'  ["{js_str(row[0])}", "{js_str(row[1])}", "{js_str(row[2])}", "{js_str(row[3])}"]')
    kpi_target_js = "[\n" + ",\n".join(kpi_target_lines) + "\n]"

    # Strength subsection
    top_s = strength_analysis[0] if strength_analysis else {"theme": "立地・アクセス", "count": 0, "quotes": []}
    top_s2 = strength_analysis[1] if len(strength_analysis) > 1 else {"theme": "清潔感", "count": 0}

    strength_sub1_bullets = [
        f"複数のOTAサイトで{top_s['theme']}に関する高評価コメントが多数見られます",
        f"特に{strength_analysis[0]['quotes'][0] if strength_analysis[0]['quotes'] else '利便性の高さ'}が繰り返し言及されています",
        "当該強みを今後もOTA説明文・写真で積極的にアピールすることを推奨します",
        "この強みを軸にしたリピーター獲得施策の検討が有効です",
    ]

    # Conclusion paragraphs
    top_issue = weaknesses[0]["cat"] if weaknesses and weaknesses[0]["count"] > 0 else "サービス全般"
    conclusion_pars = [
        f"本分析では、{hotel_name}の{ANALYSIS_PERIOD}の口コミ{total}件を包括的に評価しました。総合評価{avg:.2f}点（10点換算）、高評価率{high_rate:.1f}%という結果から、当ホテルのサービス品質は概ね高い水準にあることが確認できます。",
        f"特に「{top_s['theme']}」が顧客満足度の主要なドライバーとなっており、この強みをOTAサイトでの集客施策に積極的に活用することが収益向上に直結します。",
        f"一方、{top_issue}に関する指摘が見られ、本レポートで提示した3フェーズの改善施策を実施することで、さらなる評価向上が期待できます。",
        f"Phase 1の即時対応施策から着実に実行し、6ヶ月以内に全体平均{min(9.5, round(avg+0.3,1))}点以上、高評価率{min(95, round(high_rate+5))}%以上を目指すことを推奨いたします。",
    ]
    conclusion_final = "プライムチェンジは、本レポートの施策実施において引き続きサポートいたします。ご質問・ご相談はいつでもお気軽にお問い合わせください。"

    site_data_js = format_site_data_js(data.get("site_stats", []))

    dist_text = f"{total}件のレビューにおける評価分布を10点換算で分析しました。高評価（8-10点）が{high_count}件（{high_rate:.1f}%）を占め、"
    if low_count == 0:
        dist_text += "低評価（1-4点）は0件と、良好な評価分布が維持されています。"
    else:
        dist_text += f"低評価（1-4点）が{low_count}件（{low_rate:.1f}%）となっています。"

    def esc(s): return js_str(s)

    js = template
    js = js.replace('const OUTPUT_DIR = "/Users/mitsugutakahashi/ホテル口コミ";',
                    f'const OUTPUT_DIR = "{OUT_DIR}";')
    js = js.replace('const HOTEL_NAME = "【ホテル名】";',
                    f'const HOTEL_NAME = "{esc(hotel_name)}";')
    js = js.replace('const ANALYSIS_PERIOD = "YYYY年M月〜M月";',
                    f'const ANALYSIS_PERIOD = "{ANALYSIS_PERIOD}";')
    js = js.replace('const REPORT_DATE = "YYYY年M月D日";',
                    f'const REPORT_DATE = "{REPORT_DATE}";')
    js = js.replace('const REVIEW_COUNT = "XX件（重複除外後）";',
                    f'const REVIEW_COUNT = "{total}件（重複除外後）";')
    js = js.replace('const TARGET_SITES = "対象サイト：Booking.com / Trip.com / じゃらん / 楽天トラベル / Agoda / Google";',
                    f'const TARGET_SITES = "{target_sites_str}";')
    js = js.replace(
        '''const KPI_CARDS = [
  { label: "全体平均(10pt換算)", value: "X.XX", color: "27AE60", bgColor: "E8F5E9" },
  { label: "高評価率(8-10点)", value: "XX.X%", color: "27AE60", bgColor: "E8F5E9" },
  { label: "低評価率(1-4点)", value: "X.X%", color: "27AE60", bgColor: "E8F5E9" },
  { label: "レビュー総数", value: "XX件", color: "1B3A5C", bgColor: "D5E8F0" },
];''',
        f'const KPI_CARDS = {kpi_cards_js};')
    js = js.replace(
        'const EXECUTIVE_SUMMARY_INTRO = "YYYY年M月〜M月に各OTAサイト・口コミサイトに投稿されたXX件のレビューを包括的に分析しました。以下が主要な発見事項です。";',
        f'const EXECUTIVE_SUMMARY_INTRO = "{esc(intro)}";')
    js = js.replace(
        'const EXECUTIVE_SUMMARY_EVALUATION_1 = "【総合評価の第1段落テキスト】";',
        f'const EXECUTIVE_SUMMARY_EVALUATION_1 = "{esc(eval1)}";')
    js = js.replace(
        'const EXECUTIVE_SUMMARY_EVALUATION_2 = "【総合評価の第2段落テキスト】";',
        f'const EXECUTIVE_SUMMARY_EVALUATION_2 = "{esc(eval2)}";')
    js = js.replace(
        'const KEY_FINDING_STRENGTH = "【Strength テキスト（例：立地の利便性（XX件で言及）、清潔感（XX件）、スタッフの親切さ（XX件））】";',
        f'const KEY_FINDING_STRENGTH = "{esc(strength_str)}";')
    js = js.replace(
        'const KEY_FINDING_WEAKNESS = "【Weakness テキスト（例：水回り不備（X件）、部屋狭小感（X件）、エレベーター待ち（X件））】";',
        f'const KEY_FINDING_WEAKNESS = "{esc(weakness_str)}";')
    js = js.replace(
        'const KEY_FINDING_OPPORTUNITY = "【Opportunity テキスト（例：リピート意向が高く、体験価値の向上で高評価率をさらに伸ばせるポテンシャル大）】";',
        f'const KEY_FINDING_OPPORTUNITY = "{esc(opportunity)}";')
    js = js.replace(
        '''const SITE_DATA = [
  ["サイト名1", "XX", "X.XX", "/10", "X.XX", "X.X", "優秀"],
  ["サイト名2", "XX", "X.XX", "/5",  "X.XX", "X.X", "良好"],
  // ... 必要な行数だけ追加
];''',
        f'const SITE_DATA = {site_data_js};')
    js = js.replace(
        '''const DISTRIBUTION_DATA = [
  [10, 0, "0.0%"],
  [9,  0, "0.0%"],
  [8,  0, "0.0%"],
  [7,  0, "0.0%"],
  [6,  0, "0.0%"],
  [5,  0, "0.0%"],
  // 4以下が必要なら追加
];''',
        f'const DISTRIBUTION_DATA = {dist_data_js};')
    js = js.replace(
        'const DISTRIBUTION_MAX_COUNT = 1; // 分布バーの最大件数（最多の件数を設定）',
        f'const DISTRIBUTION_MAX_COUNT = {dist_max};')
    js = js.replace(
        'const HIGH_RATING_SUMMARY = "XX件（XX.X%）";',
        f'const HIGH_RATING_SUMMARY = "{high_count}件（{high_rate:.1f}%）";')
    js = js.replace(
        'const MID_RATING_SUMMARY = "XX件（XX.X%）";',
        f'const MID_RATING_SUMMARY = "{mid_count}件（{mid_rate:.1f}%）";')
    js = js.replace(
        'const LOW_RATING_SUMMARY = "X件（X.X%）";',
        f'const LOW_RATING_SUMMARY = "{low_count}件（{low_rate:.1f}%）";')
    js = js.replace(
        'const LOW_RATING_COLOR = "27AE60"; // 低評価0件なら緑、あれば "C0392B"(赤) に変更',
        f'const LOW_RATING_COLOR = "{low_color}";')
    js = js.replace(
        'const DATA_OVERVIEW_SITE_TEXT = "各予約サイトの評価基準が異なる点に留意が必要です。海外OTA（Booking.com/Trip.com/Agoda）は10点満点、国内サイト（じゃらん/楽天トラベル/Google）は5点満点のため、統一比較のため10点換算（5点満点×2）を行いました。";',
        f'const DATA_OVERVIEW_SITE_TEXT = "各予約サイトの評価基準が異なる点に留意が必要です。海外OTA（Booking.com/Trip.com/Agoda）は10点満点、国内サイト（じゃらん/楽天トラベル/Google）は5点満点のため、統一比較のため10点換算（5点満点×2）を行いました。";')
    js = js.replace(
        'const DATA_OVERVIEW_DIST_TEXT = "【評価分布の説明テキスト】";',
        f'const DATA_OVERVIEW_DIST_TEXT = "{esc(dist_text)}";')

    # STRENGTH_THEMES
    old_strength = '''const STRENGTH_THEMES = [
  ["テーマ1", "XX件", "「代表的なコメント1」「代表的なコメント2」"],
  ["テーマ2", "XX件", "「代表的なコメント1」「代表的なコメント2」"],
  ["テーマ3", "XX件", "「代表的なコメント1」「代表的なコメント2」"],
  ["テーマ4", "XX件", "「代表的なコメント1」「代表的なコメント2」"],
  ["テーマ5", "XX件", "「代表的なコメント1」「代表的なコメント2」"],
  ["テーマ6", "XX件", "「代表的なコメント1」「代表的なコメント2」"],
];'''
    js = js.replace(old_strength, f'const STRENGTH_THEMES = {strength_js};')

    # STRENGTH_SUB_1
    sub1_title = f"3.1 最大の強み：「{top_s['theme']}」"
    sub1_text = f"{hotel_name}の最大の強みは「{top_s['theme']}」です。{total}件のレビュー中{top_s['count']}件でこのテーマへの言及が確認されました。"
    sub1_bullets_js = "\n  ".join([f'"{esc(b)}",' for b in strength_sub1_bullets])
    sub2_title = f"3.2 第2の強み：「{top_s2['theme']}」"
    sub2_text = f"「{top_s2['theme']}」も{top_s2['count']}件のレビューで高く評価されており、{hotel_name}の重要な差別化要因となっています。"

    js = js.replace('const STRENGTH_SUB_1_TITLE = "3.1 最大の強み：「【強み名】」";',
                    f'const STRENGTH_SUB_1_TITLE = "{esc(sub1_title)}";')
    js = js.replace('const STRENGTH_SUB_1_TEXT = "【最大の強みの説明テキスト】";',
                    f'const STRENGTH_SUB_1_TEXT = "{esc(sub1_text)}";')
    js = js.replace(
        '''const STRENGTH_SUB_1_BULLETS = [
  "【ポイント1】",
  "【ポイント2】",
  "【ポイント3】",
  "【ポイント4】",
];''',
        f'const STRENGTH_SUB_1_BULLETS = [\n  {sub1_bullets_js}\n];')
    js = js.replace('const STRENGTH_SUB_2_TITLE = "3.2 【第2の強みタイトル】";',
                    f'const STRENGTH_SUB_2_TITLE = "{esc(sub2_title)}";')
    js = js.replace('const STRENGTH_SUB_2_TEXT = "【第2の強みの説明テキスト】";',
                    f'const STRENGTH_SUB_2_TEXT = "{esc(sub2_text)}";')

    # WEAKNESS_PRIORITY_DATA
    old_weakness = '''const WEAKNESS_PRIORITY_DATA = [
  ["S", "【課題カテゴリ1】", "【具体的内容】", "【影響度・X件】"],
  ["A", "【課題カテゴリ2】", "【具体的内容】", "【影響度・X件】"],
  ["A", "【課題カテゴリ3】", "【具体的内容】", "【影響度・X件】"],
  ["B", "【課題カテゴリ4】", "【具体的内容】", "【影響度・X件】"],
  ["B", "【課題カテゴリ5】", "【具体的内容】", "【影響度・X件】"],
  ["C", "【課題カテゴリ6】", "【具体的内容】", "【影響度・X件】"],
];'''
    js = js.replace(old_weakness, f'const WEAKNESS_PRIORITY_DATA = {weakness_js};')

    # Phase 1
    old_phase1 = '''const PHASE1_DESCRIPTION = "投資不要・オペレーション改善で対応可能な施策";
const PHASE1_ITEMS = [
  {
    title: "(1) 【施策タイトル】",
    bullets: [
      "【施策詳細1】",
      "【施策詳細2】",
      "【施策詳細3】",
    ],
  },
  {
    title: "(2) 【施策タイトル】",
    bullets: [
      "【施策詳細1】",
      "【施策詳細2】",
      "【施策詳細3】",
    ],
  },
  // ... 必要な数だけ追加
];'''
    new_phase1 = f'const PHASE1_DESCRIPTION = "投資不要・オペレーション改善で対応可能な施策";\nconst PHASE1_ITEMS = {format_phase_items_js(phase1_items)};'
    js = js.replace(old_phase1, new_phase1)

    old_phase2 = '''const PHASE2_DESCRIPTION = "一定の投資を伴うが、比較的早期に実行可能な施策";
const PHASE2_ITEMS = [
  {
    title: "(1) 【施策タイトル】",
    bullets: [
      "【施策詳細1】",
      "【施策詳細2】",
    ],
  },
  // ... 必要な数だけ追加
];'''

    def phase_to_docx_format(items):
        result = []
        for item in items:
            result.append({"title": f"(1) {item['title']}", "bullets": item["items"]})
        return result

    new_phase2 = f'const PHASE2_DESCRIPTION = "一定の投資を伴うが、比較的早期に実行可能な施策";\nconst PHASE2_ITEMS = {format_phase_items_js(phase_to_docx_format(phase2_items))};'
    js = js.replace(old_phase2, new_phase2)

    old_phase3 = '''const PHASE3_DESCRIPTION = "設備投資を伴う抜本的な改善施策";
const PHASE3_ITEMS = [
  {
    title: "(1) 【施策タイトル】",
    bullets: [
      "【施策詳細1】",
      "【施策詳細2】",
    ],
  },
  // ... 必要な数だけ追加
];'''
    new_phase3 = f'const PHASE3_DESCRIPTION = "設備投資を伴う抜本的な改善施策";\nconst PHASE3_ITEMS = {format_phase_items_js(phase_to_docx_format(phase3_items))};'
    js = js.replace(old_phase3, new_phase3)

    # KPI targets
    old_kpi = '''const KPI_TARGET_DATA = [
  ["全体平均(10pt換算)", "X.XX点", "X.X点以上", "YYYY年M月"],
  ["高評価率（8-10点）", "XX.X%", "XX%以上", "YYYY年M月"],
  ["低評価率（1-4点）", "X.X%", "X%維持", "YYYY年M月"],
  // ... 必要な行数だけ追加
];'''
    js = js.replace(old_kpi, f'const KPI_TARGET_DATA = {kpi_target_js};')

    # Conclusion
    conclusion_pars_js = ",\n  ".join([f'"{esc(p)}"' for p in conclusion_pars])
    old_conclusion = '''const CONCLUSION_PARAGRAPHS = [
  "【総括の第1段落テキスト】",
  "【総括の第2段落テキスト】",
  "【総括の第3段落テキスト】",
  "【総括の第4段落テキスト】",
];
const CONCLUSION_FINAL_PARAGRAPH = "【総括の最終段落（太字・NAVY色で表示）】";'''
    new_conclusion = f'const CONCLUSION_PARAGRAPHS = [\n  {conclusion_pars_js},\n];\nconst CONCLUSION_FINAL_PARAGRAPH = "{esc(conclusion_final)}";'
    js = js.replace(old_conclusion, new_conclusion)

    return js



def generate_pptx_js(hotel_name, data, strength_analysis, weaknesses):
    with open(PPTX_TEMPLATE, "r", encoding="utf-8") as f:
        template = f.read()

    total = data["total_reviews"]
    avg = data["overall_avg_10pt"]
    high_count = data["high_count"]
    high_rate = data["high_rate"]
    mid_count = data["mid_count"]
    mid_rate = data["mid_rate"]
    low_count = data["low_count"]
    low_rate = data["low_rate"]

    def esc(s): return js_str(s)

    phase1_items, phase2_items, phase3_items = generate_improvement_actions(weaknesses, strength_analysis)
    kpi_targets = build_kpi_targets(data)

    # site chart
    site_labels = [s["site"] for s in data.get("site_stats", [])]
    site_values = [round(s["avg_10pt"], 2) for s in data.get("site_stats", [])]

    site_labels_js = "[" + ", ".join([f'"{esc(s)}"' for s in site_labels]) + "]"
    site_values_js = "[" + ", ".join([str(v) for v in site_values]) + "]"

    # site table rows
    site_table_lines = []
    for s in data.get("site_stats", []):
        site_table_lines.append(f'  ["{esc(s["site"])}", "{s["count"]}", "{s["native_avg"]:.2f}", "{esc(s["scale"])}", "{s["avg_10pt"]:.2f}"]')
    site_table_js = "[\n" + ",\n".join(site_table_lines) + "\n]"

    # distribution
    dist = data.get("distribution", [])
    bar_labels = [f"{d['score']}点" for d in dist]
    bar_values = [d["count"] for d in dist]
    bar_labels_js = "[" + ", ".join([f'"{l}"' for l in bar_labels]) + "]"
    bar_values_js = "[" + ", ".join([str(v) for v in bar_values]) + "]"
    doughnut_vals_js = f"[{high_count}, {mid_count}]"

    # summary cards
    low_col = "DC2626" if low_count > 0 else "16A34A"
    low_bg = "FEE2E2" if low_count > 0 else "DCFCE7"
    dist_summary_js = f"""[
  {{ label: "高評価 8-10点", val: "{high_count}件（{high_rate:.1f}%）", col: "16A34A", bg: "DCFCE7" }},
  {{ label: "中評価 5-7点", val: "{mid_count}件（{mid_rate:.1f}%）", col: "D4A843", bg: "FFF7ED" }},
  {{ label: "低評価 1-4点", val: "{low_count}件（{low_rate:.1f}%）", col: "{low_col}", bg: "{low_bg}" }},
]"""

    # Strengths cards (6 items)
    strengths_card_lines = []
    theme_descs = {
        "立地・アクセス": "駅近・交通利便性に関する好評価",
        "清潔感": "客室・共用部の清潔さへの高評価",
        "スタッフ対応": "フロントスタッフの丁寧な対応への好評",
        "部屋・設備": "客室の快適性・設備の充実度への好評",
        "朝食": "朝食の品質・バリエーションへの高評価",
        "コスパ・リピート": "価格対価値の高さとリピート意向",
    }
    for s in strength_analysis[:6]:
        theme = s["theme"]
        count = s["count"]
        desc = theme_descs.get(theme, f"{theme}への好評価")
        quote = s["quotes"][0] if s["quotes"] else "口コミより"
        strengths_card_lines.append(
            f'  {{ theme: "{esc(theme)}", count: "{count}件", desc: "{esc(desc)}", quote: "「{esc(quote)}」" }}'
        )
    # Pad to 6
    all_themes = [t for t, _ in STRENGTH_THEMES]
    for th in all_themes:
        if len(strengths_card_lines) >= 6:
            break
        if not any(th in line for line in strengths_card_lines):
            desc = theme_descs.get(th, "")
            strengths_card_lines.append(f'  {{ theme: "{esc(th)}", count: "0件", desc: "{esc(desc)}", quote: "「口コミより」" }}')
    strengths_cards_js = "[\n" + ",\n".join(strengths_card_lines[:6]) + "\n]"

    # Weakness items
    weakness_items_lines = []
    for w in weaknesses[:9]:
        pri = w["priority"]
        cat = w["cat"]
        detail = w.get("detail", "")[:40]
        count = w["count"]
        color = w["color"]
        weakness_items_lines.append(
            f'  {{ pri: "{pri}", cat: "{esc(cat)}", detail: "{esc(detail)}", count: "{count}件", color: "{color}" }}'
        )
    weakness_items_js = "[\n" + ",\n".join(weakness_items_lines) + "\n]"

    # Summary strengths/weaknesses (top 3 each)
    sum_strengths_lines = []
    for s in strength_analysis[:3]:
        desc = theme_descs.get(s["theme"], f"{s['theme']}への好評価") + f"（{s['count']}件）"
        sum_strengths_lines.append(f'  {{ title: "{esc(s["theme"])}", desc: "{esc(desc)}" }}')
    sum_strengths_js = "[\n" + ",\n".join(sum_strengths_lines) + "\n]"

    sum_weaknesses_lines = []
    for w in weaknesses[:3]:
        if w["count"] > 0:
            desc = f"{w['count']}件の指摘あり。改善対応を推奨"
        else:
            desc = "継続的な品質管理・予防的対応を推奨"
        sum_weaknesses_lines.append(f'  {{ title: "{esc(w["cat"])}", desc: "{esc(desc)}" }}')
    if len(sum_weaknesses_lines) < 3:
        sum_weaknesses_lines.append('  { title: "継続的改善", desc: "品質維持のための定期モニタリングを推奨" }')
    sum_weaknesses_js = "[\n" + ",\n".join(sum_weaknesses_lines[:3]) + "\n]"

    # Site insight
    top_site = data["site_stats"][0] if data.get("site_stats") else {"site": "サイト", "avg_10pt": 0}
    top2_site = data["site_stats"][1] if len(data.get("site_stats", [])) > 1 else {"site": "サイト2", "avg_10pt": 0}
    insight_texts_js = f"""[
  {{ text: "サイト別評価インサイト", options: {{ bold: true, fontSize: 12, breakLine: true }} }},
  {{ text: "", options: {{ fontSize: 6, breakLine: true }} }},
  {{ text: "最高評価サイト", options: {{ bold: true, color: "16A34A", breakLine: true }} }},
  {{ text: "{esc(top_site['site'])}: {top_site['avg_10pt']:.2f}点", options: {{ fontSize: 9, color: "64748B", breakLine: true }} }},
  {{ text: "", options: {{ fontSize: 6, breakLine: true }} }},
  {{ text: "全体平均", options: {{ bold: true, color: "16A34A", breakLine: true }} }},
  {{ text: "{avg:.2f}点（10pt換算）", options: {{ fontSize: 9, color: "64748B", breakLine: true }} }},
  {{ text: "", options: {{ fontSize: 6, breakLine: true }} }},
  {{ text: "注：", options: {{ bold: true, fontSize: 9, color: "64748B" }} }},
  {{ text: "海外OTAは/10、国内サイトは/5×2換算", options: {{ fontSize: 9, color: "64748B" }} }},
]"""

    # Phase 1 cards (4 items)
    p1_card_lines = []
    for item in phase1_items[:4]:
        bullets_js = ", ".join([f'"{esc(b)}"' for b in item["bullets"][:4]])
        title = item["title"].replace('"', '\\"')
        p1_card_lines.append(f'  {{ title: "{esc(title)}", items: [{bullets_js}] }}')
    phase1_cards_js = "[\n" + ",\n".join(p1_card_lines) + "\n]"

    # Phase 2 items
    p2_lines = []
    for item in phase2_items[:3]:
        bullets_js = ", ".join([f'"{esc(b)}"' for b in item["items"][:3]])
        p2_lines.append(f'  {{ title: "{esc(item["title"])}", items: [{bullets_js}] }}')
    phase2_items_js = "[\n" + ",\n".join(p2_lines) + "\n]"

    # Phase 3 items
    p3_lines = []
    for item in phase3_items[:3]:
        bullets_js = ", ".join([f'"{esc(b)}"' for b in item["items"][:3]])
        p3_lines.append(f'  {{ title: "{esc(item["title"])}", items: [{bullets_js}] }}')
    phase3_items_js = "[\n" + ",\n".join(p3_lines) + "\n]"

    # KPI targets
    kpi_target_lines = []
    for row in kpi_targets[:7]:
        kpi_target_lines.append(f'  ["{esc(row[0])}", "{esc(row[1])}", "{esc(row[2])}", "{esc(row[3])}"]')
    kpi_targets_js = "[\n" + ",\n".join(kpi_target_lines) + "\n]"

    # KPI note
    top_issue = weaknesses[0]["cat"] if weaknesses and weaknesses[0]["count"] > 0 else "サービス全般"
    kpi_note = f"月次でOTA評価スコアを確認し、{top_issue}改善の進捗をトラッキングします。四半期ごとに目標値を見直し、継続的な改善サイクルを確立してください。"

    # Conclusion
    top_s = strength_analysis[0] if strength_analysis else {"theme": "立地・アクセス"}
    closing_pars_js = f"""[
  {{ text: "{esc(hotel_name)}の{ANALYSIS_PERIOD}口コミ分析（{total}件）では、総合評価{avg:.2f}点（10pt換算）、高評価率{high_rate:.1f}%という結果が得られました。", options: {{ breakLine: true, fontSize: 12 }} }},
  {{ text: "", options: {{ fontSize: 6, breakLine: true }} }},
  {{ text: "「{esc(top_s['theme'])}」が最大の強みとして確認され、本レポートの3フェーズ改善施策の実施により、さらなる評価向上が期待されます。", options: {{ breakLine: true, fontSize: 12 }} }},
  {{ text: "", options: {{ fontSize: 6, breakLine: true }} }},
  {{ text: "6ヶ月以内に全体平均{min(9.5, round(avg+0.3,1))}点以上、高評価率{min(95, round(high_rate+5))}%以上を目指し、継続的な改善サイクルを確立することを推奨します。", options: {{ breakLine: true, fontSize: 12 }} }},
  {{ text: "", options: {{ fontSize: 6, breakLine: true }} }},
  {{ text: "プライムチェンジは引き続き改善施策の実施をサポートいたします。", options: {{ bold: true, fontSize: 12, color: "D4A843" }} }},
]"""

    js = template
    js = js.replace('const OUTPUT_DIR = "/Users/mitsugutakahashi/ホテル口コミ";',
                    f'const OUTPUT_DIR = "{OUT_DIR}";')
    js = js.replace('const HOTEL_NAME = "【ホテル名】";',
                    f'const HOTEL_NAME = "{esc(hotel_name)}";')
    js = js.replace('const ANALYSIS_PERIOD = "YYYY年M月〜M月";',
                    f'const ANALYSIS_PERIOD = "{ANALYSIS_PERIOD}";')
    js = js.replace('const REVIEW_COUNT = "XX件（6サイト）";',
                    f'const REVIEW_COUNT = "{total}件";')
    js = js.replace('const REPORT_DATE = "YYYY年M月D日";',
                    f'const REPORT_DATE = "{REPORT_DATE}";')

    js = js.replace('const KPI_AVG = "0.00";           // 全体平均(10pt換算)',
                    f'const KPI_AVG = "{avg:.2f}";')
    js = js.replace('const KPI_HIGH_RATE = "0.0%";     // 高評価率(8-10点)',
                    f'const KPI_HIGH_RATE = "{high_rate:.1f}%";')
    js = js.replace('const KPI_LOW_RATE = "0.0%";      // 低評価率(1-4点)',
                    f'const KPI_LOW_RATE = "{low_rate:.1f}%";')
    js = js.replace('const KPI_TOTAL_COUNT = "0件";    // レビュー総数',
                    f'const KPI_TOTAL_COUNT = "{total}件";')

    # Summary strengths
    old_sum_str = '''const SUMMARY_STRENGTHS = [
  { title: "【強み1タイトル】", desc: "【強み1の説明文】" },
  { title: "【強み2タイトル】", desc: "【強み2の説明文】" },
  { title: "【強み3タイトル】", desc: "【強み3の説明文】" },
];'''
    js = js.replace(old_sum_str, f'const SUMMARY_STRENGTHS = {sum_strengths_js};')

    old_sum_weak = '''const SUMMARY_WEAKNESSES = [
  { title: "【弱み1タイトル】", desc: "【弱み1の説明文】" },
  { title: "【弱み2タイトル】", desc: "【弱み2の説明文】" },
  { title: "【弱み3タイトル】", desc: "【弱み3の説明文】" },
];'''
    js = js.replace(old_sum_weak, f'const SUMMARY_WEAKNESSES = {sum_weaknesses_js};')

    js = js.replace('const SITE_CHART_LABELS = ["Site1", "Site2", "Site3", "Site4", "Site5", "Site6"];',
                    f'const SITE_CHART_LABELS = {site_labels_js};')
    js = js.replace('const SITE_CHART_VALUES = [0, 0, 0, 0, 0, 0]; // 10pt換算平均',
                    f'const SITE_CHART_VALUES = {site_values_js};')

    old_insight = '''const SITE_INSIGHT_TEXTS = [
  { text: "【インサイト見出し】", options: { bold: true, fontSize: 12, breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "【カテゴリ1】", options: { bold: true, color: "16A34A", breakLine: true } },
  { text: "【カテゴリ1の詳細】", options: { fontSize: 9, color: "64748B", breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "【カテゴリ2】", options: { bold: true, color: "16A34A", breakLine: true } },
  { text: "【カテゴリ2の詳細】", options: { fontSize: 9, color: "64748B", breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "注：", options: { bold: true, fontSize: 9, color: "64748B" } },
  { text: "【注記テキスト】", options: { fontSize: 9, color: "64748B" } },
];'''
    js = js.replace(old_insight, f'const SITE_INSIGHT_TEXTS = {insight_texts_js};')

    old_site_table = '''const SITE_TABLE_ROWS_DATA = [
  ["Site1", "0", "0.00", "/10", "0.00"],
  ["Site2", "0", "0.00", "/5", "0.00"],
  ["Site3", "0", "0.00", "/5", "0.00"],
  ["Site4", "0", "0.00", "/10", "0.00"],
  ["Site5", "0", "0.00", "/10", "0.00"],
  ["Site6", "0", "0.00", "/5", "0.00"],
];'''
    js = js.replace(old_site_table, f'const SITE_TABLE_ROWS_DATA = {site_table_js};')

    js = js.replace('const DIST_DOUGHNUT_LABELS = ["高評価(8-10)", "中評価(5-7)"];',
                    'const DIST_DOUGHNUT_LABELS = ["高評価(8-10)", "中評価(5-7)"];')
    js = js.replace('const DIST_DOUGHNUT_VALUES = [0, 0];',
                    f'const DIST_DOUGHNUT_VALUES = {doughnut_vals_js};')
    js = js.replace('const DIST_BAR_LABELS = ["10点", "9点", "8点", "7点", "6点", "5点"];',
                    f'const DIST_BAR_LABELS = {bar_labels_js};')
    js = js.replace('const DIST_BAR_VALUES = [0, 0, 0, 0, 0, 0];',
                    f'const DIST_BAR_VALUES = {bar_values_js};')

    old_dist_summary = '''const DIST_SUMMARY_CARDS = [
  { label: "高評価 8-10点", val: "0件（0.0%）", col: "16A34A", bg: "DCFCE7" },
  { label: "中評価 5-7点", val: "0件（0.0%）", col: "D4A843", bg: "FFF7ED" },
  { label: "低評価 1-4点", val: "0件（0.0%）", col: "16A34A", bg: "DCFCE7" },
];'''
    js = js.replace(old_dist_summary, f'const DIST_SUMMARY_CARDS = {dist_summary_js};')

    old_strengths_cards = '''const STRENGTHS_CARDS = [
  { theme: "【テーマ1】", count: "0件", desc: "【説明1】", quote: "「【引用1】」" },
  { theme: "【テーマ2】", count: "0件", desc: "【説明2】", quote: "「【引用2】」" },
  { theme: "【テーマ3】", count: "0件", desc: "【説明3】", quote: "「【引用3】」" },
  { theme: "【テーマ4】", count: "0件", desc: "【説明4】", quote: "「【引用4】」" },
  { theme: "【テーマ5】", count: "0件", desc: "【説明5】", quote: "「【引用5】」" },
  { theme: "【テーマ6】", count: "0件", desc: "【説明6】", quote: "「【引用6】」" },
];'''
    js = js.replace(old_strengths_cards, f'const STRENGTHS_CARDS = {strengths_cards_js};')

    old_weakness_items = '''const WEAKNESS_ITEMS = [
  { pri: "S", cat: "【カテゴリ1】", detail: "【詳細1】", count: "0件", color: "DC2626" },
  { pri: "A", cat: "【カテゴリ2】", detail: "【詳細2】", count: "0件", color: "EA580C" },
  { pri: "A", cat: "【カテゴリ3】", detail: "【詳細3】", count: "0件", color: "EA580C" },
  { pri: "B", cat: "【カテゴリ4】", detail: "【詳細4】", count: "0件", color: "D4A843" },
  { pri: "B", cat: "【カテゴリ5】", detail: "【詳細5】", count: "0件", color: "D4A843" },
  { pri: "B", cat: "【カテゴリ6】", detail: "【詳細6】", count: "0件", color: "D4A843" },
  { pri: "B", cat: "【カテゴリ7】", detail: "【詳細7】", count: "0件", color: "D4A843" },
  { pri: "C", cat: "【カテゴリ8】", detail: "【詳細8】", count: "0件", color: "94A3B8" },
  { pri: "C", cat: "【カテゴリ9】", detail: "【詳細9】", count: "0件", color: "94A3B8" },
];'''
    js = js.replace(old_weakness_items, f'const WEAKNESS_ITEMS = {weakness_items_js};')

    old_phase1_cards = '''const PHASE1_CARDS = [
  { title: "【施策1タイトル】", items: ["【アクション1-1】", "【アクション1-2】", "【アクション1-3】", "【アクション1-4】"] },
  { title: "【施策2タイトル】", items: ["【アクション2-1】", "【アクション2-2】", "【アクション2-3】"] },
  { title: "【施策3タイトル】", items: ["【アクション3-1】", "【アクション3-2】", "【アクション3-3】"] },
  { title: "【施策4タイトル】", items: ["【アクション4-1】", "【アクション4-2】", "【アクション4-3】"] },
];'''
    js = js.replace(old_phase1_cards, f'const PHASE1_CARDS = {phase1_cards_js};')

    old_phase2_items = '''const PHASE2_ITEMS = [
  { title: "【P2施策1タイトル】", items: ["【P2アクション1-1】", "【P2アクション1-2】", "【P2アクション1-3】"] },
  { title: "【P2施策2タイトル】", items: ["【P2アクション2-1】", "【P2アクション2-2】", "【P2アクション2-3】"] },
  { title: "【P2施策3タイトル】", items: ["【P2アクション3-1】", "【P2アクション3-2】"] },
];'''
    js = js.replace(old_phase2_items, f'const PHASE2_ITEMS = {phase2_items_js};')

    old_phase3_items = '''const PHASE3_ITEMS = [
  { title: "【P3施策1タイトル】", items: ["【P3アクション1-1】", "【P3アクション1-2】", "【P3アクション1-3】"] },
  { title: "【P3施策2タイトル】", items: ["【P3アクション2-1】", "【P3アクション2-2】"] },
  { title: "【P3施策3タイトル】", items: ["【P3アクション3-1】", "【P3アクション3-2】", "【P3アクション3-3】"] },
];'''
    js = js.replace(old_phase3_items, f'const PHASE3_ITEMS = {phase3_items_js};')

    old_kpi_rows = '''const KPI_TARGET_ROWS = [
  ["全体平均(10pt換算)", "0.00点", "0.0点以上", "YYYY年M月"],
  ["高評価率（8-10点）", "0.0%", "0%以上", "YYYY年M月"],
  ["低評価率（1-4点）", "0.0%", "0%維持", "YYYY年M月"],
  ["【サイト1】平均評価", "0.00/5点", "0.0/5点以上", "YYYY年M月"],
  ["【サイト2】平均評価", "0.00/5点", "0.0/5点以上", "YYYY年M月"],
  ["口コミ返信率", "未計測", "100%（48h以内）", "YYYY年M月"],
  ["【重点課題】クレーム", "0件/2ヶ月", "0件以下/2ヶ月", "YYYY年M月"],
];'''
    js = js.replace(old_kpi_rows, f'const KPI_TARGET_ROWS = {kpi_targets_js};')

    js = js.replace('const KPI_NOTE_BOLD = "モニタリング方針：";',
                    'const KPI_NOTE_BOLD = "モニタリング方針：";')
    js = js.replace('const KPI_NOTE_TEXT = "【モニタリング方針の説明文をここに記載】";',
                    f'const KPI_NOTE_TEXT = "{esc(kpi_note)}";')

    old_closing = '''const CLOSING_PARAGRAPHS = [
  { text: "【総括パラグラフ1】", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "【総括パラグラフ2】", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "【総括パラグラフ3】", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "【締めくくりメッセージ】", options: { bold: true, fontSize: 12, color: "D4A843" } },
];'''
    js = js.replace(old_closing, f'const CLOSING_PARAGRAPHS = {closing_pars_js};')

    return js


# ─────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────

def run_hotel(key, hotel_name):
    json_path = f"{JSON_DIR}/{key}_analysis.json"
    if not os.path.exists(json_path):
        print(f"  [SKIP] JSON not found: {json_path}")
        return False, False

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    comments = data.get("comments", [])
    strength_analysis = analyze_strengths(comments)
    weaknesses = analyze_weaknesses(comments)

    tmpdir = tempfile.mkdtemp()
    docx_success = False
    pptx_success = False

    try:
        # Generate DOCX
        docx_js = generate_docx_js(hotel_name, data, strength_analysis, weaknesses)
        docx_js_path = os.path.join(tmpdir, f"{key}_docx.js")
        with open(docx_js_path, "w", encoding="utf-8") as f:
            f.write(docx_js)

        result = subprocess.run(
            ["node", docx_js_path],
            capture_output=True, text=True, timeout=60,
            env={**os.environ, "NODE_PATH": NODE_MODULES}
        )
        if result.returncode == 0:
            expected = f"{OUT_DIR}/{hotel_name}_口コミ分析改善レポート.docx"
            if os.path.exists(expected):
                print(f"  [OK] DOCX: {hotel_name}_口コミ分析改善レポート.docx")
                docx_success = True
            else:
                print(f"  [WARN] DOCX node ran OK but file not found: {expected}")
                if result.stdout:
                    print(f"    stdout: {result.stdout.strip()}")
        else:
            print(f"  [FAIL] DOCX node error (rc={result.returncode})")
            if result.stderr:
                print(f"    stderr: {result.stderr.strip()[:300]}")
            if result.stdout:
                print(f"    stdout: {result.stdout.strip()[:200]}")

        # Generate PPTX
        pptx_js = generate_pptx_js(hotel_name, data, strength_analysis, weaknesses)
        pptx_js_path = os.path.join(tmpdir, f"{key}_pptx.js")
        with open(pptx_js_path, "w", encoding="utf-8") as f:
            f.write(pptx_js)

        result = subprocess.run(
            ["node", pptx_js_path],
            capture_output=True, text=True, timeout=60,
            env={**os.environ, "NODE_PATH": NODE_MODULES}
        )
        if result.returncode == 0:
            expected = f"{OUT_DIR}/{hotel_name}_口コミ分析レポート.pptx"
            if os.path.exists(expected):
                print(f"  [OK] PPTX: {hotel_name}_口コミ分析レポート.pptx")
                pptx_success = True
            else:
                print(f"  [WARN] PPTX node ran OK but file not found: {expected}")
                if result.stdout:
                    print(f"    stdout: {result.stdout.strip()}")
        else:
            print(f"  [FAIL] PPTX node error (rc={result.returncode})")
            if result.stderr:
                print(f"    stderr: {result.stderr.strip()[:300]}")
            if result.stdout:
                print(f"    stdout: {result.stdout.strip()[:200]}")

    except Exception as e:
        print(f"  [ERROR] {e}")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    return docx_success, pptx_success


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    docx_ok = 0
    pptx_ok = 0
    failed = []

    for key, name in HOTELS.items():
        print(f"\n[{key}] {name}")
        d, p = run_hotel(key, name)
        if d:
            docx_ok += 1
        if p:
            pptx_ok += 1
        if not d or not p:
            failed.append(key)

    print(f"\n{'='*50}")
    print(f"Results: {docx_ok}/19 DOCX, {pptx_ok}/19 PPTX created successfully")
    if failed:
        print(f"Failed/partial: {', '.join(failed)}")
    print(f"Output dir: {OUT_DIR}")


if __name__ == "__main__":
    main()
