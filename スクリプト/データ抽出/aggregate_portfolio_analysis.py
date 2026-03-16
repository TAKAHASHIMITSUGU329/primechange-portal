#!/usr/bin/env python3
"""
PRIMECHANGE Portfolio Aggregation Analysis
全19ホテルの口コミ分析JSONを統合し、清掃キーワード分析を行う
"""
import json
import os
import re
import sys
from collections import defaultdict

# ===== Configuration =====

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.join(SCRIPT_DIR, '..', '..', 'データ', '分析結果JSON')

# File key -> (analysis JSON filename, Japanese hotel name)
HOTEL_MAP = {
    'daiwa_osaki': ('daiwa_osaki_analysis.json', 'ダイワロイネットホテル東京大崎'),
    'chisan': ('chisan_analysis.json', 'チサンホテル浜松町'),
    'hearton': ('hearton_analysis.json', 'ハートンホテル東品川'),
    'keyakigate': ('keyakigate_analysis.json', 'ホテルケヤキゲート東京府中'),
    'richmond_mejiro': ('richmond_mejiro_analysis.json', 'リッチモンドホテル東京目白'),
    'keisei_kinshicho': ('keisei_kinshicho_analysis.json', '京成リッチモンドホテル東京錦糸町'),
    'daiichi_ikebukuro': ('daiichi_ikebukuro_analysis.json', '第一イン池袋'),
    'comfort_roppongi': ('comfort_roppongi_analysis.json', 'コンフォートイン六本木'),
    'comfort_suites_tokyobay': ('comfort_suites_tokyobay_analysis.json', 'コンフォートスイーツ東京ベイ'),
    'comfort_era_higashikanda': ('comfort_era_higashikanda_analysis.json', 'コンフォートホテルERA東京東神田'),
    'comfort_yokohama_kannai': ('comfort_yokohama_kannai_analysis.json', 'コンフォートホテル横浜関内'),
    'comfort_narita': ('comfort_narita_analysis.json', 'コンフォートホテル成田'),
    'apa_kamata': ('apa_kamata_analysis.json', 'アパホテル蒲田駅東'),
    'apa_sagamihara': ('apa_sagamihara_analysis.json', 'アパホテル相模原橋本駅東'),
    'court_shinyokohama': ('court_shinyokohama_analysis.json', 'コートホテル新横浜'),
    'comment_yokohama': ('comment_yokohama_analysis.json', 'ホテルコメント横浜関内'),
    'kawasaki_nikko': ('kawasaki_nikko_analysis.json', '川崎日航ホテル'),
    'henn_na_haneda': ('henn_na_haneda_analysis.json', '変なホテル東京羽田'),
    'comfort_hakata': ('comfort_hakata_analysis.json', 'コンフォートホテル博多'),
}

# Cleaning keyword categories
CLEANING_CATEGORIES = {
    'カビ・モールド': ['カビ', 'かび', 'mold', 'mould', 'カビ臭', 'カビが'],
    '毛髪・髪の毛': ['髪の毛', '毛が', '毛髪', 'hair', '毛が落', '抜け毛', '陰毛'],
    'カーペット汚れ': ['カーペット', 'carpet', '絨毯', 'じゅうたん'],
    '排水・詰まり': ['排水', 'drain', '詰ま', '水はけ', '流れが悪', '流れない'],
    '臭気・異臭': ['臭い', '匂い', '臭', 'smell', 'odor', '下水', 'くさい', '異臭', 'におい'],
    '清掃品質全般': ['清掃', '掃除', '清潔', 'clean', '汚い', '汚れ', '不潔', 'dirty', '拭き残', '拭いて'],
    'ほこり・塵': ['ほこり', '埃', 'dust', 'ホコリ'],
    'シミ・汚れ跡': ['シミ', 'stain', '染み', 'しみ'],
    'タバコ関連': ['タバコ', '煙草', 'cigarette', 'smoke', 'たばこ', '喫煙'],
    '害虫': ['虫', 'bug', 'insect', 'ゴキブリ', '蟻', 'アリ', 'クモ'],
    'ゴミ残置': ['ゴミ', 'garbage', 'trash', 'ごみ'],
}

# Action mapping for each cleaning category
ACTION_MAP = {
    'カビ・モールド': {
        'phase1': [
            'エアコンフィルター・内部のカビ除去と消毒',
            '加湿器の分解洗浄・カビ除去',
            '浴室天井・壁・目地のカビ取り処理',
        ],
        'phase2': [
            '防カビコーティング施工（浴室・エアコン周り）',
            'エアコン月次メンテナンスサイクル導入',
        ],
        'phase3': [
            '換気設備の改善提案（換気扇能力の見直し）',
            '除湿対策の強化（除湿機導入検討）',
        ],
    },
    '毛髪・髪の毛': {
        'phase1': [
            'ベッドメイク後の粘着ローラー仕上げ工程追加',
            '排水口ヘアキャッチャーの毎回清掃・交換強化',
            'バスルーム清掃後の目視チェック強化',
        ],
        'phase2': [
            '清掃チェックリストに毛髪確認項目を明記',
            '白色・黒色素材箇所の重点チェック体制構築',
        ],
        'phase3': [],
    },
    'カーペット汚れ': {
        'phase1': [
            '汚損箇所のスポットクリーニング即時実施',
            '掃除機がけの重点エリア指定と徹底',
        ],
        'phase2': [
            '四半期ごとのカーペットディープクリーニング導入',
            'カーペットの劣化度調査と交換計画策定',
        ],
        'phase3': [
            '高耐久・防汚カーペットへの段階的張替え提案',
        ],
    },
    '排水・詰まり': {
        'phase1': [
            '排水口の徹底清掃と詰まり除去',
            '排水トラップの点検・洗浄',
        ],
        'phase2': [
            '月次排水管高圧洗浄の定期実施',
            '排水口ストレーナーの定期交換サイクル導入',
        ],
        'phase3': [
            '排水管の老朽化調査と更新計画の策定',
        ],
    },
    '臭気・異臭': {
        'phase1': [
            '臭気発生源の特定と即時対処（排水トラップ、換気扇）',
            '消臭・脱臭処理の実施',
        ],
        'phase2': [
            '定期的な換気・消臭ルーティンの標準化',
            'エアコン・排水系の臭気源の根本対策',
        ],
        'phase3': [
            'オゾン脱臭機の導入検討',
            '空気品質モニタリングの導入',
        ],
    },
    '清掃品質全般': {
        'phase1': [
            '清掃品質基準の明確化と全スタッフへの共有',
            '重点箇所（水回り・ベッド周り）の清掃手順見直し',
        ],
        'phase2': [
            '品質管理チェックリストの導入と実施',
            'QCインスペクション体制の構築（ランダム抜き打ち検査）',
        ],
        'phase3': [
            '清掃スタッフの定期研修プログラム導入',
            'デジタル清掃記録システムの検討',
        ],
    },
    'ほこり・塵': {
        'phase1': [
            '高所・棚上部・エアコン上部のほこり除去',
            'テレビ裏・家具裏の清掃徹底',
        ],
        'phase2': [
            '週次の高所清掃スケジュール導入',
        ],
        'phase3': [],
    },
    'シミ・汚れ跡': {
        'phase1': [
            'シミ箇所の特定と専用クリーナーによる除去',
        ],
        'phase2': [
            'シミ発見時の即時報告・対処フロー構築',
        ],
        'phase3': [],
    },
    'タバコ関連': {
        'phase1': [
            '禁煙室の残留臭調査と脱臭処理',
            '喫煙室との動線分離確認',
        ],
        'phase2': [
            '廊下・共用部の定期脱臭処理スケジュール導入',
        ],
        'phase3': [
            'フロア単位での完全禁煙化提案',
        ],
    },
    '害虫': {
        'phase1': [
            '害虫駆除業者による即時消毒実施',
            '発生箇所の特定と封鎖',
        ],
        'phase2': [
            '月次予防駆除の定期契約導入',
            '侵入経路の点検と封鎖工事',
        ],
        'phase3': [],
    },
    'ゴミ残置': {
        'phase1': [
            '清掃後のゴミ残置チェック項目の追加',
            'ゴミ箱周辺の清掃手順見直し',
        ],
        'phase2': [
            '清掃完了後のファイナルチェック体制の強化',
        ],
        'phase3': [],
    },
}


def load_hotel_data():
    """Load all 19 hotel analysis JSONs"""
    hotels = []
    for key, (filename, hotel_name) in HOTEL_MAP.items():
        filepath = os.path.join(BASE_DIR, filename)
        if not os.path.exists(filepath):
            print("WARNING: Missing file: %s (%s)" % (filename, hotel_name), file=sys.stderr)
            continue

        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)

        data['_key'] = key
        data['_name'] = hotel_name
        data['_filename'] = filename
        hotels.append(data)

    return hotels


def analyze_cleaning_keywords(comments, hotel_name):
    """
    Analyze cleaning-related keywords in negative/neutral comments (rating <= 7).
    Returns category counts and sample comments.
    """
    category_counts = defaultdict(int)
    category_samples = defaultdict(list)
    total_cleaning_mentions = 0

    for c in comments:
        rating = c.get('rating_10pt', 10)
        if rating > 7:
            continue  # Only analyze negative/neutral reviews

        # Combine all comment text fields
        text_parts = []
        for field in ['comment', 'good', 'bad', 'translated', 'translated_good', 'translated_bad']:
            val = c.get(field, '')
            if val:
                text_parts.append(str(val))
        text = ' '.join(text_parts)

        if not text.strip():
            continue

        matched_categories = []
        for category, keywords in CLEANING_CATEGORIES.items():
            for kw in keywords:
                if kw.lower() in text.lower():
                    category_counts[category] += 1
                    matched_categories.append(category)

                    # Keep sample comments (max 3 per category per hotel)
                    if len(category_samples[category]) < 3:
                        sample_text = text[:200]  # Truncate
                        category_samples[category].append({
                            'hotel': hotel_name,
                            'rating': rating,
                            'text': sample_text,
                            'site': c.get('site', ''),
                        })
                    break  # Only count once per keyword category per comment

        if matched_categories:
            total_cleaning_mentions += 1

    return dict(category_counts), dict(category_samples), total_cleaning_mentions


def classify_priority(avg, cleaning_issue_rate, cleaning_issue_count):
    """Classify hotel priority based on cleaning metrics"""
    if cleaning_issue_rate > 10 or (avg < 7.5 and cleaning_issue_count > 3):
        return 'URGENT'
    elif cleaning_issue_rate > 5 or (avg < 8.0 and cleaning_issue_count > 2):
        return 'HIGH'
    elif cleaning_issue_count > 0:
        return 'STANDARD'
    else:
        return 'MAINTENANCE'


def generate_action_plan(hotel_name, priority, category_counts, avg):
    """Generate specific action plans based on detected issues"""
    actions = {
        'phase1_immediate': {'timeline': '1-2週間', 'actions': []},
        'phase2_short_term': {'timeline': '1-3ヶ月', 'actions': []},
        'phase3_medium_term': {'timeline': '3-6ヶ月', 'actions': []},
    }

    # Sort categories by count (most frequent first)
    sorted_cats = sorted(category_counts.items(), key=lambda x: x[1], reverse=True)

    for cat, count in sorted_cats:
        if cat in ACTION_MAP:
            mapping = ACTION_MAP[cat]
            for a in mapping.get('phase1', []):
                actions['phase1_immediate']['actions'].append({
                    'action': a,
                    'category': cat,
                    'issue_count': count,
                })
            for a in mapping.get('phase2', []):
                actions['phase2_short_term']['actions'].append({
                    'action': a,
                    'category': cat,
                })
            for a in mapping.get('phase3', []):
                actions['phase3_medium_term']['actions'].append({
                    'action': a,
                    'category': cat,
                })

    # Set target
    if avg < 7.5:
        target = 8.0
    elif avg < 8.0:
        target = 8.5
    elif avg < 8.5:
        target = 9.0
    else:
        target = avg + 0.3

    return {
        'hotel': hotel_name,
        'priority_level': priority,
        'current_avg': avg,
        'target_avg': round(target, 1),
        'phase1_immediate': actions['phase1_immediate'],
        'phase2_short_term': actions['phase2_short_term'],
        'phase3_medium_term': actions['phase3_medium_term'],
    }


def main():
    print("=== PRIMECHANGE Portfolio Aggregation Analysis ===\n")

    # 1. Load all hotel data
    hotels = load_hotel_data()
    print("Loaded %d hotels" % len(hotels))

    if len(hotels) < 19:
        print("WARNING: Expected 19 hotels, got %d" % len(hotels))

    # 2. Compute portfolio metrics
    total_reviews = sum(h['total_reviews'] for h in hotels)
    weighted_avg = sum(h['overall_avg_10pt'] * h['total_reviews'] for h in hotels) / total_reviews
    avg_scores = sorted([h['overall_avg_10pt'] for h in hotels])
    median_idx = len(avg_scores) // 2
    median_score = avg_scores[median_idx]

    best_hotel = max(hotels, key=lambda h: h['overall_avg_10pt'])
    worst_hotel = min(hotels, key=lambda h: h['overall_avg_10pt'])

    total_high = sum(h['high_count'] for h in hotels)
    total_low = sum(h['low_count'] for h in hotels)
    portfolio_high_rate = round(total_high / total_reviews * 100, 1)
    portfolio_low_rate = round(total_low / total_reviews * 100, 1)

    print("Total reviews: %d" % total_reviews)
    print("Portfolio weighted avg: %.2f" % weighted_avg)
    print("Best: %s (%.2f)" % (best_hotel['_name'], best_hotel['overall_avg_10pt']))
    print("Worst: %s (%.2f)" % (worst_hotel['_name'], worst_hotel['overall_avg_10pt']))

    # 3. Cleaning keyword analysis for each hotel
    all_category_totals = defaultdict(int)
    all_category_hotels = defaultdict(set)
    all_category_samples = defaultdict(list)
    hotel_cleaning_data = []
    total_cleaning_mentions = 0

    for h in hotels:
        comments = h.get('comments', [])
        neg_comments = [c for c in comments if c.get('rating_10pt', 10) <= 7]
        cat_counts, cat_samples, mention_count = analyze_cleaning_keywords(
            comments, h['_name']
        )

        cleaning_issue_rate = round(mention_count / h['total_reviews'] * 100, 1) if h['total_reviews'] > 0 else 0

        priority = classify_priority(
            h['overall_avg_10pt'], cleaning_issue_rate, mention_count
        )

        hotel_cleaning_data.append({
            'key': h['_key'],
            'name': h['_name'],
            'avg': h['overall_avg_10pt'],
            'total_reviews': h['total_reviews'],
            'high_rate': h['high_rate'],
            'low_rate': h['low_rate'],
            'cleaning_issue_count': mention_count,
            'cleaning_issue_rate': cleaning_issue_rate,
            'priority': priority,
            'categories': dict(cat_counts),
            'neg_review_count': len(neg_comments),
        })

        total_cleaning_mentions += mention_count

        for cat, count in cat_counts.items():
            all_category_totals[cat] += count
            all_category_hotels[cat].add(h['_name'])

        for cat, samples in cat_samples.items():
            all_category_samples[cat].extend(samples)

    # Sort hotels by priority then by cleaning issue rate descending
    priority_order = {'URGENT': 0, 'HIGH': 1, 'STANDARD': 2, 'MAINTENANCE': 3}
    hotel_cleaning_data.sort(
        key=lambda h: (priority_order.get(h['priority'], 9), -h['cleaning_issue_rate'], -h['cleaning_issue_count'])
    )

    # Overall cleaning issue rate
    portfolio_cleaning_rate = round(total_cleaning_mentions / total_reviews * 100, 1) if total_reviews > 0 else 0

    print("\nCleaning Analysis:")
    print("Total cleaning mentions: %d (%.1f%%)" % (total_cleaning_mentions, portfolio_cleaning_rate))
    for h in hotel_cleaning_data:
        if h['cleaning_issue_count'] > 0:
            print("  [%s] %s: %d issues (%.1f%%) avg=%.2f" % (
                h['priority'], h['name'], h['cleaning_issue_count'],
                h['cleaning_issue_rate'], h['avg']
            ))

    # 4. Build category summary
    category_summary = []
    for cat in CLEANING_CATEGORIES.keys():
        total = all_category_totals.get(cat, 0)
        if total > 0:
            hotels_affected = len(all_category_hotels.get(cat, set()))
            if total > 10:
                severity = 'CRITICAL'
            elif total > 5:
                severity = 'HIGH'
            elif total > 2:
                severity = 'MEDIUM'
            else:
                severity = 'LOW'

            category_summary.append({
                'category': cat,
                'total_mentions': total,
                'hotels_affected': hotels_affected,
                'severity': severity,
                'sample_comments': all_category_samples.get(cat, [])[:5],
            })

    category_summary.sort(key=lambda x: x['total_mentions'], reverse=True)

    # 5. Build priority matrix
    priority_matrix = {
        'urgent': [],
        'high': [],
        'standard': [],
        'maintenance': [],
    }

    for h in hotel_cleaning_data:
        entry = {
            'hotel': h['name'],
            'avg': h['avg'],
            'total_reviews': h['total_reviews'],
            'cleaning_issues': h['cleaning_issue_count'],
            'cleaning_rate': h['cleaning_issue_rate'],
            'key_problems': [cat for cat, _ in sorted(h['categories'].items(), key=lambda x: x[1], reverse=True)[:3]] if h['categories'] else [],
        }
        priority_matrix[h['priority'].lower()].append(entry)

    # 6. Generate action plans
    action_plans = []
    for h in hotel_cleaning_data:
        if h['priority'] in ('URGENT', 'HIGH'):
            plan = generate_action_plan(
                h['name'], h['priority'], h['categories'], h['avg']
            )
            plan['priority_rank'] = len(action_plans) + 1
            action_plans.append(plan)

    # 7. Cross-cutting recommendations
    cross_cutting = [
        {
            'theme': '清掃チェックリスト標準化',
            'applicable_hotels': '全19ホテル',
            'priority': 'HIGH',
            'description': '統一された50項目の清掃チェックリストを全ホテルに導入。客室タイプ別（シングル・ダブル・ツイン）の基準を明確化し、チェック漏れを防止する。',
            'items': [
                'バスルーム清掃チェック（排水口、鏡、タオルラック、アメニティ配置）',
                'ベッド周り清掃チェック（シーツ交換、枕カバー、粘着ローラー仕上げ）',
                'フロア清掃チェック（掃除機、カーペットスポット確認）',
                'エアコン・換気扇チェック（フィルター状態、異臭確認）',
                '水回り動作確認（排水速度、水圧、温度）',
            ],
        },
        {
            'theme': '品質管理（QC）体制構築',
            'applicable_hotels': '全19ホテル（特にURGENT/HIGHホテル優先）',
            'priority': 'HIGH',
            'description': 'ランダム抜き打ち検査と定期品質監査の二層体制を構築。ゲスト口コミデータとの連動により、改善効果を定量的にモニタリングする。',
            'items': [
                '週次ランダム抜き打ち検査（各ホテル最低3室/週）',
                '月次品質監査レポートの作成と共有',
                'ゲスト口コミスコアの月次トラッキング',
                '清掃クレーム発生時の即時フィードバックループ構築',
            ],
        },
        {
            'theme': '予防保全プログラム',
            'applicable_hotels': '全19ホテル',
            'priority': 'MEDIUM',
            'description': '清掃関連設備の定期メンテナンスにより、カビ・排水・臭気等の根本原因を予防する。',
            'items': [
                'エアコンフィルター月次清掃・半年次内部洗浄',
                '排水管四半期高圧洗浄',
                'カーペットディープクリーニング四半期実施',
                '換気扇・排気設備の半年次点検',
                '害虫予防駆除の月次定期契約',
            ],
        },
        {
            'theme': 'テクノロジー活用による清掃DX',
            'applicable_hotels': 'パイロット3ホテル → 全展開',
            'priority': 'MEDIUM',
            'description': 'デジタルツールを活用した清掃品質の可視化と効率化。URGENTホテルでパイロット導入し、効果検証後に全展開する。',
            'items': [
                'タブレット型デジタル清掃チェックリスト（写真記録付き）',
                '空気品質モニタリングセンサー（臭気・湿度）',
                'UV-C除菌デバイスの導入（枕・リモコン等の高接触面）',
                '清掃完了報告のリアルタイムダッシュボード',
            ],
        },
    ]

    # 8. KPI framework
    kpi_framework = {
        'portfolio_targets': [
            {
                'kpi': 'ポートフォリオ平均スコア',
                'current': '%.2f' % weighted_avg,
                'target': '%.2f' % min(weighted_avg + 0.5, 9.5),
                'deadline': '2026年9月',
            },
            {
                'kpi': '清掃関連クレーム率',
                'current': '%.1f%%' % portfolio_cleaning_rate,
                'target': '%.1f%%以下' % max(portfolio_cleaning_rate * 0.4, 1.0),
                'deadline': '2026年9月',
            },
            {
                'kpi': '高評価率（8-10点）',
                'current': '%.1f%%' % portfolio_high_rate,
                'target': '%.1f%%以上' % min(portfolio_high_rate + 5, 95),
                'deadline': '2026年9月',
            },
            {
                'kpi': '低評価率（1-4点）',
                'current': '%.1f%%' % portfolio_low_rate,
                'target': '%.1f%%以下' % max(portfolio_low_rate * 0.5, 1.0),
                'deadline': '2026年9月',
            },
        ],
        'per_hotel_targets': [],
    }

    # Per-hotel KPI targets
    for h in hotel_cleaning_data:
        if h['priority'] in ('URGENT', 'HIGH'):
            kpi_framework['per_hotel_targets'].append({
                'hotel': h['name'],
                'priority': h['priority'],
                'current_avg': h['avg'],
                'target_avg': round(min(h['avg'] + 0.5, 9.5), 1),
                'current_cleaning_rate': h['cleaning_issue_rate'],
                'target_cleaning_rate': round(max(h['cleaning_issue_rate'] * 0.3, 0.5), 1),
            })

    # 9. ROI estimation
    roi_estimation = {
        'methodology': '口コミスコア0.1pt改善 ≒ RevPAR約1%向上（業界ベンチマーク）',
        'scenarios': [
            {
                'scenario': 'シナリオA: URGENTホテル集中改善',
                'target_hotels': len(priority_matrix['urgent']),
                'estimated_cost': '300-500万円',
                'expected_improvement': '+0.5-1.0pt（対象ホテル平均）',
                'revenue_impact': 'RevPAR 5-10%向上見込み',
                'roi_period': '6-12ヶ月',
            },
            {
                'scenario': 'シナリオB: URGENT + HIGHホテル改善',
                'target_hotels': len(priority_matrix['urgent']) + len(priority_matrix['high']),
                'estimated_cost': '500-800万円',
                'expected_improvement': '+0.3-0.7pt（対象ホテル平均）',
                'revenue_impact': 'RevPAR 3-7%向上見込み',
                'roi_period': '8-14ヶ月',
            },
            {
                'scenario': 'シナリオC: 全社品質管理システム導入',
                'target_hotels': 19,
                'estimated_cost': '800-1,200万円',
                'expected_improvement': '+0.3-0.5pt（ポートフォリオ全体）',
                'revenue_impact': 'ポートフォリオ全体RevPAR 3-5%向上',
                'roi_period': '12-18ヶ月',
            },
        ],
    }

    # 10. Hotels ranked for portfolio overview
    hotels_ranked = []
    sorted_by_avg = sorted(hotels, key=lambda h: h['overall_avg_10pt'], reverse=True)
    for rank, h in enumerate(sorted_by_avg, 1):
        # Find cleaning data
        cleaning_data = next((cd for cd in hotel_cleaning_data if cd['key'] == h['_key']), None)

        avg = h['overall_avg_10pt']
        if avg >= 9.0:
            tier = '優秀'
        elif avg >= 8.5:
            tier = '良好'
        elif avg >= 8.0:
            tier = '概ね良好'
        elif avg >= 7.0:
            tier = '要改善'
        else:
            tier = '要緊急対応'

        hotels_ranked.append({
            'rank': rank,
            'name': h['_name'],
            'key': h['_key'],
            'avg': avg,
            'total_reviews': h['total_reviews'],
            'high_rate': h['high_rate'],
            'low_rate': h['low_rate'],
            'cleaning_issue_count': cleaning_data['cleaning_issue_count'] if cleaning_data else 0,
            'cleaning_issue_rate': cleaning_data['cleaning_issue_rate'] if cleaning_data else 0,
            'priority': cleaning_data['priority'] if cleaning_data else 'MAINTENANCE',
            'tier': tier,
        })

    # 11. Build output JSON
    output = {
        'report_metadata': {
            'title': 'PRIMECHANGE ホテル清掃戦略レポート',
            'subtitle': 'ゲスト口コミデータに基づく清掃品質改善提案書',
            'date': '2026年3月11日',
            'analysis_period': '2025年12月〜2026年3月',
            'total_hotels': len(hotels),
            'total_reviews': total_reviews,
            'portfolio_avg': round(weighted_avg, 2),
            'prepared_for': '株式会社PRIMECHANGE 代表取締役社長',
            'prepared_by': '経営コンサルティングチーム',
            'confidential': True,
        },
        'portfolio_overview': {
            'avg_score': round(weighted_avg, 2),
            'median_score': round(median_score, 2),
            'best_hotel': {'name': best_hotel['_name'], 'avg': best_hotel['overall_avg_10pt']},
            'worst_hotel': {'name': worst_hotel['_name'], 'avg': worst_hotel['overall_avg_10pt']},
            'total_reviews': total_reviews,
            'portfolio_high_rate': portfolio_high_rate,
            'portfolio_low_rate': portfolio_low_rate,
            'hotels_ranked': hotels_ranked,
        },
        'cleaning_deep_dive': {
            'portfolio_cleaning_issue_rate': portfolio_cleaning_rate,
            'total_cleaning_mentions': total_cleaning_mentions,
            'category_summary': category_summary,
            'hotel_cleaning_matrix': hotel_cleaning_data,
        },
        'priority_matrix': priority_matrix,
        'action_plans': action_plans,
        'cross_cutting_recommendations': cross_cutting,
        'kpi_framework': kpi_framework,
        'roi_estimation': roi_estimation,
    }

    # Write output
    output_path = os.path.join(BASE_DIR, 'primechange_portfolio_analysis.json')
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print("\n=== Output saved to %s ===" % output_path)
    print("Hotels: %d" % len(hotels))
    print("URGENT: %d, HIGH: %d, STANDARD: %d, MAINTENANCE: %d" % (
        len(priority_matrix['urgent']),
        len(priority_matrix['high']),
        len(priority_matrix['standard']),
        len(priority_matrix['maintenance']),
    ))


if __name__ == '__main__':
    main()
