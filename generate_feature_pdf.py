#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PRIMECHANGE V3 Dashboard Feature Summary PDF Generator"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

# Register Japanese fonts
pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
pdfmetrics.registerFont(UnicodeCIDFont('HeiseiMin-W3'))

PRIME_RED = HexColor('#C23B3A')
PRIME_DARK = HexColor('#1A1A2E')
PRIME_LIGHT = HexColor('#F8F4F4')
GREEN = HexColor('#10B981')
BLUE = HexColor('#3B82F6')
ORANGE = HexColor('#F59E0B')
GRAY = HexColor('#64748B')
LIGHT_GRAY = HexColor('#E2E8F0')
WHITE = white

# Styles
FONT_GO = 'HeiseiKakuGo-W5'
FONT_MIN = 'HeiseiMin-W3'

sTitle = ParagraphStyle('Title', fontName=FONT_GO, fontSize=24, leading=32, textColor=PRIME_RED, spaceAfter=4*mm)
sSubtitle = ParagraphStyle('Subtitle', fontName=FONT_MIN, fontSize=11, leading=16, textColor=GRAY, spaceAfter=8*mm)
sH1 = ParagraphStyle('H1', fontName=FONT_GO, fontSize=16, leading=22, textColor=PRIME_DARK, spaceBefore=8*mm, spaceAfter=4*mm)
sH2 = ParagraphStyle('H2', fontName=FONT_GO, fontSize=12, leading=17, textColor=PRIME_RED, spaceBefore=5*mm, spaceAfter=3*mm)
sH3 = ParagraphStyle('H3', fontName=FONT_GO, fontSize=10, leading=14, textColor=PRIME_DARK, spaceBefore=3*mm, spaceAfter=2*mm)
sBody = ParagraphStyle('Body', fontName=FONT_MIN, fontSize=9, leading=15, textColor=PRIME_DARK, spaceAfter=2*mm)
sBullet = ParagraphStyle('Bullet', fontName=FONT_MIN, fontSize=8.5, leading=14, textColor=PRIME_DARK, leftIndent=12, spaceAfter=1*mm, bulletIndent=4, bulletFontSize=8)
sSmall = ParagraphStyle('Small', fontName=FONT_MIN, fontSize=7.5, leading=11, textColor=GRAY, spaceAfter=1*mm)
sPageLabel = ParagraphStyle('PageLabel', fontName=FONT_GO, fontSize=8, leading=11, textColor=WHITE)
sFooter = ParagraphStyle('Footer', fontName=FONT_MIN, fontSize=7, leading=10, textColor=GRAY, alignment=1)


def hr():
    return HRFlowable(width="100%", thickness=0.5, color=LIGHT_GRAY, spaceAfter=3*mm, spaceBefore=2*mm)

def sp(h=3):
    return Spacer(1, h*mm)

def page_badge(label, color=PRIME_RED):
    """Create a colored page number/label badge as a table"""
    t = Table([[Paragraph(label, sPageLabel)]], colWidths=[None])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), color),
        ('TEXTCOLOR', (0,0), (-1,-1), WHITE),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ('ROUNDEDCORNERS', [4,4,4,4]),
    ]))
    return t

def feature_table(data, col_widths=None):
    """Create a styled feature table"""
    style = TableStyle([
        ('FONTNAME', (0,0), (-1,0), FONT_GO),
        ('FONTNAME', (0,1), (-1,-1), FONT_MIN),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('LEADING', (0,0), (-1,-1), 13),
        ('BACKGROUND', (0,0), (-1,0), PRIME_DARK),
        ('TEXTCOLOR', (0,0), (-1,0), WHITE),
        ('BACKGROUND', (0,1), (-1,-1), WHITE),
        ('TEXTCOLOR', (0,1), (-1,-1), PRIME_DARK),
        ('GRID', (0,0), (-1,-1), 0.4, LIGHT_GRAY),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [WHITE, HexColor('#FAFAFA')]),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ])
    t = Table(data, colWidths=col_widths, repeatRows=1)
    t.setStyle(style)
    return t

def build_pdf(output_path):
    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        leftMargin=18*mm, rightMargin=18*mm,
        topMargin=20*mm, bottomMargin=18*mm
    )
    W = A4[0] - 36*mm  # usable width

    story = []

    # ===== COVER =====
    story.append(Spacer(1, 30*mm))
    story.append(Paragraph('PRIMECHANGE', ParagraphStyle('CoverBrand', fontName=FONT_GO, fontSize=14, leading=18, textColor=PRIME_RED, spaceAfter=2*mm)))
    story.append(Paragraph('V3 ダッシュボード', sTitle))
    story.append(Paragraph('機能仕様書', ParagraphStyle('CoverSub', fontName=FONT_GO, fontSize=20, leading=28, textColor=PRIME_DARK, spaceAfter=6*mm)))
    story.append(HRFlowable(width="40%", thickness=2, color=PRIME_RED, spaceAfter=8*mm))
    story.append(Paragraph('Portfolio Quality Management &amp; Revenue Analytics Platform', ParagraphStyle('CoverEn', fontName=FONT_MIN, fontSize=10, leading=15, textColor=GRAY, spaceAfter=4*mm)))
    story.append(Paragraph('19ホテルの口コミ・売上・ES（従業員満足度）を横断的に分析し、<br/>経営判断とアクション実行を支援する統合ダッシュボード', sBody))
    story.append(Spacer(1, 20*mm))

    # Overview table
    overview_data = [
        ['項目', '内容'],
        ['対象ホテル数', '19施設'],
        ['分析口コミ数', '1,968件（2025/12 〜 2026/03）'],
        ['ページ数', '10ページ'],
        ['データソース', '6 OTAサイト + 7種分析JSON'],
        ['技術構成', 'Node.js静的ビルド + クライアントJS'],
        ['ブランド', 'PRIMECHANGE コーポレートカラー (#C23B3A)'],
        ['フォント', 'Noto Sans JP'],
        ['レスポンシブ', 'PC / タブレット / モバイル対応'],
    ]
    story.append(feature_table(overview_data, col_widths=[35*mm, W-35*mm]))
    story.append(Spacer(1, 15*mm))
    story.append(Paragraph('作成日: 2026年3月14日', sSmall))
    story.append(Paragraph('\u00a9 2026 PRIME CHANGE Corporation', sSmall))

    story.append(PageBreak())

    # ===== TABLE OF CONTENTS =====
    story.append(Paragraph('目次', sH1))
    story.append(hr())
    toc_items = [
        ('1', 'EXECUTIVE SUMMARY', 'エグゼクティブサマリー', '経営会議用ダッシュボード'),
        ('2', 'PORTAL', 'ポータル', 'KPI概要・緊急対応・トレンド'),
        ('3', 'HOTEL DASHBOARD', 'ホテル別ダッシュボード', '19ホテル個別分析・モーダル詳細'),
        ('4', 'CLEANING STRATEGY', '清掃戦略', 'カテゴリ分析・ヒートマップ・優先度'),
        ('5', 'DEEP ANALYSIS', '深掘り分析', '7種分析タブ + CS6軸 + NPS'),
        ('6', 'REVENUE SIMULATOR', '売上シミュレーター', '回帰分析・What-If・スライダー'),
        ('7', 'OTA STRATEGY', 'OTA戦略', 'サイト横断マトリクス・ギャップ分析'),
        ('8', 'ACTION PLANS', 'アクションプラン', '進捗管理・ステータス・Export/Import'),
        ('9', 'ES DASHBOARD', 'ES管理', 'スタッフ負荷・クレーム集中・人員充足'),
        ('10', 'REVENUE IMPACT', '品質×売上', '回帰分析・シナリオ・ベンチマーク'),
    ]
    toc_data = [['#', 'Page', '機能', '概要']]
    for num, en, ja, desc in toc_items:
        toc_data.append([num, en, ja, desc])
    story.append(feature_table(toc_data, col_widths=[8*mm, 42*mm, 42*mm, W-92*mm]))

    story.append(sp(8))
    story.append(Paragraph('共通機能', sH2))
    common_items = [
        '\u2022 10ページナビゲーション（スティッキーヘッダー、モバイルハンバーガーメニュー）',
        '\u2022 日付フィルター（プリセット + カスタム範囲 + スナップショット切替）',
        '\u2022 変化検知アラートバナー（スコア低下・清掃率悪化を自動検出）',
        '\u2022 レスポンシブデザイン（PC / タブレット / モバイル）',
        '\u2022 印刷対応スタイルシート',
        '\u2022 file:// プロトコル対応（インラインデータ埋め込み）',
    ]
    for item in common_items:
        story.append(Paragraph(item, sBullet))

    story.append(PageBreak())

    # ===== PAGE 1: EXECUTIVE SUMMARY =====
    story.append(page_badge('PAGE 1'))
    story.append(sp(2))
    story.append(Paragraph('EXECUTIVE SUMMARY', sH1))
    story.append(Paragraph('エグゼクティブサマリー \u2014 経営会議用ダッシュボード', sBody))
    story.append(hr())

    story.append(Paragraph('セクション構成', sH3))
    exec_sections = [
        ['セクション', '内容', '表示要素'],
        ['アラートバナー', '変化検知による自動通知', '重要度別カラー（赤/緑/青）、アイコン付き'],
        ['KPI目標進捗', 'ポートフォリオ目標の達成率', '4指標のプログレスバー、現在値\u2192目標値、達成率%、期限'],
        ['売上概要', 'ポートフォリオ全体の財務指標', '月間売上、平均稼働率、月間改善余地（\u00a5）'],
        ['リスクアラート', '緊急対応が必要なホテル', 'TOP5表示、スコア、清掃課題率、問題点、損失額'],
        ['ROIシナリオ', '投資対効果の3パターン分析', '投資額、改善見込、売上効果、回収期間'],
        ['優先アクション', '今月実行すべきアクション', 'URGENT/HIGH優先度、ホテル名、具体施策、損失額'],
        ['ポートフォリオNPS', '顧客ロイヤリティ推定スコア', 'NPSスコア大表示、推奨者/中立者/批判者の内訳%'],
    ]
    story.append(feature_table(exec_sections, col_widths=[30*mm, 50*mm, W-80*mm]))

    story.append(sp(4))
    story.append(Paragraph('KPI進捗の計算ロジック', sH3))
    kpi_logic = [
        '\u2022 目標値の文字列パース: "2.0%以下" \u2192 2.0（数値変換 + 以下/以上判定）',
        '\u2022 達成率 = 現在値 / 目標値 \u00d7 100（"以下"系KPIは逆算ロジック適用）',
        '\u2022 カラー判定: 80%以上=緑、50%以上=橙、50%未満=赤',
        '\u2022 NPS計算: 9-10点=推奨者、7-8点=中立者、1-6点=批判者',
    ]
    for item in kpi_logic:
        story.append(Paragraph(item, sBullet))

    story.append(PageBreak())

    # ===== PAGE 2: PORTAL =====
    story.append(page_badge('PAGE 2'))
    story.append(sp(2))
    story.append(Paragraph('PORTAL', sH1))
    story.append(Paragraph('ポートフォリオ品質管理ポータル', sBody))
    story.append(hr())

    portal_sections = [
        ['セクション', '内容'],
        ['KPIカード (5枚)', 'ホテル数 / 口コミ数 / 平均スコア / 高評価率 / 清掃クレーム率\n各カードに前回比（\u25b2\u25bc）+ 目標達成バッジ'],
        ['ナビリンクカード (6枚)', '各ページへのショートカット（アイコン + 説明文）'],
        ['優先アクション TOP3', '緊急度の高い3ホテルのアクション、月間損失額付き'],
        ['緊急対応ホテル', '優先マトリクスのURGENT/HIGHホテルをカード表示'],
        ['ポートフォリオトレンド', '日次口コミスコアの折れ線チャート'],
        ['SVGゲージ (4基)', '平均スコア / 清掃率 / 高評価率 / 低評価率\n目標達成率%を数値表示'],
    ]
    story.append(feature_table(portal_sections, col_widths=[35*mm, W-35*mm]))

    story.append(sp(4))
    story.append(Paragraph('差分検知（Delta Engine）', sH3))
    delta_items = [
        '\u2022 前回ビルドのスナップショットと自動比較',
        '\u2022 閾値: スコア-0.1pt以上\u2192赤アラート、清掃率+2%以上\u2192赤、スコア+0.05以上\u2192緑',
        '\u2022 KPIカードに\u25b2\u25bc表示 + 色分け',
    ]
    for item in delta_items:
        story.append(Paragraph(item, sBullet))

    story.append(PageBreak())

    # ===== PAGE 3: HOTEL DASHBOARD =====
    story.append(page_badge('PAGE 3'))
    story.append(sp(2))
    story.append(Paragraph('HOTEL DASHBOARD', sH1))
    story.append(Paragraph('ホテル別口コミダッシュボード', sBody))
    story.append(hr())

    story.append(Paragraph('19ホテルをカードグリッドで一覧表示。フィルタ・ソート・検索・モーダル詳細を提供。', sBody))

    hotel_card = [
        ['要素', '表示内容'],
        ['ランクバッジ', '順位（1-3位: ゴールド/ブルー/グレー円形バッジ）'],
        ['スコア表示', '10点満点スコア（大文字） + カラーバー'],
        ['ティアバッジ', '優秀(緑) / 良好(青) / 概ね良好(橙) / 要改善(赤)'],
        ['統計チップ', '高評価率% / 低評価率% / 清掃課題率%'],
        ['OTAサイトドット', 'サイト別スコアを色付きドットで表示'],
        ['売上改善バッジ', '月間改善余地（\u00a5XX万/月）'],
        ['目標ギャップ', '目標スコアとの差分（\u00b1X.XX）'],
    ]
    story.append(Paragraph('ホテルカード構成', sH3))
    story.append(feature_table(hotel_card, col_widths=[30*mm, W-30*mm]))

    story.append(sp(3))
    story.append(Paragraph('インタラクティブ機能', sH3))
    hotel_interactive = [
        '\u2022 ティアフィルター: 全て / 優秀 / 良好 / 概ね良好 / 要改善',
        '\u2022 ホテル名検索: リアルタイムインクリメンタル検索',
        '\u2022 ソート: ランク / 口コミ数 / 高評価率 / 低評価率 / 清掃率',
        '\u2022 モーダル詳細: カードクリックで詳細表示（サイト別評価、スコア分布、口コミ一覧、トレンドチャート）',
    ]
    for item in hotel_interactive:
        story.append(Paragraph(item, sBullet))

    story.append(PageBreak())

    # ===== PAGE 4: CLEANING STRATEGY =====
    story.append(page_badge('PAGE 4'))
    story.append(sp(2))
    story.append(Paragraph('CLEANING STRATEGY', sH1))
    story.append(Paragraph('清掃品質改善戦略', sBody))
    story.append(hr())

    cleaning_sections = [
        ['セクション', '内容'],
        ['KPIカード (4枚)', '清掃クレーム率 / クレーム件数 / カテゴリ数 / 緊急対応ホテル数'],
        ['カテゴリ別バーチャート', '清掃課題カテゴリ（汚れ、髪の毛、カビ等）の出現頻度\n重要度カラー + 影響ホテル数 + 推定損失額/月'],
        ['ホテル\u00d7カテゴリ ヒートマップ', '行:ホテル、列:カテゴリ、セル:指摘件数\n色グラデーション: 白(0)\u2192黄(1)\u2192橙(3-5)\u2192赤(5+)'],
        ['優先度マトリクス', 'URGENT(赤) / HIGH(橙) / STANDARD(青) / MAINTENANCE(緑)\n各ホテルのスコア、課題率、重要問題'],
        ['横断施策提言', '全社共通の改善施策（タイトル + 具体アクション）'],
    ]
    story.append(feature_table(cleaning_sections, col_widths=[40*mm, W-40*mm]))

    story.append(PageBreak())

    # ===== PAGE 5: DEEP ANALYSIS =====
    story.append(page_badge('PAGE 5'))
    story.append(sp(2))
    story.append(Paragraph('DEEP ANALYSIS', sH1))
    story.append(Paragraph('深掘り分析（7種分析 + CS6軸 + NPS）', sBody))
    story.append(hr())

    story.append(Paragraph('7つのタブ切替式分析', sH3))
    analysis_tabs = [
        ['分析#', 'テーマ'],
        ['分析1', '口コミトレンド・時系列分析'],
        ['分析2', 'スタッフ別パフォーマンス分析'],
        ['分析3', '人員配置最適化分析'],
        ['分析4', '時間帯別パフォーマンス分析'],
        ['分析5', '清掃品質詳細分析'],
        ['分析6', '品質\u2192売上弾力性分析（回帰）'],
        ['分析7', '競合ベンチマーク分析'],
    ]
    story.append(feature_table(analysis_tabs, col_widths=[20*mm, W-20*mm]))

    story.append(sp(3))
    story.append(Paragraph('CS（顧客満足度）6軸分析', sH3))
    cs_items = [
        '\u2022 6軸キーワード分類: 接客態度 / 立地 / 朝食 / 設備 / 清掃 / コスパ',
        '\u2022 6軸\u00d719ホテル マトリクス（+ポジティブ / -ネガティブ件数、色分け）',
        '\u2022 キーワード頻度 TOP20 バーチャート（カテゴリバッジ付き）',
        '\u2022 ポートフォリオNPS表示（推奨者/中立者/批判者の内訳）',
    ]
    for item in cs_items:
        story.append(Paragraph(item, sBullet))

    story.append(PageBreak())

    # ===== PAGE 6: REVENUE SIMULATOR =====
    story.append(page_badge('PAGE 6', BLUE))
    story.append(sp(2))
    story.append(Paragraph('REVENUE SIMULATOR', sH1))
    story.append(Paragraph('売上シミュレーター \u2014 口コミスコア改善による売上インパクト推計', sBody))
    story.append(hr())

    sim_sections = [
        ['セクション', '内容', 'データ'],
        ['回帰分析サマリー', '3指標の回帰係数を大型カード表示', 'Score\u2192RevPAR (+229\u00a5/pt)\nScore\u2192ADR (+255\u00a5/pt)\nScore\u2192稼働率 (+3.8%/pt)'],
        ['スコア帯別テーブル', '4スコア帯のパフォーマンス比較', 'スコア帯 / ホテル数 / 稼働率 / ADR / RevPAR'],
        ['改善機会ランキング', '19ホテルの損失額順位表', '現在スコア / 目標 / GAP / 客室数 / 月間損失'],
        ['What-Ifシナリオ', '3パターンのシミュレーション', '\u2460全ホテル+0.5点\n\u2461下位5ホテル引上げ\n\u2462清掃問題30%削減'],
        ['インタラクティブスライダー', 'スコア改善幅\u2192月間売上増加をリアルタイム計算', '0.0〜2.0点（0.1刻み）\nRevPAR係数\u00d7客室数\u00d730日'],
    ]
    story.append(feature_table(sim_sections, col_widths=[32*mm, 50*mm, W-82*mm]))

    story.append(sp(3))
    story.append(Paragraph('売上機会損失の計算式', sH3))
    story.append(Paragraph('月間損失 = (目標スコア - 現在スコア) \u00d7 RevPAR回帰係数(228.9) \u00d7 客室数 \u00d7 30日', sBody))
    story.append(Paragraph('\u203b R\u00b2=35.7%（N=19）の回帰分析に基づく推計値。因果関係ではなく相関関係である点に留意。', sSmall))

    story.append(PageBreak())

    # ===== PAGE 7: OTA STRATEGY =====
    story.append(page_badge('PAGE 7', BLUE))
    story.append(sp(2))
    story.append(Paragraph('OTA STRATEGY', sH1))
    story.append(Paragraph('OTA戦略分析 \u2014 サイト横断パフォーマンス', sBody))
    story.append(hr())

    ota_sections = [
        ['セクション', '内容'],
        ['OTAサイト別サマリー (6枚)', '各OTAの平均スコア・口コミ件数・掲載ホテル数\n対象: Google, Booking.com, Expedia, 楽天トラベル, じゃらん, 一休.com'],
        ['サイト\u00d7ホテル\nクロスマトリクス', '19ホテル \u00d7 6サイトのヒートマップテーブル\nセル色: 緑(8+) / 黄(5-8) / 赤(<5) / グレー(データなし)\n全体平均列付き'],
        ['サイト別ランキング', 'OTAサイトごとのTOP3 / BOTTOM3ホテル\nスコア + ティア色表示'],
        ['ギャップ分析', 'ホテルごとの最高/最低サイトのスコア差\nGAP 1.0以上を黄色ハイライト\nGAPカラー: 緑(<1.0) / 橙(1.0-2.0) / 赤(>2.0)'],
    ]
    story.append(feature_table(ota_sections, col_widths=[35*mm, W-35*mm]))

    story.append(sp(3))
    story.append(Paragraph('分析の目的', sH3))
    ota_purpose = [
        '\u2022 OTAサイト間での評価一貫性を可視化し、サイト別対策の優先度を特定',
        '\u2022 特定サイトで低スコアのホテルに対し、写真・説明文・返信対応などサイト固有の改善施策を立案',
        '\u2022 ポートフォリオ全体のOTA戦略（掲載最適化・プラン設定・レビュー管理）を策定',
    ]
    for item in ota_purpose:
        story.append(Paragraph(item, sBullet))

    story.append(PageBreak())

    # ===== PAGE 8: ACTION PLANS =====
    story.append(page_badge('PAGE 8', ORANGE))
    story.append(sp(2))
    story.append(Paragraph('ACTION PLANS', sH1))
    story.append(Paragraph('アクションプラン管理', sBody))
    story.append(hr())

    action_sections = [
        ['セクション', '内容'],
        ['進捗サマリーバー', '全アクションの完了率（プログレスバー + %表示）'],
        ['優先度KPI (4枚)', 'URGENT / HIGH / STANDARD / MAINTENANCE の件数'],
        ['フィルター', 'ステータス別 / ホテル別 / フェーズ別 の3軸フィルタリング'],
        ['KPI目標テーブル', 'ポートフォリオKPI: 指標 / 現在値 / 目標値 / 期限'],
        ['ホテル別アコーディオン', '優先度順に19ホテル表示\n各ホテルに3フェーズのアクション一覧'],
    ]
    story.append(feature_table(action_sections, col_widths=[35*mm, W-35*mm]))

    story.append(sp(3))
    story.append(Paragraph('3フェーズ構成', sH3))
    phase_data = [
        ['フェーズ', '期間', 'ROI配分', 'カラー'],
        ['Phase 1: 即時対応', '〜1ヶ月', '改善余地の50%', '赤 (#C23B3A)'],
        ['Phase 2: 短期施策', '1〜3ヶ月', '改善余地の30%', '橙 (#F59E0B)'],
        ['Phase 3: 中期施策', '3〜6ヶ月', '改善余地の20%', '青 (#3B82F6)'],
    ]
    story.append(feature_table(phase_data, col_widths=[35*mm, 25*mm, 35*mm, 30*mm]))

    story.append(sp(3))
    story.append(Paragraph('アクション管理機能', sH3))
    action_features = [
        '\u2022 ステータス管理: 未着手 \u2192 進行中 \u2192 完了（クリックで切替）',
        '\u2022 担当者・期限フィールド',
        '\u2022 LocalStorage保存（ブラウザ内永続化）',
        '\u2022 JSON Export / Import 機能（チーム間共有）',
    ]
    for item in action_features:
        story.append(Paragraph(item, sBullet))

    story.append(PageBreak())

    # ===== PAGE 9: ES DASHBOARD =====
    story.append(page_badge('PAGE 9', GREEN))
    story.append(sp(2))
    story.append(Paragraph('ES DASHBOARD', sH1))
    story.append(Paragraph('ES（従業員満足度）ダッシュボード', sBody))
    story.append(hr())

    es_sections = [
        ['セクション', '内容'],
        ['KPIカード (5枚)', '総スタッフ数 / 平均出勤日数 / 平均清掃室数/日 / クレーム有りメイド数 / 最大清掃室数/日'],
        ['スタッフ負荷分析', 'ホテル別 rooms/person/day の水平バーチャート\n色: 赤(>15室) / 橙(10-15室) / 緑(<10室)'],
        ['出勤率分析', '平均出勤日数 / 高出勤スタッフ数 / 総スタッフ数'],
        ['クレーム集中リスク', 'TOP15メイド + TOP15チェッカーのクレーム件数テーブル\n上位3名を赤ハイライト'],
        ['人員充足度', 'TOP5 vs BOTTOM5ホテルのスタッフ配置比較\n全ホテル一覧（出勤日数/メイド数/チェッカー数/比率/スコア/クレーム率）'],
        ['相関分析 (6枚)', 'スタッフ数 vs スコアの相関係数\nr値 / R\u00b2 / 相関強度バッジ（強/中/弱）'],
        ['ES改善提言', '優先度付きアコーディオン（HIGH/MEDIUM/LOW）\n根拠 + 具体アクション'],
    ]
    story.append(feature_table(es_sections, col_widths=[32*mm, W-32*mm]))

    story.append(PageBreak())

    # ===== PAGE 10: REVENUE IMPACT =====
    story.append(page_badge('PAGE 10', ORANGE))
    story.append(sp(2))
    story.append(Paragraph('REVENUE IMPACT', sH1))
    story.append(Paragraph('品質\u00d7売上 インパクト分析', sBody))
    story.append(hr())

    rev_sections = [
        ['セクション', '内容'],
        ['ポートフォリオKPI (5枚)', '月間総売上 / 平均稼働率 / 平均ADR / 平均RevPAR / 総改善余地'],
        ['回帰分析テーブル', '指標 / 回帰係数 / 相関r / R\u00b2 / 解釈'],
        ['スコア帯別パフォーマンス', 'スコア帯ごとの稼働率・ADR・RevPAR・月間売上比較 + 閾値効果説明'],
        ['売上インパクトシナリオ', '+0.1pt / +0.2pt等の改善シナリオ\nRevPAR変動 / 月間売上変動 / 年間売上変動'],
        ['ホテル別改善機会テーブル', '19ホテルの現在スコア / 目標 / GAP / 客室数 / 月間売上 / 改善余地'],
        ['業界ベンチマーク比較', '業界平均との比較分析'],
        ['改善ポテンシャルテーブル', 'ホテル別の優先度 + 月間/年間インパクト額'],
        ['改善提言', '優先度付き提言（アコーディオン）'],
    ]
    story.append(feature_table(rev_sections, col_widths=[38*mm, W-38*mm]))

    story.append(PageBreak())

    # ===== ARCHITECTURE =====
    story.append(Paragraph('システムアーキテクチャ', sH1))
    story.append(hr())

    story.append(Paragraph('ビルドパイプライン', sH2))
    arch_data = [
        ['レイヤー', 'ファイル', '役割'],
        ['データ読込', 'data-loader.js', '全JSON読込・KPI目標パーサー・データ正規化'],
        ['差分計算', 'delta-engine.js', '前回スナップショットとの差分計算・アラート生成'],
        ['売上計算', 'revenue-calc.js', '売上機会損失計算（回帰係数\u00d7客室数\u00d730日）'],
        ['CS分析', 'cs-analyzer.js', '6軸キーワード分類・NPS推定・頻度分析'],
        ['共通基盤', 'common-v2.js', 'HTML/CSS生成、ナビ、ヘッダー/フッター'],
        ['ページ生成', 'page-*.js (10ファイル)', '各ページのHTML生成器'],
        ['オーケストレーター', 'build_all_v2.js', '全フェーズ統括・スナップショット管理'],
    ]
    story.append(feature_table(arch_data, col_widths=[25*mm, 40*mm, W-65*mm]))

    story.append(sp(4))
    story.append(Paragraph('クライアントサイドJS', sH2))
    client_data = [
        ['スクリプト', '機能'],
        ['date-filter-v2.js', '日付フィルターロジック（口コミデータの範囲指定、KPI再計算）'],
        ['date-filter-ui-v2.js', '日付フィルターUI（プリセットボタン、カレンダー入力、適用ボタン）'],
        ['hotel-dashboard-v2.js', 'ホテルカードのフィルタ・ソート・モーダル詳細表示'],
        ['svg-gauge-v2.js', 'SVGゲージ描画（達成率%表示、アニメーション）'],
        ['trend-chart-v2.js', '折れ線トレンドチャート描画（目標ライン付き）'],
        ['alert-engine.js', '変化検知アラートバナーの表示制御'],
        ['action-tracker.js', 'アクションステータス管理（LocalStorage + Export/Import）'],
        ['page-snapshot-v2.js', 'スナップショット切替（過去データの閲覧）'],
        ['cs-charts.js', 'CS分析チャート・マトリクス描画'],
        ['es-dashboard.js', 'ESページのインタラクション制御'],
    ]
    story.append(feature_table(client_data, col_widths=[38*mm, W-38*mm]))

    story.append(PageBreak())

    # ===== DESIGN SYSTEM =====
    story.append(Paragraph('デザインシステム', sH1))
    story.append(hr())

    story.append(Paragraph('カラーパレット', sH2))
    color_data = [
        ['用途', 'カラー', 'コード'],
        ['プライマリ', 'PRIME RED', '#C23B3A'],
        ['ダーク', 'PRIME DARK', '#1A1A2E'],
        ['ライト', 'PRIME LIGHT', '#F8F4F4'],
        ['成功・優秀', 'GREEN', '#10B981'],
        ['警告・概ね良好', 'ORANGE', '#F59E0B'],
        ['エラー・要改善', 'RED', '#EF4444'],
        ['情報・良好', 'BLUE', '#3B82F6'],
        ['テキスト補助', 'GRAY', '#64748B'],
    ]
    story.append(feature_table(color_data, col_widths=[35*mm, 35*mm, 30*mm]))

    story.append(sp(4))
    story.append(Paragraph('ティア判定基準', sH2))
    tier_data = [
        ['ティア', 'スコア範囲', 'カラー', '対応方針'],
        ['優秀', '9.0以上', '#10B981 (緑)', '維持・ベストプラクティス共有'],
        ['良好', '8.5〜9.0', '#3B82F6 (青)', '微調整で上位到達可能'],
        ['概ね良好', '8.0〜8.5', '#F59E0B (橙)', '重点改善対象'],
        ['要改善', '8.0未満', '#EF4444 (赤)', '緊急対応必要'],
    ]
    story.append(feature_table(tier_data, col_widths=[20*mm, 25*mm, 30*mm, W-75*mm]))

    story.append(sp(4))
    story.append(Paragraph('UIコンポーネント', sH2))
    ui_components = [
        '\u2022 カード: 白背景、12px角丸、シャドウ、カラーボーダー付きバリエーション',
        '\u2022 テーブル: ダークヘッダー(#1A1A2E)、交互行、ホバーエフェクト',
        '\u2022 バッジ: ステータス / 優先度 / 達成度 / 売上の4種類',
        '\u2022 プログレスバー: カラーコード付き達成率表示',
        '\u2022 アコーディオン: クリック展開、矢印回転アニメーション',
        '\u2022 モーダル: オーバーレイ背景、中央配置、閉じるボタン',
        '\u2022 アラートバナー: 3段階色分け（危険赤 / 改善緑 / 情報青）',
        '\u2022 ゲージ: SVG円弧型、達成率数値付き',
    ]
    for item in ui_components:
        story.append(Paragraph(item, sBullet))

    # Footer
    story.append(Spacer(1, 15*mm))
    story.append(HRFlowable(width="100%", thickness=1, color=PRIME_RED, spaceAfter=4*mm))
    story.append(Paragraph('PRIMECHANGE V3 Dashboard \u2014 \u00a9 2026 PRIME CHANGE Corporation', sFooter))

    doc.build(story)
    print(f'PDF generated: {output_path}')

if __name__ == '__main__':
    build_pdf('/Users/mitsugutakahashi/ホテル口コミ/PRIMECHANGE_V3_機能仕様書.pdf')
