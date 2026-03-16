#!/usr/bin/env python3
"""Shared utilities for hotel XLSX data extraction."""

import os
import datetime

DOWNLOADS_DIR = os.path.expanduser("~/Downloads")

HOTEL_FILES = {
    'daiwa_osaki': 'R8_P1_ダイワロイネットホテル東京大崎_集計表.xlsx',
    'chisan': 'R8_P1_チサンホテル浜松町_集計表.xlsx',
    'hearton': 'R8_P1_ハートンホテル東品川_集計表.xlsx',
    'keyakigate': 'R8_P1_ホテルケヤキゲート東京府中_集計表.xlsx',
    'richmond_mejiro': 'R8_P1_リッチモンドホテル東京目白_集計表.xlsx',
    'keisei_richmond': 'R8_P1_京成リッチモンドホテル東京錦糸町_集計表.xlsx',
    'daiichi_ikebukuro': 'R8_P1_第一イン池袋_集計表.xlsx',
    'comfort_roppongi': 'R8_P2_コンフォートイン六本木_集計表.xlsx',
    'comfort_suites_tokyobay': 'R8_P2_コンフォートスイーツ東京ベイ_集計表.xlsx',
    'comfort_era_higashikanda': 'R8_P2_コンフォートホテルERA東京東神田_集計表.xlsx',
    'comfort_narita': 'R8_P2_コンフォートホテル成田_集計表.xlsx',
    'comfort_yokohama': 'R8_P2_コンフォートホテル横浜関内_集計表.xlsx',
    'apa_sagamihara': 'R8_P3_アパホテル相模原橋本駅東_集計表.xlsx',
    'apa_kamata': 'R8_P3_アパホテル蒲田駅東_集計表.xlsx',
    'court_shinyokohama': 'R8_P3_コートホテル新横浜_集計表.xlsx',
    'comment_yokohama': 'R8_P3_ホテルコメント横浜関内_集計表.xlsx',
    'henn_na_haneda': 'R8_P3_変なホテル東京羽田_集計表.xlsx',
    'kawasaki_nikko': 'R8_P3_川崎日航ホテル_集計表.xlsx',
    'comfort_hakata': 'R8_P4_コンフォートホテル博多_集計表.xlsx',
}

# Map quality JSON keys to revenue JSON keys
KEY_MAP = {
    'keisei_kinshicho': 'keisei_richmond',
    'comfort_yokohama_kannai': 'comfort_yokohama',
}


def get_xlsx_path(hotel_key):
    """Get the full path to a hotel's XLSX file."""
    fname = HOTEL_FILES.get(hotel_key)
    if not fname:
        return None
    return os.path.join(DOWNLOADS_DIR, fname)


def open_workbook(hotel_key):
    """Open a hotel's XLSX workbook with openpyxl (data_only mode)."""
    import openpyxl
    path = get_xlsx_path(hotel_key)
    if not path or not os.path.exists(path):
        raise FileNotFoundError(f"XLSX not found for {hotel_key}: {path}")
    return openpyxl.load_workbook(path, data_only=True, read_only=True)


def safe_number(val, default=0):
    """Convert cell value to number, handling None/empty/string."""
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return val
    if isinstance(val, str):
        val = val.strip().replace(',', '')
        if not val:
            return default
        try:
            return float(val)
        except ValueError:
            return default
    return default


def safe_time(val):
    """Convert cell value to decimal hours (e.g., 14:30 → 14.5)."""
    if val is None:
        return None
    if isinstance(val, datetime.time):
        return val.hour + val.minute / 60.0
    if isinstance(val, datetime.datetime):
        return val.hour + val.minute / 60.0
    if isinstance(val, (int, float)):
        # Might be stored as fraction of day (Excel format)
        if 0 < val < 1:
            hours = val * 24
            return hours
        return val
    if isinstance(val, str):
        val = val.strip()
        if ':' in val:
            parts = val.split(':')
            try:
                return int(parts[0]) + int(parts[1]) / 60.0
            except (ValueError, IndexError):
                return None
    return None


def revenue_key(quality_key):
    """Map a quality data key to the corresponding revenue data key."""
    return KEY_MAP.get(quality_key, quality_key)
