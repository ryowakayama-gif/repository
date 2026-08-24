# -*- coding: utf-8 -*-
"""
出力ファイル名と文字エンコードの共通設定。

行政系システムへの提出では、次の2点を要求されることが多い。
  1. ファイル名が半角英数字（ASCII）のみであること
  2. CSVの文字コードが Shift_JIS(CP932) であること

本モジュールはその2点を切り替え可能にし、build_excel.py と
build_component_images.py の両方から同じ命名規則を参照させる。

使い方:
    python3 build_excel.py            # 従来どおり日本語ファイル名
    python3 build_excel.py --ascii    # 半角英数字ファイル名で出力
    ASCII_FILENAMES=1 python3 build_excel.py   # 環境変数でも指定可

CSVの文字コード:
    CSV_ENCODING=cp932     （既定）Shift_JIS。行政系CSVで最も多い指定
    CSV_ENCODING=utf-8-sig BOM付きUTF-8。UTF-8指定かつExcelで開く場合
    CSV_ENCODING=utf-8     BOMなしUTF-8
"""

import csv
import os
import sys

OUT_DIR = "/home/user/repository/output"
IMG_DIR = os.path.join(OUT_DIR, "images_basic")
CSV_DIR = os.path.join(OUT_DIR, "csv")


# ============================================================
# ASCIIファイル名モードの判定
# ============================================================
def ascii_filenames_enabled(argv=None):
    """--ascii / --no-ascii / 環境変数 ASCII_FILENAMES からモードを決める。"""
    argv = sys.argv[1:] if argv is None else argv
    if "--no-ascii" in argv:
        return False
    if "--ascii" in argv:
        return True
    return os.environ.get("ASCII_FILENAMES", "").strip().lower() in ("1", "true", "yes", "on")


ASCII_FILENAMES = ascii_filenames_enabled()


# ============================================================
# ブック名（キー -> (日本語名, 半角英数名)）
# ============================================================
BOOKS = {
    "master":     ("00_全計画マスター管理表.xlsx",   "00_master_all_plans.xlsx"),
    "common":     ("01_共通_基本コラム部品.xlsx",     "01_common_basic_columns.xlsx"),
    "senior":     ("02_高齢者介護保険事業計画.xlsx",  "02_senior_care_plan.xlsx"),
    "disability": ("03_障がい福祉計画.xlsx",          "03_disability_welfare_plan.xlsx"),
    "child":      ("04_こども計画.xlsx",              "04_child_plan.xlsx"),
}


def book_filename(key):
    """ブックのファイル名（拡張子込み）を返す。"""
    ja, en = BOOKS[key]
    return en if ASCII_FILENAMES else ja


def book_path(key):
    """ブックの絶対パスを返す。"""
    return os.path.join(OUT_DIR, book_filename(key))


# ============================================================
# 部品画像名（部品ID -> 半角英数スラッグ）
# ============================================================
PART_SLUGS = {
    "BC-01": "point",
    "BC-02": "column",
    "BC-03": "case_study",
    "BC-04": "explanation",
    "BC-05": "data_guide",
    "BC-06": "caution",
    "SR-01": "dementia_support",
    "SR-02": "care_prevention",
    "SR-03": "home_living",
    "DS-01": "reasonable_accommodation",
    "DS-02": "communication_support",
    "DS-03": "community_inclusion",
    "CH-01": "childrens_voice",
    "CH-02": "parenting_tips",
    "CH-03": "community_growth",
}


def image_filename(part_id, part_name, ext=".png"):
    """部品見本画像のファイル名を返す。"""
    if ASCII_FILENAMES:
        # 未登録IDでも必ずASCIIになるようフォールバックする
        slug = PART_SLUGS.get(part_id, part_id.lower().replace("-", "_"))
        return f"{part_id}_{slug}{ext}"
    return f"{part_id}_{part_name}{ext}"


def image_path(part_id, part_name, ext=".png"):
    return os.path.join(IMG_DIR, image_filename(part_id, part_name, ext))


# ============================================================
# CSV出力（既定 Shift_JIS / CP932）
# ============================================================
CSV_ENCODING = os.environ.get("CSV_ENCODING", "cp932").strip() or "cp932"

# CP932に存在しないが、対応する全角文字に置き換えれば通る文字。
# Unicode正規化のゆれ（波ダッシュ問題など）を吸収する。
CP932_SUBSTITUTIONS = {
    "〜": "～",  # 〜 WAVE DASH        -> ～ FULLWIDTH TILDE
    "−": "－",  # − MINUS SIGN        -> － FULLWIDTH HYPHEN-MINUS
    "—": "―",  # — EM DASH           -> ― HORIZONTAL BAR
    "‖": "∥",  # ‖ DOUBLE VERTICAL   -> ∥ PARALLEL TO
    "¢": "￠",  # ¢                   -> ￠
    "£": "￡",  # £                   -> ￡
    "¬": "￢",  # ¬                   -> ￢
    "✓": "*",       # ✓ CHECK MARK        -> *（CP932に該当なし）
    "✔": "*",       # ✔                   -> *
    "‐": "-",       # ‐ HYPHEN            -> -
}


def sanitize_for_encoding(text, encoding=None, report=None):
    """指定エンコードで表現できない文字を安全な文字へ置き換える。

    置き換えが発生した文字は report（set）へ記録し、呼び出し側で警告できる。
    """
    encoding = encoding or CSV_ENCODING
    if not isinstance(text, str):
        return text
    try:
        text.encode(encoding)
        return text
    except UnicodeEncodeError:
        pass

    out = []
    for ch in text:
        try:
            ch.encode(encoding)
            out.append(ch)
            continue
        except UnicodeEncodeError:
            pass
        alt = CP932_SUBSTITUTIONS.get(ch, "?")
        try:
            alt.encode(encoding)
        except UnicodeEncodeError:
            alt = "?"
        if report is not None:
            report.add((ch, alt))
        out.append(alt)
    return "".join(out)


def write_csv(path, header, rows, encoding=None):
    """行政系提出を想定したCSVを書き出す。

    - 既定の文字コードは Shift_JIS(CP932)
    - 改行は CRLF（Excel／行政系システムの慣例）
    - 表現できない文字は置き換え、件数を標準出力へ警告
    """
    encoding = encoding or CSV_ENCODING
    os.makedirs(os.path.dirname(path), exist_ok=True)

    report = set()
    safe_header = [sanitize_for_encoding(h, encoding, report) for h in header]
    safe_rows = [
        [sanitize_for_encoding("" if v is None else str(v), encoding, report) for v in row]
        for row in rows
    ]

    # newline="" は csv モジュールの作法。改行コードは lineterminator で指定する
    with open(path, "w", encoding=encoding, newline="") as f:
        writer = csv.writer(f, lineterminator="\r\n")
        writer.writerow(safe_header)
        writer.writerows(safe_rows)

    if report:
        detail = "、".join(f"{c!r}→{a!r}" for c, a in sorted(report))
        print(f"    ! {encoding} で表現できない文字を置換しました: {detail}")
    return path


def describe_mode():
    """現在の出力モードを1行で説明する。"""
    name_mode = "半角英数（ASCII）" if ASCII_FILENAMES else "日本語"
    return f"ファイル名: {name_mode} / CSV文字コード: {CSV_ENCODING}"
