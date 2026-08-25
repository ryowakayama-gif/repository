# -*- coding: utf-8 -*-
"""
出力ファイル名と文字エンコードの規則（環境基本計画 営業資料シリーズ）

背景:
  文字コード調査報告（encodinginvestigation.md）の結論を、環境基本計画の
  営業資料（営業整理v8／営業先リスト／士幌町レビュー）へ適用したもの。

  同報告の結論:
    - リポジトリのエンコードは全レイヤーUTF-8であり、UTF-8化の修正対象はない
    - 行政系で弾かれる要因は ①出力ファイル名が日本語（非ASCII） ②CSVの文字コード
      （行政系はUTF-8ではなくShift_JIS(CP932)指定が多い）の2点
  したがって本モジュールは「UTF-8へ直す」ためのものではなく、
  「提出先の要求に合わせて切り替える」ためのものである。

規則:
  - .xlsx の内部XMLはOOXML仕様上UTF-8固定。openpyxlに選択の余地はなく、対象外
  - テキスト出力（CSV）は必ず本モジュールの write_csv() を通す
  - open() で直接テキストを書き出す場合は必ず encoding= を明示する
    （Windows日本語環境では locale.getpreferredencoding() が cp932 を返すため）
"""

import csv
import io
import os
import sys

# ------------------------------------------------------------
# ブック名の 日本語 / 半角英数 対応表
#   新しいブックを追加したら、ここにも追記すること
# ------------------------------------------------------------
BOOKS = {
    "環境基本計画_営業先リスト_北海道東北": "env_plan_sales_list_hokkaido_tohoku",
    "環境基本計画_営業整理_v8": "env_plan_sales_pack_v8",
    "士幌町環境基本計画_レビューと提案": "shihoro_env_plan_review",
}

# ------------------------------------------------------------
# シート名の半角英数スラッグ（CSVファイル名に使う）
#   未登録のシートは sheet_slug() が機械的に生成するため、
#   未登録でも非ASCIIファイル名にはならない
# ------------------------------------------------------------
SHEET_SLUGS = {
    "00_概要": "00_overview",
    "01_R9満了候補": "01_r9_expiry",
    "02_営業先マスタ": "02_target_master",
    "03_確認手順": "03_verify_steps",
    "04_アプローチ設計": "04_approach",
    "05_更新メモ": "05_update_log",
    "06_町村改訂時期_20260825": "06_town_revision_20260825",
    "07_提出仕様_文字コード": "07_submission_encoding",
    "01_財源整理": "01_funding",
    "02_国の動向": "02_national_trends",
    "03_県の動向": "03_pref_trends",
    "04_市場再整理": "04_market",
    "04b_市場数試算": "04b_market_estimate",
    "05_現行計画チェック": "05_plan_check",
    "06_ヒアリング項目": "06_hearing_items",
    "07_営業先優先度": "07_priority",
    "08_出典URL": "08_sources",
    "00_サマリー": "00_summary",
    "01_計画チェック30項目": "01_plan_check_30",
    "02_レビュー指摘": "02_findings",
    "03_財源の確認": "03_funding_check",
    "04_施策内容の確認": "04_policy_check",
    "05_提案メニュー": "05_proposals",
    "06_アプローチ": "06_approach",
    "07_確認事項と出典": "07_todo_sources",
}

# ------------------------------------------------------------
# CP932で表現できない文字の代替表
#   調査時点で本シリーズに実在したのは '—'(U+2014) のみだが、
#   将来の追記に備えて一般的なものを登録している
# ------------------------------------------------------------
CP932_SUBSTITUTIONS = {
    "—": "―",   # U+2014 EM DASH        -> U+2015 HORIZONTAL BAR
    "–": "-",   # U+2013 EN DASH
    "−": "-",   # U+2212 MINUS SIGN
    "✓": "*",   # U+2713 CHECK MARK
    "✔": "*",   # U+2714
    "☑": "*",   # U+2611
    "№": "No.",  # U+2116
    "⇒": "=>",  # U+21D2
    "•": "・",   # U+2022 BULLET
    "…": "…",   # U+2026（cp932可。表に残すのは意図の明示のため）
}

DEFAULT_CSV_ENCODING = "cp932"


def ascii_mode(argv=None):
    """半角英数ファイル名モードか。--ascii オプションまたは ASCII_FILENAMES=1。"""
    argv = sys.argv if argv is None else argv
    if "--ascii" in argv:
        return True
    return os.environ.get("ASCII_FILENAMES", "") == "1"


def csv_encoding():
    """CSVの文字コード。既定は cp932。環境変数 CSV_ENCODING で変更する。"""
    return os.environ.get("CSV_ENCODING", DEFAULT_CSV_ENCODING)


def book_stem(japanese_stem, ascii_names=False):
    """ブックのファイル名（拡張子なし）を返す。未登録名はそのまま返す。"""
    if not ascii_names:
        return japanese_stem
    slug = BOOKS.get(japanese_stem)
    if slug:
        return slug
    # 未登録でも非ASCIIにしないための機械生成
    return _mechanical_slug(japanese_stem, "book")


def sheet_slug(sheet_name, index=0):
    """シート名の半角英数スラッグ。未登録でも必ずASCIIを返す。"""
    slug = SHEET_SLUGS.get(sheet_name)
    if slug:
        return slug
    return _mechanical_slug(sheet_name, f"sheet{index:02d}")


def _mechanical_slug(text, fallback):
    """先頭の半角英数・記号だけを拾い、無ければ fallback を返す。"""
    kept = "".join(ch for ch in text if ch.isascii() and (ch.isalnum() or ch in "-_"))
    kept = kept.strip("-_")
    return kept if kept else fallback


def sanitize(text, encoding=None):
    """
    指定エンコードで表現できない文字を代替表で置換する。
    戻り値: (置換後の文字列, {元の文字: 代替文字})
    代替表にも無い場合は '?' へ落とす。
    """
    encoding = encoding or csv_encoding()
    if text is None:
        return "", {}
    text = str(text)
    replaced = {}
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
        replaced[ch] = alt
        out.append(alt)
    return "".join(out), replaced


def check_encodable(values, encoding=None):
    """
    値の並びを検査し、指定エンコードで表現できない文字を集計して返す。
    戻り値: {文字: 出現数}
    """
    encoding = encoding or csv_encoding()
    bad = {}
    for v in values:
        if not isinstance(v, str):
            continue
        for ch in v:
            try:
                ch.encode(encoding)
            except UnicodeEncodeError:
                bad[ch] = bad.get(ch, 0) + 1
    return bad


def write_csv(path, rows, encoding=None, quiet=False, flatten_newlines=True):
    """
    文字コードと改行コードを制御してCSVを書き出す。
      - 文字コード: 既定 cp932（環境変数 CSV_ENCODING で変更）
      - 改行コード: CRLF固定（Excel・行政系システムの慣例）
      - セル内改行: 既定で半角スペースへ畳む（flatten_newlines=False で保持）
        引用符で囲めばRFC4180上は正しいが、行政系の取り込みで行数がずれる
        原因になりやすいため、提出用では畳むことを既定とする
      - 表現できない文字: CP932_SUBSTITUTIONS で置換し、内容を標準出力へ警告
    戻り値: 置換した {元の文字: 代替文字}
    """
    encoding = encoding or csv_encoding()
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)

    replaced_all = {}
    buf = io.StringIO(newline="")
    writer = csv.writer(buf, lineterminator="\r\n")
    for row in rows:
        clean = []
        for cell in row:
            if cell is None:
                clean.append("")
                continue
            if isinstance(cell, (int, float)):
                clean.append(cell)
                continue
            text, replaced = sanitize(cell, encoding)
            replaced_all.update(replaced)
            if flatten_newlines:
                text = text.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
            clean.append(text)
        writer.writerow(clean)

    with open(path, "w", encoding=encoding, newline="") as f:
        f.write(buf.getvalue())

    if replaced_all and not quiet:
        detail = "、".join(f"{k!r}→{v!r}" for k, v in replaced_all.items())
        print(f"! {encoding} で表現できない文字を置換しました: {detail}  ({os.path.basename(path)})")
    return replaced_all


def mode_label(ascii_names=None):
    """現在のモードを人が読める形で返す（エラーメッセージ用）。"""
    ascii_names = ascii_mode() if ascii_names is None else ascii_names
    return f"ファイル名: {'半角英数' if ascii_names else '日本語'} / CSV文字コード: {csv_encoding()}"
