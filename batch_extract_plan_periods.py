# -*- coding: utf-8 -*-
"""
計画PDFのフォルダを一括処理し、計画期間の一覧を作る

前提:
  extract_plan_period.py と同じ抽出ロジックを使う。
  ファイル名の先頭を自治体名にしておくと、そのまま突合できる。
    例) 函館市_環境基本計画第3次.pdf → 自治体名「函館市」

使い方:
  python3 batch_extract_plan_periods.py <PDFフォルダ> [--pages 8] [--out 結果.csv]

出力:
  自治体名 / 計画期間 / 終期 / R9満了判定 / スコア / 根拠行 / ファイル名 の表を
  CSV（既定 cp932・CRLF）で出す。そのまま 02_営業先マスタ へ転記できる。

この環境では自治体サイトへの直接アクセスが組織のegressポリシーで遮断されて
いるため、PDFの取得は制限のない環境で行い、本スクリプトはそのフォルダに対して
実行する想定。
"""

import argparse
import os
import re
import sys
import traceback

import extract_plan_period as X
import naming

RESULT_HEADERS = ["自治体名", "計画期間", "終期年度", "R9満了判定", "スコア", "確認状況", "根拠行", "ファイル名"]


def guess_name(filename):
    """ファイル名の先頭から自治体名を推定する。"""
    stem = os.path.splitext(os.path.basename(filename))[0]
    m = re.match(r"([^\s_\-　]+?[都道府県市区町村])", stem)
    return m.group(1) if m else stem


def analyze(pdf, pages):
    chars, ok = X.has_text_layer(pdf)
    if not ok:
        return {"status": "画像PDF（OCR必要）", "chars": chars}

    docx_path = os.path.splitext(pdf)[0] + ".docx"
    docx_text = ""
    try:
        X.pdf_to_docx(pdf, docx_path, pages)
        docx_text = X.text_from_docx(docx_path)
    except Exception:
        pass
    plain = X.text_from_pdftotext(pdf, pages)

    best, best_score = None, -1
    for text in (docx_text, plain):
        for c in X.find_periods(text):
            s = X.score(c)
            if s > best_score:
                best, best_score = c, s
    if not best or best_score <= 0:
        return {"status": "期間を抽出できず", "chars": chars}

    y1, y2 = best[0], best[1]
    return {
        "status": "確認済" if best_score >= 20 else "要目視確認（スコア低）",
        "y1": y1, "y2": y2, "score": best_score, "line": best[3][:120], "chars": chars,
    }


def main():
    ap = argparse.ArgumentParser(description="計画PDFを一括処理して計画期間の一覧を作る")
    ap.add_argument("folder", help="PDFの入ったフォルダ")
    ap.add_argument("--pages", type=int, default=8, help="先頭から何ページを対象にするか（既定8）")
    ap.add_argument("--out", default=None, help="出力CSV（既定: フォルダ内 計画期間抽出結果.csv）")
    args = ap.parse_args()

    if not os.path.isdir(args.folder):
        print(f"エラー: フォルダがありません -> {args.folder}")
        sys.exit(1)
    pdfs = sorted(f for f in os.listdir(args.folder) if f.lower().endswith(".pdf"))
    if not pdfs:
        print(f"エラー: PDFが1件もありません -> {args.folder}")
        sys.exit(1)

    rows = [RESULT_HEADERS]
    hit_r9 = []
    print(f"{len(pdfs)}件を処理します（先頭{args.pages}ページ）\n")
    for f in pdfs:
        path = os.path.join(args.folder, f)
        name = guess_name(f)
        try:
            r = analyze(path, args.pages)
        except Exception:
            print(f"  {name:10} 例外: {traceback.format_exc(limit=1).strip().splitlines()[-1][:80]}")
            rows.append([name, "", "", "処理エラー", "", "処理エラー", "", f])
            continue

        if "y1" not in r:
            print(f"  {name:10} {r['status']}（テキスト{r['chars']:,}文字）")
            rows.append([name, "", "", "要確認", "", r["status"], "", f])
            continue

        y1, y2 = r["y1"], r["y2"]
        period = f"{X.seireki_to_wareki(y1)}〜{X.seireki_to_wareki(y2)}"
        hantei = "★該当" if y2 == 2027 else ("◇今年度満了" if y2 == 2026 else "非該当")
        if y2 == 2027:
            hit_r9.append(name)
        print(f"  {name:10} {period}（{y2-y1+1}年間） {hantei} score={r['score']}")
        rows.append([name, period, X.seireki_to_wareki(y2), hantei, r["score"], r["status"], r["line"], f])

    out = args.out or os.path.join(args.folder, "計画期間抽出結果.csv")
    naming.write_csv(out, rows)
    print(f"\n出力: {out}（{naming.csv_encoding()} / CRLF）")
    print(f"R9(2027)満了の該当: {len(hit_r9)}件" + (f" -> {'、'.join(hit_r9)}" if hit_r9 else ""))
    print("結果を 02_営業先マスタ の G〜K列へ転記してください。")


if __name__ == "__main__":
    main()
