# -*- coding: utf-8 -*-
"""
経営戦略の計画期間 自動抽出スクリプト（外部ネットワークが使える環境で実行する用）

本調査環境は組織のイーグレスポリシーにより自治体サイトへの直接アクセスが
遮断されているため、確度「低」の残件は計画期間を確定できていない。
通常のインターネット接続がある環境で本スクリプトを実行すると、
research/findings.jsonl の出典URL（HTML/PDF）を順に取得し、
「計画期間」「策定」の記述を正規表現で抽出して findings.jsonl を更新する。

使い方:
    pip install requests pypdf      # PDFを読む場合のみ pypdf が必要
    python3 tools/fetch_periods.py            # 確度「低」の全件を対象
    python3 tools/fetch_periods.py --muni 江差町   # 特定自治体のみ
    python3 tools/fetch_periods.py --dry-run       # 抽出結果の確認のみ（保存しない）
"""

import argparse
import io
import json
import os
import re
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from senryaku_db import load, add, FINDINGS, to_fy  # noqa: E402

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")

# 「平成29年度から平成38年度」「令和7年度～令和16年度」「令和7年度－令和16年度」等
ERA = r"(平成|令和|Ｈ|Ｒ)\s*([0-9０-９]{1,2}|元)"
PERIOD_PATTERNS = [
    re.compile(ERA + r"\s*年度\s*(?:から|～|〜|-|－|ー|~)\s*" + ERA + r"\s*年度"),
    re.compile(r"計\s*画\s*期\s*間[^0-9０-９]{0,20}" + ERA + r"\s*年度[^0-9０-９]{0,6}" + ERA + r"\s*年度"),
    # 「2025（令和7）年度～2034（令和16）年度」のような括弧付き併記にも対応
    re.compile(r"(20[0-9]{2})\s*(?:（[^）]*）|\([^)]*\))?\s*年度\s*(?:から|～|〜|-|－|ー|~)\s*"
               r"(20[0-9]{2})\s*(?:（[^）]*）|\([^)]*\))?\s*年度"),
]
MADE_PATTERN = re.compile(r"(策\s*定|改\s*定|改\s*訂)\s*(?:日|年月)?\s*[：:]?\s*" + ERA + r"\s*年\s*([0-9０-９]{1,2})\s*月")

Z2H = str.maketrans("０１２３４５６７８９", "0123456789")


def norm(s):
    return s.translate(Z2H).replace("Ｈ", "平成").replace("Ｒ", "令和")


def era_to_fy(era, num):
    num = 1 if num == "元" else int(num)
    return (2018 if era == "令和" else 1988) + num


def extract(text):
    """本文から (開始年度, 終了年度, 策定/改定年月文字列) を推定する。"""
    t = norm(re.sub(r"\s+", " ", text))
    start = end = None
    for pat in PERIOD_PATTERNS:
        m = pat.search(t)
        if not m:
            continue
        g = m.groups()
        if len(g) == 4:
            start, end = era_to_fy(g[0], g[1]), era_to_fy(g[2], g[3])
        else:
            start, end = int(g[0]), int(g[1])
        break
    made = ""
    m = MADE_PATTERN.search(t)
    if m:
        made = f"{m.group(2)}{m.group(3)}年{m.group(4)}月{m.group(1)}"
    return start, end, made


def fetch(url, timeout=45):
    import requests
    r = requests.get(url, headers={"User-Agent": UA}, timeout=timeout)
    r.raise_for_status()
    ctype = r.headers.get("Content-Type", "")
    if "pdf" in ctype.lower() or url.lower().endswith(".pdf"):
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(r.content))
        pages = reader.pages[:4]  # 表紙〜数ページに計画期間が書かれている
        return "\n".join((p.extract_text() or "") for p in pages)
    r.encoding = r.apparent_encoding or r.encoding
    return re.sub(r"<[^>]+>", " ", r.text)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--muni", help="対象自治体名（部分一致）")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--sleep", type=float, default=1.5, help="リクエスト間隔（秒）")
    args = ap.parse_args()

    recs = load()
    targets = [r for r in recs
               if r.get("confidence") == "低" and r.get("source", "").startswith("http")
               and (not args.muni or args.muni in r["muni"])]
    print(f"対象 {len(targets)} 件")

    updates, failed = [], []
    for i, r in enumerate(targets, 1):
        label = f"{r['muni']}/{r['jigyo']}"
        try:
            text = fetch(r["source"])
            start, end, made = extract(text)
        except Exception as e:  # noqa: BLE001
            failed.append((label, str(e)[:80]))
            print(f"[{i}/{len(targets)}] NG   {label}: {e}")
            time.sleep(args.sleep)
            continue
        if start or end:
            rec = dict(r)
            rec["start_fy"], rec["end_fy"] = start, end
            if made:
                rec["note"] = (rec.get("note", "") + f"／自動抽出：{made}").strip("／")
            rec["confidence"] = "中"  # 自動抽出のため要目視確認
            updates.append(rec)
            print(f"[{i}/{len(targets)}] OK   {label}: {start}〜{end} {made}")
        else:
            failed.append((label, "計画期間の記述を検出できず"))
            print(f"[{i}/{len(targets)}] ---  {label}: 抽出できず")
        time.sleep(args.sleep)

    print(f"\n抽出成功 {len(updates)} 件 / 失敗・未検出 {len(failed)} 件")
    if updates and not args.dry_run:
        add(updates)
        print(f"{FINDINGS} を更新しました。python3 tools/build_report.py で帳票を再生成してください。")
    if failed:
        print("\n--- 手動確認が必要 ---")
        for label, reason in failed:
            print(f"  {label}: {reason}")


if __name__ == "__main__":
    main()
