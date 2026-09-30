# -*- coding: utf-8 -*-
"""見える化システムから一括で受領したブックを読み、本村の値を1本の表にまとめる.

受領した実物は容量が大きくリポジトリに収めないため、本スクリプトで
**本村の値だけ**を取り出して data/mieruka_batch.csv に落とす。

  python3 scripts/parse_mieruka_batch.py <受領したフォルダ>

**他自治体の値は取り出さない。** 受領したブックは県内の全市町村を含むが、
本村以外は参照値であり、成果品にも中間の表にも持ち込まない。
福島県・全国の行は公表値であり、比較の相手として取り出す。

様式が揃っていないため、次のように探索して読む。
  ・「表形式」で始まるシートを使う（グラフのシートは使わない）
  ・本村の行は、左の数列に「北塩原」を含むセルがある行
  ・期の見出しは、その行より上で、値の列に文字の見出しが並ぶ行
"""
import csv
import glob
import io
import os
import re
import sys

sys.dont_write_bytecode = True

from openpyxl import load_workbook

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "data", "mieruka_batch.csv")
# 取り出す地域。本村と、比較の相手として公表されている県・全国のみ
KEEP = ("北塩原", "福島県", "全国")


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


RE_REGION = re.compile(r"^(北塩原村?|福島県|全国)$")


def as_region(v):
    """セルがそれ自体で地域を表すときにその名を返す。

    「高齢化率（福島県）」のように地域名を含むだけの見出しを地域と
    取り違えないため、完全一致で見る。
    """
    if not isinstance(v, str):
        return None
    t = v.strip()
    return t if RE_REGION.match(t) else None


def region_in(v):
    """行の見出しが（福島県）のように地域を名指ししていればその名を返す"""
    if not isinstance(v, str):
        return None
    for k in ("福島県", "全国"):
        if k in v:
            return k
    return None


def header_row(rows):
    """期の見出しの行と、値が始まる列を返す。

    値の側から探すと、先頭の期が「-」（データなし）の地域に当たったときに
    列を見つけられない。見出しの行から決める。
    """
    best = None
    for i, r in enumerate(rows[:12]):
        # 文字の見出しが3つ以上並び、数がほとんどない行
        cells = [(j, v) for j, v in enumerate(r) if v not in (None, "")]
        lab = [(j, v) for j, v in cells if not is_num(v)]
        if len(lab) < 3:
            continue
        if sum(1 for _, v in cells if is_num(v)) > len(cells) // 2:
            continue
        # 見出しが右に連なり始める列
        js = [j for j, _ in lab]
        start = None
        for k in range(len(js) - 2):
            if js[k + 1] == js[k] + 1 and js[k + 2] == js[k] + 2:
                start = js[k]
                break
        if start is None or start < 2:
            continue
        if best is None or len(lab) > best[2]:
            best = (i, start, len(lab))
    return (best[0], best[1]) if best else (None, None)


def parse(path):
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        return [], f"開けない（{e}）"
    sheets = [s for s in wb.sheetnames if s.startswith("表形式")] or \
             [s for s in wb.sheetnames if s.startswith("表")]
    if not sheets:
        wb.close()
        return [], "表形式のシートがない"
    out, note = [], ""
    # ブックの冒頭（どのシートでもよい）から「確認している地域」を拾う
    book_head = ""
    for sn0 in wb.sheetnames:
        try:
            for r in wb[sn0].iter_rows(min_row=1, max_row=4, values_only=True):
                book_head += " ".join(str(v) for v in r if isinstance(v, str)) + " "
        except Exception:
            pass
    for sn in sheets:
        ws = wb[sn]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        if not rows:
            continue
        h, vcol = header_row(rows)
        if vcol is None:
            note = "期の見出しの行が見つからない"
            continue
        periods = []
        for j in range(vcol, max(len(r) for r in rows)):
            v = rows[h][j] if j < len(rows[h]) else None
            periods.append(str(v).strip().replace("\n", "") if v not in (None, "")
                           else f"列{j}")
        # 様式3：見出しの行が地域名のとき（地域別）。列が地域、行が指標。
        # この様式では、地域名の列を採らずに本村の値として読むと、
        # 福島県だけが収録されたブックの値を本村の値として取り違える。
        hdr_regions = [(j, str(rows[h][j]).strip())
                       for j in range(vcol, len(rows[h]))
                       if isinstance(rows[h][j], str)
                       and re.search(r"[都道府県市区町村]$", str(rows[h][j]).strip())]
        if hdr_regions:
            keep_cols = [(j, nm) for j, nm in hdr_regions if as_region(nm)]
            if not keep_cols:
                note = "地域別だが本村・県・全国の列がない（収録：%s）" % \
                       "・".join(nm for _, nm in hdr_regions[:3])
                continue
            for i, r in enumerate(rows):
                if i <= h:
                    continue
                lab = " ".join(str(v).strip() for v in r[:vcol]
                               if isinstance(v, str) and str(v).strip()
                               and not str(v).strip().startswith("（"))
                unit = next((str(v).strip() for v in r[:vcol]
                             if isinstance(v, str) and str(v).strip().startswith("（")), "")
                if not lab:
                    continue
                for j, nm in keep_cols:
                    if j >= len(r) or not is_num(r[j]):
                        continue
                    out.append({"file": os.path.basename(path), "sheet": sn,
                                "region": nm, "indicator": lab, "unit": unit,
                                "period": "", "value": r[j]})
            continue

        tgt = []
        for i, r in enumerate(rows):
            if i <= h:
                continue
            for j, v in enumerate(r[:vcol]):
                nm = as_region(v)
                if nm:
                    tgt.append((i, j, nm))
                    break
        if not tgt:
            # 様式2：ブック全体が1つの地域のもの。
            # 1行目に「確認している地域：福島県北塩原村」とあり、
            # 行は地域ではなく計画値・実績値などの区分になっている
            # 地域名は「表形式」の冒頭にないことがある（グラフのシートの1行目にある）。
            # ブック全体の冒頭を見る
            head = " ".join(str(v) for r in rows[:10] for v in r if isinstance(v, str))
            head += " " + book_head
            who = next((k for k in KEEP if k in head), None)
            if who:
                for i, r in enumerate(rows):
                    if i <= h:
                        continue
                    lab = next((str(v).strip() for v in r[:vcol]
                                if isinstance(v, str) and str(v).strip()), "")
                    if lab:
                        # 「高齢化率（福島県）」のような行は、その地域の値である
                        tgt.append((i, 0, region_in(lab) or who))
            if not tgt:
                note = "本村・県・全国の行がない"
                continue
        for i, j, name in tgt:
            r = rows[i]
            ind, unit = "", ""
            parts = []
            for k in range(0 if j == 0 else j + 1, vcol):
                v = r[k]
                if isinstance(v, str) and v.strip():
                    if v.strip().startswith("（"):
                        unit = v.strip()
                    elif not (any(kk in v for kk in KEEP) and k == j):
                        parts.append(v.strip())
            ind = " ".join(parts)
            for off, j2 in enumerate(range(vcol, len(r))):
                v = r[j2]
                if not is_num(v):
                    continue
                out.append({
                    "file": os.path.basename(path), "sheet": sn, "region": name,
                    "indicator": ind, "unit": unit,
                    "period": periods[off] if off < len(periods) else f"列{j2}",
                    "value": v,
                })
    wb.close()
    return out, note


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else None
    if not src or not os.path.isdir(src):
        print("受領したフォルダを渡してください")
        return 1
    files = sorted(set(glob.glob(os.path.join(src, "**", "*.xlsx"), recursive=True)))
    rows, bad = [], []
    for p in files:
        got, note = parse(p)
        rows += got
        if note:
            bad.append((os.path.basename(p), note))
    with io.open(OUT, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, ["file", "sheet", "region", "indicator", "unit",
                               "period", "value"])
        w.writeheader()
        w.writerows(rows)
    n_mura = sum(1 for r in rows if "北塩原" in r["region"])
    print(f"ブック {len(files)}冊を読み、{len(rows)}行を取り出しました")
    print(f"  本村 {n_mura}行／県・全国 {len(rows) - n_mura}行")
    print(f"  取り出せなかったブック {len(bad)}冊")
    for b in bad[:10]:
        print(f"    {b[0][:64]}　{b[1]}")
    print("保存:", OUT)
    return 0


if __name__ == "__main__":
    sys.exit(main())
