# -*- coding: utf-8 -*-
"""資料編 付表1 を v8 の凡例（分母＝回答者数142・無回答非該当を含む）で検算する。"""
import re
import sys
from decimal import Decimal, ROUND_HALF_UP
from load_jittai import load
from parse_tables import parse

SUBMAP = {"訪問介護": "訪問介護", "訪問入浴介護": "訪問入浴", "訪問看護": "訪問看護",
          "訪問リハビリテーション": "訪問リハ", "通所介護": "通所介護",
          "通所リハビリテーション": "通所リハ", "夜間対応型訪問介護": "夜間対応型訪問",
          "定期巡回・随時対応型訪問介護看護": "定期巡回",
          "小規模多機能型居宅介護": "小多機",
          "看護小規模多機能型居宅介護": "看多機",
          "短期入所生活介護": "ショートステイ", "居宅療養管理指導": "居宅療養管理"}
N_ALL = 142


def keyof(cap):
    m = re.search(r"([AB]問\d)", cap)
    if not m:
        return None
    q = m.group(1)
    if q != "A問8":
        return q
    for k in sorted(SUBMAP, key=len, reverse=True):
        if k in cap:
            return "A問8_" + SUBMAP[k]
    return None


def pct(c, n):
    return float(Decimal(str(c * 100 / n)).quantize(Decimal("0.1"),
                                                    rounding=ROUND_HALF_UP))


blocks, recs = load()
BY = {b["key"]: b for b in blocks}
tables, paras = parse(sys.argv[1] if len(sys.argv) > 1 else "N_B.txt")

ok = 0
ng = []
ncell = 0
ntbl = 0
for num, cap, rows in tables:
    if "付表1-" not in cap:
        continue
    ntbl += 1
    key = keyof(cap)
    if key is None or key not in BY:
        ng.append((cap, f"設問不明 key={key}"))
        continue
    blk = BY[key]
    real = [o for c, o in blk["opts"]]
    ans = [r["ans"][key] for r in recs]
    valid = [s for s in ans if s]
    n_valid = len(valid)
    seen = []
    for row in rows[1:]:
        lab = row[0].strip()
        got_s = re.sub(r"[^\d]", "", row[1])
        got = int(got_s) if got_s else None
        if lab.startswith("回答者数"):
            ncell += 1
            if got != N_ALL:
                ng.append((cap, f"回答者数(n) {got} ≠ {N_ALL}"))
            else:
                ok += 1
            if row[2].strip() not in ("－", "-", "−", ""):
                ng.append((cap, f"回答者数(n) の割合欄が空でない: {row[2]}"))
            continue
        if lab.startswith("無回答"):
            exp = N_ALL - n_valid
            ncell += 1
            if got != exp:
                ng.append((cap, f"無回答・非該当 {got} ≠ {exp}（142−有効{n_valid}）"))
            else:
                ok += 1
            cnt = exp
        else:
            cands = [o for o in real if o.replace(" ", "") == lab.replace(" ", "")]
            if not cands:
                ng.append((cap, f"選択肢が個票にない: {lab}"))
                continue
            o = cands[0]
            seen.append(o)
            cnt = sum(1 for s in valid if o in s)
            ncell += 1
            if got != cnt:
                ng.append((cap, f"{lab} 件数{got} ≠ 再集計{cnt}"))
            else:
                ok += 1
        p = row[2].replace("%", "").replace("％", "").strip()
        if p not in ("－", "-", "−", ""):
            ncell += 1
            e = pct(cnt, N_ALL)
            if abs(float(p) - e) > 0.051:
                ng.append((cap, f"{lab} 割合{p}% ≠ 再計算{e}%（{cnt}/{N_ALL}）"))
            else:
                ok += 1
    miss = [o for o in real if o not in seen]
    if miss:
        ng.append((cap, f"表に載っていない選択肢: {miss}"))

print(f"付表1 {ntbl}表／照合 {ncell}項目／一致 {ok}／不一致 {len(ng)}")
for cap, m in ng:
    print("  ×", cap[:46], "|", m)
