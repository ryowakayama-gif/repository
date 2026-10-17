# -*- coding: utf-8 -*-
"""グラフの系列値を、同じ設問の表の割合と突き合わせる。"""
import re
import sys
import unicodedata
import importlib.util as _ilu
import os as _os
_spec = _ilu.spec_from_file_location(
    "dump_chart_data",
    _os.path.join(_os.path.dirname(_os.path.abspath(__file__)),
                  "dump_chart_data_R8.9.16.py"))
_m = _ilu.module_from_spec(_spec)
_spec.loader.exec_module(_m)
charts = _m.charts
from parse_tables import parse


def norm(s):
    s = unicodedata.normalize("NFKC", s)
    s = re.sub(r"[\s　]", "", s)
    return re.sub(r"[（）()「」・,、。／/？?＋+~〜～\-－—]", "", s)


docx, dump = sys.argv[1], sys.argv[2]
cs = charts(docx)
tables, paras = parse(dump)
ok = ng = unmatched = 0
msgs = []
for k in sorted(cs):
    title, cats, vals = cs[k]
    tt = norm(re.sub(r"^図[\d\-－]+", "", title))
    best, score = None, 0.0
    for num, cap, rows in tables:
        ct = norm(re.sub(r"^表[\d\-－]+", "", cap))
        if not ct:
            continue
        a, b = set(zip(tt, tt[1:])), set(zip(ct, ct[1:]))
        if not a or not b:
            continue
        s = len(a & b) / min(len(a), len(b))
        if s > score:
            best, score = (num, cap, rows), s
    if score < 0.6:
        unmatched += 1
        msgs.append(f"  ? chart{k}「{title[:40]}」 対応表なし(最良{score:.2f})")
        continue
    num, cap, rows = best
    # 表から割合を順に拾う
    tp = []
    for row in rows[1:]:
        for c in row[1:]:
            m = re.search(r"(\d+\.\d)\s*[%％]", c.replace("⏎", ""))
            if m:
                tp.append(float(m.group(1)))
                break
    flat = [round(float(x) * 100, 1) for v in vals for x in v]
    if len(flat) == len(tp):
        for i, (a, b) in enumerate(zip(flat, tp)):
            if abs(a - b) > 0.051:
                ng += 1
                msgs.append(f"  × chart{k}/{cap[:24]} 第{i+1}点 グラフ{a} ≠ 表{b}")
            else:
                ok += 1
    else:
        unmatched += 1
        msgs.append(f"  ? chart{k}「{title[:30]}」 点数{len(flat)} vs 表{len(tp)}（{cap[:26]}）")
print(f"{docx.split('/')[-1]}: 照合{ok+ng}点／一致{ok}／不一致{ng}／突合できず{unmatched}")
for m in msgs:
    print(m)
