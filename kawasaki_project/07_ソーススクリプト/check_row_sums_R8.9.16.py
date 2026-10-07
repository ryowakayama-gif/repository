# -*- coding: utf-8 -*-
"""表の行合計・列合計・割合合計を検査する。"""
import re
import sys
from parse_tables import parse

NUMC = re.compile(r"^(\d+)\s*(?:件|人|名)?$")
PCT = re.compile(r"^(\d+\.\d)\s*[%％]$")
CELL = re.compile(r"^(\d+)\s*件?\s*[（(](\d+\.\d)[%％][）)]$")


def val(s):
    s = s.replace("⏎", "").replace(" ", "").replace("　", "")
    m = CELL.match(s)
    if m:
        return int(m.group(1)), float(m.group(2))
    m = NUMC.match(s)
    if m:
        return int(m.group(1)), None
    m = PCT.match(s)
    if m:
        return None, float(m.group(1))
    return None, None


for path in sys.argv[1:]:
    print(f"########## {path}")
    tables, paras = parse(path)
    for num, cap, rows in tables:
        if len(rows) < 2:
            continue
        head = rows[0]
        # ①「選択肢｜件数｜割合」型（末尾に有効回答数）
        ncols = len(head)
        # 行ごとの割合合計
        for i, row in enumerate(rows[1:], 1):
            ps = []
            cs = []
            for c in row[1:]:
                k, p = val(c)
                if p is not None:
                    ps.append(p)
                if k is not None:
                    cs.append(k)
            if len(ps) >= 3 and abs(sum(ps) - 100.0) > 0.35 and sum(ps) < 130:
                print(f"  表{num} {cap[:26]} 行{i}「{row[0][:20]}」"
                      f"割合計 {sum(ps):.1f}%")
