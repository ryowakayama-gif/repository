# -*- coding: utf-8 -*-
"""本文中の「割合（n=分母・分子件）」「分子／分母件」形式の算術を検算する。"""
import re
import sys
from decimal import Decimal, ROUND_HALF_UP
from parse_tables import parse


def pct(c, n):
    return float(Decimal(str(c * 100 / n)).quantize(Decimal("0.1"),
                                                    rounding=ROUND_HALF_UP))


PATS = [
    # 55.6％（n=27・15件）
    (re.compile(r"(\d+\.\d)[%％][（(]n[=＝](\d+)[・,、](\d+)件[）)]"), "p,n,c"),
    # 55.6％（15／27件）
    (re.compile(r"(\d+\.\d)[%％][（(](\d+)[／/](\d+)件?[）)]"), "p,c,n"),
    # 15／27件 の 55.6％   → 別形式
    (re.compile(r"(\d+\.\d)[%％][・･](\d+)[／/](\d+)件"), "p,c,n"),
    # （n=27・15件）55.6％  はまれ
]

ng, ok = [], 0
for path in sys.argv[1:]:
    tables, paras = parse(path)
    texts = [(f"[P{p}]", t) for p, s, t in paras]
    for num, cap, rows in tables:
        for r in rows:
            texts.append((f"[表{num}]", " ".join(r)))
    for loc, t in texts:
        for rx, order in PATS:
            for m in rx.finditer(t):
                g = [x for x in m.groups()]
                if order == "p,n,c":
                    p, n, c = float(g[0]), int(g[1]), int(g[2])
                else:
                    p, c, n = float(g[0]), int(g[1]), int(g[2])
                if n == 0:
                    continue
                e = pct(c, n)
                if abs(p - e) > 0.051:
                    ng.append((path, loc, m.group(0), f"{c}/{n}={e}%"))
                else:
                    ok += 1
print(f"検算した本文中の割合 {ok + len(ng)}件／一致 {ok}／不一致 {len(ng)}")
for path, loc, s, e in ng:
    print(f"  × {path} {loc} 「{s}」 → 再計算 {e}")
