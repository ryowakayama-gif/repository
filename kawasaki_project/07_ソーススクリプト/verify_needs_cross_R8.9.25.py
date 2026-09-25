# -*- coding: utf-8 -*-
"""ニーズ調査 第3章のクロス表とχ²検定を個票から再計算して照合する。"""
import math
import re
from decimal import Decimal, ROUND_HALF_UP
from load_needs import load
from parse_tables import parse

blocks, recs = load()


def pct(c, n):
    return float(Decimal(str(c * 100 / n)).quantize(Decimal("0.1"),
                                                    rounding=ROUND_HALF_UP))


def sel(rec, i):
    return rec[i]


# 表番号 -> (行設問, 行ラベル→個票の選択肢, 列設問, 列ラベル→個票の選択肢)
D = {}
D[86] = (10, {"何度もある": ["何度もある"], "1度ある": ["1度ある"], "ない": ["ない"]},
         11, {"とても不安": ["とても不安である"], "やや不安": ["やや不安である"],
              "あまり不安でない": ["あまり不安でない"], "不安でない": ["不安でない"]})
D[87] = (11, {"とても不安": ["とても不安である"], "やや不安": ["やや不安である"],
              "あまり不安でない": ["あまり不安でない"], "不安でない": ["不安でない"]},
         14, {"はい": ["はい"], "いいえ": ["いいえ"]})
D[88] = (42, {"趣味あり": ["趣味あり"], "思いつかない": ["思いつかない"]},
         43, {"生きがいあり": ["生きがいあり"], "思いつかない": ["思いつかない"]})
D[89] = (26, {"はい": ["はい"], "いいえ": ["いいえ"]},
         71, {"知っている": ["はい"], "知らない": ["いいえ"]})
D[91] = (5, {k: [k] for k in ["大変苦しい", "やや苦しい", "ふつう",
                              "ややゆとりがある", "大変ゆとりがある"]},
         64, {k: [k] for k in ["とてもよい", "まあよい", "あまりよくない", "よくない"]})
D[92] = (65, {"あった": ["はい"], "なかった": ["いいえ"]},
         42, {"趣味あり": ["趣味あり"], "思いつかない": ["思いつかない"]})
D[93] = (65, {"あった": ["はい"], "なかった": ["いいえ"]},
         43, {"生きがいあり": ["生きがいあり"], "思いつかない": ["思いつかない"]})
D[94] = (13, {k: [k] for k in ["とても減っている", "減っている",
                               "あまり減っていない", "減っていない"]},
         10, {"あり(何度も+1度)": ["何度もある", "1度ある"], "ない": ["ない"]})
D[95] = (52, {k: [k] for k in ["是非参加したい", "参加してもよい",
                               "参加したくない", "既に参加している"]}, None, None)

ACT = list(range(44, 52))   # 問5の8活動


def grid(t):
    ri, rmap, ci, cmap = D[t]
    rows = []
    for rl, rv in rmap.items():
        pool = [r for r in recs if any(v in r[ri] for v in rv)]
        if t == 95:
            a = sum(1 for r in pool
                    if any(r[k] and r[k][0] != "参加していない" for k in ACT))
            b = sum(1 for r in pool
                    if any(r[k] for k in ACT)
                    and all(r[k] == ["参加していない"] for k in ACT if r[k]))
            rows.append((rl, [("いずれか参加している", a), ("全て非参加", b)]))
        else:
            pool = [r for r in pool if r[ci]]
            rows.append((rl, [(cl, sum(1 for r in pool
                                       if any(v in r[ci] for v in cv)))
                              for cl, cv in cmap.items()]))
    return rows


def chi2(mat, yates=False):
    R = len(mat)
    C = len(mat[0])
    N = sum(sum(r) for r in mat)
    rs = [sum(r) for r in mat]
    cs = [sum(mat[i][j] for i in range(R)) for j in range(C)]
    x = 0.0
    for i in range(R):
        for j in range(C):
            e = rs[i] * cs[j] / N
            d = abs(mat[i][j] - e)
            if yates:
                d = max(0.0, d - 0.5)
            x += d * d / e
    v = math.sqrt(x / (N * min(R - 1, C - 1)))
    return x, (R - 1) * (C - 1), v, N


RE_CELL = re.compile(r"(\d+\.\d)[%％][（(](\d+)件[）)]")
tables, paras = parse("N_needs.txt")
T = {n: (c, r) for n, c, r in tables}
bad = []
nchk = 0
print("表  n照合  セル照合  χ²再計算")
for t in sorted(D):
    cap, rows = T[t]
    g = dict(grid(t))
    mat = []
    for row in rows[1:]:
        lab = row[0].replace("⏎", "")
        m = re.match(r"(.+?)\s*[（(]n=(\d+)[）)]", lab)
        if not m:
            bad.append((t, f"行ラベルが読めない {lab!r}"))
            continue
        key, nrep = m.group(1).strip(), int(m.group(2))
        if key not in g:
            bad.append((t, f"行が対応しない {key!r} (候補 {list(g)})"))
            continue
        cells = g[key]
        ntrue = sum(c for _, c in cells)
        nchk += 1
        if nrep != ntrue:
            bad.append((t, f"{key} n={nrep} ≠ 再集計{ntrue}"))
        mat.append([c for _, c in cells])
        for j, (cl, cnt) in enumerate(cells):
            cell = row[1 + j].replace("⏎", "").replace(" ", "")
            if cell in ("−", "－", "-"):
                continue
            mm = RE_CELL.match(cell)
            if mm:
                nchk += 2
                if int(mm.group(2)) != cnt:
                    bad.append((t, f"{key}／{cl} 件数{mm.group(2)} ≠ {cnt}"))
                e = pct(cnt, ntrue)
                if abs(float(mm.group(1)) - e) > 0.051:
                    bad.append((t, f"{key}／{cl} 割合{mm.group(1)}% ≠ {e}%"))
            elif re.fullmatch(r"\d+件", cell):
                nchk += 1
                if int(cell[:-1]) != cnt:
                    bad.append((t, f"{key}／{cl} 件数{cell} ≠ {cnt}"))
            else:
                bad.append((t, f"{key}／{cl} 読めない {cell!r}"))
    if mat and all(len(r) == len(mat[0]) for r in mat):
        y = (len(mat) == 2 and len(mat[0]) == 2)
        x, df, v, N = chi2(mat, y)
        print(f"  表{t}: χ²={x:.3f} df={df} V={v:.3f} n={N}"
              + ("（イェーツ補正）" if y else ""))
print(f"\n照合 {nchk}項目／不一致 {len(bad)}件")
for t, m in bad:
    print("  ×", f"表{t}", m)
