# -*- coding: utf-8 -*-
"""ニーズ調査結果報告書（本編）の単純集計表を個票から全数再集計して照合する。"""
import re
import sys
from decimal import Decimal, ROUND_HALF_UP
from load_needs import load
from parse_tables import parse


def norm(s):
    s = str(s or "")
    for a, b in (("（", "("), ("）", ")"), ("〜", "～"), ("　", ""), (" ", ""),
                 ("・", ""), ("，", ""), ("、", ""), ("。", ""), ("．", ""),
                 ("１", "1"), ("２", "2"), ("３", "3"), ("４", "4"), ("５", "5"),
                 ("６", "6"), ("７", "7"), ("８", "8"), ("９", "9"), ("０", "0")):
        s = s.replace(a, b)
    return s.strip()


def pct(c, n):
    return float(Decimal(str(c * 100 / n)).quantize(Decimal("0.1"),
                                                    rounding=ROUND_HALF_UP))


blocks, recs = load()
tables, paras = parse(sys.argv[1] if len(sys.argv) > 1 else "N_A.txt")

# 本文の表キャプション（表n　設問文（n=xxx））
CAP = re.compile(r"^表(\d+)\s*[　 ]\s*(.+?)(?:（n=(\d+)）)?$")

used = set()
ok = 0
ng = []
ncell = 0
done = []
for num, cap, rows in tables:
    m = CAP.match(cap.strip())
    if not m:
        continue
    body = norm(m.group(2))
    if not m.group(3):
        continue
    n_rep = int(m.group(3))
    head = [norm(h) for h in rows[0]]
    if "件数" not in head or "選択肢" not in "".join(head):
        continue
    # 設問ブロックを本文から特定
    cands = [i for i, b in enumerate(blocks)
             if norm(b["text"]) and (norm(b["text"]) in body
                                     or body in norm(b["text"]))]
    cands = [i for i in cands if i not in used] or cands
    if not cands:
        continue
    # 選択肢の一致度で最良のブロックを選ぶ
    labs = [norm(r[0]) for r in rows[1:] if norm(r[0])]
    best, score = None, -1
    for i in cands:
        opts = [norm(o) for o in blocks[i]["opts"]]
        s = sum(1 for l in labs if l in opts)
        if s > score:
            best, score = i, s
    if score <= 0:
        continue
    used.add(best)
    blk = blocks[best]
    opts = blk["opts"]
    ans = [r[best] for r in recs]
    valid = [s for s in ans if s]
    ci = head.index("件数")
    pi = next((j for j, h in enumerate(head) if "割合" in h), None)
    ncell += 1
    if n_rep != len(valid):
        ng.append((f"表{m.group(1)}", f"n={n_rep} ≠ 再集計{len(valid)}"))
    else:
        ok += 1
    for row in rows[1:]:
        lab = norm(row[0])
        g0 = re.sub(r"[^\d]", "", row[ci]) if ci < len(row) else ""
        if not lab or not g0 or "合計" in lab or lab.startswith("n="):
            continue
        hit = [o for o in opts if norm(o) == lab]
        if not hit:
            ng.append((f"表{m.group(1)}", f"選択肢が個票にない: {row[0][:28]}"))
            continue
        cnt = sum(1 for s in valid if hit[0] in s)
        g = re.sub(r"[^\d]", "", row[ci]) if ci < len(row) else ""
        if g:
            ncell += 1
            if int(g) != cnt:
                ng.append((f"表{m.group(1)}",
                           f"{row[0][:24]} 件数{g} ≠ 再集計{cnt}"))
            else:
                ok += 1
        if pi is not None and pi < len(row):
            p = row[pi].replace("%", "").replace("％", "").strip()
            if re.match(r"^\d+(\.\d+)?$", p):
                ncell += 1
                e = pct(cnt, len(valid))
                if abs(float(p) - e) > 0.051:
                    ng.append((f"表{m.group(1)}",
                               f"{row[0][:24]} 割合{p}% ≠ 再計算{e}%"
                               f"（{cnt}/{len(valid)}）"))
                else:
                    ok += 1
    done.append(m.group(1))

print(f"ニーズ調査 単純集計 {len(done)}表／照合 {ncell}項目／一致 {ok}／不一致 {len(ng)}")
print("照合した表:", " ".join("表" + d for d in done))
for t, msg in ng:
    print("  ×", t, "|", msg)
