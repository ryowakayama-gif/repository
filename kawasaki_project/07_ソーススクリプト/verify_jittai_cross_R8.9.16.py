# -*- coding: utf-8 -*-
"""本編 第9章・第10章のクロス表を個票から再集計して照合する。"""
import re
from decimal import Decimal, ROUND_HALF_UP
from load_jittai import load, layers, AGE3
from parse_tables import parse

blocks, recs = load()
BY = {b["key"]: b for b in blocks}
A = lambda r, k: r["ans"][k]


def pct(c, n):
    return float(Decimal(str(c * 100 / n)).quantize(Decimal("0.1"),
                                                    rounding=ROUND_HALF_UP))


def one(r, k):
    s = A(r, k)
    return s[0] if s else None


# 行の定義：表番号 -> (行の絞り込み関数, 列の設問キー, 列ラベル→選択肢)
ROWDEF = {}
ROWDEF["表9-1"] = (lambda v: [(x, lambda r, x=x: one(r, "A問2") == x)
                              for x in ["単身世帯", "夫婦のみ世帯", "その他"]], "A問9")
ROWDEF["表9-2"] = (lambda v: [(x, lambda r, x=x: one(r, "A問5") == x)
                              for x in ["要支援１", "要支援２", "要介護１", "要介護２",
                                        "要介護３", "要介護４", "要介護５", "わからない"]], "A問6")
ROWDEF["表9-3"] = (ROWDEF["表9-2"][0], "A問7")
ROWDEF["表9-4"] = (lambda v: [(x, lambda r, x=x: one(r, "A問9") == x)
                              for x in ["家族・親族の介護はあるが、週に１日よりも少ない",
                                        "週に１～２日ある", "週に３～４日ある", "ほぼ毎日ある"]], "B問3")
ROWDEF["表9-5"] = (lambda v: [(x, lambda r, x=x: one(r, "B問2") == x)
                              for x in ["20歳未満", "20代", "30代", "40代", "50代",
                                        "60代", "70代", "80歳以上", "わからない"]], "B問4")
ROWDEF["表9-6"] = (lambda v: [(x, lambda r, x=x: one(r, "B問4") == x)
                              for x in ["フルタイムで働いている", "パートタイムで働いている",
                                        "働いていない", "主な介護者に確認しないと、わからない"]], "B問5")
B5 = {"特に行っていない": "特に行っていない",
      "労働時間調整": "介護のために、「労働時間を調整（残業免除、短時間勤務、遅出・早帰・中抜け等）」しながら、働いている",
      "休暇取得": "介護のために、「休暇（年休や介護休暇等）」を取りながら、働いている",
      "在宅勤務利用": "介護のために、「在宅勤務」を利用しながら、働いている",
      "２～４以外の調整": "介護のために、２～４以外の調整をしながら、働いている",
      "わからない": "主な介護者に確認しないと、わからない"}
ROWDEF["表9-7"] = (lambda v: [(k, lambda r, o=o: o in A(r, "B問5"))
                              for k, o in B5.items()], "B問6")
ROWDEF["表9-8"] = (lambda v: [
    ("離職者あり（主な介護者が離職）",
     lambda r: "主な介護者が仕事を辞めた（転職除く）" in A(r, "B問1")),
    ("離職なし",
     lambda r: "介護のために仕事を辞めた家族・親族はいない" in A(r, "B問1"))], "B問3")
ROWDEF["表9-9"] = (lambda v: [
    ("介護は「ない」", lambda r: one(r, "A問9") == "ない"),
    ("介護が「ある」", lambda r: one(r, "A問9") not in (None, "ない"))], "A問6")
ROWDEF["表9-10"] = (ROWDEF["表9-1"][0], "A問6")
ROWDEF["表9-11"] = (lambda v: [(x, lambda r, x=x: one(r, "A問7") == x)
                               for x in ["利用した", "利用していない"]], "A問6")
LAY = layers(recs)
ROWDEF["表10-2"] = (lambda v: [(s, lambda r, s=s: r["性別"] == s)
                               for s in ["男性", "女性"]], "A問2")
ROWDEF["表10-3"] = (ROWDEF["表10-2"][0], "A問9")
ROWDEF["表10-4"] = (ROWDEF["表10-2"][0], "A問6")
ROWDEF["表10-5"] = (ROWDEF["表10-2"][0], "B問2")
ROWDEF["表10-6"] = (ROWDEF["表10-2"][0], "B問4")
AGES = ["65歳未満", "65～69歳", "70～74歳", "75～79歳", "80～84歳", "85～89歳", "90歳以上"]
ROWDEF["表10-7"] = (lambda v: [(a, lambda r, a=a: r["年齢階級"] == a)
                               for a in AGES], "A問6")
ROWDEF["表10-8"] = (ROWDEF["表10-7"][0], "A問7")
ROWDEF["表10-9"] = (ROWDEF["表10-7"][0], "A問9")
ROWDEF["表10-11"] = (lambda v: [(x, lambda r, x=x: one(r, "A問5") == x)
                                for x in ["要介護１", "要介護２", "要介護３",
                                          "要介護４", "要介護５"]], "B問3")
SX = [(f"{s}／{a}", (s, a)) for s in ["男性", "女性"]
      for a in ["74歳以下", "75～84歳", "85歳以上"]]
ROWDEF["表10-12"] = (lambda v: [(k, lambda r, t=t: r["性別"] == t[0]
                                 and AGE3.get(r["年齢階級"]) == t[1])
                                for k, t in SX], "A問9")
ROWDEF["表10-13"] = (ROWDEF["表10-12"][0], "A問6")

NUM = re.compile(r"^(\d+)件?\s*(?:[（(](\d+\.\d)%[）)])?$")
tables, paras = parse("N_jittai.txt")
ok = 0
ng = []
cells = 0
checked = []
for num, cap, rows in tables:
    m = re.match(r"表(9|10)-(\d+)", cap)
    if not m:
        continue
    tname = f"表{m.group(1)}-{m.group(2)}"
    if tname not in ROWDEF:
        continue
    checked.append(tname)
    rowfn, qkey = ROWDEF[tname]
    real = [o for c, o in BY[qkey]["opts"]]
    head = rows[0]
    cols = head[2:]
    idx = []
    for c in cols:
        c2 = c.replace("⏎", "").strip()
        cand = [o for o in real if o.replace(" ", "") == c2.replace(" ", "")]
        idx.append(cand[0] if cand else None)
    if any(i is None for i in idx):
        ng.append((tname, f"列見出しが個票と不一致: {[c for c,i in zip(cols,idx) if i is None]}"))
        continue
    defs = dict(rowfn(None))
    defs["全体"] = lambda r: True
    if tname == "表10-5":
        defs["男性"] = lambda r: r["性別"] == "男性"
        defs["女性"] = lambda r: r["性別"] == "女性"
    for row in rows[1:]:
        lab = row[0].replace("⏎", "").replace("\u3000", "").strip()
        if lab not in defs:
            ng.append((tname, f"行ラベル不明: {lab}"))
            continue
        pool = [r for r in recs if defs[lab](r)]
        valid = [r for r in pool if A(r, qkey)]
        n = len(valid)
        cells += 1
        try:
            nrep = int(re.sub(r"[^\d]", "", row[1]))
        except ValueError:
            ng.append((tname, f"{lab} n欄不正 {row[1]}"))
            continue
        if nrep != n:
            ng.append((tname, f"{lab} n={nrep} ≠ 再集計{n}"))
        else:
            ok += 1
        for j, o in enumerate(idx):
            cell = row[2 + j].replace("⏎", "").strip()
            cnt = sum(1 for r in valid if o in A(r, qkey))
            if cell in ("−", "－", "-"):
                continue
            mm = NUM.match(cell.replace(" ", "").replace("　", ""))
            if not mm:
                ng.append((tname, f"{lab}／{cols[j][:14]} 読めない {cell!r}"))
                continue
            cells += 1
            if int(mm.group(1)) != cnt:
                ng.append((tname, f"{lab}／{cols[j][:14]} 件数{mm.group(1)} ≠ {cnt}"))
            else:
                ok += 1
            if mm.group(2):
                cells += 1
                e = pct(cnt, n) if n else None
                if e is None or abs(float(mm.group(2)) - e) > 0.051:
                    ng.append((tname, f"{lab}／{cols[j][:14]} 割合{mm.group(2)}% ≠ {e}%（{cnt}/{n}）"))
                else:
                    ok += 1

print("照合した表:", " ".join(checked))
print(f"照合 {cells}項目／一致 {ok}／不一致 {len(ng)}")
for t, m in ng:
    print("  ×", t, "|", m)
