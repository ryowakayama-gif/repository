# -*- coding: utf-8 -*-
"""第2回策定委員会 資料の自己点検。

   資料は計画素案と交付金の算定から数値を書き写しているため、
   元を直したときの取り残しを検出する。不適合があれば終了コード1を返す。
"""
import csv, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import shiryo_content as SH
import soan_content as SO

R = []
def chk(no, name, ok, detail=""):
    R.append((no, name, ok, detail))

SEC = {s["no"]: s for c in SH.CH for s in c["sections"]}
def tbl(no, head0):
    for b in SEC[no]["blocks"]:
        if b["t"] == "table" and b["head"][0] == head0:
            return b
    return None

# 1 資料4-4 の施策一覧が素案の28施策と一致すること
soan = []
for sec in {c["no"]: c for c in SO.CH}["第4章"]["sections"]:
    if sec["no"].startswith("施策"):
        ti = sec["title"]
        soan.append((sec["no"].replace("施策", ""), ti.split("　【")[0], ti.split("【")[1].rstrip("】")))
    elif sec["no"] == "4-5":
        for b in sec["blocks"]:
            if b["t"] == "table":
                soan += [(r[0].split(" ", 1)[0], r[0].split(" ", 1)[1], "新設") for r in b["rows"]]
shiryo = [tuple(r[:3]) for b in SEC["4-4"]["blocks"] if b["t"] == "table" and b["head"][0] == "#"
          for r in b["rows"]]
bad = [f"{a[0]}:{a[1]}≠{b[1]}" for a, b in zip(soan, shiryo) if a[0] != b[0] or a[1] != b[1]]
chk(1, "資料4-4 の施策一覧が計画素案と一致", len(soan) == len(shiryo) == 28 and not bad,
    f"素案{len(soan)}／資料{len(shiryo)}・" + "・".join(bad[:3]) if (bad or len(shiryo) != 28)
    else "施策28本すべて一致")

# 2 資料2-4 の目標別表が交付金の算定と一致すること
F = "/home/user/repository/docs/北塩原村_第10期/data/交付金_目標別の県内比較_令和8年度.csv"
t = tbl("2-4", "交付金・目標")
if not os.path.exists(F) or t is None:
    chk(2, "資料2-4 の目標別表と交付金の算定の一致", False, "CSVまたは表がない")
else:
    D = {r["区分"]: r for r in csv.DictReader(open(F, encoding="utf-8-sig"))}
    bad2 = []
    for r in t["rows"]:
        k = ("推進" if r[0].startswith("推進") else "支援") + r[0][2]
        d = D.get(k)
        if not d:
            bad2.append(f"{k}が算定にない"); continue
        for i, key in ((1, "本村"), (2, "県平均"), (3, "県内順位"), (4, "全国平均"), (5, "全国順位")):
            if float(r[i].replace("位", "").replace(",", "")) != float(d[key]):
                bad2.append(f"{k} {key} 資料{r[i]}≠算定{d[key]}")
    chk(2, "資料2-4 の目標別表と交付金の算定の一致", not bad2,
        "・".join(bad2[:4]) if bad2 else "8目標×5項目すべて一致")

# 3 資料7-3c の取得をめざす評価指標が11項目・61点で、素案6-3と一致すること
t3 = tbl("7-3c", "交付金・指標")
so63 = None
for sec in {c["no"]: c for c in SO.CH}["第6章"]["sections"]:
    if sec["no"] == "6-3":
        for b in sec["blocks"]:
            if b["t"] == "table" and b["head"][0] == "交付金・指標":
                so63 = b
if t3 is None or so63 is None:
    chk(3, "資料7-3c と素案6-3 の一致", False, "表がない")
else:
    rows, last = t3["rows"][:-1], t3["rows"][-1]
    tot = sum(int(r[1]) for r in rows)
    n_ok = len(rows) == len(so63["rows"]) - 1 == 11
    pt_ok = all(int(a[1]) == int(b[1].replace("点", "")) for a, b in zip(rows, so63["rows"][:-1]))
    chk(3, "資料7-3c と素案6-3 の一致", tot == 61 and last[1] == "61" and n_ok and pt_ok,
        f"合計{tot}点／表記{last[1]}／項目{len(rows)}" if not (tot == 61 and n_ok and pt_ok)
        else "11項目・61点、素案6-3と配点が一致")

# 4 資料に貼る図のファイルが実在すること
figs = [b for c in SH.CH for s in c["sections"] for b in s["blocks"] if b["t"] == "fig"]
miss = [b["file"] for b in figs
        if not os.path.exists(f"/home/user/repository/output/figures/{b['file']}")]
chk(4, "資料に貼る図のファイルの実在", figs and not miss,
    f"欠落 {miss}" if miss else f"{len(figs)}点すべて実在")

# 5 古い誤り（認知症総合支援を「下回る唯一の目標」とする記述）が残っていないこと
src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "shiryo_content.py"), encoding="utf-8").read()
chk(5, "「下回る唯一の目標」の誤りが残っていないこと", "下回る唯一の目標" not in src)

w = max(len(n) for _, n, _, _ in R)
print("■ 第2回策定委員会 資料の自己点検")
ng = 0
for no, name, ok, detail in R:
    print(f"  {no:>3}  {name:<{w}}  {'適合' if ok else '不適合'}" + (f"  {detail}" if detail else ""))
    ng += 0 if ok else 1
print(f"\n  {len(R)}件のうち適合{len(R)-ng}件／不適合{ng}件")
sys.exit(1 if ng else 0)
