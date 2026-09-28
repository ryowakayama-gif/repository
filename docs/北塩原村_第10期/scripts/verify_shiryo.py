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

# 6 資料3 は集計前の骨子であり、数値欄が【集計後】で埋まっていること
s3 = [x for x in SH.CH if x["no"] == "資料3"]
if not s3:
    chk(6, "資料3 の骨子", False, "資料3がない")
else:
    secs = s3[0]["sections"]
    tbls = [b for sc in secs for b in sc["blocks"] if b["t"] == "table"]
    # 集計後に埋める欄が1つもない表は、数値を書き込み済みか枠の作り忘れ
    nofill = [i for i, b in enumerate(tbls)
              if not any("【集計後】" in str(c) for r in b["rows"] for c in r)]
    chk(6, "資料3 の骨子（節数・数値欄）",
        len(secs) == 10 and len(tbls) == 16 and len(nofill) <= 2,
        f"節{len(secs)}／表{len(tbls)}／数値欄のない表{len(nofill)}"
        if not (len(secs) == 10 and len(tbls) == 16 and len(nofill) <= 2)
        else f"10節・16表、うち14表に【集計後】の欄")

# 7 資料3 の作成に必要な前提（未決の2件）が確認事項に残っていること
import re as _re
KK = open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "wbs_kakunin.py"), encoding="utf-8").read()
need = [("問4(15)(16)", "リスク判定の領域数"), ("要介護認定データの提供時期", "認定データの関連付け")]
miss7 = [lab for key, lab in need if key not in KK]
chk(7, "資料3 の前提となる確認事項の存在", not miss7,
    f"確認事項から消えている: {miss7}" if miss7 else "2件とも確認事項に存在")

# 8 資料8 の日付が仕様書・確定済みの日程と食い違っていないこと
s8 = [x for x in SH.CH if x["no"] == "資料8"]
if not s8:
    chk(8, "資料8 のスケジュール", False, "資料8がない")
else:
    txt = "".join(str(b.get("v", "")) + str(b.get("rows", ""))
                  for sc in s8[0]["sections"] for b in sc["blocks"])
    need = {"令和9年3月31日": "履行期限（仕様書3）",
            "令和8年10月16日": "調査票の回収期限",
            "令和9年度から令和11年度": "計画期間（仕様書）",
            "882人": "ニーズ調査の対象者数",
            "118人": "在宅介護実態調査の対象者数"}
    miss8 = [v for k, v in need.items() if k not in txt]
    # 村が実施する工程に受託者を主体として書いていないこと
    bad8 = [r[0] for sc in s8[0]["sections"] for b in sc["blocks"] if b["t"] == "table"
            and b["head"][0] == "時期"
            for r in b["rows"] if "パブリックコメントの実施" in r[1] and r[2] != "村"]
    chk(8, "資料8 の日程・主体", not miss8 and not bad8,
        ("欠落: " + "・".join(miss8) if miss8 else "") +
        ("／パブコメの主体が村でない" if bad8 else "")
        if (miss8 or bad8) else f"{len(need)}件の日程・数値と主体の区分が整合")

# 9 資料1 の次第・議論事項・配付資料一覧が資料の実体と合っていること
s1 = [x for x in SH.CH if x["no"] == "資料1"]
if not s1:
    chk(9, "資料1 の骨子", False, "資料1がない")
else:
    S1 = {sc["no"]: sc for sc in s1[0]["sections"]}
    bad9 = []
    # 次第の所要時間が120分であること
    t11 = [b for b in S1["1-1"]["blocks"] if b["t"] == "table"][0]
    mins = sum(int(r[3].replace("分", "")) for r in t11["rows"]
               if r[3].endswith("分") and r[3] != "120分")
    if mins != 120:
        bad9.append(f"次第の所要時間の合計が{mins}分")
    if t11["rows"][-1][3] != "120分":
        bad9.append(f"次第の合計欄が{t11['rows'][-1][3]}")
    # ご議論いただきたい事項が9点であること
    t12 = [b for b in S1["1-2"]["blocks"] if b["t"] == "table"][0]
    if len(t12["rows"]) != 9:
        bad9.append(f"議論事項が{len(t12['rows'])}点")
    # 配付資料一覧の名称が実体と一致すること
    t15 = [b for b in S1["1-5"]["blocks"] if b["t"] == "table"][0]
    real = {c["no"]: c["title"] for c in SH.CH}
    for r in t15["rows"]:
        if r[0] in real and r[1] != real[r[0]]:
            bad9.append(f"{r[0]} の名称 {r[1]}≠{real[r[0]]}")
    listed = {r[0] for r in t15["rows"]}
    miss9 = [k for k in real if k not in listed]
    if miss9:
        bad9.append(f"配付資料一覧にない: {miss9}")
    chk(9, "資料1 の次第・議論事項・配付資料一覧", not bad9,
        "・".join(bad9[:3]) if bad9 else "次第120分・議論9点・全8資料を掲載")

# 10 資料1 の別冊（計画素案）の頁数が積算と大きくずれていないこと
try:
    import estimate_pages as EP
    rows, _tot, _nt, _nf = EP.run()
    est = sum(__import__("math").ceil(r[5]) for r in rows) + sum(n for _, n in EP.FRONT)
    t15 = [b for b in [x for x in SH.CH if x["no"] == "資料1"][0]["sections"][4]["blocks"]
           if b["t"] == "table"][0]
    bess = [r[2] for r in t15["rows"] if r[0] == "別冊"]
    got = int(str(bess[0]).replace("約", "")) if bess else None
    chk(10, "資料1 の別冊の頁数と積算の整合", got is not None and abs(got - est) <= 3,
        f"資料1は{got}頁・積算は{est}頁" if got is None or abs(got - est) > 3
        else f"資料1 約{got}頁／積算{est}頁")
except Exception as e:
    chk(10, "資料1 の別冊の頁数と積算の整合", False, f"照合できない（{e}）")

w = max(len(n) for _, n, _, _ in R)
print("■ 第2回策定委員会 資料の自己点検")
ng = 0
for no, name, ok, detail in R:
    print(f"  {no:>3}  {name:<{w}}  {'適合' if ok else '不適合'}" + (f"  {detail}" if detail else ""))
    ng += 0 if ok else 1
print(f"\n  {len(R)}件のうち適合{len(R)-ng}件／不適合{ng}件")
sys.exit(1 if ng else 0)
