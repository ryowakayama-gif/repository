# -*- coding: utf-8 -*-
"""第2回策定委員会 資料の自己点検。

   資料は計画素案と交付金の算定から数値を書き写しているため、
   元を直したときの取り残しを検出する。不適合があれば終了コード1を返す。
"""
import csv, os, sys
sys.dont_write_bytecode = True   # 古いバイトコードで誤った結果が出ることを防ぐ
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import shiryo_content as SH
import soan_content as SO
import shukei_data as SK
from shihyo_dict import S as SHIHYO

MD26 = "/home/user/repository/docs/北塩原村_第10期/26_第2回策定委員会_資料構成.md"
MD24 = "/home/user/repository/docs/北塩原村_第10期/24_アンケート調査報告書_骨子案.md"

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
    # 集計後に埋める表は【集計後】の欄を持つ。対象者名簿から既に埋めた表（属性・課題の一覧）は除く
    withfill = [b for b in tbls
                if any("【集計後】" in str(c) for r in b["rows"] for c in r)]
    ok6 = len(secs) == 10 and len(tbls) >= 16 and len(withfill) >= 14
    chk(6, "資料3 の骨子（節数・数値欄）", ok6,
        f"節{len(secs)}／表{len(tbls)}／【集計後】を持つ表{len(withfill)}"
        if not ok6 else
        f"10節・{len(tbls)}表、うち{len(withfill)}表に【集計後】の欄")

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

# ── ここから分母の方針の反映の点検（doc26 §19）────────────
S3 = [x for x in SH.CH if x["no"] == "資料3"][0]["sections"]

def flat(sec):
    """1節のすべての文字列を連結する。"""
    out = []
    for b in sec["blocks"]:
        out.append(str(b.get("v", "")))
        out.append(str(b.get("head", "")) + str(b.get("rows", "")))
    return "".join(out)

S3TXT = "".join(flat(sc) for sc in S3)
md26 = open(MD26, encoding="utf-8").read()
md24 = open(MD24, encoding="utf-8").read()

# 11 【集計後】の欄を持つ表がある節に、分母の注記があること
need11 = [sc["no"] for sc in S3
          if any(b["t"] == "table" and any("【集計後】" in str(c) for r in b["rows"] for c in r)
                 for b in sc["blocks"])]
miss11 = [sc["no"] for sc in S3 if sc["no"] in need11
          and not any(b["t"] == "note" and "分母" in str(b["v"]) for b in sc["blocks"])]
chk(11, "資料3 の各節に分母の注記があること", need11 and not miss11,
    f"分母の注記がない節 {miss11}" if miss11 else f"{len(need11)}節すべてに分母の注記")

# 12 資料3 の分母が集計仕様書の定義と対応していること
X12 = ["X-01", "X-02", "X-03", "X-04", "X-05", "X-06",
       "X-08", "X-09", "X-10", "X-15", "X-16", "X-18"]
C = {c[0]: c for c in SK.CROSS}
miss12 = [k for k in X12 if k not in S3TXT]                      # 資料3が参照しているか
bad12 = [k for k in X12 if str(C[k][5]) not in md26]             # doc26 §19-2 に定義のまま載っているか
chk(12, "資料3 の分母が集計仕様書の定義と対応", not miss12 and not bad12,
    (f"資料3が参照していない {miss12}" if miss12 else "") +
    (f"／doc26に定義がない {bad12}" if bad12 else "")
    if (miss12 or bad12) else f"{len(X12)}表とも対応表と一致")

# 13 「回収した票の総数では割らない」の宣言と、回答方法の例外の明示
ok13 = ("回収した票の総数では割りません" in S3TXT
        and "回収した票の総数を分母とする" in S3TXT
        and "例外" in S3TXT)
chk(13, "総数で割らない宣言と例外の明示", ok13,
    "宣言または例外の注記がない" if not ok13 else "約束ごと1と唯一の例外（回答方法）を明記")

# 14 第1号被保険者数が指標辞書と一致し、資料5-2の2つの表で同じ値であること
SEC52 = [sc for sc in {c["no"]: c for c in SH.CH}["資料5"]["sections"] if sc["no"] == "5-2"][0]
t_an = [b for b in SEC52["blocks"] if b["t"] == "table" and b["head"][0] == "案"]
t_su = [b for b in SEC52["blocks"] if b["t"] == "table"
        and b["rows"] and b["rows"][0][0] == "第1号被保険者数"]
want = {y: f"{SHIHYO['第1号被保険者数_' + y][0]:,}人" for y in ("令和9年", "令和10年", "令和11年")}
bad14 = []
if not t_an or not t_su:
    bad14.append("資料5-2の表が見つからない")
else:
    ban = [r for r in t_an[0]["rows"] if r[0] == "B案"][0]
    for i, y in enumerate(("令和9年", "令和10年", "令和11年")):
        if ban[2 + i] != want[y]:
            bad14.append(f"B案 {y} 資料{ban[2+i]}≠辞書{want[y]}")
    head, row = t_su[0]["head"], t_su[0]["rows"][0]
    for y in ("令和9年", "令和10年", "令和11年"):
        if y in head and row[head.index(y)] != want[y]:
            bad14.append(f"推計表 {y} 資料{row[head.index(y)]}≠辞書{want[y]}")
    if f"{SHIHYO['通いの場_分母'][0]:,}" not in "".join(
            str(b.get("v", "")) for sc in {c["no"]: c for c in SH.CH}["資料7"]["sections"]
            for b in sc["blocks"]):
        bad14.append("資料7に通いの場の分母1,009人の記載がない")
chk(14, "第1号被保険者数が指標辞書と一致", not bad14,
    "・".join(bad14[:3]) if bad14 else "令和9〜11年がB案・推計表・辞書で一致（1,004／995／987人）")

# 15 認定状況の4区分が資料3・doc24・集計仕様書で一致すること
KUBUN = ["認定なし", "事業対象者", "要支援1・2", "要介護1以上"]
axE = [a for a in SK.AXES if a[0] == "E"][0][2]
t33 = [b for sc in S3 if sc["no"] == "3-3" for b in sc["blocks"]
       if b["t"] == "table" and b["head"][0] == "認定の状況"]
bad15 = []
if not t33:
    bad15.append("資料3-3の表が見つからない")
elif [r[0] for r in t33[0]["rows"][:4]] != KUBUN:
    bad15.append("資料3-3の区分 " + "／".join(r[0] for r in t33[0]["rows"][:4]))
if "／".join(KUBUN) not in axE:
    bad15.append(f"集計軸E {axE}")
if "／".join(KUBUN) not in md24:
    bad15.append("doc24に4区分の記載がない")
chk(15, "認定状況の4区分が3文書で一致", not bad15,
    "・".join(bad15[:3]) if bad15 else "認定なし／事業対象者／要支援1・2／要介護1以上")

# 16 doc26 §18-5 の資料の状況（節数・表数）が資料の実体と一致すること
#    資料を直して一覧の更新を忘れることを検出する（§19-7）。
blk = md26.split("### 18-5")[-1].split("\n---\n")[0] if "### 18-5" in md26 else ""
listed = {}
for ln in blk.split("\n"):
    cs = [c.strip() for c in ln.strip().strip("|").split("|")] if ln.strip().startswith("|") else []
    if len(cs) >= 4 and (cs[0].startswith("資料") or "合計" in cs[1]):
        try:
            listed[cs[0] or "合計"] = (int(cs[2].strip("*")),
                                       int(cs[3].strip("*").split("＋")[0]))
        except ValueError:
            pass
real = {c["no"]: (len(c["sections"]),
                  sum(1 for sc in c["sections"] for b in sc["blocks"] if b["t"] == "table"))
        for c in SH.CH}
real["合計"] = (sum(v[0] for v in real.values()), sum(v[1] for v in real.values()))
bad16 = [f"{k} 一覧{listed[k]}≠実体{real[k]}" for k in real if k in listed and listed[k] != real[k]]
miss16 = [k for k in real if k not in listed]
chk(16, "doc26 §18-5 の節数・表数が資料と一致", not bad16 and not miss16,
    ("・".join(bad16[:3]) + ("／一覧にない " + str(miss16) if miss16 else ""))
    if (bad16 or miss16) else f"全8資料と合計（{real['合計'][0]}節・{real['合計'][1]}表）が一致")

# 17 資料5-2 の認定者数の系列が計画素案 5-3 と一致すること
#    他案件（金ケ崎町）で「同じ指標の推計が資料ごとに別系列だった」ことに倣う。
S52 = [sc for sc in {c["no"]: c for c in SH.CH}["資料5"]["sections"] if sc["no"] == "5-2"][0]
t_sh = [b for b in S52["blocks"] if b["t"] == "table"
        and b["rows"] and b["rows"][0][0] == "第1号被保険者数"]
S53 = [sc for sc in {c["no"]: c for c in SO.CH}["第5章"]["sections"] if sc["no"] == "5-3"][0]
t_so = [b for b in S53["blocks"] if b["t"] == "table" and b["head"][0] == ""
        and any(r[0].startswith("令和9") for r in b["rows"])]
LV = ["要支援1", "要支援2", "要介護1", "要介護2", "要介護3", "要介護4", "要介護5"]
bad17 = []
if not t_sh or not t_so:
    bad17.append("資料5-2または素案5-3の表がない")
else:
    sh, so = t_sh[0], t_so[0]
    soan = {r[0].replace("（実績）", ""): r for r in so["rows"]}
    for j, y in enumerate(sh["head"][1:], start=1):
        key = y.replace("（実績）", "")
        if key not in soan:
            continue
        for lv in LV:
            row = [r for r in sh["rows"] if r[0] == lv]
            if not row:
                bad17.append(f"資料に{lv}の行がない"); continue
            a = str(row[0][j])
            b = str(soan[key][LV.index(lv) + 1])
            if a != b:
                bad17.append(f"{key} {lv} 資料{a}≠素案{b}")
        kei = [r for r in sh["rows"] if r[0] == "認定者計"]
        if kei and str(kei[0][j]).replace("人", "") != str(soan[key][8]):
            bad17.append(f'{key} 計 資料{kei[0][j]}≠素案{soan[key][8]}')
        nri = [r for r in sh["rows"] if r[0] == "認定率"]
        if nri and str(nri[0][j]) != str(soan[key][9]):
            bad17.append(f"{key} 認定率 資料{nri[0][j]}≠素案{soan[key][9]}")
chk(17, "資料5-2 の認定者数が計画素案 5-3 と一致", not bad17,
    "・".join(bad17[:3]) if bad17 else "令和8〜11年の要介護度別・計・認定率が一致")

# 18 委員会資料の本文に受託者の内部の仕組みの語がないこと
NAIGO = ["scripts/", ".py", "verify_", ".csv", "doc2", "doc4", "doc1", "Python"]
hit18 = []
for c in SH.CH:
    for sc in c["sections"]:
        for b in sc["blocks"]:
            for v in ([str(b.get("v", ""))]
                      + [str(x) for r in b.get("rows", []) for x in r]
                      + [str(x) for x in b.get("head", [])]
                      + [str(b.get("caption", "")), str(b.get("source", ""))]):
                for g in NAIGO:
                    if g in v:
                        hit18.append(f'{c["no"]}{sc["no"]}:{g}')
chk(18, "委員会資料の本文に内部の仕組みの語がないこと", not hit18,
    "・".join(sorted(set(hit18))[:4]) if hit18 else f"{len(NAIGO)}語のいずれも本文にない")

# 19 交付金の表が、計画素案と委員会資料（資料2・資料7）で一致すること
#    同じ数字を2つの成果品に書き写しているため、片方を直したときの取り残しを検出する。
import re as _re2

def _num(x):
    """数値のセルだけを比べる。単位・記号は落とし、節番号（4-7 等）は比較の対象外とする。"""
    v = _re2.sub(r'[\s\u3000点位％%人項目計配得,〜～]', '', str(x))
    v = v.replace('▲', '-').replace('＋', '+')
    return v if _re2.match(r'^[-+]?[\d.]+$', v) else None

def _find(sec, pred):
    for b in sec['blocks']:
        if b['t'] == 'table' and pred(b):
            return b
    return None

SO_S = {sc['no']: sc for c in SO.CH for sc in c['sections']}
SH_S = {sc['no']: sc for c in SH.CH for sc in c['sections']}
PRED = {
 '3か年の得点推移': lambda b: b['head'][0] == '' and b['rows'][0][0].startswith('保険者機能強化推進交付金'),
 '推移の4区分':    lambda b: b['head'] == ['区分', '項目数', '点'],
 '目標別の8目標':   lambda b: b['head'][0] == '交付金・目標' and '県内順位' in b['head'] and '全国順位' in b['head'],
 '政策領域の差':    lambda b: '本村の差' in b['head'] and '福島県の差' in b['head'],
 '未取得の主な項目': lambda b: b['head'] == ['交付金・目標', '評価指標', '配点', '全国該当率'],
 '61点の一覧':     lambda b: b['head'][0] == '交付金・指標',
 '122点の4区分':   lambda b: b['head'] == ['区分', '項目数', '点', '内容'],
}
WHERE = {
 '3か年の得点推移': ('2-8', '2-4'), '推移の4区分': ('2-8', '2-4'),
 '目標別の8目標': ('2-8', '2-4'), '政策領域の差': ('2-8', '2-4'),
 '未取得の主な項目': ('2-8', '2-4'), '61点の一覧': ('6-3', '7-3c'),
 '122点の4区分': ('6-3', '7-3c'),
}
bad19 = []
n_ok = 0
for name, (so_no, sh_no) in WHERE.items():
    a = _find(SO_S[so_no], PRED[name])
    b = _find(SH_S[sh_no], PRED[name])
    if a is None or b is None:
        bad19.append(f'{name}が{"素案" if a is None else "資料"}にない')
        continue
    if len(a['rows']) != len(b['rows']):
        bad19.append(f'{name} 行数 素案{len(a["rows"])}≠資料{len(b["rows"])}')
        continue
    for i, (ra, rb) in enumerate(zip(a['rows'], b['rows'])):
        va = [y for y in (_num(x) for x in ra) if y is not None]
        vb = [y for y in (_num(x) for x in rb) if y is not None]
        if va != vb:
            bad19.append(f'{name} 行{i+1}（{str(ra[0])[:14]}） 素案{va}≠資料{vb}')
    n_ok += 1
chk(19, "交付金の表が計画素案と委員会資料で一致すること", not bad19,
    "・".join(bad19[:3]) if bad19 else f"{n_ok}表（3か年の推移・4区分・目標別・政策領域・未取得・61点・122点）が一致")

w = max(len(n) for _, n, _, _ in R)
print("■ 第2回策定委員会 資料の自己点検")
ng = 0
for no, name, ok, detail in R:
    print(f"  {no:>3}  {name:<{w}}  {'適合' if ok else '不適合'}" + (f"  {detail}" if detail else ""))
    ng += 0 if ok else 1
print(f"\n  {len(R)}件のうち適合{len(R)-ng}件／不適合{ng}件")
sys.exit(1 if ng else 0)
