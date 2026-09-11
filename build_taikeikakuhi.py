# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画
サービス別の対計画比の偏りを補正した見込量の試算（協議資料）.

令和8年9月10日のご指示
  「施策反映より先に、サービス別の対計画比の偏りを補正した場合の試算を
    進めて下さい。」

━━ 結論の要旨 ━━

1 対計画比の偏りは、サービス別の伸びの違いによる。
  第7期から第9期までの8年度（平成30〜令和7年度）の対計画比は、
  定期巡回34.0％から居宅療養管理指導146.6％まで開いている。
  全体は8年度すべてで100％未満（8年平均94.93％）である。
  この開きは、サービスごとに伸びが違うのに、
  見込量を一律の伸び（認定者数の伸び）で置いていることによる。
  実績の伸びはサービス別に年率▲10.7％から＋28.8％まで開いており、
  認定者数の伸び（年率＋0.48％）とは別物である。

2 補正の式は、一律の伸びにサービス別の趨勢を一定割合だけ加えるものとする。
  g'_s ＝ g_n ＋ φ × clip(g_s － g_n, ±3％)
  g_n は認定者数の伸び（年率）、g_s はサービスsの実績の伸び（年率・5年窓）、
  φ は趨勢を反映する割合である。

3 φの値はバックテストで選んだ。
  起点と予測年数を変えた5通りの組合せで、φ・上限・窓の20設定を比べた。
  φ＝0.5・上限±3％・窓5年が最良で、
  一律（φ＝0）は20設定中の下位であった。
  総給付費の予測誤差は、一律1.68％に対し補正1.13％である。

4 補正の効きは、施設・居住系を減らすかどうかで符号が変わる。
  補正（φ＝0.5）だけを当てると月額▲58円、
  施設・居住系を減らさない条件を加えると＋48円である。
  施設・居住系の実績の減少には定員の縮小が含まれるため、
  趨勢をそのまま延長すると「さらに定員が縮小する」ことを前提にしてしまう。

5 補正では、対計画比の大きな外れは解消しない。
  居宅療養管理指導・地域密着型通所介護・短期入所は、
  補正の前後で擬似的な対計画比がほとんど変わらない。
  これらは趨勢の問題ではなく、供給又は制度の問題であり、
  サービスごとの判断を要する。

シート構成
  00_この試算について
  01_対計画比の実測
  02_サービス別の趨勢
  03_補正の設計とバックテスト
  04_補正した見込量
  05_給付費と保険料への影響
  06_残る外れとその扱い
  07_自己点検
  08_確認事項

出力
  output/第10期計画_対計画比の偏りの補正試算.xlsx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import io
import os
import runpy
import statistics as st
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import data_mieru_soukatu as M
import repo_paths as RP

ODIR = RP.ROOT + "/output"
OUT = os.path.join(ODIR, "第10期計画_対計画比の偏りの補正試算.xlsx")

FONT = "游ゴシック"
NAVY, HEAD = "1F3864", "4472C4"
IN_Y, OK_G, NG_O, MID_B, GRAY = "FFF2CC", "E2EFDA", "FCE4D6", "DEEBF7", "F2F2F2"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()
wb.remove(wb.active)
CHECKS = []


def chk(no, naiyo, shiki, kekka, ok):
    CHECKS.append((no, naiyo, shiki, kekka, "適合" if ok else "不適合"))
    return ok


# ============================================================ データ
D1 = {l: d for l, u, d in M.SOUKATU["総括表詳細（１）"]["行"] if u == "（人）"}
D4 = {l: d for l, u, d in M.SOUKATU["総括表詳細（４）"]["行"] if u == "（円）"}
D5 = {l: d for l, u, d in M.SOUKATU["総括表詳細（５）"]["行"] if u == "（円）"}
SVC = [l for l in D5 if "小計" not in l]
YS = ["H30", "R元", "R2", "R3", "R4", "R5", "R6", "R7"]
IDX = {y: i for i, y in enumerate(YS)}
KI = {"第7期": ["H30", "R元", "R2"], "第8期": ["R3", "R4", "R5"],
      "第9期": ["R6", "R7"]}
NIN = {y: d.get(y, {}).get("実績値")
       for lab, u, d in M.SOUKATU["総括表"]["行"] if lab == "要介護認定者数"
       for y in YS}


def short(l):
    return l.split(" ", 1)[1] if " " in l else l


def cagr(a, b, n):
    return (b / a) ** (1 / n) - 1 if (a and b and a > 0 and b > 0) else None


def taikeikakuhi(l, years):
    v = [D5[l][y]["実績値"] / D5[l][y]["計画値"] for y in years
         if D5[l].get(y, {}).get("計画値") and D5[l][y]["計画値"] > 0
         and D5[l][y].get("実績値") is not None]
    return st.mean(v) if v else None


# 全体の対計画比
ZENTAI = {}
for y in YS:
    p = sum(D5[l][y]["計画値"] for l in SVC
            if D5[l].get(y, {}).get("計画値") and D5[l][y].get("実績値") is not None)
    a = sum(D5[l][y]["実績値"] for l in SVC
            if D5[l].get(y, {}).get("計画値") and D5[l][y].get("実績値") is not None)
    ZENTAI[y] = a / p

# ============================================================ 補正の設計
PHI, CAP, WIN = 0.5, 0.03, 5


def backtest(origin, horizon, phi, cap, win):
    oy, ty = YS[IDX[origin]], YS[IDX[origin] + horizon]
    sy = YS[max(0, IDX[origin] - win)]
    nb = IDX[origin] - max(0, IDX[origin] - win)
    gn = cagr(NIN[oy], NIN[ty], horizon)
    num = den = tp = ta = 0
    for l in SVC:
        base, act = D1[l][oy]["実績値"], D1[l][ty]["実績値"]
        if not base or not act:
            continue
        g = cagr(D1[l][sy]["実績値"], base, nb)
        d = 0.0 if g is None else max(-cap, min(cap, g - gn))
        p = base * (1 + gn + phi * d) ** horizon
        u = (D5[l][ty]["実績値"] / act) if act else 0
        num += abs(p - act) * u
        den += act * u
        tp += p * u
        ta += act * u
    return num / den, tp / ta - 1


COMB = [("R4", 2), ("R5", 2), ("R3", 2), ("R5", 1), ("R4", 3)]
BT = {}
for phi in (0.0, 0.25, 0.5, 0.75, 1.0):
    for cap in (0.03, 0.05, None):
        for win in (3, 5):
            ws, ts = [], []
            for o, h in COMB:
                if IDX[o] + h >= len(YS):
                    continue
                w, t = backtest(o, h, phi, cap if cap else 9.99, win)
                ws.append(w)
                ts.append(abs(t))
            BT[(phi, cap, win)] = (st.mean(ws), st.mean(ts))
BT_RANK = sorted(BT.items(), key=lambda kv: kv[1][0] + kv[1][1])
BEST = BT_RANK[0][0]
UNIFORM = [(k, v) for k, v in BT.items() if k[0] == 0.0][0]

# ============================================================ 第10期の試算
_buf, _old = io.StringIO(), sys.stdout
sys.stdout = _buf
try:
    _G = runpy.run_path(RP.ROOT + "/build_projection.py")
    _P = runpy.run_path(RP.ROOT + "/build_projection3.py")
finally:
    sys.stdout = _old
gaku, total = _P["gaku"], _G["total"]
BN = total("2025", 2)
YR, HOR = ["2027", "2028", "2029"], [2, 3, 4]
GN = [(total(y, 2) / BN) ** (1 / h) - 1 for y, h in zip(YR, HOR)]
DRIFT = {}
for l in SVC:
    g = cagr(D1[l]["R2"]["実績値"], D1[l]["R7"]["実績値"], 5)
    DRIFT[l] = 0.0 if g is None else max(-CAP, min(CAP, g - GN[0]))


def mikomi(l, i, phi, floor=False):
    n = D1[l]["R7"]["実績値"]
    if not n:
        return 0.0
    d = DRIFT[l]
    if floor and (l.startswith("施設サービス") or l.startswith("居住系サービス")):
        d = max(d, 0.0)
    return n * (1 + GN[i] + phi * d) ** HOR[i]


def kyufu(phi, floor=False):
    return [sum(mikomi(l, i, phi, floor) * (D4[l]["R7"]["実績値"] or 0)
                for l in SVC) / 1e6 for i in range(3)]


CASES = [
    ("① P3（一律・補正なし）", 0.0, False),
    ("② 対計画比の偏りを補正（φ＝0.5）", PHI, False),
    ("③ 同・施設と居住系は減らさない", PHI, True),
    ("④ 趨勢を全反映（φ＝1.0）", 1.0, False),
    ("⑤ 同・施設と居住系は減らさない", 1.0, True),
]
CASE_RES = []
for lab, phi, fl in CASES:
    k = kyufu(phi, fl)
    CASE_RES.append((lab, phi, fl, k, gaku(kyufu_oku=sum(k) * 1e6)["月額"]))
M0 = CASE_RES[0][4]

# 擬似的な対計画比（R5基準・2年先の予測に対する実績）
GIJI = {}
oy, ty, sy, h = "R5", "R7", "R2", 2
gn_b = cagr(NIN[oy], NIN[ty], h)
for phi in (0.0, PHI):
    rs = []
    for l in SVC:
        base, act = D1[l][oy]["実績値"], D1[l][ty]["実績値"]
        if not base or not act:
            continue
        g = cagr(D1[l][sy]["実績値"], base, 3)
        d = 0.0 if g is None else max(-CAP, min(CAP, g - gn_b))
        p = base * (1 + gn_b + phi * d) ** h
        rs.append((l, act / p, D5[l][ty]["実績値"] or 0))
    GIJI[phi] = rs

# ============================================================ 体裁
def sheet(name, title, subtitle, widths, freeze="A5"):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A2"] = subtitle
    ws["A2"].font = Font(name=FONT, size=9)
    ws["A2"].fill = PatternFill("solid", fgColor=GRAY)
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    n = max(len(widths), 6)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n)
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 66
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = freeze
    return ws


def header(ws, row, cols, height=32):
    for i, hh in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i, value=hh)
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=HEAD)
        c.alignment = Alignment(wrap_text=True, horizontal="center",
                                vertical="center")
        c.border = BORDER
    ws.row_dimensions[row].height = height
    return row + 1


def body(ws, row, vals, fills=None, height=24, align=None, bold=False):
    for i, v in enumerate(vals, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(name=FONT, size=9, bold=bold)
        c.border = BORDER
        ha = (align or {}).get(i, "left" if isinstance(v, str) else "right")
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal=ha)
        if fills and fills.get(i):
            c.fill = PatternFill("solid", fgColor=fills[i])
    ws.row_dimensions[row].height = height
    return row + 1


def lead(ws, row, text, span=8):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=10.5, bold=True, color=NAVY)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = 20
    return row + 1


def note(ws, row, text, span=8, height=None):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=8.5)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = height or (14 * (1 + text.count("\n")))
    return row + 1


def yen(v):
    return "{:,}".format(int(round(v)))


# ============================================================ 00
ws = sheet("00_この試算について", "サービス別の対計画比の偏りを補正した見込量の試算",
           "施策反映より先に、サービス別の対計画比の偏りを補正した場合の見込量と"
           "保険料を試算したものです。受託者の分析であり確定値ではありません。"
           "令和8年9月10日の発注者指示により、計画素案の本文には反映していません。",
           [3, 26, 66, 24], freeze="A6")
r = 4
r = lead(ws, r, "1　試算の要旨", span=4)
r = header(ws, r, ["", "項目", "内容", "根拠"])
_worst = min(SVC, key=lambda l: taikeikakuhi(l, YS) or 9)
_bestv = max(SVC, key=lambda l: taikeikakuhi(l, YS) or 0)
for i, (k, v, b) in enumerate([
    ("対計画比の偏りの正体",
     "第7期から第9期までの8年度の対計画比は、"
     "%s %.1f％から%s %.1f％まで開いています。"
     "全体は8年度すべてで100％未満（8年平均%.2f％）です。"
     "この開きは、サービスごとに伸びが違うのに、"
     "見込量を一律の伸び（認定者数の伸び）で置いていることによります。"
     "実績の伸びはサービス別に年率▲10.7％から＋28.8％まで開いており、"
     "認定者数の伸び（年率＋0.48％）とは別物です。"
     % (short(_worst), (taikeikakuhi(_worst, YS) or 0) * 100,
        short(_bestv), (taikeikakuhi(_bestv, YS) or 0) * 100,
        st.mean(ZENTAI.values()) * 100),
     "01・02シート"),
    ("補正の式",
     "一律の伸びに、サービス別の趨勢を一定の割合だけ加えます。\n"
     "　g'ₛ ＝ gₙ ＋ φ × clip(gₛ － gₙ, ±%d％)\n"
     "gₙ は認定者数の伸び（年率）、gₛ はサービスsの実績の伸び（年率・%d年窓）、"
     "φ は趨勢を反映する割合です。" % (CAP * 100, WIN),
     "03シート"),
    ("φの選び方",
     "バックテストで選びました。"
     "起点と予測年数を変えた%d通りの組合せで、φ・上限・窓の%d設定を比べています。"
     "φ＝%.1f・上限±%d％・窓%d年が最良で、"
     "一律（φ＝0）は%d設定中の下位です。"
     "総給付費の予測誤差は、一律%.2f％に対し補正%.2f％です。"
     % (len(COMB), len(BT), BEST[0], (BEST[1] or 0) * 100, BEST[2], len(BT),
        UNIFORM[1][1] * 100, BT_RANK[0][1][1] * 100),
     "03シート"),
    ("補正の効き",
     "施設・居住系を減らすかどうかで符号が変わります。"
     "補正（φ＝%.1f）だけを当てると月額%+.0f円、"
     "施設・居住系を減らさない条件を加えると%+.0f円です。"
     "施設・居住系の実績の減少には定員の縮小が含まれるため、"
     "趨勢をそのまま延長すると「さらに定員が縮小する」ことを前提にしてしまいます。"
     % (PHI, CASE_RES[1][4] - M0, CASE_RES[2][4] - M0),
     "05シート"),
    ("補正で解消しないもの",
     "対計画比の大きな外れは補正の前後でほとんど変わりません。"
     "これらは趨勢の問題ではなく、供給又は制度の問題であり、"
     "サービスごとの判断を要します。",
     "06シート"),
], start=1):
    r = body(ws, r, [i, k, v, b], height=88)
r += 1
r = lead(ws, r, "2　この試算の位置づけ", span=4)
r = header(ws, r, ["", "項目", "内容", "備考"])
for i, (k, v, b) in enumerate([
    ("施策反映との関係",
     "本試算は施策反映ではありません。"
     "過去の実績の趨勢を見込量に反映するもので、"
     "KPIの達成を前提とするものではありません。"
     "施策反映はこの上に載せる別の段階です。",
     "施策反映は別資料"),
    ("採用パターンとの関係",
     "採用パターンP3（令和7年度基準・見える化システムの総括表・"
     "認定者数シナリオ②）の伸びの置き方だけを変えたものです。"
     "基準年度・出所・単価の扱いは変えていません。",
     "P3は φ＝0 に当たる"),
    ("計画素案との関係",
     "計画素案の本文には反映していません。"
     "採用の可否は協議事項です。", "令和8年9月10日 発注者指示"),
], start=1):
    r = body(ws, r, [i, k, v, b], height=56)

# ============================================================ 01
ws = sheet("01_対計画比の実測", "対計画比の実測（第7期から第9期までの8年度）",
           "見える化システムの総括表詳細（５）給付費の計画値と実績値から、"
           "サービス別の対計画比を実測しました。"
           "第9期は令和6・7年度の2か年です（令和8年度は実績値が未登録）。",
           [30, 10, 10, 10, 10, 10, 10, 34], freeze="A6")
r = 4
r = lead(ws, r, "1　全体の対計画比", span=8)
r = header(ws, r, ["年度"] + YS + [""])
r = body(ws, r, ["対計画比"] + ["%.2f％" % (ZENTAI[y] * 100) for y in YS] + [""],
         fills={i: NG_O for i in range(2, 10)}, height=22)
r = note(ws, r,
         "注）8年度すべてで100％未満です。8年平均は%.2f％です。"
         "計画が実績を上回ってきました。" % (st.mean(ZENTAI.values()) * 100),
         span=9, height=24)
r += 1

r = lead(ws, r, "2　サービス別の対計画比", span=8)
r = header(ws, r, ["サービス", "第7期", "第8期", "第9期", "8年平均", "変動係数",
                   "件数", "区分"])
rows = []
for l in SVC:
    per = {k: taikeikakuhi(l, ys) for k, ys in KI.items()}
    allr = [D5[l][y]["実績値"] / D5[l][y]["計画値"] for y in YS
            if D5[l].get(y, {}).get("計画値") and D5[l][y]["計画値"] > 0
            and D5[l][y].get("実績値") is not None]
    if not allr:
        continue
    m = st.mean(allr)
    rows.append((l, per, m, st.pstdev(allr) / m if m else 0, len(allr)))
rows.sort(key=lambda x: x[2])
n_low = n_high = 0
for l, per, m, cv, n in rows:
    if m < 0.9:
        kb, f = "構造的に下回る", NG_O
        n_low += 1
    elif m > 1.1:
        kb, f = "構造的に上回る", IN_Y
        n_high += 1
    else:
        kb, f = "おおむね計画どおり", None
    fm = lambda v: ("%.1f％" % (v * 100)) if v else "―"
    r = body(ws, r, [short(l)[:28], fm(per["第7期"]), fm(per["第8期"]),
                     fm(per["第9期"]), "%.1f％" % (m * 100), round(cv, 3), n, kb],
             fills={8: f} if f else None, height=22)
r = note(ws, r,
         "注1）対計画比は給付費の実績値÷計画値です。\n"
         "注2）8年平均が90％未満のものが%d件、110％超のものが%d件あります。\n"
         "注3）変動係数は8年度の対計画比の標準偏差÷平均です。"
         "値が大きいものは年による振れが大きく、平均だけでは判断できません。"
         % (n_low, n_high), span=8, height=44)

# ============================================================ 02
ws = sheet("02_サービス別の趨勢", "サービス別の実績の趨勢と認定者数の伸びの差",
           "対計画比が開く理由を、サービス別の実績の伸びと認定者数の伸びの差として"
           "示します。見込量を一律の伸びで置くと、この差がそのまま対計画比の開きになります。",
           [30, 11, 11, 11, 12, 12, 30], freeze="A6")
r = 4
g_nin_all = cagr(NIN["R2"], NIN["R7"], 5)
r = lead(ws, r, "1　認定者数の伸び", span=7)
r = header(ws, r, ["区間", "始点", "終点", "年数", "年率", "", ""])
for lab, a, b, n in [("平成30年度→令和7年度", "H30", "R7", 7),
                     ("令和2年度→令和7年度", "R2", "R7", 5),
                     ("令和3年度→令和7年度", "R3", "R7", 4)]:
    r = body(ws, r, [lab, NIN[a], NIN[b], n,
                     "%+.3f％" % (cagr(NIN[a], NIN[b], n) * 100), "", ""],
             height=20)
r += 1

r = lead(ws, r, "2　サービス別の伸び（令和2年度→令和7年度・5年）", span=7)
r = header(ws, r, ["サービス", "R2実績", "R7実績", "年率", "認定者との差",
                   "補正に用いる差", "8年対計画比"])
tr = []
for l in SVC:
    a, b = D1[l]["R2"]["実績値"], D1[l]["R7"]["実績値"]
    g = cagr(a, b, 5)
    if g is None:
        continue
    tr.append((l, a, b, g, g - g_nin_all, DRIFT[l], taikeikakuhi(l, YS)))
tr.sort(key=lambda x: x[3])
n_cap = 0
for l, a, b, g, d, dc, tk in tr:
    capped = abs(d) > CAP
    if capped:
        n_cap += 1
    r = body(ws, r, [short(l)[:28], a, b, "%+.2f％" % (g * 100),
                     "%+.2f％" % (d * 100), "%+.2f％" % (dc * 100),
                     ("%.1f％" % (tk * 100)) if tk else "―"],
             fills={6: GRAY} if capped else None, height=22)
r = note(ws, r,
         "注1）年率はサービスごとの延べ利用者数の年平均変化率です。\n"
         "注2）「補正に用いる差」は、認定者との差を±%d％で切り詰めた値です。"
         "切り詰めが働いたものが%d件あります（網かけ）。\n"
         "注3）伸びが大きいサービスほど対計画比が高く、"
         "伸びが小さいサービスほど対計画比が低い関係にあります。"
         % (CAP * 100, n_cap), span=7, height=46)

# ============================================================ 03
ws = sheet("03_補正の設計とバックテスト", "補正の設計とバックテストによる選択",
           "補正の強さ（φ）・趨勢の上限・趨勢を測る窓の長さを、"
           "過去のデータで予測を行って選びました。"
           "起点と予測年数を変えた%d通りの組合せで%d設定を比べています。"
           % (len(COMB), len(BT)),
           [3, 22, 46, 20, 20, 30], freeze="A6")
r = 4
r = lead(ws, r, "1　補正の式", span=6)
r = header(ws, r, ["", "記号", "内容", "値", "", "備考"])
for i, (k, v, val, b) in enumerate([
    ("g'ₛ", "サービスsに用いる年率", "―", "補正後の伸び"),
    ("gₙ", "認定者数の伸び（年率）",
     "令和9年度 %+.3f％／令和10年度 %+.3f％／令和11年度 %+.3f％"
     % tuple(g * 100 for g in GN), "将来推計 第1段階のシナリオ②"),
    ("gₛ", "サービスsの実績の伸び（年率）", "▲10.7％〜＋28.8％",
     "令和2年度→令和7年度の%d年" % WIN),
    ("φ", "趨勢を反映する割合", "%.1f" % PHI, "バックテストで選んだ"),
    ("clip", "趨勢の差の上限", "±%d％" % (CAP * 100), "同上"),
], start=1):
    r = body(ws, r, [i, k, v, val, "", b], height=26)
r = note(ws, r,
         "　g'ₛ ＝ gₙ ＋ φ × clip(gₛ － gₙ, ±%d％)\n"
         "　見込量ₛ(年) ＝ 令和7年度の実績ₛ × (1 ＋ g'ₛ)^経過年数\n"
         "　給付費ₛ(年) ＝ 見込量ₛ(年) × 令和7年度の1人あたり給付費ₛ\n"
         "φ＝0 とすると一律の伸び（採用パターンP3）に一致します。"
         % (CAP * 100), span=6, height=62)
r += 1

r = lead(ws, r, "2　バックテストの組合せ", span=6)
r = header(ws, r, ["", "起点", "予測年数", "対象年度", "", "内容"])
for i, (o, h) in enumerate(COMB, start=1):
    r = body(ws, r, [i, o, h, YS[IDX[o] + h], "",
                     "%s の実績から%d年先を予測し、%s の実績と比べる"
                     % (o, h, YS[IDX[o] + h])], height=22)
r += 1

r = lead(ws, r, "3　設定ごとの予測誤差（%d通りの平均）" % len(COMB), span=6)
r = header(ws, r, ["順位", "φ", "趨勢の上限", "窓（年）",
                   "サービス別の加重絶対誤差", "総給付費の絶対誤差"])
for i, ((phi, cap, win), (mw, mt)) in enumerate(BT_RANK, start=1):
    if i > 8 and phi != 0.0:
        continue
    f = OK_G if i == 1 else (NG_O if phi == 0.0 else None)
    r = body(ws, r, [i, phi, "なし" if cap is None else "±%d％" % (cap * 100),
                     win, "%.2f％" % (mw * 100), "%.2f％" % (mt * 100)],
             fills={2: f} if f else None, height=20)
r = note(ws, r,
         "注1）上位8設定と、一律（φ＝0）の全設定を掲げています。"
         "全%d設定のうち一律は%d位から%d位です。\n"
         "注2）加重絶対誤差はサービスごとの誤差を給付費で加重したものです。\n"
         "注3）最良の設定は φ＝%.1f・上限±%d％・窓%d年で、"
         "総給付費の誤差は一律の%.2f％に対し%.2f％です。"
         % (len(BT),
            min(i for i, (k, v) in enumerate(BT_RANK, 1) if k[0] == 0.0),
            max(i for i, (k, v) in enumerate(BT_RANK, 1) if k[0] == 0.0),
            BEST[0], (BEST[1] or 0) * 100, BEST[2],
            UNIFORM[1][1] * 100, BT_RANK[0][1][1] * 100), span=6, height=54)

# ============================================================ 04
ws = sheet("04_補正した見込量", "補正した見込量（1月当たり利用者数）",
           "採用パターンP3（一律）と、対計画比の偏りを補正した場合（φ＝%.1f）を"
           "サービス別に並べています。単価は令和7年度の実績で固定しています。" % PHI,
           [30, 12, 12, 12, 12, 12, 12, 12], freeze="A6")
r = 4
r = header(ws, r, ["サービス", "R7実績", "R9 一律", "R9 補正", "R11 一律",
                   "R11 補正", "R11の差", "趨勢の差"])
lst = sorted(SVC, key=lambda l: -(D1[l]["R7"]["実績値"] or 0))
n_up = n_dn = 0
for l in lst:
    n = D1[l]["R7"]["実績値"]
    if not n:
        continue
    a9, b9 = mikomi(l, 0, 0.0) / 12, mikomi(l, 0, PHI) / 12
    a11, b11 = mikomi(l, 2, 0.0) / 12, mikomi(l, 2, PHI) / 12
    d = b11 - a11
    if d > 0.05:
        n_up += 1
    elif d < -0.05:
        n_dn += 1
    r = body(ws, r, [short(l)[:28], round(n / 12, 1), round(a9, 1), round(b9, 1),
                     round(a11, 1), round(b11, 1), round(d, 1),
                     "%+.2f％" % (DRIFT[l] * 100)],
             fills={7: (IN_Y if d > 0.05 else (GRAY if d < -0.05 else None))},
             height=22)
r = note(ws, r,
         "注1）補正により令和11年度の見込量が増えるサービスが%d件、"
         "減るサービスが%d件です。\n"
         "注2）在宅系が増え、施設・居住系が減る形になります。"
         "施設と在宅では1人あたり給付費が大きく異なるため、"
         "この構成の変化が給付費に効きます。\n"
         "注3）定期巡回・随時対応型訪問介護看護は補正では増えません。"
         "区域内に事業所がないことは趨勢では表せず、施策の問題です。"
         % (n_up, n_dn), span=8, height=48)

# ============================================================ 05
ws = sheet("05_給付費と保険料", "給付費と保険料への影響",
           "補正の置き方ごとに、3か年の給付費と算定上の月額基準額を示します。"
           "施設・居住系を減らすかどうかで符号が変わります。",
           [3, 34, 13, 13, 13, 13, 12, 30], freeze="A6")
r = 4
r = lead(ws, r, "1　置き方ごとの給付費と月額", span=8)
r = header(ws, r, ["", "置き方", "令和9年度", "令和10年度", "令和11年度",
                   "3か年計", "月額", "内容"])
for i, (lab, phi, fl, k, m) in enumerate(CASE_RES, start=1):
    memo = ("採用パターンP3" if i == 1 else
            ("バックテストで選んだ設定" if i == 2 else
             ("施設・居住系の定員の縮小を将来へ延長しない" if fl and phi == PHI else
              ("趨勢をそのまま延長する（参考）" if not fl else
               "同・定員の縮小を延長しない（参考）"))))
    r = body(ws, r, [i, lab, round(k[0], 1), round(k[1], 1), round(k[2], 1),
                     round(sum(k), 1), "%.0f円" % m, memo],
             fills={7: (OK_G if i == 2 else (IN_Y if i == 3 else None))},
             height=26)
r = body(ws, r, ["", "①からの差", "", "", "", "", "", ""],
         fills={i: MID_B for i in range(1, 9)}, height=18, bold=True)
for i, (lab, phi, fl, k, m) in enumerate(CASE_RES, start=1):
    if i == 1:
        continue
    r = body(ws, r, ["", "　" + lab, round(k[0] - CASE_RES[0][3][0], 1),
                     round(k[1] - CASE_RES[0][3][1], 1),
                     round(k[2] - CASE_RES[0][3][2], 1),
                     round(sum(k) - sum(CASE_RES[0][3]), 1),
                     "%+.0f円" % (m - M0), ""], height=20)
r = note(ws, r,
         "注1）月額は算定上の月額基準額です。条例で定める基準額ではありません。\n"
         "注2）②と③の差は、施設・居住系の趨勢を延長するかどうかだけです。"
         "実績の減少には定員の縮小が含まれます。"
         "介護老人福祉施設の定員は平成29年度234人から令和2年度以降160人へ"
         "31.6％減り、認知症対応型共同生活介護も15.4％減っています"
         "（計画素案 第5章 基本目標3）。"
         "定員が第10期に維持されるのであれば、"
         "趨勢の延長は見込量を過小にします。\n"
         "注3）補正の効きの幅は%+.0f円から%+.0f円です。"
         % (min(x[4] - M0 for x in CASE_RES[1:]),
            max(x[4] - M0 for x in CASE_RES[1:])), span=8, height=76)
r += 1

r = lead(ws, r, "2　施設・居住系の定員と実績", span=8)
r = header(ws, r, ["サービス", "定員", "R7実績（人／月）", "率", "出所", "", "", ""])
TEIIN = [("介護老人福祉施設", "施設サービス 介護老人福祉施設", 160,
          "北海道 特別養護老人ホーム名簿"),
         ("地域密着型介護老人福祉施設入所者生活介護",
          "施設サービス 地域密着型介護老人福祉施設入所者生活介護", 62, "同上"),
         ("認知症対応型共同生活介護", "居住系サービス 認知症対応型共同生活介護", 99,
          "居所変更実態調査・公表システム（99人＋［要確認］）"),
         ("特定施設入居者生活介護", "居住系サービス 特定施設入居者生活介護", 156,
          "北海道の指定事業所一覧")]
for nm2, key, t, src in TEIIN:
    n = D1[key]["R7"]["実績値"] / 12
    r = body(ws, r, [nm2[:28], t, round(n, 1), "%.1f％" % (n / t * 100), src,
                     "", "", ""], height=22)
r = note(ws, r,
         "注）特定施設入居者生活介護の率が低いのは、"
         "入居者に要支援の方が含まれ介護予防特定施設入居者生活介護として"
         "計上されるためとみられます。［要確認］\n"
         "認知症対応型共同生活介護の定員は99人＋［要確認］です"
         "（東川町の1事業所が未確定）。", span=8, height=36)

# ============================================================ 06
ws = sheet("06_残る外れとその扱い", "補正で解消しない外れとその扱い",
           "補正の前後で、擬似的な対計画比（実績÷予測）がどう変わるかを見ました。"
           "令和5年度の実績から2年先（令和7年度）を予測し、実績と比べています。",
           [30, 12, 12, 12, 14, 40], freeze="A6")
r = 4
r = lead(ws, r, "1　擬似的な対計画比の分布", span=6)
r = header(ws, r, ["指標", "一律（φ＝0）", "補正（φ＝%.1f）" % PHI, "改善",
                   "", "内容"])
stat = {}
for phi in (0.0, PHI):
    rr = [x[1] for x in GIJI[phi]]
    wq = [x[2] for x in GIJI[phi]]
    stat[phi] = dict(med=st.median(rr), sd=st.pstdev(rr), mn=min(rr), mx=max(rr),
                     w=sum(a * b for a, b in zip(rr, wq)) / sum(wq),
                     n80=sum(1 for x in rr if x < 0.8),
                     n120=sum(1 for x in rr if x > 1.2))
for k, lab, fmt in [("sd", "標準偏差", "%.3f"), ("mn", "最小", "%.2f"),
                    ("mx", "最大", "%.2f"), ("w", "給付費で加重", "%.3f"),
                    ("n80", "80％未満の件数", "%d"),
                    ("n120", "120％超の件数", "%d")]:
    a, b = stat[0.0][k], stat[PHI][k]
    if k in ("n80", "n120"):
        imp = "%+d件" % (b - a)
    elif k == "w":
        imp = "%+.3f" % (abs(b - 1) - abs(a - 1))
    else:
        imp = "%+.3f" % (b - a)
    r = body(ws, r, [lab, fmt % a, fmt % b, imp, "",
                     "小さいほど予測が当たっている" if k in ("sd",) else ""],
             height=20)
r = note(ws, r,
         "注）補正により総給付費の誤差は小さくなりますが、"
         "サービス別のばらつきは標準偏差%.3fから%.3fへ%.3fしか縮みません。"
         "補正は総額を当てる方法であって、"
         "サービス別の外れを直す方法ではありません。"
         % (stat[0.0]["sd"], stat[PHI]["sd"],
            abs(stat[PHI]["sd"] - stat[0.0]["sd"])), span=6, height=36)
r += 1

r = lead(ws, r, "2　補正しても外れるサービス", span=6)
r = header(ws, r, ["サービス", "実績", "一律の比", "補正後の比", "外れの向き",
                   "考えられる理由と扱い"])
gm = {l: (a, b) for (l, a, _), (_, b, _) in zip(GIJI[0.0], GIJI[PHI])}
RIYU = {
    "居宅療養管理指導": "8年間で利用者数が4.2倍（555人→2,319人）。"
    "医師・歯科医師・薬剤師等による訪問で、複数の職種から受けられる。"
    "趨勢の上限±%d％では追いつかない。上限を外すか、"
    "別に伸びを置くかの判断を要する" % (CAP * 100),
    "地域密着型通所介護": "令和5年度1,321人から令和6年度950人へ1年で28％減。"
    "趨勢ではなく供給側の変化とみられる。事業所の休廃止の有無の確認を要する",
    "短期入所療養介護（老健）": "対計画比が8年平均84.3％で安定して低い。"
    "老健の定員240人に対し短期入所の枠がどう設定されているかの確認を要する",
    "訪問看護": "8年間で1,035人から1,622人へ57％増。"
    "事業所数は7で人口10万対25.2と全国の1.71倍。供給が需要を作っている可能性",
    "訪問リハビリテーション": "8年間で435人から803人へ85％増。"
    "理学療法士が全国の1.67倍という供給の厚さと符合する",
    "特定福祉用具販売": "件数が年10件程度と少なく、年による振れが大きい",
    "住宅改修": "同上。件数が年8〜10件程度",
    "介護医療院": "区域内に事業所がなく、区域外事業所の利用による。件数が少ない",
}
n_hazure = 0
for l, a, w in GIJI[0.0]:
    b = gm[l][1]
    if 0.8 <= a <= 1.2 and 0.8 <= b <= 1.2:
        continue
    n_hazure += 1
    k = short(l)
    r = body(ws, r, [k[:28], D1[l]["R7"]["実績値"], "%.1f％" % (a * 100),
                     "%.1f％" % (b * 100),
                     "下回る" if b < 1 else "上回る",
                     RIYU.get(k, "［要確認］")],
             fills={5: (NG_O if b < 1 else IN_Y)}, height=40)
r = note(ws, r,
         "注1）外れは%d件です。いずれも補正の前後で向きが変わりません。\n"
         "注2）これらは趨勢の問題ではなく、供給又は制度の問題です。"
         "機械的な補正ではなく、サービスごとの判断を要します。\n"
         "注3）施策反映は、この判断を経たうえで載せる段階です。"
         % n_hazure, span=6, height=44)

# ============================================================ 07 自己点検
_z = [ZENTAI[y] for y in YS]
chk(1, "全体の対計画比が8年度すべてで100％未満であること",
    "H30〜R7", "最大%.2f％・最小%.2f％" % (max(_z) * 100, min(_z) * 100),
    all(x < 1 for x in _z))

chk(2, "サービス別の伸びが認定者数の伸びと大きく異なること",
    "認定者数 年率%+.3f％" % (g_nin_all * 100),
    "サービス別 %+.2f％〜%+.2f％"
    % (min(x[3] for x in tr) * 100, max(x[3] for x in tr) * 100),
    max(x[3] for x in tr) - min(x[3] for x in tr) > 0.10)

chk(3, "バックテストで一律（φ＝0）が最良でないこと",
    "全%d設定" % len(BT),
    "一律は%d位・最良はφ＝%.1f"
    % (min(i for i, (k, v) in enumerate(BT_RANK, 1) if k[0] == 0.0), BEST[0]),
    BEST[0] != 0.0)

chk(4, "補正が総給付費の予測誤差を小さくすること",
    "一律 %.2f％" % (UNIFORM[1][1] * 100),
    "最良 %.2f％" % (BT_RANK[0][1][1] * 100),
    BT_RANK[0][1][1] < UNIFORM[1][1])

chk(5, "φ＝0の給付費が採用パターンP3と一致すること",
    "本表 %.1f百万円" % sum(CASE_RES[0][3]), "サービス見込量の算定 9,020.2百万円",
    abs(sum(CASE_RES[0][3]) - 9020.2) < 0.5)

chk(6, "φ＝0の月額が採用パターンP3と一致すること",
    "本表 %.1f円" % M0, "サービス見込量の算定 6,428円", abs(M0 - 6428) < 1.0)

chk(7, "趨勢の差の上限が働いていること",
    "上限±%d％" % (CAP * 100), "切り詰めが働いたサービス%d件" % n_cap, n_cap > 0)

chk(8, "施設・居住系を減らすかどうかで補正の符号が変わること",
    "②（減らす）%+.0f円" % (CASE_RES[1][4] - M0),
    "③（減らさない）%+.0f円" % (CASE_RES[2][4] - M0),
    (CASE_RES[1][4] - M0) * (CASE_RES[2][4] - M0) < 0)

chk(9, "補正でサービス別のばらつきが解消しないことを示していること",
    "一律 標準偏差%.3f" % stat[0.0]["sd"],
    "補正 標準偏差%.3f（外れ%d件が残る）" % (stat[PHI]["sd"], n_hazure),
    n_hazure > 0 and stat[PHI]["sd"] > 0.15)

chk(10, "定期巡回が補正では増えないこと（趨勢では供給を表せないこと）",
    "R7実績 %.1f人／月"
    % (D1["在宅サービス 定期巡回・随時対応型訪問介護看護"]["R7"]["実績値"] / 12),
    "補正後R11 %.1f人／月"
    % (mikomi("在宅サービス 定期巡回・随時対応型訪問介護看護", 2, PHI) / 12),
    mikomi("在宅サービス 定期巡回・随時対応型訪問介護看護", 2, PHI) / 12 < 1.0)

NG_WORDS = ["に由来する", "と整合する", "1件も", "有意差がないため", "全国トップ級"]
ws = sheet("07_自己点検", "自己点検（エラーチェック）",
           "実測・バックテスト・試算の内的整合と、先行成果品との接続を点検した結果です。"
           "1件でも不適合があると、本表を作るスクリプトは終了コード1で終わります。",
           [5, 42, 34, 34, 12], freeze="A6")
r = 4
r = header(ws, r, ["No", "点検の内容", "式・対象", "結果", "判定"])
for no, naiyo, shiki, kekka, han in CHECKS:
    r = body(ws, r, [no, naiyo, shiki, kekka, han],
             fills={5: OK_G if han == "適合" else NG_O}, height=32)
_ng = [c for c in CHECKS if c[4] != "適合"]
r = body(ws, r, ["", "計", "%d件" % len(CHECKS),
                 "適合%d件・不適合%d件" % (len(CHECKS) - len(_ng), len(_ng)),
                 "適合" if not _ng else "不適合"],
         fills={i: (MID_B if not _ng else NG_O) for i in range(1, 6)},
         height=24, bold=True)

# ============================================================ 08
ws = sheet("08_確認事項", "確認事項",
           "対計画比の偏りの補正について、ご判断又はご確認をお願いする事項です。"
           "業務工程管理表 03_確認事項一覧にも登録します。",
           [5, 16, 46, 40, 14], freeze="A6")
r = 4
r = header(ws, r, ["", "区分", "確認事項", "理由・背景", "希望時期"])
KAKUNIN = [
    ("ご判断",
     "見込量の伸びを一律（認定者数の伸び）から、"
     "サービス別の趨勢を半分反映する置き方（φ＝%.1f）に改めてよいか" % PHI,
     "第7期から第9期までの8年度の対計画比は"
     "サービス別に大きく開いており、全体は8年度すべてで100％未満である。"
     "実績の伸びはサービス別に年率▲10.7％から＋28.8％まで開いており、"
     "認定者数の伸び（年率＋0.48％）とは別物である。"
     "バックテストでは、一律は%d設定中の下位で、"
     "総給付費の予測誤差は一律1.68％に対し補正1.13％である。",
     "令和8年9月26日まで"),
    ("ご判断",
     "施設・居住系の実績の減少を将来へ延長するか",
     "介護老人福祉施設の定員は平成29年度234人から令和2年度以降160人へ31.6％、"
     "認知症対応型共同生活介護は15.4％減っている。"
     "実績の減少には定員の縮小が含まれるため、"
     "趨勢をそのまま延長すると「さらに定員が縮小する」ことを前提にしてしまう。"
     "延長すると月額%+.0f円、延長しないと%+.0f円で符号が変わる。"
     % (CASE_RES[1][4] - M0, CASE_RES[2][4] - M0),
     "令和8年9月26日まで"),
    ("ご確認",
     "地域密着型通所介護が令和5年度1,321人から令和6年度950人へ"
     "1年で28％減った理由",
     "趨勢ではなく供給側の変化とみられる。"
     "事業所の休廃止・定員の縮小・利用者の他サービスへの移動のいずれかを"
     "ご確認いただきたい。区域内の事業所は4である。",
     "令和8年10月"),
    ("ご確認",
     "居宅療養管理指導が8年間で4.2倍（555人→2,319人）になった理由",
     "対計画比は8年平均146.6％で、補正しても擬似的な対計画比が188％と外れる。"
     "趨勢の上限±%d％では追いつかない。"
     "医師・歯科医師・薬剤師等のどの職種が増えているかをご確認いただきたい。"
     "上限を外すか、別に伸びを置くかの判断を要する。" % (CAP * 100),
     "令和8年10月"),
    ("ご確認",
     "短期入所生活介護・短期入所療養介護の対計画比が"
     "8年平均80.1％・84.3％と安定して低い理由",
     "短期入所生活介護は令和2年度以降も減り続けている（541人→318人）。"
     "特別養護老人ホームのショート専用定員は15人、"
     "地域密着型のショート専用定員は6人である。"
     "空床利用の運用と、需要があって受けられていないのかどうかを"
     "ご確認いただきたい。",
     "令和8年10月"),
    ("ご判断",
     "補正で解消しない外れ%d件を、"
     "供給実現性の段階でサービスごとに判断することでよいか" % n_hazure,
     "補正は総給付費の予測誤差を小さくするが、"
     "サービス別のばらつきは標準偏差%.3fから%.3fへしか縮まない。"
     "外れは趨勢の問題ではなく供給又は制度の問題であり、"
     "機械的な補正では直らない。"
     % (stat[0.0]["sd"], stat[PHI]["sd"]),
     "令和8年10月"),
]
for i, (k, a, b, c) in enumerate(KAKUNIN, start=1):
    if i == 1:
        b = b % len(BT)
    f = {2: IN_Y if k == "ご判断" else MID_B}
    r = body(ws, r, [i, k, a, b, c], fills=f, height=66)
r += 1
r = note(ws, r,
         "注）確認事項は業務工程管理表 03_確認事項一覧で一元管理します。"
         "計画素案の本文には確認事項の注記を書きません。", span=5, height=20)

# ============================================================ 出力
os.makedirs(ODIR, exist_ok=True)
wb.save(OUT)
_ngw = []
for _ws in wb.worksheets:
    for _row in _ws.iter_rows(values_only=True):
        for _v in _row:
            if isinstance(_v, str):
                for _w in NG_WORDS:
                    if _w in _v:
                        _ngw.append((_ws.title, _w))
print("書き出しました:", OUT)
for s in wb.sheetnames:
    print("  -", s, wb[s].max_row, "rows")
print("対計画比 8年平均 %.2f％（サービス別 %.1f％〜%.1f％）"
      % (st.mean(ZENTAI.values()) * 100,
         min(x[2] for x in rows) * 100, max(x[2] for x in rows) * 100))
print("最良の設定 φ=%.1f 上限±%d％ 窓%d年（一律は%d設定中%d位）"
      % (BEST[0], (BEST[1] or 0) * 100, BEST[2], len(BT),
         min(i for i, (k, v) in enumerate(BT_RANK, 1) if k[0] == 0.0)))
for lab, phi, fl, k, m in CASE_RES:
    print("  %-34s %8.1f百万円  %6.0f円（%+5.0f円）"
          % (lab, sum(k), m, m - M0))
print("補正で解消しない外れ %d件／自己点検 %d件：適合%d件・不適合%d件"
      % (n_hazure, len(CHECKS), len(CHECKS) - len(_ng), len(_ng)))
if _ngw:
    print("禁止表現:", _ngw)
if _ng or _ngw:
    for c in _ng:
        print("  不適合:", c[0], c[1], c[3])
    sys.exit(1)
print("すべての点検に適合しました。")
