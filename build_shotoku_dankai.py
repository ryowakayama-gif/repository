# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画
所得段階別第1号被保険者数の将来推計（確認事項No.156）.

令和8年9月24日のご指示
  「No.156とNo.157を当方で算定して進めてください」

━━ 何を算定するか ━━

見える化システムの「保険料額の算定」画面は、所得段階別の第1号被保険者数を
**令和9・10・11・12・17・22・27・32年度の8年度**について求める。
当方はこれまで所得段階別加入割合補正後の被保険者数（③）を
**係数で一括して**算定しており、段階別の将来人数を出していなかった。

本表は、**実績の構成比を固定し、将来の第1号被保険者数（案C）に乗じて**
段階別の人数を出す。整数化は最大剰余法により合計を保つ。

━━ 構成比をどの年度で置くかは既存の選択と一体である ★ ━━

段階別の人数を実績の構成比で置くと、そこから求まる係数
（Σ（段階別×乗率）÷Σ段階別）は、その年度の係数そのものになる。
当方の保険料算定は**令和6年度の実績による係数0.991352を基本ケース**とし、
令和7年度（1.015890）を感度としている（将来推計 第3段階 04シート）。

  令和6年度の構成比　係数0.991352　月額の基本ケース（±0円）
  令和7年度の構成比　係数1.015890　月額▲163円
  第9期計画の構成比　係数0.969232　月額＋154円

**本表は令和6年度の構成比を採る。** 当方が採用している係数と同じであり、
見える化システムが本表の人数から算定する③が当方の③と一致するためである。
構成比を令和7年度に改めるかどうかは係数の選択そのものであり、
本表の作り方の問題ではない（04シートの感度）。

━━ 当連合は16段階である ━━

見える化システムの標準は13段階であるが、
**当連合の条例は第9期において16段階**である（第13段階を多段階化）。
入力画面では所得段階の設定を「弾力化」とし16段階で入力する。
第10期の段階数・乗率は政令改正により変わり得る（確認事項No.33）。

━━ 数値の出所 ━━

  実績　　　`data_nenpo.SHOTOKU`（年報の所得段階別被保険者数・各年度末）
  将来人口　`build_mikomiryo_santei.py` の `HIHO`（案C・第1号被保険者数）
  乗率　　　`build_projection3.py` の `JORITSU`（第9期の条例による16段階）
  係数・月額　`build_projection3.py` の `gaku()`

**固定値を書かない。** いずれも `runpy` 又は import で読む。

シート構成
  00_この表について
  01_構成比の置き方
  02_所得段階別第1号被保険者数（8年度）
  03_補正後被保険者数の照合
  04_自己点検

出力
  output/第10期計画_所得段階別第1号被保険者数の将来推計.xlsx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import io
import os
import runpy
import sys
from decimal import Decimal, ROUND_HALF_UP

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import data_nenpo as N
import repo_paths as RP

# 直接実行したときだけ標準出力を UTF-8 に包み直す。
# 他のスクリプトから runpy で読まれるときは差し替えられた出力を壊さない。
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

ODIR = RP.OUTPUT
OUT = os.path.join(ODIR, "第10期計画_所得段階別第1号被保険者数の将来推計.xlsx")
KIJUNBI = "令和8年9月24日"

FONT = "游ゴシック"
NAVY, HEAD = "1F3864", "4472C4"
IN_Y, OK_G, NG_O, MID_B, GRAY = ("FFF2CC", "E2EFDA", "FCE4D6",
                                 "DEEBF7", "F2F2F2")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()
wb.remove(wb.active)

CHECKS = []


def chk(no, naiyo, shiki, kekka, ok):
    CHECKS.append((no, naiyo, shiki, kekka, "適合" if ok else "不適合"))
    return ok


# ==================================================== 前段を読む
def _load(name):
    buf, old = io.StringIO(), sys.stdout
    sys.stdout = buf
    try:
        return runpy.run_path(os.path.join(RP.ROOT, name))
    finally:
        sys.stdout = old


_M = _load("build_mikomiryo_santei.py")
_P = _load("build_projection3.py")

HIHO = _M["HIHO"]                       # 第1号被保険者数（案C）
Y3, Y3L = _M["Y3"], _M["Y3L"]
YLONG, YLONGL = _M["YLONG"], _M["YLONGL"]
BASE_Y = _M["BASE_Y"]
YALL = Y3 + YLONG                       # 令和9・10・11・12・17・22・27・32年度
YALLL = Y3L + YLONGL

JORITSU = _P["JORITSU"]
DANKAI_KEY = _P["DANKAI_KEY"]
DANKAI_NAME = _P["DANKAI_NAME"]
KEISU_R6, KEISU_R7, KEISU_K9 = _P["KEISU_R6"], _P["KEISU_R7"], _P["KEISU_K9"]
gaku, BASE = _P["gaku"], _P["BASE"]
G_DO = _M["G_DO"]                       # 第10期の採用（要介護度別）

NDAN = len(DANKAI_KEY)                  # 16段階
N_R6 = [N.SHOTOKU[k]["R6"] for k in DANKAI_KEY]
N_R7 = [N.SHOTOKU[k]["R7"] for k in DANKAI_KEY]
# 第9期計画の所得段階別被保険者数（3年計）。04シートと同じ値をソースから読む。
PLAN_N = _P["plan_n"] if "plan_n" in _P else None


def keisu_of(ns):
    return sum(a * b for a, b in zip(ns, JORITSU)) / sum(ns)


# 構成比の3案
KOSEI = {
    "令和6年度": [v / sum(N_R6) for v in N_R6],
    "令和7年度": [v / sum(N_R7) for v in N_R7],
}
if PLAN_N:
    KOSEI["第9期計画"] = [v / sum(PLAN_N) for v in PLAN_N]

SAIYO = "令和6年度"            # 当方が採用している係数と同じ構成比


def to_int(vals, total):
    """合計を total（整数）に保ったまま整数化する（最大剰余法）。"""
    base = [int(v) for v in vals]
    need = total - sum(base)
    order = sorted(range(len(vals)), key=lambda i: -(vals[i] - int(vals[i])))
    out = list(base)
    k = 0
    while need > 0:
        out[order[k % len(order)]] += 1
        need -= 1
        k += 1
    while need < 0:
        for i in sorted(range(len(vals)),
                        key=lambda i: (vals[i] - int(vals[i]))):
            if out[i] > 0 and need < 0:
                out[i] -= 1
                need += 1
    return out


def hiho_int(y):
    """その年度の第1号被保険者数（整数）。"""
    return int(Decimal(str(HIHO[y])).quantize(Decimal("1"), ROUND_HALF_UP))


def dankai(y, kosei=None):
    """所得段階別の第1号被保険者数（整数・16段階）。"""
    k = KOSEI[kosei or SAIYO]
    n = hiho_int(y)
    return to_int([n * x for x in k], n)


DANKAI = {y: dankai(y) for y in YALL}
DANKAI_R7J = {y: dankai(y, "令和7年度") for y in YALL}


def yen(v):
    return "{:,}".format(int(round(v)))


# ==================================================== 体裁
def _plain(v):
    """** による強調の指定を落とす（xlsx は Markdown を解釈しない）。"""
    return v.replace("**", "") if isinstance(v, str) else v


def sheet(name, title, subtitle, widths, freeze="A5"):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A2"] = _plain(subtitle)
    ws["A2"].font = Font(name=FONT, size=9)
    ws["A2"].fill = PatternFill("solid", fgColor=GRAY)
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    n = max(len(widths), 6)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n)
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 58
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = freeze
    return ws


def header(ws, row, cols, height=30):
    for i, h in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i, value=_plain(h))
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=HEAD)
        c.alignment = Alignment(wrap_text=True, horizontal="center",
                                vertical="center")
        c.border = BORDER
    ws.row_dimensions[row].height = height
    return row + 1


def body(ws, row, vals, fills=None, height=20, align=None, bold=False,
         fmt=None):
    for i, v in enumerate(vals, start=1):
        c = ws.cell(row=row, column=i, value=_plain(v))
        c.font = Font(name=FONT, size=9, bold=bold)
        c.alignment = Alignment(wrap_text=True, vertical="center",
                                horizontal=(align or {}).get(i, "left"))
        c.border = BORDER
        if fills and fills.get(i):
            c.fill = PatternFill("solid", fgColor=fills[i])
        if fmt and i in fmt:
            c.number_format = fmt[i]
    ws.row_dimensions[row].height = height
    return row + 1


def lead(ws, row, text, span=6, height=20):
    c = ws.cell(row=row, column=1, value=_plain(text))
    c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = height
    return row + 1


def note(ws, row, text, span=6, height=34, fill=GRAY):
    c = ws.cell(row=row, column=1, value=_plain(text))
    c.font = Font(name=FONT, size=9)
    c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = height
    return row + 1


# ============================================================ 00
ws = sheet("00_この表について",
           "所得段階別第1号被保険者数の将来推計（確認事項No.156）",
           "見える化システムの「保険料額の算定」画面は、所得段階別の"
           "第1号被保険者数を令和9・10・11・12・17・22・27・32年度の8年度に"
           "ついて求めます。"
           "当方はこれまで所得段階別加入割合補正後の被保険者数（③）を"
           "係数で一括して算定しており、段階別の将来人数を出していませんでした。"
           "本表は実績の構成比を固定し、将来の第1号被保険者数（案C）に"
           "乗じて段階別の人数を出したものです。"
           "令和8年9月24日のご指示により当方で算定しました。",
           [4, 34, 22, 22, 22, 26], freeze="A4")
r = 4
r = note(ws, r,
         "【要点】\n"
         "・構成比は**令和6年度の実績**で固定しました。"
         "当方の保険料算定が採用している係数（0.991352）と同じであり、"
         "見える化システムが本表の人数から算定する補正後被保険者数が"
         "当方の値と一致するためです。\n"
         "・構成比を令和7年度に改めると係数は1.015890となり、"
         "**算定上の月額が163円下がります**。"
         "これは段階別人数の作り方の問題ではなく、"
         "係数をどの年度で置くかという既存の選択です"
         "（将来推計 第3段階 04シート）。\n"
         "・**当連合の所得段階は16段階です**（第9期の条例。第13段階を多段階化）。"
         "見える化システムの標準は13段階であるため、"
         "入力画面では所得段階の設定を「弾力化」とします。\n"
         "・第10期の段階数・乗率は政令改正により変わり得ます（確認事項No.33）。"
         "本表は第9期の16段階・乗率をそのまま用いた暫定の値です。",
         span=6, height=132, fill=IN_Y)

r += 1
r = lead(ws, r, "1　算定の手順")
r = header(ws, r, ["", "手順", "内容", "出所", "", ""])
for i, (te, naiyo, syu) in enumerate([
    ("① 実績の構成比を求める",
     "令和6年度の段階別被保険者数（計%s人）を合計で除す" % yen(sum(N_R6)),
     "年報（`data_nenpo.SHOTOKU`）"),
    ("② 将来の第1号被保険者数を置く",
     "案C（総人口＝地方創生総合戦略、年齢階級別＝住民基本台帳の実績趨勢）。"
     "令和12年度以降は令和11年度を起点に社人研推計の伸び率で接続",
     "見込量の算定（`HIHO`）"),
    ("③ 乗じて段階別の人数を出す", "②×①。各年度16段階", "―"),
    ("④ 整数にする",
     "最大剰余法により合計を第1号被保険者数（整数）に保つ。"
     "単純に四捨五入すると合計がずれる", "―"),
    ("⑤ 照合する",
     "Σ（段階別×乗率）÷Σ段階別が採用している係数と一致することを確かめる",
     "将来推計 第3段階 04シート"),
], start=1):
    r = body(ws, r, [i, te, naiyo, syu, "", ""], height=34)

r += 1
r = lead(ws, r, "2　結果の要点")
r = header(ws, r, ["", "項目", "令和9年度", "令和11年度", "令和22年度",
                   "令和32年度"])
_KEY_Y = ["2027", "2029", "2040", "2050"]
for i, (nm, f) in enumerate([
    ("第1号被保険者数（人）", lambda y: hiho_int(y)),
    ("第1〜第3段階（公費軽減の対象・人）",
     lambda y: sum(DANKAI[y][:3])),
    ("第9段階以上（人）", lambda y: sum(DANKAI[y][8:])),
    ("補正後被保険者数（人）",
     lambda y: sum(a * b for a, b in zip(DANKAI[y], JORITSU))),
], start=1):
    r = body(ws, r, [i, nm] + [round(f(y), 1) for y in _KEY_Y],
             fmt={j: "#,##0.0" for j in range(3, 7)},
             align={j: "right" for j in range(3, 7)}, height=22)
r = note(ws, r,
         "注）構成比を固定しているため、段階別の人数は"
         "第1号被保険者数の増減にそのまま比例します。"
         "所得分布そのものの変化（令和6年度から令和7年度にかけて"
         "第9段階以上が490人から597人へ増えたような動き）は"
         "織り込んでいません。"
         "2年分の実績では趨勢と言えないためです。",
         span=6, height=48)

# ============================================================ 01
ws = sheet("01_構成比の置き方", "構成比をどの年度で置くか",
           "段階別の人数を実績の構成比で置くと、そこから求まる係数は"
           "その年度の係数そのものになります。"
           "構成比の選択は係数の選択と一体です。",
           [4, 18, 12, 14, 14, 44], freeze="A5")
r = 4
r = lead(ws, r, "1　3案の比較")
r = header(ws, r, ["", "構成比", "係数", "補正後（3か年計・人）",
                   "月額への効き", "内容"])
_NAI = {
    "令和6年度": "当方の保険料算定が採用している係数と同じ。"
                 "見える化システムが本表の人数から算定する補正後被保険者数が"
                 "当方の値と一致する",
    "令和7年度": "直近の実績。第9段階以上が令和6年度490人から597人へ増え、"
                 "係数が上がっている。見込量の基準年度（令和7年度）とはそろうが、"
                 "2年分では趨勢と言えない",
    "第9期計画": "第9期計画が前提としていた分布。実績はこれを上回っている",
}
_KS = {"令和6年度": KEISU_R6, "令和7年度": KEISU_R7, "第9期計画": KEISU_K9}
for i, nm in enumerate([k for k in ("令和6年度", "令和7年度", "第9期計画")
                        if k in _KS], start=1):
    ks = _KS[nm]
    g = gaku(keisu=ks)
    r = body(ws, r, [i, nm + "の構成比", round(ks, 6),
                     round(sum(_P["HIHOKEN"]) * ks, 1),
                     "―" if nm == SAIYO
                     else "%+d円" % round(g["月額"] - BASE["月額"]),
                     _NAI.get(nm, "")],
             fills={2: (OK_G if nm == SAIYO else None)},
             fmt={3: "0.000000", 4: "#,##0.0"},
             align={3: "right", 4: "right", 5: "center"}, height=46)
r = note(ws, r,
         "注1）**本表は令和6年度の構成比を採ります。**"
         "当方の保険料算定の基本ケースと同じ係数であり、"
         "段階別の人数を入れたときにシステムが出す補正後被保険者数が"
         "当方の値と一致します。\n"
         "注2）月額への効きは、給付費・地域支援事業費など他の前提を"
         "すべて同じにしたうえで係数だけを置き換えた場合の差です"
         "（基本ケース%.0f円に対する差）。\n"
         "注3）構成比を令和7年度に改めるかどうかは、"
         "本表の作り方ではなく係数の選択そのものです。"
         "将来推計 第3段階 04シートの感度として既に掲げています。"
         % BASE["月額"],
         span=6, height=76)

r += 1
r = lead(ws, r, "2　実績の構成比（16段階）")
r = header(ws, r, ["", "段階", "乗率", "令和6年度（人）", "令和7年度（人）",
                   "構成比の差（ポイント）"])
for i in range(NDAN):
    p6 = N_R6[i] / sum(N_R6) * 100
    p7 = N_R7[i] / sum(N_R7) * 100
    r = body(ws, r, [i + 1, DANKAI_NAME[i], JORITSU[i], N_R6[i], N_R7[i],
                     "%+.2f" % (p7 - p6)],
             fills={2: (IN_Y if i < 3 else None)},
             fmt={4: "#,##0", 5: "#,##0"},
             align={3: "center", 4: "right", 5: "right", 6: "right"},
             height=18)
r = body(ws, r, ["", "計", "―", sum(N_R6), sum(N_R7), "―"],
         fills={i: MID_B for i in range(1, 7)}, bold=True,
         fmt={4: "#,##0", 5: "#,##0"},
         align={3: "center", 4: "right", 5: "right", 6: "center"}, height=20)
r = note(ws, r,
         "注1）第1段階から第3段階は公費軽減の対象です。"
         "乗率は**公費軽減前**の値であり、"
         "補正後被保険者数はこの乗率で算定します"
         "（公費軽減分は国・都道府県・市町村の公費で補填されるため）。\n"
         "注2）**当連合は16段階です**（第9期の条例。第13段階を多段階化）。"
         "見える化システムの標準は13段階であるため、"
         "入力画面では所得段階の設定を「弾力化」とします。\n"
         "注3）第10期の段階数・乗率は政令改正により変わり得ます"
         "（確認事項No.33）。",
         span=6, height=64)

# ============================================================ 02
ws = sheet("02_段階別の将来推計",
           "所得段階別第1号被保険者数の将来推計（採用＝%sの構成比）" % SAIYO,
           "見える化システムの「保険料額の算定」画面に入れる値です。"
           "各年度の合計は第1号被保険者数（案C）と一致します。"
           "整数化は最大剰余法により合計を保っています。"
           "**必要資料が未受領であることによる暫定値です。**",
           [4, 18, 10] + [11] * len(YALL) + [12], freeze="D5")
r = 4
_ALIGN_DAN = {3: "center"}
_ALIGN_DAN.update({j: "right" for j in range(4, 5 + len(YALL))})


def _fills_dan(i):
    f = {j: IN_Y for j in range(4, 4 + len(YALL))}
    if i < 3:
        f[2] = IN_Y
    return f


r = lead(ws, r, "所得段階別第1号被保険者数（人）", span=4 + len(YALL))
r = header(ws, r, ["", "段階", "乗率"] + YALLL + ["R7実績"])
for i in range(NDAN):
    r = body(ws, r, [i + 1, DANKAI_NAME[i], JORITSU[i]]
             + [DANKAI[y][i] for y in YALL] + [N_R7[i]],
             fills=_fills_dan(i),
             fmt={j: "#,##0" for j in range(4, 5 + len(YALL))},
             align=_ALIGN_DAN, height=18)
r = body(ws, r, ["", "計", "―"] + [sum(DANKAI[y]) for y in YALL]
         + [sum(N_R7)],
         fills={i: MID_B for i in range(1, 5 + len(YALL))}, bold=True,
         fmt={j: "#,##0" for j in range(4, 5 + len(YALL))},
         align=_ALIGN_DAN, height=20)
r = body(ws, r, ["", "第1号被保険者数（案C）", "―"]
         + [hiho_int(y) for y in YALL] + [sum(N_R7)],
         fmt={j: "#,##0" for j in range(4, 5 + len(YALL))},
         align=_ALIGN_DAN, height=20)
r = body(ws, r, ["", "補正後被保険者数", "―"]
         + [round(sum(a * b for a, b in zip(DANKAI[y], JORITSU)), 1)
            for y in YALL] + ["―"],
         fills={i: OK_G for i in range(1, 5 + len(YALL))}, bold=True,
         fmt={j: "#,##0.0" for j in range(4, 4 + len(YALL))},
         align=_ALIGN_DAN, height=20)
r = note(ws, r,
         "注1）**必要資料が未受領であることによる暫定値です。確定値ではありません。**\n"
         "注2）第1号被保険者数は案Cによります"
         "（総人口＝地方創生総合戦略、年齢階級別＝住民基本台帳の実績趨勢）。"
         "令和12年度以降は令和11年度を起点に社人研推計の伸び率で接続しています。"
         "見える化システムの初期値は社人研推計であるため、"
         "総人口の設定を案Cに改めないと食い違います（確認事項No.127）。\n"
         "注3）構成比は%sの実績で固定しています。"
         "所得分布そのものの変化は織り込んでいません。\n"
         "注4）各年度の「計」は第1号被保険者数と一致します（自己点検2）。"
         % SAIYO,
         span=4 + len(YALL), height=76)

# ============================================================ 03
ws = sheet("03_補正後被保険者数の照合",
           "補正後被保険者数（③）と保険料の照合",
           "本表の段階別人数から求まる補正後被保険者数が、"
           "当方の保険料算定の③と一致することを確かめます。",
           [4, 34, 18, 18, 18, 30], freeze="A5")
r = 4
r = lead(ws, r, "1　第10期（令和9〜11年度）の照合")
r = header(ws, r, ["", "項目", "本表から求めた値", "当方の保険料算定", "差",
                   "備考"])
_H3 = sum(hiho_int(y) for y in Y3)
_HOSEI3 = sum(sum(a * b for a, b in zip(DANKAI[y], JORITSU)) for y in Y3)
_KS3 = _HOSEI3 / _H3
for i, (nm, a, b, bk) in enumerate([
    ("第1号被保険者数（3か年計・人）", _H3, sum(_P["HIHOKEN"]),
     "整数化による差"),
    ("補正後被保険者数（③・人）", _HOSEI3, G_DO["③"],
     "Σ（段階別×乗率）"),
    ("係数（③÷第1号）", _KS3, KEISU_R6, "採用している係数"),
], start=1):
    r = body(ws, r, [i, nm, round(a, 4), round(b, 4), round(a - b, 4), bk],
             fmt={3: "#,##0.0000", 4: "#,##0.0000", 5: "#,##0.0000"},
             align={3: "right", 4: "right", 5: "right"}, height=22)
r = body(ws, r, ["", "算定上の月額基準額（円）",
                 round(gaku(keisu=_KS3, hiho=_H3,
                            kyufu_oku=sum(_M["KYUFU3"]))["月額"], 2),
                 round(G_DO["月額"], 2),
                 round(gaku(keisu=_KS3, hiho=_H3,
                            kyufu_oku=sum(_M["KYUFU3"]))["月額"]
                       - G_DO["月額"], 2),
                 "本表の人数で算定し直した場合"],
         fills={i: OK_G for i in range(1, 7)}, bold=True,
         fmt={3: "#,##0.00", 4: "#,##0.00", 5: "#,##0.00"},
         align={3: "right", 4: "right", 5: "right"}, height=24)
r = note(ws, r,
         "注1）第1号被保険者数は、当方の保険料算定では小数のまま"
         "（%.1f人）用いており、本表は整数化しています。"
         "整数化による差は%.1f人で、月額への影響は1円未満です。\n"
         "注2）係数が一致するのは、構成比を令和6年度で固定しており、"
         "係数がその構成比から定まるためです。"
         "本表の人数を見える化システムに入れれば、"
         "システムが出す補正後被保険者数は当方の値と一致します。"
         % (sum(_P["HIHOKEN"]), _H3 - sum(_P["HIHOKEN"])),
         span=6, height=58)

r += 1
r = lead(ws, r, "2　中長期（参考）")
r = header(ws, r, ["", "年度", "第1号被保険者数", "補正後被保険者数", "係数",
                   "備考"])
for i, (y, lb) in enumerate(zip(YALL, YALLL), start=1):
    h = sum(a * b for a, b in zip(DANKAI[y], JORITSU))
    r = body(ws, r, [i, lb, hiho_int(y), round(h, 1),
                     round(h / hiho_int(y), 6),
                     "第10期" if y in Y3 else "中長期（確認事項No.157）"],
             fmt={3: "#,##0", 4: "#,##0.0", 5: "0.000000"},
             align={3: "right", 4: "right", 5: "right"}, height=20)
r = note(ws, r,
         "注）係数が全年度で同じになるのは、構成比を固定しているためです。"
         "整数化による端数で小数第4位以下がわずかに動きます。",
         span=6, height=30)

# ============================================================ 04
ws = sheet("04_自己点検", "自己点検",
           "本表の内的整合を、出典から独立に計算して確かめた記録です。"
           "1件でも不適合があると本表を作るスクリプトは終了コード1で終わります。",
           [5, 44, 34, 34, 10], freeze="A5")

chk(1, "所得段階が16段階であること",
    "len(DANKAI_KEY) == len(JORITSU)",
    "%d段階" % NDAN, NDAN == 16 and len(JORITSU) == 16)

_sum_ok = all(sum(DANKAI[y]) == hiho_int(y) for y in YALL)
chk(2, "各年度の段階別の合計が第1号被保険者数と一致すること",
    "Σ段階別 == round(HIHO[y])",
    "全%d年度で一致" % len(YALL) if _sum_ok else "不一致あり", _sum_ok)

_neg = [(y, i) for y in YALL for i, v in enumerate(DANKAI[y]) if v < 0]
chk(3, "負の人数がないこと", "DANKAI[y][i] >= 0",
    "該当%d件" % len(_neg), not _neg)

chk(4, "採用した構成比の係数が保険料算定の係数と一致すること",
    "Σ（R6実績×乗率）÷Σ R6実績 == KEISU_R6",
    "%.6f 対 %.6f" % (keisu_of(N_R6), KEISU_R6),
    abs(keisu_of(N_R6) - KEISU_R6) < 1e-9)

chk(5, "令和7年度の構成比の係数が感度の係数と一致すること",
    "Σ（R7実績×乗率）÷Σ R7実績 == KEISU_R7",
    "%.6f 対 %.6f" % (keisu_of(N_R7), KEISU_R7),
    abs(keisu_of(N_R7) - KEISU_R7) < 1e-9)

chk(6, "本表から求めた第10期の係数が採用値と小数第4位まで一致すること",
    "Σ（段階別×乗率）÷Σ段階別（3か年計）",
    "%.6f 対 %.6f" % (_KS3, KEISU_R6), abs(_KS3 - KEISU_R6) < 1e-4)

_g_here = gaku(keisu=_KS3, hiho=_H3, kyufu_oku=sum(_M["KYUFU3"]))["月額"]
chk(7, "本表の人数で算定し直した月額が当方の値と1円未満で一致すること",
    "gaku(keisu=本表の係数, hiho=本表の合計)",
    "%.2f円 対 %.2f円（差%.2f円）" % (_g_here, G_DO["月額"],
                                      _g_here - G_DO["月額"]),
    abs(_g_here - G_DO["月額"]) < 1.0)

chk(8, "対象の年度が見える化システムの求める8年度であること",
    "Y3 ＋ YLONG", "／".join(YALLL), len(YALL) == 8)

_r7_keisu_gap = gaku(keisu=KEISU_R7)["月額"] - BASE["月額"]
chk(9, "令和7年度の構成比に改めた場合の月額の差を示していること",
    "gaku(keisu=KEISU_R7) − 基本ケース",
    "%+.0f円" % _r7_keisu_gap, abs(_r7_keisu_gap) > 100)

_pub = [sum(DANKAI[y][:3]) for y in YALL]
chk(10, "公費軽減の対象（第1〜第3段階）が全年度で正であること",
     "Σ DANKAI[y][:3]",
     "令和9年度%d人〜令和32年度%d人" % (_pub[0], _pub[-1]),
     all(v > 0 for v in _pub))

_int_ok = all(isinstance(v, int) for y in YALL for v in DANKAI[y])
chk(11, "すべての欄が整数であること", "isinstance(v, int)",
     "%d年度×%d段階" % (len(YALL), NDAN), _int_ok)

NG_WORDS = ["に由来する", "と整合する", "1件も", "有意差がないため関係がない",
            "全国トップ級"]
_ast = sum(1 for s in wb.worksheets
           for row in s.iter_rows(values_only=True)
           for v in row if isinstance(v, str) and "**" in v)
chk(12, "強調の指定（**）がセルに残っていないこと",
     "xlsx は Markdown を解釈しない", "残り%d件" % _ast, _ast == 0)

r = header(ws, 4, ["No.", "点検した内容", "式・条件", "結果", "判定"])
for c in CHECKS:
    body(ws, r, list(c), fills={5: (OK_G if c[4] == "適合" else NG_O)},
         height=28)
    r += 1
_NG = sum(1 for c in CHECKS if c[4] != "適合")
body(ws, r, ["", "自己点検の結果", "",
             "適合%d件・不適合%d件" % (len(CHECKS) - _NG, _NG),
             "適合" if _NG == 0 else "不適合"],
     fills={5: (OK_G if _NG == 0 else NG_O)}, height=24, bold=True)


# ============================================================ 出力
os.makedirs(ODIR, exist_ok=True)
wb.save(OUT)

_ngw = []
for _ws in wb.worksheets:
    if _ws.title == "04_自己点検":
        continue
    for _row in _ws.iter_rows(values_only=True):
        for _v in _row:
            if isinstance(_v, str):
                for _w in NG_WORDS:
                    if _w in _v:
                        _ngw.append((_ws.title, _w))

_ng = [c for c in CHECKS if c[4] != "適合"]
print("書き出しました:", OUT)
for s in wb.sheetnames:
    print("  -", s, wb[s].max_row, "rows")
print("所得段階 %d段階（弾力化）／年度 %d（%s）"
      % (NDAN, len(YALL), "・".join(YALLL)))
print("採用した構成比 %s（係数 %.6f）" % (SAIYO, KEISU_R6))
print("令和7年度の構成比に改めた場合 係数 %.6f・月額 %+.0f円"
      % (KEISU_R7, _r7_keisu_gap))
print("第10期3か年 第1号 %s人／補正後 %.1f人／係数 %.6f"
      % ("{:,}".format(_H3), _HOSEI3, _KS3))
print("本表の人数による月額 %.2f円（当方 %.2f円・差 %.2f円）"
      % (_g_here, G_DO["月額"], _g_here - G_DO["月額"]))
print("自己点検 %d件：適合%d件・不適合%d件"
      % (len(CHECKS), len(CHECKS) - len(_ng), len(_ng)))
if _ngw:
    print("禁止表現:", _ngw)
for c in _ng:
    print("  不適合:", c[0], c[1], c[3])
if _ng or _ngw:
    sys.exit(1)
print("すべての点検に適合しました。")
