# -*- coding: utf-8 -*-
"""保険者機能強化推進交付金等　市町村分の評価指標と構成3町の評価.

保険者機能強化推進交付金及び介護保険保険者努力支援交付金の
**市町村分**の評価指標について、指標の構造（交付金・目標・指標群・
大項目・明細列・配点）と、構成3町の3か年の得点を1冊にまとめたものである。

出典
  「令和８年度保険者機能強化推進交付金及び介護保険保険者努力支援交付金
    （市町村分）に係る評価指標」（42頁。令和8年9月2日受領）
  「同（市町村分）に係る全国集計結果」令和6年度・令和7年度・令和8年度
    （令和8年8月28日受領、令和8年9月2日に原本を再受領）

評価は市町村単位で行われる。保険者は大雪地区広域連合であるが、
公表資料の広域連合の行に得点はなく、構成3町それぞれの行に記載がある。

年度の対応　令和8年度交付金は令和8年度の評価指標により評価され、
　　　　　　評価の対象は令和7年度（2025年度）に実施した取組である。

シート構成
  00_この表について
  01_評価指標の構造
  02_大項目別の評価（令和8年度）
  03_明細列別の評価（令和8年度）
  04_3か年の推移
  05_全国・北海道との比較
  06_得点していない項目
  07_自己点検

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import data_kofukin as KF                            # noqa: E402
import data_kofukin_detail as KD                     # noqa: E402
import data_kofukin_item as KI                       # noqa: E402
import data_kofukin_zenkoku as Z                     # noqa: E402
import repo_paths as RP                              # noqa: E402

ODIR = RP.ROOT + "/output"
OUT = os.path.join(ODIR, "第10期計画_交付金 市町村分評価指標と3町の評価.xlsx")

TOWNS = ["東川町", "美瑛町", "東神楽町"]
YEARS = [("R6", "令和6年度"), ("R7", "令和7年度"), ("R8", "令和8年度")]
Y8, Y8L = "R8", "令和8年度"

FONT = "游ゴシック"
NAVY, HEAD = "1F3864", "4472C4"
IN_Y, OK_G, NG_O, MID_B, GRAY = "FFF2CC", "E2EFDA", "FCE4D6", "DEEBF7", "F2F2F2"
_thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

wb = Workbook()
wb.remove(wb.active)
CHK = []


def chk(no, naiyo, kekka, ok):
    CHK.append((no, naiyo, kekka, "適合" if ok else "不適合"))


def sheet(name, title, subtitle, widths, freeze="A5", landscape=True):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    c = ws.cell(row=1, column=1, value=title)
    c.font = Font(name=FONT, size=14, bold=True, color=NAVY)
    ws.row_dimensions[1].height = 22
    c = ws.cell(row=2, column=1, value=subtitle)
    c.font = Font(name=FONT, size=9)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=2, start_column=1, end_row=2,
                   end_column=len(widths))
    ws.row_dimensions[2].height = 46
    ws.freeze_panes = freeze
    # 印刷（PDF）の体裁
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "4:4"
    ws.oddFooter.center.text = "&A　&P/&N"
    ws.oddFooter.center.size = 8
    return ws


def header(ws, row, cols, height=30):
    for i, v in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=HEAD)
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = height
    return row + 1


def body(ws, row, vals, fills=None, height=24, align=None, bold=False,
         fmt=None):
    for i, v in enumerate(vals, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(name=FONT, size=9, bold=bold)
        c.alignment = Alignment(wrap_text=True, vertical="top",
                                horizontal=(align or {}).get(i, "left"))
        c.border = BORDER
        if fmt and i in fmt:
            c.number_format = fmt[i]
        if fills and fills.get(i):
            c.fill = PatternFill("solid", fgColor=fills[i])
    ws.row_dimensions[row].height = height
    return row + 1


def lead(ws, row, text, span):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=10, bold=True, color=NAVY)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = 18
    return row + 1


def note(ws, row, text, span, height=70):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=8.5)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = height
    return row + 1


N0 = "#,##0"
PCT1 = "0.0"

# ------------------------------------------------------------ 集計
ITEM8 = KI.ITEM[Y8]
DET8 = KD.DETAIL[Y8L]
GOKEI = {t: KF.KOF[t][Y8]["推進・支援合計"] for t in TOWNS}
SUISHIN = {t: KF.KOF[t][Y8]["推進合計"] for t in TOWNS}
SHIEN = {t: KF.KOF[t][Y8]["支援合計"] for t in TOWNS}


def item_sum(kubun=None, moku=None, year=Y8):
    """大項目の配点と3町の得点の合計。"""
    rows = [r for r in KI.ITEM[year]
            if (kubun is None or r[0] == kubun)
            and (moku is None or r[1] == moku)]
    return (sum(r[4] for r in rows),
            {t: sum(r[5][KI.TOWNS.index(t)] for r in rows) for t in TOWNS})


MOKU = []
for r in ITEM8:
    if (r[0], r[1]) not in MOKU:
        MOKU.append((r[0], r[1]))

# ============================================================ 00
ws = sheet("00_この表について",
           "保険者機能強化推進交付金等　市町村分の評価指標と構成3町の評価",
           "市町村分の評価指標の構造（交付金・目標・指標群・大項目・明細列・"
           "配点）と、構成3町の3か年の得点をまとめたものです。"
           "評価は市町村単位で行われ、広域連合の行に得点はありません。",
           [22, 56, 46], landscape=False)
r = 4
r = header(ws, r, ["区分", "内容", "備考"])
for a, b, c in [
    ("交付金の種類",
     "保険者機能強化推進交付金（推進）と"
     "介護保険保険者努力支援交付金（支援）の2つがあります。",
     "本表では「推進」「支援」と略しています。"),
    ("評価の単位", "市町村単位で評価されます。",
     "保険者は大雪地区広域連合ですが、公表資料の広域連合の行に得点はなく、"
     "構成3町それぞれの行に記載があります。"),
    ("評価の対象年度",
     "令和8年度の評価指標による評価の対象は、"
     "令和7年度（2025年度）に実施した取組です。",
     "評価指標（42頁）の記載により確定しました。"),
    ("指標の構成",
     "交付金2種類×目標Ⅰ〜Ⅳ。目標Ⅰ〜Ⅲは体制・取組指標群（ⅰ）と"
     "活動指標群（ⅱ）、目標Ⅳは成果指標群から成ります。",
     "大項目は%d件、明細列（ア・イ・ウ・エ及び①〜④）は%d件です。"
     % (len(ITEM8), len(DET8))),
    ("配点の見方",
     "大項目の配点を合算すると%d点になりますが、これは満点ではありません。"
     % sum(r[4] for r in ITEM8),
     "目標Ⅳ成果指標群は選択肢が排他的であり、"
     "明細列の配点を合算すると満点（各100点）を超えます。"
     "得点の推移と町間の比較に用います。"),
    ("全国該当率",
     "当該明細列に得点のある市町村の数を集計対象（1,741市町村）で"
     "除したものです。",
     "公表資料のヘッダーには得点合計と平均が入るため、"
     "原本から算定しました。"
     "全国の多くの市町村が得点しているのに当区域が0点である項目を"
     "洗い出すために用います。"),
    ("都道府県分との違い",
     "都道府県分は市町村分とは別の評価であり、"
     "当広域連合及び構成3町の得点はありません。",
     "同じ目標名でも評価項目の中身が違います"
     "（認知症総合支援は都道府県分が活動指標4項目、市町村分が3項目）。"),
    ("個人情報の取扱い", "本表は集計値のみで構成しています。",
     "担当者名・電話番号・メールアドレスは収録していません。"),
]:
    r = body(ws, r, [a, b, c], height=52, fills={1: GRAY})
r += 1
r = lead(ws, r, "令和8年度の得点（推進・支援合計）", 3)
r = header(ws, r, ["町", "得点", "内訳・全国の位置"])
for t in TOWNS:
    r = body(ws, r, [t, GOKEI[t],
                     "推進%d点・支援%d点／全国の下位から%.1f％の位置"
                     % (SUISHIN[t], SHIEN[t], Z.PCT[Y8L][t])],
             height=24, fills={1: GRAY}, fmt={2: N0})
r = note(ws, r, "※ 3町の得点を足しても広域連合の評価にはなりません。"
                "評価が市町村単位で行われるためです。"
                "サービス見込量・給付費・保険料は保険者（広域連合）を"
                "単位として算定しますが、交付金の評価は町ごとに見ます。", 3,
         height=40)

# ============================================================ 01
ws = sheet("01_評価指標の構造", "評価指標の構造（令和8年度）",
           "交付金・目標・指標群ごとの配点と、構成3町の得点です。",
           [8, 40, 10, 12, 12, 12, 12, 12], landscape=True)
r = 4
r = header(ws, r, ["交付金", "目標", "大項目数", "配点",
                   "東川町", "美瑛町", "東神楽町", "配点に対する割合（3町平均）"])
for kubun, moku in MOKU:
    hai, sc = item_sum(kubun, moku)
    n = len([x for x in ITEM8 if x[0] == kubun and x[1] == moku])
    avg = sum(sc.values()) / 3.0 / hai if hai else None
    r = body(ws, r, [kubun, moku, n, hai,
                     sc["東川町"], sc["美瑛町"], sc["東神楽町"], avg],
             height=32, align={1: "center", 3: "center"},
             fills={1: GRAY},
             fmt={3: N0, 4: N0, 5: N0, 6: N0, 7: N0, 8: "0.0%"})
for kubun in ("推進", "支援"):
    hai, sc = item_sum(kubun)
    r = body(ws, r, [kubun, "小計", "", hai,
                     sc["東川町"], sc["美瑛町"], sc["東神楽町"],
                     sum(sc.values()) / 3.0 / hai],
             height=22, bold=True, align={1: "center"},
             fills={i: MID_B for i in range(1, 9)},
             fmt={4: N0, 5: N0, 6: N0, 7: N0, 8: "0.0%"})
hai, sc = item_sum()
r = body(ws, r, ["合計", "推進・支援", len(ITEM8), hai,
                 sc["東川町"], sc["美瑛町"], sc["東神楽町"],
                 sum(sc.values()) / 3.0 / hai],
         height=22, bold=True, align={1: "center", 3: "center"},
         fills={i: NG_O for i in range(1, 9)},
         fmt={3: N0, 4: N0, 5: N0, 6: N0, 7: N0, 8: "0.0%"})
r = note(ws, r, "※ 上表の得点は公表されている推進合計・支援合計・"
                "推進・支援合計に一致します（3町とも）。"
                "配点の合計は満点ではありません（00シート）。", 8, height=34)

# ============================================================ 02
ws = sheet("02_大項目別の評価", "大項目別の評価（令和8年度）",
           "評価指標の大項目ごとの配点と、構成3町の得点です。"
           "3町とも0点の項目、町により差がある項目に色を付けています。",
           [8, 34, 24, 40, 8, 9, 9, 9, 8, 26], landscape=True)
r = 4
r = header(ws, r, ["交付金", "目標", "指標群", "大項目", "配点",
                   "東川町", "美瑛町", "東神楽町", "差", "見方"])
ZERO3, SA = [], []
for kubun, moku, gun, nm, hai, sc in ITEM8:
    d = dict(zip(KI.TOWNS, sc))
    sa = max(sc) - min(sc)
    if all(v == 0 for v in sc):
        mi, f = "3町とも0点", NG_O
        ZERO3.append((kubun, moku, nm, hai))
    elif sa > 0:
        mi, f = "町により%d点の差がある" % sa, IN_Y
        SA.append((kubun, moku, nm, hai, sa))
    else:
        mi, f = "3町とも同じ", None
    r = body(ws, r, [kubun, moku, gun, nm, hai,
                     d["東川町"], d["美瑛町"], d["東神楽町"], sa, mi],
             height=26, align={1: "center", 5: "center", 9: "center"},
             fills={1: GRAY, 10: f},
             fmt={5: N0, 6: N0, 7: N0, 8: N0, 9: N0})
r = body(ws, r, ["計", "", "", "", sum(x[4] for x in ITEM8)]
         + [item_sum()[1][t] for t in TOWNS] + ["", ""],
         height=22, bold=True, align={1: "center"},
         fills={i: MID_B for i in range(1, 11)},
         fmt={5: N0, 6: N0, 7: N0, 8: N0})
r = note(ws, r, "※ 大項目%d件のうち、3町とも0点は%d件、"
                "町により差があるものは%d件です。"
                "得点0が「取組がない」のか「取組はあったが要件を"
                "満たさない」のか「報告していない」のかは、"
                "公表資料からは判別できません。"
                % (len(ITEM8), len(ZERO3), len(SA)), 10, height=42)

# ============================================================ 03
ws = sheet("03_明細列別の評価", "明細列別の評価（令和8年度）",
           "大項目の下の明細列（ア・イ・ウ・エ及び①〜④）ごとの配点、"
           "全国該当率、構成3町の得点です。"
           "全国該当率は、当該明細列に得点のある市町村の割合です。",
           [8, 30, 22, 34, 6, 7, 11, 9, 9, 9], landscape=True)
r = 4
r = header(ws, r, ["交付金", "目標", "指標群", "大項目", "枝番", "配点",
                   "全国該当率（％）", "東川町", "美瑛町", "東神楽町"])
for kubun, moku, gun, nm, eda, hai, ritsu, sc in DET8:
    d = dict(zip(KD.TOWNS, sc))
    zero = all((v or 0) == 0 for v in sc)
    r = body(ws, r, [kubun, moku, gun, nm, eda, hai, ritsu,
                     d["東川町"], d["美瑛町"], d["東神楽町"]],
             height=22, align={1: "center", 5: "center", 6: "center"},
             fills={1: GRAY,
                    7: (NG_O if (zero and ritsu >= 45.0) else None)},
             fmt={6: N0, 7: PCT1, 8: N0, 9: N0, 10: N0})
r = note(ws, r, "※ 全国該当率の欄に色が付いているものは、"
                "3町とも0点でありながら全国では45％以上の市町村が"
                "得点している明細列です（06シートに一覧）。"
                "明細列の得点を交付金別に合算すると、"
                "公表の推進合計・支援合計に一致します。", 10, height=40)

# ============================================================ 04
ws = sheet("04_3か年の推移", "3か年の推移",
           "令和6年度・令和7年度・令和8年度の得点です。"
           "年度は交付金の年度であり、評価の対象はその前年度の取組です。",
           [10, 34, 24, 36, 8, 9, 9, 9, 9, 9, 9], landscape=True)
r = 4
r = lead(ws, r, "1　合計の推移", 11)
r = header(ws, r, ["年度", "区分", "", "", "配点",
                   "東川町", "美瑛町", "東神楽町", "全国平均", "北海道平均",
                   "全国の位置（下位から％）"])
for y, yl in YEARS:
    hai, sc = item_sum(year=y)
    zen = Z.ZEN[yl].get("推進・支援合計")
    hok = Z.HOK[yl].get("推進・支援合計")
    r = body(ws, r, [yl, "推進・支援合計", "", "", hai,
                     KF.KOF["東川町"][y]["推進・支援合計"],
                     KF.KOF["美瑛町"][y]["推進・支援合計"],
                     KF.KOF["東神楽町"][y]["推進・支援合計"],
                     zen, hok,
                     "東川%.1f／美瑛%.1f／東神楽%.1f"
                     % (Z.PCT[yl]["東川町"], Z.PCT[yl]["美瑛町"],
                        Z.PCT[yl]["東神楽町"])],
             height=24, fills={1: GRAY}, bold=(y == Y8),
             fmt={5: N0, 6: N0, 7: N0, 8: N0, 9: "0.0", 10: "0.0"})
r += 1
r = lead(ws, r, "2　目標別の推移", 11)
r = header(ws, r, ["年度", "交付金", "目標", "", "配点",
                   "東川町", "美瑛町", "東神楽町", "", "", ""])
for kubun, moku in MOKU:
    for y, yl in YEARS:
        hai, sc = item_sum(kubun, moku, y)
        r = body(ws, r, [yl, kubun, moku, "", hai,
                         sc["東川町"], sc["美瑛町"], sc["東神楽町"], "", "", ""],
                 height=22, align={2: "center"},
                 fills={1: GRAY, 3: (MID_B if y == Y8 else None)},
                 fmt={5: N0, 6: N0, 7: N0, 8: N0}, bold=(y == Y8))
r = note(ws, r, "※ 各年度の配点は評価指標の改正により変わります。"
                "得点の増減を見るときは配点の変化も併せて見てください。", 11,
         height=30)

# ============================================================ 05
ws = sheet("05_全国・北海道との比較", "全国・北海道との比較（令和8年度）",
           "全国1,741市町村・北海道179市町村の平均との比較です。",
           [30, 10, 12, 12, 12, 12, 12, 34], landscape=True)
r = 4
r = header(ws, r, ["項目", "東川町", "美瑛町", "東神楽町", "全国平均",
                   "北海道平均", "全国平均との差（3町平均）", "見方"])
KOMOKU = ["推進Ⅰ合計", "推進Ⅱ合計", "推進Ⅲ合計", "推進Ⅳ合計", "推進合計",
          "支援Ⅰ合計", "支援Ⅱ合計", "支援Ⅲ合計", "支援Ⅳ合計", "支援合計",
          "推進・支援合計"]
HIKAKU = []
for k in KOMOKU:
    zen = Z.ZEN[Y8L].get(k)
    hok = Z.HOK[Y8L].get(k)
    vals = [Z.MACHI[Y8L][t].get(k) for t in TOWNS]
    if zen is None or any(v is None for v in vals):
        continue
    sa = sum(vals) / 3.0 - zen
    HIKAKU.append((k, sa))
    mi = ("全国平均を%.1f点%s" % (abs(sa), "上回る" if sa >= 0 else "下回る"))
    r = body(ws, r, [k] + vals + [zen, hok, sa, mi],
             height=22, fills={1: GRAY,
                               8: (OK_G if sa >= 0 else NG_O)},
             bold=(k == "推進・支援合計"),
             fmt={2: N0, 3: N0, 4: N0, 5: "0.0", 6: "0.0", 7: "+0.0;-0.0;0.0"})
r = note(ws, r, "※ 全国平均を下回る区分が、第10期に取組を強める候補です。"
                "ただし評価指標は毎年度改正されるため、"
                "同じ項目名でも年度により中身が変わることがあります。", 8,
         height=34)

# ============================================================ 06
ws = sheet("06_得点していない項目", "3町とも得点していない項目（令和8年度）",
           "3町とも0点でありながら、全国では多くの市町村が得点している"
           "明細列です。第10期の取組の候補として掲げています。",
           [8, 30, 22, 34, 6, 7, 12, 40], landscape=True)
r = 4
r = lead(ws, r, "1　全国該当率が45％以上のもの", 8)
r = header(ws, r, ["交付金", "目標", "指標群", "大項目", "枝番", "配点",
                   "全国該当率（％）", "見方"])
TORI = KD.torikoboshi(Y8L, 45.0)
for kubun, moku, gun, nm, eda, hai, ritsu, _sc in TORI:
    r = body(ws, r, [kubun, moku, gun, nm, eda, hai, ritsu,
                     "全国の%.1f％の市町村が得点しています。" % ritsu],
             height=24, align={1: "center", 5: "center", 6: "center"},
             fills={1: GRAY, 7: NG_O},
             fmt={6: N0, 7: PCT1})
r += 1
r = lead(ws, r, "2　大項目の単位で3町とも0点のもの", 8)
r = header(ws, r, ["交付金", "目標", "大項目", "", "", "配点", "", "見方"])
for kubun, moku, nm, hai in ZERO3:
    r = body(ws, r, [kubun, moku, nm, "", "", hai, "",
                     "大項目の全ての明細列で3町とも0点です。"],
             height=24, align={1: "center", 6: "center"},
             fills={1: GRAY}, fmt={6: N0})
r = note(ws, r, "※ 得点0が「取組がない」のか「取組はあったが要件を"
                "満たさない」のか「報告していない」のかは、"
                "公表資料からは判別できません。評価調書によりご確認ください。"
                "第10期の施策に位置づけるかどうかは、"
                "取組の必要性から判断するものであり、"
                "得点のみを理由とはしません。", 8, height=44)

# ============================================================ 07 自己点検
chk(1, "大項目の得点の合計が公表の推進・支援合計に一致すること",
    "・".join("%s %d点" % (t, item_sum()[1][t]) for t in TOWNS),
    all(item_sum()[1][t] == GOKEI[t] for t in TOWNS))
chk(2, "交付金別の合計が公表の推進合計・支援合計に一致すること",
    "推進・支援とも3町で一致",
    all(item_sum("推進")[1][t] == SUISHIN[t]
        and item_sum("支援")[1][t] == SHIEN[t] for t in TOWNS))
_det = {t: sum(r[7][KD.TOWNS.index(t)] or 0 for r in DET8) for t in TOWNS}
chk(3, "明細列の得点の合計が大項目の得点の合計に一致すること",
    "・".join("%s %d点" % (t, _det[t]) for t in TOWNS),
    all(_det[t] == GOKEI[t] for t in TOWNS))
chk(4, "3か年とも同じ検算が成り立つこと",
    "3か年×3町の9通りで一致",
    all(item_sum(year=y)[1][t] == KF.KOF[t][y]["推進・支援合計"]
        for y, _ in YEARS for t in TOWNS))
chk(5, "大項目が53件であること", "%d件" % len(ITEM8), len(ITEM8) == 53)
chk(6, "明細列が290件以上あること", "%d件" % len(DET8), len(DET8) >= 290)
chk(7, "集計対象の市町村数が3か年とも1,741であること",
    "・".join("%s %s" % (yl, format(Z.N[yl], ",")) for _y, yl in YEARS),
    all(Z.N[yl] == 1741 for _y, yl in YEARS))
chk(8, "全国該当率が0〜100の範囲にあること",
    "最小%.1f％・最大%.1f％" % (min(r[6] for r in DET8),
                                max(r[6] for r in DET8)),
    all(0.0 <= r[6] <= 100.0 for r in DET8))
chk(9, "3町とも0点の大項目が1件以上あること",
    "%d件" % len(ZERO3), len(ZERO3) >= 1)
chk(10, "町により差がある大項目が1件以上あること",
     "%d件" % len(SA), len(SA) >= 1)
chk(11, "全国の位置（パーセンタイル）が3町とも50％未満であること",
     "・".join("%s %.1f％" % (t, Z.PCT[Y8L][t]) for t in TOWNS),
     all(Z.PCT[Y8L][t] < 50.0 for t in TOWNS))
chk(12, "認知症総合支援（支援 目標Ⅱ）の得点が公表値に一致すること",
     "・".join("%s %d点" % (t, KF.KOF[t][Y8]["Ⅱ合計_2"]) for t in TOWNS),
     all(item_sum("支援", "目標Ⅱ 認知症総合支援を推進する")[1][t]
         == KF.KOF[t][Y8]["Ⅱ合計_2"] for t in TOWNS))
_ALL = []
for _ws in wb.worksheets:
    for _row in _ws.iter_rows(values_only=True):
        for _v in _row:
            if isinstance(_v, str):
                _ALL.append(_v)
_NG = ["〜に由来する", "全国トップ級", "1件も"]
chk(13, "禁止表現が含まれていないこと",
     "・".join(w for w in _NG if any(w in v for v in _ALL)) or "0件",
     not [w for w in _NG if any(w in v for v in _ALL)])
import re as _re                                       # noqa: E402
chk(14, "個人情報（電話番号・メールアドレス）が含まれていないこと", "0件",
     not [v for v in _ALL
          if _re.search(r"[\w.+-]+@[\w.-]+", v)
          or _re.search(r"0\d{1,4}[-(（]\d{2,4}[-)）]\d{3,4}", v)])
_NAIBU = ["固定値", "実物から", "runpy", ".py", "スクリプト", "再実行"]
chk(15, "受託者の内部の仕組み・作業経過の語が残っていないこと",
     "・".join(w for w in _NAIBU if any(w in v for v in _ALL)) or "0件",
     not [w for w in _NAIBU if any(w in v for v in _ALL)])
chk(16, "強調記号が本文に残っていないこと", "0件",
     not [v for v in _ALL if "**" in v])

ws = sheet("07_自己点検", "自己点検",
           "本表の値が公表資料の収録値と合っていることを機械で確かめた"
           "結果です。1件でも不適合があれば作成を止めます。",
           [6, 52, 46, 12], landscape=False)
r = header(ws, 4, ["No.", "確かめたこと", "結果", "判定"])
for no, naiyo, kekka, han in CHK:
    r = body(ws, r, [no, naiyo, kekka, han], height=28,
             align={1: "center", 4: "center"},
             fills={4: OK_G if han == "適合" else NG_O})

os.makedirs(ODIR, exist_ok=True)
wb.save(OUT)
print("出力:", OUT)
print("大項目%d件・明細列%d件／3町とも0点の大項目%d件・"
      "全国該当率45％以上で0点の明細列%d件"
      % (len(ITEM8), len(DET8), len(ZERO3), len(TORI)))
print("自己点検 %d件　不適合 %d件"
      % (len(CHK), sum(1 for x in CHK if x[3] != "適合")))
for x in CHK:
    if x[3] != "適合":
        print("  不適合 No.%s %s → %s" % (x[0], x[1], x[2]))
if any(x[3] != "適合" for x in CHK):
    sys.exit(1)
