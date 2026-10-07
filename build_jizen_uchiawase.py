# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画
　　構成町ヒアリング前の事前打合せ（令和8年10月8日）の確認事項.

令和8年10月6日のご依頼
  「8に構成町のヒアリング前の事前打ち合わせが入りました。
    確認内容を影響度と仮置きデータをもとにご教示下さい。」

━━ 本資料の位置づけ ━━

構成3町への町別ヒアリング（令和8年10月中旬から下旬）の前に、
広域連合と受託者で行う事前打合せで用いるものである。

ヒアリングの場で3町へ伺う事項は既に別冊にまとめている。
**本資料は「ヒアリングに入る前に広域連合で決めていただくこと」と
「ヒアリングで3町へ伺うことの最終確認」に絞る。**

━━ 数え方 ━━

確認事項は業務工程管理表（正の台帳）から、
決着しない場合の当方の扱い（仮置き）は既定値の一覧から、
影響度の点は日次の表と同じ採点から、
見込量・給付費・保険料は算定から読む。**台帳を二重に持たない。**

━━ 最低限の確認事項（令和8年10月7日のご依頼） ━━

  「1〜2時間の打ち合わせのため、最低限確認すべき事項と優先順位も
    合わせてご教示下さい。」

未決は160件を超えるため、**この場で決めるものを絞る。**
A群（この場で決めるもの）・B群（時間があれば）・C群（既定値のまま進めるもの）に
分け、所要の目安を分で付した（01シート）。
**A群は次の4つの観点のうち2つ以上に当たるものとした。**
  ①ヒアリングの問いが立たないもの（広域連合の置き方が先に要る）
  ②月額への効きが大きいもの（100円以上）
  ③法定記載事項に直結するもの
  ④第2次概算（令和8年10月30日）の前提になるもの

シート構成
  00_この打合せについて
  01_最低限の確認事項と時間配分
  02_先に申し上げること
  03_事前に決めていただきたいこと
  04_ヒアリングで3町へ伺うこと
  05_月額を動かす前提
  06_ヒアリングの進め方
  07_自己点検

出力
  output/第10期計画_ヒアリング前事前打合せ_確認事項.xlsx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import datetime
import io
import os
import re
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import repo_paths as RP

sys.path.insert(0, RP.ROOT)
import kakunin_score as KS                                   # noqa: E402

if isinstance(sys.stdout, io.TextIOWrapper):
    sys.stdout.reconfigure(encoding="utf-8")

ODIR = RP.OUTPUT
OUT = os.path.join(ODIR, "第10期計画_ヒアリング前事前打合せ_確認事項.xlsx")

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


# ============================================================ 暦
UCHIAWASE = datetime.date(2026, 10, 8)        # 事前打合せの日
KIJUNBI = KS.wareki(datetime.date.today())
UBI = KS.wareki(UCHIAWASE)
NOW_M = KS.yyyymm(UCHIAWASE)

SC = KS.sorted_check(NOW_M)                   # (点, 内訳, 行) の降順
KITEI = KS.KITEI
S = KS.SANTEI
G = S["G_DO"]
SUEOKI = KS.SUEOKI
KUBUN = KS.KUBUN

# シートの番号は本辞書で定める。注記の中の「○○シート」もここから引く。
# 番号を直書きすると、シートを増やしたときに注記の側がずれる。
SH = {"about": "00", "agenda": "01", "saki": "02", "kimeru": "03",
      "machi": "04", "lever": "05", "susume": "06", "chk": "07"}


def shn(key):
    """「03シート」のような参照の文字列。"""
    return SH[key] + "シート"


MACHI_W = ("3町", "東川", "美瑛", "東神楽")
TEN_MIN = 4             # 決めていただきたいことに掲げる点の下限（絞り込みの条件）


def _is_machi(x):
    return any(w in str(x[6]) for w in MACHI_W)


# ヒアリングで3町へ伺うもの
MACHI_KO = [(p, u, x) for p, u, x in SC if _is_machi(x)]
MACHI_BETSU = [t for t in MACHI_KO if "3町" not in str(t[2][6])]
MACHI_KYO = [t for t in MACHI_KO if "3町" in str(t[2][6])]
# 資料提供依頼のうち入手先に町を含むもの
LACK_MACHI = [x for x in KS.LACK_MACHI
              if any(w in str(x[5]) for w in MACHI_W)]


def kitei_of(no):
    return KITEI.get(no, "―")


def yen_of(no):
    g = KS.GETSUGAKU.get(no)
    return ("最大%s円" % "{:,}".format(g[0])) if g else "―"


# ============================================================ 最低限の確認事項
# 令和8年10月7日のご依頼による。打合せが1〜2時間であるため、
# この場で決めるものを絞る。
# A群は次の4つの観点のうち2つ以上に当たるものとした。
#   ①ヒアリングの問いが立たない（広域連合の置き方が先に要る）
#   ②月額への効きが大きい（100円以上）
#   ③法定記載事項に直結する
#   ④第2次概算（令和8年10月30日）の前提になる
KANTEN = [
    ("①", "ヒアリングの問いが立たない",
     "広域連合としての置き方が決まらないと、各町へ伺う問いが立たない"),
    ("②", "月額への効きが大きい",
     "保険料の月額を100円以上動かし得る"),
    ("③", "法定記載事項に直結する",
     "介護保険法第117条が定める記載事項そのものである"),
    ("④", "第2次概算の前提になる",
     "令和8年10月30日に提示する算定の前提として要る"),
]

# (記号, 確認事項No.の並び, 決めること, なぜこの場で決めるか, 当たる観点, 分)
AGENDA_A = [
    ("A1", [88, 113, 84],
     "日常生活圏域と必要利用定員総数（基盤整備の単位）",
     "**これが決まらないと、各町へ整備の方針を伺うことができません。**"
     "計画素案の未確定箇所は第6章第4節（施設の見込み）に最も多く集中しており、"
     "基盤整備の単位が決まればその大半が解消します。"
     "法第117条第2項第1号の法定記載事項です。",
     "①③④", 15),
    ("A2", [137, 104, 109, 110],
     "1人1月あたり給付費（単価）をどの年度で固定するか",
     "**月額への効きが最も大きい事項です（最大＋327円）。**"
     "現在は令和7年度の実績で固定しており、"
     "令和8年度及び令和9年度の介護報酬改定と、"
     "実績として現れている単価の伸び（年率1.32％から1.75％）を"
     "見込んでいません。第2次概算の前提になります。",
     "②④", 15),
    ("A3", [112, 150],
     "地域支援事業及び認知症総合支援事業の量の見込み",
     "**影響度の点が最も高い2件です。**"
     "法第117条第2項第2号の法定記載事項であり、"
     "3町へどの資料をご依頼するかがここで決まります。"
     "3町の実施報告書で確認できた実施状況は単位が延べ人数・回数・食数で"
     "混在しており合計できないため、"
     "令和6年度の利用者実人数を据え置く扱いでよいかをお決めいただきたい。",
     "①③④", 10),
    ("A4", [152, 153, 174],
     "人口の将来推計を案Cのままとするか",
     "**町別データシートで町別の将来推計をお示しするかに直結し、"
     "ヒアリングで配付する資料の作りが変わります。**"
     "令和8年9月10日のご了承では、町別は現況のみを計画本文に加え、"
     "将来推計は3町計のままとしています。"
     "第1号被保険者数の系列が3つあるため、どれを正式に採るかも併せて要ります。",
     "①④", 10),
    ("A5", [106, 107, 105],
     "サービス別の趨勢の補正と施策反映を行うか",
     "第2次概算の前提になります。"
     "補正のみで月額▲58円、施設及び居住系を減らさない条件を加えると＋48円です。"
     "現在はいずれも**安全側**（保険料が不足する側に誤らない向き）に"
     "織り込まないまま置いています。",
     "②④", 5),
]
AGENDA_B = [
    ("B1", [93, 86],
     "調整交付金の扱い（適正化の取組による減額・交付見込額の控除）",
     "令和8年9月18日の社会保障審議会介護保険部会で、"
     "年齢区分の精緻化により交付額が増加する保険者が"
     "介護給付適正化の主要3事業のいずれかを実施していない場合に"
     "増加額の5％を減額することが示されました。"
     "当広域連合はケアプラン等の点検を実施していないため該当します。"
     "交付見込額を収納必要額から控除しない当方の扱いの確認も併せて。", 5),
    ("B2", [131, 114],
     "計画本文の掲げ方（見込量の表・中長期推計）",
     "既に計画素案へ反映している内容のご確認です。"
     "見込量の表に利用者数・利用回（日）数・給付費を掲げ、"
     "中長期推計は令和17年度・令和22年度を掲げています。", 5),
    ("B3", [175],
     "介護情報基盤及びケアプランデータ連携への対応の工程",
     "令和10年4月1日までに3町の介護保険事務システムの標準化対応と"
     "データ移行を終えることが国から示されています。"
     "**ヒアリングで3町へ伺う項目として加えるかをご確認いただきたい。**", 5),
    ("B4", [130],
     "市町村認知症施策推進計画と意見の聴取",
     "3町がそれぞれ市町村計画を策定する案Bで素案へ反映済みです。"
     "認知症の人及びその家族等からの意見の聴取の主体・方法・時期が"
     "残っています。", 5),
]
_A_NO = [n for a in AGENDA_A for n in a[1]]
_B_NO = [n for b in AGENDA_B for n in b[1]]
AB_NO = _A_NO + _B_NO
A_MIN = sum(a[5] for a in AGENDA_A)
B_MIN = sum(b[4] for b in AGENDA_B)

# 広域連合で決まるもの（確認先に町を含まないもの）のうち点の高いもの。
# **A群・B群に挙げたものは点が下限に届かなくても掲げる**
# （この場で決めていただくものの「決まらない場合の扱い」が
#  一覧に現れないことがないようにするため）。
KIMERU = [(p, u, x) for p, u, x in SC
          if not _is_machi(x) and (p >= TEN_MIN or x[0] in AB_NO)]

# 時間配分の案。(枠, 区分, 内容, 分)
WARI = [
    (90, "1", "冒頭　仮置きによる算定の結果と、低く出る側に寄っていること"
     "（%s）" % shn("saki"), 5),
    (90, "2", "A群　この場で決めること（%d件・%s）" % (len(AGENDA_A),
                                               shn("agenda")), A_MIN),
    (90, "3", "3町へ伺う事項と順序の確認（%s）" % shn("machi"), 15),
    (90, "4", "ヒアリングの進め方（日程・配付資料・記録の様式。%s）"
     % shn("susume"), 10),
    (90, "5", "予備", 5),
    (120, "1", "冒頭　仮置きによる算定の結果と、低く出る側に寄っていること"
     "（%s）" % shn("saki"), 5),
    (120, "2", "A群　この場で決めること（%d件・%s）" % (len(AGENDA_A),
                                                shn("agenda")), A_MIN),
    (120, "3", "B群　時間があれば決めること（%d件・%s）" % (len(AGENDA_B),
                                                  shn("agenda")), B_MIN),
    (120, "4", "3町へ伺う事項と順序の確認（%s）" % shn("machi"), 20),
    (120, "5", "ヒアリングの進め方（日程・配付資料・記録の様式。%s）"
     % shn("susume"), 15),
    (120, "6", "予備", 5),
]


# ============================================================ 体裁
def _plain(v):
    """xlsx は Markdown を解釈しないため、書き出しの時点で ** を落とす。"""
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
    if freeze:
        ws.freeze_panes = freeze
    ws.sheet_view.showGridLines = False
    return ws


def header(ws, row, cols, height=30):
    for i, c in enumerate(cols, start=1):
        cell = ws.cell(row, i, _plain(c))
        cell.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=HEAD)
        cell.alignment = Alignment(wrap_text=True, vertical="center",
                                   horizontal="center")
        cell.border = BORDER
    ws.row_dimensions[row].height = height
    return row + 1


def body(ws, row, vals, fills=None, height=22, align=None, bold=False):
    for i, v in enumerate(vals, start=1):
        cell = ws.cell(row, i, _plain(v))
        cell.font = Font(name=FONT, size=9, bold=bold)
        cell.alignment = Alignment(
            wrap_text=True, vertical="top",
            horizontal=(align or {}).get(i, "left"))
        cell.border = BORDER
        if fills and i in fills:
            cell.fill = PatternFill("solid", fgColor=fills[i])
    ws.row_dimensions[row].height = height
    return row + 1


def mcols(ws, row, c1, c2):
    """1行の中で列をつなぐ（長い文を幅の広い範囲に収めるため）。"""
    ws.merge_cells(start_row=row, start_column=c1, end_row=row,
                   end_column=c2)


def lead(ws, row, text, span=8):
    ws.cell(row, 1, _plain(text))
    ws.cell(row, 1).font = Font(name=FONT, size=11, bold=True, color=NAVY)
    ws.merge_cells(start_row=row, start_column=1,
                   end_row=row, end_column=span)
    ws.row_dimensions[row].height = 22
    return row + 1


def note(ws, row, text, span=8, height=None, fill=GRAY):
    ws.cell(row, 1, _plain(text))
    ws.cell(row, 1).font = Font(name=FONT, size=9)
    ws.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row, 1).fill = PatternFill("solid", fgColor=fill)
    ws.merge_cells(start_row=row, start_column=1,
                   end_row=row, end_column=span)
    ws.row_dimensions[row].height = height or (16 * (text.count("\n") + 3))
    return row + 1


def num(v, nd=0):
    return ("{:,.%df}" % nd).format(v)


# ============================================================ 00
ws = sheet(SH["about"] + "_この打合せについて",
           "構成町ヒアリング前の事前打合せ　確認事項",
           "構成3町への町別ヒアリング（令和8年10月中旬から下旬）の前に、"
           "広域連合と受託者で行う事前打合せ（" + UBI + "）で"
           "ご確認・ご判断をお願いしたい事項をまとめたものです。"
           "ヒアリングの場で3町へ伺う事項は別冊「3町ヒアリング資料」"
           "及び別冊「10月 地域課題の整理と3町協議の確認事項」にあります。"
           "作成日 " + KIJUNBI + "。",
           [4, 26, 12, 56])
r = 4
r = lead(ws, r, "1　この打合せで決めたいこと", span=4)
r = header(ws, r, ["", "事項", "シート", "なぜヒアリングの前に決めるか"])
NERAI = [
    ("最低限の確認事項と時間配分", SH["agenda"],
     "打合せが1〜2時間であるため、この場で決めるものをA群%d件に絞り、"
     "所要の目安を分で付した。90分の場合と120分の場合の時間配分の案を置く"
     % len(AGENDA_A)),
    ("ヒアリングに入る前に広域連合で決めていただくこと", SH["kimeru"],
     "3町へ伺う内容は、広域連合としての置き方が決まっていないと問いが立たない。"
     "例えば必要利用定員総数の基盤整備単位が決まらないと、"
     "各町へ整備の方針を伺うことができない"),
    ("3町へ伺うことの最終確認", SH["machi"],
     "町別の論点と3町共通の論点を、伺う相手と順序を含めて確かめる。"
     "町からご提供をお願いする資料もここで確かめる"),
    ("保険料の月額を動かす前提の共有", SH["lever"],
     "仮置きのまま算定しているものを示す。"
     "ヒアリングの場で町から整備の意向が示されると見込量が動くため、"
     "どこまでが動き得るのかを先に共有する"),
    ("ヒアリングの進め方", SH["susume"],
     "日程・配付する資料・記録の様式・ヒアリング後の工程を確かめる"),
]
for i, (a, b, c) in enumerate(NERAI, start=1):
    r = body(ws, r, [i, a, b, c], height=52, align={1: "center",
                                                    3: "center"})
r += 1
r = lead(ws, r, "2　掲げる範囲", span=4)
r = note(ws, r,
         "・完了していない確認事項は%d件あります。"
         "そのまま並べると一覧として用をなさないため、"
         "**この場で決めるものは%sのA群%d件（計%d分）に絞っています。**\n"
         "・**%sは「確認先が広域連合であり、影響度の点が%d点以上のもの」"
         "に絞って%d件**を掲げています"
         "（A群・B群に挙げたものは点が下限に届かなくても掲げています）。"
         "A群はこの中から選んでいます（%sに含まれないものは"
         "確認先に構成町を含むものです）。\n"
         "・**%sは確認先に構成町を含むもの%d件**です"
         "（町別%d件・3町共通%d件）。\n"
         "・影響度の点は、期限・止めている成果物・月額への効き・"
         "法定記載事項・決着しない場合の扱いの有無の5つで付けています。"
         "付け方は別冊「日次の状況と翌日の作業順位」と同じです。\n"
         "・**決まらない場合の当方の扱い（仮置き）は%d件すべてに置いています。**"
         "このため、ご決定をお待ちしている事項があっても"
         "第2次概算（令和8年10月30日）は組めます。"
         "仮置きは協議を進めるためのものであり、公表の確定の根拠にはなりません。"
         % (len(SC), shn("agenda"), len(AGENDA_A), A_MIN,
            shn("kimeru"), TEN_MIN, len(KIMERU), shn("kimeru"),
            shn("machi"), len(MACHI_KO),
            len(MACHI_BETSU), len(MACHI_KYO), len(SC)),
         span=4, height=180, fill=IN_Y)

# ============================================================ 01 最低限
ws = sheet(SH["agenda"] + "_最低限の確認事項と時間配分",
           "最低限ご確認いただきたい事項と優先順位",
           "打合せの時間が1〜2時間であることを踏まえ、"
           "完了していない確認事項%d件のうち、"
           "**この場で決めるものをA群%d件（計%d分）に絞りました。**"
           "B群%d件（計%d分）は時間があれば、"
           "残りは決まらない場合の当方の扱い（仮置き）のまま進めます。"
           % (len(SC), len(AGENDA_A), A_MIN, len(AGENDA_B), B_MIN),
           [6, 14, 34, 58, 7, 6])
r = 4
r = lead(ws, r, "1　A群　この場で決めていただきたいこと（計%d分）" % A_MIN,
         span=6)
r = header(ws, r, ["", "確認事項No.", "決めること",
                   "なぜこの場で決めるか", "観点", "目安"])
for kg, nos, koto, naze, kan, mn in AGENDA_A:
    r = body(ws, r,
             [kg, "・".join("No.%d" % n for n in nos), koto, naze, kan,
              "%d分" % mn],
             fills={1: MID_B, 6: IN_Y},
             height=96, align={1: "center", 5: "center", 6: "center"},
             bold=False)
r = note(ws, r,
         "注1）観点の記号は本シート3のとおりです。"
         "**A群は4つの観点のうち2つ以上に当たるものとしました。**\n"
         "注2）各件の「決まらない場合の当方の扱い（仮置き）」は%s（広域連合で"
         "決まるもの）及び%s（3町へ伺うもの）の表にあります。"
         "この場でご否認がなければ、そのまま第2次概算に用います。\n"
         "注3）A1は確認先に構成町を含みますが、**基盤整備の単位をどう置くかは"
         "広域連合としての判断であり、これを先に決めないと"
         "各町へ整備の方針を伺えません。**"
         % (shn("kimeru"), shn("machi")), 6, height=96, fill=IN_Y)
r += 1
r = lead(ws, r, "2　B群　時間があれば決めていただきたいこと（計%d分）" % B_MIN,
         span=6)
r = header(ws, r, ["", "確認事項No.", "決めること",
                   "この場で伺う理由", "観点", "目安"])
for kg, nos, koto, naze in [(b[0], b[1], b[2], b[3]) for b in AGENDA_B]:
    mn = [b[4] for b in AGENDA_B if b[0] == kg][0]
    r = body(ws, r,
             [kg, "・".join("No.%d" % n for n in nos), koto, naze, "―",
              "%d分" % mn],
             height=76, align={1: "center", 5: "center", 6: "center"})
r = note(ws, r,
         "注）B群はいずれも**既に計画素案へ反映している内容のご確認**又は"
         "**ヒアリングで伺う項目に加えるかのご確認**です。"
         "この場で決まらなくても第2次概算は組めます。", 6)
r += 1
r = lead(ws, r, "3　A群に選んだ観点", span=6)
_hr = r
r = header(ws, r, ["", "観点", "内容", "", "A群で当たるもの", ""])
mcols(ws, _hr, 3, 4)
mcols(ws, _hr, 5, 6)
for kg, nm, naiyo in KANTEN:
    _ataru = "・".join(a[0] for a in AGENDA_A if kg in a[4])
    _br = r
    r = body(ws, r, [kg, nm, naiyo, "", _ataru, ""],
             height=32, align={1: "center", 5: "center"})
    mcols(ws, _br, 3, 4)
    mcols(ws, _br, 5, 6)
r = note(ws, r,
         "注）4つの観点のうち1つだけに当たるものはB群又はC群としました。"
         "**期限は観点に用いていません。**"
         "期限が令和8年10月までのものが大半であり、"
         "期限だけでは順位が付かないためです。", 6)
r += 1
r = lead(ws, r, "4　時間配分の案", span=6)
_hr = r
r = header(ws, r, ["枠", "順", "内容", "", "", "分"])
mcols(ws, _hr, 3, 5)
for waku in (90, 120):
    _rows2 = [w for w in WARI if w[0] == waku]
    for j, (_w, kg, naiyo, mn) in enumerate(_rows2):
        _br = r
        r = body(ws, r,
                 [("%d分" % waku) if j == 0 else "", kg, naiyo, "", "",
                  "%d分" % mn],
                 fills={1: (MID_B if j == 0 else None)} if j == 0 else None,
                 height=26, align={1: "center", 2: "center", 6: "center"},
                 bold=(j == 0))
        mcols(ws, _br, 3, 5)
    _br = r
    r = body(ws, r, ["", "計", "", "", "",
                     "%d分" % sum(w[3] for w in _rows2)],
             fills={i: GRAY for i in range(1, 7)}, bold=True,
             height=22, align={2: "center", 6: "center"})
    mcols(ws, _br, 3, 5)
r = note(ws, r,
         "注1）**90分の場合はB群を扱いません。**"
         "B群は後日の持ち帰り又はメールでのご回答で足ります。\n"
         "注2）3町へ伺う事項と順序の確認（%s）は、"
         "A群が決まってから行うことを想定しています"
         "（A群の置き方により伺う問いが変わるためです）。"
         % shn("machi"), 6, height=52)
r += 1
r = lead(ws, r, "5　C群　この場では扱わないもの", span=6)
r = note(ws, r,
         "・A群・B群に挙げた確認事項は実数%d件です。"
         "残る%d件はこの場では扱わず、"
         "決まらない場合の当方の扱い（仮置き）のまま進めます。\n"
         "・このうち**国の告示・公布を待つもの（据え置きの区分C）4件は"
         "この場でお決めいただくことができません。**"
         "介護報酬改定率・第1号被保険者負担割合・所得段階の政令改正・"
         "令和9年4月新設の3区分です。"
         "告示・公布の後に再算定します。\n"
         "・**受託者の側で確定できるもの（区分A）は残っていません。**"
         "令和8年10月1日に4件（要介護度別の単価・1人1月あたり利用回（日）数・"
         "要介護度の構成・季節補正）を確定しました。\n"
         "・仮置きのまま計画素案へ反映した事項は%d件あります"
         "（別冊「3町合同意見交換会資料」第6節3に一覧があります）。"
         "いずれもご決定をお待ちしているものであり、"
         "ご否認があれば素案の側を改めます。"
         % (len(set(AB_NO)), len(SC) - len(set(AB_NO)), len(KS.HANEI)),
         span=6, height=130)

# ============================================================ 02
ws = sheet(SH["saki"] + "_先に申し上げること",
           "先に申し上げること（仮置きによる算定の結果）",
           "ご判断をお待ちしている事項はいずれも仮置きの値で算定しています。"
           "現時点の算定の結果は次のとおりです。",
           [4, 34, 22, 54])
r = 4
r = lead(ws, r, "1　現時点の算定（仮置きによるもの）", span=4)
r = header(ws, r, ["", "項目", "値", "見方"])
SAKI = [
    ("総給付費（第10期3か年計）", "%s円" % num(sum(S["KYUFU3"])),
     "要介護度別の認定者数に、要介護度別の利用率と1人1月あたり給付費を"
     "乗じて算定したものです"),
    ("標準給付費見込額（A）", "%s円" % num(G["A"]),
     "総給付費に特定入所者介護サービス費等の上乗せを加えたものです"),
    ("地域支援事業費（B）", "%s円" % num(G["B"]),
     "令和6年度決算額を3年据え置いたものです"),
    ("保険料収納必要額（J）", "%s円" % num(G["J"]),
     "J＝C＋D－E＋F＋G±H－I によります"),
    ("補正後第1号被保険者数（3か年）", "%s人" % num(G["③"], 1),
     "第9期条例の16段階・公費軽減前の乗率によります"),
    ("算定上の月額基準額", "%s円" % num(G["月額"], 2),
     "J÷予定収納率99.0％÷補正後被保険者数÷12 によります"),
    ("保険料基準額（百円未満四捨五入）", "%s円" % num(round(G["月額"] / 100) * 100),
     "第9期の基準額6,400円と同額になります"),
]
for i, (a, b, c) in enumerate(SAKI, start=1):
    r = body(ws, r, [i, a, b, c], height=36,
             fills={3: (MID_B if "基準額" in a else None)} if "基準額" in a
             else None, align={1: "center", 3: "right"},
             bold=("基準額（百円" in a))
r = note(ws, r,
         "注）いずれも仮置きによる値であり、ご決定及び国の告示により動きます。", 4)
r += 1
r = lead(ws, r, "2　この値は低く出る側に寄っています", span=4)
r = note(ws, r,
         "単価を令和7年度の実績で固定しており、"
         "**令和8年度及び令和9年度の介護報酬改定と、"
         "実績として現れている単価の伸びを見込んでいません。**\n"
         "・単価の実績の伸び（年率1.32％から1.75％）を織り込むと月額＋245円から＋327円\n"
         "・介護報酬改定率1％あたり月額＋61円\n"
         "・第1号被保険者負担割合が23％から24％になると月額約＋290円\n"
         "・**調整交付金の減額**（年齢区分の精緻化により交付額が増加する保険者が"
         "介護給付適正化の主要3事業のいずれかを実施していない場合に"
         "増加額の5％を減額。令和8年9月18日の社会保障審議会介護保険部会）。"
         "当広域連合はケアプラン等の点検を実施していないため該当しますが、"
         "増加額が示されていないため減額の額は算定できません\n"
         "一方、下向きに働くもの（介護給付適正化・サービス別の趨勢の補正・"
         "介護給付費準備基金の取崩し）はいずれも織り込んでいません。\n"
         "**上向きは国の告示・公布により生じるもので、資料のご提供とは関わりなく動きます。**"
         "下向きはご判断を要するものです。" + shn("lever") + "に一覧があります。",
         span=4, height=170, fill=IN_Y)
r += 1
r = lead(ws, r, "3　ヒアリングで見込量が動き得るところ", span=4)
r = note(ws, r,
         "本算定は、現に指定を受けている定員を第10期を通じて据え置く前提で"
         "置いています。ヒアリングの場で整備又は事業の実施の意向が示された場合は、"
         "開始年度から利用者数と給付費を加えることになります。\n"
         "特に次の3つは、町のご意向により動きます。\n"
         "・地域密着型介護老人福祉施設（区域内定員62人。令和11年度に61.9人＝99.9％に"
         "達する見込みで、令和12年度には定員を超えると見込まれます）\n"
         "・認知症対応型共同生活介護（区域内定員99人。令和11年度 月85.2人＝86.1％）\n"
         "・定期巡回・随時対応型訪問介護看護（区域内に事業所がなく、"
         "自然体では月0.6人。整備を行う場合は上乗せを置きます）",
         span=4, height=120)

# ============================================================ 02
ws = sheet(SH["kimeru"] + "_事前に決めていただきたいこと",
           "事前打合せで決めていただきたいこと（広域連合でお決めいただくもの）",
           "確認先が広域連合であり、影響度の点が%d点以上のもの"
           "（及びA群・B群に挙げたもの）%d件です。"
           "点の高い順に並べています。"
           "「決まらない場合の当方の扱い」は仮置きであり、"
           "この打合せでご否認がなければそのまま第2次概算に用います。"
           % (TEN_MIN, len(KIMERU)),
           [5, 5, 30, 20, 11, 10, 44])
r = 4
r = header(ws, 4, ["順", "No.", "確認事項", "止めている成果物", "期限",
                   "月額への効き", "決まらない場合の当方の扱い（仮置き）"])
for i, (p, u, x) in enumerate(KIMERU, start=1):
    r = body(ws, r,
             [i, x[0], x[3], x[5], x[8], yen_of(x[0]), kitei_of(x[0])],
             fills={6: (NG_O if KS.GETSUGAKU.get(x[0], (0,))[0] >= 100
                        else None)}
             if KS.GETSUGAKU.get(x[0], (0,))[0] >= 100 else None,
             height=58, align={1: "center", 2: "center", 5: "center",
                               6: "right"})
r = note(ws, r,
         "注1）**この打合せで特にお決めいただきたいのは、"
         "%sのA群%d件（計%d分）です。**"
         "本シートの中では次のものが当たります。\n%s\n"
         "注2）月額への効きの欄が「―」のものは、"
         "保険料には影響せず計画本文の記述に関わるものです。\n"
         "注3）点が%d点に届かないものも、"
         "A群・B群に挙げたものは本シートに掲げています。"
         % (shn("agenda"), len(AGENDA_A), A_MIN,
            "\n".join(
                "　%s　%s　%s（%d分）"
                % (a[0], "・".join("No.%d" % n for n in a[1]), a[2], a[5])
                for a in AGENDA_A),
            TEN_MIN),
         7, height=170, fill=IN_Y)

# ============================================================ 03
ws = sheet(SH["machi"] + "_ヒアリングで3町へ伺うこと",
           "ヒアリングで構成3町へ伺うこと",
           "確認先に構成町を含む確認事項%d件（町別%d件・3町共通%d件）と、"
           "町からご提供をお願いする資料%d件です。"
           "町別は町ごとに、3町共通は点の高い順に並べています。"
           % (len(MACHI_KO), len(MACHI_BETSU), len(MACHI_KYO),
              len(LACK_MACHI)),
           [5, 5, 12, 30, 11, 46])
r = 4
r = lead(ws, r, "1　町ごとに伺うこと", span=6)
r = header(ws, r, ["順", "No.", "町", "確認事項", "期限",
                   "決まらない場合の当方の扱い（仮置き）"])
_i = 0
for machi in ("東川", "美瑛", "東神楽"):
    for p, u, x in MACHI_BETSU:
        if machi not in str(x[6]):
            continue
        _i += 1
        r = body(ws, r, [_i, x[0], machi + "町", x[3], x[8],
                         kitei_of(x[0])],
                 height=46, align={1: "center", 2: "center", 3: "center",
                                   5: "center"})
r = note(ws, r,
         "注）いずれも別冊「10月 地域課題の整理と3町協議の確認事項」の"
         "町ごとのシートに同じ事項を掲げています。"
         "**ヒアリングの場では、この一覧の順に伺うことを想定しています。**", 6)
r += 1
r = lead(ws, r, "2　3町に共通して伺うこと", span=6)
r = header(ws, r, ["順", "No.", "確認先", "確認事項", "期限",
                   "決まらない場合の当方の扱い（仮置き）"])
for i, (p, u, x) in enumerate(MACHI_KYO, start=1):
    r = body(ws, r, [i, x[0], x[6], x[3], x[8], kitei_of(x[0])],
             height=46, align={1: "center", 2: "center", 5: "center"})
r += 1
r = lead(ws, r, "3　町からご提供をお願いする資料", span=6)
r = header(ws, r, ["順", "No.", "入手先", "資料", "希望時期",
                   "何が確定するか"])
for i, x in enumerate(LACK_MACHI, start=1):
    r = body(ws, r, [i, x[0], x[5], x[2], x[6], x[4]],
             height=46, align={1: "center", 2: "center", 5: "center"})
r = note(ws, r,
         "注）資料のご提供がない場合は、現に確認できている値で確定し、"
         "出所と制約を計画本文の注記に書く扱いとしています。"
         "算定そのものが止まるものはありません。", 6)

# ============================================================ 04
ws = sheet(SH["lever"] + "_月額を動かす前提",
           "保険料の月額を動かす前提（仮置きのまま置いているもの）",
           "据え置いた前提のうち、広域連合・3町のご判断又は資料を待つもの（区分B）と、"
           "国の告示・公布を待つもの（区分C）です。"
           "効きの大きいものから並べています。いずれも本算定には織り込んでいません。",
           [4, 6, 28, 34, 34, 11])
r = 4
_rows = []
for i, s_ in enumerate(SUEOKI, start=1):
    k = KUBUN.get(i)
    if k not in ("B", "C"):
        continue
    _rows.append((KS.max_yen(s_[5]) or 0, k, s_))
_rows.sort(key=lambda t: -t[0])
r = header(ws, 4, ["", "区分", "事項", "決まらない場合の当方の扱い（仮置き）",
                   "効き", "確認事項"])
for i, (y, k, s_) in enumerate(_rows, start=1):
    r = body(ws, r, [i, k, s_[1], KS.oki(s_[3]), KS.oki(s_[5]), s_[4]],
             fills={2: (NG_O if k == "C" else IN_Y),
                    5: (NG_O if y >= 100 else None)} if y >= 100
             else {2: (NG_O if k == "C" else IN_Y)},
             height=52, align={1: "center", 2: "center"})
r = note(ws, r,
         "注1）区分Bは広域連合・3町のご判断又は資料を待つもので、"
         "届かない場合は現に確認できている値で確定します。"
         "区分Cは国の告示・公布を待つもので、"
         "「受領できない」では済まないものです（公布後に再算定します）。\n"
         "注2）効きの欄の額は、本計画の算定の基礎で求めたものです。"
         "**効きは、据え置きの一覧で名を挙げている確認事項に結び付けています。**"
         "関係する確認事項が複数あるものは、代表の1件に額が付きます"
         "（" + shn("kimeru") + "で「―」となっているものに効きが"
         "ないという意味ではありません）。\n"
         "**介護報酬改定率を乗じることと単価の伸びを延ばすことを重ねると"
         "二重になります。**\n"
         "注3）受託者の側で確定できるもの（区分A）は残っていません"
         "（令和8年10月1日に要介護度別の単価・1人1月あたり利用回（日）数・"
         "要介護度の構成・季節補正の4件を確定しました）。", 6, height=110)

# ============================================================ 05
ws = sheet(SH["susume"] + "_ヒアリングの進め方", "ヒアリングの進め方",
           "日程・お持ちする資料・記録の様式・ヒアリング後の工程です。",
           [4, 22, 20, 56])
r = 4
r = lead(ws, r, "1　日程", span=4)
r = header(ws, r, ["", "時期", "場", "内容"])
NITTEI = [
    (UBI, "広域連合・受託者",
     "事前打合せ。本資料の" + shn("agenda") + "のA群の決定と、"
     + shn("machi") + "の伺う事項・順序の確認"),
    ("令和8年10月上旬（本打合せの後）", "3町合同",
     "意見交換会。3町共通の確認事項と地域課題。"
     "別冊「3町合同意見交換会資料」による"),
    ("令和8年10月中旬から下旬", "町ごと",
     "町別ヒアリング。" + shn("machi") + "の1の事項を町ごとに伺う"),
    ("令和8年10月30日", "広域連合", "サービス見込量 第2次概算の提示"),
    ("令和8年11月13日", "広域連合", "予算編成用の数値の提示"),
    ("令和8年11月", "広域連合・3町", "3町の意見の反映と計画の骨子の整理"),
]
for i, (a, b, c) in enumerate(NITTEI, start=1):
    r = body(ws, r, [i, a, b, c], height=36, align={1: "center"})
r += 1
r = lead(ws, r, "2　お持ちする資料（いずれも既にお送りしているもの）", span=4)
r = header(ws, r, ["", "資料", "使う場面", "内容"])
SHIRYO = [
    ("別冊「3町ヒアリング資料」", "ヒアリング本体",
     "伺う事項と当方の整理を章立てで示したもの"),
    ("別冊「キックオフ会議ヒアリングシート」", "ヒアリング本体",
     "伺った内容をその場で書き留める記録の用紙"),
    ("別冊「3町別の論点整理」", "ヒアリング本体",
     "町ごとの現在地と論点、3町共通の論点、協議の進め方"),
    ("別冊「10月 地域課題の整理と3町協議の確認事項」", "ヒアリング本体",
     "地域課題15件と、町ごとの確認事項（町別の3シート）"),
    ("別冊「3町合同意見交換会資料」", "3町合同の意見交換会",
     "策定の状況・第1次概算の結果・月額を動かす前提・地域課題15件・"
     "3町共通の確認事項・認知症施策推進計画・10月以降の工程"),
    ("別冊「町別データシート」", "町へお示しするもの",
     "町ごとの人口・認定者数・サービスの利用状況・総合事業"),
    ("別冊「交付金 市町村分評価指標と3町の評価」", "町へお示しするもの",
     "評価指標の大項目と明細列ごとの3町の得点"),
    ("別冊「交付金評価の取りまとめと第10期への反映」", "町へお示しするもの",
     "3か年の推移、全国との比較、一部の町のみ得点している項目"),
    ("別冊「サービス見込量の算定_第2次概算」", "協議の前提",
     "第1次概算からの差分と、月額を動かす前提"),
    ("別冊「資料を受領できない場合のサービス見込量」", "協議の前提",
     "資料のご提供がない場合に確定する値と、確定しない値"),
]
for i, (a, b, c) in enumerate(SHIRYO, start=1):
    r = body(ws, r, [i, a, b, c], height=32, align={1: "center"})
r = note(ws, r,
         "注）上記は「3町ヒアリング・打合せ資料一式」として束ねたものに"
         "収めています。索引に、どれをどの場面で開くかを1行ずつ付しています。", 4)
r += 1
r = lead(ws, r, "3　町別データシートをお示しする際の留意", span=4)
r = note(ws, r,
         "・**町別データシートの数値は計画作成支援ツールによるもので、"
         "3町を足しても広域連合の実績（介護保険事業状況報告）の値になりません。**"
         "区域外の施設を利用する当連合の被保険者と、"
         "区域内の施設を利用する他の保険者の被保険者の扱いが違うためです"
         "（施設・居住系の月平均利用者数は3町計527人に対し保険者単位480.17人）。\n"
         "・**サービス見込量・給付費・保険料は保険者（広域連合）を単位として"
         "算定するもので、町別には算定しません。**"
         "町別に掲げるのは実績（人口・認定者数・給付費・事業所・総合事業・"
         "交付金の得点）です。\n"
         "・将来推計を町別に掲げるかは確認事項No.152のご判断によります"
         "（令和8年9月10日のご了承では、町別は現況のみを計画本文に加え、"
         "将来推計は3町計のままとしています）。", 4, height=130, fill=IN_Y)
r += 1
r = lead(ws, r, "4　ヒアリングの場で踏み込まないこと", span=4)
r = note(ws, r,
         "・**保険料の額は3町計の概算までをお示しします。**"
         "町別への按分は広域連合の規約によるものであり、当方に案はありません。\n"
         "・**他の団体の名は協議の場に限ります。**"
         "交付金の比較に用いている団体名は計画本文には掲げません。\n"
         "・得点が0である項目について、"
         "「取組がない」のか「要件を満たさない」のか「報告していない」のかは"
         "公表資料から判別できません。評価調書によるご確認をお願いします。\n"
         "・第10期の施策に位置づけるかどうかは取組の必要性から判断するもので、"
         "交付金の得点のみを理由とはしません。", 4, height=110)

# ============================================================ 06 自己点検
_no_set = [x[0] for x in KS.CHECK]
chk(1, "確認事項の台帳に欠番・重複がないこと",
    "No.の最小から最大まで連続し、重複がないこと",
    "No.%d〜No.%d／%d件／重複 %s"
    % (min(_no_set), max(_no_set), len(_no_set),
       "あり" if len(_no_set) != len(set(_no_set)) else "なし"),
    len(_no_set) == len(set(_no_set))
    and sorted(_no_set) == list(range(min(_no_set), max(_no_set) + 1)))

chk(2, "未決の全件に決まらない場合の扱い（仮置き）があること",
    "未決の確認事項 ⊆ 既定値",
    "未決%d件／仮置きのないもの %s"
    % (len(SC), [x[0] for _p, _u, x in SC if x[0] not in KITEI] or "なし"),
    all(x[0] in KITEI for _p, _u, x in SC))

chk(3, shn("kimeru") + "と" + shn("machi")
    + "で同じ確認事項を二重に掲げていないこと",
    "確認先に町を含まないものと含むものの積集合",
    "重なり %s"
    % (sorted(set(x[0] for _p, _u, x in KIMERU)
              & set(x[0] for _p, _u, x in MACHI_KO)) or "なし"),
    not (set(x[0] for _p, _u, x in KIMERU)
         & set(x[0] for _p, _u, x in MACHI_KO)))

chk(4, shn("kimeru") + "が点の高い順に並んでいること",
    "隣り合う行の点を比べる",
    "%d件／降順 %s" % (len(KIMERU),
                      "適合" if all(KIMERU[i][0] >= KIMERU[i + 1][0]
                                    for i in range(len(KIMERU) - 1))
                      else "不適合"),
    all(KIMERU[i][0] >= KIMERU[i + 1][0] for i in range(len(KIMERU) - 1)))

chk(5, shn("kimeru") + "の絞り込みの条件を満たしていること",
    "確認先に町を含まず、点が%d点以上又はA群・B群に挙げたもの" % TEN_MIN,
    "%d件（うち点が%d点未満のもの %s）"
    % (len(KIMERU), TEN_MIN,
       sorted(x[0] for p, _u, x in KIMERU if p < TEN_MIN) or "なし"),
    all(not _is_machi(x) and (p >= TEN_MIN or x[0] in AB_NO)
        for p, _u, x in KIMERU))

chk(6, "町別に伺う事項が3町のいずれかに割り当てられていること",
    "確認先に町名を含むもののうち3町共通でないもの",
    "%d件（東川%d・美瑛%d・東神楽%d）"
    % (len(MACHI_BETSU),
       sum(1 for _p, _u, x in MACHI_BETSU if "東川" in str(x[6])),
       sum(1 for _p, _u, x in MACHI_BETSU if "美瑛" in str(x[6])),
       sum(1 for _p, _u, x in MACHI_BETSU if "東神楽" in str(x[6]))),
    len(MACHI_BETSU) == sum(
        1 for _p, _u, x in MACHI_BETSU
        if any(m in str(x[6]) for m in ("東川", "美瑛", "東神楽"))))

_a = sum(1 for v in KUBUN.values() if v == "A")
chk(7, "受託者の側で確定できる据え置き（区分A）が残っていないこと",
    "据え置きの区分がAのもの", "%d件" % _a, _a == 0)

_bc = sum(1 for v in KUBUN.values() if v in ("B", "C"))
chk(8, shn("lever") + "の件数が据え置きの区分B・Cの件数と一致すること",
    "区分B＋区分C ＝ " + shn("lever") + "の行数",
    "%d件／%s %d件" % (_bc, shn("lever"), len(_rows)), _bc == len(_rows))

# ---- 最低限の確認事項（A群・B群）の点検
_sc_no = set(x[0] for _p, _u, x in SC)
_ab_gai = sorted(n for n in AB_NO if n not in _sc_no)
chk(8.1, "A群・B群に挙げた確認事項がいずれも未決であること",
    "A群・B群のNo. ⊆ 未決の確認事項",
    "A群%d件・B群%d件／未決でないもの %s"
    % (len(_A_NO), len(_B_NO), _ab_gai or "なし"),
    not _ab_gai)

_ab_dup = sorted(n for n in set(AB_NO) if AB_NO.count(n) > 1)
chk(8.2, "A群とB群で同じ確認事項を二重に挙げていないこと",
    "A群のNo. ∩ B群のNo.", "重なり %s" % (_ab_dup or "なし"), not _ab_dup)

_ab_kitei = sorted(n for n in set(AB_NO) if n not in KITEI)
chk(8.3, "A群・B群の全件に決まらない場合の扱い（仮置き）があること",
    "A群・B群のNo. ⊆ 既定値",
    "実数%d件／仮置きのないもの %s"
    % (len(set(AB_NO)), _ab_kitei or "なし"), not _ab_kitei)

_kan_ng = [a[0] for a in AGENDA_A
           if sum(1 for k, _n, _d in KANTEN if k in a[4]) < 2]
chk(8.4, "A群が4つの観点のうち2つ以上に当たること",
    "観点の記号を数える",
    "%d件／2つ未満のもの %s" % (len(AGENDA_A), _kan_ng or "なし"),
    not _kan_ng)

_w90 = sum(w[3] for w in WARI if w[0] == 90)
_w120 = sum(w[3] for w in WARI if w[0] == 120)
chk(8.5, "時間配分の案の合計が枠に収まっていること",
    "90分の案・120分の案のそれぞれの合計",
    "90分の案 %d分／120分の案 %d分" % (_w90, _w120),
    _w90 == 90 and _w120 == 120)

chk(8.6, "A群の所要の合計が90分の案に組み込まれた分と一致すること",
    "A群の分の和 ＝ 90分の案のA群の行の分",
    "A群 %d分／90分の案 %d分"
    % (A_MIN, [w[3] for w in WARI if w[0] == 90 and "A群" in w[2]][0]),
    A_MIN == [w[3] for w in WARI if w[0] == 90 and "A群" in w[2]][0])

_a_in = sorted(n for n in _A_NO
               if n not in set(x[0] for _p, _u, x in KIMERU)
               and n not in set(x[0] for _p, _u, x in MACHI_KO))
chk(8.7, "A群の確認事項がいずれかの一覧（広域連合で決まるもの又は"
         "3町へ伺うもの）に掲げられていること",
    "A群のNo. ⊆ %s ∪ %s" % (shn("kimeru"), shn("machi")),
    "A群%d件／どちらにもないもの %s" % (len(_A_NO), _a_in or "なし"),
    not _a_in)

_g = round(G["月額"] / 100) * 100
chk(9, "算定上の月額基準額が百円未満四捨五入で6,400円になること",
    "算定上の月額 %.2f円 → 基準額" % G["月額"],
    "%s円" % num(_g), _g == 6400)

chk(10, "算定上の月額が算定の値と一致すること",
    "保険料収納必要額 ÷ 予定収納率 ÷ 補正後被保険者数 ÷ 12",
    "%.4f円（算定 %.4f円）"
    % (G["J"] / 0.99 / G["③"] / 12, G["月額"]),
    abs(G["J"] / 0.99 / G["③"] / 12 - G["月額"]) < 0.01)

chk(11, "総給付費が3か年の和になっていること",
    "令和9・10・11年度の和",
    "%s円" % num(sum(S["KYUFU3"])), len(S["KYUFU3"]) == 3)

chk(12, "月額への効きとして総額を拾っていないこと",
    "1件の効きが%s円を超えないこと" % "{:,}".format(KS.GETSU_MAX),
    "最大%s円"
    % "{:,}".format(max([v[0] for v in KS.GETSUGAKU.values()] or [0])),
    all(v[0] <= KS.GETSU_MAX for v in KS.GETSUGAKU.values()))

# ── 発注者へお送りするものであるため、禁止表現・内部の仕組みの語・
#    個人情報の形・他団体名を走査する。
NG_WORDS = ["に由来する", "と整合する", "1件も", "有意差がないため関係がない",
            "全国トップ級"]
NAIBU = ["runpy", ".py", "固定値", "実物から", "再実行", "リポジトリ",
         "書き写し", "終了コード", "スクリプト", "読んで判定", "章節ごとに"]
PI = [re.compile(r"0\d{1,4}[-(]\d{1,4}[-)]\d{3,4}"),
      re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+")]
OK_CH = re.compile(r"[ぁ-んァ-ヴ一-龥々〆ー０-９0-9A-Za-z"
                   r"、。・（）「」『』【】〔〕［］〜％　\s\n＋▲"
                   r"①-⑳Ⅰ-Ⅻ.,:：;／＼/\-－—―…‐※"
                   r"＝=＜＞<>＊*#&°±×÷｜|～〇○●◇◆□■△▲★☆"
                   r"，．！？＆＃＄＠_~\"'`!?@\(\)\[\]]")


def _cells():
    for s_ in wb.worksheets:
        for row in s_.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    yield s_.title, c.coordinate, c.value


_ng = [(t, co, w) for t, co, v in _cells()
       for w in NG_WORDS if w in v and t != "06_自己点検"]
chk(13, "禁止表現が本資料にないこと", "禁止表現5件を全てのセルで探す",
    "%d件%s" % (len(_ng), "" if not _ng else "　" + str(_ng[:3])), not _ng)

_nb = [(t, co, w) for t, co, v in _cells()
       for w in NAIBU if w in v and t != "06_自己点検"]
chk(14, "受託者の内部の仕組み・作業経過の語がないこと",
    "%d語を全てのセルで探す" % len(NAIBU),
    "%d件%s" % (len(_nb), "" if not _nb else "　" + str(_nb[:3])), not _nb)

_pi = [(t, co) for t, co, v in _cells() for p in PI if p.search(v)]
chk(15, "個人情報の形（電話番号・メールアドレス）がないこと",
    "値の形で探す", "%d件" % len(_pi), not _pi)

# 他の団体の名は収録している一覧から作る（書き写すと増減に追随しない）
_JI = ("大雪地区広域連合", "東川町", "美瑛町", "東神楽町")
_HIKAKU = {"北塩原村", "浜田地区広域行政組合", "川崎町", "金ヶ崎町", "小野町"}
try:
    import data_kofukin_rengo as DKR
    _names = ({r[1] for r in DKR.RENGO} | {r[1] for r in DKR.KAISAN}
              | {m[0] for r in DKR.RENGO for m in r[3]}
              | {m[0] for r in DKR.KAISAN for m in r[4]})
    _names = (_names | _HIKAKU) - set(_JI)
    _dan = [(t, co, n) for t, co, v in _cells() for n in _names if n in v]
except Exception:
    _dan = None
chk(16, "他の団体の固有名称がないこと",
    "全国集計から読んだ団体名%s件を全てのセルで探す"
    % ("{:,}".format(len(_names)) if _dan is not None else "―"),
    ("%d件%s" % (len(_dan), "" if not _dan else "　" + str(_dan[:3])))
    if _dan is not None else "未実施（団体名を読めない）",
    _dan is not None and not _dan)

_ast = [(t, co) for t, co, v in _cells() if "**" in v]
chk(17, "強調の記号が本文に残っていないこと",
    "** を全てのセルで探す（書き出しの時点で落としている）",
    "%d件" % len(_ast), not _ast)

_out = [(t, co, ch) for t, co, v in _cells() for ch in v
        if not OK_CH.match(ch) and t != "06_自己点検"]
chk(18, "許容する文字の外の文字がないこと",
    "日本語・英数字・定めた記号以外の文字を探す",
    "%d件%s" % (len(_out), "" if not _out
                else "　" + str(sorted({c for _t, _c, c in _out})[:8])),
    not _out)

ws = sheet(SH["chk"] + "_自己点検", "自己点検",
           "本資料の数値と記述が、台帳・既定値・算定から再現できることを"
           "確かめたものです。1件でも不適合があると作成を止めます。",
           [6, 34, 38, 38, 10])
r = header(ws, 4, ["No.", "確かめたこと", "式・条件", "結果", "判定"])
for no, naiyo, shiki, kekka, han in CHECKS:
    r = body(ws, r, [no, naiyo, shiki, kekka, han],
             {5: OK_G if han == "適合" else NG_O},
             height=36, align={1: "center", 5: "center"})

wb.save(OUT)
NG = [c for c in CHECKS if c[4] != "適合"]
print("保存 %s" % OUT)
print("シート %d" % len(wb.sheetnames))
print("最低限 A群 %d件・%d分／B群 %d件・%d分／C群 %d件"
      "（時間配分 90分の案 %d分・120分の案 %d分）"
      % (len(AGENDA_A), A_MIN, len(AGENDA_B), B_MIN,
         len(SC) - len(set(AB_NO)), _w90, _w120))
print("事前に決めていただきたいこと %d件（点%d以上又はA群・B群）／"
      "3町へ伺うこと %d件（町別%d・共通%d）／"
      "町へお願いする資料 %d件／月額を動かす前提 %d件"
      % (len(KIMERU), TEN_MIN, len(MACHI_KO), len(MACHI_BETSU),
         len(MACHI_KYO), len(LACK_MACHI), len(_rows)))
print("算定上の月額 %.2f円／保険料基準額 %s円" % (G["月額"], num(_g)))
print("自己点検 %d件：適合%d件・不適合%d件"
      % (len(CHECKS), len(CHECKS) - len(NG), len(NG)))
if NG:
    for c in NG:
        print("  不適合 No.%s %s → %s" % (c[0], c[1], c[3]))
    sys.exit(1)
print("すべての点検に適合しました。")
