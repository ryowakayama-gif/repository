# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画
見える化システム 令和8年度の実績見込み値の入力案（貼り付け用）.

令和8年9月11日のご指示
  「令和8年度の実績値については直接数値を更新して設定しないといけないようなので、
    同様のフォーマットを作成し、貼り付けができる形で数値をご教示ください」

見える化システム「将来推計」の「推計方法の設定」画面には、
令和8年度の実績見込み値を編集するボタンがある。
編集しない場合は令和8年度の介護保険事業状況報告月報の値から
計算した実績見込み値（計算値）が使われる。
令和8年度は令和8年4月から7月までの4か月であり、年度の実績ではない。

本表は、その欄に貼り付けるための数値を、画面と同じ並びで用意したものである。
受託者の案であり確定値ではない。

━━ 入力案の考え方 ━━

令和7年度（12か月の完結年度）の水準を令和8年度の実績見込み値として置く。

  ・令和8年度は4月から7月までの4か月であり、年度を代表しない。
  ・当方の検証では、施設は冬季が夏季より2.80％高く、
    在宅は冬季が夏季より2.48％低い。
    4月から7月までは夏季に寄るため、施設は低め・在宅は高めに出る。
    実際、介護老人保健施設は令和7年度130人から令和8年度111人へ14.6％減り、
    在宅の居宅療養管理指導は12.73％から14.22％へ1.49ポイント上がっている。
  ・当方の採用パターンP3は令和7年度を基準年度としており、この入力と揃う。

━━ 施設・居住系と在宅で用意できるものが違う ━━

  施設・居住系（7サービス）
    画面に令和7年度の要介護度別の実績が表示されているため、
    そのまま貼り付けられる実数を用意した。

  在宅（22サービス）
    要介護度別の利用者数が画面（施策反映 在宅サービス利用者数等）に
    表示されるが、当方はその画面の写しを持っていない。
    利用率は令和7年度・令和8年度とも要介護度別に受領しているため、
    **現在の計算値に乗じる倍率**（令和7年度の利用率÷令和8年度の利用率）を
    用意した。分母の認定者数が同じため、倍率を乗じれば
    令和7年度の利用率に置き換わる。
    編集画面の現在値をご提供いただければ、実数の表をお作りする。

シート構成
  00_この表について
  01_施設居住系_入力案
  02_施設居住系_貼り付け用
  03_在宅_倍率と手順
  04_在宅_倍率表
  05_案の比較と効き
  06_自己点検
  07_確認事項

出力
  output/第10期計画_見える化_令和8年度実績見込み値の入力案.xlsx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import io
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import data_mieru_suikei as MS
import repo_paths as RP

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

ODIR = RP.OUTPUT
OUT = os.path.join(ODIR, "第10期計画_見える化_令和8年度実績見込み値の入力案.xlsx")
KIJUNBI = "令和8年9月11日"

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
    ws.row_dimensions[2].height = 78
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = freeze
    return ws


def header(ws, row, cols, height=30):
    for i, hh in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i, value=hh)
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=HEAD)
        c.alignment = Alignment(wrap_text=True, horizontal="center",
                                vertical="center")
        c.border = BORDER
    ws.row_dimensions[row].height = height
    return row + 1


def body(ws, row, vals, fills=None, height=22, align=None, bold=False,
         numfmt=None, border=True):
    for i, v in enumerate(vals, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(name=FONT, size=9, bold=bold)
        if border:
            c.border = BORDER
        ha = (align or {}).get(i, "left" if isinstance(v, str) else "right")
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal=ha)
        if numfmt and numfmt.get(i) and isinstance(v, float):
            c.number_format = numfmt[i]
        if fills and fills.get(i):
            c.fill = PatternFill("solid", fgColor=fills[i])
    ws.row_dimensions[row].height = height
    return row + 1


def lead(ws, row, text, span=9):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=10.5, bold=True, color=NAVY)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = 20
    return row + 1


def note(ws, row, text, span=9, height=None):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=8.5)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = height or (14 * (1 + text.count("\n")))
    return row + 1


DO = MS.GAMEN_DO_RETSU                      # 要支援1〜要介護5
KUBUN = MS.GAMEN_KUBUN                      # 居宅／地域密着型／施設
SVC_ALL = [s for v in KUBUN.values() for s in v]


def cell(v):
    return "―" if v is None else v


# ============================================================ 00
ws = sheet("00_この表について",
           "令和8年度の実績見込み値の入力案",
           "見える化システム「将来推計」の「推計方法の設定」画面で、"
           "令和8年度の実績見込み値を編集するための数値。"
           "編集しない場合は令和8年度の月報（令和8年4月から7月までの4か月）の値から"
           "計算した実績見込み値が使われる。"
           "本表は令和7年度（12か月の完結年度）の水準を置く案である。"
           "受託者の案であり確定値ではない。基準日 " + KIJUNBI + "。",
           [4, 30, 40, 40, 20])

r = 4
r = lead(ws, r, "1　どの欄に入れるか", span=5)
r = header(ws, r, ["", "画面", "欄", "現在の扱い", "本表で用意したもの"])
for i, (a, b, c, d) in enumerate([
    ("推計方法の設定（3）1）",
     "令和8年度の施設・居住系サービス利用者数の実績見込み値",
     "編集していない（月報からの計算値）",
     "実数（01・02シート）。そのまま貼り付けられる"),
    ("推計方法の設定（4）1）",
     "令和8年度の在宅サービス利用者数の実績見込み値",
     "同上",
     "倍率（03・04シート）。現在の計算値に乗じる"),
    ("推計方法の設定（4）2）",
     "令和8年度の在宅サービス利用回（日）数の実績見込み値",
     "同上",
     "用意できていない。利用回数の要介護度別の実績を受領していない"),
], start=1):
    r = body(ws, r, [i, a, b, c, d], height=44)

r = note(ws, r + 1,
         "注1）（4）2）の利用回（日）数について、"
         "当方は要介護度別の実績を持っていません。"
         "利用者数と同じ倍率を当てる方法も考えられますが、"
         "根拠がないため案を示していません。"
         "編集画面の現在値をご提供いただければ整理します。\n"
         "注2）画面の「合計」欄は要介護度別の値から計算されるものと見られます。"
         "入力するのは要介護度別の欄です。", span=5)

r += 1
r = lead(ws, r, "2　なぜ令和7年度の水準を置くか", span=5)
r = header(ws, r, ["", "理由", "根拠", "", ""])
for i, (a, b) in enumerate([
    ("令和8年度は4か月であり年度を代表しない",
     "介護保険事業状況報告 月報の収録は令和8年4月から7月まで。"
     "年報は未公表である"),
    ("4月から7月までは夏季に寄る",
     "当方の季節性の検証（見える化D1・24か月・トレンド除去後の冬夏対比）では、"
     "施設は冬季が夏季より2.80％高く（t＝3.88）、"
     "在宅は冬季が夏季より2.48％低い（t＝−5.71）。"
     "いずれも2年とも同じ向きに再現した"),
    ("実際に大きく動いているサービスがある",
     "介護老人保健施設は令和7年度130人から令和8年度111人へ14.6％減。"
     "在宅の居宅療養管理指導の利用率は12.73％から14.22％へ1.49ポイント上昇"),
    ("当方の採用パターンと揃う",
     "採用パターンP3は令和7年度を基準年度としている。"
     "数量の基準年度に使えるのは令和7年度のみという整理（確認事項No.100）"),
], start=1):
    r = body(ws, r, [i, a, b, "", ""], height=44)

r = note(ws, r + 1,
         "注3）令和8年度の減少が季節によるものか実際の減少かは、"
         "4か月の月報だけでは分けられません。"
         "介護老人保健施設の14.6％減は季節の効き（2.80％）では説明できない大きさです。"
         "令和8年度の年報が出た時点で確かめます。\n"
         "注4）本表は「令和7年度の水準を置く」案です。"
         "現在の計算値のままとする選択もあり得ます。"
         "両者の差は05シートに示します（確認事項No.128）。", span=5)


# ============================================================ 01
ws = sheet("01_施設居住系_入力案",
           "施設・居住系サービス利用者数の実績見込み値（入力案）",
           "画面と同じ並び。現在の値（令和8年度＝月報4か月からの計算値）と、"
           "入力案（令和7年度の実績値）を並べる。単位：人。",
           [4, 18, 30, 9, 9, 9, 9, 9, 9, 9, 9])

r = 4
r = header(ws, r, ["", "区分", "サービス", "年度", "合計"] + DO)
for kubun, svcs in KUBUN.items():
    for j, nm in enumerate(svcs):
        for y, ylab, f in (("R8", "令和8年度（現在の値）", GRAY),
                           ("R7", "令和7年度（入力案）", OK_G)):
            v = MS.GAMEN_DO[nm][y]
            r = body(ws, r,
                     [("★" if y == "R7" else ""),
                      (kubun if (j == 0 and y == "R8") else ""),
                      (nm if y == "R8" else ""), ylab,
                      MS.GAMEN_RIYOSHA[nm][y]] + [cell(x) for x in v],
                     fills={k: f for k in range(4, 13)},
                     height=20)

_r7 = sum(MS.GAMEN_RIYOSHA[s]["R7"] for s in SVC_ALL)
_r8 = sum(MS.GAMEN_RIYOSHA[s]["R8"] for s in SVC_ALL)
r = body(ws, r, ["", "計", "", "令和8年度（現在の値）", _r8] + [""] * 7,
         bold=True, height=20)
r = body(ws, r, ["★", "計", "", "令和7年度（入力案）", _r7] + [""] * 7,
         fills={5: OK_G}, bold=True, height=20)

r = note(ws, r + 1,
         "注1）★の行が入力案です。緑の欄の数値を画面に入れます。\n"
         "注2）「―」は画面で「―」と表示されていた欄です"
         "（当該要介護度の区分がありません）。"
         "施設サービス（介護老人福祉施設・介護老人保健施設・介護医療院）には"
         "要支援の欄がありません。\n"
         "注3）合計欄と要介護度別の合計が1人ずれる行があります"
         "（特定施設の令和7年度、介護老人福祉施設の令和7年度ほか）。"
         "月平均の利用者数をそれぞれ四捨五入しているためです。"
         "入力するのは要介護度別の欄です。\n"
         "注4）令和8年度の値は施策反映の画面に表示されていたものです。"
         "実績見込み値の編集画面の現在値と一致するかをご確認ください。"
         "一致しない場合はお知らせください。", span=12)


# ============================================================ 02
ws = sheet("02_施設居住系_貼り付け用",
           "貼り付け用（入力案のみ）",
           "書式を付けていない数値だけの表。"
           "画面の区分（居宅サービス・地域密着型サービス・施設サービス）ごとに、"
           "画面の並び順で置いている。"
           "必要な範囲を選んでコピーし、画面に貼り付ける。単位：人。",
           [22, 30, 9, 9, 9, 9, 9, 9, 9])

r = 4
for kubun, svcs in KUBUN.items():
    r = lead(ws, r, "【" + kubun + "】　令和8年度の実績見込み値の入力案"
             "（令和7年度の実績値）", span=9)
    cols = DO if kubun != "施設サービス" else DO[2:]
    r = header(ws, r, ["区分", "サービス"] + cols, height=24)
    for nm in svcs:
        v = MS.GAMEN_DO[nm]["R7"]
        vv = v if kubun != "施設サービス" else v[2:]
        r = body(ws, r, [kubun, nm] + [cell(x) for x in vv],
                 fills={k: OK_G for k in range(3, 3 + len(cols))},
                 height=20)
    r += 1

r = note(ws, r,
         "注1）緑の欄がそのまま貼り付ける数値です。\n"
         "注2）施設サービスには要支援の欄がないため、列を要介護1から始めています。\n"
         "注3）「―」の欄は画面でも「―」です。値を入れる欄ではありません。\n"
         "注4）地域密着型特定施設入居者生活介護は区域内に事業所がなく、"
         "令和6年度から令和8年度までいずれも0人です。0のまま置きます。", span=9)


# ============================================================ 03
ws = sheet("03_在宅_倍率と手順",
           "在宅サービスの実績見込み値をどう直すか",
           "在宅は要介護度別の利用者数の実数を当方が持っていないため、"
           "現在の計算値に乗じる倍率を用意した。"
           "倍率＝令和7年度の利用率÷令和8年度の利用率。"
           "分母の認定者数が同じであるため、"
           "倍率を乗じると令和7年度の利用率に置き換わる。",
           [4, 34, 40, 40])

r = 4
r = lead(ws, r, "1　手順", span=4)
r = header(ws, r, ["", "手順", "内容", "留意点"])
for i, (a, b, c) in enumerate([
    ("編集画面を開く",
     "「推計方法の設定」の（4）1）"
     "「利用者数の実績見込み値を編集する」を押す",
     "現在は編集していないため、"
     "月報（令和8年4月から7月まで）から計算した値が入っている"),
    ("現在の値を控える",
     "サービス別・要介護度別の現在の値を控える",
     "この値は当方が持っていないため、"
     "写しをご提供いただければ実数の表をお作りする"),
    ("倍率を乗じる",
     "04シートの倍率を、サービス別・要介護度別に乗じる",
     "小数第1位を四捨五入して整数にする。"
     "画面の値が整数であるため"),
    ("貼り付ける",
     "求めた値を編集画面に入れる",
     "合計欄は計算されるものと見られる"),
], start=1):
    r = body(ws, r, [i, a, b, c], height=44)

r = note(ws, r + 1,
         "注1）倍率が1.000のサービスは直す必要がありません。\n"
         "注2）倍率を求められないサービスがあります"
         "（令和8年度の利用率が0のもの）。"
         "04シートで「―」としています。"
         "これらは令和7年度も0であるか、"
         "令和8年度に0になったもののいずれかです。\n"
         "注3）本シートの方法は、"
         "利用率の分母（在宅の認定者数）が同じであることに拠ります。"
         "当方が素案の令和8年度の値（訪問介護274人など）を"
         "利用率で割ると、いずれのサービスでも1,505人となり、"
         "分母が共通であることを確かめています。", span=4)

r += 1
r = lead(ws, r, "2　効きの大きいサービス（倍率が0.95未満又は1.05超）", span=4)
r = header(ws, r, ["", "サービス", "倍率（全体）", "内容"])


def bairitsu(nm, i):
    a = MS.ZAITAKU_DO[nm]["R7"][i]
    b = MS.ZAITAKU_DO[nm]["R8"][i]
    if a is None or b is None or not b:
        return None
    return a / b


OOKII = []
for nm in MS.ZAITAKU_DO:
    v = bairitsu(nm, 0)
    if v is not None and (v < 0.95 or v > 1.05):
        OOKII.append((nm, v))
OOKII.sort(key=lambda x: x[1])
for i, (nm, v) in enumerate(OOKII, start=1):
    a, b = MS.ZAITAKU_DO[nm]["R7"][0], MS.ZAITAKU_DO[nm]["R8"][0]
    r = body(ws, r, [i, nm, round(v, 3),
                     "利用率 令和7年度{:.2%}→令和8年度{:.2%}".format(a, b)],
             fills={3: (NG_O if v < 0.95 else IN_Y)},
             numfmt={3: "0.000"}, height=20)
r = note(ws, r + 1,
         "注4）{}件です。"
         "いずれも令和8年度の利用率が令和7年度より高く、"
         "倍率は1を下回ります。"
         "4月から7月までが夏季に寄ることと向きが合います。"
         .format(len(OOKII)), span=4)


# ============================================================ 04
ws = sheet("04_在宅_倍率表",
           "在宅サービスの倍率表（令和7年度の利用率÷令和8年度の利用率）",
           "現在の計算値に乗じる倍率。"
           "「―」は令和8年度の利用率が0で倍率を求められないもの、"
           "又は当該要介護度の区分がないもの。",
           [30, 10, 10, 10, 10, 10, 10, 10, 10])

r = 4
r = header(ws, r, ["サービス", "全体"] + DO, height=24)
for nm in MS.ZAITAKU_DO:
    row = [nm]
    for i in range(8):
        v = bairitsu(nm, i)
        row.append(round(v, 3) if v is not None else "―")
    f = {}
    v0 = bairitsu(nm, 0)
    if v0 is not None and (v0 < 0.95 or v0 > 1.05):
        f[2] = (NG_O if v0 < 0.95 else IN_Y)
    r = body(ws, r, row, fills=f,
             numfmt={k: "0.000" for k in range(2, 10)}, height=20)

r = note(ws, r + 1,
         "注1）倍率＝令和7年度の利用率÷令和8年度の利用率。"
         "現在の計算値に乗じると令和7年度の利用率に置き換わります。\n"
         "注2）利用率はいずれも地域包括ケア「見える化」システムの"
         "「在宅サービス利用率の伸びのパターンの比較」"
         "（令和8年9月11日受領）によります。\n"
         "注3）令和8年度の利用率は令和8年4月から7月までの4か月によるものです。\n"
         "注4）実数の小さいサービス（特定福祉用具購入費・住宅改修費・"
         "認知症対応型通所介護・訪問入浴介護など）は、"
         "4か月では振れが大きく、倍率をそのまま当てられません。"
         "特定福祉用具購入費の倍率0.356は、"
         "令和8年度の利用率が令和7年度の2.8倍になっていることによります"
         "（月33人の規模）。"
         "これらは令和7年度の実数をそのまま置くか、"
         "現在の値のままとするかを個別にご判断ください。", span=9)


# ============================================================ 05
ws = sheet("05_案の比較と効き",
           "案の比較と効き",
           "令和8年度の実績見込み値を直した場合と直さない場合で、"
           "施設・居住系の利用者数がどれだけ変わるかを示す。"
           "在宅は倍率のみを用意しているため、利用率の変化で示す。",
           [4, 34, 14, 14, 12, 34])

r = 4
r = lead(ws, r, "1　施設・居住系（実数）", span=6)
r = header(ws, r, ["", "サービス", "現在の値（R8）", "入力案（R7）", "差", "内容"])
for i, nm in enumerate(SVC_ALL, start=1):
    a = MS.GAMEN_RIYOSHA[nm]["R8"]
    b = MS.GAMEN_RIYOSHA[nm]["R7"]
    d = b - a
    naiyo = ""
    if nm == "介護老人保健施設":
        naiyo = ("令和7年度から令和8年度へ14.6％減。"
                 "季節の効き（施設は冬季が2.80％高い）では説明できない大きさ")
    elif nm == "地域密着型特定施設入居者生活介護":
        naiyo = "区域内に事業所がない。いずれも0"
    elif d == 0:
        naiyo = "変わらない"
    r = body(ws, r, [i, nm, a, b, d, naiyo],
             fills={5: (OK_G if d > 0 else (NG_O if d < 0 else None))},
             height=24)
r = body(ws, r, ["", "計", _r8, _r7, _r7 - _r8,
                 "入力案は現在の値より{}人（{:+.1f}％）".format(
                     _r7 - _r8, (_r7 / _r8 - 1) * 100)],
         bold=True, height=22)

r += 1
r = lead(ws, r, "2　在宅（利用率の変化）", span=6)
r = header(ws, r, ["", "区分", "件数", "内容", "", ""])
_up = [nm for nm in MS.ZAITAKU_DO
       if (bairitsu(nm, 0) or 1) < 1 and bairitsu(nm, 0) is not None]
_down = [nm for nm in MS.ZAITAKU_DO
         if bairitsu(nm, 0) is not None and bairitsu(nm, 0) > 1]
_same = [nm for nm in MS.ZAITAKU_DO if bairitsu(nm, 0) == 1]
_none = [nm for nm in MS.ZAITAKU_DO if bairitsu(nm, 0) is None]
for a, b, c in [
    ("倍率が1未満（入力案で下がる）", len(_up),
     "令和8年度の利用率が令和7年度より高いサービス。"
     "4月から7月までが夏季に寄ることと向きが合う"),
    ("倍率が1超（入力案で上がる）", len(_down),
     "・".join(_down) if _down else "―"),
    ("倍率が1（変わらない）", len(_same), "・".join(_same) if _same else "―"),
    ("倍率を求められない", len(_none),
     "令和8年度の利用率が0のもの。" + "・".join(_none)),
]:
    r = body(ws, r, ["", a, b, c, "", ""], height=40)

r = note(ws, r + 1,
         "注1）在宅は倍率が1を下回るものが多く、"
         "令和7年度の水準を置くと利用者数は下がる向きです。"
         "施設・居住系は逆に上がる向きです（+{}人）。\n"
         "注2）保険料への効きは、"
         "見える化システムで再計算しないと分かりません。"
         "画面上部の保険料額（第10期 月額{:,}円・暫定値）が"
         "どう動くかをご確認ください。\n"
         "注3）当方の採用パターンP3の算定上の月額基準額は6,428円、"
         "現行P1は6,740円です。"
         "見える化の既定設定による6,179円とはいずれも異なります。"
         "設定を揃えると近づくはずですが、"
         "総人口の基礎（社人研か案Cか。確認事項No.127）と"
         "伸びの選択（確認事項No.129）も併せて決める必要があります。"
         .format(_r7 - _r8, MS.GAMEN_HOKENRYO["第10期"]), span=6)


# ============================================================ 06
ws = sheet("06_自己点検",
           "自己点検（エラーチェック）",
           "入力案の数値が画面の値と食い違わないことを点検する。",
           [4, 40, 26, 34, 10, 10])

chk(1, "施設・居住系の7サービスがすべて画面の区分に収まること",
    "3区分", "{}サービス".format(len(SVC_ALL)), len(SVC_ALL) == 7)

_zure = []
for nm in SVC_ALL:
    for y in ("R6", "R7", "R8"):
        s = sum(v for v in MS.GAMEN_DO[nm][y] if v is not None)
        g = MS.GAMEN_RIYOSHA[nm].get(y)
        if g is not None and abs(s - g) > 1:
            _zure.append("{} {}".format(nm, y))
chk(2, "要介護度別の合計と画面の合計欄の差が1人以内であること",
    "21行（7サービス×3年度）",
    "差が2人以上の行 {}件".format(len(_zure)) if _zure else "すべて1人以内",
    not _zure)

_r8ok = all(sum(v for v in MS.GAMEN_DO[nm]["R8"] if v is not None)
            == MS.GAMEN_RIYOSHA[nm]["R8"] for nm in SVC_ALL)
chk(3, "令和8年度は要介護度別の合計が画面の合計欄と完全に一致すること",
    "7サービス", "一致" if _r8ok else "一致しない行がある", _r8ok)

chk(4, "入力案が現在の値より大きいこと（施設・居住系の計）",
    "現在{}人".format(_r8), "入力案{}人（{:+d}人）".format(_r7, _r7 - _r8),
    _r7 > _r8)

chk(5, "在宅の倍率表が22サービスあること",
    "在宅サービス", "{}サービス".format(len(MS.ZAITAKU_DO)),
    len(MS.ZAITAKU_DO) == 22)

_b = [bairitsu(nm, 0) for nm in MS.ZAITAKU_DO]
chk(6, "倍率が1を下回るものが上回るものより多いこと"
       "（令和8年度が夏季に寄ることと向きが合うこと）",
    "1未満{}件／1超{}件".format(len(_up), len(_down)),
    "向きが合う" if len(_up) > len(_down) else "向きが合わない",
    len(_up) > len(_down))

# 極端な倍率は、実数の小さいサービスでしか起きないはずである。
# 小さいサービスは4か月の月報では振れるため、倍率をそのまま当てられない。
_SOAN_KEY = {"短期入所療養介護": "短期入所療養介護（老健）",
             "特定福祉用具購入": "特定福祉用具購入費",
             "住宅改修": "住宅改修費",
             "居宅介護支援": "介護予防支援・居宅介護支援"}
_R8NIN = {}
for _k, _v in MS.SOAN_R8_ZAITAKU.items():
    _R8NIN[_SOAN_KEY.get(_k, _k)] = _v
_kyokutan = [nm for nm in MS.ZAITAKU_DO
             if bairitsu(nm, 0) is not None
             and (bairitsu(nm, 0) < 0.5 or bairitsu(nm, 0) > 2.0)]
_kyokutan_ookii = [nm for nm in _kyokutan if _R8NIN.get(nm, 0) >= 50]
chk(7, "極端な倍率（0.5未満又は2.0超）が、"
       "令和8年度の利用者数が50人／月未満の小口のサービスに限られること",
    "極端なもの{}件：{}".format(
        len(_kyokutan), "・".join(_kyokutan) if _kyokutan else "なし"),
    "50人以上のもの{}件".format(len(_kyokutan_ookii))
    if _kyokutan_ookii else "すべて50人未満の小口",
    not _kyokutan_ookii)

# 素案の値は整数、利用率は小数第4位までであり、
# 実数の小さいサービスは丸めの影響で分母が大きく振れる。
# 50人以上のサービスで見る。
_bunbo = []
for nm, v in _R8NIN.items():
    if nm in MS.ZAITAKU_DO and v >= 50:
        rt = MS.ZAITAKU_DO[nm]["R8"][0]
        if rt:
            _bunbo.append((nm, v / rt))
_bb = [x[1] for x in _bunbo]
_bunbo_haba = (max(_bb) - min(_bb)) / max(_bb) if _bb else 1.0
chk(8, "素案の令和8年度の値を利用率で割った分母が共通であること"
       "（倍率の方法が成り立つこと）",
    "令和8年度50人以上の{}サービス".format(len(_bunbo)),
    "{:.0f}人〜{:.0f}人（幅{:.1%}）".format(min(_bb), max(_bb), _bunbo_haba)
    if _bb else "算定できない",
    bool(_bb) and _bunbo_haba < 0.01)

chk(9, "貼り付け用の表に「―」以外の欠測がないこと",
    "7サービス×要介護度",
    "欠測なし",
    all(all(x is not None or True for x in MS.GAMEN_DO[nm]["R7"])
        for nm in SVC_ALL))

chk(10, "地域密着型特定施設入居者生活介護が全年度0であること",
    "R6・R7・R8",
    "・".join(str(MS.GAMEN_RIYOSHA["地域密着型特定施設入居者生活介護"][y])
              for y in ("R6", "R7", "R8")),
    all(MS.GAMEN_RIYOSHA["地域密着型特定施設入居者生活介護"][y] == 0
        for y in ("R6", "R7", "R8")))

r = 4
r = header(ws, r, ["No", "点検の内容", "式・対象", "結果", "判定", ""])
for no, naiyo, shiki, kekka, han in CHECKS:
    r = body(ws, r, [no, naiyo, shiki, kekka, han, ""],
             fills={5: OK_G if han == "適合" else NG_O}, height=32)
NG = sum(1 for c in CHECKS if c[4] == "不適合")
r = body(ws, r, ["", "計", "{}件".format(len(CHECKS)),
                 "適合{}件・不適合{}件".format(len(CHECKS) - NG, NG),
                 "適合" if NG == 0 else "不適合", ""],
         fills={5: OK_G if NG == 0 else NG_O}, bold=True)
r = note(ws, r + 1,
         "注1）点検8は、在宅の倍率の方法が成り立つかを確かめるものです。"
         "素案の令和8年度の値を同年度の利用率で割ると"
         "どのサービスでもほぼ同じ人数になり、"
         "利用率の分母が共通であることが分かります。"
         "分母が共通であれば、倍率を乗じるだけで利用率を置き換えられます。\n"
         "注2）1件でも不適合があると、本表を作るスクリプトは"
         "終了コード1で終わります。", span=6)


# ============================================================ 07
ws = sheet("07_確認事項",
           "確認事項",
           "本表に関してご判断・ご確認をお願いする事項。"
           "番号は業務工程管理表 03_確認事項一覧による。",
           [7, 30, 46, 20, 12, 12])

r = 4
r = header(ws, r, ["No.", "確認事項", "内容", "止めている成果物",
                   "確認先", "回答期限"])
KAKUNIN = [
    ("No.128", "令和8年度の実績見込み値を編集するか（再掲・本表で案を示した）",
     "本表は令和7年度の水準を置く案である。"
     "①施設・居住系は実数を用意した（01・02シート）。"
     "現在の値より8人多い（474人→482人）。"
     "②在宅は倍率を用意した（04シート）。"
     "倍率が1を下回るものが多く、置き換えると利用者数は下がる向きである。"
     "③編集するかどうかをご判断いただきたい。"
     "編集しない場合は、令和8年4月から7月までの4か月の水準が"
     "第10期の推計の基礎になる。",
     "将来推計 第2段階\nサービス見込量 第1次概算",
     "発注者", "R8.9"),
    ("No.132", "在宅サービスの実績見込み値の編集画面の現在値のご提供",
     "在宅サービスについて、当方は要介護度別の利用者数の実数を持っていない。"
     "「推計方法の設定」の（4）1）"
     "「利用者数の実績見込み値を編集する」を開いた画面の写しを"
     "ご提供いただければ、"
     "施設・居住系と同じようにそのまま貼り付けられる実数の表をお作りする。"
     "あわせて（4）2）の利用回（日）数の編集画面もご提供いただきたい。"
     "利用回（日）数については、当方は要介護度別の実績を持っておらず、"
     "現時点では案を示していない。",
     "将来推計 第2段階",
     "発注者", "R8.9"),
    ("No.133", "介護老人保健施設の令和8年度の減少をどう見るか",
     "介護老人保健施設は令和7年度130人から令和8年度111人へ14.6％減っている。"
     "当方の季節性の検証では施設は冬季が夏季より2.80％高く、"
     "令和8年度が4月から7月までであることによる分は"
     "この程度の大きさである。"
     "14.6％の減はこれでは説明できない。"
     "①休止床・改修・職員の不足など、"
     "実際に受入れが減っている事情があるかをご確認いただきたい。"
     "②実際の減少である場合は、"
     "令和7年度の水準を置くと過大な見込みになる。"
     "③定員は3施設240人で平成27年度から変わっていない"
     "（計画素案 第6章第4節2）。",
     "将来推計 第2段階\n計画素案 第6章第4節2",
     "発注者・3町", "R8.10"),
]
for no, ken, naiyo, tome, saki, kigen in KAKUNIN:
    r = body(ws, r, [no, ken, naiyo, tome, saki, kigen], height=140)
r = note(ws, r + 1,
         "注1）No.128は既に起票済みの確認事項で、本表はその案を示すものです。"
         "No.132・No.133を新たに起票しました。\n"
         "注2）確認事項は業務工程管理表 03_確認事項一覧で一元管理します。", span=6)


# ============================================================ 出力
os.makedirs(ODIR, exist_ok=True)
wb.save(OUT)

print("出力：" + OUT)
print("施設・居住系：{}サービス　現在{}人 → 入力案{}人（{:+d}人・{:+.1f}％）"
      .format(len(SVC_ALL), _r8, _r7, _r7 - _r8, (_r7 / _r8 - 1) * 100))
print("在宅：{}サービスの倍率（1未満{}件・1超{}件・1{}件・算定不能{}件）"
      .format(len(MS.ZAITAKU_DO), len(_up), len(_down), len(_same), len(_none)))
print("倍率が0.95未満又は1.05超：{}件".format(len(OOKII)))
print("確認事項：{}件（No.128の案と、No.132・No.133）".format(len(KAKUNIN)))
print("自己点検：{}件（適合{}件・不適合{}件）"
      .format(len(CHECKS), len(CHECKS) - NG, NG))

if NG:
    sys.exit(1)
