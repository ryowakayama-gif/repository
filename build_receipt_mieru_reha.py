# -*- coding: utf-8 -*-
"""見える化システム 個別指標データの受領点検（福祉用具・住宅改修とリハビリテーション）.

令和8年9月10日のご指示
  「介護給付適正化における住宅改修等の点検および福祉用具購入・貸与調査の
    取組促進について　計画素案に修正して頂いている為、
    見えるかのデータを送付したところであります。
    4.受領点検の部分について、上記観点から確認をお願いします。」

このご指示により、受領した99ファイルを
介護保険最新情報Vol.1521「介護給付適正化における住宅改修等の点検および
福祉用具購入・貸与調査の取組促進について」（令和8年7月8日）の観点から点検する。

観点
  ① 福祉用具貸与調査の対象をどの種目に置くか
  ② 住宅改修の点検にリハビリテーション専門職等をどう関与させるか
  ③ 交付金の評価指標のうち3町が取り切れていない2項目
     （福祉用具貸与後の点検8点、福祉用具購入費・住宅改修費の検討8点）を
     取得できる資源が区域内にあるか

シート構成
  00_受領資料の概要
  01_福祉用具貸与8種目の状況
  02_リハビリテーション専門職の状況
  03_リハビリテーション関係の加算等の状況
  04_交付金評価との対応
  05_成果品への反映
  06_収録できなかったものと限界

出力
  output/第10期計画_見える化データの受領点検_福祉用具住宅改修とリハ.xlsx
"""

import os
from collections import Counter

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import data_mieruka_reha as R
import data_kofukin_detail as KD

ODIR = "/home/user/repository/output"
OUT = os.path.join(
    ODIR, "第10期計画_見える化データの受領点検_福祉用具住宅改修とリハ.xlsx")

KIJUNBI = "令和8年9月10日"
TOWNS = ["東川町", "美瑛町", "東神楽町"]

FONT = "游ゴシック"
NAVY, HEAD = "1F4E78", "5B9BD5"
IN_Y, OK_G, NG_O, MID_B, GRAY = "FFF2CC", "E2EFDA", "FCE4D6", "DEEBF7", "F2F2F2"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()
wb.remove(wb.active)


# ================================================================ 書式
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
    ws.row_dimensions[2].height = 60
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = freeze
    return ws


def header(ws, row, cols, height=32):
    for i, h in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=HEAD)
        c.alignment = Alignment(wrap_text=True, horizontal="center",
                                vertical="center")
        c.border = BORDER
    ws.row_dimensions[row].height = height
    return row + 1


def body(ws, row, vals, fills=None, height=26, align=None, bold=False):
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


def src(ws, row, text, span=8):
    c = ws.cell(row=row, column=1, value="資料：" + text)
    c.font = Font(name=FONT, size=8.5, color="595959")
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = 16
    return row + 1


def note(ws, row, text, span=8):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=9, italic=True)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = max(28, 13 * (len(text) // (span * 14) + 1))
    return row + 2


# ================================================================ 基礎数値
def pick(code, who, kw="", last=True):
    """指標コードの値を返す。kw を含む系列を選ぶ。last のとき最終年度の値。"""
    v = R.VALUE.get(code)
    if not v:
        return None, None
    rows = v["値"].get(who, [])
    cand = [r for r in rows if kw in r[0] and "１万対" in r[0]] or \
           [r for r in rows if kw in r[0]]
    if not cand:
        return None, None
    vals = cand[0][1]
    idx = None
    for i in range(len(vals) - 1, -1, -1):
        if vals[i] is not None:
            idx = i
            break
    if idx is None:
        return None, None
    return (vals[idx], v["年度"][idx]) if last else (vals, v["年度"])


def ratio(a, b):
    return None if not a or not b else round(a / b, 2)


# 福祉用具貸与8種目
YOGU = [
    ("N18-a", "車椅子"), ("N18-b", "特殊寝台"), ("N18-c", "体位変換器"),
    ("N18-d", "手すり"), ("N18-e", "スロープ"), ("N18-f", "歩行器"),
    ("N18-g", "歩行補助つえ"), ("N18-h", "移動用リフト"),
]
# 令和6年4月から貸与と販売の選択制の対象となった種目
SENTAKU = {"スロープ", "歩行器", "歩行補助つえ"}

# リハビリテーション専門職
REHA_SHOKU = [("M1-aa", "理学療法士"), ("M1-bb", "作業療法士"),
              ("M1-cc", "言語聴覚士")]

# リハビリテーション関係の加算等
REHA_KASAN = [
    ("N5", "リハビリテーションマネジメント加算Ⅱ以上", "合計"),
    ("N6", "生活機能向上連携加算", ""),
    ("N17", "入所前後訪問指導加算", ""),
    ("N1", "短期集中（個別）リハビリテーション実施加算", "合計"),
    ("N2", "認知症短期集中リハビリテーション実施加算", "合計"),
    ("N3", "個別リハビリテーション実施加算", ""),
    ("N7", "通所リハビリテーション（短時間）", ""),
    ("N14", "生活行為向上リハビリテーション実施加算", ""),
    ("K3-dd", "訪問リハビリテーション事業所数", ""),
    ("K3-hh", "通所リハビリテーション事業所数", ""),
    ("N11", "リハビリテーションマネジメント加算Ⅱ以上の事業所数", "合計"),
]


def _kd(daikomoku, nendo="令和8年度"):
    return [x for x in KD.DETAIL[nendo] if x[3] == daikomoku]


# ============================================================ 00
INV = R.INVENTORY
n_all = len(INV)
n_ok = sum(1 for t in INV if t[6])
n_ng = n_all - n_ok
ser = Counter(t[0].split("-")[0][0] for t in INV)

ws = sheet("00_受領資料の概要",
           "見える化システム 個別指標データの受領点検",
           "令和8年9月10日に、地域包括ケア「見える化」システムの"
           "個別指標の時系列データを2ファイル（zip）・計%d件受領しました。"
           "介護保険最新情報Vol.1521「介護給付適正化における住宅改修等の"
           "点検および福祉用具購入・貸与調査の取組促進について」を"
           "計画素案に反映したことを受けてご送付いただいたものと理解し、"
           "この観点から点検します。" % n_all,
           [4, 22, 40, 30, 24, 20])

r = 4
r = lead(ws, r, "【受領資料の内訳】", 6)
r = header(ws, r, ["系列", "内容", "件数", "値のある件数", "備考", ""])
SER_NAME = {
    "N": "リハビリテーション関係の加算算定者数・事業所数、福祉用具貸与の種目別",
    "K": "サービス提供事業所数",
    "M": "従事者数（職種別・サービス別）",
    "D": "受給率・利用率・給付月額",
    "O": "医療連携に関する加算の算定者数",
    "L": "在宅医療（訪問診療・往診・訪問歯科・訪問薬剤管理指導）",
    "G": "医師数",
}
for s in ["N", "K", "M", "D", "O", "L", "G"]:
    sub = [t for t in INV if t[0].split("-")[0][0] == s]
    ok = sum(1 for t in sub if t[6])
    bikou = ("すべて値なし。当区域の値が出力されない" if ok == 0 else
             "―")
    r = body(ws, r, [s, SER_NAME[s], len(sub), ok, bikou, ""],
             fills={4: (NG_O if ok == 0 else OK_G)},
             height=32, align={1: "center"})
r = body(ws, r, ["計", "", n_all, n_ok, "値なしは%d件" % n_ng, ""],
         bold=True, height=22, align={1: "center"})
r = src(ws, r, "地域包括ケア「見える化」システム（厚生労働省）"
        "出力日 令和8年9月10日", 6)

r += 1
r = lead(ws, r, "【この点検で確かめること】", 6)
for a, b in [
    ("① 福祉用具貸与調査の対象",
     "どの種目に着目して調査の対象を絞り込むか。"
     "8種目それぞれの算定者数（認定者1万対）を全国・北海道と比べます（01）。"),
    ("② リハビリテーション専門職の関与の可能性",
     "住宅改修の点検及び福祉用具購入・貸与調査に"
     "リハビリテーション専門職等を関与させられるか。"
     "区域内の理学療法士・作業療法士・言語聴覚士の配置状況を見ます（02）。"),
    ("③ 交付金評価の2項目を取得できるか",
     "「福祉用具貸与後のリハビリ専門職等による点検」（8点）と"
     "「福祉用具購入費・住宅改修費のリハビリ専門職等の検討」（8点）は"
     "3町平均で2.7点・5.3点にとどまります。"
     "取得の前提となる資源が区域内にあるかを確かめます（04）。"),
]:
    r = body(ws, r, [None, a, b, None, None, None], height=46)

r += 1
r = note(ws, r,
         "【結論】区域内のリハビリテーション専門職の配置と"
         "リハビリテーション関係の加算の算定は、"
         "いずれも全国・北海道を大きく上回っています。"
         "一方で福祉用具貸与は8種目中5種目が全国の45〜63％にとどまり、"
         "交付金評価のリハビリ専門職の関与2項目も取り切れていません。"
         "資源はあるが点検の仕組みに結びついていない、という状態です。"
         "Vol.1521が求める取組促進は、当区域では新たな人材の確保ではなく、"
         "既に区域内にいるリハビリテーション専門職を"
         "点検の仕組みに組み込むことで達成できると考えられます。", 6)
r = note(ws, r,
         "【収録の範囲】個別指標は町別の値が出力されず、すべて「−」です。"
         "本表に収める値は保険者（大雪地区広域連合）・北海道・全国の"
         "集計値のみです。個人情報は含まれません。", 6)


# ============================================================ 01
ws = sheet("01_福祉用具貸与8種目の状況",
           "福祉用具貸与の種目別の算定者数（認定者1万対）",
           "Vol.1521は福祉用具購入・貸与調査の取組促進を求めています。"
           "調査の対象を絞り込むため、8種目それぞれの算定者数を"
           "全国・北海道と比べます。"
           "適正化計画指針は、福祉用具貸与調査についても"
           "「認定調査状況と利用サービス不一致一覧表」により"
           "対象を絞り込むこととしています。",
           [4, 16, 12, 12, 12, 10, 10, 34])

r = 4
_, yrs_y = pick("N18-a", "大雪地区広域連合", last=False)
r = lead(ws, r, "【種目別の算定者数（%s）】" % (yrs_y[-1] if yrs_y else ""), 8)
r = header(ws, r, ["No.", "種目", "大雪", "北海道", "全国", "全国比",
                   "北海道比", "読み取れること"])
rows_y = []
for i, (code, nm) in enumerate(YOGU, start=1):
    a, y = pick(code, "大雪地区広域連合")
    h, _ = pick(code, "北海道")
    z, _ = pick(code, "全国")
    rz, rh = ratio(a, z), ratio(a, h)
    rows_y.append((nm, a, h, z, rz, rh))
    if rz is None:
        yomi = "―"
        f = {}
    elif rz >= 1.0:
        yomi = "全国と同水準以上。調査の対象として優先度が高い"
        f = {6: NG_O}
    elif rh is not None and rh >= 1.1:
        yomi = "全国は下回るが北海道を1割以上上回る"
        f = {7: IN_Y}
    elif rz <= 0.5:
        yomi = "全国の半分以下。供給又は把握の不足が疑われる"
        f = {6: MID_B}
    else:
        yomi = "全国を下回る"
        f = {}
    if nm in SENTAKU:
        yomi += "。令和6年4月からの貸与と販売の選択制の対象種目"
    r = body(ws, r, [i, nm, a, h, z, rz, rh, yomi], fills=f, height=32,
             align={1: "center"})
r = src(ws, r, "地域包括ケア「見える化」システム N18-a〜N18-h "
        "福祉用具貸与算定者数[認定者1万対]", 8)

r += 1
r = lead(ws, r, "【受託者の見方】", 8)
_te = [x for x in rows_y if x[0] == "手すり"][0]
_ho = [x for x in rows_y if x[0] == "歩行器"][0]
_su = [x for x in rows_y if x[0] == "スロープ"][0]
r = note(ws, r,
         "8種目のうち全国と同水準以上なのは手すり（全国比%.2f）のみです。"
         "北海道と比べると手すり・スロープ・歩行器の3種目が上回ります"
         "（スロープ%.2f倍、歩行器%.2f倍）。"
         "一方、車椅子・特殊寝台・歩行補助つえは全国の45％前後にとどまります。"
         % (_te[4], _su[5], _ho[5]), 8)
r = note(ws, r,
         "令和6年4月から貸与と販売の選択制の対象となったのは"
         "固定用スロープ、歩行器（歩行車を除く）、単点杖及び多点杖の4種目です"
         "（本表のスロープ・歩行器・歩行補助つえに相当）。"
         "このうちスロープと歩行器は当区域で北海道を1〜2割上回っており、"
         "選択制の導入による貸与から販売への移行の影響が"
         "他地域より大きく出る可能性があります。"
         "計画素案 第3章第3節は特定福祉用具販売の対計画比が128.5％である"
         "ことを掲げており、この点と接続します。", 8)
r = note(ws, r,
         "【受託者案】福祉用具貸与調査の対象は、"
         "①手すり（全国と同水準で件数が最も多い）、"
         "②スロープ・歩行器（選択制の対象で北海道を上回る）"
         "の順に優先することを提案します。"
         "ケアプラン点検実施要領 第4節の抽出源"
         "（認定調査状況と利用サービス不一致一覧表）と組み合わせます。", 8)
r = note(ws, r,
         "【この表の限界】N18系列は平成27年度から%sまでの収録で、"
         "令和2年度以降の値がありません。"
         "選択制が導入された令和6年度の影響は本表では確認できません。"
         "令和6年度以降の種目別の状況は、"
         "国保連の給付実績の帳票により把握する必要があります。"
         % (yrs_y[-1] if yrs_y else ""), 8)


# ============================================================ 02
ws = sheet("02_リハビリテーション専門職の状況",
           "リハビリテーション専門職の配置状況",
           "Vol.1521及び交付金の評価指標は、福祉用具貸与後の点検並びに"
           "福祉用具購入費及び住宅改修費の申請内容の検討に"
           "リハビリテーション専門職等が関与することを求めています。"
           "区域内にその資源があるかを確かめます。",
           [4, 18, 12, 12, 12, 10, 12, 32])

r = 4
r = lead(ws, r, "【職種別の従事者数（リハビリテーションサービス）】", 8)
r = header(ws, r, ["No.", "職種", "大雪\n（認定者1万対）", "北海道", "全国",
                   "全国比", "大雪の実数", "年度・内訳"])
for i, (code, nm) in enumerate(REHA_SHOKU, start=1):
    a, y = pick(code, "大雪地区広域連合", kw="合計")
    h, _ = pick(code, "北海道", kw="合計")
    z, _ = pick(code, "全国", kw="合計")
    v = R.VALUE.get(code, {}).get("値", {}).get("大雪地区広域連合", [])
    jitsu = None
    for lab, vals in v:
        if "合計" in lab and "１万対" not in lab:
            jitsu = [x for x in vals if x is not None][-1]
    rz = ratio(a, z)
    r = body(ws, r, [i, nm, a, h, z, rz,
                     ("%d人" % jitsu) if jitsu is not None else "―",
                     "%s。介護老人保健施設及び通所リハビリテーション"
                     "（老健）の合計" % y],
             fills={6: (OK_G if rz and rz >= 1.0 else IN_Y)},
             height=32, align={1: "center", 7: "center"})
r = src(ws, r, "地域包括ケア「見える化」システム M1-aa〜M1-cc "
        "従事者数（リハビリテーションサービス）[認定者1万対]", 8)

r += 1
r = lead(ws, r, "【リハビリテーションを提供する事業所の数】", 8)
r = header(ws, r, ["No.", "区分", "大雪\n（認定者1万対）", "北海道", "全国",
                   "全国比", "年度", "読み取れること"])
for i, (code, nm) in enumerate(
        [("K3-dd", "訪問リハビリテーション事業所数"),
         ("K3-hh", "通所リハビリテーション事業所数")], start=1):
    a, y = pick(code, "大雪地区広域連合")
    h, _ = pick(code, "北海道")
    z, _ = pick(code, "全国")
    rz = ratio(a, z)
    r = body(ws, r, [i, nm, a, h, z, rz, y,
                     "認定者1万人当たりの事業所数で全国を上回る"
                     if rz and rz > 1 else "全国を下回る"],
             fills={6: (OK_G if rz and rz >= 1.0 else IN_Y)},
             height=30, align={1: "center", 7: "center"})
r = src(ws, r, "地域包括ケア「見える化」システム K3-dd・K3-hh", 8)

r += 1
r = note(ws, r,
         "理学療法士は全国の1.7倍、言語聴覚士は全国の2.4倍で、"
         "訪問リハビリテーション・通所リハビリテーションの事業所数も"
         "全国を上回ります。"
         "リハビリテーション専門職等の関与を求める取組を行うための"
         "資源は区域内にあると判断できます。", 8)
r = note(ws, r,
         "ただし実数は理学療法士12人・作業療法士4人・言語聴覚士2人であり、"
         "いずれも介護老人保健施設と通所リハビリテーション（老健）に"
         "所属する職員です。"
         "作業療法士は令和3年度の11人から令和5年度に4人へ減っており、"
         "個々の事業所の状況により変動します。"
         "関与を求める際は、特定の事業所に負担が集中しないよう"
         "3町・事業者と調整する必要があります。", 8)


# ============================================================ 03
ws = sheet("03_リハビリテーション関係の加算等の状況",
           "リハビリテーション関係の加算等の算定状況",
           "リハビリテーション専門職が実際にどう関与しているかを、"
           "加算の算定状況から確かめます。"
           "特に生活機能向上連携加算と入所前後訪問指導加算は、"
           "リハビリテーション専門職が居宅を訪問して"
           "生活環境を評価する取組であり、"
           "住宅改修の点検と直接つながります。",
           [4, 34, 12, 12, 12, 10, 10, 26])

r = 4
r = header(ws, r, ["No.", "指標", "大雪", "北海道", "全国", "全国比",
                   "年度", "備考"])
for i, (code, nm, kw) in enumerate(REHA_KASAN, start=1):
    a, y = pick(code, "大雪地区広域連合", kw=kw)
    h, _ = pick(code, "北海道", kw=kw)
    z, _ = pick(code, "全国", kw=kw)
    rz = ratio(a, z)
    f = {}
    bikou = "―"
    if rz is not None:
        if rz >= 2.0:
            f = {6: OK_G}
            bikou = "全国の2倍以上"
        elif rz >= 1.0:
            f = {6: MID_B}
            bikou = "全国を上回る"
        elif rz < 0.5:
            f = {6: NG_O}
            bikou = "全国の半分未満"
    r = body(ws, r, [i, nm, a, h, z, rz, y, bikou], fills=f, height=30,
             align={1: "center", 7: "center"})
r = src(ws, r, "地域包括ケア「見える化」システム N1〜N17・K3系列"
        "（いずれも認定者1万対）", 8)

r += 1
r = note(ws, r,
         "【住宅改修の点検との接続】入所前後訪問指導加算は、"
         "介護老人保健施設の職員が入所の前後に居宅を訪問し、"
         "退所後の生活環境を確認して指導を行うものです。"
         "当区域の算定は全国の2.1倍であり、"
         "リハビリテーション専門職が居宅の環境を評価する取組が"
         "既に日常的に行われていることを示します。"
         "住宅改修の点検に同じ職員の関与を求めることは、"
         "新たな仕組みを立ち上げるのではなく"
         "既存の取組の対象を広げるものとして構成できます。", 8)
r = note(ws, r,
         "【生活機能向上連携加算】訪問介護・通所介護等の事業所が"
         "リハビリテーション専門職と連携して"
         "生活機能の向上を目的とした計画を作成する加算です。"
         "当区域は全国の2.2倍であり、"
         "介護事業所とリハビリテーション専門職の連携の回路が"
         "既にあることを示します。"
         "福祉用具貸与後の点検はこの回路を用いることができます。", 8)


# ============================================================ 04
ws = sheet("04_交付金評価との対応",
           "交付金の評価指標のうちリハビリテーション専門職等の関与に係る2項目",
           "令和8年度の評価指標は、福祉用具貸与後の点検（8点）と"
           "福祉用具購入費・住宅改修費の申請内容の検討（8点）について、"
           "リハビリテーション専門職等が関与する仕組みがあることを"
           "評価の対象としています。"
           "3町の取得状況と、本表で確かめた資源とを対照します。",
           [4, 30, 12, 12, 12, 12, 30, 20])

r = 4
r = lead(ws, r, "【3町の取得状況（令和8年度）】", 7)
r = header(ws, r, ["No.", "評価項目", "東川町", "美瑛町", "東神楽町",
                   "全国該当率", "本表で確かめたこと", ""])
KOF = [
    ("エ", "福祉用具の貸与後に、リハビリテーション専門職等が"
     "用具の適切な利用がなされているかどうかを点検する仕組みがある"),
    ("オ", "福祉用具購入費・住宅改修費の申請内容について、"
     "リハビリテーション専門職等がその妥当性を検討する仕組みがある"),
]
_d = _kd("給付費適正化事業の取組状況")
for i, (kigo, nm) in enumerate(KOF, start=1):
    row = [x for x in _d if x[4] == kigo][0]
    haiten, zenkoku = row[5], row[6]
    tenten = dict(zip(KD.TOWNS, row[7]))
    f = {c: (OK_G if tenten.get(t) else NG_O)
         for c, t in zip([3, 4, 5], TOWNS)}
    r = body(ws, r, ["%s（%d点）" % (kigo, haiten), nm,
                     "%d点" % tenten["東川町"],
                     "%d点" % tenten["美瑛町"],
                     "%d点" % tenten["東神楽町"],
                     "%.1f％" % zenkoku,
                     "区域内に理学療法士12人・作業療法士4人・"
                     "言語聴覚士2人が配置されている（02）。"
                     "資源の不足が理由ではない", ""],
             fills=f, height=54, align={1: "center", 3: "center",
                                        4: "center", 5: "center",
                                        6: "center"})
r = src(ws, r, "厚生労働省「保険者機能強化推進交付金及び"
        "介護保険保険者努力支援交付金（市町村分）に係る全国集計結果」"
        "令和8年度", 7)

r += 1
r = lead(ws, r, "【取得のために必要なこと】", 7)
r = header(ws, r, ["No.", "事項", "内容", "根拠", "実施主体", "時期", "", ""])
YOKEN = [
    (1, "関与の仕組みを定めること",
     "評価の対象となるのは「仕組みがある」ことである。"
     "件数の多寡ではない。"
     "ケアプラン点検実施要領 第10節に手順を定めることで満たせる。",
     "令和8年度評価指標 目標Ⅱ（ⅰ）2 エ・オ",
     "広域連合", "令和9年度当初"),
    (2, "広域団体等との連携も対象になること",
     "評価指標は「関係団体や都道府県・近隣市町村による広域団体等と"
     "連携して関与する仕組みがある場合も対象に含む」としている。"
     "区域内の職員に限る必要はない。",
     "同上　留意点", "広域連合・北海道", "令和9年度当初"),
    (3, "住宅改修費は建築専門職等も対象になること",
     "住宅改修費の申請内容の検討に係る「リハビリテーション専門職等」には、"
     "建築専門職及び福祉住環境コーディネーター検定試験二級以上を含む。",
     "同上　留意点", "広域連合・3町", "令和9年度当初"),
    (4, "購入費と住宅改修費はいずれかでよいこと",
     "オは「福祉用具購入費・住宅改修費のいずれかに"
     "リハビリテーション専門職等が関与していれば評価の対象として差し支えない」"
     "とされている。",
     "同上　留意点", "広域連合", "令和9年度当初"),
    (5, "3町で仕組みを揃えること",
     "エは美瑛町のみ、オは美瑛町・東神楽町が取得している。"
     "既に取得している町の手順を3町共通の手順とすることで、"
     "新たに設計する範囲を小さくできる。",
     "全国集計結果（令和8年度）", "広域連合・3町", "令和9年度当初"),
]
for t in YOKEN:
    r = body(ws, r, list(t) + ["", ""], height=48, align={1: "center"})

r += 1
r = note(ws, r,
         "エとオを3町とも取得した場合、"
         "1町当たり16点、3町で48点の増となります。"
         "ケアプラン点検の実施（推進Ⅱ 16点＋3事業の全て6点＋"
         "有料老人ホーム等を含む点検8点＋PDCA 24点）と合わせると、"
         "給付適正化の領域で1町当たり最大70点の増となります。", 7)


# ============================================================ 05
ws = sheet("05_成果品への反映",
           "成果品への反映",
           "本表で確かめたことを、どの成果品のどこに反映するかの整理です。"
           "ご指示により、計画素案への反映は協議のうえ決定します。",
           [4, 26, 40, 30, 22, 20])

r = 4
r = header(ws, r, ["No.", "反映先", "内容", "本表の該当箇所", "扱い", "状態"])
HANEI = [
    (1, "ケアプラン点検実施要領（案）第4節",
     "福祉用具貸与調査の対象種目の優先順位"
     "（①手すり、②スロープ・歩行器）を加える。",
     "01シート", "実施要領は別の成果品であり反映可能", "ご確認待ち"),
    (2, "ケアプラン点検実施要領（案）第10節",
     "リハビリテーション専門職等の関与について、"
     "区域内の配置状況（理学療法士12人・作業療法士4人・言語聴覚士2人）と、"
     "既存の加算算定の回路（生活機能向上連携加算・入所前後訪問指導加算）を"
     "用いることを加える。",
     "02・03シート", "同上", "ご確認待ち"),
    (3, "計画素案 第2章第3節",
     "リハビリテーション専門職の配置と"
     "リハビリテーション関係の加算の算定が全国を上回ることを"
     "サービス提供体制の特徴として加える。",
     "02・03シート", "［要協議］計画素案への反映は協議のうえ",
     "ご判断待ち"),
    (4, "計画素案 第5章 基本目標5（3）",
     "福祉用具貸与調査の対象種目の考え方を"
     "給付適正化の施策の方向性に加える。",
     "01シート", "同上", "ご判断待ち"),
    (5, "計画素案 第5章 基本目標1・2",
     "リハビリテーションの提供体制が強みであることを"
     "介護予防・重度化防止の施策に接続する。",
     "02・03シート", "同上", "ご判断待ち"),
    (6, "第9期の評価（第3章第3節）",
     "特定福祉用具販売の対計画比128.5％について、"
     "選択制の対象種目（スロープ・歩行器）が"
     "当区域で北海道を上回ることを要因の候補として示す。",
     "01シート", "同上", "ご判断待ち"),
]
for t in HANEI:
    r = body(ws, r, list(t), fills={6: IN_Y}, height=56,
             align={1: "center", 6: "center"})


# ============================================================ 06
ws = sheet("06_収録できなかったものと限界",
           "収録できなかったものと本表の限界",
           "受領した%d件のうち%d件は値が出力されていません。"
           "また収録できたものにも、年度の範囲による限界があります。"
           % (n_all, n_ng),
           [4, 26, 40, 30, 22, 20])

r = 4
r = lead(ws, r, "【値が出力されなかったもの】", 6)
r = header(ws, r, ["No.", "指標", "内容", "理由（推定）", "代替手段", ""])
for i, t in enumerate([x for x in INV if not x[6]], start=1):
    r = body(ws, r, [i, t[0], t[1],
                     "国が公表する在宅医療のデータ（平成31年）は"
                     "市町村単位で当区域の値が出力されない",
                     "北海道の医療計画・地域医療構想の資料、"
                     "医療介護連携推進会議での共有", ""],
             fills={4: NG_O}, height=40, align={1: "center", 2: "center"})
r = note(ws, r,
         "L系列12件はいずれも在宅医療（訪問診療・往診・訪問歯科診療・"
         "訪問薬剤管理指導）の患者数で、平成31年の地域別データです。"
         "計画素案 第2章第3節3及び第6章第4節3は"
         "在宅医療との整合を扱っており、"
         "本来これらの値を用いたい箇所ですが、値が得られません。"
         "医療側のデータは北海道又は上川中部の医療圏単位で"
         "取得する必要があります（確認事項として管理します）。", 6)

r += 1
r = lead(ws, r, "【年度の範囲による限界】", 6)
r = header(ws, r, ["No.", "系列", "収録年度", "限界", "影響", ""])
GEN = [
    (1, "N18系列（福祉用具貸与8種目）", "平成27年度〜令和元年度",
     "令和2年度以降の値がない",
     "令和6年4月からの貸与と販売の選択制の影響を確認できない", ""),
    (2, "N17（入所前後訪問指導加算）", "平成27年度〜令和元年度",
     "同上", "直近の状況は確認できない", ""),
    (3, "N6（生活機能向上連携加算）", "平成30年度〜令和5年度",
     "令和6年度の値がない", "第9期の評価には令和6年度の値を要する", ""),
    (4, "M1系列（リハビリテーション専門職）", "平成29年度〜令和6年度",
     "令和2年度が欠測（介護サービス施設・事業所調査の中止による）",
     "推移の連続性に注意を要する", ""),
    (5, "G7（医師数）", "平成26年度・平成28年度",
     "2時点のみ。令和以降の値がない",
     "医師数は医療側の資料により補う必要がある", ""),
]
for t in GEN:
    r = body(ws, r, list(t), height=36, align={1: "center"})

r += 1
r = note(ws, r,
         "個別指標はいずれも町別の値が「−」で出力されません。"
         "本表の値はすべて保険者単位です。"
         "町別データシートには収録できません。"
         "この点は計画素案 第2章の「保険者単位の原則」と整合します。", 6)


wb.save(OUT)
print("saved: %s  sheets=%d" % (os.path.basename(OUT), len(wb.sheetnames)))
for s in wb.sheetnames:
    print("  - %s %d rows" % (s, wb[s].max_row))
print("受領 %d件／値あり %d件／値なし %d件" % (n_all, n_ok, n_ng))
