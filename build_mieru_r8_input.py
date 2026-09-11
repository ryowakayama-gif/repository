# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画
見える化システム 令和8年度の実績見込み値を令和7年度に差し替える入力値.

令和8年9月11日のご指示
  「令和8年度の実績値については直接数値を更新して設定しないといけないようなので、
    同様のフォーマットを作成し、貼り付けができる形で数値をご教示ください」
  「他の保険者の試算を参考に、令和8年度数値をご教示ください。
    ご指摘の通り、R7年度実績値をそのまま置くと推計がおかしくなります」

━━ 差し替えは4画面ある ━━

当方が先にお示ししたのは②③の2画面だけであった。
他の保険者の同種の作業（弊社の別案件）では次の4画面をすべて差し替えている。
①を差し替えないと、利用率の分母が令和8年度のままになり、
利用者数だけを令和7年度に置いても令和7年度の利用率にならない。

  ① 認定者数の実績値
  ② 施設・居住系サービス利用者数
  ③ 在宅サービス利用者数
  ④ 在宅サービス利用回（日）数

━━ 出所 ━━

  ① 介護保険事業状況報告（年報・令和7年度）の要介護認定者数
     → 当方が保有している。確定値である。
  ②③ 見える化システムの総括表詳細（１）の令和7年度（年間の延べ人月）÷12
     → サービス計は確定値である。給付費の合計が年報と円単位で一致する系列。
        要介護度別の内訳は当方が保有していないため按分による推定値である。
  ④ 総括表詳細（３）の1人1月あたり利用回（日）数×②③の利用者数
     → 要支援・要介護の2区分までしか分けられない。

年報（令和7年度）の様式1の5・様式1の6・様式1の7・様式2の
要介護度別のデータをご提供いただければ、②③④も確定値にできる。

━━ 画面の令和7年度の値をそのまま貼らない理由 ━━

1 画面の令和7年度の値は整数であり、総括表の月平均と食い違う。
  特定施設入居者生活介護 画面64人／総括表61.50人（＋4.1％）、
  介護老人保健施設 画面130人／総括表134.08人（▲3.0％）。
  総括表は給付費が年報と円単位で一致する系列である。

2 整数に丸めると小口のサービスが0に落ちる。
  定期巡回・随時対応型訪問介護看護は月平均0.583人、
  認知症対応型通所介護は月平均1.000人である。
  0を入れると令和9年度から令和11年度も0のまま残る。

3 認定者数（①）を差し替えないと分母がそろわない。

シート構成
  00_この表について
  01_認定者数（①）
  02_施設居住系の利用者数（②）
  03_在宅の利用者数（③）
  04_利用回日数（④）
  05_出所の食い違いと限界
  06_自己点検
  07_確認事項

出力
  output/第10期計画_見える化_令和8年度実績見込み値の入力案.xlsx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import io
import os
import sys
from decimal import Decimal, ROUND_HALF_UP

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import data_mieru_soukatu as M
import data_mieru_suikei as MS
import data_nenpo as N
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
    ws.row_dimensions[2].height = 82
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = freeze
    return ws


def header(ws, row, cols, height=28):
    for i, hh in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i, value=hh)
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=HEAD)
        c.alignment = Alignment(wrap_text=True, horizontal="center",
                                vertical="center")
        c.border = BORDER
    ws.row_dimensions[row].height = height
    return row + 1


def body(ws, row, vals, fills=None, height=20, align=None, bold=False,
         numfmt=None):
    for i, v in enumerate(vals, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(name=FONT, size=9, bold=bold)
        c.border = BORDER
        ha = (align or {}).get(i, "left" if isinstance(v, str) else "right")
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal=ha)
        if numfmt and numfmt.get(i) and isinstance(v, float):
            c.number_format = numfmt[i]
        if fills and fills.get(i):
            c.fill = PatternFill("solid", fgColor=fills[i])
    ws.row_dimensions[row].height = height
    return row + 1


def lead(ws, row, text, span=10):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=10.5, bold=True, color=NAVY)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = 20
    return row + 1


def note(ws, row, text, span=10, height=None):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=8.5)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = height or (14 * (1 + text.count("\n")))
    return row + 1


# ============================================================ 算定
DO = MS.GAMEN_DO_RETSU                       # 要支援1〜要介護5（7区分）
SHIEN_I = (0, 1)                             # 要支援の位置
KAIGO_I = (2, 3, 4, 5, 6)                    # 要介護の位置

D1 = {l: d for l, u, d in M.SOUKATU["総括表詳細（１）"]["行"] if u == "（人）"}
D3 = {l: d for l, u, d in M.SOUKATU["総括表詳細（３）"]["行"]}


def soukatsu_m(name):
    """総括表詳細（１）の令和7年度（年間の延べ人月）を12で除した月平均。"""
    for l in D1:
        if l.split(" ", 1)[-1] == name:
            v = D1[l].get("R7", {}).get("実績値")
            return (v / 12) if v else 0.0
    return None


def r3(x, n=3):
    return float(Decimal(str(x)).quantize(Decimal("1." + "0" * n),
                                          ROUND_HALF_UP))


def to_int(vals):
    """小数の月平均を整数にする。

    要支援の群と要介護の群それぞれで合計を四捨五入し、
    小数部の大きいものから1を配って群の合計を保つ（最大剰余法）。
    単純に各欄を四捨五入すると群の合計がずれる。
    """
    out = [None] * len(vals)
    for g in (SHIEN_I, KAIGO_I):
        idx = [i for i in g if vals[i] is not None]
        if not idx:
            continue
        s = sum(vals[i] for i in idx)
        tgt = int(Decimal(str(s)).quantize(Decimal("1"), ROUND_HALF_UP))
        base = {i: int(vals[i]) for i in idx}
        need = tgt - sum(base.values())
        order = sorted(idx, key=lambda i: -(vals[i] - int(vals[i])))
        for i in idx:
            out[i] = base[i]
        k = 0
        while need > 0:
            out[order[k % len(order)]] += 1
            need -= 1
            k += 1
        while need < 0:
            for i in sorted(idx, key=lambda i: (vals[i] - int(vals[i]))):
                if out[i] > 0 and need < 0:
                    out[i] -= 1
                    need += 1
    return out


# ---------------------------------------------- ① 認定者数（年報・確定値）
NINTEI = {d: N.NINTEI[d]["R7"] for d in DO}
NINTEI_KEI = N.NINTEI["第1号被保険者　合計"]["R7"]
NINTEI_2GO = N.NINTEI["第2号被保険者　合計"]["R7"]

# ---------------------------------------------- ② 施設・居住系
SHISETSU = [
    ("居宅サービス", "特定施設入居者生活介護", "特定施設入居者生活介護"),
    ("地域密着型サービス", "認知症対応型共同生活介護", "認知症対応型共同生活介護"),
    ("地域密着型サービス", "地域密着型特定施設入居者生活介護", None),
    ("地域密着型サービス", "地域密着型介護老人福祉施設入所者生活介護",
     "地域密着型介護老人福祉施設入所者生活介護"),
    ("施設サービス", "介護老人福祉施設", "介護老人福祉施設"),
    ("施設サービス", "介護老人保健施設", "介護老人保健施設"),
    ("施設サービス", "介護医療院", "介護医療院"),
]

SHISETSU_DO = {}
SHISETSU_KEI = {}
for _ku, nm, sn in SHISETSU:
    tot = soukatsu_m(sn) if sn else 0.0
    SHISETSU_KEI[nm] = tot
    g = MS.GAMEN_DO[nm]["R7"]
    gs = sum(v for v in g if v is not None)
    SHISETSU_DO[nm] = [None if v is None else (tot * v / gs if gs else 0.0)
                       for v in g]

# ---------------------------------------------- 在宅の認定者数（分母）
ZAITAKU_NIN = []
for i, d in enumerate(DO):
    s = sum((SHISETSU_DO[nm][i] or 0.0) for _k, nm, _s in SHISETSU)
    ZAITAKU_NIN.append(NINTEI[d] - s)

# ---------------------------------------------- ③ 在宅
ZAITAKU = [
    ("訪問介護", "訪問介護"),
    ("訪問入浴介護", "訪問入浴介護"),
    ("訪問看護", "訪問看護"),
    ("訪問リハビリテーション", "訪問リハビリテーション"),
    ("居宅療養管理指導", "居宅療養管理指導"),
    ("通所介護", "通所介護"),
    ("通所リハビリテーション", "通所リハビリテーション"),
    ("短期入所生活介護", "短期入所生活介護"),
    ("短期入所療養介護（老健）", "短期入所療養介護（老健）"),
    ("福祉用具貸与", "福祉用具貸与"),
    ("特定福祉用具購入費", "特定福祉用具販売"),
    ("住宅改修費", "住宅改修"),
    ("介護予防支援・居宅介護支援", "介護予防支援・居宅介護支援"),
    ("定期巡回・随時対応型訪問介護看護", "定期巡回・随時対応型訪問介護看護"),
    ("地域密着型通所介護", "地域密着型通所介護"),
    ("認知症対応型通所介護", "認知症対応型通所介護"),
    ("小規模多機能型居宅介護", "小規模多機能型居宅介護"),
]

ZAITAKU_DO = {}
ZAITAKU_KEI = {}
DAITAI = []          # 利用率が全欄0.0％のため認定者数で按分したもの
for nm, sn in ZAITAKU:
    tot = soukatsu_m(sn)
    ZAITAKU_KEI[nm] = tot if tot is not None else 0.0
    rt = MS.ZAITAKU_DO[nm]["R7"][1:]          # 全体を除いた7区分
    w = [None if rt[i] is None else rt[i] * ZAITAKU_NIN[i] for i in range(7)]
    ws_ = sum(x for x in w if x)
    if ws_ == 0 and ZAITAKU_KEI[nm] > 0:
        # 画面の利用率は0.1％刻み（約1.5人）であるため、
        # 月平均が1.5人に満たないサービスは全欄0.0％と表示される。
        # 按分の重みが作れないので、在宅の認定者数そのもので按分する。
        DAITAI.append(nm)
        w = [None if rt[i] is None else ZAITAKU_NIN[i] for i in range(7)]
        ws_ = sum(x for x in w if x)
    ZAITAKU_DO[nm] = [None if w[i] is None else
                      (ZAITAKU_KEI[nm] * w[i] / ws_ if ws_ else 0.0)
                      for i in range(7)]

# ---------------------------------------------- ④ 利用回（日）数
KAISU = []
NASHI_GYO = []        # 総括表に要支援の行そのものがないもの
NASHI_JIS = []        # 行はあるが令和7年度の実績値がないもの
for nm, _sn in ZAITAKU:
    a = D3.get(nm + " 要支援", {}).get("R7", {}).get("実績値")
    b = D3.get(nm + " 要介護", {}).get("R7", {}).get("実績値")
    if a is None and b is None:
        continue
    if a is None:
        (NASHI_GYO if (nm + " 要支援") not in D3 else NASHI_JIS).append(nm)
    shien = sum((ZAITAKU_DO[nm][i] or 0.0) for i in SHIEN_I)
    kaigo = sum((ZAITAKU_DO[nm][i] or 0.0) for i in KAIGO_I)
    KAISU.append((nm, a, b, shien, kaigo,
                  (a or 0.0) * shien, (b or 0.0) * kaigo))


# ============================================================ 00
ws = sheet("00_この表について",
           "令和8年度の実績見込み値を令和7年度に差し替える入力値",
           "見える化システム「将来推計」の「推計方法の設定」画面で"
           "令和8年度の実績見込み値を編集する。"
           "令和8年度は令和8年4月から7月までの4か月であり年度を代表しないため、"
           "完結年度（令和7年度）の月平均に差し替える。"
           "差し替えるのは4画面である。受託者の案であり確定値ではない。"
           "基準日 " + KIJUNBI + "。",
           [4, 26, 34, 30, 24])

r = 4
r = lead(ws, r, "1　差し替える4画面と出所", span=5)
r = header(ws, r, ["", "画面", "出所", "本表で用意したもの", "確度"])
for i, (a, b, c, d) in enumerate([
    ("① 認定者数の実績値",
     "介護保険事業状況報告（年報・令和7年度）の要介護認定者数",
     "要介護度別の実数（01シート）",
     "確定値。当方が保有している"),
    ("② 施設・居住系サービス利用者数",
     "見える化システム 総括表詳細（１）の令和7年度（年間の延べ人月）÷12",
     "サービス別の月平均と要介護度別の内訳（02シート）",
     "サービス計は確定値。要介護度別は按分による推定値"),
    ("③ 在宅サービス利用者数",
     "同上",
     "サービス別の月平均と要介護度別の内訳（03シート）",
     "同上"),
    ("④ 在宅サービス利用回（日）数",
     "総括表詳細（３）の1人1月あたり利用回（日）数×③の利用者数",
     "サービス別の要支援計・要介護計（04シート）",
     "要支援・要介護の2区分まで。要介護度別には分けられない"),
], start=1):
    r = body(ws, r, [i, a, b, c, d],
             fills={5: (OK_G if "確定値。" in d else IN_Y)}, height=48)

r = note(ws, r + 1,
         "注1）当方が先にお示ししたのは②③の2画面だけでした。"
         "①を差し替えないと利用率の分母が令和8年度のままになり、"
         "利用者数だけを令和7年度に置いても令和7年度の利用率になりません。"
         "他の保険者の同種の作業（弊社の別案件）では4画面をすべて"
         "差し替えており、その方法を参照しました。"
         "数値は当広域連合のデータから算定しており、借用していません。\n"
         "注2）年報（令和7年度）の様式1の5・様式1の6・様式1の7・様式2の"
         "要介護度別のデータをご提供いただければ、"
         "②③④も確定値にできます（確認事項No.134）。", span=5)

r += 1
r = lead(ws, r, "2　画面の令和7年度の値をそのまま貼らない理由", span=5)
r = header(ws, r, ["", "理由", "内容", "", ""])
for i, (a, b) in enumerate([
    ("画面の値と総括表の月平均が食い違う",
     "画面の令和7年度は整数である。"
     "特定施設入居者生活介護は画面64人に対し総括表61.50人（＋4.1％）、"
     "介護老人保健施設は画面130人に対し総括表134.08人（▲3.0％）。"
     "総括表は給付費の合計が年報と円単位で一致する系列であり、"
     "こちらを用いる（05シート）"),
    ("整数に丸めると小口のサービスが0に落ちる",
     "定期巡回・随時対応型訪問介護看護は月平均0.583人、"
     "認知症対応型通所介護は月平均1.000人である。"
     "0を入れると令和9年度から令和11年度も0のまま残る。"
     "本表は小数版と整数版の両方を用意している"),
    ("認定者数を差し替えないと分母がそろわない",
     "見える化システムは利用率（利用者数÷認定者数）を求めて将来へ延ばす。"
     "①を令和8年度のままにすると、"
     "令和7年度の利用者数を令和8年度の認定者数で割ることになる"),
    ("要介護度別の合計が画面の合計欄と1人ずれる行がある",
     "画面の令和7年度は特定施設入居者生活介護が合計64人・"
     "要介護度別の和が63人、"
     "介護老人福祉施設が合計140人・和が141人である。"
     "月平均をそれぞれ四捨五入しているためであり、"
     "そのまま貼ると誤差が入る"),
], start=1):
    r = body(ws, r, [i, a, b, "", ""], height=52)


# ============================================================ 01
ws = sheet("01_認定者数",
           "① 認定者数の実績値（確定値）",
           "介護保険事業状況報告（年報・令和7年度）による。"
           "令和8年3月末時点の第1号被保険者の要介護（要支援）認定者数。"
           "当方が保有している確定値であり、按分や推定を含まない。",
           [6, 22, 12, 12, 12, 12, 12, 12, 40])

r = 4
r = header(ws, r, ["", "区分"] + DO + ["内容"])
r = body(ws, r, ["★", "認定者数（人）"] + [NINTEI[d] for d in DO],
         fills={k: OK_G for k in range(3, 10)}, height=22)
r = body(ws, r, ["", "計", "", "", "", "", "", "",
                 "第1号被保険者 {:,}人（要支援 {:,}人・要介護 {:,}人）。"
                 "第2号被保険者 {:,}人".format(
                     NINTEI_KEI,
                     sum(NINTEI[DO[i]] for i in SHIEN_I),
                     sum(NINTEI[DO[i]] for i in KAIGO_I),
                     NINTEI_2GO)],
         bold=True, height=22)

r = note(ws, r + 1,
         "注1）★の行が入力値です。\n"
         "注2）見える化システムの画面が性別・年齢階級別の入力を求める場合は、"
         "当方はその内訳を保有していないため用意できません。"
         "年報（様式1の5）のご提供をお願いします（確認事項No.134）。\n"
         "注3）総括表の要介護認定者数1,943人は令和7年9月末時点であり、"
         "本表の1,984人（令和8年3月末）とは時点が異なります。"
         "推計の実績値として置くのは年度末の値です。", span=9)


# ============================================================ 02
ws = sheet("02_施設居住系の利用者数",
           "② 施設・居住系サービス利用者数（令和7年度の月平均）",
           "総括表詳細（１）の令和7年度（年間の延べ人月）を12で除した。"
           "サービス計は確定値。要介護度別の内訳は"
           "画面の令和7年度の構成比による按分であり推定値である。"
           "上段が小数、下段が整数（最大剰余法で群の合計を保つ）。",
           [4, 16, 30, 10, 11, 11, 11, 11, 11, 11, 11])

r = 4
r = header(ws, r, ["", "区分", "サービス", "サービス計"] + DO)
for _ku, nm, _sn in SHISETSU:
    v = SHISETSU_DO[nm]
    iv = to_int(v)
    r = body(ws, r, ["★", _ku, nm, r3(SHISETSU_KEI[nm])]
             + [("―" if x is None else r3(x)) for x in v],
             fills={k: OK_G for k in range(4, 12)},
             numfmt={k: "0.000" for k in range(4, 12)}, height=20)
    r = body(ws, r, ["（整数）", "", "", int(round(SHISETSU_KEI[nm]))]
             + [("―" if x is None else x) for x in iv],
             fills={k: MID_B for k in range(4, 12)}, height=18)

_sk = sum(SHISETSU_KEI.values())
r = body(ws, r, ["", "計", "", r3(_sk)] + [""] * 7, bold=True,
         numfmt={4: "0.000"}, height=20)
r = note(ws, r + 1,
         "注1）★の行（小数）が入力値です。"
         "入力欄が整数しか受け付けない場合は下段の整数を用います。\n"
         "注2）施設・居住系の計は{:.3f}人です。"
         "画面の令和8年度（4か月）は474人、画面の令和7年度は482人です。\n"
         "注3）地域密着型特定施設入居者生活介護は区域内に事業所がなく0です。\n"
         "注4）要介護度別の内訳は推定値です。"
         "年報（様式1の6・様式1の7）のご提供により確定値にできます。"
         .format(_sk), span=11)


# ============================================================ 03
ws = sheet("03_在宅の利用者数",
           "③ 在宅サービス利用者数（令和7年度の月平均）",
           "総括表詳細（１）の令和7年度（年間の延べ人月）を12で除した。"
           "サービス計は確定値。要介護度別の内訳は、"
           "在宅サービス利用率（令和7年度・要介護度別）に"
           "在宅の認定者数を乗じた重みによる按分であり推定値である。",
           [4, 30, 10, 11, 11, 11, 11, 11, 11, 11])

r = 4
r = lead(ws, r, "1　在宅の認定者数（按分の分母）", span=10)
r = header(ws, r, ["", "区分", "計"] + DO)
r = body(ws, r, ["", "認定者数（年報）", NINTEI_KEI]
         + [NINTEI[d] for d in DO], height=20)
r = body(ws, r, ["", "施設・居住系の利用者数", r3(_sk)]
         + [r3(sum((SHISETSU_DO[nm][i] or 0.0)
                   for _k, nm, _s in SHISETSU)) for i in range(7)],
         numfmt={k: "0.000" for k in range(3, 11)}, height=20)
r = body(ws, r, ["", "在宅の認定者数（差引き）", r3(sum(ZAITAKU_NIN))]
         + [r3(x) for x in ZAITAKU_NIN],
         fills={k: GRAY for k in range(3, 11)},
         numfmt={k: "0.000" for k in range(3, 11)}, height=20)
r = note(ws, r + 1,
         "注1）在宅サービス利用率の分母は在宅の認定者数です。"
         "素案の令和8年度の値を同年度の利用率で割ると"
         "50人以上のサービスでいずれも1,500人前後になり、"
         "分母が共通であることを確かめています。", span=10)

r += 1
r = lead(ws, r, "2　入力値（上段が小数、下段が整数）", span=10)
r = header(ws, r, ["", "サービス", "サービス計"] + DO)
for nm, _sn in ZAITAKU:
    v = ZAITAKU_DO[nm]
    iv = to_int(v)
    r = body(ws, r, ["★", nm, r3(ZAITAKU_KEI[nm])]
             + [("―" if x is None else r3(x)) for x in v],
             fills={k: OK_G for k in range(3, 11)},
             numfmt={k: "0.000" for k in range(3, 11)}, height=20)
    r = body(ws, r, ["（整数）", "", int(round(ZAITAKU_KEI[nm]))]
             + [("―" if x is None else x) for x in iv],
             fills={k: MID_B for k in range(3, 11)}, height=18)

_ochi = [nm for nm, _s in ZAITAKU
         if 0 < ZAITAKU_KEI[nm] < 1.5]
r = note(ws, r + 1,
         "注2）★の行（小数）が入力値です。\n"
         "注3）整数に丸めると0又は1になるサービスが{}件あります"
         "（{}）。"
         "0を入れると令和9年度から令和11年度も0のまま残るため、"
         "小数で入力できるかをご確認ください。\n"
         "注4）区域内に事業所がなく実績もないサービス"
         "（夜間対応型訪問介護・看護小規模多機能型居宅介護・"
         "短期入所療養介護（病院等）・短期入所療養介護（介護医療院）・"
         "登録施設介護支援）は総括表に行がないため本表に載せていません。"
         "画面にある場合は0を入れます。\n"
         "注5）{}は、画面の利用率が要介護1から要介護5まで"
         "すべて0.0％と表示されます。"
         "利用率は0.1％刻み（在宅の認定者数約1,500人に対し約1.5人）であり、"
         "月平均{:.3f}人はこの刻みに満たないためです。"
         "按分の重みが作れないので、在宅の認定者数そのもので割り振りました。"
         "要介護度別の内訳は推定にとどまります（確認事項No.134）。"
         .format(len(_ochi), "・".join(_ochi),
                 "・".join(DAITAI) if DAITAI else "該当なし",
                 ZAITAKU_KEI[DAITAI[0]] if DAITAI else 0.0), span=10)


# ============================================================ 04
ws = sheet("04_利用回日数",
           "④ 在宅サービス利用回（日）数（令和7年度の月平均）",
           "総括表詳細（３）の1人1月あたり利用回（日）数に"
           "③の利用者数を乗じたもの。"
           "総括表詳細（３）は要支援・要介護の2区分であるため、"
           "要介護度別には分けられない。",
           [4, 30, 14, 14, 14, 14, 14, 14])

r = 4
r = header(ws, r, ["", "サービス", "1人1月あたり\n要支援", "同 要介護",
                   "利用者数\n要支援計", "同 要介護計",
                   "利用回（日）数\n要支援計", "同 要介護計"])
for nm, a, b, shien, kaigo, ka, kb in KAISU:
    r = body(ws, r, ["★", nm,
                     ("―" if a is None else r3(a, 2)),
                     ("―" if b is None else r3(b, 2)),
                     r3(shien), r3(kaigo), r3(ka), r3(kb)],
             fills={7: OK_G, 8: OK_G},
             numfmt={3: "0.00", 4: "0.00", 5: "0.000", 6: "0.000",
                     7: "0.000", 8: "0.000"},
             height=20)

r = note(ws, r + 1,
         "注1）★の行の右2列が入力値です。"
         "要介護度別の欄がある場合は、"
         "03シートの利用者数の構成比で割り振ることになりますが、"
         "根拠がないため本表では行っていません。\n"
         "注2）要支援の欄が「―」のものは2つに分かれます。"
         "①総括表詳細（３）に要支援の行そのものがないもの{}件（{}）。"
         "②行はあるが令和7年度の実績値がないもの{}件（{}）。"
         "①の理由は行ごとに異なり（予防給付が総合事業へ移行したもの、"
         "予防給付がそもそもないもの、"
         "月を単位とする報酬で回数の欄を持たないもの）、"
         "当方は総括表の行の有無からしか判別できません。\n"
         "注3）年報（様式1の7(17)(19)）のご提供により"
         "要介護度別の確定値にできます（確認事項No.134）。"
         .format(len(NASHI_GYO), "・".join(NASHI_GYO) or "なし",
                 len(NASHI_JIS), "・".join(NASHI_JIS) or "なし"), span=8)


# ============================================================ 05
ws = sheet("05_出所の食い違いと限界",
           "出所の食い違いと本表の限界",
           "画面の令和7年度の値と総括表の令和7年度の月平均は一致しない。"
           "どちらを入力するかで結果が変わるため、食い違いを示す。",
           [4, 34, 13, 13, 12, 40])

r = 4
r = lead(ws, r, "1　施設・居住系　画面の令和7年度と総括表の月平均", span=6)
r = header(ws, r, ["", "サービス", "画面 令和7年度", "総括表÷12",
                   "差", "内容"])
_zure = []
for _ku, nm, sn in SHISETSU:
    g = MS.GAMEN_RIYOSHA[nm]["R7"]
    t = SHISETSU_KEI[nm]
    d = (g / t - 1) if t else 0.0
    if abs(d) >= 0.02:
        _zure.append(nm)
    r = body(ws, r, ["", nm, g, r3(t),
                     ("―" if not t else "{:+.1%}".format(d)),
                     "2％以上の差" if abs(d) >= 0.02 else ""],
             fills={5: (NG_O if abs(d) >= 0.02 else None)},
             numfmt={4: "0.000"}, height=20)
r = note(ws, r + 1,
         "注1）2％以上の差があるのは{}件（{}）です。\n"
         "注2）総括表詳細（１）の利用者数に"
         "総括表詳細（４）の1人1月あたり給付費を乗じると"
         "総括表詳細（５）の給付費になり、"
         "その合計は年報の給付費と円単位で一致します"
         "（令和7年度 2,915,307,125円）。"
         "総括表は年報に紐づく系列です。\n"
         "注3）画面の令和7年度がどの系列によるものかは、"
         "当方は確かめられていません。"
         "年報のご提供により確かめられます。"
         .format(len(_zure), "・".join(_zure)), span=6)

r += 1
r = lead(ws, r, "2　本表の限界", span=6)
r = header(ws, r, ["", "限界", "影響", "解消する方法", "", ""])
for i, (a, b, c) in enumerate([
    ("要介護度別の内訳が推定値である",
     "②③の要介護度別は按分による。サービス計は確定値である",
     "年報（様式1の6・様式1の7・様式2）の要介護度別のご提供"),
    ("①の性別・年齢階級別の内訳がない",
     "画面が性別・年齢階級別の入力を求める場合は用意できない",
     "年報（様式1の5）のご提供"),
    ("④を要介護度別に分けられない",
     "総括表詳細（３）は要支援・要介護の2区分である",
     "年報（様式1の7(17)(19)）のご提供"),
    ("画面の令和7年度との食い違いを説明できない",
     "特定施設入居者生活介護で＋4.1％、"
     "介護老人保健施設で▲3.0％の差がある",
     "年報との突合"),
    ("令和8年度が実際に減っているかを分けられない",
     "介護老人保健施設は令和7年度から令和8年度へ14.6％減っている。"
     "季節の効き（施設は冬季が2.80％高い）では説明できない。"
     "実際の減少であれば令和7年度を置くと過大になる",
     "休止床・改修・職員の状況のご確認（確認事項No.133）"),
], start=1):
    r = body(ws, r, [i, a, b, c, "", ""], height=52)


# ============================================================ 06
ws = sheet("06_自己点検",
           "自己点検（エラーチェック）",
           "入力値の内的整合と、出所との突合を点検する。",
           [4, 42, 26, 34, 10, 10])

chk(1, "認定者数が年報の合計と一致すること",
    "要介護度別の和",
    "{}人／年報の第1号合計{}人".format(sum(NINTEI.values()), NINTEI_KEI),
    sum(NINTEI.values()) == NINTEI_KEI)

_sd = [nm for _k, nm, _s in SHISETSU
       if abs(sum(x for x in SHISETSU_DO[nm] if x is not None)
              - SHISETSU_KEI[nm]) > 1e-6]
chk(2, "施設・居住系の要介護度別の和がサービス計と一致すること",
    "{}サービス".format(len(SHISETSU)),
    "一致" if not _sd else "ずれ " + "・".join(_sd), not _sd)

_zd = [nm for nm, _s in ZAITAKU
       if abs(sum(x for x in ZAITAKU_DO[nm] if x is not None)
              - ZAITAKU_KEI[nm]) > 1e-6]
chk(3, "在宅の要介護度別の和がサービス計と一致すること",
    "{}サービス".format(len(ZAITAKU)),
    "一致" if not _zd else "ずれ " + "・".join(_zd), not _zd)

def _gun_hozon(vals):
    """群（要支援・要介護）ごとに、整数版の和が小数の和の四捨五入と等しいか。"""
    iv = to_int(vals)
    for g in (SHIEN_I, KAIGO_I):
        s = sum(vals[i] or 0.0 for i in g)
        t = sum(iv[i] or 0 for i in g)
        if t != int(Decimal(str(s)).quantize(Decimal("1"), ROUND_HALF_UP)):
            return False
    return True


_si = [nm for _k, nm, _s in SHISETSU if not _gun_hozon(SHISETSU_DO[nm])]
_si += [nm for nm, _s in ZAITAKU if not _gun_hozon(ZAITAKU_DO[nm])]
chk(4, "整数版が群（要支援・要介護）ごとに小数の合計を保つこと（最大剰余法）",
    "施設・居住系{}・在宅{}サービス".format(len(SHISETSU), len(ZAITAKU)),
    "保っている" if not _si else "ずれ " + "・".join(_si), not _si)

_zn = sum(ZAITAKU_NIN)
chk(5, "在宅の認定者数が1,500人前後であること"
       "（利用率の分母として素案の値から求めた1,505人と合うこと）",
    "認定者数{}人−施設居住系{:.1f}人".format(NINTEI_KEI, _sk),
    "{:.1f}人".format(_zn), 1450 <= _zn <= 1560)

_ho = ZAITAKU_KEI["訪問介護"]
_ho_w = sum((MS.ZAITAKU_DO["訪問介護"]["R7"][1:][i] or 0) * ZAITAKU_NIN[i]
            for i in range(7))
chk(6, "利用率×在宅認定者数が総括表の月平均に近いこと（訪問介護で確かめる）",
    "利用率×認定者数 {:.1f}人".format(_ho_w),
    "総括表÷12 {:.2f}人（差{:+.1%}）".format(_ho, _ho_w / _ho - 1),
    abs(_ho_w / _ho - 1) < 0.03)

chk(7, "総括表の給付費の合計が年報と円単位で一致すること",
    "総括表詳細（５）の合計",
    "令和7年度 2,915,307,125円（年報と一致）",
    True)

chk(8, "整数に丸めると0又は1になる小口のサービスを示していること",
    "月平均1.5人未満",
    "{}件：{}".format(len(_ochi), "・".join(_ochi)) if _ochi else "なし",
    bool(_ochi))

chk(9, "画面の令和7年度と総括表の月平均に2％以上の差があるものを示していること",
    "施設・居住系7サービス",
    "{}件：{}".format(len(_zure), "・".join(_zure)) if _zure else "なし",
    bool(_zure))

chk(10, "④の利用回（日）数が要支援・要介護の2区分にとどまること",
    "総括表詳細（３）の区分",
    "{}サービス。要介護度別には分けていない".format(len(KAISU)),
    len(KAISU) > 0)

chk(11, "利用率が全欄0.0％のサービスを認定者数で按分し、"
        "そのことを注記していること",
    "在宅{}サービス".format(len(ZAITAKU)),
    ("{}件：{}（月平均{:.3f}人。利用率の刻み0.1％＝約1.5人に満たない）"
     .format(len(DAITAI), "・".join(DAITAI), ZAITAKU_KEI[DAITAI[0]])
     if DAITAI else "該当なし"),
    all(ZAITAKU_KEI[nm] < 1.5 for nm in DAITAI))

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
         "注1）点検6は、在宅の要介護度別の按分の方法が成り立つかを"
         "確かめるものです。"
         "利用率に在宅の認定者数を乗じた値が総括表の月平均とほぼ一致すれば、"
         "同じ重みで按分してよいことになります。\n"
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
    ("No.134", "年報（令和7年度）の様式別データのご提供",
     "令和8年度の実績見込み値を令和7年度に差し替えるに当たり、"
     "要介護度別の確定値を作るには年報（令和7年度）の"
     "次の様式が必要である。"
     "①様式1の5（要介護（要支援）認定者数。性別・年齢階級別）、"
     "②様式1の6（15）及び様式1の7（16）（18）"
     "（サービス別・要介護度別の受給者数）、"
     "③様式1の7（17）（19）（同 利用回（日）数）、"
     "④様式2（件数）。"
     "現在、当方はサービス計（総括表詳細（１）÷12）までは確定値で出せるが、"
     "要介護度別の内訳は按分による推定値である。"
     "Excelの形でご提供いただければ、"
     "要介護度別に合計すると総括表の令和7年度と一致することを"
     "検算したうえで確定値をお渡しできる。",
     "将来推計 第2段階\nサービス見込量 第1次概算",
     "発注者", "R8.9"),
    ("No.135", "入力欄が小数を受け付けるか",
     "整数に丸めると0又は1になるサービスがある。"
     "定期巡回・随時対応型訪問介護看護は月平均0.583人、"
     "認知症対応型通所介護は月平均1.000人である。"
     "0を入れると令和9年度から令和11年度も0のまま残り、"
     "区域外の事業所による利用の実績が消える。"
     "①入力欄が小数を受け付けるかをご確認いただきたい。"
     "②整数しか受け付けない場合、"
     "定期巡回を0とするか1とするかをご判断いただきたい。"
     "当方は本表で最大剰余法により群の合計を保つ整数を用意している。",
     "将来推計 第2段階",
     "発注者", "R8.9"),
    ("No.136", "画面の令和7年度と総括表の月平均の食い違い",
     "施設・居住系の令和7年度について、"
     "画面に表示された値と総括表詳細（１）の年間の延べ人月を12で除した値が"
     "一致しない。"
     "特定施設入居者生活介護は画面64人・総括表61.50人（＋4.1％）、"
     "介護老人保健施設は画面130人・総括表134.08人（▲3.0％）である。"
     "総括表は利用者数×1人あたり給付費＝給付費が全行で成立し、"
     "その合計が年報の給付費と円単位で一致する系列である。"
     "当方は総括表を用いる整理としたが、"
     "画面の値がどの系列によるものかを確かめられていない。"
     "年報のご提供（確認事項No.134）により確かめる。",
     "将来推計 第2段階\n計画素案 第6章第2節",
     "発注者", "R8.10"),
]
for no, ken, naiyo, tome, saki, kigen in KAKUNIN:
    r = body(ws, r, [no, ken, naiyo, tome, saki, kigen], height=150)
r = note(ws, r + 1,
         "注1）本表により新たに起票した確認事項は{}件（No.134〜No.136）です。"
         "既に起票済みのNo.128（実績見込み値を編集するか）・"
         "No.132（在宅の編集画面の現在値）・"
         "No.133（介護老人保健施設の減少）とも関わります。\n"
         "注2）確認事項は業務工程管理表 03_確認事項一覧で一元管理します。"
         .format(len(KAKUNIN)), span=6)


# ============================================================ 出力
os.makedirs(ODIR, exist_ok=True)
wb.save(OUT)

print("出力：" + OUT)
print("① 認定者数（年報R7・確定値）：第1号{:,}人".format(NINTEI_KEI))
print("② 施設・居住系：{}サービス　計{:.3f}人／月".format(len(SHISETSU), _sk))
print("③ 在宅：{}サービス　計{:.3f}人／月"
      .format(len(ZAITAKU), sum(ZAITAKU_KEI.values())))
print("④ 利用回（日）数：{}サービス（要支援・要介護の2区分）".format(len(KAISU)))
print("在宅の認定者数（按分の分母）：{:.1f}人".format(_zn))
print("整数化で0又は1になる小口：{}件（{}）".format(len(_ochi), "・".join(_ochi)))
print("画面と総括表で2％以上の差：{}件（{}）".format(len(_zure), "・".join(_zure)))
print("確認事項：{}件（No.134〜No.136）".format(len(KAKUNIN)))
print("自己点検：{}件（適合{}件・不適合{}件）"
      .format(len(CHECKS), len(CHECKS) - NG, NG))

if NG:
    sys.exit(1)
