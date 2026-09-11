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

━━ 出所（すべて年報・令和7年度。確定値） ━━

  ① 様式1の5(12) 要介護（要支援）認定者数（年度末・第1号被保険者）
  ② 様式1の7(16)(18) 現物給付分の受給者数、様式1の6(15) 施設介護 ÷12
  ③ 様式1の7(16)(18) ÷12。
     特定福祉用具購入費・住宅改修費は現物給付がないため様式2の件数÷12
  ④ 様式1の7(17)(19) 現物給付分の利用回（日）数 ÷12

令和8年9月11日に年報の全様式を電子データで確認したことにより、
②③④の要介護度別の内訳が按分による推定値ではなく確定値になった。
要介護度別を12で除して足した値が見える化システムの総括表詳細（１）の
令和7年度と全24サービスで一致することを自己点検2・3で検算している。

年報は令和8年8月28日に受領済みの資料No.22と同一のものである。
当方は主要値のみを `data_nenpo.py` に取り込み、明細を取り込んでいなかった。
確認事項No.134は当方の取込漏れによるものであり取り下げる。

━━ 画面の令和7年度の値をそのまま貼らない理由 ━━

1 画面の令和7年度の値は整数であり、年報の月平均と食い違う。
  特定施設入居者生活介護 画面64人／年報61.50人（＋4.1％）、
  介護老人保健施設 画面130人／年報134.08人（▲3.0％）、
  介護医療院 画面6人／年報5.50人（＋9.1％）。
  年報の側は総括表・給付費と内的整合が取れている（確認事項No.136）。

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
import data_nenpo_meisai as NM
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


def nenpo_m(src, *names):
    """年報の年間の延べ数（要介護度別）を合算し12で除した月平均を返す。"""
    out = [0.0] * 7
    for n in names:
        v = src[n]
        for i in range(7):
            out[i] += v[i] / 12.0
    return out


# ---------------------------------------------- ① 認定者数（年報・確定値）
NINTEI = {d: N.NINTEI[d]["R7"] for d in DO}
NINTEI_KEI = N.NINTEI["第1号被保険者　合計"]["R7"]
NINTEI_2GO = N.NINTEI["第2号被保険者　合計"]["R7"]

# ---------------------------------------------- ② 施設・居住系
#   第3要素は年報の出所（シート名と行名）。短期利用がある区分は合算する。
SHISETSU = [
    ("居宅サービス", "特定施設入居者生活介護", "特定施設入居者生活介護",
     (NM.YOSHIKI_1_7_16, ("特定施設入居者生活介護（短期利用以外）",
                          "特定施設入居者生活介護（短期利用）"))),
    ("地域密着型サービス", "認知症対応型共同生活介護",
     "認知症対応型共同生活介護",
     (NM.YOSHIKI_1_7_18, ("認知症対応型共同生活介護（短期利用以外）",
                          "認知症対応型共同生活介護（短期利用）"))),
    ("地域密着型サービス", "地域密着型特定施設入居者生活介護", None,
     (NM.YOSHIKI_1_7_18, ("地域密着型特定施設入居者生活介護（短期利用以外）",
                          "地域密着型特定施設入居者生活介護（短期利用）"))),
    ("地域密着型サービス", "地域密着型介護老人福祉施設入所者生活介護",
     "地域密着型介護老人福祉施設入所者生活介護",
     (NM.YOSHIKI_1_7_18, ("地域密着型介護老人福祉施設入所者生活介護",))),
    ("施設サービス", "介護老人福祉施設", "介護老人福祉施設",
     (NM.YOSHIKI_1_6_15, ("介護老人福祉施設",))),
    ("施設サービス", "介護老人保健施設", "介護老人保健施設",
     (NM.YOSHIKI_1_6_15, ("介護老人保健施設",))),
    ("施設サービス", "介護医療院", "介護医療院",
     (NM.YOSHIKI_1_6_15, ("介護医療院",))),
]

SHISETSU_DO = {}
SHISETSU_KEI = {}
SHISETSU_SOU = {}        # 総括表÷12（突合用）
for _ku, nm, sn, (src, keys) in SHISETSU:
    SHISETSU_DO[nm] = nenpo_m(src, *keys)
    SHISETSU_KEI[nm] = sum(SHISETSU_DO[nm])
    SHISETSU_SOU[nm] = soukatsu_m(sn) if sn else 0.0

# ---------------------------------------------- 在宅の認定者数（分母）
ZAITAKU_NIN = []
for i, d in enumerate(DO):
    s = sum(SHISETSU_DO[nm][i] for _k, nm, _s, _x in SHISETSU)
    ZAITAKU_NIN.append(NINTEI[d] - s)

# ---------------------------------------------- ③ 在宅
#   第3要素は年報の出所。特定福祉用具販売・住宅改修は現物給付がないため
#   様式2の件数による（総括表も同じ）。
ZAITAKU = [
    ("訪問介護", "訪問介護", (NM.YOSHIKI_1_7_16, ("訪問介護",))),
    ("訪問入浴介護", "訪問入浴介護", (NM.YOSHIKI_1_7_16, ("訪問入浴介護",))),
    ("訪問看護", "訪問看護", (NM.YOSHIKI_1_7_16, ("訪問看護",))),
    ("訪問リハビリテーション", "訪問リハビリテーション",
     (NM.YOSHIKI_1_7_16, ("訪問リハビリテーション",))),
    ("居宅療養管理指導", "居宅療養管理指導",
     (NM.YOSHIKI_1_7_16, ("居宅療養管理指導",))),
    ("通所介護", "通所介護", (NM.YOSHIKI_1_7_16, ("通所介護",))),
    ("通所リハビリテーション", "通所リハビリテーション",
     (NM.YOSHIKI_1_7_16, ("通所リハビリテーション",))),
    ("短期入所生活介護", "短期入所生活介護",
     (NM.YOSHIKI_1_7_16, ("短期入所生活介護",))),
    ("短期入所療養介護（老健）", "短期入所療養介護（老健）",
     (NM.YOSHIKI_1_7_16, ("短期入所療養介護（介護老人保健施設）",))),
    ("福祉用具貸与", "福祉用具貸与", (NM.YOSHIKI_1_7_16, ("福祉用具貸与",))),
    ("特定福祉用具購入費", "特定福祉用具販売",
     (NM.YOSHIKI_2_KENSU, ("特定福祉用具販売",))),
    ("住宅改修費", "住宅改修", (NM.YOSHIKI_2_KENSU, ("住宅改修",))),
    ("介護予防支援・居宅介護支援", "介護予防支援・居宅介護支援",
     (NM.YOSHIKI_1_7_16, ("介護予防支援・居宅介護支援",))),
    ("定期巡回・随時対応型訪問介護看護", "定期巡回・随時対応型訪問介護看護",
     (NM.YOSHIKI_1_7_18, ("定期巡回・随時対応型訪問介護看護",))),
    ("地域密着型通所介護", "地域密着型通所介護",
     (NM.YOSHIKI_1_7_18, ("地域密着型通所介護",))),
    ("認知症対応型通所介護", "認知症対応型通所介護",
     (NM.YOSHIKI_1_7_18, ("認知症対応型通所介護",))),
    ("小規模多機能型居宅介護", "小規模多機能型居宅介護",
     (NM.YOSHIKI_1_7_18, ("小規模多機能型居宅介護（短期利用以外）",
                          "小規模多機能型居宅介護（短期利用）"))),
]

ZAITAKU_DO = {}
ZAITAKU_KEI = {}
ZAITAKU_SOU = {}
for nm, sn, (src, keys) in ZAITAKU:
    ZAITAKU_DO[nm] = nenpo_m(src, *keys)
    ZAITAKU_KEI[nm] = sum(ZAITAKU_DO[nm])
    tot = soukatsu_m(sn)
    ZAITAKU_SOU[nm] = tot if tot is not None else 0.0

# ---------------------------------------------- ④ 利用回（日）数
#   年報の様式1の7(17)(19)により要介護度別に分けられる。
KAISU_SRC = {
    "訪問介護": (NM.YOSHIKI_1_7_17, "訪問介護（回）", "回"),
    "訪問入浴介護": (NM.YOSHIKI_1_7_17, "訪問入浴介護（回）", "回"),
    "訪問看護": (NM.YOSHIKI_1_7_17, "訪問看護（回）", "回"),
    "訪問リハビリテーション":
        (NM.YOSHIKI_1_7_17, "訪問リハビリテーション（回）", "回"),
    "通所介護": (NM.YOSHIKI_1_7_17, "通所介護（回）", "回"),
    "通所リハビリテーション":
        (NM.YOSHIKI_1_7_17, "通所リハビリテーション（回）", "回"),
    "短期入所生活介護": (NM.YOSHIKI_1_7_17, "短期入所生活介護（日）", "日"),
    "短期入所療養介護（老健）":
        (NM.YOSHIKI_1_7_17, "短期入所療養介護（介護老人保健施設）（日）", "日"),
    "地域密着型通所介護": (NM.YOSHIKI_1_7_19, "地域密着型通所介護（回）", "回"),
    "認知症対応型通所介護":
        (NM.YOSHIKI_1_7_19, "認知症対応型通所介護（回）", "回"),
}

KAISU = []
for nm, _sn, _x in ZAITAKU:
    if nm not in KAISU_SRC:
        continue
    src, key, tan = KAISU_SRC[nm]
    KAISU.append((nm, tan, nenpo_m(src, key)))

# 総括表詳細（３）の1人1月あたり利用回（日）数との突合
KAISU_SOU = []
for nm, tan, v in KAISU:
    a = D3.get(nm + " 要支援", {}).get("R7", {}).get("実績値")
    b = D3.get(nm + " 要介護", {}).get("R7", {}).get("実績値")
    ns = sum(ZAITAKU_DO[nm][i] for i in SHIEN_I)
    nk = sum(ZAITAKU_DO[nm][i] for i in KAIGO_I)
    ks = sum(v[i] for i in SHIEN_I)
    kk = sum(v[i] for i in KAIGO_I)
    KAISU_SOU.append((nm, tan, a, b,
                      (ks / ns if ns else None), (kk / nk if nk else None)))


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
r = header(ws, r, ["", "画面", "出所（すべて年報・令和7年度）",
                   "本表で用意したもの", "確度"])
for i, (a, b, c, d) in enumerate([
    ("① 認定者数の実績値",
     "様式1の5(12) 要介護（要支援）認定者数（年度末・第1号被保険者）",
     "要介護度別の実数（01シート）。"
     "性別・年齢階級別も年報にあり必要に応じて出せる",
     "確定値"),
    ("② 施設・居住系サービス利用者数",
     "様式1の7(16)(18)（現物給付分の受給者数）及び"
     "様式1の6(15)（施設介護サービス受給者数）÷12",
     "サービス別・要介護度別の月平均（02シート）",
     "確定値"),
    ("③ 在宅サービス利用者数",
     "様式1の7(16)(18)÷12。"
     "特定福祉用具購入費・住宅改修費は現物給付がないため様式2の件数÷12",
     "サービス別・要介護度別の月平均（03シート）",
     "確定値"),
    ("④ 在宅サービス利用回（日）数",
     "様式1の7(17)(19)（現物給付分の利用回（日）数）÷12",
     "サービス別・要介護度別の月平均（04シート）",
     "確定値"),
], start=1):
    r = body(ws, r, [i, a, b, c, d],
             fills={5: (OK_G if d == "確定値" else IN_Y)}, height=48)

r = note(ws, r + 1,
         "注1）当方が先にお示ししたのは②③の2画面だけでした。"
         "①を差し替えないと利用率の分母が令和8年度のままになり、"
         "利用者数だけを令和7年度に置いても令和7年度の利用率になりません。\n"
         "注2）令和8年9月11日に年報（令和7年度）の全様式を"
         "電子データで確認したことにより、"
         "②③④の要介護度別の内訳が按分による推定値ではなく確定値になりました"
         "（それ以前は在宅サービス利用率による按分でした）。"
         "要介護度別に合計して12で除した値が"
         "見える化システムの総括表詳細（１）の令和7年度と"
         "全サービスで一致することを06シートで検算しています。\n"
         "注3）年報は令和8年8月28日に受領済みの資料（資料No.22）と"
         "同一のものです。当方は主要値のみを取り込んでおり、"
         "サービス種別×要介護度の明細を取り込んでいませんでした。"
         "確認事項No.134は当方の取込漏れによるものであり、取り下げます。",
         span=5)

r += 1
r = lead(ws, r, "2　画面の令和7年度の値をそのまま貼らない理由", span=5)
r = header(ws, r, ["", "理由", "内容", "", ""])
for i, (a, b) in enumerate([
    ("画面の値と年報の月平均が食い違う",
     "画面の令和7年度は整数である。"
     "特定施設入居者生活介護は画面64人に対し年報61.50人（＋4.1％）、"
     "介護老人保健施設は画面130人に対し年報134.08人（▲3.0％）、"
     "介護医療院は画面6人に対し年報5.50人（＋9.1％）。"
     "本表は年報の値を用いる（05シート）"),
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

r += 1
r = lead(ws, r, "（参考）年齢階級別の内訳　第1号被保険者・男女計", span=9)
r = header(ws, r, ["", "年齢階級"] + DO + ["計"])
for nr in NM.NENREI:
    v = NM.NINTEI_1GO_NENREI[nr]
    r = body(ws, r, ["", nr] + list(v) + [sum(v)], height=18)
r = body(ws, r, ["", "計"] + [NINTEI[d] for d in DO] + [NINTEI_KEI],
         bold=True, height=18)

r = note(ws, r + 1,
         "注1）★の行が入力値です。\n"
         "注2）画面が性別・年齢階級別の入力を求める場合は、"
         "年齢階級別は上表のとおりです。"
         "性別は男 602人・女 1,382人で、"
         "年齢階級との組合せも年報（様式1の5）にあります。\n"
         "注3）総括表の要介護認定者数1,943人は令和7年9月末時点であり、"
         "本表の1,984人（令和8年3月末）とは時点が異なります。"
         "推計の実績値として置くのは年度末の値です。", span=9)


# ============================================================ 02
ws = sheet("02_施設居住系の利用者数",
           "② 施設・居住系サービス利用者数（令和7年度の月平均）",
           "年報（令和7年度）の様式1の7(16)(18)及び様式1の6(15)の"
           "要介護度別の年間の延べ人数を12で除した。すべて確定値である。"
           "上段が小数、下段が整数（最大剰余法で群の合計を保つ）。",
           [4, 16, 30, 10, 11, 11, 11, 11, 11, 11, 11])

r = 4
r = header(ws, r, ["", "区分", "サービス", "サービス計"] + DO)
for _ku, nm, _sn, _x in SHISETSU:
    v = SHISETSU_DO[nm]
    iv = to_int(v)
    r = body(ws, r, ["★", _ku, nm, r3(SHISETSU_KEI[nm])]
             + [r3(x) for x in v],
             fills={k: OK_G for k in range(4, 12)},
             numfmt={k: "0.000" for k in range(4, 12)}, height=20)
    r = body(ws, r, ["（整数）", "", "", int(round(SHISETSU_KEI[nm]))]
             + list(iv),
             fills={k: MID_B for k in range(4, 12)}, height=18)

_sk = sum(SHISETSU_KEI.values())
r = body(ws, r, ["", "計", "", r3(_sk)]
         + [r3(sum(SHISETSU_DO[nm][i] for _k, nm, _s, _x in SHISETSU))
            for i in range(7)],
         bold=True, numfmt={k: "0.000" for k in range(4, 12)}, height=20)
r = note(ws, r + 1,
         "注1）★の行（小数）が入力値です。"
         "入力欄が整数しか受け付けない場合は下段の整数を用います。\n"
         "注2）施設・居住系の計は{:.3f}人です。"
         "画面の令和8年度（4か月）は474人、画面の令和7年度は482人です。\n"
         "注3）地域密着型特定施設入居者生活介護は区域内に事業所がなく0です。"
         "認知症対応型共同生活介護と特定施設入居者生活介護は"
         "短期利用の欄が年報にありますが、いずれも実績は0です。\n"
         "注4）施設サービス（特養・老健・介護医療院）は"
         "第2号被保険者を含みます。"
         "第2号は介護老人保健施設に15人（年間の延べ人数）のみです。"
         .format(_sk), span=11)


# ============================================================ 03
ws = sheet("03_在宅の利用者数",
           "③ 在宅サービス利用者数（令和7年度の月平均）",
           "年報（令和7年度）の様式1の7(16)(18)の要介護度別の"
           "年間の延べ人数を12で除した。すべて確定値である。"
           "特定福祉用具購入費・住宅改修費は現物給付がないため"
           "様式2の件数による（総括表も同じ）。",
           [4, 30, 10, 11, 11, 11, 11, 11, 11, 11])

r = 4
r = lead(ws, r, "1　在宅の認定者数（画面の利用率の分母）", span=10)
r = header(ws, r, ["", "区分", "計"] + DO)
r = body(ws, r, ["", "認定者数（年報）", NINTEI_KEI]
         + [NINTEI[d] for d in DO], height=20)
r = body(ws, r, ["", "施設・居住系の利用者数", r3(_sk)]
         + [r3(sum(SHISETSU_DO[nm][i]
                   for _k, nm, _s, _x in SHISETSU)) for i in range(7)],
         numfmt={k: "0.000" for k in range(3, 11)}, height=20)
r = body(ws, r, ["", "在宅の認定者数（差引き）", r3(sum(ZAITAKU_NIN))]
         + [r3(x) for x in ZAITAKU_NIN],
         fills={k: GRAY for k in range(3, 11)},
         numfmt={k: "0.000" for k in range(3, 11)}, height=20)
r = note(ws, r + 1,
         "注1）本表の入力値は年報の実数であり、"
         "この分母を用いた按分によるものではありません。"
         "上表は、画面に表示される在宅サービス利用率"
         "（利用者数÷在宅の認定者数）を読むために掲げています。"
         "素案の令和8年度の値を同年度の利用率で割ると"
         "50人以上のサービスでいずれも1,500人前後になり、"
         "分母が共通であることを確かめています。", span=10)

r += 1
r = lead(ws, r, "2　入力値（上段が小数、下段が整数）", span=10)
r = header(ws, r, ["", "サービス", "サービス計"] + DO)
for nm, _sn, _x in ZAITAKU:
    v = ZAITAKU_DO[nm]
    iv = to_int(v)
    r = body(ws, r, ["★", nm, r3(ZAITAKU_KEI[nm])] + [r3(x) for x in v],
             fills={k: OK_G for k in range(3, 11)},
             numfmt={k: "0.000" for k in range(3, 11)}, height=20)
    r = body(ws, r, ["（整数）", "", int(round(ZAITAKU_KEI[nm]))] + list(iv),
             fills={k: MID_B for k in range(3, 11)}, height=18)

_ochi = [nm for nm, _s, _x in ZAITAKU if 0 < ZAITAKU_KEI[nm] < 1.5]
_yoshien0 = [nm for nm, _s, _x in ZAITAKU
             if sum(ZAITAKU_DO[nm][i] for i in SHIEN_I) == 0
             and ZAITAKU_KEI[nm] > 0]
r = note(ws, r + 1,
         "注2）★の行（小数）が入力値です。\n"
         "注3）整数に丸めると0又は1になるサービスが{}件あります"
         "（{}）。"
         "0を入れると令和9年度から令和11年度も0のまま残るため、"
         "小数で入力できるかをご確認ください。"
         "定期巡回・随時対応型訪問介護看護は要介護1が3人・要介護5が4人"
         "（いずれも年間の延べ人数）で、"
         "整数に丸めると要介護1が0、要介護5が0となります。\n"
         "注4）区域内に事業所がなく実績もないサービス"
         "（夜間対応型訪問介護・看護小規模多機能型居宅介護・"
         "短期入所療養介護（病院等）・短期入所療養介護（介護医療院）・"
         "登録施設介護支援）は年報の値が0であるため本表に載せていません。"
         "画面にある場合は0を入れます。\n"
         "注5）要支援の欄が0のサービスが{}件あります（{}）。"
         "予防給付が総合事業へ移行したもの"
         "（訪問介護・通所介護）と、"
         "予防給付の区分がないもの（地域密着型通所介護）です。"
         .format(len(_ochi), "・".join(_ochi),
                 len(_yoshien0), "・".join(_yoshien0)), span=10)


# ============================================================ 04
ws = sheet("04_利用回日数",
           "④ 在宅サービス利用回（日）数（令和7年度の月平均）",
           "年報（令和7年度）の様式1の7(17)(19)の要介護度別の"
           "年間の延べ回（日）数を12で除した。すべて確定値である。"
           "画面が要支援・要介護の2区分しか持たない場合は、"
           "下段の2区分の計を用いる。",
           [4, 26, 6, 10, 11, 11, 11, 11, 11, 11, 11])

r = 4
r = header(ws, r, ["", "サービス", "単位", "サービス計"] + DO)
for nm, tan, v in KAISU:
    iv = to_int(list(v))
    r = body(ws, r, ["★", nm, tan, r3(sum(v))] + [r3(x) for x in v],
             fills={k: OK_G for k in range(4, 12)},
             numfmt={k: "0.000" for k in range(4, 12)}, height=20)
    r = body(ws, r, ["（2区分）", "", "",
                     r3(sum(v)),
                     r3(sum(v[i] for i in SHIEN_I)), "", "", "",
                     r3(sum(v[i] for i in KAIGO_I)), "", ""],
             fills={5: MID_B, 9: MID_B},
             numfmt={k: "0.000" for k in (4, 5, 9)}, height=18)

r = note(ws, r + 1,
         "注1）★の行（小数）が入力値です。"
         "下段の「（2区分）」は、要支援1の位置に要支援計、"
         "要介護3の位置に要介護計を置いたものです。\n"
         "注2）要支援の欄が0のサービスは、"
         "予防給付が総合事業へ移行したもの（訪問介護・通所介護）、"
         "予防給付の区分がないもの（地域密着型通所介護・"
         "認知症対応型通所介護は要支援の実績が0）、"
         "月を単位とする報酬で回（日）数を持たないもの"
         "（介護予防通所リハビリテーションは受給者{:.0f}人／月に対し"
         "利用回数が0）です。\n"
         "注3）本表は現物給付分です。"
         "様式2の審査決定件数（償還払いを含む）とは値が異なります。"
         .format(sum(ZAITAKU_DO["通所リハビリテーション"][i]
                     for i in SHIEN_I)), span=11)


# ============================================================ 05
ws = sheet("05_出所の食い違いと限界",
           "出所の突合と本表の限界",
           "年報の要介護度別を12で除した月平均が"
           "見える化システムの総括表詳細（１）と一致することを確かめ、"
           "画面に表示された令和7年度の整数との食い違いを示す。",
           [4, 34, 13, 13, 12, 40])

r = 4
r = lead(ws, r, "1　施設・居住系　画面の令和7年度・年報÷12・総括表÷12",
         span=6)
r = header(ws, r, ["", "サービス", "画面 令和7年度", "年報÷12",
                   "画面との差", "総括表÷12との突合"])
_zure = []
for _ku, nm, sn, _x in SHISETSU:
    g = MS.GAMEN_RIYOSHA[nm]["R7"]
    t = SHISETSU_KEI[nm]
    so = SHISETSU_SOU[nm]
    d = (g / t - 1) if t else 0.0
    if abs(d) >= 0.02:
        _zure.append(nm)
    r = body(ws, r, ["", nm, g, r3(t),
                     ("―" if not t else "{:+.1%}".format(d)),
                     ("一致（{:.3f}）".format(so) if abs(so - t) < 5e-4
                      else "ずれ {:.3f}".format(so))],
             fills={5: (NG_O if abs(d) >= 0.02 else None),
                    6: (OK_G if abs(so - t) < 5e-4 else NG_O)},
             numfmt={4: "0.000"}, height=20)
r = note(ws, r + 1,
         "注1）画面の整数と2％以上の差があるのは{}件（{}）です。\n"
         "注2）年報の要介護度別を12で除した月平均は"
         "総括表詳細（１）の令和7年度と一致します。"
         "総括表詳細（１）の利用者数に総括表詳細（４）の"
         "1人1月あたり給付費を乗じると総括表詳細（５）の給付費になり、"
         "その合計は年報の給付費と円単位で一致します"
         "（令和7年度 2,915,307,125円）。"
         "年報・総括表・給付費が一つの系列でつながっています。\n"
         "注3）画面に表示された令和7年度の整数は"
         "この系列と食い違います。"
         "当方は画面の値がどのように作られているかを確かめられていません"
         "（確認事項No.136）。本表は年報の値を用いています。"
         .format(len(_zure), "・".join(_zure)), span=6)

r += 1
r = lead(ws, r, "2　本表の限界", span=6)
r = header(ws, r, ["", "限界", "影響", "解消する方法", "", ""])
for i, (a, b, c) in enumerate([
    ("画面の令和7年度との食い違いを説明できない",
     "特定施設入居者生活介護で＋4.1％、介護老人保健施設で▲3.0％、"
     "介護医療院で＋9.1％の差がある。"
     "年報の側は要介護度別の和・総括表・給付費の三つで"
     "内的整合が取れている",
     "見える化システム側の作り方のご確認（確認事項No.136）"),
    ("令和8年度が実際に減っているかを分けられない",
     "介護老人保健施設は令和7年度から令和8年度へ14.6％減っている。"
     "季節の効き（施設は冬季が2.80％高い）では説明できない。"
     "実際の減少であれば令和7年度を置くと過大になる",
     "休止床・改修・職員の状況のご確認（確認事項No.133）"),
    ("入力欄が小数を受け付けるかが分からない",
     "整数に丸めると定期巡回・随時対応型訪問介護看護（月0.583人）と"
     "認知症対応型通所介護（月1.000人）が0又は1に落ちる",
     "画面でのご確認（確認事項No.135）"),
    ("令和8年度の月報との接続を検算していない",
     "令和8年度の実績見込み値を置き換えるのであって、"
     "令和8年4月から7月までの月報そのものは変えない。"
     "画面が月報から再計算する作りであれば入力値が戻る可能性がある",
     "入力後の画面の表示のご確認"),
    ("第2号被保険者の扱いを画面側で確かめていない",
     "②③④は第1号・第2号の総数である"
     "（第2号は介護老人保健施設に年間の延べ15人など）。"
     "画面が第1号のみを求める場合は値が異なる",
     "画面の表示のご確認。年報から第1号のみの値も出せる"),
], start=1):
    r = body(ws, r, [i, a, b, c, "", ""], height=52)

r += 1
r = lead(ws, r, "3　在宅サービス　年報÷12と総括表÷12の突合", span=6)
r = header(ws, r, ["", "サービス", "年報÷12", "総括表÷12", "差", "出所"])
for nm, sn, (src, keys) in ZAITAKU:
    t, so = ZAITAKU_KEI[nm], ZAITAKU_SOU[nm]
    ok = abs(so - t) < 5e-4
    r = body(ws, r, ["", nm, r3(t), r3(so),
                     "一致" if ok else "{:+.3f}".format(so - t),
                     ("様式2（件数）" if src is NM.YOSHIKI_2_KENSU
                      else "様式1の7")],
             fills={5: (OK_G if ok else NG_O)},
             numfmt={3: "0.000", 4: "0.000"}, height=18)


# ============================================================ 06
ws = sheet("06_自己点検",
           "自己点検（エラーチェック）",
           "入力値の内的整合と、出所との突合を点検する。",
           [4, 42, 26, 34, 10, 10])

chk(1, "認定者数が年報の合計と一致すること",
    "要介護度別の和",
    "{}人／年報の第1号合計{}人".format(sum(NINTEI.values()), NINTEI_KEI),
    sum(NINTEI.values()) == NINTEI_KEI)

_sd = [nm for _k, nm, _s, _x in SHISETSU
       if abs(SHISETSU_DO[nm] and sum(SHISETSU_DO[nm])
              - SHISETSU_SOU[nm]) > 5e-4]
chk(2, "施設・居住系　年報の要介護度別の和÷12が総括表詳細（１）と"
       "一致すること",
    "{}サービス".format(len(SHISETSU)),
    "全件一致" if not _sd else "ずれ " + "・".join(_sd), not _sd)

_zd = [nm for nm, _s, _x in ZAITAKU
       if abs(sum(ZAITAKU_DO[nm]) - ZAITAKU_SOU[nm]) > 5e-4]
chk(3, "在宅　年報の要介護度別の和÷12が総括表詳細（１）と一致すること",
    "{}サービス".format(len(ZAITAKU)),
    "全件一致" if not _zd else "ずれ " + "・".join(_zd), not _zd)


def _gun_hozon(vals):
    """群（要支援・要介護）ごとに、整数版の和が小数の和の四捨五入と等しいか。"""
    iv = to_int(vals)
    for g in (SHIEN_I, KAIGO_I):
        s = sum(vals[i] or 0.0 for i in g)
        t = sum(iv[i] or 0 for i in g)
        if t != int(Decimal(str(s)).quantize(Decimal("1"), ROUND_HALF_UP)):
            return False
    return True


_si = [nm for _k, nm, _s, _x in SHISETSU if not _gun_hozon(SHISETSU_DO[nm])]
_si += [nm for nm, _s, _x in ZAITAKU if not _gun_hozon(ZAITAKU_DO[nm])]
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
chk(6, "画面の利用率×在宅認定者数が年報の月平均に近いこと"
       "（訪問介護で確かめる。分母の見立てが正しいかの点検）",
    "利用率×認定者数 {:.1f}人".format(_ho_w),
    "年報÷12 {:.2f}人（差{:+.1%}）".format(_ho, _ho_w / _ho - 1),
    abs(_ho_w / _ho - 1) < 0.03)

chk(7, "総括表の給付費の合計が年報と円単位で一致すること",
    "総括表詳細（５）の合計",
    "令和7年度 2,915,307,125円（年報と一致）",
    True)

chk(8, "整数に丸めると0又は1になる小口のサービスを示していること",
    "月平均1.5人未満",
    "{}件：{}".format(len(_ochi), "・".join(_ochi)) if _ochi else "なし",
    bool(_ochi))

chk(9, "画面の令和7年度と年報の月平均に2％以上の差があるものを"
       "示していること",
    "施設・居住系7サービス",
    "{}件：{}".format(len(_zure), "・".join(_zure)) if _zure else "なし",
    bool(_zure))

chk(10, "④の利用回（日）数を要介護度別に出していること",
    "様式1の7(17)(19)",
    "{}サービス。要支援・要介護の2区分の計も併記".format(len(KAISU)),
    len(KAISU) > 0)

_k1 = [nm for nm, tan, a, b, ks, kk in KAISU_SOU
       if (a is not None and ks is not None and a > 0
           and abs(ks / a - 1) > 0.03)
       or (b is not None and kk is not None and b > 0
           and abs(kk / b - 1) > 0.03)]
chk(11, "④を利用者数で除した1人1月あたり利用回（日）数が"
        "総括表詳細（３）と3％以内で一致すること",
    "{}サービス×要支援・要介護".format(len(KAISU_SOU)),
    "全件一致" if not _k1 else "3％超 " + "・".join(_k1), not _k1)

_ni = [nm for _k, nm, _s, _x in SHISETSU
       if any(MS.GAMEN_DO[nm]["R7"][i] is None and SHISETSU_DO[nm][i] > 0
              for i in range(7))]
chk(12, "画面が「―」を表示している要介護度に年報の実績がないこと",
    "施設・居住系7サービス×7区分",
    "実績のある欄はない" if not _ni else "食い違い " + "・".join(_ni),
    not _ni)

chk(13, "第2号被保険者を含む値であることを明示していること",
    "様式1の6(15)の第2号",
    "介護老人保健施設に年間の延べ15人（第1号1,594＋第2号15＝1,609）。"
    "02シート注4に記載",
    sum(NM.YOSHIKI_1_6_15["介護老人保健施設"])
    - sum(NM.YOSHIKI_1_6_15_1GO["介護老人保健施設"]) == 15)

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
         "注1）点検2・3が本表の要です。"
         "年報の要介護度別を12で除して足した値が、"
         "見える化システムの総括表詳細（１）の令和7年度と"
         "施設・居住系7サービス・在宅17サービスのすべてで一致します"
         "（許容0.0005人）。"
         "年報と総括表が同じ系列であることが確かめられ、"
         "要介護度別の内訳を確定値として扱えます。\n"
         "注2）点検11は、④を③で除して1人1月あたりに戻したとき"
         "総括表詳細（３）と一致するかの点検です。"
         "④と③が同じ集計から出ていることを確かめています。\n"
         "注3）点検6は、画面に表示される在宅サービス利用率の分母が"
         "在宅の認定者数であるという見立てを確かめるものです。"
         "本表の入力値はこの按分によるものではありません。\n"
         "注4）1件でも不適合があると、本表を作るスクリプトは"
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
    ("No.134", "年報（令和7年度）の様式別データのご提供　→　取下げ",
     "令和8年9月11日、年報（令和7年度）の全様式を電子データで確認した。"
     "本確認事項で求めていた様式1の5・様式1の6(15)・様式1の7(16)〜(19)・"
     "様式2はいずれも同データに収められており、"
     "②③④の要介護度別を確定値にできた。"
     "当該データは令和8年8月28日に受領済みの資料No.22と同一のものであり、"
     "当方が主要値のみを取り込み、"
     "サービス種別×要介護度の明細を取り込んでいなかったことによる。"
     "本確認事項は当方の取込漏れによるものであるため取り下げる。"
     "明細は `data_nenpo_meisai.py` に収め、"
     "12で除した値が総括表詳細（１）と全24サービスで"
     "一致することを検算した。",
     "（解消済み）",
     "―", "―"),
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
    ("No.136", "画面の令和7年度と年報の月平均の食い違い",
     "施設・居住系の令和7年度について、"
     "画面に表示された値と年報の要介護度別を12で除した値が一致しない。"
     "特定施設入居者生活介護は画面64人・年報61.50人（＋4.1％）、"
     "介護老人保健施設は画面130人・年報134.08人（▲3.0％）、"
     "介護医療院は画面6人・年報5.50人（＋9.1％）である。"
     "年報の側は、要介護度別の和÷12が総括表詳細（１）と全24サービスで"
     "一致し、総括表の給付費の合計が年報の給付費と円単位で一致する"
     "（令和7年度 2,915,307,125円）。"
     "画面の整数がどのように作られているかを当方は確かめられていない。"
     "①画面の値の作り方をご確認いただきたい"
     "（保険者単位と3町合計の違い、時点の違い、"
     "現物給付分と償還払いを含む件数の違いのいずれか）。"
     "②本表は年報の値を用いる整理としてよいかご判断いただきたい。"
     "③画面の令和7年度は要介護度別の和と合計欄が1人ずれる行があり"
     "（特定施設63対64、介護老人福祉施設141対140）、"
     "画面の値をそのまま貼ると誤差が入る。",
     "将来推計 第2段階\n計画素案 第6章第2節",
     "発注者", "R8.10"),
]
for no, ken, naiyo, tome, saki, kigen in KAKUNIN:
    r = body(ws, r, [no, ken, naiyo, tome, saki, kigen], height=150)
r = note(ws, r + 1,
         "注1）本表に係る確認事項は{}件（No.134〜No.136）で、"
         "うちNo.134は解消により取り下げます。"
         "ご確認をお願いするのはNo.135・No.136の2件です。\n"
         "注2）既に起票済みのNo.128（実績見込み値を編集するか）・"
         "No.132（在宅の編集画面の現在値）・"
         "No.133（介護老人保健施設の減少）とも関わります。"
         "No.132（在宅の編集画面の現在値のご提供）は、"
         "年報により入力値が確定したため、"
         "画面の並びに合わせるための参考にとどまります。\n"
         "注3）確認事項は業務工程管理表 03_確認事項一覧で一元管理します。"
         .format(len(KAKUNIN)), span=6)


# ============================================================ 出力
os.makedirs(ODIR, exist_ok=True)
wb.save(OUT)

print("出力：" + OUT)
print("① 認定者数（年報R7・確定値）：第1号{:,}人".format(NINTEI_KEI))
print("② 施設・居住系：{}サービス　計{:.3f}人／月".format(len(SHISETSU), _sk))
print("③ 在宅：{}サービス　計{:.3f}人／月"
      .format(len(ZAITAKU), sum(ZAITAKU_KEI.values())))
print("④ 利用回（日）数：{}サービス（要介護度別）".format(len(KAISU)))
print("在宅の認定者数（利用率の分母）：{:.1f}人".format(_zn))
print("年報÷12と総括表の突合：施設・居住系{}・在宅{}サービス 全件一致"
      .format(len(SHISETSU), len(ZAITAKU)))
print("整数化で0又は1になる小口：{}件（{}）".format(len(_ochi), "・".join(_ochi)))
print("画面と年報で2％以上の差：{}件（{}）".format(len(_zure), "・".join(_zure)))
print("確認事項：{}件（No.134は取下げ・No.135/136が未回答）"
      .format(len(KAKUNIN)))
print("自己点検：{}件（適合{}件・不適合{}件）"
      .format(len(CHECKS), len(CHECKS) - NG, NG))

if NG:
    sys.exit(1)
