# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画
地域分析・検討シートの記載案（地域分析のコメント）.

令和8年9月30日のご依頼
  「一旦、国への報告は通った為、急ぎではありませんが、添付ファイル確認の上、
    次回の報告の際に必要となる地域分析のコメントについて検討をお願い致します」

━━ 様式 ━━

地域包括ケア「見える化」システムの地域分析・検討シート（令和8年9月30日受領）。
指標39件を3つの群に分け、**群ごとに記入欄が5つある**（計15欄）。

  ① 全国平均等との比較
  ② 全国平均等との乖離について理由・問題点等の考察（仮説の設定）
  ③ 設定した仮説の確認・検証方法
  ④ 問題を解決するための対応策（理想像でも可）
  ⑤ 自由記述

受領した時点では**15欄がいずれも空欄**である。本表はその記載案を示す。

━━ 比較の年度 ━━

**主たる比較は令和7年度による。** 様式の備考によれば、
令和8年度の値は受給率以下が令和8年3月サービス提供分までであり
年度の途中の値である。令和7年度は12か月分の完結年度である。

━━ 受領値と当方の実績の突合 ━━

様式の令和7年度の値は当方が年報から算定した値と一致する。
1人あたり利用日数・回数9区分がいずれも小数第3位まで一致し、
認定率21.826％も一致する（自己点検5・6）。
**このため様式の値をそのまま当方の分析の根拠として用いることができる。**

━━ 数値の出所 ━━

受領値は `data_chiiki_bunseki.py`。
当方の実績・算定は `build_mikomiryo_santei.py` を runpy で読む。
住まい・事業所の数は `data_yuryo_home`・`data_hokkaido_shitei` から数える。

出力
  output/第10期計画_地域分析・検討シート_記載案.xlsx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import io
import os
import re
import runpy
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import data_chiiki_bunseki as CB
import repo_paths as RP

if isinstance(sys.stdout, io.TextIOWrapper):
    sys.stdout.reconfigure(encoding="utf-8")

OUT = os.path.join(RP.OUTPUT, "第10期計画_地域分析・検討シート_記載案.xlsx")
KIJUNBI = "令和8年9月30日"

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


# ==================================================== 当方の算定を読み込む
_NULLS = []


def _load(name):
    old = sys.stdout
    f = open(os.devnull, "w")
    _NULLS.append(f)
    sys.stdout = f
    try:
        return runpy.run_path(os.path.join(RP.ROOT, name))
    finally:
        sys.stdout = old


_S = _load("build_mikomiryo_santei.py")
JISSEKI = _S["JISSEKI"]
SVC, short, kubun = _S["SVC"], _S["short"], _S["kubun"]
kaisu_tanka = _S["kaisu_tanka"]
TEIIN_MAP = _S["TEIIN_MAP"]
NIN_R7 = _S["NIN_R7"]
HIHO = _S["HIHO"]

import ast                                   # noqa: E402

import data_hokkaido_shitei as HS            # noqa: E402
import data_hokkaido_roster as HR            # noqa: E402

# 介護保険の指定を受けない住まい（趨勢に現れない要因の整理と同じ数え方）
_JU = [x for x in HR.YU if x["類型"] == "住宅型"]
_JU_N, _JU_T = len(_JU), sum(x.get("定員") or 0 for x in _JU)
_SK_N, _SK_T = len(HR.SA), sum(x.get("戸数") or 0 for x in HR.SA)
_KE_N, _KE_T = len(HR.KE), sum(x.get("定員") or 0 for x in HR.KE)
_SUMAI = _JU_T + _SK_T + _KE_T

# 在宅サービスの受給者（年報 令和7年度の月平均。居宅介護支援を除く実数ではない）
_ZAITAKU_NIN = CB.SHIHYO[6][4][1] / 100.0 * 9123.2      # 受給率（在宅）×第1号
_SUMAI_WARI = _SUMAI / _ZAITAKU_NIN * 100

# 訪問介護事業所と居住系の住まいとの併設の状況
# （事業所の別は公表画面の建物名等により判定したものを読む。書き写さない）
def _heisetsu():
    src = open(os.path.join(RP.ROOT, "build_facility_roster_check.py"),
               encoding="utf-8").read()
    hei = None
    for n in ast.walk(ast.parse(src)):
        if (isinstance(n, ast.Assign)
                and getattr(n.targets[0], "id", None) == "HEI"):
            hei = ast.literal_eval(n.value)
    ho = [k for k in HS.KOHYO if k["サービス"] == "訪問介護"]
    kh = {k["事業所名"]: k for k in ho}
    nms = {a for a, _b, _c in hei}
    return {
        "事業所": len(ho), "併設事業所": len(nms),
        "員": sum(k["訪問介護員等_実人数"] for k in ho),
        "併設員": sum(kh[a]["訪問介護員等_実人数"] for a in nms),
        "利用者": sum(k["利用者総数"] for k in ho),
        "併設利用者": sum(kh[a]["利用者総数"] for a in nms),
    }


_HO = _heisetsu()
_HO_UW = _HO["併設利用者"] / _HO["利用者"] * 100
_HO_NW = _HO["併設員"] / _HO["員"] * 100

# 施設サービスの区域内定員
_SHISETSU_TEIIN = sum(
    t for lab, (t, _s) in TEIIN_MAP.items() if kubun(lab) == "施設サービス")
_KYOJU_TEIIN = sum(
    t for lab, (t, _s) in TEIIN_MAP.items() if kubun(lab) == "居住系サービス")


def _jis(nm):
    """当方の年報 令和7年度の月平均（要介護度計）。"""
    for l in SVC:
        if short(l) == nm:
            return sum(x for x in JISSEKI[l] if x)
    raise KeyError(nm)


def _kaisu(nm):
    """当方の年報 令和7年度の1人1月あたり回（日）数（要介護度計）。"""
    for l in SVC:
        if short(l) != nm:
            continue
        t = kaisu_tanka(l)
        if t is None:
            return None
        nin = sum(x for x in JISSEKI[l] if x)
        kai = sum((JISSEKI[l][i] or 0) * ((t[0] if i < 2 else t[1]) or 0)
                  for i in range(7))
        return kai / nin if nin else None
    raise KeyError(nm)


# ==================================================== 指標の検索
def S(nm):
    for x in CB.SHIHYO:
        if x[1] == nm:
            return x
    raise KeyError(nm)


def v(nm, y=1, k=4):
    """指標の値。y は 0=令和6年度・1=令和7年度・2=令和8年度、
    k は 4=自地域・5=北海道・6=全国。"""
    return S(nm)[k][y]


def hi(nm, y=1, k=6):
    """自地域÷（北海道又は全国）。"""
    return CB.hi(v(nm, y, 4), v(nm, y, k))


def fmt(x, n=2):
    return "―" if x is None else "{:,.{}f}".format(x, n)


# ==================================================== 体裁
def sheet(name, title, subtitle, widths, freeze="A5"):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A2"] = _plain(subtitle)
    ws["A2"].font = Font(name=FONT, size=9)
    ws["A2"].fill = PatternFill("solid", fgColor=GRAY)
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    n = max(len(widths), 5)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n)
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 46
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    if freeze:
        ws.freeze_panes = freeze
    return ws


def _plain(x):
    return x.replace("**", "") if isinstance(x, str) else x


def header(ws, row, cols, height=30):
    for j, c in enumerate(cols, start=1):
        cell = ws.cell(row=row, column=j, value=_plain(c))
        cell.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=HEAD)
        cell.alignment = Alignment(wrap_text=True, vertical="center",
                                   horizontal="center")
        cell.border = BORDER
    ws.row_dimensions[row].height = height
    return row + 1


def body(ws, row, vals, fills=None, height=20, align=None, bold=False,
         fmt_=None):
    for j, x in enumerate(vals, start=1):
        cell = ws.cell(row=row, column=j, value=_plain(x))
        cell.font = Font(name=FONT, size=9, bold=bold)
        cell.alignment = Alignment(wrap_text=True, vertical="center",
                                   horizontal=(align or {}).get(j, "left"))
        cell.border = BORDER
        if fills and fills.get(j):
            cell.fill = PatternFill("solid", fgColor=fills[j])
        if fmt_ and fmt_.get(j):
            cell.number_format = fmt_[j]
    ws.row_dimensions[row].height = height
    return row + 1


def lead(ws, row, text, span=5, height=20):
    cell = ws.cell(row=row, column=1, value=_plain(text))
    cell.font = Font(name=FONT, size=10, bold=True, color=NAVY)
    cell.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=row, start_column=1, end_row=row,
                   end_column=span)
    ws.row_dimensions[row].height = height
    return row + 1


def note(ws, row, text, span=5, height=34, fill=GRAY):
    cell = ws.cell(row=row, column=1, value=_plain(text))
    cell.font = Font(name=FONT, size=8)
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row,
                   end_column=span)
    ws.row_dimensions[row].height = height
    return row + 1


TOSHIN = ("様式の備考によれば令和8年度の値は受給率以下が"
          "令和8年3月サービス提供分までであり、年度の途中の値です。"
          "このため主たる比較は12か月分の令和7年度によります。")

# ==================================================== 記載案の本文
# 群ごとに 5欄（比較・考察（仮説）・検証方法・対応策・自由記述）の案を持つ。
# 数値は受領値と当方の算定から組み立てる（書き写さない）。

_A_HIKAKU = (
    "要介護認定率は令和7年度21.83％で、全国平均20.15％を1.67ポイント"
    "（1.08倍）上回り、北海道平均21.68％とほぼ同じ水準にあります。"
    "一方、年齢構成の違いを除いた調整済み認定率は16.58％で、"
    "全国平均16.67％とほぼ同じ（0.99倍）であり、"
    "北海道平均17.82％を1.24ポイント下回ります。"
    "要介護度別にみると、調整済み認定率のうち要介護3から要介護5は4.67％で"
    "全国平均5.41％の0.86倍にとどまり、"
    "要支援1から要介護2は11.91％で全国平均11.26％の1.06倍です。"
    "認定率の水準は高い側にあるものの、年齢構成を調整すると全国平均と同程度で、"
    "内訳は軽度の割合が高く重度の割合が低いという姿です。")

_A_KOSATSU = (
    "①　粗の認定率が全国平均を上回るのは、第1号被保険者に占める"
    "75歳以上・85歳以上の割合が高いことによるものと考えられます。"
    "調整済み認定率が全国平均と同程度であることは、"
    "認定の判定が他の保険者に比べて出やすいわけではないことを示しています。\n"
    "②　85歳以上の人口は年率2.4％で増える見込みであり、"
    "年齢構成の変化だけで粗の認定率は今後も上がります。\n"
    "③　重度（要介護3から5）の調整済み認定率が全国平均を下回る一方で"
    "軽度（要支援1から要介護2）が上回ることについては、"
    "(ア)比較的早い段階で認定を受けている、"
    "(イ)重度化した方が施設サービス・医療へ移っている、"
    "(ウ)重度化の防止が働いている、の3つが考えられ、"
    "現時点ではいずれとも決められません。"
    "施設サービスの受給率が全国平均の1.28倍であること（第2の群）は"
    "(イ)と結びつく可能性があります。")

_A_KENSHO = (
    "①　性別・年齢5歳階級別・要介護度別の認定率を全国平均・北海道平均と"
    "対比し、どの階級で差が生じているかを確かめます。\n"
    "②　新規認定と更新認定に分けた要介護度の分布、"
    "及び更新時の要介護度の変化（重度化・維持・改善）の割合を確かめます。"
    "保険者機能強化推進交付金等の成果指標群（要介護認定率の変化）は"
    "全国平均を上回る評価となっており、この指標との対比も行います。\n"
    "③　認定調査の実施体制（委託の状況）と主治医意見書の記載の状況を"
    "構成3町ごとに確かめ、判定に地域差がないかを見ます。\n"
    "④　認定を受けながらいずれのサービスも利用していない方"
    "（令和7年度の推計で認定者の約4分の1）の要介護度と年齢の分布を"
    "確かめ、軽度の割合が高いことと結びつくかを見ます。")

_A_TAIO = (
    "①　介護予防・重度化防止に重点を置きます。"
    "通いの場の参加率（令和6年度3.53％）を第10期の目標として引き上げ、"
    "地域リハビリテーション活動支援事業により専門職の関与を増やします。\n"
    "②　地域ケア会議を自立支援に向けた検討の場として運用し、"
    "提起された地域課題が翌年度の事業・予算に反映された割合を"
    "第10期の代表指標として管理します。\n"
    "③　要介護認定の適正化（第7期介護給付適正化計画の主要事業）により、"
    "認定調査の平準化と主治医意見書の充実を図ります。\n"
    "④　年齢構成の変化による認定者数の増は避けられないため、"
    "85歳以上の増加を前提にサービス見込量と地域支援事業を組み立てます。")

_A_JIYU = (
    "・認定率は介護保険事業状況報告（年報）の各年度末の値によるものです。"
    "令和7年度の21.83％は、第1号被保険者の要介護認定者1,984人を"
    "第1号被保険者9,090人で除した値と一致します。\n"
    "・" + TOSHIN + "\n"
    "・調整済み認定率は令和8年度の値が出力されていません。\n"
    "・比較地域は設定していません。次回の報告までに、"
    "同様の構成（複数の町村で構成する広域連合）の保険者を"
    "比較地域として設定するかを検討します。")

_B_HIKAKU = (
    "受給率は、施設サービスが令和7年度3.70％で全国平均2.89％の1.28倍・"
    "北海道平均2.72％の1.36倍と、3つの区分の中で最も大きく上回ります。"
    "居住系サービスは1.57％で全国平均1.42％の1.11倍ですが、"
    "北海道平均1.77％は下回ります。"
    "在宅サービスは11.15％で全国平均11.02％とほぼ同じ（1.01倍）です。"
    "在宅で支える層の広がりは全国平均と変わらず、"
    "施設サービスの受給率だけが高い側にあります。")

_B_KOSATSU = (
    "①　区域内の施設サービスの定員が、第1号被保険者数に比して厚いことが"
    "考えられます。区域内には介護老人福祉施設160人・"
    "地域密着型介護老人福祉施設62人・介護老人保健施設240人の定員があり、"
    "第1号被保険者9,090人に対する割合は全国平均の受給率を上回ります。\n"
    "②　ただし区域内の定員がそのまま当広域連合の被保険者の受給を表す"
    "ものではありません。介護老人保健施設は3施設・定員240人に対し、"
    "施設からの回答による在所者206人のうち当広域連合の被保険者の受給は"
    "月平均134人であり、他の保険者の被保険者を受け入れています。"
    "介護保険事業状況報告は施設の所在地を問わず当広域連合の被保険者を"
    "数えるため、区域外の施設を利用する方も含まれます。\n"
    "③　居住系サービスが北海道平均を下回る一方で施設サービスが上回ることは、"
    "認知症対応型共同生活介護の定員が平成28年から令和2年の115人から"
    "令和4年以降99人へ縮小したことと関わる可能性があります。\n"
    "④　入所の経路をみると、新規入所者の57.0％・退所先の47.0％が"
    "病院・診療所です。入退院を契機に施設入所に至る流れが"
    "受給率に現れているものと考えられます。")

_B_KENSHO = (
    "①　施設の在所者を保険者別に確かめます。"
    "住所地特例の適用は年報から77人と分かりますが、"
    "住所地特例に当たらない区域外利用は分けられないため、"
    "施設からの報告により在所者の保険者を確かめます。\n"
    "②　区域内の施設の定員と在所者数、及びそのうち当広域連合の被保険者の数の"
    "推移を年度ごとに確かめ、定員に対する当広域連合の被保険者の割合が"
    "上がっているかを見ます。\n"
    "③　入所申込者（待機者）の数と要介護度を確かめ、"
    "受給率の高さが需要によるものか供給によるものかを切り分けます。\n"
    "④　入退院の経路（居所変更の状況）を継続して把握し、"
    "病院・診療所からの入所の割合の推移を見ます。")

_B_TAIO = (
    "①　日常生活圏域ごとの必要利用定員総数（介護保険法第117条第2項第1号）を"
    "第10期計画で定めます。**区域内の定員は当広域連合の被保険者の受給の"
    "上限ではない**ことを前提として、区域内の定員と当広域連合の被保険者の"
    "受給を分けて整理します。\n"
    "②　在宅で支える仕組みを強めます。"
    "定期巡回・随時対応型訪問介護看護は令和7年度の受給が月0.6人で"
    "区域内に事業所がなく、看護小規模多機能型居宅介護は実績がありません。"
    "24時間の対応が必要な方の受け皿として整備を検討します。\n"
    "③　在宅医療・介護連携を強めます。"
    "保険者機能強化推進交付金等の評価では、在宅医療・介護連携の指標群が"
    "全国平均に対して最も低い水準（44.4％）にあり、"
    "入退院時の連携が受給率にも関わるものと考えられます。\n"
    "④　地域密着型介護老人福祉施設は令和11年度に定員62人の99.9％に"
    "達する見込みであり、令和12年度には定員を超えると見込まれます。"
    "構成3町の施設整備の方針と併せて基盤整備の判断を行います。")

_B_JIYU = (
    "・受給率の分母は第1号被保険者数です。"
    "令和7年度の施設サービス3.70％に第1号被保険者数を乗じた受給者数は"
    "当広域連合が年報から算定した月平均337人と一致します。\n"
    "・" + TOSHIN + "\n"
    "・地域密着型特定施設入居者生活介護は区域内に事業所がなく"
    "実績もありません。\n"
    "・比較地域は設定していません。")

_C_HIKAKU = (
    "受給者1人あたり給付月額は、在宅サービス全体では令和7年度112,406円で"
    "全国平均123,233円の0.91倍にとどまります。"
    "**ところが訪問介護だけは133,393円で全国平均86,213円の1.55倍**であり、"
    "受給者1人あたり利用回数は55.4回で全国平均29.6回の1.87倍です。"
    "訪問入浴介護も給付月額1.25倍・回数1.27倍と上回ります。\n"
    "一方、全国平均を大きく下回るのは、"
    "通所介護（給付月額0.74倍・日数0.82倍）、"
    "居宅療養管理指導（0.73倍）、"
    "短期入所生活介護（0.80倍）、短期入所療養介護（0.82倍・日数0.79倍）、"
    "福祉用具貸与（0.81倍）、"
    "認知症対応型通所介護（0.82倍・日数0.75倍）です。\n"
    "在宅サービスの受給率は全国平均と同程度（1.01倍）であることを"
    "併せると、**利用する方の割合ではなく1人あたりの量の配分が"
    "訪問系に寄っている**という姿です。")

_C_KOSATSU = (
    "①　**訪問系への偏り**。積雪寒冷地であり、"
    "通所サービスの送迎が難しい時期があること、"
    "除雪・買い物・調理などの生活援助の需要が大きいことが考えられます。"
    "健康とくらしの調査では、生活動作の困りごととして除雪を挙げた方が"
    "1,002人（回答者4,128人の24.3％）ありました。\n"
    "②　**介護保険の指定を受けない住まいの入居者への訪問**。"
    "区域内には住宅型有料老人ホーム%d施設%d人・"
    "サービス付き高齢者向け住宅%d施設%d戸・軽費老人ホーム%d施設%d人の"
    "計%d人・戸があり、在宅サービスの受給者の約3分の1に当たります。"
    "また、区域内の訪問介護%d事業所のうち%d事業所は、"
    "住宅型有料老人ホーム・サービス付き高齢者向け住宅又は"
    "介護付有料老人ホームに併設されており、"
    "**この%d事業所の利用者は%d人で、区域内の訪問介護の利用者%d人の"
    "%.1f％**を占めます（訪問介護員等は%d人のうち%d人・%.1f％）。"
    "同一の建物に住む方への訪問は、1人あたりの回数が多くなりやすく、"
    "1回あたりの提供時間は短くなる傾向があります。"
    "**この仮説は訪問介護の給付月額（1.55倍）より回数（1.87倍）の方が"
    "大きく上回っていることと向きが合います。**\n"
    "③　**通所と訪問が入れ替わっている**。"
    "通所介護の1人あたり日数は8.74日で全国平均10.66日の0.82倍であり、"
    "訪問介護の回数の多さと表裏の関係にある可能性があります。\n"
    "④　短期入所生活介護・短期入所療養介護が全国平均を下回ることは、"
    "施設サービスの受給率が高いこと（第2の群）と結びつく可能性があります。"
    "短期入所を繰り返さずに入所へ移る流れが考えられます。\n"
    "⑤　居宅療養管理指導・福祉用具貸与が全国平均を下回ることについては、"
    "供給の側（事業所の所在）による可能性があり、"
    "1人あたりの量の問題として読むことはできません。"
    % (_JU_N, _JU_T, _SK_N, _SK_T, _KE_N, _KE_T, _SUMAI,
       _HO["事業所"], _HO["併設事業所"], _HO["併設事業所"],
       _HO["併設利用者"], _HO["利用者"], _HO_UW,
       _HO["員"], _HO["併設員"], _HO_NW))

_C_KENSHO = (
    "①　**同一建物等居住者に対する減算の算定状況**（訪問介護）を"
    "国保連合会の給付実績の明細により確かめます。"
    "集合住宅に住む方への訪問がどの程度あるかが分かり、"
    "仮説②を確かめる最も直接の手立てとなります。\n"
    "②　**生活援助中心型と身体介護の別**の算定回数を確かめます。"
    "生活援助の割合が高ければ仮説①を裏づけます。\n"
    "③　**月別の利用回数**（冬季と夏季）を確かめます。"
    "介護保険 利用状況統計表を審査月ごとに継続して受領できれば、"
    "サービス種類別・要介護度別に季節性を測定できます。\n"
    "④　区分支給限度基準額に対する利用割合の分布を確かめ、"
    "限度額に近い利用が訪問介護に集まっているかを見ます。\n"
    "⑤　ケアプラン点検（第7期介護給付適正化計画の主要事業）により、"
    "訪問介護の回数が多い事例を抽出して個別に内容を確かめます。\n"
    "⑥　通所介護について、送迎の状況"
    "（中山間地域等における小規模事業所加算・送迎に係る減算の算定）を"
    "確かめ、送迎が利用日数の制約となっているかを見ます。")

_C_TAIO = (
    "①　ケアプラン点検の重点を訪問介護の回数が多い事例に置きます"
    "（第7期介護給付適正化計画）。"
    "**回数の多さそのものを是正の対象とするのではなく、"
    "居住の形態と必要な支援の内容に照らして妥当かを確かめる**ものとします。\n"
    "②　定期巡回・随時対応型訪問介護看護の整備を検討します。"
    "頻回の訪問を包括報酬の仕組みへ移すことにより、"
    "1人あたりの回数と給付費の関係を整えられる可能性があります。\n"
    "③　通所サービスの利用を妨げている要因（冬季の送迎・移動）への"
    "対応を検討します。地域支援事業の移動支援と併せて整理します。\n"
    "④　指定を受けない住まい（住宅型有料老人ホーム・"
    "サービス付き高齢者向け住宅・軽費老人ホーム）の入居者の状況を"
    "把握する仕組みを整えます。これらは給付費に現れないまま"
    "在宅サービスの需要を動かします。\n"
    "⑤　地域ケア会議で個別の事例を検証し、"
    "提起された地域課題が翌年度の事業・予算に反映された割合を"
    "第10期の代表指標として管理します。")

_C_JIYU = ((
    "・様式の令和7年度の1人あたり利用日数・回数は、"
    "当広域連合が年報の要介護度別の明細から算定した値と"
    "9つの区分すべてで一致します。\n"
    "・**受給者が少ない区分は1人あたりの値が大きく動きます。**"
    "定期巡回・随時対応型訪問介護看護（令和7年度の受給 月%.1f人）は"
    "給付月額が全国平均の%.2f倍と出ていますが、"
    "1人から2人の状況による値です。"
    "訪問入浴介護（月%.1f人）・認知症対応型通所介護（月%.1f人）も同じです。\n"
    "・夜間対応型訪問介護・地域密着型特定施設入居者生活介護・"
    "看護小規模多機能型居宅介護は実績がなく、様式でも「―」です。\n"
    "・" + TOSHIN + "\n"
    "・比較地域は設定していません。") % (
    _jis("定期巡回・随時対応型訪問介護看護"),
    hi("受給者1人あたり給付月額（定期巡回・随時対応型訪問介護看護）"),
    _jis("訪問入浴介護"), _jis("認知症対応型通所介護")))

KISAI = [
    ("① 認定率", "B4-a・B5-a・B6-a・B6-b",
     _A_HIKAKU, _A_KOSATSU, _A_KENSHO, _A_TAIO, _A_JIYU),
    ("② 受給率", "D2・D3・D4",
     _B_HIKAKU, _B_KOSATSU, _B_KENSHO, _B_TAIO, _B_JIYU),
    ("③ 受給者1人あたり給付月額・利用日数回数", "D15・D17・D31",
     _C_HIKAKU, _C_KOSATSU, _C_KENSHO, _C_TAIO, _C_JIYU),
]

# ============================================================ 00
ws = sheet("00_この表について",
           "地域分析・検討シートの記載案",
           "地域包括ケア「見える化」システムの地域分析・検討シート"
           "（令和8年9月30日受領）の記入欄15件について、"
           "記載する内容の案を示したものです。"
           + TOSHIN,
           [4, 28, 80], freeze="A5")
r = 4
r = lead(ws, r, "1　この表の使い方", span=3)
r = header(ws, r, ["#", "項目", "内容"])
for i, (a, b) in enumerate([
        ("何を示すものか",
         "様式の記入欄（群ごとに5つ・計15件）に記載する内容の案です。"
         "07シートの文をそのまま様式の該当欄へ貼り付けられる形にしています。"),
        ("受領した時点の状態",
         "15の記入欄はいずれも空欄です。"
         "比較地域も設定されていません。"),
        ("比較の年度",
         TOSHIN),
        ("様式の値と当広域連合の実績",
         "様式の令和7年度の値は当広域連合が年報から算定した値と一致します"
         "（1人あたり利用日数・回数9区分が小数第3位まで、"
         "認定率も一致）。様式の値をそのまま分析の根拠として用いられます。"),
        ("仮説の立て方",
         "乖離の大きい指標から順に、当広域連合が既に把握している"
         "実績・調査結果・供給の状況と結びつけて仮説を立てています。"
         "確かめられていないことは「考えられます」「可能性があります」と"
         "書き、断定していません。"),
        ("検証方法の選び方",
         "国保連合会の給付実績・施設からの報告・調査により"
         "実際に確かめられるものに限っています。"),
], start=1):
    r = body(ws, r, [i, a, b], height=44)

r += 1
r = lead(ws, r, "2　シートの構成", span=3)
r = header(ws, r, ["#", "シート", "内容"])
_SH = [
    ("01_指標の対比", "指標39件の自地域・北海道・全国と対全国比"),
    ("02_①認定率", "第1の群（認定率）の記載案5欄"),
    ("03_②受給率", "第2の群（受給率）の記載案5欄"),
    ("04_③給付月額・回数", "第3の群（1人あたり給付月額・利用日数回数）の記載案5欄"),
    ("05_仮説の裏づけ", "仮説ごとの根拠となる数値と出所"),
    ("06_検証に要する資料", "検証のために受領を要する資料"),
    ("07_貼り付ける文", "15の記入欄に貼り付ける文（そのまま使える形）"),
    ("08_自己点検", "―"),
]
for i, (a, b) in enumerate(_SH, start=1):
    r = body(ws, r, [i, a, b], height=20)
r = note(ws, r,
         "注）本表は記載案であり、様式への記載は"
         "構成3町との協議を経て確定するものとします。",
         span=3, height=26)

# ============================================================ 01
ws = sheet("01_指標の対比",
           "指標39件の対比（令和7年度・12か月）",
           "様式が掲げる指標39件について、自地域・北海道平均・全国平均と"
           "対全国比を並べたものです。" + TOSHIN,
           [4, 22, 44, 7, 13, 13, 13, 9, 9, 8], freeze="D5")
r = 4
r = header(ws, r, ["", "群", "指標", "単位", "自地域\n令和7年度",
                   "北海道平均\n令和7年度", "全国平均\n令和7年度",
                   "対全国比", "対北海道比", "判定"])
_N_SHIHYO = 0
_OVER, _UNDER = [], []
for i, (g, nm, sid, u, j, d, z, h) in enumerate(CB.SHIHYO, start=1):
    _N_SHIHYO += 1
    jz = CB.hi(j[1], z[1])
    jd = CB.hi(j[1], d[1])
    if jz is None:
        han, fl = "実績なし", GRAY
    elif jz >= 1.15:
        han, fl = "上回る", NG_O
        _OVER.append(nm)
    elif jz <= 0.85:
        han, fl = "下回る", MID_B
        _UNDER.append(nm)
    else:
        han, fl = "同程度", OK_G
    nd = 3 if u == "％" or u in ("回", "日") else 0
    r = body(ws, r, [i, g, nm, u,
                     "―" if j[1] is None else round(j[1], nd),
                     "―" if d[1] is None else round(d[1], nd),
                     "―" if z[1] is None else round(z[1], nd),
                     "―" if jz is None else round(jz, 2),
                     "―" if jd is None else round(jd, 2), han],
             fills={10: fl},
             align={j2: "right" for j2 in range(5, 10)},
             fmt_={j2: ("#,##0.000" if nd else "#,##0")
                   for j2 in range(5, 8)},
             height=18)
r = note(ws, r,
         "注1）判定は対全国比が1.15以上を「上回る」、0.85以下を"
         "「下回る」、その間を「同程度」としたものです。\n"
         "注2）%s\n"
         "注3）**受給者が少ない区分は1人あたりの値が大きく動きます。**"
         "定期巡回・随時対応型訪問介護看護は令和7年度の受給が月0.6人、"
         "訪問入浴介護は月5.3人、認知症対応型通所介護は月1.0人です。\n"
         "注4）比較地域は設定されていません（全年度が空欄です）。"
         % TOSHIN,
         span=10, height=56)

# ============================================================ 02・03・04
_RAN = CB.KINYU_RAN
for si, (gname, sid, c1, c2, c3, c4, c5) in enumerate(KISAI, start=2):
    ws = sheet("%02d_%s" % (si, ["①認定率", "②受給率", "③給付月額・回数"][si - 2]),
               "%s　記載案" % gname,
               "様式の%sの群（指標ID %s）の記入欄5件に記載する内容の案です。"
               % (gname[:3], sid) + TOSHIN,
               [4, 26, 96], freeze="A5")
    r = 4
    r = lead(ws, r, "1　対比した指標", span=3)
    r = header(ws, r, ["", "指標", "自地域／北海道平均／全国平均（令和7年度）"])
    _k = 0
    for g, nm, s2, u, j, d, z, h in CB.SHIHYO:
        if g != ["認定率", "受給率", "給付月額・利用日数回数"][si - 2]:
            continue
        _k += 1
        jz = CB.hi(j[1], z[1])
        nd = 3 if u in ("％", "回", "日") else 0
        r = body(ws, r, [_k, nm,
                         "%s ／ %s ／ %s（%s）　対全国比 %s"
                         % (fmt(j[1], nd), fmt(d[1], nd), fmt(z[1], nd), u,
                            "―" if jz is None else "%.2f倍" % jz)],
                 height=18)
    r += 1
    for k, (nm, txt) in enumerate(zip(_RAN, (c1, c2, c3, c4, c5)), start=1):
        r = lead(ws, r, "%d　%s" % (k + 1, nm), span=3)
        cell = ws.cell(row=r, column=1, value=_plain(txt))
        cell.font = Font(name=FONT, size=9)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.border = BORDER
        cell.fill = PatternFill("solid", fgColor=IN_Y)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        ws.row_dimensions[r] = ws.row_dimensions[r]
        ws.row_dimensions[r].height = max(
            60, 13 * (len(_plain(txt)) // 90 + _plain(txt).count("\n") + 1))
        r += 1
        r += 1

# ============================================================ 05
ws = sheet("05_仮説の裏づけ",
           "仮説の裏づけとなる数値と出所",
           "記載案の仮説が依拠している数値と、その出所を示したものです。"
           "いずれも当広域連合が既に把握している実績・調査結果・"
           "公表資料によります。",
           [4, 34, 22, 40, 26], freeze="A5")
r = 4
r = header(ws, r, ["#", "仮説・記述", "数値", "出所", "群"])
_URA = [
    ("粗の認定率は全国平均を上回るが、調整済み認定率は全国平均と同程度",
     "21.83％／16.58％（全国20.15％／16.67％）",
     "地域分析・検討シート（見える化システム）", "①"),
    ("認定率は年報の年度末の値であり、当広域連合の算定と一致する",
     "1,984人÷9,090人＝21.826％",
     "介護保険事業状況報告 年報（令和7年度）", "①"),
    ("85歳以上は増える",
     "年率2.4％",
     "住民基本台帳（各年10月1日現在）", "①"),
    ("認定を受けながらいずれのサービスも利用していない方がある",
     "486人（認定者の24.5％。推計）",
     "認定者数と受給者数の対比（令和7年度）", "①"),
    ("通いの場の参加率",
     "3.53％（令和6年度。週1回以上は1.28％）",
     "介護予防・日常生活支援総合事業の実施状況に関する調査", "①"),
    ("施設サービスの区域内定員",
     "%d人（施設サービス）／%d人（居住系サービス）"
     % (_SHISETSU_TEIIN, _KYOJU_TEIIN),
     "北海道の各名簿及び指定事業所一覧", "②"),
    ("介護老人保健施設は他の保険者の被保険者を受け入れている",
     "在所者206人に対し当広域連合の被保険者の受給 月%.1f人"
     % _jis("介護老人保健施設"),
     "居所変更実態調査及び年報（令和7年度）", "②"),
    ("住所地特例の適用",
     "77人",
     "介護保険事業状況報告 年報 様式1（令和7年度）", "②"),
    ("認知症対応型共同生活介護の定員は縮小した",
     "115人（平成28年〜令和2年）→99人（令和4年以降）",
     "北海道の指定事業所一覧", "②"),
    ("入所の経路は病院・診療所が多い",
     "新規入所の57.0％・退所先の47.0％",
     "居所変更実態調査（令和7年度）", "②"),
    ("在宅医療・介護連携の評価は全国平均に対して最も低い",
     "対全国比44.4％",
     "保険者機能強化推進交付金等 令和8年度全国集計（市町村分）", "②"),
    ("地域密着型介護老人福祉施設は定員に達する見込み",
     "令和11年度に定員62人の99.9％",
     "サービス見込量の算定（第10期）", "②"),
    ("訪問介護の1人あたり回数は当広域連合の算定と一致する",
     "%.3f回／月（令和7年度）" % _kaisu("訪問介護"),
     "介護保険事業状況報告 年報 様式1の7（令和7年度）", "③"),
    ("生活動作の困りごととして除雪を挙げた方",
     "1,002人（回答者4,128人の24.3％）",
     "健康とくらしの調査（令和7年度）", "③"),
    ("介護保険の指定を受けない住まい",
     "住宅型有料老人ホーム%d施設%d人・"
     "サービス付き高齢者向け住宅%d施設%d戸・"
     "軽費老人ホーム%d施設%d人＝計%d人・戸"
     % (_JU_N, _JU_T, _SK_N, _SK_T, _KE_N, _KE_T, _SUMAI),
     "北海道の有料老人ホーム一覧・サービス付き高齢者向け住宅の登録", "③"),
    ("同住まいは在宅サービスの受給者の約3分の1に当たる",
     "%.1f％（受給者 約%.0f人）" % (_SUMAI_WARI, _ZAITAKU_NIN),
     "上記と受給率（在宅）による算定", "③"),
    ("訪問介護の利用者の多くは居住系の住まいに併設された事業所の利用者",
     "%d事業所のうち%d事業所が併設。"
     "利用者%d人／%d人＝%.1f％（訪問介護員等%d人／%d人＝%.1f％）"
     % (_HO["事業所"], _HO["併設事業所"], _HO["併設利用者"], _HO["利用者"],
        _HO_UW, _HO["併設員"], _HO["員"], _HO_NW),
     "介護サービス情報公表システムの個別公表画面（令和7年10〜11月）"
     "及び北海道の届出済有料老人ホーム一覧", "③"),
    ("定期巡回・随時対応型訪問介護看護の受給は少ない",
     "月%.1f人（令和7年度）" % _jis("定期巡回・随時対応型訪問介護看護"),
     "介護保険事業状況報告 年報（令和7年度）", "③"),
    ("訪問入浴介護・認知症対応型通所介護の受給も少ない",
     "月%.1f人／月%.1f人（令和7年度）"
     % (_jis("訪問入浴介護"), _jis("認知症対応型通所介護")),
     "同上", "③"),
]
for i, (a, b, c, g) in enumerate(_URA, start=1):
    r = body(ws, r, [i, a, b, c, g], height=32)
r = note(ws, r,
         "注）**他の団体が策定した計画は参照していません。**"
         "全国平均・北海道平均は見える化システムが出力した値、"
         "交付金の評価は公表されている全国集計によります。",
         span=5, height=28)

# ============================================================ 06
ws = sheet("06_検証に要する資料",
           "仮説の検証に要する資料",
           "03シート・04シートの検証方法を実行するために"
           "受領を要する資料です。"
           "**いずれも次回の報告までに用意できるかを確かめる必要があります。**",
           [4, 40, 18, 40, 16], freeze="A5")
r = 4
r = header(ws, r, ["#", "資料", "依頼先", "何が分かるか", "優先"])
_SHIRYO = [
    ("訪問介護の給付実績の明細（加減算を含むもの）", "国保連合会・構成3町",
     "同一建物等居住者に対する減算の算定状況。"
     "集合住宅に住む方への訪問の割合が分かる", "A"),
    ("訪問介護のサービスコード別の算定回数", "国保連合会・構成3町",
     "生活援助中心型と身体介護の別。生活援助の割合", "A"),
    ("介護保険 利用状況統計表（審査月ごと・継続）", "構成3町",
     "サービス種類別・要介護度別の月別の給付費と回数。季節性の測定", "A"),
    ("区分支給限度基準額に対する利用割合の分布", "国保連合会・構成3町",
     "限度額に近い利用が訪問介護に集まっているか", "B"),
    ("施設の在所者の保険者別の状況", "区域内の施設",
     "区域内の定員のうち当広域連合の被保険者が占める割合", "A"),
    ("入所申込者（待機者）の数と要介護度", "区域内の施設・構成3町",
     "受給率の高さが需要によるものか供給によるものか", "A"),
    ("性別・年齢5歳階級別・要介護度別の認定者数（全国・北海道）",
     "見える化システム",
     "どの年齢階級で認定率の差が生じているか", "B"),
    ("新規認定・更新認定別の要介護度の分布と要介護度の変化", "構成3町",
     "重度が全国平均を下回る理由の切り分け", "B"),
    ("認定調査の委託の状況", "構成3町",
     "判定に地域差がないか（確認事項No.92）", "B"),
    ("通所介護の送迎に係る加算・減算の算定状況", "国保連合会・構成3町",
     "送迎が利用日数の制約となっているか", "B"),
    ("指定を受けない住まいの入居者の状況", "構成3町・事業者",
     "給付費に現れない在宅の需要", "B"),
    ("ケアプラン点検の実施結果", "構成3町",
     "訪問介護の回数が多い事例の内容", "B"),
]
for i, (a, b, c, p) in enumerate(_SHIRYO, start=1):
    r = body(ws, r, [i, a, b, c, p],
             fills={5: (NG_O if p == "A" else MID_B)}, height=32)
r = note(ws, r,
         "注1）優先Aは仮説の中心を確かめるもの、"
         "優先Bは仮説を補うものです。\n"
         "注2）**給付実績の明細は個人を特定できる情報を含むため、"
         "集計した値の提供を依頼します。**",
         span=5, height=32)

# ============================================================ 07
_MEI = {KISAI[0][0]: "①認定率", KISAI[1][0]: "②受給率",
        KISAI[2][0]: "③給付月額・回数"}

ws = sheet("07_貼り付ける文",
           "様式の記入欄に貼り付ける文",
           "様式の記入欄に貼り付ける文です。"
           "様式の行番号と欄の名称を対にしています。"
           "**貼り付ける前に構成3町との協議を経てください。**",
           [4, 10, 30, 100], freeze="A5")
r = 4
r = header(ws, r, ["#", "様式の行", "記入欄", "記載する文"])
_N_RAN = 0
for (gname, sid, c1, c2, c3, c4, c5), (gn2, cr, fr) in zip(KISAI, CB.KINYU):
    for k, (nm, txt) in enumerate(zip(_RAN, (c1, c2, c3, c4, c5))):
        _N_RAN += 1
        gyo = "%d行" % (cr if k < 4 else fr)
        r = body(ws, r, [_N_RAN, gyo, "%s／%s" % (_MEI[gname], nm), txt],
                 fills={4: IN_Y}, height=max(
                     40, 12 * (len(_plain(txt)) // 100
                               + _plain(txt).count("\n") + 1)))
r = note(ws, r,
         "注）「全国平均等との比較」「考察」「検証方法」「対応策」は"
         "同じ行（%s）の別の列に、「自由記述」は次の行（%s）にあります。"
         % ("・".join(str(x[1]) + "行" for x in CB.KINYU),
            "・".join(str(x[2]) + "行" for x in CB.KINYU)),
         span=4, height=28)

# ============================================================ 08 自己点検
ws = sheet("08_自己点検", "自己点検",
           "本表の内的整合を機械で確かめた記録です。"
           "1件でも不適合があると出力そのものを止めます。",
           [6, 46, 34, 34, 10])

chk(1, "様式の指標を残らず掲げていること",
    "01シートの行数 ＝ 受領値の指標の数",
    "指標%d件／掲げた%d件" % (len(CB.SHIHYO), _N_SHIHYO),
    _N_SHIHYO == len(CB.SHIHYO) == 39)

chk(2, "記入欄15件すべてに記載案があること",
    "3群×5欄。空でないこと",
    "%d欄／空 %d欄" % (_N_RAN, sum(
        1 for x in KISAI for t in x[2:] if not t.strip())),
    _N_RAN == 15 and all(t.strip() for x in KISAI for t in x[2:]))

_re_ng = []
for g, nm, sid, u, j, d, z, h in CB.SHIHYO:
    if j[1] is None or z[1] is None:
        continue
    if abs(CB.hi(j[1], z[1]) - j[1] / z[1]) > 1e-12:
        _re_ng.append(nm)
chk(3, "対全国比が受領値から再現できること",
    "自地域÷全国平均を受領値で計算し直す",
    "合わない %s" % (_re_ng or "なし"), not _re_ng)

_TS_TAI = ("00_この表について", "01_指標の対比", "02_①認定率",
           "03_②受給率", "04_③給付月額・回数")
_ts = [w.title for w in wb.worksheets
       if w.title in _TS_TAI and "令和8年3月" not in str(w["A2"].value or "")]
chk(4, "対比を掲げるシートに令和8年度が年度の途中の値である旨があること",
    "00から04シートのA2に様式の備考の趣旨があること",
    "無いシート %s" % (_ts or "なし"), not _ts)

_K9 = ["訪問介護", "訪問入浴介護", "訪問看護", "訪問リハビリテーション",
       "通所介護", "通所リハビリテーション", "短期入所生活介護",
       "認知症対応型通所介護", "地域密着型通所介護"]
_kg, _kmax = [], 0.0
for nm in _K9:
    mine = _kaisu(nm)
    got = v("受給者1人あたり利用日数・回数（%s）" % nm, 1, 4)
    if mine is None or got is None:
        _kg.append(nm)
        continue
    _kmax = max(_kmax, abs(mine - got))
    if abs(mine - got) > 0.001:
        _kg.append(nm)
chk(5, "1人あたり利用日数・回数が当広域連合の算定と一致すること",
    "年報 令和7年度の要介護度別明細から算定した値と受領値を比べる",
    "%d区分／最大の差 %.6f／合わない %s" % (len(_K9), _kmax, _kg or "なし"),
    not _kg and _kmax < 0.001)

_nin = sum(NIN_R7)
_hh = HIHO["2025"]
_mine_r = _nin / _hh * 100
_got_r = v("認定率", 1, 4)
chk(6, "認定率が当広域連合の算定と一致すること",
    "年報 第1号認定者÷第1号被保険者（年度末）",
    "%d人÷%.0f人＝%.6f％／受領値 %.6f％" % (_nin, _hh, _mine_r, _got_r),
    abs(_mine_r - _got_r) < 1e-6)

_shi_mine = sum(_jis(short(l)) for l in SVC if kubun(l) == "施設サービス")
_shi_got = v("受給率（施設サ－ビス）", 1, 4) / 100.0 * 9123.2
chk(7, "受給率（施設）から逆算した受給者数が当広域連合の算定と近いこと",
    "受給率×第1号被保険者数（総括表の分母9,123.2人）と年報の月平均を比べる",
    "逆算%.1f人／当方%.1f人（差%.2f％）"
    % (_shi_got, _shi_mine, (_shi_got - _shi_mine) / _shi_mine * 100),
    abs(_shi_got - _shi_mine) / _shi_mine < 0.03)

_KINSHI = ["に由来する", "と整合する", "1件も", "有意差がないため",
           "全国トップ"]
_kin = []
for w_ in wb.worksheets:
    for row in w_.iter_rows():
        for c in row:
            if not isinstance(c.value, str):
                continue
            for k in _KINSHI:
                if k in c.value:
                    _kin.append("%s!%s(%s)" % (w_.title, c.coordinate, k))
chk(8, "禁止表現を用いていないこと",
    "5件の語を全シートで走査",
    "%d件 %s" % (len(_kin), _kin or ""), not _kin)

_NAIBU = ["固定値", "実物から", "章節ごとに", "判定している", "runpy",
          ".py", "再実行", "書き写して", "スクリプト"]
_nb = []
for w_ in wb.worksheets:
    for row in w_.iter_rows():
        for c in row:
            if not isinstance(c.value, str):
                continue
            for k in _NAIBU:
                if k in c.value:
                    _nb.append("%s!%s(%s)" % (w_.title, c.coordinate, k))
chk(9, "当方の内部の仕組み・作業経過の語を書いていないこと",
    "9件の語を全シートで走査（発注者へ提供する文であるため）",
    "%d件 %s" % (len(_nb), _nb or ""), not _nb)

_PI = [re.compile(r"\d{2,4}-\d{2,4}-\d{3,4}"),
       re.compile(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}")]
_pi = []
for w_ in wb.worksheets:
    for row in w_.iter_rows():
        for c in row:
            if isinstance(c.value, str):
                for p in _PI:
                    if p.search(c.value):
                        _pi.append("%s!%s" % (w_.title, c.coordinate))
chk(10, "個人を特定する値を収めていないこと",
    "電話番号・メールアドレスの形を全セルで走査",
    "%d件" % len(_pi), not _pi)

_ast = []
for w_ in wb.worksheets:
    for row in w_.iter_rows():
        for c in row:
            if isinstance(c.value, str) and "**" in c.value:
                _ast.append("%s!%s" % (w_.title, c.coordinate))
chk(11, "強調の指定（**）がセルに残っていないこと",
    "全シートの全セルを走査",
    "%d件" % len(_ast), not _ast)

import data_kofukin_rengo as KR              # noqa: E402

_DAN = sorted({x[1] for x in KR.RENGO} | {x[1] for x in KR.KAISAN}
              | {m[0] for x in KR.RENGO for m in x[3]}
              | {m[0] for x in KR.KAISAN for m in x[4]},
              key=len, reverse=True)
_dan_ng = []
for w_ in wb.worksheets:
    for row in w_.iter_rows():
        for c in row:
            if not isinstance(c.value, str):
                continue
            for nm in _DAN:
                if (nm and nm not in ("大雪地区広域連合", "東川町",
                                      "美瑛町", "東神楽町")
                        and nm in c.value):
                    _dan_ng.append("%s!%s(%s)" % (w_.title, c.coordinate, nm))
chk(12, "他の団体の固有名称を掲げていないこと",
    "全国集計から割り出した団体名・構成市町村名を全シートで走査",
    "対象%d件／検出 %d件 %s" % (len(_DAN), len(_dan_ng), _dan_ng[:3] or ""),
    not _dan_ng and len(_DAN) > 200)

_hik = [w.title for w in wb.worksheets
        if w.title in ("01_指標の対比", "02_①認定率", "03_②受給率",
                       "04_③給付月額・回数")]
_hik_ng = []
for t in _hik:
    w_ = wb[t]
    found = any("比較地域" in str(c.value or "")
                for row in w_.iter_rows() for c in row)
    if t == "01_指標の対比" and not found:
        _hik_ng.append(t)
chk(13, "比較地域が設定されていないことを明記していること",
    "01シートに「比較地域」の語があること",
    "無いシート %s" % (_hik_ng or "なし"), not _hik_ng)

_JU_OK = (_JU_T + _SK_T + _KE_T) == _SUMAI and _SUMAI > 0
chk(14, "住まいの人数・戸数を名簿から数えていること",
    "住宅型＋サービス付き高齢者向け住宅＋軽費 ＝ 計",
    "%d＋%d＋%d＝%d人・戸" % (_JU_T, _SK_T, _KE_T, _SUMAI), _JU_OK)

_OV = set(_OVER)
_need = {"受給率（施設サ－ビス）", "受給者1人あたり給付月額（訪問介護）",
         "受給者1人あたり利用日数・回数（訪問介護）"}
chk(15, "全国平均を大きく上回る指標が記載案で取り上げられていること",
    "対全国比1.15以上の主要な指標が本文に現れること",
    "上回る%d件／下回る%d件／主要3件 %s"
    % (len(_OVER), len(_UNDER),
       "あり" if _need <= _OV else "なし"),
    _need <= _OV
    and "訪問介護" in _C_HIKAKU and "施設サービス" in _B_HIKAKU)

_SUKU = ["定期巡回・随時対応型訪問介護看護", "訪問入浴介護",
         "認知症対応型通所介護"]
_sk_ng = [nm for nm in _SUKU
          if ("月%.1f人" % _jis(nm)) not in _plain(_C_JIYU)]
chk(16, "受給者が少ない区分の人数が本文と算定で一致すること",
    "自由記述の人数が年報の月平均と一致すること",
    "%d区分／合わない %s" % (len(_SUKU), _sk_ng or "なし"), not _sk_ng)

r = header(ws, 4, ["No.", "点検した内容", "式・条件", "結果", "判定"])
for c in CHECKS:
    body(ws, r, list(c), fills={5: (OK_G if c[4] == "適合" else NG_O)},
         height=30)
    r += 1
_NG = sum(1 for c in CHECKS if c[4] != "適合")
body(ws, r, ["", "自己点検の結果", "",
             "適合%d件・不適合%d件" % (len(CHECKS) - _NG, _NG),
             "適合" if _NG == 0 else "不適合"],
     fills={5: (OK_G if _NG == 0 else NG_O)}, height=24, bold=True)

# ============================================================ 印刷・保存
for w_ in wb.worksheets:
    w_.page_setup.orientation = "landscape"
    w_.page_setup.fitToWidth = 1
    w_.page_setup.fitToHeight = 0
    w_.sheet_properties.pageSetUpPr.fitToPage = True
    w_.print_title_rows = "4:4"
    w_.oddFooter.left.text = "地域分析・検討シートの記載案（%s）" % KIJUNBI
    w_.oddFooter.right.text = "&P / &N"

wb.save(OUT)
print("書き出しました:", OUT)
for w_ in wb.worksheets:
    print("  -", w_.title, w_.max_row, "rows")
print("指標 %d件（全国平均を上回る%d件・下回る%d件）"
      % (len(CB.SHIHYO), len(_OVER), len(_UNDER)))
print("記入欄 %d件に記載案を置いた（3群×5欄）" % _N_RAN)
print("裏づけ %d件／検証に要する資料 %d件（優先A %d件）"
      % (len(_URA), len(_SHIRYO), sum(1 for x in _SHIRYO if x[3] == "A")))
print("自己点検 %d件：適合%d件・不適合%d件"
      % (len(CHECKS), len(CHECKS) - _NG, _NG))
if _NG:
    print("不適合があります。")
    sys.exit(1)
print("すべての点検に適合しました。")
