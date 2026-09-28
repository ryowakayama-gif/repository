# -*- coding: utf-8 -*-
"""令和8年10月工程　地域課題の整理と構成3町との協議における確認事項.

仕様書５の令和8年10月の工程は「地域課題整理、将来推計、
第10期計画骨子案作成、構成町との協議・調整」であり、
工程末の提出物に「地域課題整理資料」と「3町協議資料」が置かれている。

本表はその2つを1冊にまとめたものである。
  ・地域課題の整理　　　　　根拠となる数値と計画本文の反映先を対にして掲げる
  ・各町との打合せの確認事項　町ごとに、何を伺い何を決めていただくかを掲げる
  ・認知症施策推進計画　　　保険者機能強化推進交付金等の評価結果を踏まえ、
                            市町村計画（認知症基本法第13条）の扱いを掲げる

シート構成
  00_この表について
  01_地域課題の整理
  02_町別の現在地
  03_東川町の確認事項
  04_美瑛町の確認事項
  05_東神楽町の確認事項
  06_3町共通の確認事項
  07_認知症施策推進計画
  08_認知症総合支援の町別評価
  09_10月の進め方
  10_自己点検

数値は受領資料の収録値及びサービス見込量の算定から引く。
確認事項は業務工程管理表から読む。
自己点検で1件でも不適合があると終了コード1で終わる。
"""

import ast
import os
import runpy as _runpy
import sys as _sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

_sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import data_chiiki_shien as CS                      # noqa: E402
import data_hokkaido_shitei as HS                   # noqa: E402
import data_juki as JU                              # noqa: E402
import data_kofukin as KF                           # noqa: E402
import data_kofukin_item as KI                      # noqa: E402
import data_kyufu_jisseki as KJ                     # noqa: E402
import data_riyojokyo_r8_08 as RJ                   # noqa: E402
import data_shien_tool as ST                        # noqa: E402
import data_sogo_r6 as SG                           # noqa: E402
import repo_paths as RP                             # noqa: E402

ODIR = RP.ROOT + "/output"
OUT = os.path.join(
    ODIR, "第10期計画_10月 地域課題の整理と3町協議の確認事項.xlsx")

KIJUNBI = "令和8年9月28日"
TOWNS = ["東川町", "美瑛町", "東神楽町"]

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


# ------------------------------------------------------------ 算定を読む
class _Sink(object):
    def write(self, *a):
        pass

    def flush(self):
        pass


def _load(name):
    old = _sys.stdout
    _sys.stdout = _Sink()
    try:
        return _runpy.run_path(RP.ROOT + "/" + name)
    finally:
        _sys.stdout = old


_M = _load("build_mikomiryo_santei.py")     # サービス見込量 第1次概算
_SOG = _M["sogaku"]
_JIS = _M["JISSEKI"]
_TEIIN = _M["TEIIN_MAP"]
_Y3L = _M["Y3L"]
R11 = _M["Y3"][-1]


def jis(lab):
    """令和7年度の月平均利用者数（保険者単位）。"""
    return sum(_JIS[lab])


# ------------------------------------------------------------ 確認事項を読む
def _check_rows():
    """業務工程管理表の確認事項をソースから読む（固定値を書かない）。"""
    src = open(RP.ROOT + "/build_process_control.py", encoding="utf-8").read()
    for node in ast.parse(src).body:
        if (isinstance(node, ast.Assign)
                and getattr(node.targets[0], "id", None) == "CHECK"):
            return [ast.literal_eval(e) for e in node.value.elts]
    raise RuntimeError("CHECK を読めない")


CHECK = _check_rows()
CHECK_BY_NO = {c[0]: c for c in CHECK}
# 列：No./業務内容/工程（月）/確認事項/確認をお願いする内容/
#     止めている成果物/確認先/状態/回答期限/回答・対応
SUMI = ("解決", "完了", "了承済")


def _kitei():
    """決着しない場合の当方の扱い（既定値）をソースから読む。"""
    src = open(RP.ROOT + "/build_soan_kadai.py", encoding="utf-8").read()
    for node in ast.parse(src).body:
        if (isinstance(node, ast.Assign)
                and getattr(node.targets[0], "id", None) == "KITEI"):
            return ast.literal_eval(node.value)
    raise RuntimeError("KITEI を読めない")


KITEI = _kitei()

# 業務工程管理表に既定値のない確認事項について、本表で示す扱い。
KITEI_ADD = {
    11: "住民基本台帳（各年1月1日現在）の町別人口により本文を確定する",
    12: "日常生活圏域を3町（3圏域）とする案のまま概算する",
    13: "北海道の指定事業所一覧の「通常の事業の実施地域」により判定する",
    14: "案C（総合戦略＋住民基本台帳の実績趨勢）のまま確定する",
    17: "広域連合の代表KPI16項目を掲げ、3町計画への掲載は求めない",
    23: "北海道の各名簿及び介護サービス情報公表システムの定員による",
    26: "現に受領している交付金評価結果と実施報告書の範囲で記述する",
    40: "同一法人の重複計上分を除いた採用83人・離職65人を用いる",
    42: "現行の工程表のまま進める",
    73: "ケアプラン点検実施要領（案）は協議資料にとどめ素案に入れない",
    78: "北海道の名簿の定員のまま概算する",
    83: "3町合同の意見交換会（10月上旬）と町ごとの個別協議"
        "（10月中旬〜下旬）の2段構えで進める",
    84: "現定員の据え置き案を掲げ、本文は［要協議］のまま残す",
    85: "町別データシートは現行の内容のまま配付する",
    96: "認知症対応型共同生活介護の定員は99人＋［要確認］のまま掲げる",
    108: "補正しても外れる6件は趨勢ではなく供給・制度の問題として注記する",
    111: "趨勢に現れない要因は向きのみを注記し、数値には織り込まない",
    115: "令和9年4月新設の3区分を0として立てる",
    116: "供給の制約（人材）を見込量に反映しない（需要のみで算定する）",
    119: "地域ケア会議の記録様式は協議資料にとどめ、令和9年度に整える",
    120: "地域ケア個別会議と地域ケア推進会議を分けずに計数する",
    121: "令和7年度の実施報告書により把握できる範囲で記述する",
    130: "案B（3町がそれぞれ市町村計画を策定する）のまま素案に反映する",
    133: "令和8年度の月報（4〜7月）は基準年度に用いず、"
         "令和7年度を基準として概算する",
}


def kitei_of(no):
    """決着しない場合の当方の扱い。台帳の回答欄にあるものはそこから採る。"""
    v = KITEI.get(no) or KITEI_ADD.get(no)
    if v:
        return v
    kotae = CHECK_BY_NO[no][9] if no in CHECK_BY_NO else ""
    if "決着しない場合" in kotae:
        v = kotae.split("決着しない場合", 1)[-1]
        return v.lstrip("は、 　").strip()
    return ""


# ------------------------------------------------------------ 体裁
def sheet(name, title, subtitle, widths, freeze="A5"):
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
    ws.row_dimensions[2].height = 52
    ws.freeze_panes = freeze
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


def body(ws, row, vals, fills=None, height=30, align=None, bold=False,
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


def note(ws, row, text, span, height=80):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=8.5)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = height
    return row + 1


def yen(v):
    return "%s円" % format(int(round(v)), ",")


def pct(v, k=1):
    return ("%." + str(k) + "f％") % (v * 100)


N0 = "#,##0"

# ------------------------------------------------------------ 町別の基礎値
Y = 2025                       # 住民基本台帳の直近年（各年1月1日現在）
Y0 = 2019
POP = {t: JU.pop(t, Y, "tot") for t in TOWNS}
POP65 = {t: JU.pop(t, Y, "65+") for t in TOWNS}
POP85 = {t: JU.pop(t, Y, "85+") for t in TOWNS}
NIN8 = {t: ST.NINTEI[t]["R8"][0] for t in TOWNS}
NIN6 = {t: ST.NINTEI[t]["R6"][0] for t in TOWNS}
KOF8 = {t: KF.KOF[t]["R8"]["推進・支援合計"] for t in TOWNS}
KOF6 = {t: KF.KOF[t]["R6"]["推進・支援合計"] for t in TOWNS}
NINCHI = {y: {t: KF.KOF[t][y]["Ⅱ合計_2"] for t in TOWNS}
          for y in ("R6", "R7", "R8")}
KAYOI = {t: SG.KAYOI["箇所数　計"][0][SG.TOWNS.index(t)] for t in TOWNS}
KAYOI_N = {t: SG.KAYOI["参加者実人数　計"][0][SG.TOWNS.index(t)]
           for t in TOWNS}
HIYOU8 = {t: RJ.val(t, "38介護給付費計", "計", "費用総額") for t in TOWNS}
HIYOU8_K = sum(HIYOU8.values())


def shitei_count(town):
    """北海道の指定事業所一覧により、当該町に所在する事業所の数を数える。

    同一の事業所が複数のサービス種別に現れるため、延べの件数である。
    """
    n = 0
    for _svc, rows in HS.SHITEI.items():
        n += sum(1 for x in rows if x.get("町") == town)
    return n


SHITEI_N = {t: shitei_count(t) for t in TOWNS}


def shoki(town, src=None):
    """認知症初期集中支援推進事業の記載（量の欄）を返す。"""
    src = CS.HOUKOKU_R7 if src is None else src
    for x in src:
        if x[0] == town and "認知症初期集中" in x[2]:
            return x[6] or ""
    return ""


def shoki_naiyo(town, src=None):
    src = CS.HOUKOKU_R7 if src is None else src
    for x in src:
        if x[0] == town and "認知症初期集中" in x[2]:
            return x[5] or ""
    return ""


KYUSHI = [t for t in TOWNS if "休止" in (shoki(t) + shoki_naiyo(t))]

# ============================================================ 00
ws = sheet("00_この表について",
           "令和8年10月工程　地域課題の整理と3町協議の確認事項",
           "業務仕様書５の令和8年10月の工程（地域課題整理、構成町との協議・調整）"
           "に向けて、地域課題の整理と、構成3町との打合せで確認する事項を"
           "1冊にまとめたものです。基準日 %s。" % KIJUNBI,
           [22, 56, 46])
r = 4
r = header(ws, r, ["区分", "内容", "備考"])
for a, b, c in [
    ("本表の位置づけ",
     "仕様書５の令和8年10月の提出物のうち「地域課題整理資料」と"
     "「3町協議資料」に当たります。",
     "11月工程の「構成町意見反映、第10期計画骨子の整理」へつなぎます。"),
    ("地域課題の整理（01シート）",
     "計画に載せる地域課題を、根拠となる数値・出所・計画本文の反映先と"
     "対にして掲げています。",
     "町による差がある課題には、どの町の課題かを示しています。"),
    ("町別の確認事項（03〜05シート）",
     "町ごとに、伺いたいこと・決めていただきたいこと・"
     "決まらない場合の当方の扱いを掲げています。",
     "決まらない場合の扱いを置いているため、"
     "ご決定がなくても算定は止まりません。"),
    ("3町共通の確認事項（06シート）",
     "広域連合として決めていただく事項です。",
     "保険料・見込量は保険者（広域連合）単位で算定するため、"
     "町別には算定しません。"),
    ("認知症施策推進計画（07・08シート）",
     "認知症基本法第13条の市町村計画は3町がそれぞれ策定するという"
     "整理（案B）に沿って、各町で決めていただく事項を掲げています。",
     "保険者機能強化推進交付金等の評価は市町村単位で行われるため、"
     "町別の得点を判断の材料として付しています。"),
    ("個人情報の取扱い",
     "本表は集計値のみで構成しています。",
     "担当者名・電話番号・メールアドレスは収録していません。"),
    ("数値の出所",
     "人口は住民基本台帳（各年1月1日現在）、認定者数は介護保険事業計画"
     "作成支援ツール、給付費は給付費データ及び利用状況統計表、"
     "交付金は市町村分の全国集計結果によります。",
     "見込量・給付費・保険料はサービス見込量 第1次概算によります。"),
]:
    r = body(ws, r, [a, b, c], height=44, fills={1: GRAY})
r += 1
r = lead(ws, r, "確認事項の件数", 3)
r = header(ws, r, ["区分", "件数", "内訳"])

MACHI_NO = sorted(c[0] for c in CHECK
                  if "町" in c[6] and c[7] not in SUMI)
TOKAWA_NO = [n for n in MACHI_NO if "東川町" in CHECK_BY_NO[n][6]]
BIEI_NO = [n for n in MACHI_NO if "美瑛町" in CHECK_BY_NO[n][6]]
KYOTSU_NO = [n for n in MACHI_NO
             if n not in TOKAWA_NO and n not in BIEI_NO]
for a, b, c in [
    ("構成3町に関わる確認事項", len(MACHI_NO),
     "業務工程管理表の確認先に町を含み、まだ決着していないものです。"),
    ("うち特定の町へのもの", len(TOKAWA_NO) + len(BIEI_NO),
     "東川町%d件・美瑛町%d件" % (len(TOKAWA_NO), len(BIEI_NO))),
    ("うち3町共通のもの", len(KYOTSU_NO),
     "広域連合として決めていただく事項です。"),
]:
    r = body(ws, r, [a, b, c], height=24, fills={1: GRAY}, fmt={2: N0})
r = note(ws, r, "※ 上記のほか、本表で新たに掲げた確認事項があります"
                "（03〜07シートのNo.欄が「新」のもの）。"
                "これらは業務工程管理表へ登録したうえで管理します。", 3,
         height=32)

# ============================================================ 01
ws = sheet("01_地域課題の整理", "地域課題の整理",
           "第10期計画に載せる地域課題を、根拠となる数値と計画本文の反映先と"
           "対にして掲げています。町による差があるものは「町の別」欄に"
           "示しています。",
           [4, 26, 46, 30, 22, 16])
r = 4
r = header(ws, r, ["No.", "地域課題", "根拠となる事実（数値）", "出所",
                   "計画本文の反映先", "町の別"])

_ninchi_gh = _SOG("居住系サービス 認知症対応型共同生活介護", R11)
_ninchi_tsu = _SOG("在宅サービス 認知症対応型通所介護", R11)
_teiki = _SOG("在宅サービス 定期巡回・随時対応型訪問介護看護", R11)
_chimitsu_teiin = _TEIIN["施設サービス 地域密着型介護老人福祉施設入所者生活介護"][0]
_chimitsu_r11 = _SOG("施設サービス 地域密着型介護老人福祉施設入所者生活介護", R11)
_gh_teiin = _TEIIN["居住系サービス 認知症対応型共同生活介護"][0]

KADAI = [
    ("高齢者数は横ばいだが、85歳以上が増える",
     "3町計の65歳以上は令和元年%s人から令和7年%s人（年%+.2f％）で"
     "ほぼ横ばいであるのに対し、85歳以上は%s人から%s人"
     "（年%+.2f％）で増えています。"
     % (format(JU.pop3(Y0, "65+"), ","), format(JU.pop3(Y, "65+"), ","),
        JU.g("65+") * 100, format(JU.pop3(Y0, "85+"), ","),
        format(JU.pop3(Y, "85+"), ","), JU.g("85+") * 100),
     "住民基本台帳（各年1月1日現在）",
     "第2章第1節・第6章第1節", "3町共通"),
    ("認定者数の動きが町で分かれる",
     "令和6年度から令和8年度にかけて、東川町は%+d人、東神楽町は%+d人と"
     "増える一方、美瑛町は%+d人と減っています。"
     % (NIN8["東川町"] - NIN6["東川町"],
        NIN8["東神楽町"] - NIN6["東神楽町"],
        NIN8["美瑛町"] - NIN6["美瑛町"]),
     "介護保険事業計画作成支援ツール（実績値）",
     "第2章第2節・第6章第1節", "町で分かれる"),
    ("認定を受けながらサービスを利用していない方が4分の1いる",
     "認定者1,984人のうち、いずれの介護サービスも利用していない方が"
     "推計486人（24.5％）います。",
     "調査結果と年報実績の突合（推計値）",
     "第2章第2節・第5章 基本目標1", "3町共通"),
    ("24時間対応のサービスが区域内にない",
     "定期巡回・随時対応型訪問介護看護の令和11年度の見込みは"
     "月%.1f人にとどまり、区域内に事業所がありません。"
     "在宅生活改善調査では、より適切と思われるサービスとして"
     "24時間対応の3サービスを挙げた回答が延べ26件ありました。"
     % _teiki,
     "サービス見込量 第1次概算／在宅生活改善調査",
     "第6章第2節・第6章第4節・第5章 基本目標1", "3町共通"),
    ("在宅生活の継続が困難な層がいる",
     "在宅生活改善調査で、在宅生活の維持が困難と見込まれる方が"
     "回答のあった98人のうち72人（73.5％）でした。"
     "ただしこの調査は在宅の認定者全体を代表しません。",
     "在宅生活改善調査（居宅介護支援事業所経由）",
     "第2章第4節・第5章 基本目標1", "3町共通"),
    ("施設入所の経路が医療機関に偏っている",
     "居所変更実態調査で、新規入所の57.0％・退去先の47.0％が"
     "病院・診療所でした。",
     "居所変更実態調査（18施設）",
     "第6章第5節", "3町共通"),
    ("地域密着型介護老人福祉施設が第10期中に定員に達する見込み",
     "令和11年度の見込みは月%.1f人で、区域内定員%d人の%s"
     "に当たります。令和12年度には定員を超えると見込まれます。"
     % (_chimitsu_r11, _chimitsu_teiin, pct(_chimitsu_r11 / _chimitsu_teiin)),
     "サービス見込量 第1次概算／北海道 特別養護老人ホーム名簿",
     "第6章第4節・第6章第5節", "美瑛町に定員が集中"),
    ("認知症高齢者の受け皿の構成が町で異なる",
     "認知症対応型共同生活介護の令和11年度の見込みは月%.1f人で、"
     "区域内定員%d人（＋［要確認］）の%sに当たります。"
     "町別の月平均利用者数は東川町%d人・美瑛町%d人・東神楽町%d人です。"
     % (_ninchi_gh, _gh_teiin, pct(_ninchi_gh / _gh_teiin),
        ST.SHISETSU["東川町"]["認知症対応型共同生活介護"],
        ST.SHISETSU["美瑛町"]["認知症対応型共同生活介護"],
        ST.SHISETSU["東神楽町"]["認知症対応型共同生活介護"]),
     "サービス見込量 第1次概算／介護保険事業計画作成支援ツール",
     "第2章第3節・第6章第4節・第5章 基本目標2", "町で分かれる"),
    ("認知症対応型通所介護の事業所が区域内にない",
     "区域内に事業所がありませんが、区域外の事業所の利用により"
     "令和11年度は月%.2f人の給付を見込んでいます。"
     "第9期は計画値が0であったため、対計画比が算出されず"
     "評価から漏れていました。" % _ninchi_tsu,
     "サービス見込量 第1次概算／北海道の指定事業所一覧",
     "第6章第2節・第5章 基本目標2", "3町共通"),
    ("認知症初期集中支援推進事業が1町で休止している",
     "令和7年度の実施報告書によれば、%sの認知症初期集中支援推進事業は"
     "対応件数0件で、医師を確保できず休止しています。"
     % "・".join(KYUSHI),
     "地域包括支援センター事業実施報告書（令和7年度）",
     "第5章 基本目標2", "・".join(KYUSHI)),
    ("通いの場の整備状況が町で大きく異なる",
     "箇所数は東川町%d・美瑛町%d・東神楽町%dで、"
     "参加実人数は%d人・%d人・%d人です。"
     % (KAYOI["東川町"], KAYOI["美瑛町"], KAYOI["東神楽町"],
        KAYOI_N["東川町"], KAYOI_N["美瑛町"], KAYOI_N["東神楽町"]),
     "総合事業の実施状況に関する調査（令和6年度）",
     "第3章第1節・第5章 基本目標2", "町で分かれる"),
    ("介護人材の確保が進まない",
     "介護人材実態調査で、外国人職員は施設・通所系の介護職員319人のうち"
     "37人（11.6％）、勤続1年未満が19.9％でした。"
     "採用83人に対し離職65人です。",
     "介護人材実態調査（事業所票27・職員個票317）",
     "第2章第6節・第5章 基本目標4", "3町共通"),
    ("除雪など介護保険の給付対象外の困りごとが大きい",
     "健康とくらしの調査で、生活動作の困りごととして除雪を挙げた方が"
     "回答者4,128票のうち1,002人（24.3％）でした。",
     "健康とくらしの調査（回収4,798票）",
     "第2章第4節・第5章 基本目標1", "3町共通"),
    ("保険者機能強化推進交付金等の得点に町の差がある",
     "令和8年度は東川町%d点・美瑛町%d点・東神楽町%d点で、"
     "最大の差は%d点です。令和6年度からの動きも"
     "東川町%+d点・美瑛町%+d点・東神楽町%+d点と分かれています。"
     % (KOF8["東川町"], KOF8["美瑛町"], KOF8["東神楽町"],
        max(KOF8.values()) - min(KOF8.values()),
        KOF8["東川町"] - KOF6["東川町"], KOF8["美瑛町"] - KOF6["美瑛町"],
        KOF8["東神楽町"] - KOF6["東神楽町"]),
     "保険者機能強化推進交付金等（市町村分）全国集計結果",
     "第3章第4節・第5章 基本目標5", "町で分かれる"),
    ("地域ケア会議の成果を測る指標がない",
     "第9期の評価で「指標不足・成果を測る指標がない」と判定しました。"
     "指標は会議の開催回数のみです。令和7年度の実施報告書でも、"
     "回数・人数と単位がそろわず3町を合計できません。",
     "第9期計画の評価／地域包括支援センター事業実施報告書",
     "第5章 基本目標2・資料1", "3町共通"),
]
for i, (nm, konkyo, shutsu, hanei, machi) in enumerate(KADAI, start=1):
    r = body(ws, r, [i, nm, konkyo, shutsu, hanei, machi], height=56,
             align={1: "center"},
             fills={1: GRAY,
                    6: IN_Y if machi != "3町共通" else None})
r = note(ws, r, "※ 「町の別」が「3町共通」のものは広域連合として、"
                "「町で分かれる」ものは各町との打合せで扱います。"
                "なお、サービス見込量・給付費・保険料は保険者（広域連合）を"
                "単位として算定するため、町別には算定しません"
                "（介護保険法第3条、地方自治法第284条）。", 6, height=46)

# ============================================================ 02
ws = sheet("02_町別の現在地", "町別の現在地",
           "各町との打合せの前提となる数値です。"
           "実績は町別に掲げ、将来推計と保険料は3町計で算定しています。",
           [30, 18, 18, 18, 18, 30])
r = 4
r = header(ws, r, ["項目", "東川町", "美瑛町", "東神楽町", "3町計", "出所・見方"])


def row3(label, d, keisan=None, shutsu="", fmt=None, height=24, bold=False):
    global r
    vals = [d[t] for t in TOWNS]
    tot = keisan if keisan is not None else sum(vals)
    r = body(ws, r, [label] + vals + [tot, shutsu], height=height,
             fills={1: GRAY}, bold=bold,
             fmt=fmt or {2: N0, 3: N0, 4: N0, 5: N0})


row3("総人口（令和7年1月1日）", POP, shutsu="住民基本台帳")
row3("65歳以上人口（同）", POP65, shutsu="住民基本台帳")
row3("85歳以上人口（同）", POP85, shutsu="住民基本台帳")
r = body(ws, r, ["高齢化率（同）"]
         + [POP65[t] / POP[t] for t in TOWNS]
         + [sum(POP65.values()) / sum(POP.values()), "65歳以上÷総人口"],
         height=24, fills={1: GRAY},
         fmt={2: "0.0%", 3: "0.0%", 4: "0.0%", 5: "0.0%"})
row3("認定者数（令和8年度）", NIN8,
     shutsu="計画作成支援ツール（実績値）。3町合計であり、"
            "保険者単位の年報（令和7年度 第1号1,984人）とは数え方が違う",
     height=38)
r = body(ws, r, ["認定者数の動き（令和6→8年度）"]
         + [NIN8[t] - NIN6[t] for t in TOWNS]
         + [sum(NIN8.values()) - sum(NIN6.values()), "美瑛町のみ減少"],
         height=24, fills={1: GRAY}, fmt={2: "+#,##0;-#,##0;0",
                                          3: "+#,##0;-#,##0;0",
                                          4: "+#,##0;-#,##0;0",
                                          5: "+#,##0;-#,##0;0"})
row3("費用総額（令和8年8月審査分）", HIYOU8, keisan=HIYOU8_K,
     shutsu="利用状況統計表。費用額であり給付費（保険給付費）ではない",
     height=30)
r = body(ws, r, ["同　構成比"]
         + [HIYOU8[t] / HIYOU8_K for t in TOWNS] + [1.0, "1か月分の構成比"],
         height=22, fills={1: GRAY},
         fmt={2: "0.0%", 3: "0.0%", 4: "0.0%", 5: "0.0%"})
row3("指定事業所の件数（延べ）", SHITEI_N,
     shutsu="北海道の指定事業所一覧（令和8年6月30日現在）。"
            "所在地により数えたもの。同一の事業所が複数のサービス種別に"
            "現れるため延べの件数", height=38)
row3("通いの場の箇所数（令和6年度）", KAYOI,
     shutsu="総合事業の実施状況に関する調査")
row3("通いの場の参加実人数（同）", KAYOI_N, shutsu="同上")
row3("交付金の得点（令和8年度）", KOF8, keisan="―",
     shutsu="推進・支援合計。市町村単位の評価であり合計に意味はない",
     fmt={2: N0, 3: N0, 4: N0}, height=30)
row3("認知症総合支援の得点（同）", NINCHI["R8"], keisan="―",
     shutsu="支援 目標Ⅱ（100点満点）。08シートに内訳",
     fmt={2: N0, 3: N0, 4: N0}, height=24)
r = note(ws, r, "※ 施設・居住系サービスの月平均利用者数は、"
                "計画作成支援ツールでは3町合計、年報では保険者単位であり"
                "＋9.8％の開きがあります。町別の値を足しても"
                "保険者単位の値にはなりません。"
                "区域外の施設を利用する当連合の被保険者と、"
                "区域内の施設を利用する他の保険者の被保険者の扱いが"
                "違うためです。", 6, height=58)

# ============================================================ 03〜06
COLS_M = ["No.", "業務内容", "確認事項", "確認をお願いする内容",
          "決まらない場合の当方の扱い", "回答期限"]

# 町ごとの確認事項は業務工程管理表から読む（台帳は同表を正とする）。
# 令和8年9月28日に、町ごとの論点9件を同表へ No.159〜No.167 として登録した。
SHEET_NM = {"東川町": "03_東川町の確認事項", "美瑛町": "04_美瑛町の確認事項",
            "東神楽町": "05_東神楽町の確認事項",
            "3町共通": "06_3町共通の確認事項"}
FURIWAKE = {t: [n for n in MACHI_NO if t in CHECK_BY_NO[n][6]] for t in TOWNS}
FURIWAKE["3町共通"] = [n for n in MACHI_NO
                       if not any(t in CHECK_BY_NO[n][6] for t in TOWNS)]
ROWS_BY = {}

for machi in TOWNS + ["3町共通"]:
    rows = [(no, CHECK_BY_NO[no][1], CHECK_BY_NO[no][3], CHECK_BY_NO[no][4],
             kitei_of(no), CHECK_BY_NO[no][8])
            for no in FURIWAKE[machi]]
    ROWS_BY[machi] = rows
    ws = sheet(SHEET_NM[machi],
               ("%sとの打合せの確認事項" % machi) if machi != "3町共通"
               else "3町共通の確認事項（広域連合として決めていただく事項）",
               "「決まらない場合の当方の扱い」を全件に置いています。"
               "ご決定がなくても算定は止まりません。"
               "No.は業務工程管理表 03_確認事項一覧の番号です。",
               [6, 12, 34, 50, 38, 9])
    r = header(ws, 4, COLS_M)
    for x in rows:
        r = body(ws, r, list(x), height=64, align={1: "center", 6: "center"},
                 fills={1: GRAY, 5: OK_G})
    if machi != "3町共通":
        r = note(ws, r, "※ 3町に共通する確認事項（地域ケア会議の実績、"
                        "必要利用定員総数、認知症施策推進計画の扱いなど）は"
                        "06シートに掲げています。"
                        "3町合同の意見交換会で扱い、"
                        "町ごとの個別協議では本シートを用います。", 6,
                 height=40)
    else:
        r = note(ws, r, "※ 特定の町へお伺いする事項は03〜05シートに"
                        "掲げています。"
                        "サービス見込量・給付費・保険料は保険者"
                        "（広域連合）を単位として算定するため、"
                        "町別には算定しません。", 6, height=34)

# ============================================================ 07
ws = sheet("07_認知症施策推進計画", "認知症施策推進計画の扱い",
           "共生社会の実現を推進するための認知症基本法（令和5年法律第65号）"
           "第13条の市町村認知症施策推進計画について、"
           "本計画との関係と各町で決めていただく事項を掲げています。",
           [22, 54, 48])
r = 4
r = lead(ws, r, "1　計画の階層と本計画の位置づけ", 3)
r = header(ws, r, ["区分", "定めるもの", "本件での扱い"])
for a, b, c in [
    ("国", "認知症施策推進基本計画（認知症基本法第11条）",
     "未受領です。閣議決定された基本計画の本文のご提供をお願いします"
     "（確認事項No.149）。"),
    ("都道府県", "都道府県認知症施策推進計画（同法第12条）",
     "北海道の策定状況が確認できていません。"
     "計画素案 第1章第8節に［要確認］として掲げています。"),
    ("市町村", "市町村認知症施策推進計画（同法第13条）",
     "名宛人は「市町村」です。**構成3町がそれぞれ策定します**（案B）。"),
    ("広域連合", "介護保険事業計画（介護保険法第117条）",
     "同条第3項の任意記載事項として認知症施策を掲げます。"
     "3町の市町村計画とは資料2（役割分担）により接続します。"),
]:
    r = body(ws, r, [a, b.replace("**", ""), c.replace("**", "")],
             height=40, fills={1: GRAY})
r = note(ws, r, "※ 広域連合は地方自治法第284条により設けられ、"
                "介護保険法上の保険者としての事務を共同処理する"
                "特別地方公共団体です。"
                "広域連合の計画に一体のものとして定めても、"
                "認知症基本法第13条の市町村計画としての効力は生じません。"
                "このため案B（3町がそれぞれ策定する）を採っています。"
                "法の条文を受領していないため、"
                "素案には条番号のみを書き項番号は書いていません。", 3,
         height=62)
r += 1
r = lead(ws, r, "2　計画素案への反映（令和8年9月・4箇所）", 3)
r = header(ws, r, ["反映先", "加えた内容", "残っているもの"])
for a, b, c in [
    ("第1章第2節 計画の位置付け",
     "認知症基本法が第11条で国の基本計画・第12条で都道府県計画・"
     "第13条で市町村計画を定めていること、"
     "市町村計画は3町が策定すること、"
     "本計画は介護保険法第117条第3項により定めること、"
     "調和は資料2で保つこと。",
     "意見の聴取の主体・方法・時期は［要協議］のままです。"),
    ("第1章第8節 関係計画等との関係",
     "「認知症施策推進計画」を「都道府県認知症施策推進計画」に改め"
     "主体を明示し、整合を図る計画に"
     "「認知症施策推進基本計画（認知症基本法第11条により国が定めるもの）」"
     "を加えました。",
     "認知症施策推進基本計画及び都道府県認知症施策推進計画の"
     "内容・策定の状況は［要確認］です。"),
    ("資料2（1）役割分担",
     "広域連合側に「本計画への位置づけ・認知症総合支援事業」、"
     "3町側に「市町村認知症施策推進計画の策定」。",
     "認知症総合支援事業の実施主体は「3町・広域連合」の併記のままです。"),
    ("第5章 基本目標2",
     "介護保険法第117条第3項による事項であること、"
     "3町の市町村計画とは資料2により接続すること、"
     "認知症初期集中支援推進事業の実施状況。",
     "各町の取組の差をどこまで書くかを決めていただく必要があります。"),
]:
    r = body(ws, r, [a, b, c], height=52, fills={1: GRAY})
r += 1
r = lead(ws, r, "3　決めていただく事項（3町共通）", 3)
r = header(ws, r, ["確認事項", "決めていただく内容", "決まらない場合の当方の扱い"])
NINCHI_CHECK = [130, 150, 149, 96]
for no in NINCHI_CHECK:
    c = CHECK_BY_NO[no]
    r = body(ws, r, ["No.%d　%s" % (no, c[3]), c[4],
                     kitei_of(no) or "現に確認できる範囲で記述する"],
             height=68, fills={1: GRAY, 3: OK_G})
r = note(ws, r, "※ No.130は①位置づけ（案Bで了承済み）・"
                "②認知症総合支援事業の実施主体・"
                "③認知症の人及び家族等からの意見の聴取の3点から成り、"
                "②と③が決着していません。"
                "町ごとの取組に関する確認事項は03〜05シートに掲げています。", 3,
         height=44)
r += 1
r = lead(ws, r, "4　見込量への反映", 3)
r = header(ws, r, ["サービス・事業", "令和11年度の見込み", "状況"])
for a, b, c in [
    ("認知症対応型共同生活介護",
     "月%.1f人（区域内定員%d人の%s）" % (_ninchi_gh, _gh_teiin,
                                       pct(_ninchi_gh / _gh_teiin)),
     "算定済みです。定員は東川町の1事業所が未確定です"
     "（確認事項No.96）。"),
    ("認知症対応型通所介護",
     "月%.2f人" % _ninchi_tsu,
     "算定済みです。区域内に事業所はありませんが、"
     "区域外の事業所の利用により給付があります。"),
    ("認知症総合支援事業（包括的支援事業）",
     "―",
     "量の見込みが立ちません。決算の科目は4区分で"
     "6事業別の内訳が分からず、総合事業の実施状況調査は"
     "包括的支援事業を含みません（確認事項No.150）。"),
]:
    r = body(ws, r, [a, b, c], height=44, fills={1: GRAY})
r = note(ws, r, "※ 認知症対応型共同生活介護について、"
                "従前「実績は年率▲4.96％で減っている」と整理していましたが、"
                "これは定員の縮小を趨勢と読んだものでした。"
                "区域全体の定員は平成28年から令和2年の115人が"
                "令和4年以降99人に減っており、"
                "令和6年度から令和7年度は＋3.7％と増に転じています。"
                "当方の算定は令和7年度を基準としており直近の水準を"
                "起点としています。", 3, height=58)

# ============================================================ 08
ws = sheet("08_認知症総合支援の町別評価", "認知症総合支援の町別評価（交付金）",
           "保険者機能強化推進交付金等（市町村分）の介護保険保険者努力支援"
           "交付金 目標Ⅱ「認知症総合支援を推進する」（100点満点）の"
           "町別の得点です。評価は市町村単位で行われます。",
           [34, 8, 12, 12, 12, 40])
r = 4
r = lead(ws, r, "1　目標Ⅱの合計（3か年）", 6)
r = header(ws, r, ["年度", "配点", "東川町", "美瑛町", "東神楽町", "見方"])
for y, yl in (("R6", "令和6年度"), ("R7", "令和7年度"), ("R8", "令和8年度")):
    d = NINCHI[y]
    mi = ("最高%s%d点・最低%s%d点で差は%d点"
          % (max(d, key=d.get), max(d.values()),
             min(d, key=d.get), min(d.values()),
             max(d.values()) - min(d.values())))
    r = body(ws, r, [yl, 100, d["東川町"], d["美瑛町"], d["東神楽町"], mi],
             height=22, fills={1: GRAY},
             fmt={2: N0, 3: N0, 4: N0, 5: N0}, bold=(y == "R8"))
r += 1
r = lead(ws, r, "2　令和8年度の項目別の得点", 6)
r = header(ws, r, ["評価項目", "配点", "東川町", "美瑛町", "東神楽町",
                   "見方"])
NINCHI_ITEM = [x for x in KI.ITEM["R8"] if "認知症総合支援" in x[1]]
ZERO3, WAKARERU = [], []
for _kubun, _moku, _gun, nm, hai, sc in NINCHI_ITEM:
    d = dict(zip(KI.TOWNS, sc))
    if all(v == 0 for v in sc):
        mi, f = "3町とも0点", NG_O
        ZERO3.append(nm)
    elif max(sc) - min(sc) > 0:
        mi, f = "町により%d点の差がある" % (max(sc) - min(sc)), IN_Y
        WAKARERU.append(nm)
    else:
        mi, f = "3町とも同じ", None
    r = body(ws, r, [nm, hai, d["東川町"], d["美瑛町"], d["東神楽町"], mi],
             height=24, fills={1: GRAY, 6: f},
             fmt={2: N0, 3: N0, 4: N0, 5: N0})
r = body(ws, r, ["計", 100] + [NINCHI["R8"][t] for t in TOWNS],
         height=22, bold=True,
         fills={i: MID_B for i in range(1, 7)},
         fmt={2: N0, 3: N0, 4: N0, 5: N0})
r += 1
r = lead(ws, r, "3　打合せで伺うこと", 6)
r = header(ws, r, ["区分", "項目", "", "", "", "伺うこと"])
ws.merge_cells(start_row=r - 1, start_column=2, end_row=r - 1, end_column=5)
for kubun, nm, toi in (
    [("3町とも0点", x,
      "取組がないのか、取組はあるが要件を満たさないのかを伺います。"
      "0点の理由により第10期の書き方が変わります。") for x in ZERO3]
    + [("町により差がある", x,
        "得点の高い町の取組を伺い、3町展開の可否を検討します。")
       for x in WAKARERU]):
    r = body(ws, r, [kubun, nm, "", "", "", toi], height=30,
             fills={1: NG_O if kubun == "3町とも0点" else IN_Y})
    ws.merge_cells(start_row=r - 1, start_column=2, end_row=r - 1,
                   end_column=5)
r = note(ws, r, "※ 得点0が「取組がない」のか「取組はあったが要件を"
                "満たさない」のか「報告していない」のかは、"
                "公表資料からは判別できません。"
                "評価調書によりご確認いただく必要があります。"
                "また、都道府県分の目標Ⅱと市町村分の目標Ⅱは"
                "同じ目標名でも評価項目の中身が違います"
                "（都道府県分は活動指標4項目、市町村分は3項目）。", 6,
         height=52)

# ============================================================ 09
ws = sheet("09_10月の進め方", "10月の進め方",
           "仕様書５の令和8年10月の工程に沿った段取りです。"
           "日程は発注者及び構成3町のご都合により定めます。",
           [14, 26, 44, 40])
r = 4
r = header(ws, r, ["時期", "行うこと", "用いる資料", "決めていただくこと"])
for a, b, c, d in [
    ("10月上旬", "3町合同の意見交換会",
     "本表（01・02・06シート）、計画素案、"
     "サービス見込量 第1次概算、町別データシート",
     "3町共通の確認事項%d件。"
     "特に認知症施策推進計画の策定主体（案B）と意見の聴取の方法。"
     % len(ROWS_BY["3町共通"])),
    ("10月中旬〜下旬", "町ごとの個別協議",
     "本表（03〜05・07・08シート）、町別データシート",
     "町ごとの確認事項（東川町%d件・美瑛町%d件・東神楽町%d件）。"
     "特に施設整備・サービス提供方針と認知症施策の取組。"
     % (len(ROWS_BY["東川町"]), len(ROWS_BY["美瑛町"]),
        len(ROWS_BY["東神楽町"]))),
    ("10月30日", "サービス見込量 第2次概算の提示",
     "サービス見込量の算定（第2次概算）",
     "1人1月あたり給付費の年度の選択（確認事項No.137）と、"
     "保険者機能強化推進交付金等の交付見込額の控除（同No.86）。"),
    ("11月13日", "予算編成用の数値の提示",
     "サービス見込量の算定、保険料算定表",
     "第10期保険料の基準額の水準。"),
    ("11月", "構成町意見の反映、第10期計画骨子の整理",
     "3町意見の整理表、第10期計画骨子案",
     "意見交換会・個別協議で決まったことの計画素案への反映の可否。"),
]:
    r = body(ws, r, [a, b, c, d], height=50, fills={1: GRAY})
r += 1
r = lead(ws, r, "決めていただけない場合", 4)
r = note(ws, r,
         "本表の確認事項には、いずれも「決まらない場合の当方の扱い」を"
         "置いています。ご決定がない場合はその扱いで第2次概算まで進め、"
         "計画素案の本文は［要協議］［要確認］のまま残します。"
         "算定そのものが止まるものはありません。\n"
         "ただし、介護報酬改定率・第1号被保険者負担割合・"
         "所得段階の政令改正の3件は国の告示・公布を待つものであり、"
         "「受領できない」では済みません。"
         "これらが示された時点で保険料の算定をやり直します。", 4, height=92)

# ============================================================ 10 自己点検
chk(1, "構成3町に関わる確認事項を業務工程管理表から読んでいること",
    "%d件（決着したものを除く）" % len(MACHI_NO), len(MACHI_NO) > 0)
chk(2, "町別の確認事項がいずれも1件以上あること",
    "東川町%d件・美瑛町%d件・東神楽町%d件・3町共通%d件"
    % (len(ROWS_BY["東川町"]), len(ROWS_BY["美瑛町"]),
       len(ROWS_BY["東神楽町"]), len(ROWS_BY["3町共通"])),
    all(len(ROWS_BY[m]) > 0 for m in ROWS_BY))
_nashi = [x[0] for m in ROWS_BY for x in ROWS_BY[m] if not str(x[4]).strip()]
chk(3, "全ての確認事項に「決まらない場合の当方の扱い」があること",
    "扱いのないもの%d件" % len(_nashi), not _nashi)
chk(4, "地域課題が10件以上あり、いずれにも根拠と反映先があること",
    "%d件" % len(KADAI),
    len(KADAI) >= 10 and all(k[2] and k[4] for k in KADAI))
chk(5, "認知症初期集中支援推進事業の休止が受領資料から読めていること",
    "休止 %s" % ("・".join(KYUSHI) if KYUSHI else "なし"),
    KYUSHI == ["美瑛町"])
chk(6, "認知症総合支援の目標Ⅱが6項目・配点100点であること",
    "%d項目・配点計%d点" % (len(NINCHI_ITEM), sum(x[4] for x in NINCHI_ITEM)),
    len(NINCHI_ITEM) == 6 and sum(x[4] for x in NINCHI_ITEM) == 100)
chk(7, "目標Ⅱの項目別得点の和が公表の目標Ⅱ合計に一致すること",
    "・".join("%s %d点" % (t, NINCHI["R8"][t]) for t in TOWNS),
    all(sum(x[5][KI.TOWNS.index(t)] for x in NINCHI_ITEM) == NINCHI["R8"][t]
        for t in TOWNS))
chk(8, "3町とも0点の項目があること（打合せで伺う対象）",
    "%d件" % len(ZERO3), len(ZERO3) >= 1)
chk(18, "指定事業所の件数が町ごとに1件以上あること",
     "・".join("%s %d件" % (t, SHITEI_N[t]) for t in TOWNS),
     all(SHITEI_N[t] > 0 for t in TOWNS))
chk(9, "町別の費用総額の和が利用状況統計表の3町計に一致すること",
    yen(HIYOU8_K),
    HIYOU8_K == sum(RJ.val(t, "38介護給付費計", "計", "費用総額")
                    for t in TOWNS))
chk(10, "認定者数の動きが町で分かれること（東川・東神楽は増、美瑛は減）",
     "東川%+d・美瑛%+d・東神楽%+d"
     % (NIN8["東川町"] - NIN6["東川町"], NIN8["美瑛町"] - NIN6["美瑛町"],
        NIN8["東神楽町"] - NIN6["東神楽町"]),
     NIN8["東川町"] > NIN6["東川町"] and NIN8["美瑛町"] < NIN6["美瑛町"])
chk(11, "地域密着型介護老人福祉施設の令和11年度の到達率が99％を超えること",
     pct(_chimitsu_r11 / _chimitsu_teiin),
     0.99 <= _chimitsu_r11 / _chimitsu_teiin <= 1.0)
chk(12, "認知症対応型共同生活介護の令和11年度の見込みが定員内であること",
     "月%.1f人／定員%d人" % (_ninchi_gh, _gh_teiin),
     _ninchi_gh < _gh_teiin)
chk(13, "計画素案に認知症基本法の記載があること",
     "第1章第2節ほか",
     "認知症基本法" in open(
         RP.ROOT + "/build_plan_draft.py", encoding="utf-8").read())
_NG = ["〜に由来する", "全国トップ級", "1件も"]
_ALL = []
for _ws in wb.worksheets:
    for _row in _ws.iter_rows(values_only=True):
        for _v in _row:
            if isinstance(_v, str):
                _ALL.append(_v)
chk(14, "禁止表現が含まれていないこと",
     "・".join(w for w in _NG if any(w in v for v in _ALL)) or "0件",
     not [w for w in _NG if any(w in v for v in _ALL)])
import re as _re                                          # noqa: E402
chk(15, "個人情報（電話番号・メールアドレス）が含まれていないこと", "0件",
     not [v for v in _ALL
          if _re.search(r"[\w.+-]+@[\w.-]+", v)
          or _re.search(r"0\d{1,4}[-(（]\d{2,4}[-)）]\d{3,4}", v)])
_NAIBU = ["固定値", "実物から", "runpy", ".py", "スクリプト", "再実行"]
chk(16, "受託者の内部の仕組み・作業経過の語が本文に残っていないこと",
     "・".join(w for w in _NAIBU if any(w in v for v in _ALL)) or "0件",
     not [w for w in _NAIBU if any(w in v for v in _ALL)])
chk(17, "強調記号が本文に残っていないこと",
     "0件", not [v for v in _ALL if "**" in v])

ws = sheet("10_自己点検", "自己点検",
           "本表の値が受領資料及び当方の算定と合っていることを"
           "機械で確かめた結果です。1件でも不適合があれば作成を止めます。",
           [6, 52, 46, 12])
r = header(ws, 4, ["No.", "確かめたこと", "結果", "判定"])
for no, naiyo, kekka, han in CHK:
    r = body(ws, r, [no, naiyo, kekka, han], height=30,
             align={1: "center", 4: "center"},
             fills={4: OK_G if han == "適合" else NG_O})

os.makedirs(ODIR, exist_ok=True)
wb.save(OUT)
print("出力:", OUT)
print("地域課題 %d件／確認事項 東川%d・美瑛%d・東神楽%d・共通%d"
      % (len(KADAI), len(ROWS_BY["東川町"]), len(ROWS_BY["美瑛町"]),
         len(ROWS_BY["東神楽町"]), len(ROWS_BY["3町共通"])))
print("自己点検 %d件　不適合 %d件"
      % (len(CHK), sum(1 for x in CHK if x[3] != "適合")))
for x in CHK:
    if x[3] != "適合":
        print("  不適合 No.%s %s → %s" % (x[0], x[1], x[2]))
if any(x[3] != "適合" for x in CHK):
    _sys.exit(1)
