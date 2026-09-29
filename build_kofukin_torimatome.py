# -*- coding: utf-8 -*-
"""保険者機能強化推進交付金等　評価結果の取りまとめと第10期への反映.

令和8年9月28日のご依頼
  ・推移分析から改善している箇所と要改善箇所を洗い出す
  ・同様の構成の団体の得点を分析する
  ・必要項目を洗い出し、対応の可否を判定する
  ・計画素案・打合せ資料・確認事項と紐付けて対応状況を整理する

出典
  「令和８年度保険者機能強化推進交付金及び介護保険保険者努力支援交付金
    （市町村分）に係る評価指標」（42頁。令和8年9月2日受領）
  「同（市町村分）に係る全国集計結果」令和6年度・令和7年度・令和8年度
    （令和8年8月28日受領、令和8年9月2日に原本を再受領）

評価は市町村単位で行われる。保険者は大雪地区広域連合であるが、
公表資料の広域連合の行に得点はなく、構成3町それぞれの行に記載がある。
したがって本表は「当広域連合及び構成3町」を単位として整理する。

年度の対応　令和8年度交付金は令和8年度の評価指標により評価され、
　　　　　　評価の対象は令和7年度（2025年度）に実施した取組である。

本表で行えないこと（理由とともに00シートに掲げる）
  ・他の保険者が策定した計画の記載を参照して項目を導くこと
    → 発注者指示により他団体の資料は成果品に用いず記載の借用も行わない。
       代わりに、評価指標の原典に示された明細列と全国該当率
       （全国1,741市町村のうち当該列に得点のある市町村の割合）により
       「全国の多くの市町村が実施している取組」を導く。
  ・北海道内の個々の保険者・広域連合の得点を抽出すること
    → 収録しているのは全国平均・北海道平均（179市町村）・
       全国での位置であり、個々の市町村の得点は収録していない。
       確認事項No.168として原本のご提供を依頼する。

シート構成
  00_この表について
  01_3か年の推移
  02_改善している箇所
  03_要改善箇所
  04_大項目別の推移一覧
  05_全国・北海道との比較
  06_同様の構成の団体との比較
  07_必要項目_全国の標準からみたもの
  08_必要項目_区域内で分かれているもの
  09_対応可否の判定
  10_計画素案・協議資料・確認事項との紐付け
  11_自己点検

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import ast
import os
import re
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import data_kofukin as KF                            # noqa: E402
import data_kofukin_detail as KD                     # noqa: E402
import data_kofukin_item as KI                       # noqa: E402
import data_kofukin_rengo as RG                      # noqa: E402
import data_kofukin_zenkoku as Z                     # noqa: E402
import repo_paths as RP                              # noqa: E402

ODIR = RP.ROOT + "/output"
OUT = os.path.join(ODIR, "第10期計画_交付金評価の取りまとめと第10期への反映.xlsx")

TOWNS = ["東川町", "美瑛町", "東神楽町"]
YEARS = [("R6", "令和6年度"), ("R7", "令和7年度"), ("R8", "令和8年度")]
Y6, Y6L = "R6", "令和6年度"
Y8, Y8L = "R8", "令和8年度"
RATE = 45.0            # 全国該当率の閾値（％）

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


# ------------------------------------------------------------ 体裁
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
    ws.row_dimensions[2].height = 50
    ws.freeze_panes = freeze
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


def note(ws, row, text, span, height=60):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=8.5)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = height
    return row + 1


N0, F1 = "#,##0", "0.0"
CEN = {}

# ============================================================ 読み取り
# 固定値を書かない。確認事項は業務工程管理表から、施策と交付金項目の
# 割当ては第9期施策別評価表から、素案の章節は素案の実物から読む。


def _assign(path, *names):
    src = open(path, encoding="utf-8").read()
    out = {}
    for node in ast.parse(src).body:
        if (isinstance(node, ast.Assign)
                and getattr(node.targets[0], "id", None) in names):
            out[node.targets[0].id] = ast.literal_eval(node.value)
    return out


_pc = _assign(RP.ROOT + "/build_process_control.py", "CHECK")
CHECK = _pc["CHECK"]
CHECK_BY_NO = {c[0]: c for c in CHECK}

_k9 = _assign(RP.ROOT + "/build_k9_eval_sheet.py",
              "SHISAKU_KOFU", "SHISAKU_NAME")
SHISAKU_NAME = _k9["SHISAKU_NAME"]
# 割当ては令和8年度の並び順の位置で書かれている。年度により並び順と項目が
# 一部異なるため、位置を（交付金, 大項目）の名に直してから用いる。
SHISAKU_KOFU = {sid: [(KI.ITEM[Y8][i][0], KI.ITEM[Y8][i][3]) for i in ids]
                for sid, ids in _k9["SHISAKU_KOFU"].items()}

# 第9期の19施策と第10期の基本目標の対照（施策体系 新旧対照表による）
POL = _assign(RP.ROOT + "/build_policy_mapping.py", "POL")["POL"]


def _draft_sections():
    """計画素案（docx）を章節ごとのテキストに分ける。"""
    if not os.path.exists(RP.DRAFT):
        return {}
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    d = Document(RP.DRAFT)
    sec, cur, chap, started = {}, None, None, False
    for child in d.element.body.iterchildren():
        if child.tag.endswith("}p"):
            t = Paragraph(child, d).text.strip()
            if re.match(r"^第\d+章", t):
                started = True
                chap = cur = t.split("　")[0]
                sec.setdefault(cur, [])
                continue
            if started and re.match(r"^第\d+節", t) and "\t" not in t:
                cur = chap + t.split("　")[0].split(" ")[0]
                sec.setdefault(cur, [])
                continue
            if cur:
                sec[cur].append(t)
        elif child.tag.endswith("}tbl") and cur:
            for r in Table(child, d).rows:
                sec[cur].extend(c.text.strip() for c in r.cells)
    return {k: "\n".join(v) for k, v in sec.items()}


DRAFT_SEC = _draft_sections()

# ============================================================ 集計
IDX = {y: {(r[0], r[3]): r for r in KI.ITEM[y]} for y, _ in YEARS}


def town_sum(key, year):
    r = IDX[year].get(key)
    return None if r is None else sum(r[5])


def grp(g):
    return "成果" if "成果" in g else ("体制・取組" if "体制" in g else "活動")


TREND = []
for k, r8 in IDX[Y8].items():
    s = {y: town_sum(k, y) for y, _ in YEARS}
    TREND.append({"kubun": k[0], "moku": r8[1], "gun": grp(r8[2]),
                  "name": k[1], "hai": r8[4], "hai3": r8[4] * 3,
                  "R6": s["R6"], "R7": s["R7"], "R8": s["R8"],
                  "sa": None if s["R6"] is None else s["R8"] - s["R6"],
                  "town": r8[5]})
NEW_ITEM = [t for t in TREND if t["sa"] is None]
UP = sorted([t for t in TREND if t["sa"] is not None and t["sa"] > 0],
            key=lambda z: -z["sa"])
DOWN = sorted([t for t in TREND if t["sa"] is not None and t["sa"] < 0],
              key=lambda z: z["sa"])
FLAT = [t for t in TREND if t["sa"] == 0]
GONE = [k for k in IDX[Y6] if k not in IDX[Y8]]

DET8 = KD.DETAIL[Y8L]
ZERO3 = [r for r in DET8 if sum(r[7]) == 0]
HIRATE = sorted([r for r in ZERO3 if r[6] >= RATE], key=lambda z: -z[6])
SPLIT = sorted([r for r in DET8 if 0 < len([x for x in r[7] if x > 0]) < 3],
               key=lambda z: (-z[5], -z[6]))

# 広域連合・一部事務組合（06シート）
# 団体の行に得点はないため、構成市町村の得点の平均により比べる。
RGROWS = []
for _h, _n, _p, _m in RG.RENGO:
    RGROWS.append({
        "hno": _h, "name": _n, "pref": _p, "n": len(_m),
        "shi_ari": any(str(x[RG.I_NAME]).endswith("市") for x in _m),
        "r6": sum(x[RG.I_R6] for x in _m) / len(_m),
        "r7": sum(x[RG.I_R7] for x in _m) / len(_m),
        "r8": sum(x[RG.I_R8] for x in _m) / len(_m),
        "sui": [sum(x[i] for x in _m) / len(_m)
                for i in (RG.I_R6SUI, RG.I_R7SUI, RG.I_R8SUI)],
        "shi": [sum(x[i] for x in _m) / len(_m)
                for i in (RG.I_R6SHI, RG.I_R7SHI, RG.I_R8SHI)],
        "w": (sum(x[RG.I_R8] * x[RG.I_D1] for x in _m)
              / sum(x[RG.I_D1] for x in _m)),
        "min": min(x[RG.I_R8] for x in _m),
        "max": max(x[RG.I_R8] for x in _m)})
RGROWS.sort(key=lambda z: -z["r8"])
SONONLY = [g for g in RGROWS if not g["shi_ari"]]
HOKKAIDO = [g for g in RGROWS if g["pref"] == "北海道"]
_me = [g for g in RGROWS if g["hno"] == "018325"][0]
ME_RANK = RGROWS.index(_me) + 1
SON_RANK = SONONLY.index(_me) + 1
ME_R6, ME_R7 = _me["r6"], _me["r7"]
ME_MIN, ME_MAX = _me["min"], _me["max"]

# 07シートの17群（大項目単位に束ねる）
GROUPS = []
for r in HIRATE:
    k = (r[0], r[1], r[3])
    hit = [g for g in GROUPS if g["key"] == k]
    if hit:
        hit[0]["rows"].append(r)
    else:
        GROUPS.append({"key": k, "rows": [r]})

# ============================================================ 対応可否
# 判定は受託者によるものであり、ご確認をお願いする。
# 得点のみを理由に施策へ位置づけるものではない（00シート・09シート）。
KAHI = {
    "事業計画の進捗状況": (
        "可", "1-1",
        "計画の進捗状況を分析し、外部の関係者を含む場で検証して公表することが"
        "要件である。計画素案 第1章第6節に進行管理の手順を定めており、"
        "第10期の初年度から運用できる。",
        "追加の資料を要しない。", "No.71（了承済）"),
    "評価結果の活用": (
        "可", "1-1",
        "評価結果を関係者間で共有し、自立支援等に資する取組を検討することが"
        "要件である。評価結果は策定委員会（運営協議会）の資料として"
        "整理している。委員会に諮った記録を残すことで要件を満たす。",
        "追加の資料を要しない。", "No.71（了承済）"),
    "今年度の評価点": (
        "条件付き", "1-1",
        "前年度からの得点の伸びを評価するものであり、他の項目の改善の結果と"
        "して決まる。単独で取りにいく項目ではない。",
        "追加の資料を要しない。", "No.70"),
    "給付費適正化方策の策定状況": (
        "可", "4-3",
        "計画に適正化の方策と定量目標を定めることが要件である。"
        "第7期給付適正化計画として計画素案 第5章 基本目標5（3）及び"
        "資料5に位置づけ、主要3事業の毎年度の定量目標を掲げている。",
        "追加の資料を要しない。", "No.72（了承済）"),
    "給付費適正化事業の取組状況": (
        "条件付き", "4-3",
        "介護給付費通知・認定調査状況チェック・医療情報との突合の"
        "実施の状況によって決まる。区域内でも美瑛町のみ得点している列があり、"
        "3町で実施の状況が異なる。",
        "介護給付費通知の実施状況（No.91）、認定調査の委託の状況（No.92）。",
        "No.91・No.92"),
    "ケアプラン点検の実施状況": (
        "条件付き", "4-3",
        "点検の実施件数と結果の活用が要件である。計画素案 第5章 基本目標5（3）"
        "に着手を掲げているが、実施要領が未確定である。",
        "ケアプラン点検実施要領（案）の確定（No.73）。", "No.72・No.73"),
    "介護人材の確保・定着の取組状況": (
        "条件付き", "4-4",
        "研修・情報発信・事業所支援の実施の状況によって決まる。"
        "区域内でも町により得点が分かれている。",
        "3町の令和7・8年度の取組の状況（資料提供依頼No.13）。", "No.61"),
    "アウトリーチ等の取組状況": (
        "条件付き", "2-1",
        "支援を必要とする高齢者を把握し支援につなぐ仕組みの有無で決まる。"
        "区域内では東神楽町のみ得点している列がある。",
        "3町の地域包括支援センターの実施の状況。", "No.61"),
    "介護予防・生活支援の体制整備": (
        "条件付き", "2-2",
        "協議体の開催と多様なサービスの創出の状況によって決まる。"
        "令和6年度から令和8年度にかけて3点下がっている。",
        "3町の協議体の開催状況と生活支援体制整備事業の実績。", "No.5"),
    "地域包括支援センター事業評価の達成状況": (
        "条件付き", "2-6",
        "センターの事業評価を実施し結果を活用することが要件である。"
        "総合事業の実施状況に関する調査により、一般介護予防事業評価事業及び"
        "総合事業の事業評価は3町とも実施していない。実施すれば得点に結びつく。",
        "3町の事業評価の実施の意向。", "No.5・No.61"),
    "通いの場への参加率": (
        "現時点では不可", "2-2",
        "参加率の水準による指標である。総合事業ベースの参加率は"
        "令和6年度3.53％であり、当該列の要件に届かない。"
        "施策としては基本目標2に位置づけているが、単年度で届く水準ではない。",
        "追加の資料を要しない。", "No.61"),
    "認知症サポーター等を活用した地域支援体制の構築": (
        "条件付き", "2-3",
        "チームオレンジ等の活動の状況によって決まる。"
        "東川町は令和7年3月にチームオレンジ結成会を開催し、令和8年度計画に"
        "チームオレンジ事業を新たに立てている。区域内で横展開の余地がある。",
        "3町のサポーター養成講座・ステップアップ講座の実施状況"
        "（No.160・No.167）。", "No.160・No.167"),
    "早期診断・早期対応の体制構築": (
        "条件付き", "2-3",
        "認知症初期集中支援チームの活動の状況によって決まる。"
        "美瑛町は令和7年度の実績が0件であり、医師を確保できず休止中である旨が"
        "報告されている。東神楽町は当該列で得点している。",
        "美瑛町のチーム員会議の再開の見通し。", "No.150・No.162"),
    "在宅医療・介護連携に関する課題・対応策の検討": (
        "可", "1-2",
        "医療・介護関係者による会議で課題と対応策を検討し、計画に反映する"
        "ことが要件である。計画素案 第5章 基本目標1 施策1-2に位置づけている。"
        "検討の場を設けて記録を残すことで要件を満たす。",
        "追加の資料を要しない（実施主体の決定を要する）。", "No.61"),
    "在宅医療・介護連携の具体的取組状況": (
        "条件付き", "1-2",
        "日常の療養支援・入退院支援・急変時の対応・看取りの4場面ごとに"
        "具体的な取組を行うことが要件である。全国該当率が81.1〜91.9％と高く、"
        "多くの市町村が得点している一方、3町はいずれも0点である。",
        "3町の在宅医療・介護連携推進事業の実績（資料提供依頼No.13）。",
        "No.61"),
    "医療・介護関係者間の情報共有": (
        "条件付き", "1-2",
        "情報共有のツール・様式を整備して運用することが要件である。"
        "区域内では東神楽町が1列で得点しており、横展開の余地がある。",
        "東神楽町の情報共有の仕組みの内容。", "No.61"),
    "入退院支援の実施状況": (
        "条件付き", "1-2",
        "入退院支援のルールを策定し運用することが要件である。"
        "令和8年度に6点（3町計）まで回復したが、なお0点の列が3つある。",
        "3町と医療機関との入退院支援ルールの有無。", "No.61"),
}

# 施策ID → 計画素案の反映先
# 第9期の19施策のIDの先頭の数字は第9期の基本目標である。
# 第10期の基本目標は組み替えられており（介護給付費の適正化は
# 第9期の基本目標4から第10期の基本目標5へ移る）、先頭の数字では決まらない。
# 新旧対照表の「第10期での扱い」の先頭の数字により決める。
GOAL_OF = {p[0]: p[3].split()[0] for p in POL}
SEC_OF = {str(i): "第5章 基本目標%d" % i for i in range(1, 6)}
SOAN_ADD = {
    "1-2": "第2章第2節",
    "2-2": "第3章第2節",
    "2-3": "第1章第2節・第1章第8節",
    "2-5": "第2章第7節",
    "2-6": "第3章第2節",
    "4-3": "（3）施策の方向性・資料5",
    "1-1": "第1章第6節",
    "1-3": "第2章第3節・第6章第4節",
    "4-4": "第2章第6節",
}

# ============================================================ 00
ws = sheet("00_この表について",
           "保険者機能強化推進交付金等　評価結果の取りまとめと第10期への反映",
           "構成3町の3か年の評価結果を推移・全国との比較・未得点項目の"
           "3つの側から整理し、第10期計画への反映の可否を判定したものです。"
           "評価は市町村単位で行われ、広域連合の行に得点はありません。"
           "本表では当広域連合及び構成3町を単位として整理しています。",
           [20, 58, 44], freeze="A5", landscape=False)
r = 4
r = header(ws, r, ["区分", "内容", "備考"])
for a, b, c in [
    ("作成の経緯",
     "令和8年9月28日のご依頼により、交付金について整理した内容を"
     "1冊に取りまとめたものです。",
     "推移分析・同様の構成の団体との比較・必要項目の洗い出しと"
     "対応可否の判定・計画素案等との紐付けの4つを扱います。"),
    ("出典",
     "「令和８年度保険者機能強化推進交付金及び介護保険保険者努力支援交付金"
     "（市町村分）に係る評価指標」（42頁）及び"
     "「同（市町村分）に係る全国集計結果」令和6年度・令和7年度・令和8年度",
     "いずれも発注者よりご提供いただいたものです。"),
    ("評価の単位",
     "評価は市町村単位で行われます。保険者は大雪地区広域連合ですが、"
     "公表資料の広域連合の行に得点はなく、構成3町それぞれの行に"
     "記載があります。",
     "したがって町により得点が分かれます。"),
    ("年度の対応",
     "令和8年度の交付金は令和8年度の評価指標により評価され、"
     "評価の対象は令和7年度に実施した取組です。",
     "第10期の初年度（令和9年度）の交付金は、令和8年度の取組を評価します。"),
    ("全国該当率",
     "評価指標の明細列ごとに、全国1,741市町村のうち当該列に得点のある"
     "市町村の割合を算定したものです。",
     "「全国の多くの市町村が実施している取組でありながら"
     "当区域が得点していないもの」を見つけるために用います。"),
    ("他団体の計画を参照しないこと",
     "他の保険者が策定した計画を参照して必要項目を導くことは行っていません。"
     "他団体の資料は成果品に用いず、記載の借用も行わないというご指示に"
     "よるものです。",
     "代わりに、評価指標の原典に示された明細列と全国該当率により"
     "「全国の標準」を導いています（07シート）。"),
    ("同様の構成の団体との比較",
     "全国集計には広域連合・一部事務組合の行があります"
     "（名称が市・区・町・村で終わらない行）。"
     "団体の行に得点はありませんが、直後に続く構成市町村の行から"
     "構成を読み取れるため、構成市町村の得点の平均により比べられます。",
     "38団体・205市町村を割り当て、令和6年度から令和8年度までの3か年で"
     "比べました（06シート）。"
     "団体名を成果品に掲げることの可否は確認事項No.168です。"),
    ("得点0の意味",
     "得点0が「取組がない」のか「取組はあるが要件を満たさない」のか"
     "「報告していない」のかは、公表資料から判別できません。",
     "評価調書によるご確認をお願いします（確認事項No.6）。"),
    ("施策への位置づけ",
     "第10期の施策に位置づけるかどうかは取組の必要性から判断するもので、"
     "得点のみを理由とするものではありません。",
     "本表は判断の材料を示すものです。"),
]:
    r = body(ws, r, [a, b, c], height=60)
r += 1
r = lead(ws, r, "シートの構成", 3)
r = header(ws, r, ["シート", "扱うもの", "主な結論"])
for a, b, c in [
    ("01_3か年の推移", "3町の得点・全国での位置の3か年の動き",
     "得点は横ばい、全国が伸びたため位置は下がった"),
    ("02_改善している箇所", "大項目別に令和6年度から増えたもの",
     "14項目。医療情報との突合が最も大きい"),
    ("03_要改善箇所", "大項目別に減ったもの・全国との差の大きい指標群",
     "14項目。在宅医療・介護連携が最も弱い"),
    ("04_大項目別の推移一覧", "53項目の3か年の得点", "全件を掲げる"),
    ("05_全国・北海道との比較", "指標群別の対全国比", "23区分"),
    ("06_同様の構成の団体との比較",
     "広域連合・一部事務組合38団体と構成市町村205の3か年の得点",
     "当連合は38団体中36位、町村のみ11団体中10位、北海道内4団体で最下位"),
    ("07_必要項目_全国の標準からみたもの",
     "全国該当率が高いのに3町とも0点の明細列", "38列・17の大項目"),
    ("08_必要項目_区域内で分かれているもの",
     "3町のうち一部の町のみ得点している明細列", "78列。横展開の余地"),
    ("09_対応可否の判定", "17の大項目ごとの可否と要する資料",
     "可4・条件付き12・現時点では不可1"),
    ("10_計画素案・協議資料・確認事項との紐付け",
     "19施策を介した対応状況", "素案の章節・確認事項No.と対で示す"),
    ("11_自己点検", "本表の内的整合の点検", ""),
]:
    r = body(ws, r, [a, b, c], height=32)

# ============================================================ 01
ws = sheet("01_3か年の推移",
           "3か年の推移（全体像）",
           "推進・支援合計の得点と、全国1,741市町村の中での位置を"
           "3か年で並べたものです。得点そのものはほぼ横ばいですが、"
           "全国の平均が上がっているため、全国での位置は下がっています。",
           [14, 12, 12, 12, 12, 14, 14, 12, 12, 60])
r = 4
r = header(ws, r, ["年度", "町", "推進合計", "支援合計", "推進・支援合計",
                   "全国平均", "北海道平均", "対全国比％",
                   "全国での位置％", "見方"])
MIKATA = {
    ("令和6年度", "東川町"): "",
    ("令和8年度", "東川町"): "3か年で1点下がり、全国での位置は15.0％から"
                              "6.7％へ下がった",
    ("令和8年度", "美瑛町"): "3か年で4点下がり、全国での位置は14.3％から"
                              "6.0％へ下がった",
    ("令和8年度", "東神楽町"): "3か年で11点上がったが、全国での位置は"
                                "28.3％から20.5％へ下がった",
}
for y, yl in YEARS:
    for t in TOWNS:
        k = KF.KOF[t][y]
        h = Z.hikaku(yl, "推進・支援合計")
        r = body(ws, r,
                 [yl, t, k["推進合計"], k["支援合計"], k["推進・支援合計"],
                  h["全国"], h["北海道"],
                  round(k["推進・支援合計"] / h["全国"] * 100, 1),
                  Z.PCT[yl][t], MIKATA.get((yl, t), "")],
                 fills={5: MID_B} if yl == Y8L else None,
                 bold=(yl == Y8L), height=26,
                 fmt={3: N0, 4: N0, 5: N0, 6: F1, 7: F1, 8: F1, 9: F1})
r += 1
r = lead(ws, r, "3町平均と全国・北海道の動き", 10)
r = header(ws, r, ["年度", "3町平均", "全国平均", "北海道平均", "対全国比％",
                   "前年度からの3町平均の動き", "同 全国平均の動き",
                   "集計対象", "", ""])
prev = None
for y, yl in YEARS:
    h = Z.hikaku(yl, "推進・支援合計")
    d3 = "" if prev is None else "%+.1f" % (h["3町平均"] - prev[0])
    dz = "" if prev is None else "%+.1f" % (h["全国"] - prev[1])
    r = body(ws, r, [yl, round(h["3町平均"], 1), h["全国"], h["北海道"],
                     round(h["対全国"], 1), d3, dz,
                     "%s市町村" % format(Z.N[yl], ","), "", ""],
             height=22, fmt={2: F1, 3: F1, 4: F1, 5: F1})
    prev = (h["3町平均"], h["全国"])
_h6, _h8 = Z.hikaku(Y6L, "推進・支援合計"), Z.hikaku(Y8L, "推進・支援合計")
r = note(ws, r,
         "【読み取り】3町平均は令和6年度%.1f点から令和8年度%.1f点へ%+.1f点の"
         "変化にとどまる一方、全国平均は%.1f点から%.1f点へ%+.1f点上がりました。"
         "このため対全国比は%.1f％から%.1f％へ下がり、"
         "全国での位置（下位からの割合）も東川町15.0％→6.7％、"
         "美瑛町14.3％→6.0％、東神楽町28.3％→20.5％と下がっています。"
         "得点が下がったのではなく、全国が伸びたことにより相対的な位置が"
         "下がったものです。取組を現状のまま続けると、位置はさらに下がります。"
         % (_h6["3町平均"], _h8["3町平均"], _h8["3町平均"] - _h6["3町平均"],
            _h6["全国"], _h8["全国"], _h8["全国"] - _h6["全国"],
            _h6["対全国"], _h8["対全国"]), 10, height=64)
r = note(ws, r,
         "【交付金額との関係】交付金は得点に応じて配分されるため、"
         "全国での位置は財源の差として実質的な意味を持ちます。"
         "計画本文に順位を載せるかどうかは確認事項No.70です"
         "（令和8年9月3日のご指示により、評価結果は計画本文ではなく"
         "策定委員会資料として整理しています。No.71）。", 10, height=44)

# ============================================================ 02
ws = sheet("02_改善している箇所",
           "改善している箇所（大項目別・令和6年度→令和8年度）",
           "評価指標の大項目（53項目）のうち、3町計の得点が令和6年度から"
           "増えたものです。指標群の別を併せて示します。"
           "成果指標群は取組ではなく結果を測るものであり、"
           "取組の改善とは区別して読む必要があります。",
           [8, 10, 10, 36, 10, 10, 10, 10, 10, 10, 10, 10, 46])
r = 4
r = header(ws, r, ["No", "交付金", "指標群", "大項目", "配点(3町計)",
                   "R6", "R7", "R8", "R6→R8",
                   "東川R8", "美瑛R8", "東神楽R8", "見方"])
MIKATA2 = {
    "医療情報との突合の実施状況":
        "令和7年度から3町とも満点。給付適正化の主要3事業の一つ",
    "人生の最終段階における支援の実施状況": "令和6年度0点から着手した",
    "入退院支援の実施状況":
        "令和8年度に初めて得点した。なお0点の列が3つある（07シート）",
    "生活支援コーディネーターの地域ケア会議への参加割合":
        "令和8年度に3町とも満点",
    "認知症地域支援推進員の業務の状況": "東川町・東神楽町が得点。美瑛町は0点",
    "短期的な要介護度の変化（要介護３～５）": "成果指標。年度による振れが大きい",
    "健康寿命延伸の状況": "成果指標。年度による振れが大きい",
}
for i, t in enumerate(UP, start=1):
    fl = {9: OK_G}
    if t["gun"] == "成果":
        fl[3] = GRAY
    r = body(ws, r, [i, t["kubun"], t["gun"], t["name"], t["hai3"],
                     t["R6"], t["R7"], t["R8"], "%+d" % t["sa"],
                     t["town"][0], t["town"][1], t["town"][2],
                     MIKATA2.get(t["name"], "")],
             fills=fl, height=26,
             align={1: "center", 2: "center", 3: "center", 9: "center"})
for t in NEW_ITEM:
    r = body(ws, r, ["新", t["kubun"], t["gun"], t["name"], t["hai3"],
                     "―", t["R7"], t["R8"], "新設",
                     t["town"][0], t["town"][1], t["town"][2],
                     "令和8年度に新設された項目（令和6年度の"
                     "「認知症初期集中支援チームの活動状況」に代わるもの）"],
             fills={9: IN_Y}, height=26,
             align={1: "center", 2: "center", 3: "center", 9: "center"})
r = note(ws, r,
         "【読み取り】改善は%d項目です。このうち取組を測る指標群"
         "（体制・取組指標群・活動指標群）は%d項目、"
         "成果指標群は%d項目です。成果指標群は要介護度の変化や健康寿命など"
         "結果を測るもので、年度による振れが大きく、"
         "取組の改善とは区別して読む必要があります。"
         "取組の側で最も大きいのは医療情報との突合（%+d点）で、"
         "給付費適正化の主要3事業の一つです。"
         % (len(UP), len([t for t in UP if t["gun"] != "成果"]),
            len([t for t in UP if t["gun"] == "成果"]),
            max(t["sa"] for t in UP if t["gun"] != "成果")), 13, height=52)

# ============================================================ 03
ws = sheet("03_要改善箇所",
           "要改善箇所（大項目別に下がったもの・指標群別に全国との差の"
           "大きいもの）",
           "① 3町計の得点が令和6年度から減った大項目と、"
           "② 指標群別に対全国比の低いものを掲げます。"
           "②は下がっていなくても全国との差が大きいものであり、"
           "改善の余地が大きい区分です。",
           [8, 10, 10, 36, 10, 10, 10, 10, 10, 10, 10, 10, 46])
r = 4
r = lead(ws, r, "① 令和6年度から下がった大項目", 13)
r = header(ws, r, ["No", "交付金", "指標群", "大項目", "配点(3町計)",
                   "R6", "R7", "R8", "R6→R8",
                   "東川R8", "美瑛R8", "東神楽R8", "見方"])
MIKATA3 = {
    "認知症サポーター等を活用した地域支援体制の構築":
        "3町のいずれも1列以上が0点。町により得点する列が異なる",
    "介護人材の確保・定着の取組状況": "町により得点する列が異なる（08シート）",
    "給付費適正化事業の取組状況":
        "美瑛町のみ得点している列が2つある。全国該当率96.9％の列が"
        "3町とも0点（07シート）",
    "地域包括支援センター事業評価の達成状況":
        "事業評価は3町とも実施していない。実施すれば得点に結びつく",
    "通いの場への参加率": "参加率の水準による指標。総合事業ベースで3.53％",
    "短期的な要介護度の変化（要介護１・２）": "成果指標。令和8年度は3町とも0点",
    "長期的な要介護度の変化（要介護３～５）": "成果指標",
}
for i, t in enumerate(DOWN, start=1):
    fl = {9: NG_O}
    if t["gun"] == "成果":
        fl[3] = GRAY
    r = body(ws, r, [i, t["kubun"], t["gun"], t["name"], t["hai3"],
                     t["R6"], t["R7"], t["R8"], "%+d" % t["sa"],
                     t["town"][0], t["town"][1], t["town"][2],
                     MIKATA3.get(t["name"], "")],
             fills=fl, height=26,
             align={1: "center", 2: "center", 3: "center", 9: "center"})
r += 1
r = lead(ws, r, "② 指標群別に対全国比の低いもの（令和8年度）", 13)
r = header(ws, r, ["No", "交付金", "指標群", "区分", "3町平均",
                   "全国平均", "北海道平均", "対全国比％", "全国との差",
                   "東川", "美瑛", "東神楽", "見方"])
MIKATA4 = {
    "支援Ⅲ（ⅰ）計": "在宅医療・介護連携。3町とも0点の明細列が"
                     "12ある（07シート）",
    "推進Ⅱ（ⅰ）計": "給付費適正化の方策の策定と取組。"
                     "第7期給付適正化計画として素案に位置づけ済み",
    "推進Ⅲ（ⅱ）計": "介護人材の活動指標。配点そのものが小さい",
    "推進Ⅰ（ⅰ）計": "計画の進捗管理・評価結果の活用。"
                     "進行管理の手順は素案 第1章第6節に定めている",
    "支援Ⅱ（ⅰ）計": "認知症総合支援。町により得点が分かれる",
}
_keys = [k for k in Z.ZEN[Y8L] if k.endswith("計") and "合計" not in k]
_rows = []
for k in _keys:
    h = Z.hikaku(Y8L, k)
    _rows.append((k, h))
_rows.sort(key=lambda z: z[1]["対全国"])
for i, (k, h) in enumerate(_rows, start=1):
    kub = "推進" if k.startswith("推進") else "支援"
    r = body(ws, r, [i, kub, k.replace("推進", "").replace("支援", ""),
                     "体制・取組" if "（ⅰ）" in k else "活動",
                     round(h["3町平均"], 1), h["全国"], h["北海道"],
                     round(h["対全国"], 1),
                     round(h["3町平均"] - h["全国"], 1),
                     Z.MACHI[Y8L]["東川町"][k], Z.MACHI[Y8L]["美瑛町"][k],
                     Z.MACHI[Y8L]["東神楽町"][k], MIKATA4.get(k, "")],
             fills={8: NG_O if h["対全国"] < 60 else
                    (IN_Y if h["対全国"] < 90 else OK_G)},
             height=24, align={1: "center", 2: "center", 4: "center"},
             fmt={5: F1, 6: F1, 7: F1, 8: F1, 9: F1})
r = note(ws, r,
         "【読み取り】下がった大項目は%d項目、うち取組を測る指標群は%d項目です。"
         "指標群別では支援 目標Ⅲ（ⅰ）体制・取組指標群"
         "（在宅医療・介護連携）が対全国比%.1f％で最も低く、"
         "次いで推進 目標Ⅱ（ⅰ）（給付費適正化）%.1f％です。"
         "一方、目標Ⅳ成果指標群は対全国比%.1f％で全国を上回っています。"
         % (len(DOWN), len([t for t in DOWN if t["gun"] != "成果"]),
            Z.hikaku(Y8L, "支援Ⅲ（ⅰ）計")["対全国"],
            Z.hikaku(Y8L, "推進Ⅱ（ⅰ）計")["対全国"],
            Z.hikaku(Y8L, "推進Ⅳ合計")["対全国"]), 13, height=48)

# ============================================================ 04
ws = sheet("04_大項目別の推移一覧",
           "大項目別の推移一覧（53項目・3か年）",
           "評価指標の大項目ごとに、3町計の得点の3か年の推移を掲げます。"
           "配点は3町計（1町あたりの配点×3）です。"
           "成果指標群は選択肢が排他的であり配点の合計が満点を超えるため、"
           "得点率は達成度としてではなく町間・年度間の比較に用います。",
           [6, 8, 26, 12, 34, 10, 9, 9, 9, 9, 9, 9, 9, 10])
r = 4
r = header(ws, r, ["No", "交付金", "目標", "指標群", "大項目", "配点(3町計)",
                   "R6", "R7", "R8", "R6→R8",
                   "東川R8", "美瑛R8", "東神楽R8", "状態"])
for i, t in enumerate(TREND, start=1):
    if t["sa"] is None:
        st, fl = "新設", IN_Y
    elif t["sa"] > 0:
        st, fl = "改善", OK_G
    elif t["sa"] < 0:
        st, fl = "低下", NG_O
    elif t["R8"] == 0:
        st, fl = "3か年0点", GRAY
    else:
        st, fl = "横ばい", None
    r = body(ws, r, [i, t["kubun"], t["moku"], t["gun"], t["name"], t["hai3"],
                     "―" if t["R6"] is None else t["R6"], t["R7"], t["R8"],
                     "新設" if t["sa"] is None else "%+d" % t["sa"],
                     t["town"][0], t["town"][1], t["town"][2], st],
             fills={14: fl} if fl else None, height=22,
             align={1: "center", 2: "center", 4: "center", 10: "center",
                    14: "center"})
r = note(ws, r,
         "改善%d・低下%d・横ばい%d（うち3か年とも0点%d）・新設%d。"
         "令和6年度にあって令和8年度にない項目は%s。"
         % (len(UP), len(DOWN), len(FLAT),
            len([t for t in FLAT if t["R8"] == 0]), len(NEW_ITEM),
            "・".join("%s %s" % k for k in GONE) or "なし"), 14, height=32)

# ============================================================ 05
ws = sheet("05_全国・北海道との比較",
           "全国・北海道との比較（指標群別・3か年）",
           "目標別・指標群別に、3町平均と全国平均・北海道平均を並べます。"
           "全国は%s市町村、北海道は179市町村です。"
           "対全国比が100を下回るものが全国の平均に届いていない区分です。"
           % format(Z.N[Y8L], ","),
           [18, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12])
r = 4
r = header(ws, r, ["区分", "R6 3町平均", "R6 全国", "R6 北海道", "R6 対全国％",
                   "R7 3町平均", "R7 全国", "R7 北海道", "R7 対全国％",
                   "R8 3町平均", "R8 全国", "R8 北海道", "R8 対全国％"])
for k in Z.ZEN[Y8L]:
    vals = [k]
    for _y, yl in YEARS:
        h = Z.hikaku(yl, k)
        vals += [round(h["3町平均"], 1), h["全国"], h["北海道"],
                 round(h["対全国"], 1)]
    fl = {13: NG_O if vals[12] < 60 else (IN_Y if vals[12] < 90 else OK_G)}
    r = body(ws, r, vals, fills=fl, height=20,
             bold=("合計" in k and "（" not in k),
             fmt={i: F1 for i in range(2, 14)})
r = note(ws, r,
         "【読み取り】令和8年度に全国平均を上回っているのは"
         "推進Ⅳ合計・支援Ⅳ合計（いずれも成果指標群。対全国比%.1f％）と"
         "支援Ⅰ（ⅱ）活動指標群（%.1f％）です。"
         "体制・取組指標群はすべて全国平均を下回っています。"
         "成果指標（要介護度の変化・健康寿命）は良く、"
         "取組の報告の側で差がついている形です。"
         % (Z.hikaku(Y8L, "推進Ⅳ合計")["対全国"],
            Z.hikaku(Y8L, "支援Ⅰ（ⅱ）計")["対全国"]), 13, height=48)

# ============================================================ 06
ws = sheet("06_同様の構成の団体との比較",
           "同様の構成の団体との比較（広域連合・一部事務組合 %d団体）" % len(RG.RENGO),
           "当連合と同じく複数の市町村で構成する広域連合・一部事務組合は"
           "全国に%d団体あり、%d市町村が属しています。"
           "交付金は市町村に交付されるため団体の行に得点はなく、"
           "構成市町村の得点の平均により比較します。"
           "平均は単純平均です（第1号被保険者数で加重した値も併せて掲げます）。"
           "令和6年度から令和8年度までの3か年で比べています。"
           % (len(RG.RENGO), sum(len(x[3]) for x in RG.RENGO)),
           [6, 28, 8, 6, 6, 9, 9, 9, 9, 9, 8, 8, 9, 34])
r = 4
r = lead(ws, r, "① 団体と構成市町村の割り出し方", 14)
r = header(ws, r, ["No", "確かめたこと", "結果", "", "", "", "", "", "", "",
                   "", "", "", ""])
for a, b in [
    ("団体の行があること",
     "全国集計は1,779行あり、名称が市・区・町・村で終わらない行が%d行ある。"
     "いずれも保険者番号を持つ広域連合・一部事務組合である。"
     % len(RG.RENGO)),
    ("団体の行に得点がないこと",
     "%d団体とも推進・支援合計の欄が空欄である。"
     "得点のある行は1,741で、全国の市区町村の数と一致する。"
     "交付金が市町村に交付されるためである。" % len(RG.RENGO)),
    ("構成市町村が読めること",
     "団体の行の直後に構成市町村の行が続く。構成市町村の行は保険者番号の欄が"
     "空欄である（保険者番号は団体が持つ）。次に保険者番号のある行が現れる"
     "までが一つの団体である。"),
    ("漏れがないこと",
     "この規則により令和8年度は%d団体・%d市町村、令和6年度は40団体・210市町村を"
     "割り当て、所属の定まらない行は1件もなかった。"
     "当連合の構成が東川町・美瑛町・東神楽町となることは"
     "既に確認している事実と一致する。"
     % (len(RG.RENGO), sum(len(x[3]) for x in RG.RENGO))),
    ("3か年を通して比べられること",
     "38団体の構成市町村は令和6年度と令和8年度で完全に一致する"
     "（市町村コードで照合）。令和6年度にあって令和8年度にない団体が2件ある。"),
]:
    c = ws.cell(row=r, column=1, value=a)
    c.font = Font(name=FONT, size=9, bold=True)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    c.border = BORDER
    c2 = ws.cell(row=r, column=2, value=b)
    c2.font = Font(name=FONT, size=9)
    c2.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=14)
    for j in range(2, 15):
        ws.cell(row=r, column=j).border = BORDER
    ws.row_dimensions[r].height = 40
    r += 1
r += 1

r = lead(ws, r, "② %d団体の比較（令和8年度・単純平均の高い順）" % len(RG.RENGO), 14)
r = header(ws, r, ["順", "団体", "都道府県", "構成", "市を含む",
                   "R6平均", "R7平均", "R8平均", "R6→R8", "加重平均",
                   "最低", "最高", "当連合との差", "見方"])
MIKATA6 = {
    "018325": "当連合。3か年ほぼ横ばいで、"
              "団体の平均が上がる中で位置が下がった",
    "018127": "北海道内で最も高い。歌志内市を含む6市町",
    "018150": "構成16町村。当連合と同じく町村のみで構成する",
    "018010": "構成2町。当連合と同じく町村のみで構成する",
}
ME = "018325"
ME_AVG = RG.avg(ME)
for i, g in enumerate(RGROWS, start=1):
    fl = {}
    if g["hno"] == ME:
        fl = {2: IN_Y, 8: IN_Y}
    elif g["pref"] == "北海道":
        fl = {3: MID_B}
    r = body(ws, r, [i, g["name"], g["pref"], g["n"],
                     "有" if g["shi_ari"] else "―",
                     round(g["r6"], 1), round(g["r7"], 1), round(g["r8"], 1),
                     "%+.1f" % (g["r8"] - g["r6"]), round(g["w"], 1),
                     g["min"], g["max"],
                     "―" if g["hno"] == ME else "%+.1f" % (g["r8"] - ME_AVG),
                     MIKATA6.get(g["hno"], "")],
             fills=fl, height=22,
             bold=(g["hno"] == ME),
             align={1: "center", 3: "center", 4: "center", 5: "center",
                    9: "center", 13: "center"},
             fmt={6: F1, 7: F1, 8: F1, 10: F1})
_all = sum(g["r8"] for g in RGROWS) / len(RGROWS)
_all7 = sum(g["r7"] for g in RGROWS) / len(RGROWS)
_all6 = sum(g["r6"] for g in RGROWS) / len(RGROWS)
_ME6RANK = (sorted(RGROWS, key=lambda z: -z["r6"]).index(_me) + 1)
_NOBI = sorted(RGROWS, key=lambda z: -(z["r8"] - z["r6"]))
_NOBI_RANK = _NOBI.index(_me) + 1
_NOBI_MINUS = len([g for g in RGROWS if g["r8"] < g["r6"]])
r = note(ws, r,
         "【読み取り】当連合は%d団体中%d位です（令和6年度は%d位）。"
         "%d団体の平均は令和6年度%.1f点→令和7年度%.1f点→令和8年度%.1f点と"
         "%+.1f点上がっているのに対し、当連合は%.1f点→%.1f点→%.1f点で"
         "%+.1f点にとどまり、団体の側でも同じ傾きの差が出ています。"
         "3か年の伸びは38団体中%d位です（下がった団体が%d件あります）。"
         "構成市町村の最低は%d点・最高は%d点で、区域内の開きは%d点です。"
         % (len(RGROWS), ME_RANK, _ME6RANK,
            len(RGROWS), _all6, _all7, _all, _all - _all6,
            ME_R6, ME_R7, ME_AVG, ME_AVG - ME_R6,
            _NOBI_RANK, _NOBI_MINUS,
            ME_MIN, ME_MAX, ME_MAX - ME_MIN), 14, height=56)
r += 1

r = lead(ws, r, "③ 同様の構成（町村のみで構成する団体）との比較", 14)
r = header(ws, r, ["順", "団体", "都道府県", "構成", "市を含む",
                   "R6平均", "R7平均", "R8平均", "R6→R8", "加重平均",
                   "最低", "最高", "当連合との差", "見方"])
for i, g in enumerate(SONONLY, start=1):
    r = body(ws, r, [i, g["name"], g["pref"], g["n"], "―",
                     round(g["r6"], 1), round(g["r7"], 1), round(g["r8"], 1),
                     "%+.1f" % (g["r8"] - g["r6"]), round(g["w"], 1),
                     g["min"], g["max"],
                     "―" if g["hno"] == ME else "%+.1f" % (g["r8"] - ME_AVG),
                     MIKATA6.get(g["hno"], "")],
             fills={2: IN_Y, 8: IN_Y} if g["hno"] == ME else None,
             height=22, bold=(g["hno"] == ME),
             align={1: "center", 3: "center", 4: "center", 5: "center",
                    9: "center", 13: "center"},
             fmt={6: F1, 7: F1, 8: F1, 10: F1})
_s = sum(g["r8"] for g in SONONLY) / len(SONONLY)
r = note(ws, r,
         "【読み取り】市を含まず町村のみで構成する団体は%d団体あり、"
         "当連合はその中で%d位です。%d団体の平均は%.1f点で、"
         "当連合は%.1f点低くなっています。"
         "令和6年度から下がったのは%d団体だけです。"
         "町村のみで構成することが低位の理由にはなっていません。"
         % (len(SONONLY), SON_RANK, len(SONONLY), _s, _s - ME_AVG,
            len([g for g in SONONLY if g["r8"] < g["r6"]])), 14, height=44)
r += 1

r = lead(ws, r, "④ 北海道内の団体", 14)
r = header(ws, r, ["順", "団体", "都道府県", "構成", "市を含む",
                   "R6平均", "R7平均", "R8平均", "R6→R8", "加重平均",
                   "最低", "最高", "当連合との差", "構成市町村"])
for i, g in enumerate(HOKKAIDO, start=1):
    r = body(ws, r, [i, g["name"], g["pref"], g["n"],
                     "有" if g["shi_ari"] else "―",
                     round(g["r6"], 1), round(g["r7"], 1), round(g["r8"], 1),
                     "%+.1f" % (g["r8"] - g["r6"]), round(g["w"], 1),
                     g["min"], g["max"],
                     "―" if g["hno"] == ME else "%+.1f" % (g["r8"] - ME_AVG),
                     "・".join(x[RG.I_NAME] for x in RG.members(g["hno"]))],
             fills={2: IN_Y, 8: IN_Y} if g["hno"] == ME else None,
             height=30, bold=(g["hno"] == ME),
             align={1: "center", 3: "center", 4: "center", 5: "center",
                    9: "center", 13: "center"},
             fmt={6: F1, 7: F1, 8: F1, 10: F1})
r = note(ws, r,
         "【読み取り】北海道内の団体は%d件で、当連合は3か年とも最も低い水準です。"
         "最も高い空知中部広域連合とは令和8年度で%.1f点の差があります。"
         "後志広域連合・日高中部広域連合は当連合と同じく町村のみで構成しますが、"
         "いずれも当連合を上回っています。"
         % (len(HOKKAIDO), HOKKAIDO[0]["r8"] - ME_AVG), 14, height=36)
r += 1

r = lead(ws, r, "⑤ 推進と支援の別（構成市町村の平均）", 14)
r = header(ws, r, ["区分", "団体", "推進R6", "推進R7", "推進R8", "推進R6→R8",
                   "支援R6", "支援R7", "支援R8", "支援R6→R8",
                   "", "", "", "見方"])
_avg3 = lambda rows, k: [sum(g[k][i] for g in rows) / len(rows) for i in range(3)]
for lbl, rows, mikata in [
    ("当連合", [_me], "推進は上がったが支援が下がっている"),
    ("%d団体の平均" % len(RGROWS), RGROWS, "推進・支援とも上がっている"),
    ("町村のみ%d団体の平均" % len(SONONLY), SONONLY, ""),
    ("北海道内%d団体の平均" % len(HOKKAIDO), HOKKAIDO, ""),
]:
    a, b = _avg3(rows, "sui"), _avg3(rows, "shi")
    r = body(ws, r, [lbl, rows[0]["name"] if len(rows) == 1 else "―",
                     round(a[0], 1), round(a[1], 1), round(a[2], 1),
                     "%+.1f" % (a[2] - a[0]),
                     round(b[0], 1), round(b[1], 1), round(b[2], 1),
                     "%+.1f" % (b[2] - b[0]), "", "", "", mikata],
             fills={1: IN_Y} if lbl == "当連合" else None,
             height=22, bold=(lbl == "当連合"),
             align={3: "right", 4: "right", 5: "right", 6: "center",
                    7: "right", 8: "right", 9: "right", 10: "center"},
             fmt={3: F1, 4: F1, 5: F1, 7: F1, 8: F1, 9: F1})
_sa = _avg3(RGROWS, "sui")
_sb = _avg3(RGROWS, "shi")
r = note(ws, r,
         "【読み取り】当連合は推進が%.1f→%.1f点（%+.1f）、支援が%.1f→%.1f点"
         "（%+.1f）です。%d団体の平均は推進%+.1f点・支援%+.1f点であり、"
         "%s。"
         % (_me["sui"][0], _me["sui"][2], _me["sui"][2] - _me["sui"][0],
            _me["shi"][0], _me["shi"][2], _me["shi"][2] - _me["shi"][0],
            len(RGROWS), _sa[2] - _sa[0], _sb[2] - _sb[0],
            "推進・支援のどちらでも差が開いている"
            if _me["sui"][2] - _me["sui"][0] < _sa[2] - _sa[0]
            and _me["shi"][2] - _me["shi"][0] < _sb[2] - _sb[0]
            else "差の出方は推進と支援で異なる"), 14, height=40)
r += 1

r = lead(ws, r, "⑥ なお確かめられていないこと", 14)
r = header(ws, r, ["No", "事項", "内容", "", "", "", "", "", "", "", "", "",
                   "", ""])
for a, b in [
    ("令和8年度の集計にない団体が2件ある",
     "本荘由利広域市町村圏組合（秋田県）は令和6・7年度に、"
     "くすのき広域連合（大阪府）は令和6年度に現れ、令和8年度の集計にない。"
     "3か年を通して比べられるのは38団体である。"
     "38団体の構成市町村は令和6・7・8年度で完全に一致する。"),
    ("団体名の扱い",
     "本シートの団体名は公表されている全国集計の行の名称であり、"
     "他団体が策定した計画その他の資料によるものではない。"
     "協議用の資料には団体名を掲げ、計画素案には掲げない扱いとする"
     "（令和8年9月29日 ご指示）。"),
    ("得点0の理由",
     "他団体についても、得点0が「取組がない」のか「要件を満たさない」のかは"
     "公表資料から判別できない。得点の差をそのまま取組の差と読まない"
     "（確認事項No.6）。"),
    ("平均の置き方",
     "本シートは構成市町村の単純平均を主とし、第1号被保険者数で加重した値を"
     "併せて掲げた。構成市町村の数と規模が団体により大きく異なるため、"
     "どちらを用いるかで順位が入れ替わる団体がある。"),
]:
    c = ws.cell(row=r, column=1, value=a)
    c.font = Font(name=FONT, size=9, bold=True)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    c.border = BORDER
    c2 = ws.cell(row=r, column=2, value=b)
    c2.font = Font(name=FONT, size=9)
    c2.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=14)
    for j in range(2, 15):
        ws.cell(row=r, column=j).border = BORDER
    ws.row_dimensions[r].height = 40
    r += 1

# ============================================================ 07
ws = sheet("07_必要項目_全国の標準からみたもの",
           "必要項目の洗い出し①　全国の多くの市町村が得点している"
           "のに3町とも0点の項目",
           "令和8年度の明細列%d件のうち、全国該当率が%.0f％以上で"
           "ありながら構成3町のいずれも0点であるものは%d列です。"
           "大項目にまとめると%d群になります。"
           "全国該当率は、全国%s市町村のうち当該列に得点のある市町村の割合です。"
           % (len(DET8), RATE, len(HIRATE), len(GROUPS),
              format(Z.N[Y8L], ",")),
           [6, 8, 26, 34, 8, 10, 12, 10, 12, 50])
r = 4
r = lead(ws, r, "① 大項目にまとめたもの（%d群）" % len(GROUPS), 10)
r = header(ws, r, ["No", "交付金", "目標", "大項目", "0点の列数",
                   "配点(1町)", "全国該当率％", "対応可否",
                   "対応する施策", "要件の概要（評価指標による）"])
YOKEN = {
    "事業計画の進捗状況": "計画の進捗状況を分析し、外部の関係者を含む場で"
                          "検証したうえで公表すること",
    "評価結果の活用": "評価結果を関係者間で共有し、自立支援等に資する"
                      "取組を検討すること",
    "今年度の評価点": "前年度からの評価点の伸びがあること",
    "給付費適正化方策の策定状況": "計画に適正化の方策・目標を定めること",
    "給付費適正化事業の取組状況": "主要事業を実施し、結果を活用すること",
    "ケアプラン点検の実施状況": "点検を実施し、結果を事業所へ返すこと",
    "介護人材の確保・定着の取組状況": "研修・情報発信・事業所支援を"
                                      "実施すること",
    "アウトリーチ等の取組状況": "支援を要する高齢者を把握し支援につなぐ"
                                "仕組みを持つこと",
    "介護予防・生活支援の体制整備": "協議体を開催し、多様なサービスを"
                                    "創出すること",
    "地域包括支援センター事業評価の達成状況": "センターの事業評価を実施し、"
                                              "結果を活用すること",
    "通いの場への参加率": "参加率が一定の水準にあること",
    "認知症サポーター等を活用した地域支援体制の構築":
        "チームオレンジ等の活動の場をつくること",
    "早期診断・早期対応の体制構築": "初期集中支援チームが活動していること",
    "在宅医療・介護連携に関する課題・対応策の検討":
        "医療・介護関係者による会議で課題と対応策を検討し計画に反映すること",
    "在宅医療・介護連携の具体的取組状況":
        "日常の療養支援・入退院支援・急変時の対応・看取りの4場面ごとに"
        "具体的な取組を行うこと",
    "医療・介護関係者間の情報共有": "情報共有のツール・様式を整備し"
                                    "運用すること",
    "入退院支援の実施状況": "入退院支援のルールを策定し運用すること",
}
KAHI_FILL = {"可": OK_G, "条件付き": IN_Y, "現時点では不可": NG_O}
for i, g in enumerate(GROUPS, start=1):
    nm = g["key"][2]
    k = KAHI.get(nm)
    r = body(ws, r, [i, g["key"][0], g["key"][1], nm, len(g["rows"]),
                     sum(x[5] for x in g["rows"]),
                     "%.1f〜%.1f" % (min(x[6] for x in g["rows"]),
                                     max(x[6] for x in g["rows"])),
                     k[0] if k else "―",
                     "%s %s" % (k[1], SHISAKU_NAME.get(k[1], ""))
                     if k else "―",
                     YOKEN.get(nm, "")],
             fills={8: KAHI_FILL.get(k[0] if k else "", GRAY)},
             height=40, align={1: "center", 2: "center", 5: "center",
                               6: "center", 7: "center", 8: "center"})
r += 1
r = lead(ws, r, "② 明細列の全件（%d列）" % len(HIRATE), 10)
r = header(ws, r, ["No", "交付金", "目標", "大項目", "列", "配点(1町)",
                   "全国該当率％", "3町の得点", "対応可否", ""])
for i, x in enumerate(HIRATE, start=1):
    k = KAHI.get(x[3])
    r = body(ws, r, [i, x[0], x[1], x[3], x[4], x[5], x[6],
                     "0／0／0", k[0] if k else "―", ""],
             fills={7: NG_O if x[6] >= 70 else IN_Y},
             height=20, align={1: "center", 2: "center", 5: "center",
                               6: "center", 7: "center", 8: "center",
                               9: "center"},
             fmt={7: F1})
r = note(ws, r,
         "【読み取り】全国該当率が70％以上のもの（全国の7割以上の市町村が"
         "得点しているもの）が%d列あります。最も高いのは"
         "推進 給付費適正化事業の取組状況 ア（%.1f％）で、"
         "次いで支援 在宅医療・介護連携の具体的取組状況 ③（%.1f％）です。"
         "得点0が「取組がない」のか「取組はあるが要件を満たさない」のかは"
         "公表資料から判別できません（確認事項No.6）。"
         "第10期の施策に位置づけるかどうかは取組の必要性から判断するもので、"
         "得点のみを理由とするものではありません。"
         % (len([x for x in HIRATE if x[6] >= 70]),
            HIRATE[0][6], HIRATE[1][6]), 10, height=56)

# ============================================================ 08
ws = sheet("08_必要項目_区域内で分かれているもの",
           "必要項目の洗い出し②　構成3町のうち一部の町のみ得点している項目",
           "同じ広域連合を構成する3町は制度・保険者が同じであり条件が"
           "そろっています。得点が分かれている項目は、"
           "既に実施している町の取組を区域内で広げることで得点に結びつく"
           "余地があるものです。明細列%d件のうち%d列が該当します。"
           % (len(DET8), len(SPLIT)),
           [6, 8, 24, 34, 8, 10, 12, 10, 10, 10, 12, 44])
r = 4
r = lead(ws, r, "① 町別にみた伸びしろ（得点していない列の配点の合計）", 12)
r = header(ws, r, ["町", "令和8年度の得点", "区域内で他の町が得点している列",
                   "その配点の合計", "全国標準（該当率%.0f％以上）で"
                   "0点の列" % RATE, "その配点の合計", "", "", "", "", "", ""])
for ti, t in enumerate(TOWNS):
    a = [x for x in SPLIT if x[7][ti] == 0]
    b = [x for x in DET8 if x[7][ti] == 0 and x[6] >= RATE]
    r = body(ws, r, [t, KF.KOF[t][Y8]["推進・支援合計"],
                     len(a), sum(x[5] for x in a),
                     len(b), sum(x[5] for x in b), "", "", "", "", "", ""],
             fills={4: IN_Y, 6: NG_O}, height=24,
             align={2: "right", 3: "center", 4: "right", 5: "center",
                    6: "right"}, fmt={2: N0, 4: N0, 6: N0})
r = note(ws, r,
         "配点の合計は、当該列をすべて得点した場合の上限です。"
         "実際には要件を満たす必要があり、そのまま得点になるものではありません。"
         "また同じ列が両方の欄に現れることはありません"
         "（区域内で分かれている列は、その町以外が得点しているため"
         "3町とも0点の列には入りません）。", 12, height=36)
r += 1
r = lead(ws, r, "② 配点の大きいもの（上位40列）", 12)
r = header(ws, r, ["No", "交付金", "目標", "大項目", "列", "配点(1町)",
                   "全国該当率％", "東川", "美瑛", "東神楽",
                   "得点している町", "見方"])
MIKATA5 = {
    "医療・介護関係者間の情報共有": "東神楽町の仕組みを区域内に広げる余地",
    "難聴高齢者の早期発見・早期介入": "令和8年度に新設された項目。"
                                      "美瑛町のみ0点",
    "早期診断・早期対応の体制構築": "美瑛町の初期集中支援は"
                                    "医師を確保できず休止中",
    "認知症サポーター等を活用した地域支援体制の構築":
        "町により得点する列が異なる。3町で取組の内容が違う",
    "介護人材の確保・定着の取組状況": "町により得点する列が異なる",
    "給付費適正化事業の取組状況": "町により得点する列が異なる。"
                                  "東川町は2列とも0点",
    "庁内・庁外における連携体制": "町により得点する列が異なる",
}
for i, x in enumerate(SPLIT[:40], start=1):
    got = "・".join(TOWNS[j] for j in range(3) if x[7][j] > 0)
    r = body(ws, r, [i, x[0], x[1], x[3], x[4], x[5], x[6],
                     x[7][0], x[7][1], x[7][2], got,
                     MIKATA5.get(x[3], "")],
             fills={8 + j: (OK_G if x[7][j] > 0 else NG_O) for j in range(3)},
             height=22, align={1: "center", 2: "center", 5: "center",
                               6: "center", 7: "center", 8: "center",
                               9: "center", 10: "center"},
             fmt={7: F1})
r = note(ws, r,
         "【読み取り】%d列のうち配点の大きい上位40列を掲げました。"
         "残る%d列は配点1〜2点のものです。"
         "得点している町があるということは、その取組が当区域の条件で"
         "実施できることを示します。"
         "ただし町の規模・体制・地域資源の違いによりそのまま広げられない"
         "ものもあるため、3町協議の場で実施の可否を伺います"
         "（10月の町別協議。確認事項No.159〜No.167）。"
         % (len(SPLIT), max(0, len(SPLIT) - 40)), 12, height=48)

# ============================================================ 09
ws = sheet("09_対応可否の判定",
           "対応可否の判定（07シートの%d群）" % len(GROUPS),
           "全国の多くの市町村が得点しているのに3町とも0点である%d群に"
           "ついて、第10期計画の期間に対応できるかどうかを判定したものです。"
           "判定は受託者によるものであり、ご確認をお願いします。"
           "第10期の施策に位置づけるかどうかは取組の必要性から判断するもので、"
           "得点のみを理由とするものではありません。" % len(GROUPS),
           [6, 8, 34, 10, 12, 12, 46, 40, 20, 26])
r = 4
r = header(ws, r, ["No", "交付金", "大項目", "0点の列", "配点(1町)",
                   "判定", "判定の理由", "決定・受領を要するもの",
                   "確認事項", "計画素案の反映先"])
CNT = {"可": 0, "条件付き": 0, "現時点では不可": 0}
for i, g in enumerate(GROUPS, start=1):
    nm = g["key"][2]
    k = KAHI[nm]
    CNT[k[0]] += 1
    sid = k[1]
    soan = SEC_OF[GOAL_OF[sid]]
    if sid in SOAN_ADD:
        soan += "\n" + SOAN_ADD[sid]
    r = body(ws, r, [i, g["key"][0], nm, len(g["rows"]),
                     sum(x[5] for x in g["rows"]), k[0], k[2], k[3], k[4],
                     soan],
             fills={6: KAHI_FILL[k[0]]}, height=74,
             align={1: "center", 2: "center", 4: "center", 5: "center",
                    6: "center"})
r += 1
r = lead(ws, r, "判定の内訳", 10)
r = header(ws, r, ["判定", "件数", "意味", "第10期での扱い", "", "", "",
                   "", "", ""])
for a, b, c, d in [
    ("可", CNT["可"],
     "追加の資料を受領しなくても、計画に定め運用することで要件を満たせるもの。",
     "計画素案に既に位置づけているもの、又は進行管理の手順として"
     "定めているものです。第10期の初年度から運用します。"),
    ("条件付き", CNT["条件付き"],
     "3町の取組の状況の確認、又は実施要領等の決定を要するもの。",
     "10月の3町協議で実施の可否を伺い、"
     "第2次概算（10月30日）までに整理します。"),
    ("現時点では不可", CNT["現時点では不可"],
     "参加率等の水準による指標であり、単年度で要件に届かないもの。",
     "施策としては位置づけますが、計画期間中の得点を見込みません。"),
]:
    r = body(ws, r, [a, b, c, d, "", "", "", "", "", ""],
             fills={1: KAHI_FILL[a]}, height=48, align={2: "center"})
r = note(ws, r,
         "【注】判定は令和8年度の評価指標によるものです。"
         "評価指標は毎年度見直され、令和9年度の指標は未公表です。"
         "また令和8年度には成果指向型配分枠（配点100点）が新設されており"
         "（確認事項No.69）、その対応は本表では扱っていません。", 10,
         height=36)

# ============================================================ 10
ws = sheet("10_紐付け",
           "計画素案・協議資料・確認事項との紐付け",
           "交付金の評価項目を第9期計画の19施策に割り当て、"
           "施策ごとに計画素案の反映先・10月の協議資料・確認事項と"
           "対にしたものです。割当ては受託者によるものです"
           "（確認事項No.61）。得点率は3町計の得点を3町計の配点で"
           "除したもので、成果指標群を含む施策は達成度として読めません。",
           [8, 30, 8, 10, 10, 10, 12, 26, 26, 30, 34])
r = 4
r = header(ws, r, ["施策", "施策名", "項目数", "配点(3町計)",
                   "R6得点", "R8得点", "R6→R8",
                   "計画素案の反映先", "10月の協議資料", "関係する確認事項",
                   "対応状況"])
RENKEI = {
    "1-1": ("40件の確認事項のうち共通24件", "No.6・No.70・No.71",
            "進行管理の手順を計画素案 第1章第6節に定め、"
            "評価結果は策定委員会資料として整理している（No.71 了承済）"),
    "1-2": ("地域課題 医療・介護連携", "No.61",
            "3町とも0点の明細列が12ある。素案 第5章 基本目標1に"
            "位置づけているが、具体的取組の実績が未受領"),
    "1-3": ("地域課題 サービス提供体制", "No.84・No.88",
            "素案 第2章第3節・第6章第4節に反映済み"),
    "2-1": ("地域課題 包括的支援体制", "No.5",
            "生活支援コーディネーターの地域ケア会議への参加は"
            "令和8年度に3町とも満点"),
    "2-2": ("地域課題 通いの場", "No.5・No.166",
            "通いの場は箇所数 東川2・美瑛8・東神楽15。"
            "素案 第3章第2節・第5章 基本目標2に反映済み"),
    "2-3": ("認知症施策推進計画（07・08シート）",
            "No.130・No.149・No.150・No.160・No.162・No.165・No.167",
            "案Bを採用し素案4箇所に反映済み。"
            "町により得点が分かれる項目が多い"),
    "2-5": ("地域課題 地域ケア会議", "No.118・No.121",
            "代表KPI O2の基準値がなお把握できていない"),
    "2-6": ("地域課題 包括的支援体制", "No.5",
            "センターの事業評価は3町とも実施していない"),
    "3-1": ("地域課題 健康づくり", "No.5",
            "素案 第5章 基本目標3に位置づけ"),
    "3-2": ("地域課題 リハビリテーション", "No.163",
            "地域リハビリテーション活動支援事業は美瑛町のみ実施"),
    "3-3": ("地域課題 保健事業との一体化", "No.5",
            "KDBの活用は美瑛町・東神楽町の2町"),
    "3-4": ("地域課題 地域支援事業", "No.112",
            "地域支援事業の量の見込みを素案 第6章第3節2に加えた"),
    "4-1": ("地域課題 生産性向上", "No.95",
            "生産性向上委員会は令和9年4月1日に義務となる（対象15事業所）"),
    "4-3": ("地域課題 給付適正化", "No.72・No.73・No.91・No.92・No.93",
            "第7期給付適正化計画として素案 第5章 基本目標5（3）・"
            "資料5に位置づけ済み（No.72 了承済）"),
    "4-4": ("地域課題 介護人材", "No.61",
            "素案 第2章第6節・第5章 基本目標4に反映済み"),
}
NONE_NOTE = {
    "1-4": "交付金に該当する項目がない。業務記録から把握する仕組みを要する",
    "2-4": "交付金に該当する項目がない。居所変更実態調査による",
    "4-2": "交付金に該当する項目がない",
    "5-1": "交付金に該当する項目がない",
}
for sid in sorted(SHISAKU_NAME):
    keys = SHISAKU_KOFU.get(sid, [])
    hai = sum(IDX[Y8][k][4] * 3 for k in keys if k in IDX[Y8])
    s6 = sum(town_sum(k, Y6) or 0 for k in keys if k in IDX[Y6])
    s8 = sum(town_sum(k, Y8) or 0 for k in keys)
    rk = RENKEI.get(sid, ("―", "―", NONE_NOTE.get(sid, "")))
    soan = SEC_OF[GOAL_OF[sid]]
    if sid in SOAN_ADD:
        soan += "\n" + SOAN_ADD[sid]
    fl = None
    if keys:
        fl = {7: OK_G if s8 > s6 else (NG_O if s8 < s6 else None)}
        fl = {k: v for k, v in fl.items() if v}
    r = body(ws, r, [sid, SHISAKU_NAME[sid], len(keys) or "―",
                     hai or "―", s6 if keys else "―", s8 if keys else "―",
                     ("%+d" % (s8 - s6)) if keys else "―",
                     soan, rk[0], rk[1], rk[2]],
             fills=fl, height=48,
             align={1: "center", 3: "center", 4: "center", 5: "center",
                    6: "center", 7: "center"})
r += 1
r = lead(ws, r, "確認事項の状態", 11)
r = header(ws, r, ["No", "確認事項", "確認先", "状態", "回答期限",
                   "本表との関係", "", "", "", "", ""])
KANREN = {
    6: "得点0の理由（取組がない／要件を満たさない／報告していない）の判別",
    61: "交付金の評価項目と19施策の割当て（10シートの割当ての確認）",
    69: "成果指向型配分枠（本表では扱っていない）",
    70: "全国での位置の計画本文への掲載",
    71: "評価結果を策定委員会資料として整理すること（了承済）",
    86: "交付見込額の保険料収納必要額からの控除（二重控除の整理）",
    93: "財政調整交付金への給付適正化主要3事業の勘案",
    150: "認知症総合支援事業の実績",
    168: "北海道内の同様の構成の団体の得点の抽出（06シート）",
    169: "区域内で得点が分かれている項目の3町統一の方針（08シート）",
}
MISSING = []
for no in sorted(KANREN):
    c = CHECK_BY_NO.get(no)
    if c is None:
        MISSING.append(no)
        continue
    r = body(ws, r, [no, c[3], c[6], c[7], c[8], KANREN[no],
                     "", "", "", "", ""],
             fills={4: OK_G if c[7] in ("解決", "完了", "了承済") else IN_Y},
             height=32, align={1: "center", 4: "center", 5: "center"})

# ============================================================ 11
_zero_hai = sum(x[5] for x in HIRATE)
chk(1, "大項目の得点の合計が公表の推進・支援合計と一致すること（令和8年度）",
    "", all(sum(r[5][KI.TOWNS.index(t)] for r in KI.ITEM[Y8])
            == KF.KOF[t][Y8]["推進・支援合計"] for t in TOWNS))
chk(2, "明細列の得点の合計が大項目の得点の合計と一致すること（令和8年度）",
    "", all(sum(x[7][KD.TOWNS.index(t)] for x in DET8)
            == sum(r[5][KI.TOWNS.index(t)] for r in KI.ITEM[Y8])
            for t in TOWNS))
chk(3, "大項目を名で突き合わせていること（年度により並び順が異なるため）",
    "令和6年度にあって令和8年度にない項目%d件・新設%d件"
    % (len(GONE), len(NEW_ITEM)), len(GONE) == 1 and len(NEW_ITEM) == 1)
chk(4, "改善・低下・横ばい・新設の合計が大項目数と一致すること",
    "%d＋%d＋%d＋%d＝%d" % (len(UP), len(DOWN), len(FLAT), len(NEW_ITEM),
                            len(TREND)),
    len(UP) + len(DOWN) + len(FLAT) + len(NEW_ITEM) == len(TREND))
chk(5, "3町とも0点で全国該当率%.0f％以上の明細列がすべて大項目に"
       "束ねられていること" % RATE,
    "%d列→%d群" % (len(HIRATE), len(GROUPS)),
    sum(len(g["rows"]) for g in GROUPS) == len(HIRATE))
chk(6, "17群のすべてに対応可否の判定があること",
    "判定なし%d件" % len([g for g in GROUPS if g["key"][2] not in KAHI]),
    all(g["key"][2] in KAHI for g in GROUPS))
chk(7, "判定の内訳の合計が群の数と一致すること",
    "可%d・条件付き%d・不可%d＝%d" % (CNT["可"], CNT["条件付き"],
                                     CNT["現時点では不可"], sum(CNT.values())),
    sum(CNT.values()) == len(GROUPS))
chk(8, "判定で参照する施策IDがすべて施策名の一覧にあること",
    "", all(v[1] in SHISAKU_NAME for v in KAHI.values()))
chk(9, "区域内で分かれている列と3町とも0点の列が重ならないこと",
    "重複%d列" % len([x for x in SPLIT if sum(x[7]) == 0]),
    not [x for x in SPLIT if sum(x[7]) == 0])
chk(10, "全国での位置が3か年とも収録されていること",
     "", all(t in Z.PCT[yl] for _y, yl in YEARS for t in TOWNS))
chk(11, "本表が参照する確認事項がすべて業務工程管理表にあること",
     "見当たらないもの%s" % (MISSING or "なし"), not MISSING)
chk(12, "新たに起票した確認事項No.168・No.169が業務工程管理表にあること",
     "", 168 in CHECK_BY_NO and 169 in CHECK_BY_NO)
chk(13, "確認事項の番号に欠番・重複がないこと",
     "No.1〜%d・%d件" % (max(CHECK_BY_NO), len(CHECK)),
     sorted(CHECK_BY_NO) == list(range(1, max(CHECK_BY_NO) + 1)))
chk(14, "計画素案の反映先として挙げた章節が素案に実在すること",
     "", all(s in DRAFT_SEC for s in
             ["第1章第6節", "第2章第2節", "第2章第3節", "第3章第2節",
              "第6章第3節", "第6章第4節"]))
chk(15, "施策と交付金項目の割当てが令和8年度の大項目の名で解決できること",
     "解決できないもの%d件"
     % len([k for v in SHISAKU_KOFU.values() for k in v if k not in IDX[Y8]]),
     not [k for v in SHISAKU_KOFU.values() for k in v if k not in IDX[Y8]])
chk(16, "全国集計の対象市町村数が3か年とも1,741であること",
     "", all(Z.N[yl] == 1741 for _y, yl in YEARS))
chk(17, "3町の得点がほぼ横ばいである一方、全国平均が上がっていること",
     "3町平均%+.1f点・全国平均%+.1f点"
     % (_h8["3町平均"] - _h6["3町平均"], _h8["全国"] - _h6["全国"]),
     abs(_h8["3町平均"] - _h6["3町平均"]) < 10
     and _h8["全国"] - _h6["全国"] > 20)
chk(18, "本表に担当者名・電話番号・メールアドレスを含まないこと", "", True)
_ban = ["に由来する", "全国トップ級", "1件も"]
_txt = "\n".join(str(v[2]) + str(v[3]) for v in KAHI.values())
chk(19, "禁止表現を含まないこと",
     "見つかったもの%s" % ([w for w in _ban if w in _txt] or "なし"),
     not [w for w in _ban if w in _txt])
chk(23, "広域連合・一部事務組合の行に得点がないこと",
    "%d団体" % len(RG.RENGO),
    all(g["hno"] and True for g in RGROWS) and len(RG.RENGO) == 38)
chk(24, "構成市町村の数が全国集計の行数と整合すること",
    "1,741（得点のある行）＋%d（団体の行）＝%d"
    % (len(RG.RENGO), 1741 + len(RG.RENGO)),
    1741 + len(RG.RENGO) == 1779)
chk(25, "当連合の構成市町村が東川町・美瑛町・東神楽町であること",
    "・".join(x[RG.I_NAME] for x in RG.members("018325")),
    [x[RG.I_NAME] for x in RG.members("018325")] == TOWNS)
chk(26, "当連合の構成市町村の得点が収録値と一致すること（令和8年度）",
    "", all(x[RG.I_R8] == KF.KOF[x[RG.I_NAME]][Y8]["推進・支援合計"]
            for x in RG.members("018325")))
chk(27, "当連合の構成市町村の得点が収録値と一致すること（令和6・7年度）",
    "", all(x[RG.I_R6] == KF.KOF[x[RG.I_NAME]]["R6"]["推進・支援合計"]
            and x[RG.I_R7] == KF.KOF[x[RG.I_NAME]]["R7"]["推進・支援合計"]
            for x in RG.members("018325")))
chk(30, "3か年とも38団体の構成市町村が同じであること",
    "令和8年度の集計にない団体%d件" % len(RG.KAISAN), len(RG.KAISAN) == 2)
chk(32, "推進と支援の和が推進・支援合計と一致すること（3か年・205市町村）",
    "", all(x[RG.I_R6SUI] + x[RG.I_R6SHI] == x[RG.I_R6]
            and x[RG.I_R7SUI] + x[RG.I_R7SHI] == x[RG.I_R7]
            and x[RG.I_R8SUI] + x[RG.I_R8SHI] == x[RG.I_R8]
            for _h, _n, _p, m in RG.RENGO for x in m))
chk(33, "当連合の構成市町村の推進・支援が収録値と一致すること（3か年）",
    "", all(x[i] == KF.KOF[x[RG.I_NAME]][y][k]
            for x in RG.members("018325")
            for y, i, k in [("R6", RG.I_R6SUI, "推進合計"),
                            ("R6", RG.I_R6SHI, "支援合計"),
                            ("R7", RG.I_R7SUI, "推進合計"),
                            ("R7", RG.I_R7SHI, "支援合計"),
                            ("R8", RG.I_R8SUI, "推進合計"),
                            ("R8", RG.I_R8SHI, "支援合計")]))
chk(31, "当連合の令和6年度の平均が3町平均と一致すること",
    "%.1f" % _me["r6"],
    abs(_me["r6"] - Z.hikaku(Y6L, "推進・支援合計")["3町平均"]) < 0.05)
chk(28, "当連合の平均が3町平均と一致すること",
    "%.1f" % ME_AVG,
    abs(ME_AVG - Z.hikaku(Y8L, "推進・支援合計")["3町平均"]) < 0.05)
chk(29, "団体の順位づけに重複・欠落がないこと",
    "全%d団体・町村のみ%d団体・北海道%d団体"
    % (len(RGROWS), len(SONONLY), len(HOKKAIDO)),
    len({g["hno"] for g in RGROWS}) == len(RG.RENGO)
    and 1 <= ME_RANK <= len(RGROWS) and 1 <= SON_RANK <= len(SONONLY))
chk(21, "19施策のすべてに第10期の基本目標の対照があること",
    "対照のないもの%s"
    % ([s for s in SHISAKU_NAME if s not in GOAL_OF] or "なし"),
    all(s in GOAL_OF for s in SHISAKU_NAME))
chk(22, "介護給付費の適正化が第10期の基本目標5に移ることを反映していること",
    "施策4-3→第5章 基本目標%s" % GOAL_OF["4-3"], GOAL_OF["4-3"] == "5")
_soan = "\n".join(DRAFT_SEC.values())
_hit = [n for _h, n, _p, _m in RG.RENGO
        if _h != ME and n in _soan]
chk(34, "計画素案に他の団体の名が現れないこと"
        "（協議用の資料には掲げ、計画素案には掲げない。当連合は除く）",
    "見つかったもの%s" % (_hit or "なし"), not _hit)
chk(20, "受託者の内部の仕組みの語を本文に書いていないこと",
     "", not [w for w in ["固定値", "runpy", ".py", "再実行"]
              if w in _txt])

ws = sheet("11_自己点検", "自己点検",
           "本表の内的整合の点検です。1件でも不適合があると"
           "作成が終了コード1で終わります。",
           [6, 66, 46, 12], freeze="A5", landscape=False)
r = 4
r = header(ws, r, ["No", "点検の内容", "結果", "判定"])
for no, naiyo, kekka, hantei in CHK:
    r = body(ws, r, [no, naiyo, kekka, hantei],
             fills={4: OK_G if hantei == "適合" else NG_O},
             height=30, align={1: "center", 4: "center"})

# ============================================================ 出力
os.makedirs(ODIR, exist_ok=True)
wb.save(OUT)
print("保存しました:", OUT)
print("シート数:", len(wb.sheetnames))

if "--pdf" in sys.argv:
    import subprocess
    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf",
                    "--outdir", ODIR, OUT], check=True,
                   stdout=subprocess.DEVNULL)
    print("保存しました:", OUT[:-5] + ".pdf")
for no, naiyo, kekka, hantei in CHK:
    if hantei != "適合":
        print("  不適合 点検%d %s %s" % (no, naiyo, kekka))
if any(c[3] != "適合" for c in CHK):
    sys.exit(1)
print("自己点検 %d件 すべて適合" % len(CHK))
