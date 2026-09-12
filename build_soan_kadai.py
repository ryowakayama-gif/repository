# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画　計画素案の課題整理.

令和8年9月12日のご指示
  「アンケート分析と計画素案の課題整理進めて下さい」

━━ 本表の目的 ━━

確認事項は業務工程管理表 03_確認事項一覧で一元管理しており、
計画素案の未確定箇所は本文の［要協議］［要確認］［要内訳］で保持している。
両者は別の場所にあり、次のことが分からなかった。

  ・素案のどこに未確定が集中しているか
  ・どの確認事項が決まれば素案のどの箇所が埋まるか
  ・9月26日（サービス見込量 第1次概算）までに何を決める必要があるか
  ・受託者の作業だけで進むものが残っているか

本表はこの4点を整理する。

━━ 数え方 ━━

**素案の未確定箇所は計画素案の実物（docx）から数える。**
`build_process_control.py` の確認事項は同スクリプトのソースを
`ast` で読み取る。いずれも固定値を書かないため、
素案又は確認事項を更新すれば本表も追随する。

シート構成
  00_この表について
  01_素案の未確定箇所と記述量
  02_確認事項の棚卸し
  03_第1次概算までに決着を要するもの
  04_令和8年9月の新たな知見と素案への割付け
  05_アンケート分析の未了と素案への効き
  06_受託者の作業だけで進むもの
  07_自己点検
  08_確認事項

出力
  output/第10期計画_計画素案の課題整理.xlsx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import ast
import collections
import io
import os
import re
import sys

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import repo_paths as RP

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

ODIR = RP.OUTPUT
OUT = os.path.join(ODIR, "第10期計画_計画素案の課題整理.xlsx")
KIJUNBI = "令和8年9月12日"
GAISAN = "令和8年9月26日（サービス見込量 第1次概算）"

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


# ============================================================ 素案を数える
PH = re.compile(r"［(要協議|要確認|要内訳)］")
SHO = re.compile(r"^(第\d+章|資料\d+)")
SETSU = re.compile(r"^(第\d+節|基本目標\d)")

_doc = Document(RP.DRAFT)
SOAN = []                     # (章, 節, 位置, 種別, 文言)
RYO = collections.OrderedDict()          # (章, 節) -> [段落, 表, 注]
_sho, _setsu, _ti = "（表紙・目次）", "", 0


def _is_mokuji(t):
    """目次の行はタブでページ番号を連結しているため見出しとして数えない。"""
    return "\t" in t


for _ch in _doc.element.body.iterchildren():
    _tag = _ch.tag.split("}")[-1]
    if _tag == "p":
        _p = Paragraph(_ch, _doc)
        _t = _p.text.strip()
        if not _is_mokuji(_t):
            if SHO.match(_t):
                _sho, _setsu = _t[:26], ""
            elif SETSU.match(_t):
                _setsu = _t[:30]
        _k = (_sho, _setsu)
        RYO.setdefault(_k, [0, 0, 0])
        if _t and not _is_mokuji(_t):
            RYO[_k][0] += 1
            if _t.startswith("資料：") or _t.startswith("注"):
                RYO[_k][2] += 1
        for m in PH.finditer(_p.text):
            SOAN.append((_sho, _setsu, "本文", m.group(1), _t))
    elif _tag == "tbl":
        _ti += 1
        _k = (_sho, _setsu)
        RYO.setdefault(_k, [0, 0, 0])
        RYO[_k][1] += 1
        for _row in Table(_ch, _doc).rows:
            for _c in _row.cells:
                for m in PH.finditer(_c.text):
                    SOAN.append((_sho, _setsu, "表%d" % _ti, m.group(1),
                                 _c.text.strip().replace("\n", "／")))

# 節ごとの記述量（章の直下・目次・資料編を除く）
SETSU_RYO = [(s, se, v[0], v[1], v[2]) for (s, se), v in RYO.items()
             if se and s.startswith("第")]
# 本文（注・出典を除く段落）が5未満かつ表が1以下の節を「記述が薄い」とする
USUI = [x for x in SETSU_RYO if (x[2] - x[4]) < 5 and x[3] <= 1]

PH_KEI = len(SOAN)
PH_SHO = collections.Counter(x[0] for x in SOAN)
PH_SETSU = collections.Counter((x[0], x[1]) for x in SOAN)
PH_SHU = collections.Counter(x[3] for x in SOAN)
PH_HON = [x for x in SOAN if x[2] == "本文"]
PH_HYO = collections.Counter(x[2] for x in SOAN if x[2] != "本文")


# ============================================================ 確認事項を読む
def load_check():
    """build_process_control.py のソースから CHECK を読み取る。"""
    src = io.open(os.path.join(RP.ROOT, "build_process_control.py"),
                  encoding="utf-8").read()
    for node in ast.parse(src).body:
        if isinstance(node, ast.Assign) and any(
                getattr(t, "id", None) == "CHECK" for t in node.targets):
            return ast.literal_eval(node.value)
    raise RuntimeError("CHECK が見つからない")


CHECK = load_check()
# (0)No (1)区分 (2)起票日 (3)表題 (4)内容 (5)止めている成果物
# (6)確認先 (7)状態 (8)期限 (9)備考
KANRYO = ("完了", "了承済", "了承済（保管せず廃棄）")
MATI = [x for x in CHECK if x[7] not in KANRYO]
GAISAN_LIST = sorted([x for x in MATI if "第1次概算" in x[5]],
                     key=lambda x: (x[8], x[0]))
SOAN_LIST = [x for x in MATI if "計画素案" in x[5]]
NO_SET = set(x[0] for x in CHECK)


# ============================================================ 00
ws = sheet("00_この表について",
           "計画素案の課題整理",
           "確認事項は業務工程管理表で、素案の未確定箇所は本文の"
           "［要協議］［要確認］［要内訳］で保持しており、両者は別の場所にある。"
           "本表は両者を突き合わせ、"
           "素案のどこに未確定が集中しているか、"
           "どの確認事項が決まればどこが埋まるか、"
           "第1次概算までに何を決める必要があるかを整理する。"
           "基準日 " + KIJUNBI + "。",
           [4, 30, 14, 44])

r = 4
r = lead(ws, r, "1　全体の数", span=4)
r = header(ws, r, ["", "項目", "件数", "内容"])
for i, (a, b, c_) in enumerate([
    ("計画素案の未確定箇所", "{}箇所".format(PH_KEI),
     "要協議{}・要確認{}・要内訳{}。"
     "本文{}箇所・表{}種{}箇所（01シート）"
     .format(PH_SHU.get("要協議", 0), PH_SHU.get("要確認", 0),
             PH_SHU.get("要内訳", 0), len(PH_HON),
             len(PH_HYO), sum(PH_HYO.values()))),
    ("確認事項（全件）", "{}件".format(len(CHECK)),
     "業務工程管理表 03_確認事項一覧。うち完了・了承済{}件"
     .format(len(CHECK) - len(MATI))),
    ("確認事項（未回答）", "{}件".format(len(MATI)),
     "このうち計画素案を止めているものが{}件（02シート）"
     .format(len(SOAN_LIST))),
    ("第1次概算までに決着を要するもの", "{}件".format(len(GAISAN_LIST)),
     GAISAN + "を止めている確認事項（03シート）"),
], start=1):
    r = body(ws, r, [i, a, b, c_], height=40)

r = note(ws, r + 1,
         "注1）未確定箇所は計画素案の実物（{}）から数えています。"
         "確認事項は `build_process_control.py` のソースから読み取っています。"
         "いずれも固定値を書いていないため、"
         "素案又は確認事項を更新すれば本表も追随します。\n"
         "注2）本表は素案の本文には反映しません。"
         "計画素案の本文に確認事項の注記を書かないという指示によります。"
         "本文に残すのは図表を読むために必要な注記"
         "（出典・単位・欠測・数値の性質・記号の意味）に限っています。\n"
         "注3）［要協議］［要確認］［要内訳］は本文に残します。"
         "確定しない値を推測して埋めることはしません。"
         .format(RP.draft_label()), span=4)

r += 1
r = lead(ws, r, "2　本表で分かること", span=4)
r = header(ws, r, ["", "分かること", "所在", "内容"])
for i, (a, b, c_) in enumerate([
    ("未確定は第6章第4節に集中している",
     "01シート",
     "{}箇所のうち{}箇所（{:.0f}％）が"
     "第6章第4節（施設の見込み）にあります。"
     "日常生活圏域ごとの必要利用定員総数（確認事項No.88）が"
     "決まらないためです"
     .format(PH_KEI, PH_SETSU.get(("第6章　介護保険事業等の見込み",
                                   "第4節 介護保険サービス等における施設の見込み"), 0),
             PH_SETSU.get(("第6章　介護保険事業等の見込み",
                           "第4節 介護保険サービス等における施設の見込み"), 0)
             / PH_KEI * 100)),
    ("第1次概算を止めているものが{}件ある".format(len(GAISAN_LIST)),
     "03シート",
     "9月26日まで残り期間が短く、"
     "うち期限が令和8年9月のものが{}件あります"
     .format(len([x for x in GAISAN_LIST if x[8].startswith("R8.9")]))),
    ("令和8年9月に新たに判明した論点が素案の複数節に及ぶ",
     "04シート",
     "認知症基本法・地域ケア会議のKPI・年報の要介護度別・"
     "調査の代表性・区域を越えた施設利用の5群です"),
    ("受託者の作業だけで進むものが残っている",
     "06シート",
     "決定を待たずに着手できるものを挙げています"),
], start=1):
    r = body(ws, r, [i, a, b, c_], height=52)


# ============================================================ 01
ws = sheet("01_素案の未確定箇所と記述量",
           "計画素案の未確定箇所の所在と節ごとの記述量（実物から数える）",
           "計画素案（{}）の［要協議］［要確認］［要内訳］を"
           "章節別・表別に数えたもの。合計{}箇所。"
           "あわせて節ごとの記述量を測り、記述が薄い節を示す。"
           .format(RP.draft_label(), PH_KEI),
           [4, 30, 34, 10, 10, 10, 10, 34])

r = 4
r = lead(ws, r, "1　章別", span=8)
r = header(ws, r, ["", "章", "", "計", "要協議", "要確認", "要内訳",
                   "内容"])
SHO_NAIYO = {
    "第6章　介護保険事業等の見込み":
        "必要利用定員総数と見込量の表。確認事項No.88・No.100〜No.101ほか",
    "第5章　施策の展開":
        "施策ごとの目標値。基準値がないもの・事業設計の決定を待つもの",
    "第2章　高齢者及び介護保険の状況":
        "調査の点検事項と出所の食い違い。確認事項No.8ほか",
    "第4章　計画の基本理念及び基本目標":
        "代表KPIの定義変更と目標の方向。確認事項No.4ほか",
    "第3章　第9期計画の実施状況と評価":
        "第9期評価の表現と役割分担",
}
for sho, n in PH_SHO.most_common():
    kk = collections.Counter(x[3] for x in SOAN if x[0] == sho)
    r = body(ws, r, ["", sho, "", n, kk.get("要協議", 0),
                     kk.get("要確認", 0), kk.get("要内訳", 0),
                     SHO_NAIYO.get(sho, "")],
             fills={4: (NG_O if n >= 30 else None)}, height=32)
r = body(ws, r, ["", "計", "", PH_KEI, PH_SHU.get("要協議", 0),
                 PH_SHU.get("要確認", 0), PH_SHU.get("要内訳", 0), ""],
         bold=True, height=20)

r += 1
r = lead(ws, r, "2　節別（3箇所以上）", span=8)
r = header(ws, r, ["", "章", "節", "計", "要協議", "要確認", "要内訳",
                   "主に止めている確認事項"])
SETSU_KAKUNIN = {
    "第4節 介護保険サービス等における施設の見込み":
        "No.88（必要利用定員総数）・No.113（混合型特定施設）・"
        "No.141（区域を越えた施設利用）",
    "第3節 施策の体系と成果指標": "No.4（代表KPIの代理指標）",
    "第2節 介護給付等サービスの見込量":
        "No.100・No.101（基準年度と数値の差し替え）・No.131（表に何を掲げるか）",
    "第8節 中長期推計からみた需要と財政": "No.114（中長期推計の範囲）",
    "第4節 第9期計画の評価と第10期への課題": "No.75（代表KPIの再整理案）",
    "第3節 サービス提供体制の状況": "No.8（3調査の点検事項）",
    "第6節 介護事業所の現状": "No.8（同）",
    "第1節 高齢者の状況": "No.3（町別人口）",
    "第1節 要介護認定者数の見込み": "No.1（認定率のシナリオ）",
    "第6節 第1号被保険者の介護保険料試算": "No.86（交付金の控除）ほか",
    "第5節 在宅介護の実態": "No.8（同）",
    "第3節 介護保険サービス事業費の見込み": "No.112（地域支援事業の量）",
    "基本目標1　地域包括ケア・相談・医療介護連携":
        "確認事項B No.19（医療側の連携の実施状況）",
    "基本目標2　介護予防・生活支援・認知症・家族支援":
        "No.118〜No.121（地域ケア会議のKPI）・No.130（認知症基本法）",
    "基本目標3　持続可能なサービス提供・住まい":
        "No.98（省令上の類型）・確認事項A No.15（24時間対応の確保方策）",
    "基本目標4　介護人材・生産性・質・経営支援":
        "No.10（採用者数・退職者数）・No.116（人材の制約）",
    "基本目標5　保険運営・レジリエンス・PDCA":
        "No.91〜No.94（給付適正化）・確認事項A No.14（住宅改修等の点検）",
}
for (sho, setsu), n in sorted(PH_SETSU.items(), key=lambda x: -x[1]):
    if n < 3:
        continue
    kk = collections.Counter(x[3] for x in SOAN
                             if x[0] == sho and x[1] == setsu)
    r = body(ws, r, ["", sho, setsu or "（節の前）", n,
                     kk.get("要協議", 0), kk.get("要確認", 0),
                     kk.get("要内訳", 0),
                     SETSU_KAKUNIN.get(setsu, "")],
             fills={4: (NG_O if n >= 30 else None)}, height=32)

r += 1
r = lead(ws, r, "3　本文の未確定箇所（{}件）".format(len(PH_HON)), span=8)
r = header(ws, r, ["", "章", "節", "種別", "本文（抜粋）", "", "", ""])
for sho, setsu, _i, shu, t in PH_HON:
    r = body(ws, r, ["", sho, setsu, shu, t[:180], "", "", ""],
             fills={4: (IN_Y if shu == "要協議" else None)}, height=44)

r = note(ws, r + 1,
         "注1）表の中の未確定箇所は{}種の表に{}箇所あります。"
         "多くは施策の目標値の欄と必要利用定員総数の欄です。\n"
         "注2）本シートは実物から数えています。"
         "素案を更新して本表を作り直せば数は自動で変わります。"
         "固定値で書くと実物とずれるため、数える形にしています。"
         .format(len(PH_HYO), sum(PH_HYO.values())), span=8)

r += 1
r = lead(ws, r, "4　記述が薄い節（未確定箇所とは別の課題）", span=8)
r = header(ws, r, ["", "章", "節", "段落", "うち注・出典", "表",
                   "未確定", "本来入るべき内容と所在"])
USUI_NAIYO = {
    "第4節 高齢者の生活実態":
        "健康とくらしの調査（4,729票）の結果が本文に入っていない。"
        "集計は調査クロス集計・分析（24シート）にあるが、"
        "本節は導入文と資料注のみである。"
        "アンケート分析の未了No.4に当たる",
    "第5節 介護施設整備に係る基本方針":
        "整備の方針。第4節4の必要利用定員総数（確認事項No.88）の"
        "決定を待っている",
    "第7節 低所得者支援":
        "所得段階と公費軽減。第10期の政令改正（確認事項No.33）を待っている",
    "第4節　計画期間": "計画期間の記述であり、短いことは妥当",
    "第5節　計画の策定体制": "策定体制の記述であり、短いことは妥当",
    "第1節 計画の基本理念": "基本理念であり、短いことは妥当",
    "第2節 計画の基本目標": "基本目標の一覧であり、短いことは妥当",
    "第1節　計画策定の趣旨": "趣旨の記述であり、短いことは妥当",
    "第2節　計画の位置付け":
        "他計画との関係の記述であり短いことは妥当だが、"
        "認知症基本法第13条の市町村認知症施策推進計画をどう扱うかが"
        "決まれば本節に加筆する（確認事項No.130）。"
        "給付適正化計画（第7期）の位置づけは反映済み",
}
for s, se, p_, t_, n_ in USUI:
    ph = PH_SETSU.get((s, se), 0)
    y = USUI_NAIYO.get(se, "［要確認］")
    r = body(ws, r, ["", s, se, p_, n_, t_, ph, y],
             fills={4: (NG_O if "妥当" not in y else None)}, height=44)

_toriaezu = [x for x in USUI
             if "妥当" not in USUI_NAIYO.get(x[1], "［要確認］")]
r = note(ws, r + 1,
         "注3）未確定箇所（［要協議］等）は「書く内容は決まっているが"
         "値が決まらない」箇所です。"
         "本節は「書くべき内容がまだ入っていない」節を示すもので、"
         "性質が異なります。"
         "記号が置かれていないため未確定箇所の集計には現れません。\n"
         "注4）記述が薄い節は{}節あり、"
         "うち短いことが妥当なもの（計画期間・基本理念など）を除くと"
         "{}節です。\n"
         "注5）第2章第4節は{}段落・表{}と最も薄く、"
         "調査の集計結果が本文に入っていない状態です。"
         "受託者の作業だけで進むため、着手します（06シート）。"
         .format(len(USUI), len(_toriaezu),
                 next((x[2] for x in USUI if x[1] == "第4節 高齢者の生活実態"), 0),
                 next((x[3] for x in USUI if x[1] == "第4節 高齢者の生活実態"), 0)),
         span=8)


# ============================================================ 02
ws = sheet("02_確認事項の棚卸し",
           "確認事項の棚卸し（{}件）".format(len(CHECK)),
           "業務工程管理表 03_確認事項一覧を"
           "状態・期限・確認先・止めている成果物の別に集計したもの。"
           "同スクリプトのソースから読み取っている。",
           [4, 26, 12, 12, 12, 12, 44])

r = 4
r = lead(ws, r, "1　状態別", span=7)
r = header(ws, r, ["", "状態", "件数", "", "", "", "内容"])
for k, v in collections.Counter(x[7] for x in CHECK).most_common():
    r = body(ws, r, ["", k, v, "", "", "",
                     "完了・了承済を除く未回答は{}件".format(len(MATI))
                     if k == "確認待ち" else ""],
             fills={3: (NG_O if k == "確認待ち" else OK_G)}, height=18)

r += 1
r = lead(ws, r, "2　期限別（未回答のみ）", span=7)
r = header(ws, r, ["", "期限", "件数", "うち計画素案を止めている",
                   "うち第1次概算を止めている", "", "内容"])
for k, v in sorted(collections.Counter(x[8] for x in MATI).items()):
    so = len([x for x in MATI if x[8] == k and "計画素案" in x[5]])
    ga = len([x for x in MATI if x[8] == k and "第1次概算" in x[5]])
    r = body(ws, r, ["", k, v, so, ga, "",
                     "期限が過ぎているもの" if k < "R8.9" else ""],
             fills={3: (NG_O if k < "R8.9" else None)}, height=18)

r += 1
r = lead(ws, r, "3　確認先別（未回答のみ）", span=7)
r = header(ws, r, ["", "確認先", "件数", "", "", "", "内容"])
for k, v in collections.Counter(x[6] for x in MATI).most_common():
    r = body(ws, r, ["", k, v, "", "", "", ""], height=18)

r = note(ws, r + 1,
         "注1）期限が令和8年8月以前のもの（{}件）は期限を過ぎています。"
         "いずれも当方から催促する性質のものではなく、"
         "決定又は資料のご提供をお待ちしている状態です。\n"
         "注2）確認先が「発注者・3町」のもの（{}件）は、"
         "3町への照会の様式と時期の決定を先に要します。\n"
         "注3）本シートは件数の把握のためのものです。"
         "個々の内容は業務工程管理表 03_確認事項一覧をご覧ください。"
         .format(len([x for x in MATI if x[8] < "R8.9"]),
                 len([x for x in MATI if x[6] == "発注者・3町"])), span=7)


# ============================================================ 03
ws = sheet("03_第1次概算までに決着",
           "第1次概算までに決着を要するもの（{}件）".format(len(GAISAN_LIST)),
           GAISAN + "を止めている確認事項。"
           "決着しない場合の当方の扱い（既定値）を併せて示す。"
           "決着しなくても概算は出せるが、その前提を明示することになる。",
           [4, 8, 34, 12, 12, 44])

KITEI = {
    88: "3町（3圏域）を基盤整備単位とする案のまま概算する。"
        "必要利用定員総数は［要協議］のまま残す",
    100: "採用パターンP3（令和7年度基準・総括表・"
         "認定者数シナリオ②）で概算する",
    101: "素案 第6章第2節の表は見える化の自然体推計値のまま置き、"
         "P3による値を別掲する",
    103: "令和8年度を基準年度に用いず、季節補正も行わないまま概算する",
    105: "施策反映を行わないまま概算する（安全側）",
    106: "一律の伸びのまま概算し、補正の効き（月額▲58円又は＋48円）を併記する",
    107: "施設・居住系の実績の減少を将来へ延長しないまま概算する",
    110: "給付適正化の効果を織り込まないまま概算する（安全側）",
    113: "混合型特定施設の必要利用定員総数を定めないまま概算する",
    123: "見える化システムの推計値は現行の自然体推計のまま用いる",
    126: "基本推計に近いP3で概算する",
    131: "第6章第2節は利用者数のみを掲げたまま概算する",
    132: "在宅の実績見込み値は当方の入力案のまま置く",
    137: "1人1月あたり給付費は令和7年度で固定して概算する"
         "（月額で＋185〜＋374円に当たる）",
    141: "必要利用定員総数は区域内の施設の定員とは別に整理したまま概算する",
    98: "省令上の類型を計画に掲げないまま概算する",
    82: "町別按分は当方の案（第1次概算）のまま示す",
    142: "本表そのものに係る確認事項である。"
         "決着しない場合は、本シートの既定値によったことを"
         "第1次概算の資料に明示して提出する",
}
r = 4
r = header(ws, r, ["", "No.", "確認事項", "期限", "確認先",
                   "決着しない場合の当方の扱い（既定値）"])
for i, x in enumerate(GAISAN_LIST, start=1):
    r = body(ws, r, [i, "No.%d" % x[0], x[3], x[8], x[6],
                     KITEI.get(x[0], "［要協議］のまま残す")],
             fills={4: (NG_O if x[8].startswith("R8.9") else IN_Y)},
             height=44)

r = note(ws, r + 1,
         "注1）{}件のうち期限が令和8年9月のものは{}件です。\n"
         "注2）既定値はいずれも当方が作業上置いているものであり、"
         "ご了承をいただいたものではありません。"
         "第1次概算をお出しする際は、"
         "どの既定値によったかを明示します。\n"
         "注3）安全側と記したものは、"
         "決着しないまま概算すると保険料が高めに出る向きのものです。"
         "施策の達成や適正化の効果を前提に下げると、"
         "保険料が不足する側に誤ります。\n"
         "注4）No.137（単価の年度）は月額で＋185〜＋374円に当たり、"
         "本表のうち最も効きが大きいものです。"
         .format(len(GAISAN_LIST),
                 len([x for x in GAISAN_LIST
                      if x[8].startswith("R8.9")])), span=6)


# ============================================================ 04
ws = sheet("04_9月の新たな知見と割付け",
           "令和8年9月に新たに判明した論点と計画素案への割付け",
           "令和8年9月10日から12日までに判明した論点を、"
           "計画素案のどの箇所に効くかで整理する。"
           "いずれも確認事項として起票済みである。",
           [4, 12, 26, 30, 30, 12])

SHIRUSHI = [
    ("No.130", "認知症基本法第13条",
     "共生社会の実現を推進するための認知症基本法（令和5年法律第65号）は"
     "市町村認知症施策推進計画について定めているが、"
     "計画素案に同法への言及がない",
     "第1章第2節（計画の位置付け）・第5章 基本目標2。"
     "策定主体を広域連合とするか3町とするかの決定を要する",
     "R8.9"),
    ("No.140", "調査①の代表性",
     "在宅生活改善調査の要介護度分布は在宅の認定者の分布と食い違う"
     "（χ²＝43.5・自由度6。1％点16.81）。"
     "要介護3が期待の2.8倍、要支援1が0.16倍",
     "第2章第5節（在宅介護の実態）・第4章第3節（代表KPI H12）。"
     "「在宅で困難を抱える方の状況」として引く整理に改める",
     "R8.9"),
    ("No.141", "区域を越えた施設利用",
     "施設の在籍者と保険者の受給者は一致しない。"
     "介護老人保健施設だけが在籍206人＞受給134.1人／月。"
     "特別養護老人ホーム等は受給197.5人／月が回答施設の定員152人を上回る",
     "第6章第4節4（必要利用定員総数）。"
     "区域内の施設の定員とは別に整理する",
     "R8.9"),
    ("No.134〜No.138", "年報の要介護度別と見える化の入力",
     "年報（令和7年度）の要介護度別の明細を収めたことにより、"
     "見える化システムの実績見込み値を確定値で置けるようになった。"
     "入力欄は人数が整数・利用回（日）数が小数第1位まで",
     "第6章第2節（見込量）。"
     "No.136（画面と年報の食い違い）とNo.138（認定者数の編集項目）は"
     "入力を始める前の確認を要する",
     "R8.9〜10"),
    ("No.137", "1人1月あたり給付費の実績値の年度",
     "見える化システムに令和7年度とするか令和8年度とするかを選ぶ項目があり、"
     "施設・居住系と在宅で別に選べる。"
     "令和7年度で固定することは月額で＋185〜＋374円に当たる",
     "第6章第2節・第6節（保険料試算）。"
     "第1次概算の前提として決定を要する",
     "R8.9"),
    ("No.118〜No.122", "地域ケア会議のKPI",
     "施策2-5は第9期評価で「指標不足」と判定済み。"
     "3視点13指標のKPI設計案を作成したが、"
     "12指標は現時点で測れない",
     "第5章 基本目標2。"
     "令和9年度を基準値の把握に充て、令和10・11年度に目標値を置く",
     "R8.10"),
    ("No.131", "第6章第2節の導入文と表の食い違い",
     "導入文は「利用者数、回数・日数、必要定員及び給付費」としているが、"
     "表は利用者数のみである",
     "第6章第2節。表頭の「R8実績」の注記は改めた",
     "R8.9"),
    ("No.112", "地域支援事業の量の見込み",
     "法第117条第2項第2号の基本的記載事項であるが、"
     "第6章第3節2は費用の額のみで量の表がない",
     "第6章第3節2。総合事業の令和7年度実績の提供を要する",
     "R8.10"),
    ("No.139", "健康とくらしの調査の個票の再提供",
     "調査クロス集計・分析（24シート）を作るスクリプトは個票を読むため"
     "再実行できない。成果品は収録済みで内容は確定している",
     "第2章第4節。作り直しが必要になったときだけ再提供を要する",
     "R8.9"),
]
r = 4
r = header(ws, r, ["", "確認事項", "論点", "内容", "素案への割付け",
                   "期限"])
for i, (no, ron, nai, wari, kig) in enumerate(SHIRUSHI, start=1):
    r = body(ws, r, [i, no, ron, nai, wari, kig], height=64)

r = note(ws, r + 1,
         "注1）{}群のうち、素案の本文に加筆を要するのは"
         "No.130（認知症基本法）・No.131（第6章第2節の導入文）・"
         "No.112（地域支援事業の量）の3件です。"
         "いずれも決定を待って加筆します。\n"
         "注2）No.140（調査①の代表性）は、"
         "既に本文にある記述の引き方を改めるものです。"
         "数値そのものは変わりません。\n"
         "注3）新たに節や章を起こすものは「諮る資料」として別に作成し、"
         "計画素案には入れない整理としています。"
         .format(len(SHIRUSHI)), span=6)


# ============================================================ 05
ws = sheet("05_アンケート分析の未了",
           "アンケート分析の未了と計画素案への効き",
           "未了8件の現況と、それぞれが素案のどの箇所を止めているか。"
           "うち受託者の作業だけで進むものは1件である。",
           [4, 26, 34, 26, 12, 12])

MIRYO = [
    ("照会票の発出", "点検事項のうち事業所への照会により解消できるもの"
     "（重大6件・要確認21件）。文面と管理表は作成済み",
     "第2章第3節・第5節・第6節（未確定{}箇所）"
     .format(PH_SETSU.get(("第2章　高齢者及び介護保険の状況",
                           "第3節 サービス提供体制の状況"), 0)
             + PH_SETSU.get(("第2章　高齢者及び介護保険の状況",
                             "第5節 在宅介護の実態"), 0)
             + PH_SETSU.get(("第2章　高齢者及び介護保険の状況",
                             "第6節 介護事業所の現状"), 0)),
     "No.8", "確認待ち"),
    ("確定値の確定", "介護職員の総数、小規模多機能の需要、施設利用者数、"
     "特別養護老人ホームの待機者数の採用値",
     "同上。成果品11箇所の更新を伴う", "No.8", "未着手"),
    ("成果品11箇所の更新", "受領点検と集計・集計分析報告書・結果報告書・"
     "計画素案・図表集",
     "同上", "No.8", "未着手"),
    ("④の分析結果の計画本文への反映",
     "健康とくらしの調査の地区別集計・年齢調整による地域比較・"
     "社会参加や移動制約との関連分析",
     "第2章第4節（高齢者の状況）・第5章 基本目標1",
     "No.9（地区の定義の部分）", "実施中"),
    ("代表KPIの振替の確定", "H07・H08を健康とくらしの調査へ、"
     "H12を在宅生活改善調査へ振り替える案",
     "第4章第3節・資料1（未確定{}箇所）"
     .format(PH_SETSU.get(("第4章　計画の基本理念及び基本目標",
                           "第3節 施策の体系と成果指標"), 0)),
     "No.4", "確認待ち"),
    ("H14の基準値の算定", "採用率と離職率の差。"
     "各年9月30日現在の在籍者数と年間の採用者数・退職者数を要する",
     "第4章第3節", "No.10", "確認待ち"),
    ("回収率の確定", "①②③は配布数の記録がないため回収率を算定していない。"
     "公表データによる母数を用いた把握率を暫定値としている",
     "第2章第5節・第6節", "No.9", "確認待ち"),
    ("未回答施設の取扱いの確定",
     "居所変更実態調査は区域内31施設のうち13施設が未回答。"
     "把握率58.1％のまま用いるか追加照会を行うか",
     "第2章第3節・第6章第4節4", "No.8", "確認待ち"),
]
r = 4
r = header(ws, r, ["", "未了の作業", "内容", "素案のどこを止めているか",
                   "確認事項", "状態"])
for i, (a, b, c_, d, e) in enumerate(MIRYO, start=1):
    r = body(ws, r, [i, a, b, c_, d, e],
             fills={6: {"実施中": OK_G, "確認待ち": IN_Y,
                        "未着手": GRAY}.get(e)}, height=52)

_jisshi = [x for x in MIRYO if x[4] == "実施中"]
r = note(ws, r + 1,
         "注1）8件のうち受託者の作業だけで進むのは{}件（④の反映）です。"
         "5件は発注者のご判断又は資料の提供を待っており、"
         "2件は前の作業が終わらないと着手できません。\n"
         "注2）No.8（3調査の点検事項の取扱い）が最も多くの作業を"
         "止めています。この1件の決定により、"
         "照会の発出から成果品の更新まで一連の作業が動きます。\n"
         "注3）令和8年9月11日の突合（確認事項No.140）により、"
         "在宅生活改善調査の結果は"
         "在宅の認定者全体を代表しないことが確かめられました。"
         "代表KPI H12の振替（No.4）は、"
         "同じ母集団の指標となることを踏まえてご判断いただく必要があります。"
         .format(len(_jisshi)), span=6)


# ============================================================ 06
ws = sheet("06_受託者の作業だけで進むもの",
           "決定を待たずに着手できるもの",
           "確認事項の回答を待たずに受託者の作業だけで進むもの。"
           "本表の作成時点で着手できるものを挙げる。",
           [4, 30, 40, 16, 14, 24])

DEKIRU = [
    ("④の分析結果の第2章第4節への反映",
     "健康とくらしの調査の地区別集計・年齢調整による地域比較を"
     "本文に反映する。地区の定義（小学校区とするか）のみ確認を要するため、"
     "定義に依存しない部分から進める",
     "第2章第4節", "着手できる",
     "アンケート分析 未了No.4"),
    ("中長期推計（令和22年度まで）の算定",
     "P3の1人あたり給付費を固定し認定者数を乗じれば算定できる。"
     "認定者数はR17 2,228人・R22 2,288人まで算定済み",
     "第6章第1節・第8節", "着手できる",
     "確認事項No.114。第2次概算までに行う"),
    ("地域支援事業の量の見込みの様式作成",
     "第3章第2節の第9期実績から様式を起こすことはできる。"
     "数値は総合事業の令和7年度実績のご提供を待つ",
     "第6章第3節2", "様式のみ着手できる",
     "確認事項No.112"),
    ("見える化システムの入力値の確定",
     "年報の要介護度別により②③④は確定値で用意済み。"
     "①認定者数の編集項目の有無のみ確認を要する",
     "第6章第2節", "確認待ち（①のみ）",
     "確認事項No.138"),
    ("認知症基本法への対応方針の整理",
     "策定主体・実施主体・意見聴取の3点を整理した協議資料は作成できる。"
     "計画素案への反映は決定を待つ",
     "第1章第2節・第5章 基本目標2", "協議資料は着手できる",
     "確認事項No.130"),
    ("素案の未確定箇所の表示の統一",
     "［要協議］［要確認］［要内訳］の使い分けを揃える。"
     "現在は要協議{}・要確認{}・要内訳{}で、"
     "同じ性質のものに異なる記号を用いている箇所がある"
     .format(PH_SHU.get("要協議", 0), PH_SHU.get("要確認", 0),
             PH_SHU.get("要内訳", 0)),
     "素案全体", "着手できる", "本表01シート"),
]
r = 4
r = header(ws, r, ["", "作業", "内容", "素案の箇所", "状態", "関係"])
for i, (a, b, c_, d, e) in enumerate(DEKIRU, start=1):
    r = body(ws, r, [i, a, b, c_, d, e],
             fills={5: (OK_G if "着手できる" in d else IN_Y)}, height=56)

r = note(ws, r + 1,
         "注1）本シートは当方の作業予定であり、"
         "ご判断をお願いするものではありません。\n"
         "注2）着手する順序は、"
         "第1次概算（9月26日）に効くものを先にします。"
         "中長期推計は第2次概算（10月30日）までに行います。", span=6)


# ============================================================ 07
ws = sheet("07_自己点検", "自己点検（エラーチェック）",
           "本表の数え方と、確認事項との突合を点検する。",
           [4, 44, 26, 34, 10, 10])

chk(1, "素案の未確定箇所を実物から数えていること",
    "計画素案の docx",
    "{}／{}箇所".format(RP.draft_label(), PH_KEI), PH_KEI > 0)

chk(2, "種別の和が合計と一致すること",
    "要協議＋要確認＋要内訳",
    "{}＝{}".format(sum(PH_SHU.values()), PH_KEI),
    sum(PH_SHU.values()) == PH_KEI)

chk(3, "章別の和が合計と一致すること",
    "{}章".format(len(PH_SHO)),
    "{}＝{}".format(sum(PH_SHO.values()), PH_KEI),
    sum(PH_SHO.values()) == PH_KEI)

chk(4, "確認事項をソースから読み取っていること",
    "build_process_control.py の CHECK",
    "{}件（未回答{}件）".format(len(CHECK), len(MATI)), len(CHECK) > 0)

chk(5, "確認事項の番号に重複がないこと",
    "No の一意性",
    "{}件／一意{}件".format(len(CHECK), len(NO_SET)),
    len(CHECK) == len(NO_SET))

_04no = []
for no, _r, _n, _w, _k in SHIRUSHI:
    for m in re.finditer(r"No\.(\d+)", no):
        _04no.append(int(m.group(1)))
_nai = [n for n in _04no if n not in NO_SET]
chk(6, "04シートに挙げた確認事項が一覧に実在すること",
    "{}件".format(len(_04no)),
    "全件実在" if not _nai else "不明 " + "・".join(str(x) for x in _nai),
    not _nai)

_05no = []
for x in MIRYO:
    for m in re.finditer(r"No\.(\d+)", x[3]):
        _05no.append(int(m.group(1)))
_nai5 = [n for n in _05no if n not in NO_SET]
chk(7, "05シートに挙げた確認事項が一覧に実在すること",
    "{}件".format(len(_05no)),
    "全件実在" if not _nai5 else "不明 " + "・".join(str(x) for x in _nai5),
    not _nai5)

_kt = [n for n in KITEI if n not in NO_SET]
chk(8, "03シートの既定値が一覧に実在する確認事項に対応すること",
    "{}件".format(len(KITEI)),
    "全件実在" if not _kt else "不明 " + "・".join(str(x) for x in _kt),
    not _kt)

_gno = set(x[0] for x in GAISAN_LIST)
_km = [n for n in _gno if n not in KITEI]
chk(9, "第1次概算を止めている全件に既定値を置いていること",
    "{}件".format(len(GAISAN_LIST)),
    "全件に置いている" if not _km
    else "未記載 " + "・".join("No.%d" % n for n in sorted(_km)),
    not _km)

_s6 = PH_SETSU.get(("第6章　介護保険事業等の見込み",
                    "第4節 介護保険サービス等における施設の見込み"), 0)
chk(10, "未確定箇所が第6章第4節に集中していること",
    "同節／全体",
    "{}／{}箇所（{:.0f}％）".format(_s6, PH_KEI, _s6 / PH_KEI * 100),
    _s6 / PH_KEI >= 0.3)

chk(11, "記述が薄い節を記述量から判定していること",
    "本文（注・出典を除く）5段落未満かつ表1以下",
    "{}節（うち短いことが妥当なものを除くと{}節）"
    .format(len(USUI), len(_toriaezu)), len(USUI) > 0)

chk(12, "第2章第4節に調査結果が入っていないことを示していること",
    "同節の段落数と表数",
    "{}段落・表{}。健康とくらしの調査の集計は"
    "調査クロス集計・分析（24シート）にある"
    .format(next((x[2] for x in USUI if x[1] == "第4節 高齢者の生活実態"), -1),
            next((x[3] for x in USUI if x[1] == "第4節 高齢者の生活実態"), -1)),
    any(x[1] == "第4節 高齢者の生活実態" for x in USUI))

chk(13, "計画素案の本文に確認事項の注記を書いていないこと",
    "本表は素案を変更しない",
    "本スクリプトは素案を読むだけで書き換えていない", True)

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
         "注1）点検6〜9は、本表に書いた確認事項番号が"
         "業務工程管理表に実在することを確かめるものです。"
         "番号だけを書き写すと一覧との食い違いが起きるため、"
         "ソースから読み取った番号と突き合わせています。\n"
         "注2）点検9は、第1次概算を止めている確認事項の全件について"
         "「決着しない場合の扱い」を置いているかを確かめるものです。"
         "置いていないものがあると、概算の前提を説明できません。\n"
         "注3）1件でも不適合があると、本表を作るスクリプトは"
         "終了コード1で終わります。", span=6)


# ============================================================ 08
ws = sheet("08_確認事項", "確認事項",
           "本表に関してご判断をお願いする事項。"
           "番号は業務工程管理表 03_確認事項一覧による。",
           [7, 30, 46, 20, 12, 12])

r = 4
r = header(ws, r, ["No.", "確認事項", "内容", "止めている成果物",
                   "確認先", "回答期限"])
KAKUNIN = [
    ("No.142", "第1次概算を既定値のまま出してよいか",
     "{}を止めている確認事項が{}件あり、"
     "うち期限が令和8年9月のものが{}件ある。"
     "9月26日までに全件の決着は難しいと見込まれる。"
     "①決着していないものは当方の既定値（03シート）により概算し、"
     "どの既定値によったかを明示する扱いでよいか。"
     "②既定値のうち保険料を高めに出す向きのもの"
     "（施策反映を行わない・給付適正化の効果を織り込まない・"
     "施設の減少を延長しない）は安全側の置き方である。"
     "この向きでよいか。"
     "③効きが最も大きいのはNo.137（1人1月あたり給付費の年度。"
     "月額＋185〜＋374円）であり、"
     "これだけは9月26日までに決定をいただきたい。"
     .format(GAISAN, len(GAISAN_LIST),
             len([x for x in GAISAN_LIST if x[8].startswith("R8.9")])),
     "サービス見込量 第1次概算",
     "発注者", "R8.9"),
    ("No.143", "素案の未確定箇所の表示の統一",
     "計画素案の未確定箇所は{}箇所ある"
     "（要協議{}・要確認{}・要内訳{}）。"
     "①［要協議］は決定を要するもの、"
     "［要確認］は事実の確認を要するもの、"
     "［要内訳］は内訳の提供を要するものという使い分けでよいか。"
     "②現在は同じ性質のものに異なる記号を用いている箇所がある。"
     "使い分けを揃える作業は受託者の作業だけで進む。"
     "③策定委員会にお諮りする版では、"
     "未確定箇所をそのまま残すか、"
     "一覧表に移して本文からは外すかをご判断いただきたい。"
     .format(PH_KEI, PH_SHU.get("要協議", 0), PH_SHU.get("要確認", 0),
             PH_SHU.get("要内訳", 0)),
     "計画素案（策定委員会用）",
     "発注者", "R8.10"),
]
for no, ken, naiyo, tome, saki, kigen in KAKUNIN:
    r = body(ws, r, [no, ken, naiyo, tome, saki, kigen], height=170)
r = note(ws, r + 1,
         "注1）本表により新たに起票した確認事項は{}件"
         "（No.142・No.143）です。\n"
         "注2）本表は確認事項を新たに増やすことよりも、"
         "既にお願いしている{}件の優先順位を示すことを目的としています。"
         .format(len(KAKUNIN), len(MATI)), span=6)


# ============================================================ 出力
os.makedirs(ODIR, exist_ok=True)
wb.save(OUT)

print("出力：" + OUT)
print("計画素案：{}".format(RP.draft_label()))
print("未確定箇所：{}箇所（要協議{}・要確認{}・要内訳{}）"
      .format(PH_KEI, PH_SHU.get("要協議", 0), PH_SHU.get("要確認", 0),
              PH_SHU.get("要内訳", 0)))
print("  第6章第4節に{}箇所（{:.0f}％）が集中".format(_s6, _s6 / PH_KEI * 100))
print("確認事項：{}件（未回答{}件・計画素案を止めている{}件）"
      .format(len(CHECK), len(MATI), len(SOAN_LIST)))
print("第1次概算を止めているもの：{}件（うち令和8年9月期限{}件）"
      .format(len(GAISAN_LIST),
              len([x for x in GAISAN_LIST if x[8].startswith("R8.9")])))
print("記述が薄い節：{}節（短いことが妥当なものを除くと{}節）"
      .format(len(USUI), len(_toriaezu)))
print("受託者の作業だけで進むもの：{}件".format(len(DEKIRU)))
print("新たな確認事項：{}件（No.142・No.143）".format(len(KAKUNIN)))
print("自己点検：{}件（適合{}件・不適合{}件）"
      .format(len(CHECKS), len(CHECKS) - NG, NG))

if NG:
    sys.exit(1)
