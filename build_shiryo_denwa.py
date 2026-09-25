# -*- coding: utf-8 -*-
"""サービス見込量の算定に要する資料と、電話照会の進め方.

令和8年9月17日のご依頼
  「見込量算定を行う上での必要書類について今一度ご教示ください。
    電話連絡する際の例もお示しください。
    美瑛町（014597）『R7給付費データ集計／介護度別給付費シート（総計）』と
    『【大雪】R7.給付費データ集計』に相違がある点について、回答がないので」

本書は次の3つを1冊にまとめたものである。
  ① 見込量の算定に要する資料の一覧（何が届かないと何が止まるか）
  ② 資料が届かない場合の当方の扱い（既定値）
  ③ 電話で照会するときの言い方の例（逐語）

**数値は算定スクリプト及びデータモジュールから読む。固定値で書かない。**
  給付費データの差　`data_kyufu_jisseki.KYUFU_TOTAL`
  据え置き・確認事項　`build_mikomiryo_santei.py` の `SUEOKI`
  確認事項の期限　`build_process_control.py` のソースから `ast` で読む

**担当者名・電話番号は書かない**（発注者指示。CLAUDE.md §1）。
名宛人と連絡先は〔　〕の記入枠としている。

━━ 本件（確認事項No.87）は令和8年9月25日に解決した ━━

令和8年8月28日受領分では、令和7年度のみ3町のファイルの合計が
広域連合のファイルの表を129,580,267円（4.20％）上回っていた。

令和8年9月25日に令和7年度の広域連合単位のファイルの再提供を受け、
「介護度別給付費」の表の総計が3,213,054,558円となり、
要支援1から要介護5までの全7区分で3町合計と円単位で一致した。
従前のファイルも、その明細を集計すると同じ額になる。
**差は抽出条件の違いではなく、表が明細と一致していなかったことによる。**

第4節の例1は、照会の言い方の例として残す。

出力
  output/第10期計画_見込量算定に要する資料と電話照会の進め方.docx
"""

import docx_fix
import ast
import io
import os
import re
import runpy
import sys

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import data_kyufu_jisseki as KJ
import repo_paths as RP

OUT = (RP.ROOT + "/output/"
       "第10期計画_見込量算定に要する資料と電話照会の進め方.docx")
KIJUNBI = "令和8年9月17日"
FONT = "游ゴシック"
NAVY = RGBColor(0x1F, 0x38, 0x64)

TOWNS = ["東川町", "美瑛町", "東神楽町"]
DO = ("要支援1", "要支援2", "要介護1", "要介護2", "要介護3", "要介護4",
      "要介護5")
# 証記載保険者番号（市町村番号）。個人情報ではない。
HOKENSHA = {"東川町": "014589", "美瑛町": "014597", "東神楽町": "014530"}

CHECKS = []


def chk(no, naiyo, kekka, ok):
    CHECKS.append((no, naiyo, kekka, ok))
    return ok


# ============================================================ 事実の算定
def tot(machi, y):
    return KJ.KYUFU_TOTAL[(machi, y)]


def machi_kei(y):
    return [sum((tot(m, y)[i] or 0) for m in TOWNS) for i in range(8)]


SA = {}
for _y in ("R6", "R7", "R8"):
    a, b = machi_kei(_y), tot("大雪", _y)
    SA[_y] = (a, [b[i] or 0 for i in range(8)],
              [a[i] - (b[i] or 0) for i in range(8)])

# 令和8年8月28日受領分の【大雪】令和7年度の「介護度別給付費」の表の値。
# 再提供（令和8年9月25日）により差は解消したため、以下は経緯の記録である。
OLD = KJ.R7_KOIKI_KYU_HYO
SA_OLD = [SA["R7"][0][i] - OLD[i] for i in range(8)]
SA_R7 = SA_OLD[7]
SA_R7_P = SA_R7 / OLD[7] * 100
DO_P = [SA_OLD[i] / OLD[i] * 100 for i in range(7)]

# 令和8年度の町別構成比で広域連合の令和7年度を按分する
NOBI = {m: tot(m, "R7")[7] / tot(m, "R6")[7] - 1 for m in TOWNS}
NOBI["大雪"] = tot("大雪", "R7")[7] / tot("大雪", "R6")[7] - 1

chk(1, "3町計と広域連合が3か年とも一致すること",
    "／".join("%s 差%d円" % (y, SA[y][2][7]) for y in ("R6", "R7", "R8")),
    all(SA[y][2][7] == 0 for y in ("R6", "R7", "R8")))
chk(2, "令和7年度の要介護度別も全7区分で一致すること",
    "不一致 %d区分" % sum(1 for i in range(7) if SA["R7"][2][i] != 0),
    all(SA["R7"][2][i] == 0 for i in range(7)))
chk(3, "従前の表の値との差が129,580,267円であったこと",
    "差 %s円（%.2f％）" % ("{:,}".format(SA_R7), SA_R7_P),
    SA_R7 == 129580267)
chk(4, "従前の表との差が要介護度により一様でなかったこと",
    "%.2f％から%.2f％まで" % (min(DO_P), max(DO_P)),
    max(DO_P) - min(DO_P) > 2.0)
chk(5, "再提供後の3町計が従前の表の値を上回っていること",
    "3町計 %s円／従前の表 %s円"
    % ("{:,}".format(SA["R7"][0][7]), "{:,}".format(OLD[7])),
    SA["R7"][0][7] > OLD[7])
chk(6, "令和6年度から令和7年度への伸びが3町・広域連合とも正であること",
    "広域連合%+.2f％／3町%+.2f〜%+.2f％"
    % (NOBI["大雪"] * 100, min(NOBI[m] for m in TOWNS) * 100,
       max(NOBI[m] for m in TOWNS) * 100),
    NOBI["大雪"] > 0 and min(NOBI[m] for m in TOWNS) > 0)

MISMATCH = [x for x in KJ.SHEET_NAME_MISMATCH]
chk(7, "美瑛町 令和7年度のシート名の不一致が1件あること",
    "／".join("%s %s %s" % x for x in MISMATCH), len(MISMATCH) == 1)


# ============================================================ 据え置きを読む
def _load(name):
    class Sink(object):
        closed = False
        encoding = "utf-8"
        errors = "strict"
        newlines = None
        line_buffering = False
        name = "<sink>"
        mode = "w"

        def __init__(self):
            self.buffer = self

        def write(self, *_a, **_k):
            return 0

        def writelines(self, _l):
            pass

        def flush(self):
            pass

        def close(self):
            pass

        def fileno(self):
            raise OSError("sink")

        def isatty(self):
            return False

        def readable(self):
            return False

        def writable(self):
            return True

        def seekable(self):
            return False

        def detach(self):
            return self

        def reconfigure(self, **_k):
            pass

    old = sys.stdout
    sys.stdout = Sink()
    try:
        return runpy.run_path(RP.ROOT + "/" + name)
    finally:
        sys.stdout = old


_M = _load("build_mikomiryo_santei.py")
SUEOKI = _M["SUEOKI"]


def _kakunin():
    """確認事項の期限・確認先をソースから読む（固定値で書き写さない）。"""
    src = io.open(RP.ROOT + "/build_process_control.py",
                  encoding="utf-8").read()
    m = re.search(r"^CHECK = (\[.*?^\])", src, re.S | re.M)
    if not m:
        return {}
    out = {}
    for c in ast.literal_eval(m.group(1)):
        out[c[0]] = {"期限": c[2], "件名": c[3], "確認先": c[6],
                     "状態": c[7]}
    return out


KAKUNIN = _kakunin()
chk(8, "確認事項をソースから読めること",
    "%d件（No.87は「%s」）"
    % (len(KAKUNIN), KAKUNIN.get(87, {}).get("件名", "―")),
    87 in KAKUNIN)


# ============================================================ docx の体裁
doc = Document()
st = doc.styles["Normal"]
st.font.name = FONT
st.font.size = Pt(10.5)
st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
st.paragraph_format.space_after = Pt(4)
st.paragraph_format.line_spacing = 1.15
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.top_margin = sec.bottom_margin = Cm(2.0)
sec.left_margin = sec.right_margin = Cm(1.9)
TEXTW = 21.0 - 1.9 * 2


def _shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    # CT_TcPrBase は要素の順序が定められており、shd は
    # noWrap・tcMar・textDirection・tcFitText・vAlign・hideMark より前に置く
    # （CLAUDE.md §4）。
    after = None
    for tag in ("w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText",
                "w:vAlign", "w:hideMark"):
        after = tcPr.find(qn(tag))
        if after is not None:
            break
    if after is None:
        tcPr.append(shd)
    else:
        after.addprevious(shd)


def _cellmargin(t):
    mar = OxmlElement("w:tblCellMar")
    for tag, v in (("top", 0.06), ("left", 0.12), ("bottom", 0.06),
                   ("right", 0.12)):
        e = OxmlElement("w:" + tag)
        e.set(qn("w:w"), str(int(v * 567)))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    # CT_TblPrBase は要素の順序が定められており、
    # tblCellMar は tblLook・tblCaption・tblDescription より前に置く
    # （CLAUDE.md §4）。
    tblPr = t._tbl.tblPr
    after = None
    for tag in ("w:tblLook", "w:tblCaption", "w:tblDescription"):
        after = tblPr.find(qn(tag))
        if after is not None:
            break
    if after is None:
        tblPr.append(mar)
    else:
        after.addprevious(mar)


def P(text="", size=10.5, bold=False, align=None, indent=0.0, after=4):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold
    r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    if align:
        p.alignment = align
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(after)
    return p


def H1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = FONT
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = NAVY
    r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)


def H2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(9)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = FONT
    r.font.size = Pt(11)
    r.font.bold = True
    r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)


def BUL(text, size=10.5, mark="・"):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.6)
    p.paragraph_format.first_line_indent = Cm(-0.6)
    r = p.add_run(mark + text)
    r.font.name = FONT
    r.font.size = Pt(size)
    r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    return p


def SERIF(who, text):
    """電話の逐語。話し手を太字にして本文を続ける。"""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(1.6)
    p.paragraph_format.first_line_indent = Cm(-1.6)
    p.paragraph_format.space_after = Pt(5)
    a = p.add_run(who + "　")
    a.font.name = FONT
    a.font.size = Pt(10)
    a.font.bold = True
    a.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    b = p.add_run(text)
    b.font.name = FONT
    b.font.size = Pt(10)
    b.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    return p


def NOTE(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.4)
    r = p.add_run("※ " + text)
    r.font.name = FONT
    r.font.size = Pt(9)
    r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    return p


TBLNO = [0]


def CAP(text):
    TBLNO[0] += 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    r = p.add_run("表%d　%s" % (TBLNO[0], text))
    r.font.name = FONT
    r.font.size = Pt(10)
    r.font.bold = True
    r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)


def TBL(head, rows, widths, size=9, center=None, first_bold=False):
    center = center or set()
    t = doc.add_table(rows=1 + len(rows), cols=len(head))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    s = sum(widths)
    if s > TEXTW:
        widths = [w * TEXTW / s for w in widths]
    _cellmargin(t)
    tr = t.rows[0]
    tr._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
    for j, v in enumerate(head):
        c = tr.cells[j]
        c.width = Cm(widths[j])
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _shade(c, "1F3864")
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(v)
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    for i, row in enumerate(rows):
        tr = t.rows[i + 1]
        tr._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        for j, v in enumerate(row):
            c = tr.cells[j]
            c.width = Cm(widths[j])
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if i % 2 == 1:
                _shade(c, "F2F5FA")
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if j in center:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run("" if v is None else str(v))
            r.font.name = FONT
            r.font.size = Pt(size)
            r.font.bold = first_bold and j == 0
            r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    return t


def yen(v):
    return "{:,}".format(int(round(v)))


# ============================================================ 表紙
p = P("大雪地区広域連合　第10期介護保険事業計画策定支援業務", size=12,
      align=WD_ALIGN_PARAGRAPH.CENTER)
p = P("サービス見込量の算定に要する資料と　電話照会の進め方", size=17,
      bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
for r in p.runs:
    r.font.color.rgb = NAVY
P("基準日　%s　　受託者　ビズアップ公共コンサルティング株式会社" % KIJUNBI,
  size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
P("")

P("本書は、サービス見込量の算定を進めるうえで、"
  "何が届かないと何が止まるのかを一覧にし、"
  "届かない場合に当方がどう置くか（既定値）を示したうえで、"
  "電話で照会するときの言い方の例を添えたものです。")
P("直近の工程は、サービス見込量 第2次概算（令和8年10月30日）と"
  "予算編成用（令和8年11月13日）です。"
  "第1次概算（令和8年9月26日）は、"
  "資料が届かないものを既定値で置いて算定を終えており、"
  "提示できる状態にあります。")

# ============================================================ 第1節
H1("第1節　見込量の算定に要する資料")

H2("１　資料が届かないと算定そのものが止まるもの")
P("現時点で、算定そのものが止まるものはありません。"
  "資料が届かないものは既定値で置き、"
  "何をどう置いたかを別冊「サービス見込量の算定 第1次概算」15シートに"
  "残らず掲げています（23件）。"
  "以下は「届けば置き換えられる」という趣旨の一覧です。")

H2("２　保険料の月額を100円以上動かし得るもの（最優先）")
CAP("月額を100円以上動かし得る資料（3件）")
TBL(["資料", "入手先", "現在の置き方", "届いた場合の効き"],
    [["第10期の介護報酬改定率\n（令和9年度は改定の年）", "国（告示待ち）",
      "令和7年度の実績単価で固定",
      "改定率1％あたり月額＋61円。\n単価の趨勢を見込まないことは＋185〜＋374円"],
     ["第1号被保険者負担割合\n（第10期の政令）", "国（公布待ち）",
      "第9期と同じ23％",
      "1ポイントの違いで月額が約290円動く"],
     ["令和7年度末の\n介護給付費準備基金残高", "発注者",
      "取崩しを行わない（I＝0）",
      "1億円の取崩しで月額が約▲310円"]],
    [4.6, 2.6, 4.4, 5.6], first_bold=True)
NOTE("上の2件は国の告示・公布を待つほかなく、当方から催促できません。"
     "基金残高は発注者にご提供いただけるものです。")

H2("３　法定記載事項を満たすために要るもの")
CAP("記載事項に直結する資料")
TBL(["資料", "入手先", "何が定まらないか", "確認事項"],
    [["3町の施設整備・サービス提供方針",
      "3町（東川町は令和8年10月、\n東神楽町は令和8年10月・11月に検討予定）",
      "必要利用定員総数（介護保険法第117条第2項第1号）。\n"
      "計画素案 第6章第4節の未確定箇所58箇所がこれで解消する",
      "No.88・No.84"],
     ["認知症対応型共同生活介護\nくるみの郷の定員", "東川町",
      "区域内定員が99人＋［要確認］のまま。\n"
      "北海道の名簿は7事業所108人、見える化は5事業所99人で一致しない",
      "No.88"],
     ["特定施設の類型\n（混合型・介護専用型）", "北海道・発注者",
      "混合型特定施設の必要利用定員総数を定めるかどうか",
      "No.113"],
     ["地域密着型サービスの\n人員・設備・運営の基準条例", "広域連合",
      "令和6年厚生労働省令第16号の改正に対応しているかを確認できない。\n"
      "令和9年4月1日に義務となる2件（協力医療機関の要件強化・"
      "生産性向上委員会）も条例の対応規定を経て及ぶ",
      "No.95"],
     ["総合事業の令和7年度実績", "発注者・3町",
      "地域支援事業の量の見込み（同項第2号）。\n"
      "現在は令和6年度の実績を3年据え置いている",
      "No.112"],
     ["認知症総合支援事業の実績", "発注者・3町",
      "認知症施策の量の見込み。決算の科目は4区分で6事業別の内訳がない",
      "No.150"]],
    [4.0, 3.0, 7.4, 2.4], first_bold=True)

H2("４　算定の精度・整合に効くもの")
CAP("精度と整合に効く資料")
TBL(["資料", "入手先", "何のために要るか", "確認事項"],
    [["給付費データ 令和7年度\n（受領済み）",
      "発注者",
      "令和8年9月25日に再提供を受け、3町計と広域連合単位が\n"
      "要介護度別まで円単位で一致した。町別の実績値を書ける",
      "No.87（解決）"],
     ["見える化システムの\n操作手引・認証後画面の写し", "発注者",
      "施設・居住系の画面から先へ進めない件、\n"
      "1人1月あたり給付費の年度の選択、認定者数の編集項目の有無",
      "No.123・No.137\n・No.138・No.148"],
     ["令和7年度の月報", "発注者",
      "見える化の画面の令和7年度が年報でなく月報によるものか。\n"
      "画面と年報の月平均が2％以上食い違う3件の説明がつく",
      "No.136"],
     ["事業所・施設の定員と\n稼働の実績", "発注者・3町",
      "地域密着型通所介護の稼働率、施設・居住系の定員の縮小の経緯。\n"
      "定員の縮小を趨勢と読むかどうかの判断",
      "No.107"],
     ["未利用認定者の内訳", "発注者",
      "認定者の24.5％（486人）がサービスを利用していない要因の分解",
      "No.23"],
     ["第9期計画本文", "発注者",
      "給付適正化計画（第6期）の記載の確認",
      "No.90"]],
    [4.0, 3.0, 7.4, 2.4], first_bold=True)

H2("５　受領済みで、現時点で不足のないもの")
for t in [
    "介護保険事業状況報告（年報）令和6年度・令和7年度。"
    "令和8年9月11日の再送により要介護度別の明細まで受領しました。"
    "見込量・給付費の算定はこれを基礎としています。",
    "介護保険事業状況報告（月報）令和8年4月〜7月分。",
    "見える化システムの総括表（令和6年度・令和7年度）。"
    "給付費の合計が年報と円単位で一致します。",
    "介護保険事業計画作成支援ツールの実績値。",
    "北海道の施設・住まいの名簿5種、介護保険事業所一覧（47区分）。",
    "介護保険特別会計の決算額、保険料収納率の実績、所得段階別被保険者数、"
    "介護給付費準備基金の残高（令和5年度末・令和6年度末）。",
]:
    BUL(t, size=10)

doc.add_page_break()

# ============================================================ 第2節
H1("第2節　資料が届かない場合の当方の扱い（既定値）")
P("第1次概算では、届かない資料を次のように置きました。"
  "全23件を別冊15シートに掲げています。"
  "「安全側」とあるものは、置き方を誤った場合に"
  "保険料が不足する側ではなく余る側に振れることを意味します。")
CAP("据え置き・既定値（%d件のうち主なもの）" % len(SUEOKI))
_ROWS = []
for s in SUEOKI[:12]:
    _ROWS.append([str(s[0]), str(s[3])[:60], str(s[4]), str(s[5])[:46]])
TBL(["項目", "本表での置き方", "確認事項", "月額への効き"],
    _ROWS, [3.4, 7.0, 2.2, 4.2], first_bold=True)
NOTE("全23件は別冊「サービス見込量の算定 第1次概算」15シートにあります。")
P("")
P("安全側に置いたものは次の3件です。", bold=True)
for t in ["施策による見込量の増減を織り込まない（自然体）",
          "給付適正化による抑制を織り込まない",
          "施設・居住系の実績の減少を将来へ延長しない"]:
    BUL(t, size=10)

doc.add_page_break()

# ============================================================ 第3節
H1("第3節　電話で照会するときの進め方")

H2("１　電話の前に手元に置くもの")
for t in ["照会する事項を書いた紙（1件につき1枚。本書 第5節の様式）",
          "突き合わせた数値そのもの（本書 第6節）",
          "こちらの希望期限と、届かない場合にどうするか（既定値）",
          "業務工程管理表 03_確認事項一覧の該当行（No.と件名）"]:
    BUL(t, size=10)

H2("２　話す順序")
CAP("電話の組み立て（5段階）")
TBL(["段階", "話すこと", "言い方のねらい"],
    [["① 名乗り", "会社名・氏名・何の業務で電話しているか",
      "相手が「どの案件か」をすぐに思い出せるようにする"],
     ["② 用件の一言",
      "「○○の件で、確認をお願いしたいことが1件あります」",
      "先に件数と所要時間の見当を伝える"],
     ["③ 事実の確認",
      "こちらで突き合わせた数値を、判断を交えずに読み上げる",
      "「誤っている」ではなく「食い違っている」と言う"],
     ["④ お願いすること",
      "何を・いつまでに・どの形式で",
      "相手の作業量が見えるようにする"],
     ["⑤ 届かない場合",
      "「いただけない場合は当方で○○として扱います」",
      "相手に判断の余地を残し、こちらの作業も止めない"]],
    [2.0, 7.0, 8.0], first_bold=True)

H2("３　言わないこと")
for t in ["「誤っています」「間違っています」。"
          "どちらのファイルが正しいかは当方では判定できません。",
          "「回答がありません」「督促です」。"
          "相手の側で止まっているとは限りません。"
          "「行き違いでしたら失礼します」と前置きします。",
          "他町の名前を出して比べること。"
          "本件は3町とも同じ向きに差が出ているため、"
          "特定の町の問題として話すと事実と違います。",
          "担当者の個人名・携帯番号をこちらの資料に書き取ること"
          "（発注者指示により保管しません）。"]:
    BUL(t, size=10)

doc.add_page_break()

# ============================================================ 第4節
H1("第4節　電話連絡の例")

H2("例1　給付費データ 令和7年度の相違（確認事項No.87。解決済み）")
P("宛先　〔　　　　　　　　　〕　電話　〔　　　　　　　　　〕", size=10)
P("※ 本件は令和8年9月25日に令和7年度のファイルの再提供を受けて"
  "解決しました。以下は、数値の食い違いを照会するときの"
  "言い方の例として残すものです。", size=9)
P("")

SERIF("受託者",
      "お世話になっております。"
      "大雪地区広域連合の第10期介護保険事業計画の策定支援を"
      "お請けしております、ビズアップ公共コンサルティングの"
      "〔　　　〕と申します。"
      "介護保険のご担当の方はいらっしゃいますでしょうか。")
SERIF("（取次ぎ）", "")
SERIF("受託者",
      "お忙しいところ恐れ入ります。"
      "8月28日にご提供いただきました給付費データの集計について、"
      "確認をお願いしたいことが1件ございます。"
      "5分ほどお時間をいただけますでしょうか。")
SERIF("先方", "（応答）")
SERIF("受託者",
      "ありがとうございます。"
      "令和7年度分の数値を突き合わせましたところ、"
      "3町それぞれのファイルの合計と、"
      "「【大雪】R7.給付費データ集計」の数値とが食い違っております。")
SERIF("受託者",
      "具体的には、介護度別給付費シートの総計で、"
      "3町の合計が%s円、大雪地区広域連合のファイルが%s円で、"
      "%s円、率にして%.1f％の開きがございます。"
      % (yen(SA["R7"][0][7]), yen(SA["R7"][1][7]), yen(SA_R7), SA_R7_P))
SERIF("受託者",
      "令和6年度と令和8年度は円単位で一致しておりまして、"
      "令和7年度だけが合わない状態です。")
SERIF("受託者",
      "どちらかが誤っているという趣旨ではございません。"
      "抽出の条件が年度によって違っているのではないかと考えておりまして、"
      "そこを教えていただきたいというお願いでございます。")
SERIF("先方", "（どういう違いか、という趣旨の応答）")
SERIF("受託者",
      "はい。差が要介護度によって%.1f％から%.1f％まで幅がございまして、"
      "1か月分が抜けたというような一様な差ではないようです。"
      "対象月で切っておられるのか、審査月で切っておられるのか、"
      "あるいは住所地特例の方の扱いが年度で違うのか、"
      "そのあたりではないかと見ております。"
      % (min(DO_P), max(DO_P)))
SERIF("受託者",
      "お願いしたいことは2つでございます。"
      "1つは、令和7年度分について、"
      "どの条件で抽出されたものかを教えていただくこと。"
      "もう1つは、条件が違っていた場合に、"
      "同じ条件で作り直したファイルを1つご提供いただくことです。")
SERIF("受託者",
      "作り直していただく場合、個人情報が含まれる1枚目のシートは不要です。"
      "集計のシートだけで足ります。"
      "介護度別給付費、1月あたり利用者、1人1月あたり給付費、"
      "1人1月あたり利用回数の4枚です。")
SERIF("先方", "（いつまでか、という趣旨の応答）")
SERIF("受託者",
      "サービス見込量の第2次概算を10月30日にお出しする予定ですので、"
      "10月中旬ごろまでにいただけますと、そこに反映できます。"
      "難しいようでしたらご相談ください。")
SERIF("受託者",
      "なお、いただけない場合でも算定は止まりません。"
      "保険料の算定には介護保険特別会計の決算額を用いておりますので、"
      "本件は計画書に載せる実績の記載にかかわるものでございます。"
      "その場合は、広域連合単位の値のみを用いて、"
      "町別の内訳は載せない扱いにいたします。")
SERIF("先方", "（了解の趣旨の応答）")
SERIF("受託者",
      "ありがとうございます。"
      "本日の内容を書面にしてお送りいたします。"
      "恐れ入りますが、ご担当の部署名をお教えいただけますでしょうか。"
      "送付先として記録させていただきます。")
SERIF("受託者",
      "お忙しいところありがとうございました。失礼いたします。")

P("")
NOTE("「【大雪】R7.給付費データ集計」と町別ファイルは、"
     "いずれも令和8年8月28日にご提供いただいたものです。"
     "個人情報を含む1枚目のシートは、発注者のご指示により"
     "作業終了後ただちに消去しており、当方には残っておりません。"
     "集計シートの値のみを当方の資料に収めています。")

H2("例2　相手が不在・折り返しになる場合")
SERIF("受託者",
      "承知いたしました。それでは改めてご連絡いたします。"
      "念のため用件だけお伝えしますと、"
      "いただきました資料の数値について、"
      "3町の合計と広域連合の数値が食い違っておりまして、"
      "抽出の条件を教えていただきたいという内容でございます。"
      "お戻りになりましたら、その旨お伝えいただけますと助かります。")
NOTE("折り返しを待つ場合でも、その日のうちに同じ内容を書面で送ります。"
     "口頭だけにしないことで、担当が替わっても内容が残ります。")

H2("例3　3町へ施設整備の方針をうかがう場合（確認事項No.88・No.84）")
SERIF("受託者",
      "第10期の計画には、日常生活圏域ごとの必要利用定員総数を"
      "載せることが介護保険法で定められております。"
      "第9期の計画には記載がございませんので、今回から新しく設けます。")
SERIF("受託者",
      "つきましては、令和9年度から11年度までの間に"
      "施設の整備や定員の変更をご予定されているかどうかを"
      "お聞かせいただけますでしょうか。"
      "ご予定がないということであれば、"
      "現在の定員を3年間据え置く案でお示しいたします。")
SERIF("受託者",
      "10月にご検討の予定とうかがっておりますので、"
      "ご検討の結果が出ましたらお知らせいただければ結構です。"
      "11月13日に予算編成用の見込量をお出しする予定ですので、"
      "そこに間に合えば反映できます。")

H2("例4　資料の提供を改めてお願いする場合（総合事業の利用者実人数）")
SERIF("受託者",
      "地域支援事業につきまして、計画に「量の見込み」を"
      "載せることが法律で定められております。"
      "9月25日にいただきました3町の実施報告書で、"
      "令和7年度にどの事業をどれだけ実施されたかは分かりました。"
      "ありがとうございました。")
SERIF("受託者",
      "ただ、実施報告書の人数は延べの人数でございまして、"
      "計画に載せる量は一人の方を1回として数えた実人数になります。"
      "国の総合事業の実施状況調査の様式に、"
      "事業ごとの利用者実人数の欄がございますので、"
      "令和7年度分をその様式のままいただけますと置き換えられます。")
SERIF("受託者",
      "10月中旬ごろまでにいただけますと第2次概算に反映できます。"
      "難しいようでしたら令和6年度の据え置きのままお出しし、"
      "その旨を注記いたします。")

doc.add_page_break()

# ============================================================ 第5節
H1("第5節　照会の記録様式")
P("電話で照会したときは、次の項目を書き取って残します。"
  "確認事項一覧の備考へ転記し、口頭のやり取りを記録に残します。")
CAP("電話照会の記録様式")
TBL(["項目", "記入欄"],
    [["日時", "令和　　年　　月　　日　　　時　　分（通話　　分）"],
     ["相手方", "〔団体名・部署名〕　　　　　　　　　　※個人名は記録しない"],
     ["確認事項No.", "No.　　　（件名　　　　　　　　　　　　　　　　）"],
     ["こちらから伝えたこと", ""],
     ["先方のご回答", ""],
     ["お願いしたこと", "内容　　　　　　　　　　　　　期限　　月　　日"],
     ["いただけない場合の扱い", ""],
     ["書面の送付", "送付日　　月　　日　　／　　宛先〔　　　　　　　〕"],
     ["次の行動", ""]],
    [4.0, 13.0], first_bold=True)
NOTE("発注者のご指示により、担当者の個人名・電話番号・"
     "メールアドレスは記録しません。部署名までとします。")

doc.add_page_break()

# ============================================================ 第6節
H1("第6節　確認事項No.87の経緯と決着")
P("令和7年度の給付費が3町のファイルの合計と広域連合のファイルで"
  "食い違っていた件の経緯です。"
  "同じ食い違いが他の資料で起きたときの当たり方としても残します。")

H2("１　年度別の突合（再提供後）")
CAP("給付費の総計（3町のファイルの合計と広域連合のファイル）")
TBL(["年度", "3町のファイルの合計", "【大雪】のファイル", "差", "率"],
    [["令和6年度", yen(SA["R6"][0][7]), yen(SA["R6"][1][7]),
      yen(SA["R6"][2][7]) + "円", "一致"],
     ["令和7年度", yen(SA["R7"][0][7]), yen(SA["R7"][1][7]),
      yen(SA["R7"][2][7]) + "円", "一致"],
     ["令和8年度\n（年度途中）", yen(SA["R8"][0][7]), yen(SA["R8"][1][7]),
      yen(SA["R8"][2][7]) + "円", "一致"]],
    [2.6, 4.4, 4.4, 3.6, 2.2], center={4}, first_bold=True)
NOTE("単位は円。令和7年度は令和8年9月25日に再提供を受けたファイルの"
     "「介護度別給付費」シートの総計です。"
     "令和6年度・令和8年度は令和8年8月28日受領分によります。")

H2("２　従前の表との差（令和7年度）")
CAP("従前の表の値と3町合計の差（要介護度別）")
TBL(["要介護度", "3町の合計", "従前の表の値", "差", "率"],
    [[DO[i], yen(SA["R7"][0][i]), yen(OLD[i]),
      "＋" + yen(SA_OLD[i]), "＋%.2f％" % DO_P[i]]
     for i in range(7)]
    + [["総計", yen(SA["R7"][0][7]), yen(OLD[7]),
        "＋" + yen(SA_R7), "＋%.2f％" % SA_R7_P]],
    [2.6, 4.4, 4.4, 3.6, 2.2], center={4}, first_bold=True)
NOTE("差は%.2f％から%.2f％まで幅があり、"
     "1か月分が欠けたような一様な差ではありませんでした。"
     "この幅は、表が明細と一致していなかったことの現れです。"
     % (min(DO_P), max(DO_P)))

H2("３　差の原因")
for t in ["従前のファイルも、その明細（個票）を集計すると"
          "3,213,054,558円であり、要介護度別も3町合計と円単位で一致する。"
          "表の側だけが明細と一致していなかった。",
          "したがって差は、3町と広域連合で対象月・審査月・住所地特例の"
          "扱いが違ったことによるものではない。",
          "再提供を受けたファイルの表は明細と一致しており、"
          "3町合計＝広域連合単位となった。",
          "従前お示ししていた「町別に按分して比べると3町とも上回る」"
          "「広域連合のファイルだけ伸びが趨勢から外れている」という整理は、"
          "表の値を広域連合の実績として扱っていたことによるものであり、"
          "取り下げる。"]:
    BUL(t, size=10)

H2("４　残っている手掛かり")
P("美瑛町の令和7年度のファイルは、"
  "1人1月あたり利用回（日）数のシートの名称が"
  "「%s」となっております。"
  "収録されている値は東神楽町のファイルと一致せず美瑛町の値でしたので、"
  "当方ではシート名の書き誤りとして扱っております。"
  "給付費の差が解消した後も、この点は書き誤りとして残ります。"
  % MISMATCH[0][2])

H2("５　同じ食い違いが起きたときの当たり方")
for t in ["表（ピボット）の値と、同じファイルの明細を集計した額とを"
          "先に突き合わせる。表が明細と一致していなければ、"
          "抽出条件を照会する前に表の更新をお願いすればよい。",
          "電話では「どちらかが誤っている」と言わず、"
          "「どの条件で抽出したものか」を伺う形にする。"
          "他町の名称を出して比べない。",
          "保険料の算定には介護保険特別会計の決算額を用いています。"
          "給付費データの食い違いは計画書に載せる実績の記載に"
          "かかわるものです。",
          "サービス見込量・給付費・保険料は、"
          "介護保険事業状況報告（年報）の広域連合単位の値により算定しています"
          "（令和7年度の給付費2,915,307,125円。"
          "見える化システムの総括表と円単位で一致します）。"]:
    BUL(t, size=10)

doc.add_page_break()

# ============================================================ 第7節
H1("第7節　自己点検")
CAP("本書の数値の点検")
TBL(["No.", "点検の内容", "結果", "判定"],
    [[str(n), a, b, "適合" if ok else "不適合"]
     for n, a, b, ok in CHECKS],
    [1.4, 6.0, 7.6, 2.0], center={0, 3}, first_bold=True)
NOTE("本書の数値は、給付費データ集計の集計値、"
     "サービス見込量の算定 第1次概算、"
     "及び業務工程管理表の確認事項一覧から読み込んで作成しています。"
     "固定値で書き写していないため、"
     "算定を改めれば本書の数値も追随します。")


os.makedirs(os.path.dirname(OUT), exist_ok=True)
# OOXML の順序と w:zoom の必須属性を直す（CLAUDE.md §4）
docx_fix.fix(doc)
doc.save(OUT)

_ng = [c for c in CHECKS if not c[3]]
print("saved:", OUT)
print("段落 %d / 表 %d" % (len(doc.paragraphs), len(doc.tables)))
print("令和7年度 3町計と広域連合の差 %d円（再提供後）" % SA["R7"][2][7])
print("従前の表との差 %s円（%.2f％）" % (yen(SA_R7), SA_R7_P))
print("要介護度別の差 %.2f％〜%.2f％" % (min(DO_P), max(DO_P)))
print("R6→R7の伸び 大雪%+.2f％／3町%+.2f〜%+.2f％"
      % (NOBI["大雪"] * 100, min(NOBI[m] for m in TOWNS) * 100,
         max(NOBI[m] for m in TOWNS) * 100))
print("据え置き %d件／確認事項 %d件" % (len(SUEOKI), len(KAKUNIN)))
print("自己点検 %d件：適合%d件・不適合%d件"
      % (len(CHECKS), len(CHECKS) - len(_ng), len(_ng)))
if _ng:
    for c in _ng:
        print("  不適合:", c[0], c[1], c[2])
    sys.exit(1)
print("すべての点検に適合しました。")
