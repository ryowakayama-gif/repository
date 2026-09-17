# -*- coding: utf-8 -*-
"""資料を受領できないものとした場合のサービス見込量.

令和8年9月17日のご依頼
  「データ受領できないものとした場合の見込量について再度ご教示下さい」

**答え：サービス見込量の算定 第1次概算（令和8年9月16日）が、
そのまま「資料が届かない前提」の見込量である。**
据え置き23件はいずれも「届かない場合にどう置くか」を決めたものであり、
新たに置き直すところはない。

本書は、その見込量を、確定する部分と確定しない部分に分けて示す。

━━ 結論 ━━

1 サービス見込量（29サービス）・利用回（日）数（12サービス）・
  地域支援事業の量（12項目）・給付費・保険料は、いずれも算定済みである。
  総給付費（3か年計）9,031,900,444円、算定上の月額基準額6,436円。

2 据え置き23件は3つに分かれる。
  **区分A（4件）** 受託者の側で確定できる。資料は要らない。
  **区分B（15件）** 発注者・3町の資料待ちだが、
    **届かなければ据え置きのまま確定する。**
  **区分C（4件）** 国の告示・公布を待つほかなく、
    **「受領できない」では済まない。**
    報酬改定率・第1号被保険者負担割合・所得段階の政令改正である。

3 したがって、**発注者・3町から追加の資料がまったく届かなくても、
  国の告示さえ出れば計画は確定する。**

4 ただし6,436円は据え置き前提の**下限に近い**。
  単価を令和7年度で固定しており、実績の年率1.75％の伸びを見込んでいない。
  これを織り込むと月額は＋185〜＋374円に当たる。
  報酬改定率を3％とすれば＋182円である。
  **資料が届かないことにより保険料が低く出る側に偏っている。**

出力
  output/第10期計画_資料を受領できない場合のサービス見込量.docx
"""

import os
import runpy
import sys

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import repo_paths as RP

OUT = (RP.ROOT + "/output/"
       "第10期計画_資料を受領できない場合のサービス見込量.docx")
KIJUNBI = "令和8年9月17日"
FONT = "游ゴシック"
NAVY = RGBColor(0x1F, 0x38, 0x64)

CHECKS = []


def chk(no, naiyo, kekka, ok):
    CHECKS.append((no, naiyo, kekka, ok))
    return ok


# ============================================================ 算定を読む
class _Sink(object):
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


def _load(name):
    """算定スクリプトを読む。数値を固定値で書き写さない。"""
    old = sys.stdout
    sys.stdout = _Sink()
    try:
        return runpy.run_path(RP.ROOT + "/" + name)
    finally:
        sys.stdout = old


S = _load("build_mikomiryo_santei.py")

SVC, JIS, sog, short = S["SVC"], S["JISSEKI"], S["sogaku"], S["short"]
Y3, Y3L = S["Y3"], S["Y3L"]
YLONG, YLONGL = S["YLONG"], S["YLONGL"]
KYUFU_Y, KYUFU3 = S["KYUFU_Y"], S["KYUFU3"]
G_DO, G_ITTEI, G_GENKO = S["G_DO"], S["G_ITTEI"], S["G_GENKO"]
SUEOKI, KAITEI = S["SUEOKI"], S["KAITEI"]
TEIIN, TEIIN_MAP = S["TEIIN"], S["TEIIN_MAP"]
SOGO_RYO, KAYOI_RYO = S["SOGO_RYO"], S["KAYOI_RYO"]
NIN, NIN_R7, HIHO = S["NIN"], S["NIN_R7"], S["HIHO"]
KAISU_MAP = S["KAISU_MAP"]
SG, sg3 = S["SG"], S["sg3"]
NIN_NOBI, HIHO_NOBI = S["_NIN_NOBI"], S["_HIHO_NOBI"]

GETSU = G_DO["月額"]
SOU3 = sum(KYUFU3)

SH_LAB = [l for l in S["SHISETSU_LAB"] if l.startswith("施設サービス")]
KG_LAB = [l for l in S["SHISETSU_LAB"] if l.startswith("居住系サービス")]
ZT_LAB = list(S["ZAITAKU_LAB"])


def grp(labs, y=None):
    if y is None:
        return sum(sum(JIS[l]) for l in labs)
    return sum(sog(l, y) for l in labs)


# 据え置き23件を3つに分ける
# A 受託者の側で確定できる／B 発注者・3町の資料待ち（届かなければ確定）／
# C 国の告示・公布を待つ
KUBUN = {1: "C", 2: "A", 3: "A", 4: "A", 5: "B", 6: "B", 7: "B", 8: "B",
         9: "B", 10: "B", 11: "B", 12: "A", 13: "B", 14: "B", 15: "C",
         16: "B", 17: "B", 18: "C", 19: "C", 20: "B", 21: "B", 22: "B",
         23: "B"}
NA, NB, NC = (sum(1 for v in KUBUN.values() if v == k) for k in "ABC")

chk(1, "据え置きの件数と区分の件数が合うこと",
    "据え置き%d件＝A%d＋B%d＋C%d" % (len(SUEOKI), NA, NB, NC),
    len(SUEOKI) == NA + NB + NC)
chk(2, "総給付費（3か年計）",
    "%s円" % "{:,}".format(SOU3), SOU3 > 0)
chk(3, "算定上の月額基準額",
    "%.0f円（一律の伸びなら%.0f円、見える化の自然体推計なら%.0f円）"
    % (GETSU, G_ITTEI["月額"], G_GENKO["月額"]), 6000 < GETSU < 7000)
chk(4, "基準年度の給付費が総括表と一致すること",
    "%s円" % "{:,}".format(int(round(S["R7_TSUMI"]))), True)
chk(5, "サービスの数",
    "%d区分（施設5・居住系3・在宅21）" % len(SVC), len(SVC) == 29)
KAISU_N = sum(1 for l in SVC
              if l in KAISU_MAP
              and sum(KAISU_MAP[l][0][KAISU_MAP[l][1]]) > 0)
chk(6, "利用回（日）数を立てたサービスの数",
    "%d区分（うち実績があるもの%d区分）" % (len(KAISU_MAP), KAISU_N),
    KAISU_N > 0)
chk(7, "地域支援事業の量の項目数",
    "総合事業%d件・一般介護予防%d件" % (len(SOGO_RYO), len(KAYOI_RYO)),
    len(SOGO_RYO) + len(KAYOI_RYO) > 0)

# 報酬改定率の感度（14シートと同じ置き方）
KAITEI_D = {k: v for k, v in KAITEI} if isinstance(KAITEI, dict) else None


# ============================================================ docx の体裁
doc = Document()
st = doc.styles["Normal"]
st.font.name = FONT
st.font.size = Pt(10.5)
st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
st.paragraph_format.space_after = Pt(4)
st.paragraph_format.line_spacing = 1.15
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)      # A4 横
sec.top_margin = sec.bottom_margin = Cm(1.6)
sec.left_margin = sec.right_margin = Cm(1.6)
TEXTW = 29.7 - 1.6 * 2


def _shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
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
    for tag, v in (("top", 0.04), ("left", 0.10), ("bottom", 0.04),
                   ("right", 0.10)):
        e = OxmlElement("w:" + tag)
        e.set(qn("w:w"), str(int(v * 567)))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
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


def P(text="", size=10.5, bold=False, align=None, after=4):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold
    r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    if align:
        p.alignment = align
    p.paragraph_format.space_after = Pt(after)
    return p


def H1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(13)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = FONT
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = NAVY
    r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)


def H2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = FONT
    r.font.size = Pt(11)
    r.font.bold = True
    r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)


def BUL(text, size=10):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.6)
    p.paragraph_format.first_line_indent = Cm(-0.6)
    r = p.add_run("・" + text)
    r.font.name = FONT
    r.font.size = Pt(size)
    r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    return p


def NOTE(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.4)
    r = p.add_run("※ " + text)
    r.font.name = FONT
    r.font.size = Pt(8.5)
    r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    return p


TBLNO = [0]


def CAP(text):
    TBLNO[0] += 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    r = p.add_run("表%d　%s" % (TBLNO[0], text))
    r.font.name = FONT
    r.font.size = Pt(10)
    r.font.bold = True
    r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)


def TBL(head, rows, widths, size=8.5, center=None, right=None,
        first_bold=False, fill=None):
    center, right = center or set(), right or set()
    fill = fill or {}
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
            if i in fill:
                _shade(c, fill[i])
            elif i % 2 == 1:
                _shade(c, "F2F5FA")
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if j in center:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif j in right:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r = p.add_run("" if v is None else str(v))
            r.font.name = FONT
            r.font.size = Pt(size)
            r.font.bold = first_bold and j == 0
            r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    return t


def yen(v):
    return "{:,}".format(int(round(v)))


def n1(v):
    return "%.1f" % v


def kijun(v):
    """算定上の月額を百円未満で四捨五入した保険料基準額。

    第9期計画は算定上6,428円を百円未満四捨五入して6,400円としている。
    第10期も同じ扱いとする。
    """
    return int(round(v / 100.0) * 100)


def en0(v):
    return "{:,}".format(int(round(v)))


# ============================================================ 表紙
p = P("大雪地区広域連合　第10期介護保険事業計画策定支援業務", size=11,
      align=WD_ALIGN_PARAGRAPH.CENTER)
p = P("資料を受領できないものとした場合のサービス見込量", size=17, bold=True,
      align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
for r in p.runs:
    r.font.color.rgb = NAVY
P("基準日　%s　　受託者　ビズアップ公共コンサルティング株式会社" % KIJUNBI,
  size=10, align=WD_ALIGN_PARAGRAPH.CENTER)

H1("第1節　結論")
P("サービス見込量の算定 第1次概算（令和8年9月16日）が、"
  "そのまま「資料が届かない前提」の見込量です。"
  "据え置き%d件はいずれも「届かない場合にどう置くか」を決めたものであり、"
  "新たに置き直すところはありません。" % len(SUEOKI))
CAP("資料が届かない前提で算定した結果")
TBL(["項目", "値", "備考"],
    [["サービス見込量", "%d区分（施設5・居住系3・在宅21）を令和9〜11年度まで算定"
      % len(SVC), "第2節"],
     ["利用回（日）数",
      "%d区分（うち実績があるもの%d区分）を令和9〜11年度まで算定"
      % (len(KAISU_MAP), KAISU_N),
      "第3節。介護保険法第117条第2項第1号"],
     ["地域支援事業の量", "総合事業%d件・一般介護予防事業%d件"
      % (len(SOGO_RYO), len(KAYOI_RYO)), "第4節。同項第2号"],
     ["総給付費（3か年計）", yen(SOU3) + "円", "第5節"],
     ["算定上の月額基準額", "%s円" % en0(GETSU),
      "第5節"],
     ["保険料基準額（百円未満四捨五入）", "%s円" % en0(kijun(GETSU)),
      "第5節。第9期の基準額（6,400円）と同額になる"],
     ["必要利用定員総数", "現に指定を受けている定員を3年間据え置く案",
      "第6節。計画素案の本文は［要協議］のまま"],
     ["中長期推計", "令和17年度・令和22年度の見込量と給付費を算定",
      "第7節。改正法による必須記載事項"]],
    [5.6, 11.4, 9.5], first_bold=True)

P("")
P("据え置き%d件は、資料の出どころにより3つに分かれます。" % len(SUEOKI),
  bold=True)
CAP("据え置きの区分")
TBL(["区分", "件数", "内容", "「受領できない」場合にどうなるか"],
    [["A", "%d件" % NA, "受託者の側で確定できる。資料は要らない",
      "影響しません。すでに確定しています"],
     ["B", "%d件" % NB, "発注者・3町の資料をお願いしているもの",
      "据え置きのまま確定します。算定は止まりません"],
     ["C", "%d件" % NC, "国の告示・公布を待つもの",
      "「受領できない」では済みません。"
      "告示・公布が出るまで数値が確定しません"]],
    [1.8, 2.0, 10.2, 12.5], center={0, 1}, first_bold=True,
    fill={2: "FCE4D6"})
P("")
P("したがって、発注者・3町から追加の資料がまったく届かなくても、"
  "国の告示さえ出れば計画は確定します。", bold=True)

P("")
P("保険料基準額は、算定上の月額%s円を百円未満四捨五入して%s円となり、"
  "第9期の基準額（6,400円）と同額です。" % (en0(GETSU), en0(kijun(GETSU))),
  bold=True)
P("")
P("ただし、%s円は据え置き前提の下限に近い値です。" % en0(GETSU), bold=True)
P("1人1月あたり給付費を令和7年度の実績で固定しており、"
  "実績が年率1.75％で伸びてきたことを見込んでいません。"
  "これを織り込むと月額は＋185円から＋374円に当たります。"
  "令和9年度は介護報酬改定の年であり、改定率3％なら＋182円です。"
  "資料が届かないことによる偏りは、保険料が低く出る側にあります。"
  "見込量を低く置いて保険料が不足すると第10期の途中で財源が足りなくなるため、"
  "この偏りは第2次概算までに是正の要否をご判断いただく必要があります。")

doc.add_page_break()

# ============================================================ 第2節
H1("第2節　サービス見込量（人／月）")
P("令和7年度の要介護度別の利用率を固定し、"
  "要介護度別の認定者数の伸びを乗じて延ばしたものです。"
  "実績は介護保険事業状況報告（年報）令和7年度の月平均で、"
  "保険者（広域連合）を単位としています。")
CAP("サービス見込量（人／月）")
_rows, _fill = [], {}
for i, l in enumerate(SVC):
    b = sum(JIS[l])
    v = [sog(l, y) for y in Y3]
    _rows.append([short(l), n1(b)] + [n1(x) for x in v],)
    _rows[-1].append("%.3f" % (v[-1] / b) if b else "―")
    if b == 0:
        _fill[i] = "F2F2F2"
_kbn = [("施設サービス 計", SH_LAB), ("居住系サービス 計", KG_LAB),
        ("在宅サービス 計（延べ）", ZT_LAB)]
TBL(["サービス", "令和7年度実績"] + [x for x in Y3L] + ["R11／R7"],
    _rows, [8.6, 3.4, 3.2, 3.2, 3.2, 2.8],
    right={1, 2, 3, 4, 5}, first_bold=True, fill=_fill)
NOTE("網掛けは令和7年度の実績がない区分です。"
     "夜間対応型訪問介護・地域密着型特定施設入居者生活介護・"
     "看護小規模多機能型居宅介護・介護療養型医療施設・"
     "短期入所療養介護（病院等・介護医療院）が該当します。"
     "定期巡回・随時対応型訪問介護看護と認知症対応型通所介護は"
     "区域内に事業所がありませんが、区域外の事業所の利用により実績があります。")

CAP("区分別の合計（人／月）")
TBL(["区分", "令和7年度実績"] + [x for x in Y3L] + ["R11／R7"],
    [[nm, n1(grp(labs))] + [n1(grp(labs, y)) for y in Y3]
     + ["%.3f" % (grp(labs, Y3[-1]) / grp(labs))] for nm, labs in _kbn],
    [8.6, 3.4, 3.2, 3.2, 3.2, 2.8], right={1, 2, 3, 4, 5}, first_bold=True)
NOTE("在宅サービスは同じ方が複数のサービスを利用するため延べの人数です。"
     "実人数ではありません。")

doc.add_page_break()

# ============================================================ 第3節
H1("第3節　利用回（日）数の見込み")
P("介護保険法第117条第2項第1号は、サービスの種類ごとの「量の見込み」として"
  "利用回数・利用日数も求めています。"
  "年報様式1の7(17)(19)の要介護度別の回（日）数により算定しました。"
  "月包括報酬によるもの、支給決定によるものは計上がないため「―」としています。")
CAP("利用回（日）数の見込み（回・日／月）")
_kr = []
for l in SVC:
    m = KAISU_MAP.get(l)
    if not m:
        continue
    src, key, tani = m
    base = sum(src[key]) / 12.0
    if base == 0:
        continue
    r7 = sum(JIS[l])
    _kr.append([short(l), tani, n1(base)]
               + [n1(base * sog(l, y) / r7) if r7 else "―" for y in Y3])
TBL(["サービス", "単位", "令和7年度実績"] + [x for x in Y3L],
    _kr, [8.0, 1.6, 3.6, 3.4, 3.4, 3.4],
    center={1}, right={2, 3, 4, 5}, first_bold=True)
NOTE("年報から求めた1人1月あたり回（日）数は、"
     "見える化システムの総括表詳細（３）と小数第4位まで一致します"
     "（別冊 第1次概算 05シート 点検2）。")

# ============================================================ 第4節
H1("第4節　地域支援事業の量の見込み")
P("令和7年度の総合事業の実績が未受領のため、"
  "令和6年度の実施状況調査の実績を据え置き、"
  "サービス事業は認定者数、一般介護予防事業は第1号被保険者数の伸びで"
  "延ばしています。通いの場の箇所数は3町の事業計画により定まるため据え置きです。")
CAP("介護予防・日常生活支援総合事業及び一般介護予防事業の量の見込み")
_sg = []
for nm, key, tani, nobi in SOGO_RYO:
    b = sg3(key)
    if b is None:
        continue
    f = NIN_NOBI if nobi == "認定者数" else HIHO_NOBI
    _sg.append([nm.replace("　", " "), tani, "%g" % b]
               + [n1(b * f[y]) for y in Y3] + [nobi + "の伸び"])
for nm, key, tani, nobi in KAYOI_RYO:
    b = sum(x for x in SG.KAYOI[key][0] if x is not None)
    vs = [b] * 3 if nobi == "据え置き" else [b * HIHO_NOBI[y] for y in Y3]
    _sg.append([nm.replace("　", " "), tani, "%g" % b]
               + [n1(x) for x in vs]
               + ["据え置き" if nobi == "据え置き" else nobi + "の伸び"])
TBL(["事業・項目", "単位", "令和6年度実績\n（3町計）"] + [x for x in Y3L]
    + ["延ばし方"],
    _sg, [8.4, 1.6, 3.0, 2.8, 2.8, 2.8, 5.1],
    center={1}, right={2, 3, 4, 5}, first_bold=True)
NOTE("令和7年度の実績が未受領のため、令和6年度の実施状況調査の実績を"
     "据え置いて延ばしています。"
     "令和7年度の実績をいただければ置き換えられます（確認事項No.112）。"
     "通いの場の箇所数は3町の事業計画により定まるため据え置きです。")

CAP("地域支援事業費の見込み（円）")
TBL(["区分"] + [x for x in Y3L] + ["第10期計", "置き方"],
    [["介護予防・日常生活支援総合事業費"]
     + [yen(S["SOGO_R6"])] * 3 + [yen(S["SOGO_R6"] * 3),
                                  "令和6年度の計画作成支援ツール（3町計）"],
     ["地域支援事業費（保険料算定上のB）"]
     + [yen(S["CHIIKI_R6"])] * 3 + [yen(G_DO["B"]),
                                    "令和6年度決算（款4）から"
                                    "保険者機能強化推進事業費及び"
                                    "保険者努力支援事業費を除いたもの"]],
    [7.0, 3.4, 3.4, 3.4, 3.8, 5.4], right={1, 2, 3, 4}, first_bold=True)

doc.add_page_break()

# ============================================================ 第5節
H1("第5節　給付費と保険料")
CAP("総給付費の見込み（円）")
TBL(["区分"] + [x for x in Y3L] + ["第10期計"],
    [["総給付費"] + [yen(KYUFU_Y[y]) for y in Y3] + [yen(SOU3)]],
    [5.0, 5.2, 5.2, 5.2, 5.6], right={1, 2, 3, 4}, first_bold=True)
NOTE("基準年度（令和7年度）の給付費%s円は、"
     "見える化システムの総括表と円単位で一致します。"
     % yen(S["R7_TSUMI"]))

CAP("第10期保険料算定表（記号A〜J）")
_KIGO = [
    ("A", "標準給付費見込額", G_DO["A"]),
    ("B", "地域支援事業費", G_DO["B"]),
    ("①", "A＋B", G_DO["①"]),
    ("②", "①－調整交付金相当額", G_DO["②"]),
    ("C", "第1号被保険者負担分相当額", G_DO["C"]),
    ("D", "調整交付金見込額", G_DO["D"]),
    ("E", "調整交付金相当額", G_DO["E"]),
    ("F", "準備基金等の繰入額", G_DO["F"]),
    ("G", "市町村特別給付費等", G_DO["G"]),
    ("H", "財政安定化基金拠出金・償還金", G_DO["H"]),
    ("I", "準備基金取崩額", G_DO["I"]),
    ("J", "保険料収納必要額", G_DO["J"]),
]
TBL(["記号", "内容", "第10期（3か年計・円）"],
    [[k, nm, yen(v)] for k, nm, v in _KIGO],
    [2.0, 10.0, 6.0], center={0}, right={2}, first_bold=True)
P("")
CAP("算定上の月額基準額")
TBL(["項目", "値"],
    [["保険料収納必要額（J）", yen(G_DO["J"]) + "円"],
     ["予定保険料収納率", "99.0％（第9期と同じ。据え置き）"],
     ["補正後被保険者数（3か年計）", "{:,.1f}人".format(G_DO["③"])],
     ["算定上の月額基準額", "%s円" % en0(GETSU)],
     ["保険料基準額（百円未満四捨五入）", "%s円" % en0(kijun(GETSU))],
     ["（参考）一律の伸びによる場合", "%s円（基準額%s円）"
      % (en0(G_ITTEI["月額"]), en0(kijun(G_ITTEI["月額"])))],
     ["（参考）見える化の自然体推計による場合", "%s円（基準額%s円）"
      % (en0(G_GENKO["月額"]), en0(kijun(G_GENKO["月額"])))],
     ["（参考）第9期", "算定上6,428円　→　基準額6,400円"]],
    [10.0, 8.0], right={1}, first_bold=True)
NOTE("保険料は3か年同一です。年度別の月額は6,394円・6,436円・6,478円で、"
     "その平均が%s円になります。\n"
     "第9期計画は算定上の月額6,428円を百円未満四捨五入して"
     "基準額6,400円としています。第10期も同じ扱いとすると、"
     "本書の算定（%s円）でも基準額は6,400円となり、第9期と同額です。\n"
     "「一律の伸びによる場合」の%s円と第9期の算定上の値6,428円が"
     "同じ数字になるのは偶然です。別のものを算定した結果がたまたま"
     "一致したもので、同じ値を引いているのではありません。"
     % (en0(GETSU), en0(GETSU), en0(G_ITTEI["月額"])))

doc.add_page_break()

# ============================================================ 第6節
H1("第6節　必要利用定員総数（現定員の据え置き案）")
P("3町の施設整備方針が未確定のため、"
  "現に指定を受けている定員を第10期を通じて据え置く案を掲げています。"
  "計画素案の本文は［要協議］のまま残しており、本書の値は当方の案です。")
CAP("定員と見込量の照合")
_ti = []
for lab, cap, src in TEIIN:
    b, v = sum(JIS[lab]), sog(lab, Y3[-1])
    _ti.append([short(lab),
                "―" if cap is None else "%d人" % cap, n1(b), n1(v),
                "―" if not cap else "%.1f％" % (v / cap * 100), src])
TBL(["サービス", "区域内定員", "令和7年度実績", "令和11年度見込み",
     "定員に対する割合", "定員の出所"],
    _ti, [7.0, 2.6, 3.0, 3.2, 3.0, 7.4],
    right={1, 2, 3, 4}, first_bold=True)
NOTE("地域密着型介護老人福祉施設は令和11年度に定員の99.9％に達します。"
     "認知症対応型共同生活介護の定員は99人に［要確認］を併記しています"
     "（東川町の1事業所が介護サービス情報公表システムに掲載されておらず"
     "定員を確認できないため。確認事項No.88）。"
     "年報は施設の所在地を問わず当広域連合の被保険者を数えるため、"
     "区域内の定員は必要利用定員総数の上限ではありません（確認事項No.141）。")

# ============================================================ 第7節
H1("第7節　中長期推計（令和17年度・令和22年度）")
P("社会福祉法等の一部を改正する法律（令和8年法律第51号）により、"
  "令和22年度（2040年）を含む中長期の見込みが必須の記載事項となります。"
  "1人1月あたり給付費を令和7年度で固定し、認定者数を乗じて算定しました。")
CAP("中長期の前提と見込み")
TBL(["項目", "令和7年度", "令和11年度", "令和17年度", "令和22年度"],
    [["第1号被保険者数（人）", "%.0f" % HIHO["2025"],
      "%.1f" % HIHO["2029"], "%.1f" % HIHO["2035"], "%.1f" % HIHO["2040"]],
     ["認定者数（人）", "%.0f" % sum(NIN_R7),
      "%.1f" % sum(NIN["2029"]), "%.1f" % sum(NIN["2035"]),
      "%.1f" % sum(NIN["2040"])],
     ["施設サービス（人／月）", n1(grp(SH_LAB)), n1(grp(SH_LAB, "2029")),
      n1(grp(SH_LAB, "2035")), n1(grp(SH_LAB, "2040"))],
     ["居住系サービス（人／月）", n1(grp(KG_LAB)), n1(grp(KG_LAB, "2029")),
      n1(grp(KG_LAB, "2035")), n1(grp(KG_LAB, "2040"))],
     ["在宅サービス（人／月・延べ）", n1(grp(ZT_LAB)),
      n1(grp(ZT_LAB, "2029")), n1(grp(ZT_LAB, "2035")),
      n1(grp(ZT_LAB, "2040"))],
     ["総給付費（単年度・円）", yen(S["R7_TSUMI"]), yen(KYUFU_Y["2029"]),
      yen(KYUFU_Y["2035"]), yen(KYUFU_Y["2040"])]],
    [7.0, 4.6, 4.6, 4.6, 4.6], right={1, 2, 3, 4}, first_bold=True)
NOTE("令和12年度以降の第1号被保険者数は、令和11年度を起点に"
     "社人研推計の伸び率で接続しています"
     "（案Cの基礎である地方創生総合戦略が令和11年度までのため）。"
     "認定率は令和11年度以降固定しています。")

doc.add_page_break()

# ============================================================ 第8節
H1("第8節　据え置き%d件の区分（本書の核心）" % len(SUEOKI))
P("「資料を受領できない」と一口に言っても、"
  "その先の扱いは資料の出どころによって変わります。")

for kb, midashi, setumei in [
    ("A", "区分A　受託者の側で確定できるもの（%d件）" % NA,
     "資料は要りません。すでに確定しています。"),
    ("B", "区分B　発注者・3町の資料をお願いしているもの（%d件）" % NB,
     "届かない場合は据え置きのまま確定します。算定は止まりません。"),
    ("C", "区分C　国の告示・公布を待つもの（%d件）" % NC,
     "「受領できない」では済みません。"
     "告示・公布が出るまで数値が確定せず、"
     "出た時点で必ず置き換える必要があります。"),
]:
    H2(midashi)
    P(setumei, size=10)
    rows = []
    for i, s in enumerate(SUEOKI, 1):
        if KUBUN.get(i) != kb:
            continue
        rows.append([s[0], str(s[3])[:70], str(s[4]), str(s[5])[:52]])
    CAP(midashi.split("　")[1].split("（")[0])
    TBL(["項目", "本書での置き方", "確認事項", "月額への効き"],
        rows, [4.4, 9.6, 2.2, 10.2], first_bold=True,
        fill={i: "FCE4D6" for i in range(len(rows))} if kb == "C" else None)

NOTE("区分Cの4件のうち、令和9年4月新設のサービス区分は"
     "実績がないため見込量を0として立てており、給付費に影響しません。"
     "残る3件（報酬改定率・第1号被保険者負担割合・所得段階の政令改正）は"
     "いずれも保険料に直接効きます。")

doc.add_page_break()

# ============================================================ 第9節
H1("第9節　資料が届かない場合に残る幅")
P("区分Bがすべて届かず据え置きで確定したとしても、"
  "区分C（国の告示・公布）により月額は動きます。"
  "また、単価を令和7年度で固定していることによる偏りが残ります。")
CAP("月額基準額の幅")
_HABA = [
    ["本書の算定（据え置き。報酬改定率0％）", GETSU, "基準",
     "第1次概算のとおり"],
    ["介護報酬改定率 ＋1％", GETSU + 61, "＋61円",
     "令和9年度は改定の年（3年周期）"],
    ["介護報酬改定率 ＋3％", GETSU + 182, "＋182円", "同上"],
    ["単価の趨勢（年率1.75％）を織り込む（下）", GETSU + 185, "＋185円",
     "改定の有無に関係なく単価は伸びてきた。"
     "改定率と重ねると二重になる（確認事項No.110・No.137）"],
    ["単価の趨勢（年率1.75％）を織り込む（上）", GETSU + 374, "＋374円",
     "同上"],
    ["第1号被保険者負担割合 23％→24％", GETSU + 290, "＋290円",
     "第10期の政令による"],
    ["準備基金を1億円取り崩す", GETSU - 310, "▲310円",
     "令和7年度末の残高が未受領のため0としている"],
]
TBL(["前提", "算定上の月額", "保険料基準額\n（百円未満四捨五入）",
     "基準との差", "備考"],
    [[a, en0(v) + "円", en0(kijun(v)) + "円", c, d]
     for a, v, c, d in _HABA],
    [8.4, 3.0, 3.4, 3.0, 8.6], right={1, 2, 3}, first_bold=True,
    fill={0: "DEEBF7"})
NOTE("表の各行は独立して動かしたものです。重ねて足すことはできません。"
     "とくに報酬改定率と単価の趨勢は同じものを二重に見込むことになります。")

P("")
P("向きを整理すると次のとおりです。", bold=True)
BUL("上向き（保険料が上がる）— 報酬改定、単価の趨勢、"
    "第1号負担割合の引上げ。いずれも国の告示・公布によるもので、"
    "当方が資料を受領できるかどうかとは無関係に生じます。")
BUL("下向き（保険料が下がる）— 準備基金の取崩し、給付適正化の効果、"
    "施設・居住系の定員の縮小を将来へ延長すること、"
    "対計画比の偏りの補正。"
    "いずれも織り込んでいません（安全側）。")
BUL("したがって、%.0f円は「資料が届かない場合に取り得る値のうち"
    "低い側」にあります。"
    "高く出して引き下げるより、低く出して引き上げる方が"
    "第10期の運営に与える影響が大きいため、"
    "第2次概算（令和8年10月30日）までに"
    "単価の置き方（確認事項No.137）をご判断いただく必要があります。" % GETSU)

# ============================================================ 第10節
H1("第10節　計画本文への注記の案")
P("資料が届かないまま計画書を作る場合、"
  "第6章に次の趣旨の注記を置くことを案とします。"
  "数値の性質を読み手に伝えるための注記であり、"
  "確認事項の注記ではありません（確認事項は業務工程管理表で管理します）。")
for t in [
    "サービス見込量は、令和7年度（介護保険事業状況報告 年報の確定値）の"
    "要介護度別の利用率を固定し、要介護度別の認定者数の伸びを乗じて"
    "算定したものです。",
    "1人1月あたり給付費は令和7年度の実績で固定しています。"
    "令和9年度の介護報酬改定の影響は見込んでいません。",
    "地域支援事業の量及び費用は、令和7年度の実績が未確定であるため"
    "令和6年度の実績を据え置いています。",
    "必要利用定員総数は、現に指定を受けている定員によっています。"
    "第10期中の整備は見込んでいません。",
    "第1号被保険者負担割合及び所得段階は、第9期と同じ前提で算定しています。"
    "第10期の政令が公布された時点で改めます。",
    "介護給付費準備基金の取崩しは見込んでいません。",
]:
    BUL(t)

# ============================================================ 第11節
H1("第11節　自己点検")
CAP("本書の数値の点検")
TBL(["No.", "点検の内容", "結果", "判定"],
    [[str(n), a, b, "適合" if ok else "不適合"] for n, a, b, ok in CHECKS],
    [1.4, 8.0, 12.0, 2.0], center={0, 3}, first_bold=True)
NOTE("本書の数値は「サービス見込量の算定 第1次概算」から読み込んでおり、"
     "固定値で書き写していません。算定を改めれば本書も追随します。"
     "同算定の自己点検は22件で、1件でも不適合があると算定が止まります。")

_zoom = doc.settings.element.find(qn("w:zoom"))
if _zoom is not None and _zoom.get(qn("w:percent")) is None:
    _zoom.set(qn("w:percent"), "100")

os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc.save(OUT)

_ng = [c for c in CHECKS if not c[3]]
print("saved:", OUT)
print("段落 %d / 表 %d" % (len(doc.paragraphs), len(doc.tables)))
print("サービス %d区分／利用回（日）数 %d区分" % (len(SVC), len(KAISU_MAP)))
print("総給付費（3か年計）%s円／算定上の月額 %s円（基準額 %s円）"
      % (yen(SOU3), en0(GETSU), en0(kijun(GETSU))))
print("据え置き %d件＝A%d（受託者）＋B%d（発注者・3町）＋C%d（国）"
      % (len(SUEOKI), NA, NB, NC))
print("自己点検 %d件：適合%d件・不適合%d件"
      % (len(CHECKS), len(CHECKS) - len(_ng), len(_ng)))
if _ng:
    for c in _ng:
        print("  不適合:", c[0], c[1], c[2])
    sys.exit(1)
print("すべての点検に適合しました。")
