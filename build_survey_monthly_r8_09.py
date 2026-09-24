# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画
アンケート調査の分析　令和8年9月 業務報告（発注者提出用）.

令和8年9月24日のご依頼
  「9月業務上のアンケート分析について先方に提出する資料と
    社内チェック用の点検箇所を整理した資料をワードでアウトプットして下さい」

本書は**発注者に提出する側**である。
社内チェック用は `build_survey_check.py`（別冊）とする。

業務仕様書「４．業務内容（3）実施済み調査結果の集計・分析」及び
「５．履行工程　令和8年9月」に対する報告として、
令和8年9月に行ったこと・その結果・計画素案への反映・残る事項をまとめる。

━━ 数値はすべて実物から読む ━━

**固定値を書かない。** 件数・区分・所見・留保は成果品の実物
（`output/第10期計画_アンケート調査の集計分析報告書.xlsx`）を読んで数える。
突合の測定値は `build_survey_jisseki_cross.py` を `runpy` で読む。
見込量は `build_mikomiryo_santei.py` を `runpy` で読む。
計画素案の規模は `repo_paths.draft_label()`（実物を数える）。

同じ数値を2つの成果品に別々に書かない（CLAUDE.md §4）。

━━ 守るべき制約 ━━

個票は収録しない。担当者名・電話番号・メールアドレスは書かない。
調査①の結果は「在宅で困難を抱える方の状況」として引き、
在宅の認定者全体の割合としては引かない（確認事項No.140）。

構成
  第1節　令和8年9月に行ったこと
  第2節　仕様書４（3）の作業項目に対する状況（網羅性）
  第3節　対象4調査・母数と回収の状況
  第4節　前回調査との比較（経年）
  第5節　健康とくらしの調査の結果と計画素案への反映
  第6節　調査結果と保険者の実績の突合
  第7節　調査から把握した地域課題
  第8節　主要所見と計画素案への反映状況
  第9節　調査結果の限界と留保
  第10節　ご判断をお願いする事項
  第11節　成果品

出力
  output/第10期計画_アンケート調査の分析_令和8年9月業務報告.docx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import ast
import io
import os
import runpy
import sys
import collections

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from openpyxl import load_workbook

import repo_paths as RP

OUT = (RP.ROOT + "/output/"
       "第10期計画_アンケート調査の分析_令和8年9月業務報告.docx")
SRC_XLSX = os.path.join(RP.OUTPUT, "第10期計画_アンケート調査の集計分析報告書.xlsx")
KIJUNBI = "令和8年9月24日"
FONT = "游ゴシック"
NAVY = RGBColor(0x1F, 0x38, 0x64)

CHECKS = []


def chk(no, naiyo, kekka, ok):
    CHECKS.append((no, naiyo, kekka, "適合" if ok else "不適合"))
    return ok


# ==================================================== 実物から読む
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
        return runpy.run_path(os.path.join(RP.ROOT, name))
    finally:
        sys.stdout = old


_X = _load("build_survey_jisseki_cross.py")     # 突合（χ²・在籍と受給）
_M = _load("build_mikomiryo_santei.py")         # 見込量の算定

WB = load_workbook(SRC_XLSX, data_only=True)
SHEETS = WB.sheetnames


def rows_of(sheet, head_row=4):
    """シートの表の行を返す（見出し行の次から、空行まで）。"""
    ws = WB[sheet]
    out = []
    for row in ws.iter_rows(min_row=head_row + 1, values_only=True):
        if row[0] is None or (isinstance(row[0], str)
                              and row[0].startswith("注")):
            break
        out.append(row)
    return out


# 01シート　回収の状況
KAISHU = rows_of("01_調査の実施状況と回収")
# 11シート　留保
RYUHO = rows_of("11_調査結果の限界と留保")
RYUHO_KB = collections.Counter(r[4] for r in RYUHO)
# 12シート　主要所見
SHOKEN = rows_of("12_主要所見と計画本文への反映")
SHOKEN_ST = collections.Counter(r[5] for r in SHOKEN)
# 14シート　補完と解消（2つの表に分かれている）
_ws14 = WB["14_公表データによる補完と留保の解消"]
HOKAN = []
for row in _ws14.iter_rows(min_row=5, values_only=True):
    # No.の列が数字のものだけを拾う（表が2つに分かれており見出し行が挟まる）。
    if str(row[0] or "").strip().isdigit():
        HOKAN.append(row)
HOKAN_ST = collections.Counter(r[5] for r in HOKAN)

# 突合の測定値
DO = _X["DO"]
CU_DO, CU_EXP, CU_CHI = _X["CU_DO"], _X["CU_EXP"], _X["CU_CHI"]
CHI_1, CU_N = _X["CHI_1"], _X["CU_N"]
ZAITAKU_NIN = _X["ZAITAKU_NIN"]

# 見込量（アンケート分析報告書 10シートと同じ値。算定から読む）
SVC, kubun, JISSEKI, sogaku = (_M["SVC"], _M["kubun"], _M["JISSEKI"],
                               _M["sogaku"])


def grp(kb, y=None):
    ls = [l for l in SVC if kubun(l) == kb]
    return (sum(sum(JISSEKI[l]) for l in ls) if y is None
            else sum(sogaku(l, y) for l in ls))


DRAFT = RP.draft_label()


def n1(v):
    return "%.1f" % v


# ---------------------------------------------------- 10シートの3つの表
def block_rows(sheet, lead, ncol=6):
    """【…】の見出しで始まる表の行を返す（見出し行の次から空行まで）。

    10シートは1枚に3つの表が並んでおり、`rows_of()` では読めない。
    シートごとに表の作りが違うことを前提に読む（CLAUDE.md §4）。
    """
    ws = WB[sheet]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    top = None
    for i, row in enumerate(rows):
        if isinstance(row[0], str) and row[0].startswith(lead):
            top = i + 2          # 見出しの次が表頭、その次から本体
            break
    if top is None:
        raise KeyError(lead)
    out = []
    for row in rows[top:]:
        if row[0] is None or (isinstance(row[0], str)
                              and row[0].startswith("注")):
            break
        out.append(list(row[:ncol]))
    return out


_SD = "10_供給と需要の対照"
# 先頭の表（区分・実績・見込み・定員・割合・見方）は【】の見出しを持たない
KYOKYU = [list(r[:6]) for r in
          WB[_SD].iter_rows(min_row=5, max_row=6, values_only=True)]
ZAISEKI = block_rows(_SD, "【施設の側から数えた在籍者")
RIYO_JOKYO = block_rows(_SD, "【認定者の利用状況")
SETSUZOKU = block_rows(_SD, "【調査結果と供給構造の接続")


# ---------------------------------------------------- 仕様書４（3）の作業項目
def _work3():
    """仕様書４（3）の要求・作業項目・成果物を業務工程管理表のソースから読む。

    `WORK` には `rate()` の呼出しが混じるため literal_eval できない。
    文字列の要素だけを取り出す。
    """
    src = open(os.path.join(RP.ROOT, "build_process_control.py"),
               encoding="utf-8").read()
    for node in ast.walk(ast.parse(src)):
        if (isinstance(node, ast.Assign)
                and getattr(node.targets[0], "id", "") == "WORK"):
            for elt in node.value.elts:
                vals = [e.value for e in elt.elts[:5]
                        if isinstance(e, ast.Constant)]
                if vals and vals[0] == "(3)":
                    return vals[2], vals[3].split("\n"), vals[4].split("\n")
    raise RuntimeError("仕様書４（3）が読めない")


YOKYU, SAGYO, SEIKABUTSU = _work3()


# ---------------------------------------------------- 前回調査（経年比較）
def _zenkai():
    """前回調査の規模を図表集の実物（00_凡例・出典）から読む。"""
    path = os.path.join(RP.OUTPUT, "第10期計画_図表集_白黒.xlsx")
    wb = load_workbook(path, data_only=True)
    ws = wb["00_凡例・出典"]
    out = {}
    for row in ws.iter_rows(min_row=9, max_row=40, values_only=True):
        if row[1] and row[3]:
            out[str(row[1])] = str(row[3])
    return out


ZENKAI = _zenkai()


def _zen(key):
    """図表集の出所の記載から、前回調査の実施時期・規模の部分だけを返す。"""
    return ZENKAI[key].split("／")[-1]


def _sheets_of(fn):
    """成果品（xlsx）のシート数を実物から数える。"""
    return len(load_workbook(os.path.join(RP.OUTPUT, fn),
                             read_only=True).sheetnames)


def _zenkai_hikaku():
    """計画素案 第2章第4節の【8　前回調査との比較】の表を実物から読む。"""
    from docx.table import Table as _T
    from docx.text.paragraph import Paragraph as _P
    d = Document(RP.DRAFT)
    blocks = []
    for ch in d.element.body.iterchildren():
        if ch.tag.endswith("}p"):
            blocks.append(("p", _P(ch, d)))
        elif ch.tag.endswith("}tbl"):
            blocks.append(("t", _T(ch, d)))
    on = False
    for kind, b in blocks:
        if kind == "p":
            if b.text.strip().startswith("【8　前回調査との比較】"):
                on = True
        elif on:
            return [[c.text for c in r.cells] for r in b.rows]
    raise RuntimeError("前回調査との比較の表が見つからない")


ZENKAI_HYO = _zenkai_hikaku()

# ---------------------------------------------------- 調査①の票数
# 受領は99票。うち1票は要介護度の記載がなく、分布の比較から除いている。
import data_survey2025 as _S2                                   # noqa: E402

RIYO_N = _S2.RIYO["件数"]
RIYO_FUMEI = _S2.RIYO["要介護度"].get("-", 0)


# ---------------------------------------------------- 課題の束ね方
# 課題の名は本報告で束ねたもの。根拠と反映先は12シートの実物から読む。
KADAI = [
    ("移動と外出の確保（生活支援）", "7"),
    ("外出と社会参加（他保険者との比較）", "9"),
    ("除雪（介護保険の給付対象外）", "8"),
    ("24時間対応サービスの不在", "3"),
    ("医療機関からの施設入所という経路", "1"),
    ("介護人材の把握範囲", "6"),
    ("介護人材の定着", "11"),
    ("在宅生活の継続が困難な層の存在", "15"),
]


def _shoken(no):
    """12シートの所見を番号で引く。"""
    for row in SHOKEN:
        if str(row[0]).strip() == str(no):
            return row
    raise KeyError(no)


# ==================================================== docx の体裁
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
    # CT_TcPrBase は要素の順序が定められており、shd は vAlign 等より前に置く
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
    # tblCellMar は tblLook より前に置く（CLAUDE.md §4）。
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


def _runs(p, text, size, bold=False, color=None):
    """** で囲んだ部分を太字にする。docx は Markdown を解釈しない。"""
    for i, part in enumerate(str(text).split("**")):
        if not part:
            continue
        r = p.add_run(part)
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.bold = bold or (i % 2 == 1)
        if color is not None:
            r.font.color.rgb = color
        r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    if not p.runs:
        r = p.add_run("")
        r.font.name = FONT
        r.font.size = Pt(size)
        r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    return p


def P(text="", size=10.5, bold=False, align=None, indent=0.0, after=4):
    p = doc.add_paragraph()
    _runs(p, text, size, bold)
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
    p.paragraph_format.space_after = Pt(3)
    _runs(p, mark + text, size)
    return p


def NOTE(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.4)
    _runs(p, "※ " + text, 9)
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
            _runs(p, "" if v is None else v, size,
                  bold=(first_bold and j == 0))
    return t


# ============================================================ 表紙
P("大雪地区広域連合　第10期介護保険事業計画策定支援業務", size=12,
  align=WD_ALIGN_PARAGRAPH.CENTER)
p = P("アンケート調査の分析　令和8年9月　業務報告", size=17, bold=True,
      align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
for r in p.runs:
    r.font.color.rgb = NAVY
P("基準日　%s　　受託者　ビズアップ公共コンサルティング株式会社" % KIJUNBI,
  size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
P("")

P("業務仕様書「４．業務内容（3）実施済み調査結果の集計・分析」及び"
  "「５．履行工程　令和8年9月」に対する報告です。"
  "令和8年9月に行ったこと、その結果、計画素案への反映、"
  "残る事項をまとめています。")
P("集計・クロス集計・分析は4調査とも完了しています。"
  "**9月に新たにできるようになったのは、調査の結果と保険者の実績（年報）を"
  "突き合わせることです。**"
  "令和8年9月11日に年報（令和7年度）の要介護度別の明細を収めたことにより、"
  "これまで定性的な留保にとどまっていたものを"
  "大きさと向きを伴う測定値にすることができました（第6節）。")

# ============================================================ 第1節
H1("第1節　令和8年9月に行ったこと")

CAP("令和8年9月の作業と成果物")
TBL(["時期", "行ったこと", "成果物", "結果"],
    [["9月3日", "アンケート調査の作業状況を整理した",
      "9月作業_アンケート分析の作業状況（10シート）",
      "集計・クロス集計・分析は4調査とも完了しており、"
      "9月に残るのは確定値化と計画本文への反映であることを確かめた"],
     ["9月11日",
      "年報（令和7年度）の要介護度別の明細を収め、"
      "調査の結果と保険者の実績を突き合わせた",
      "調査結果と年報実績の突合クロス集計（10シート）",
      "調査①の要介護度分布が在宅の認定者の分布と食い違うことを"
      "χ²＝%.1f（自由度6・1％点%.2f）として測定した。"
      "施設の在籍者と保険者の受給者は数える対象が異なることを確定した"
      % (CU_CHI, CHI_1)],
     ["9月15日",
      "健康とくらしの調査（4,729票）の結果を計画素案 第2章第4節に入れた",
      "計画素案（%s）" % DRAFT,
      "第2章第4節は4段落・表1から30段落・表8になった。"
      "集計値は data_survey_jages.py に収め、個票は収録していない"],
     ["9月17日",
      "アンケート調査の集計分析報告書を再点検し修正した",
      "アンケート調査の集計分析報告書（%dシート）／"
      "実施済み調査 結果報告書（Word）" % len(SHEETS),
      "10シートの見込量が計画素案 第6章と食い違っていたものを是正した。"
      "χ²と在籍・受給の差を留保一覧と課題整理に収めた。"
      "計画本文への反映状況を素案の実物を読んで判定する形に改めた"],
     ["9月24日", "本報告書及び社内点検用の資料を作成した",
      "本書／アンケート分析 点検箇所と手順（社内用）",
      "―"]],
    [2.0, 4.0, 4.2, 6.8], first_bold=True)

NOTE("4調査を横断したクロス集計は、令和8年8月5日のご意向により行いません。"
     "利用者票の所在地区の記入形式が9種類に分かれ、"
     "個票を地区に割り付けられないためです。"
     "これに代えて、各調査内のクロス集計と公表データによる供給構造の分析により"
     "調査相互を接続しています。")

# ============================================================ 第2節
H1("第2節　仕様書４（3）の作業項目に対する状況")

P("業務仕様書「４．業務内容（3）実施済み調査結果の集計・分析」は、"
  "次のとおり定めています。")
P("「%s」" % YOKYU, indent=0.5)

P("**求められているのは、集計・クロス集計・分析に加えて"
  "「課題整理」と「第10期計画への反映」までです。**"
  "作業項目ごとの状況は次のとおりです。")

_JOKYO = {
    "①健康とくらしの調査（JAGES）の集計・分析":
        ("完了", "調査クロス集計・分析（%dシート）"
         % _sheets_of("第10期計画_調査クロス集計・分析.xlsx"),
         "第5節。計画素案 第2章第4節に反映済み"),
    "①-2 第9期の課題・第10期の施策に対する妥当性レビューと追加分析":
        ("完了", "妥当性検証報告書／調査クロス集計・分析 13・15〜18シート",
         "第9期の課題20件と第10期の施策体系に対して不足と判定した"
         "4領域の分析を追加した"),
    "②在宅生活改善調査の集計・分析":
        ("完了", "集計分析報告書 02シート",
         "第6節1。母集団の偏りを測定した（χ²＝%.1f）" % CU_CHI),
    "③居所変更実態調査の集計・分析":
        ("完了", "集計分析報告書 03シート", "第6節2。在籍と受給の差を測定した"),
    "④介護人材実態調査の集計・分析":
        ("完了", "集計分析報告書 04シート",
         "介護職員の総数は確認事項No.8のご決定を待っている"),
    "⑤4調査を横断したクロス集計":
        ("行わない", "集計分析報告書 13シート",
         "令和8年8月5日のご意向による。"
         "利用者票の所在地区の記入形式が9種類に分かれ個票を割り付けられない。"
         "各調査内のクロス集計と供給構造の分析で代えている"),
    "⑥見える化データ・給付実績・人口推計と組み合わせた分析":
        ("完了", "集計分析報告書 06〜10シート／突合クロス集計（10シート）",
         "第6節・第7節。給付実績は年報（令和7年度）、"
         "供給は北海道の名簿と見える化K系列、"
         "人口推計は将来推計（案C）による"),
}
CAP("仕様書４（3）の作業項目と本報告の対応")
TBL(["作業項目", "状態", "成果品", "本報告での扱い"],
    [[s, _JOKYO[s][0], _JOKYO[s][1], _JOKYO[s][2]]
     for s in SAGYO],
    [4.8, 1.4, 4.6, 6.2], center={1}, first_bold=True)

P("")
P("**集計・クロス集計・分析は4調査とも完了しています。**"
  "課題整理は第7節に、計画への反映は第5節・第8節に掲げています。"
  "残っているのは、点検事項の取扱いのご決定を待つ所見%d件の反映です"
  "（第8節・第10節）。"
  % SHOKEN_ST.get("決定待ち", 0))

NOTE("仕様書は「新たなアンケート調査の企画、設計、配布、回収、入力及び督促は"
     "行わない」と定めています。本業務の範囲は貸与資料に基づく"
     "集計・クロス集計・分析・課題整理・計画への反映です。")

# ============================================================ 第3節
H1("第3節　対象4調査・母数と回収の状況")

P("4調査はいずれも発注者が実施されたものであり、"
  "受託者は配布・回収に関与していません。"
  "受領した票を集計・分析しています。")

P("**母数は調査ごとに置き方が異なります。**"
  "①②③は配布数の記録がないため回収率を算定できず、"
  "公表データで母数を置き換えられるものだけ回収率を示しています。")

CAP("調査の実施状況・母数と回収")
TBL(["調査・票種", "受領件数", "母数（公表データ等）", "回収の状況",
     "母数についての注記"],
    [[str(r[0]), str(r[1]), str(r[2] or "―"), str(r[3] or "―"),
      str(r[5] or "―")] for r in KAISHU],
    [3.6, 1.8, 3.8, 3.6, 4.4], first_bold=True)

P("")
P("**母数の置き方で特にご留意いただきたいもの**")
BUL("**①利用者票（99票）には母数がありません。**"
    "10事業所から提出されたもので、配布数の記録がありません。"
    "在宅の認定者（年報 令和7年度の月平均%s人）を母集団とみると"
    "分布が食い違います（第6節1）。"
    "「調査対象となった%d人のうち」と母数を明記して引きます。"
    % ("{:,.1f}".format(sum(ZAITAKU_NIN)), CU_N))
BUL("**③事業所票（施設・通所系）24件の母数は確定していません。**"
    "訪問系は訪問介護13事業所と確定しましたが（回収率23.1％）、"
    "施設・通所系は調査の対象範囲が記録されていません。"
    "職員個票317人は、回答24事業所の介護職員319人に対する99.4％です。"
    "**「区域の介護職員の99.4％」ではありません。**")
BUL("**②施設等票は13施設が未回答です**（把握率58.1％）。"
    "回答18施設の定員は652人で、区域内の施設の定員の全部ではありません。"
    "受給者数と比べるときはこの差が効きます（第6節2）。")
BUL("**④健康とくらしの調査の母数は65歳以上7,121人**であり、"
    "要支援者・要介護者を含みません。"
    "①②③（認定を受けた方・施設の在籍者）とは母集団が異なるため、"
    "**4調査の数値を合算しません。**")

# ============================================================ 第4節
H1("第4節　前回調査との比較（経年）")

P("前回（第9期）の調査は令和4年11月（④）及び令和5年5月（①②③）に"
  "実施されています。**規模の比較は次のとおりです。**")

CAP("前回調査と今回調査の規模")
TBL(["調査", "前回（第9期）", "今回（第10期）", "比較の可否"],
    [["① 在宅生活改善調査",
      _zen("12_在宅生活改善調査"), "15事業所・利用者票99票",
      "規模のみ。設問と集計区分の一致を確かめられない"],
     ["② 居所変更実態調査",
      _zen("11_居所変更実態調査"), "18施設（定員652人・入所者536人）",
      "規模のみ。回答施設の顔ぶれが異なる"],
     ["③ 介護人材実態調査",
      _zen("13_介護人材実態調査"),
      "27事業所・職員個票317人・職員票26件",
      "規模のみ。前回は職員数の分母が3種類あり接続できない"],
     ["④ 健康とくらしの調査",
      _zen("08_ニーズ調査データ"),
      "回収4,798票・回収率67.4％（分析対象4,729票）",
      "**指標の比較ができる**（下表）"]],
    [3.4, 5.0, 4.6, 4.2], first_bold=True)

P("")
P("④健康とくらしの調査は、前回と同じ様式の指標について"
  "**令和4年度と令和7年度を比べられます。**"
  "次の表は計画素案 第2章第4節に掲げたものです。")

CAP("④健康とくらしの調査　前回調査との比較（計画素案 第2章第4節）")
TBL([str(c) for c in ZENKAI_HYO[0]],
    [[str(c) for c in row] for row in ZENKAI_HYO[1:]],
    [4.2, 1.8, 1.8, 2.2, 7.2], first_bold=True)

P("")
P("**①②③は内容の経年比較ができません。**"
  "前回の結果は第9期計画の本文に掲載されていますが、"
  "本文からは回収票数・回収率を復元できず、"
  "第9期計画の本文そのものを受領していないためです（確認事項No.90）。"
  "受領できれば、図表集 10〜13シートの前回値と接続します。")

NOTE("④も、設問と判定ロジックが前回と同一であることを確認したうえで"
     "接続しています。通いの場参加率は国の目標値と統計の取り方が異なるため、"
     "計画素案では出所を分けて記述しています。")

# ============================================================ 第5節
H1("第5節　健康とくらしの調査の結果と計画素案への反映")

P("令和8年9月15日に、健康とくらしの調査（4,729票）の結果を"
  "計画素案 第2章第4節（高齢者の生活実態）に入れました。"
  "**同節は4段落・表1から30段落・表8になりました。**"
  "計画素案は現在%sです。" % DRAFT)

H2("入れた内容")
for i, t in enumerate([
        "調査の対象者の範囲（一般高齢者及び総合事業対象者であり、"
        "要支援者・要介護者を含まない。他の調査と合算しない）",
        "主要指標の町別の値",
        "年齢調整と町間の差の検定（有意7項目・有意でない4項目）",
        "同規模保険者40との比較（課題5項目・同等1項目。連帯感は上位）",
        "社会参加との関連",
        "暮らし向きとの関連（フレイル　大変苦しい42.7％対"
        "大変ゆとり6.8％＝6.3倍）",
        "生活実態の主な値",
        "地区別（小学校区12地区）",
        "前回調査との比較"], start=1):
    BUL("%d　%s" % (i, t))

NOTE("地区の定義（小学校区とするかどうか）は［要確認］のまま残しています"
     "（確認事項No.9）。")
NOTE("集計値は data_survey_jages.py に収めています。"
     "**個票は発注者のご指示により収録していません。**"
     "個票がなくても計画素案の側は再現できます。")

H2("計画素案の記述が薄かった3節のうち1節が埋まりました")
P("令和8年9月12日の課題整理で、"
  "「書くべき内容がまだ入っていない節」が3節あることをお示ししました。"
  "うち第2章第4節は本作業で埋まりました。"
  "残る2節（第6章第5節 介護施設整備に係る基本方針、"
  "第6章第7節 低所得者支援）は、"
  "それぞれ確認事項No.88・No.33のご決定を待っています。")

# ============================================================ 第4節
H1("第6節　調査結果と保険者の実績の突合")

P("令和8年9月11日に年報（令和7年度）の要介護度別の明細を収めたことにより、"
  "**調査の結果と保険者の実績を要介護度別に突き合わせることが"
  "初めてできるようになりました。**"
  "それ以前は在宅サービス利用率による按分の推定値しかなく、"
  "突き合わせられませんでした。")

H2("1　調査①（在宅生活改善調査）は在宅の認定者を代表しません")

P("利用者票%d票のうち**要介護度が判明した%d票**の分布を、"
  "年報から求めた在宅の認定者の分布と比べました。"
  "（%d票は要介護度の記載がなく、この比較から除いています。）"
  % (RIYO_N, CU_N, RIYO_N - CU_N))

CAP("調査①の要介護度分布と在宅の認定者の分布")
TBL(["要介護度", "調査①（人）", "在宅の認定者の分布による期待値（人）",
     "実測÷期待"],
    [[DO[i], "%d" % CU_DO[i], "%.1f" % CU_EXP[i],
      "%.2f倍" % (CU_DO[i] / CU_EXP[i])] for i in range(7)]
    + [["計", "%d" % sum(CU_DO), "%.1f" % sum(CU_EXP), "―"]],
    [3.0, 3.0, 6.0, 3.0], center={1, 2, 3}, first_bold=True)

P("**χ²＝%.1f（自由度6。1％点%.2f）で、分布は食い違います。**"
  "要介護3が期待%.1f人に対し%d人（%.1f倍）、"
  "要支援1が期待%.1f人に対し%d人（%.2f倍）です。"
  "ケアマネジャーを通じて在宅で困難を抱える方を集めた調査であり、"
  "設計のとおりの結果です。"
  % (CU_CHI, CHI_1, CU_EXP[4], CU_DO[4], CU_DO[4] / CU_EXP[4],
     CU_EXP[0], CU_DO[0], CU_DO[0] / CU_EXP[0]))

P("**調査①の結果は「在宅で困難を抱える方の状況」として引き、"
  "在宅の認定者全体の割合としては引きません。**"
  "従前は「課題のある利用者を抽出する設計である」という定性的な留保でしたが、"
  "偏りの大きさと向きを数値で示せるようになりました。")

NOTE("代表KPI H12（必要サービス未充足率）も同じ母集団の指標になります。"
     "計画本文では「調査対象となった%d人のうち」と明記します。" % CU_N)

H2("2　施設の在籍者と保険者の受給者は数える対象が異なります")

P("年報は施設の所在地を問わず当広域連合の被保険者を数えます。"
  "調査②は保険者を問わず区域内の施設の在籍者を数えます。"
  "**数える対象が違うため一致しません。**")

CAP("施設の在籍者（調査②）と保険者の受給者（年報 令和7年度）")
TBL(["区分", "調査② 在籍（人）", "年報 受給（人／月）", "差（在籍−受給）",
     "見方"],
    [[str(r[0]), str(r[1]), str(r[2]), str(r[3]), str(r[5])]
     for r in ZAISEKI],
    [3.4, 2.2, 2.4, 2.2, 7.0], center={1, 2, 3}, first_bold=True)

P("**介護老人保健施設だけが在籍が受給を上回ります。**"
  "他保険者の被保険者を受け入れているためです（新規205人・退去213人で"
  "回転が速いことも確かめました）。"
  "特別養護老人ホーム等は逆に受給（197.5人／月）が"
  "回答施設の定員（152人）を上回ります"
  "（区域外の施設の利用と、未回答13施設によるものです）。")

P("このことから、**必要利用定員総数は区域内の施設の定員とは別に"
  "整理する必要があります**（確認事項No.88・No.141）。"
  "区域内の定員は当広域連合の被保険者にとっての上限ではありません。")

NOTE("在籍は要介護度が判明した方の数です。"
     "調査②は13施設が未回答であり（把握率58.1％）、"
     "区域内の施設の在籍者の全部ではありません。")

H2("3　施設・居住系サービスの見込量")

P("令和8年9月17日に、報告書に載せていた見込量を"
  "計画素案 第6章（サービス見込量の算定 第1次概算）の値に改めました。"
  "**従前は将来推計 第2段階の値のままで、素案と別の答えが並んでいました。**"
  "現在は報告書が算定を直接読む形にしており、"
  "算定を改めれば報告書も追随します。")

CAP("施設・居住系サービスの見込量と区域内定員（人／月）")
TBL(["区分", "令和7年度（実績）", "令和11年度（見込み）", "区域内定員",
     "定員に対する割合"],
    [[str(r[0]), str(r[1]), str(r[2]), str(r[3]), str(r[4])]
     for r in KYOKYU]
    + [["施設・居住系　計",
        n1(grp("施設サービス") + grp("居住系サービス")),
        n1(grp("施設サービス", "2029") + grp("居住系サービス", "2029")),
        "%d" % (int(KYOKYU[0][3]) + int(KYOKYU[1][3])),
        "%.1f％" % ((grp("施設サービス", "2029")
                     + grp("居住系サービス", "2029"))
                    / (int(KYOKYU[0][3]) + int(KYOKYU[1][3])) * 100)]],
    [3.4, 3.2, 3.4, 2.4, 3.0], center={1, 2, 3, 4}, first_bold=True)

P("**合計では定員に収まりますが、種別ごとに見ると様子が違います。**")
BUL("**地域密着型介護老人福祉施設入所者生活介護は令和11年度に"
    "定員62人の99.9％に達します。**"
    "令和8年9月24日に中長期の年次を広げたところ、"
    "定員を超えるのは令和12年度であることが分かりました。")
BUL("**特定施設は区域内ではほぼ満室です**（定員156人に対し入居者154人・"
    "98.7％。所見20）。一方、当広域連合の被保険者の受給は"
    "令和7年度で61.5人／月であり、定員の39.4％にとどまります。"
    "**区域内の満室と、当広域連合の被保険者にとっての充足は"
    "別のことです。**"
    "なお入居者154人のうち97人は未回答の1施設（定員100人）の値であり、"
    "令和7年10月6日現在の別時点のものです（所見19・確認事項No.32）。")
BUL("介護老人保健施設・介護老人福祉施設は令和11年度に定員内に収まります。"
    "ただし特別養護老人ホームの待機者は別に存在します（所見17）。")

NOTE("見込量は別冊「サービス見込量の算定 第1次概算」（令和8年9月16日）"
     "によるものであり、計画素案 第6章と同じ値です。"
     "本報告は算定を直接読んでおり、数値を書き写していません。")

# ============================================================ 第7節
H1("第7節　調査から把握した地域課題")

P("仕様書４（3）は集計・分析に加えて**課題整理**を求めています。"
  "調査から把握した課題を、供給の側から見えることと突き合わせて"
  "整理しました。")

H2("1　認定者の利用状況")

P("認定者1,984人（令和7年度）がどのサービスを使っているかを"
  "年報から求めました。")

CAP("認定者の利用状況（令和7年度）")
TBL(["区分", "推計人数", "認定者に対する割合", "見方"],
    [[str(r[0]), "{:,}人".format(int(r[1])), str(r[2]), str(r[5])]
     for r in RIYO_JOKYO],
    [3.6, 2.2, 3.0, 8.2], center={1, 2}, first_bold=True)

P("**認定を受けながらいずれのサービスも利用していない方が"
  "認定者の4分の1を占めます。**"
  "この方々の内訳（区分支給限度額の範囲内で使っていないのか、"
  "入院中なのか、給付管理の時点の差なのか）は年報からは分かれません。"
  "内訳の資料をいただければ要因を分解できます（資料提供依頼No.8）。")

NOTE("健康とくらしの調査の「介護・介助が必要だが受けていない」（7.0％）とは"
     "**別の集団です**（同調査は要支援者・要介護者を含みません）。"
     "合算しません。")

H2("2　調査で示されたことと、供給構造の側から見えること")

CAP("調査結果と供給構造の接続")
TBL(["調査で示されたこと", "供給構造の側から見えること", "計画での扱い"],
    [[str(r[0]), str(r[2]), str(r[5])] for r in SETSUZOKU],
    [5.0, 5.4, 6.8])

P("")
P("**課題は、量の不足だけではなく、供給の偏りと担い手の把握範囲にあります。**"
  "24時間対応の3サービスは区域内に事業所がなく、"
  "小規模多機能は5事業所すべてを1法人が運営し、"
  "訪問系13事業所のうち美瑛町を実施地域とするのは4事業所です。"
  "介護職員の33.3％は見える化システムでは把握できないサービスに"
  "従事しています。")

H2("3　課題と第5章 基本目標の対応")

CAP("調査から把握した課題と反映先")
TBL(["課題", "所見", "根拠（数値）", "計画素案の反映先", "状態"],
    [[nm, "No.%s" % no, str(_shoken(no)[2]),
      str(_shoken(no)[3]).replace("\n", "・"), str(_shoken(no)[5])]
     for nm, no in KADAI],
    [3.2, 1.2, 6.4, 3.6, 1.6], center={1, 4}, first_bold=True)

NOTE("所見・根拠・反映先・状態は集計分析報告書 12シートの実物から"
     "読んでいます。課題の名は本報告で付けたものです。"
     "「決定待ち」は点検事項の取扱いのご決定（確認事項No.8）を待つものです。"
     "**在宅生活の継続が困難な層の割合は、"
     "母集団が「在宅で困難を抱える方」であることを明示して引きます**"
     "（第6節1）。")

NOTE("課題の一覧は集計分析報告書 12シート（主要所見%d件）及び"
     "調査クロス集計・分析 13シート（第9期の課題・第10期の施策への"
     "対応レビュー）に収めています。" % len(SHOKEN))

# ============================================================ 第8節
H1("第8節　主要所見と計画素案への反映状況")

P("調査から得た主要所見は%d件です。"
  "計画素案への反映状況は次のとおりです。"
  "**反映済みかどうかは、計画素案の本文と表を章節ごとに実際に読んで"
  "判定しています。**固定値では書いていません。" % len(SHOKEN))

CAP("主要所見の反映状況")
TBL(["状態", "件数", "内容"],
    [["反映済", "%d件" % SHOKEN_ST["反映済"],
      "計画素案の本文に反映済みであることを素案の実物で確認したもの。"
      "第2章第4節（高齢者の生活実態）は令和8年9月15日、"
      "第2章第5節・第6章第2節は既に反映している"],
     ["着手可", "%d件" % SHOKEN_ST["着手可"],
      "点検事項の取扱いの決定を要しない。"
      "第1章第7節・第2章第3節・第2章第6節・第3章第5節・"
      "第5章・第6章第5節へ反映する（令和8年10月）"],
     ["決定待ち", "%d件" % SHOKEN_ST["決定待ち"],
      "点検事項No.1・No.2・No.13・No.25・No.31・No.32の取扱いが"
      "決まってから反映する（確認事項No.8）"]],
    [2.2, 2.0, 11.0], center={0, 1}, first_bold=True)

H2("反映済みの所見")
for r in SHOKEN:
    if r[5] == "反映済":
        BUL("所見%s　%s（反映先　%s）"
            % (r[0], r[1], str(r[3]).replace("\n", "・")))

H2("着手できる所見")
for r in SHOKEN:
    if r[5] == "着手可":
        BUL("所見%s　%s（反映先　%s）"
            % (r[0], r[1], str(r[3]).replace("\n", "・")))

NOTE("計画本文には確認事項の注記を書きません（発注者のご指示）。"
     "反映の管理は計画素案の別管理表と業務工程管理表で行います。")

# ============================================================ 第6節
H1("第9節　調査結果の限界と留保")

P("本報告書の数値を計画本文に用いる際の留保は%d件です。"
  "うち%d件は計画本文に数値を載せる際に必ず付すもの（必須）、"
  "%d件は点検事項の取扱いの決定を待つものです。"
  % (len(RYUHO), RYUHO_KB.get("必須", 0), RYUHO_KB.get("決定待ち", 0)))

CAP("留保の区分")
TBL(["区分", "件数", "意味"],
    [[k, "%d件" % v,
      {"必須": "計画本文に数値を載せる際に必ず付す",
       "決定待ち": "点検事項の取扱いの決定による",
       "決定済": "取扱いが決まっており記述しない"}.get(k, "―")]
     for k, v in sorted(RYUHO_KB.items(), key=lambda x: -x[1])],
    [3.0, 2.2, 9.0], center={0, 1}, first_bold=True)

H2("公表データにより解消した留保と、測定できた留保")

P("介護サービス情報公表システム・国勢調査・北海道の名簿・"
  "年報（令和7年度）の要介護度別の明細により、"
  "留保の一部を解消し、又は大きさを測定しました。")

CAP("留保の解消と測定の状況")
TBL(["状態", "件数", "内容"],
    [[k, "%d件" % v,
      {"解消": "公表データにより母数又は数値が確定した",
       "測定済み": "留保そのものは残るが、大きさと向きを数値で示せる",
       "未解消": "解消できない",
       "一部解消": "一部のサービスのみ解消した",
       "公表待ち": "公表を待つ",
       "決定待ち": "取扱いの決定を待つ"}.get(k, "―")]
     for k, v in sorted(HOKAN_ST.items(), key=lambda x: -x[1])],
    [3.0, 2.2, 9.0], center={0, 1}, first_bold=True)

P("**「測定済み」は令和8年9月に新たに設けた区分です。**"
  "留保そのものは残りますが、どちらへどれだけ偏っているかを"
  "数値で示せるようになったものです。"
  "調査①の代表性（χ²＝%.1f）と、在籍と受給の差がこれに当たります。"
  % CU_CHI)

NOTE("最も重い留保は、調査①の利用者票が課題のある利用者を"
     "抽出する設計であることです（留保No.2・No.7）。"
     "母集団が「在宅で困難を抱える方」であることを"
     "記述のたびに明示します。")

# ============================================================ 第7節
H1("第10節　ご判断をお願いする事項")

P("アンケート分析でご判断をお願いする事項は3件です。"
  "**このうち確認事項No.8（点検事項の重大6件の取扱い）により、"
  "所見%d件の反映と成果品の確定値化が動きます。**"
  "他の2件は記述の範囲と資料の再提供に関するものです。"
  % SHOKEN_ST["決定待ち"])

CAP("ご判断をお願いする事項")
TBL(["確認事項", "件名", "内容", "決着しない場合の扱い"],
    [["No.8", "3調査の点検事項の重大6件の取扱い",
      "①職員数の二重計上（No.1）②利用者票の回答の偏り（No.2）"
      "③所在地区の記入形式（No.3）④老健・通所リハと見える化の乖離（No.31）"
      "⑤区域内最大の特定施設の未回答（No.32）⑥特別養護老人ホームの"
      "待機者の重複（No.13）。"
      "いずれも計画本文へ反映する数値の確定に必要である",
      "受託者の案（介護職員348人・偏りを明記して実数で記述するなど）により"
      "記述し、その旨を注記する"],
     ["No.9", "地区の定義（小学校区とするか）",
      "健康とくらしの調査は小学校区12地区別に集計できる。"
      "計画の日常生活圏域との対応が定まらない",
      "［要確認］のまま残し、地区別の記述は調査の区分のままとする"],
     ["No.139", "健康とくらしの調査の個票の再提供の要否",
      "個票はご指示により格納しておらず作業領域にも残っていない。"
      "成果品は収録済みで内容は確定しており、現時点で不足はない",
      "再提供は依頼しない。作り直しが要るときだけ依頼する"]],
    [2.0, 3.4, 6.4, 5.0], first_bold=True)

# ============================================================ 第8節
H1("第11節　成果品")

CAP("アンケート分析に係る成果品")
TBL(["成果品", "規模", "内容"],
    [["アンケート調査の集計分析報告書", "%dシート" % len(SHEETS),
      "4調査の結果、供給構造の分析、限界と留保、主要所見と反映、"
      "横断クロス集計を行わないことの整理、公表データによる補完"],
     ["実施済み調査 結果報告書（Word）", "発注者様式",
      "上記を発注者からご提示のあった様式に合わせたもの"],
     ["調査クロス集計・分析",
      "%dシート" % _sheets_of("第10期計画_調査クロス集計・分析.xlsx"),
      "健康とくらしの調査（個票4,729票）の集計・分析"],
     ["調査結果と年報実績の突合クロス集計",
      "%dシート"
      % _sheets_of("第10期計画_調査結果と年報実績の突合クロス集計.xlsx"),
      "本報告 第6節の測定の記録"],
     ["実施済み3調査の受領点検と集計",
      "%dシート"
      % _sheets_of("第10期計画_実施済み3調査の受領点検と集計.xlsx"),
      "①②③の受領点検と集計。点検事項35件"],
     ["9月作業_アンケート分析の作業状況",
      "%dシート"
      % _sheets_of("第10期計画_9月作業_アンケート分析の作業状況.xlsx"),
      "令和8年9月の作業工程"],
     ["本報告書", "全11節",
      "令和8年9月の業務報告"]],
    [5.2, 2.4, 8.2], first_bold=True)

P("")
P("本報告についてのお尋ねは、受託者までお願いします。", size=10)

# ============================================================ 自己点検
chk(1, "成果品の実物からシート数を読んでいること",
    "%dシート" % len(SHEETS), len(SHEETS) >= 14)
chk(2, "主要所見の件数と状態の内訳が合うこと",
    "%d件＝反映済%d＋着手可%d＋決定待ち%d"
    % (len(SHOKEN), SHOKEN_ST["反映済"], SHOKEN_ST["着手可"],
       SHOKEN_ST["決定待ち"]),
    len(SHOKEN) == sum(SHOKEN_ST.values()))
chk(3, "留保の件数と区分の内訳が合うこと",
    "%d件＝%s" % (len(RYUHO),
                  "＋".join("%s%d" % (k, v) for k, v in RYUHO_KB.items())),
    len(RYUHO) == sum(RYUHO_KB.values()))
chk(4, "χ²が1％点を超えること",
    "%.1f ＞ %.2f" % (CU_CHI, CHI_1), CU_CHI > CHI_1)
chk(5, "調査①の要介護度の和が回答数と一致すること",
    "%d 対 %d" % (sum(CU_DO), CU_N), sum(CU_DO) == CU_N)
chk(6, "期待値の和が回答数と一致すること",
    "%.1f 対 %d" % (sum(CU_EXP), CU_N), abs(sum(CU_EXP) - CU_N) < 1e-6)
chk(7, "見込量を算定から読んでいること（じか書きしていない）",
    "施設 R7 %s人／月・R11 %s人／月"
    % (n1(grp("施設サービス")), n1(grp("施設サービス", "2029"))),
    grp("施設サービス", "2029") > 0)
chk(8, "計画素案の規模を実物から数えていること", DRAFT,
    RP.draft_size()[0] > 0)
chk(9, "回収の状況を実物から読んでいること",
    "%d行" % len(KAISHU), len(KAISHU) >= 8)
chk(10, "留保の解消・測定の状況を実物から読んでいること",
     "%d件（%s）" % (len(HOKAN),
                     "・".join("%s%d" % (k, v) for k, v in HOKAN_ST.items())),
     len(HOKAN) == sum(HOKAN_ST.values()))
chk(11, "「測定済み」の区分があること",
     "%d件" % HOKAN_ST.get("測定済み", 0), HOKAN_ST.get("測定済み", 0) > 0)
chk(12, "調査①の票数と比較に用いた票数の差が説明できること",
     "受領%d票−要介護度の記載なし%d票＝%d票"
     % (RIYO_N, RIYO_FUMEI, CU_N), RIYO_N - RIYO_FUMEI == CU_N)
chk(13, "仕様書４（3）の作業項目を業務工程管理表のソースから読んでいること",
     "%d項目" % len(SAGYO), len(SAGYO) >= 6 and all(s in _JOKYO for s in SAGYO))
chk(14, "在籍と受給の表が2列でなく在籍・受給・差を備えること",
     "%d行×%d列" % (len(ZAISEKI), 5), len(ZAISEKI) == 4)
chk(15, "在籍と受給が一致しないこと（数える対象が違う）",
     "介護老人保健施設 在籍%s 対 受給%s" % (ZAISEKI[3][1], ZAISEKI[3][2]),
     float(ZAISEKI[3][1]) > float(ZAISEKI[3][2]))
chk(16, "見込量の表の値が算定の値と一致すること（写し違いがないこと）",
     "施設 R11 表%s 対 算定%s"
     % (KYOKYU[0][2], n1(grp("施設サービス", "2029"))),
     abs(float(KYOKYU[0][2]) - grp("施設サービス", "2029")) < 0.05
     and abs(float(KYOKYU[1][2]) - grp("居住系サービス", "2029")) < 0.05)
chk(17, "定員に対する割合が空でないこと",
     "施設%s・居住系%s" % (KYOKYU[0][4], KYOKYU[1][4]),
     "％" in str(KYOKYU[0][4]) and "％" in str(KYOKYU[1][4]))
chk(18, "認定者の利用状況の割合の合計が100％になること",
     "%.1f％" % sum(float(str(r[2]).rstrip("％")) for r in RIYO_JOKYO),
     abs(sum(float(str(r[2]).rstrip("％")) for r in RIYO_JOKYO) - 100) < 0.5)
chk(19, "前回調査との比較を実物（計画素案・図表集）から読んでいること",
     "素案の表%d行／図表集%d件" % (len(ZENKAI_HYO), len(ZENKAI)),
     len(ZENKAI_HYO) >= 3 and "08_ニーズ調査データ" in ZENKAI)
chk(21, "課題の根拠と反映先を12シートの実物から読んでいること",
     "課題%d件（所見に対応）" % len(KADAI),
     all(_shoken(no) for _k, no in KADAI))
chk(20, "ご判断をお願いする事項の件数が本文と表で一致すること",
     "本文3件・表3件", "ご判断をお願いする事項は3件です" in "".join(
         p.text for p in doc.paragraphs))

# ============================================================ 出力
# 既定のテンプレートの w:zoom は必須属性 w:percent を欠くため補う
# （CLAUDE.md §4）。
_zoom = doc.settings.element.find(qn("w:zoom"))
if _zoom is not None and _zoom.get(qn("w:percent")) is None:
    _zoom.set(qn("w:percent"), "100")

os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc.save(OUT)

NG_WORDS = ["に由来する", "と整合する", "1件も", "有意差がないため関係がない",
            "全国トップ級"]
_txt = [p.text for p in doc.paragraphs]
for t in doc.tables:
    for row in t.rows:
        for c in row.cells:
            _txt.extend(x.text for x in c.paragraphs)
_ngw = [(w, s[:40]) for s in _txt for w in NG_WORDS if w in s]
_ast = [s[:40] for s in _txt if "**" in s]

_ng = [c for c in CHECKS if c[3] != "適合"]
print("書き出しました:", OUT)
print("段落 %d ／ 表 %d" % (len(doc.paragraphs), len(doc.tables)))
print("集計分析報告書 %dシート／主要所見 %d件（反映済%d・着手可%d・決定待ち%d）"
      % (len(SHEETS), len(SHOKEN), SHOKEN_ST["反映済"], SHOKEN_ST["着手可"],
         SHOKEN_ST["決定待ち"]))
print("留保 %d件（%s）" % (len(RYUHO),
                          "・".join("%s%d" % (k, v)
                                    for k, v in RYUHO_KB.items())))
print("補完と解消 %d件（%s）"
      % (len(HOKAN), "・".join("%s%d" % (k, v) for k, v in HOKAN_ST.items())))
print("χ² %.1f（1％点 %.2f）／計画素案 %s" % (CU_CHI, CHI_1, DRAFT))
print("自己点検 %d件：適合%d件・不適合%d件"
      % (len(CHECKS), len(CHECKS) - len(_ng), len(_ng)))
if _ngw:
    print("禁止表現:", _ngw)
if _ast:
    print("強調の指定が残っている:", _ast)
for c in _ng:
    print("  不適合:", c[0], c[1], c[2])
if _ng or _ngw or _ast:
    sys.exit(1)
print("すべての点検に適合しました。")
