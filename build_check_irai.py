# -*- coding: utf-8 -*-
"""他のメンバーへのチェック依頼　確認箇所と手順.

令和8年9月18日のご依頼
  「他メンバーにチェック依頼をお願いする為、確認箇所と手順を
    まとめて整理したものをワードファイルに作成」

本書は、第10期介護保険事業計画の策定支援業務の成果品を、
この作業に携わっていない者が点検できるようにするための手引きである。
点検する側が、何を、どこで、どうやって確かめるのかが分かることを目的とする。

**件数・シート数・進捗・素案の規模は、いずれも実物又はソースから数える。**
固定値で書き写すと、成果品を改訂したときに本書とずれる（CLAUDE.md §4）。
  成果品の件数と区分　`data_dispatch.DISPATCH`
  ビルドスクリプトの本数　`build_*.py` を実際に数える
  自己点検の件数　各スクリプトのソースを `ast` で読み `chk(` の呼出しを数える
  計画素案の規模　`repo_paths.draft_label()`（実物を数える）
  未確定箇所　計画素案の実物から［要協議］等を数える
  確認事項　`build_process_control.py` の `CHECK` を `ast` で読む
  算定値　`build_mikomiryo_santei.py` を `runpy` で読む

**担当者名は書かない**（発注者指示。CLAUDE.md §1）。
分担の欄は〔　〕の記入枠としている。

出力
  output/第10期計画_チェック依頼_確認箇所と手順.docx
"""

import ast
import glob
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

import repo_paths as RP
import data_dispatch as DP
import data_progress as G

OUT = RP.ROOT + "/output/第10期計画_チェック依頼_確認箇所と手順.docx"
KIJUNBI = "令和8年9月18日"
FONT = "游ゴシック"
NAVY = RGBColor(0x1F, 0x38, 0x64)

# 禁止表現（発注者指示。CLAUDE.md §1）
NG_WORDS = ["に由来する", "と整合する", "1件も", "有意差がないため関係がない",
            "全国トップ級"]

# 受領資料をセッション固有の場所から読むため再実行できないもの
NO_RERUN = ["build_kickoff_fix.py", "build_minutes_fix.py",
            "build_survey_crosstab.py"]

CHECKS = []


def chk(no, naiyo, kekka, ok):
    CHECKS.append((no, naiyo, kekka, "適合" if ok else "不適合"))
    return ok


# ============================================================ 実物から数える
def scripts():
    return sorted(os.path.basename(p)
                  for p in glob.glob(os.path.join(RP.ROOT, "build_*.py")))


SCRIPTS = scripts()


def selfcheck_counts():
    """各スクリプトの自己点検の件数を、ソースの chk( の呼出しから数える。

    自己点検を持つスクリプトは、不適合があると終了コード1で終わる。
    点検する側は「実行して終了コード0であること」を見ればよい。
    """
    out = {}
    for name in SCRIPTS:
        try:
            src = io.open(os.path.join(RP.ROOT, name), encoding="utf-8").read()
            tree = ast.parse(src)
        except (SyntaxError, UnicodeDecodeError):
            continue
        n = 0
        for node in ast.walk(tree):
            if (isinstance(node, ast.Call)
                    and isinstance(node.func, ast.Name)
                    and node.func.id == "chk"):
                n += 1
        if n:
            out[name] = n
    return out


SELFCHK = selfcheck_counts()


def sheets_of(fname):
    """成果品（xlsx）のシート数を実物から数える。"""
    path = os.path.join(RP.OUTPUT, fname)
    if not os.path.exists(path):
        return None
    from openpyxl import load_workbook
    return len(load_workbook(path, read_only=True).sheetnames)


def ph_count():
    """計画素案に残る未確定箇所を実物から数える。"""
    out = {"要協議": 0, "要確認": 0, "要内訳": 0}
    if not os.path.exists(RP.DRAFT):
        return out
    from docx import Document as _D
    d = _D(RP.DRAFT)
    buf = [p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            for c in row.cells:
                buf.extend(p.text for p in c.paragraphs)
    s = "\n".join(buf)
    for k in out:
        out[k] = s.count("［%s］" % k)
    return out


PH = ph_count()


def kakunin():
    """確認事項を build_process_control.py のソースから読む。"""
    src = io.open(RP.ROOT + "/build_process_control.py",
                  encoding="utf-8").read()
    m = re.search(r"^CHECK = (\[.*?^\])", src, re.S | re.M)
    if not m:
        return []
    return ast.literal_eval(m.group(1))


KAKUNIN = kakunin()
KUBUN = {}
for _c in DP.DISPATCH.values():
    KUBUN[_c[0]] = KUBUN.get(_c[0], 0) + 1


# ============================================================ 算定値を読む
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


_M = _load("build_mikomiryo_santei.py")
G_DO = _M["G_DO"]
G_ITTEI = _M["G_ITTEI"]
SUEOKI = _M["SUEOKI"]
SOU3 = sum(_M["KYUFU3"])                 # 総給付費（3か年計）
GEPPGAKU = int(round(G_DO["月額"]))
KIJUN_GAKU = int(round(GEPPGAKU / 100.0) * 100)


def yen(v):
    return "{:,}".format(int(round(v)))


# ============================================================ 自己点検
chk(1, "ビルドスクリプトを実物から数えていること",
    "%d本" % len(SCRIPTS), len(SCRIPTS) > 100)
chk(2, "自己点検を持つスクリプトを数えられること",
    "%d本・計%d件" % (len(SELFCHK), sum(SELFCHK.values())),
    len(SELFCHK) >= 10)
chk(3, "成果品を data_dispatch から数えていること",
    "%d件（%s）" % (len(DP.DISPATCH),
                    "・".join("%s%d" % (k, v) for k, v in sorted(
                        KUBUN.items(), key=lambda x: -x[1]))),
    len(DP.DISPATCH) > 100)
chk(4, "確認事項をソースから読めること",
    "%d件（最大No.%d）" % (len(KAKUNIN), max(c[0] for c in KAKUNIN)),
    len(KAKUNIN) > 100)
chk(5, "計画素案の規模を実物から数えていること",
    RP.draft_label(), RP.draft_size()[0] > 0)
chk(6, "未確定箇所を実物から数えていること",
    "%d箇所（要協議%d・要確認%d・要内訳%d）"
    % (sum(PH.values()), PH["要協議"], PH["要確認"], PH["要内訳"]),
    sum(PH.values()) > 0)
chk(7, "算定値を build_mikomiryo_santei.py から読めること",
    "総給付費 %s円／算定上の月額 %s円／保険料基準額 %s円"
    % (yen(SOU3), yen(GEPPGAKU), yen(KIJUN_GAKU)),
    SOU3 > 0 and GEPPGAKU > 0)
chk(8, "再実行できないスクリプト3本が実在すること",
    "／".join(NO_RERUN),
    all(os.path.exists(os.path.join(RP.ROOT, n)) for n in NO_RERUN))
chk(9, "据え置きの一覧を読めること", "%d件" % len(SUEOKI), len(SUEOKI) > 20)


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
    """** で囲んだ部分を太字にして段落へ流し込む。

    docx は Markdown を解釈しないため、** をそのまま書くと
    本文に ** が出る（CLAUDE.md §4 で実際に起きた誤り）。
    ここで強調の指定として解釈し、** 自体は出力しない。
    """
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
    if not p.runs:                      # 空文字でも段落は残す
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


def CODE(text):
    """コマンドの行。等幅で示す。"""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    r.font.name = "Consolas"
    r.font.size = Pt(9.5)
    r.element.rPr.rFonts.set(qn("w:eastAsia"), "Consolas")
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
            _runs(p, "" if v is None else v, size,
                  bold=(first_bold and j == 0))
    return t


# ============================================================ 表紙
P("大雪地区広域連合　第10期介護保険事業計画策定支援業務", size=12,
  align=WD_ALIGN_PARAGRAPH.CENTER)
p = P("成果品のチェック依頼　　確認箇所と手順", size=17, bold=True,
      align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
for r in p.runs:
    r.font.color.rgb = NAVY
P("基準日　%s　　受託者　ビズアップ公共コンサルティング株式会社" % KIJUNBI,
  size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
P("")

P("本書は、この業務に携わっていない方に成果品を点検していただくための"
  "手引きです。何を、どこで、どうやって確かめるのかを示しています。")
P("点検の対象は成果品%d件（%s）と計画素案（%s）です。"
  "ビルドスクリプトは%d本あり、うち%d本が自らの計算を検算する"
  "自己点検を備えています（計%d件）。"
  % (len(DP.DISPATCH),
     "・".join("%s%d件" % (k, v)
                for k, v in sorted(KUBUN.items(), key=lambda x: -x[1])),
     RP.draft_label(), len(SCRIPTS), len(SELFCHK), sum(SELFCHK.values())))
P("すべてを見ていただく必要はありません。"
  "第1節の優先順に沿って、上から順にお願いします。"
  "第8節に、この業務で実際に起きた誤りを挙げています。"
  "同じ形の誤りが残っていないかという見方が、最も早く不備に行き当たります。")


# ============================================================ 第1節
H1("第1節　点検の優先順")

P("点検には3つの層があります。上の層ほど、誤りがあったときの影響が大きく、"
  "かつ短い時間で確かめられます。")

CAP("点検の3層と優先順")
TBL(["層", "何を見るか", "所要", "誰でもできるか"],
    [["第1層\n機械で確かめる",
      "スクリプトを再実行し、自己点検が通ること（終了コード0）と、"
      "docx が validate.py に適合することを確かめる。"
      "計算の誤りはここでほぼ捕まる",
      "30分",
      "できる。第7節の手順どおりに実行するだけ"],
     ["第2層\n数値の出所を辿る",
      "成果品に載っている数値が、受領資料のどこから来たものかを辿る。"
      "同じ数値が2つの成果品で食い違っていないかを見る",
      "2〜3時間",
      "できる。第4節に出所の対照表がある"],
     ["第3層\n記述の当否を見る",
      "計画素案の本文が、制度の定めと算定の結果に照らして正しいか。"
      "守るべき制約（個人情報・二次情報・禁止表現）に触れていないか",
      "半日",
      "制度の知識があると早い。第5節・第6節"]],
    [2.4, 8.2, 1.6, 5.0], first_bold=True)

NOTE("第1層だけでも行っていただく値打ちがあります。"
     "この業務で実際に見つかった誤りのうち、算定に関わるものは"
     "いずれも自己点検か再実行で捕まえたものです。")


# ============================================================ 第2節
H1("第2節　点検を始める前に")

H2("１　作業の場所")
P("リポジトリの直下で作業します。出力はすべて output/ に落ちます。"
  "スクリプトは自らの置かれている場所を基準に出力先を決めるため"
  "（repo_paths.py の ROOT）、どのディレクトリから実行しても"
  "そのリポジトリの output/ に書き込まれます。")
CODE("cd <リポジトリ>")
CODE("python3 build_mikomiryo_santei.py")

H2("２　必要なもの")
BUL("Python 3（python-docx、openpyxl）")
BUL("docx の検証には /mnt/skills/public/docx/scripts/office/validate.py")
BUL("目次のページ番号を作り直す場合は LibreOffice Writer。"
    "入っていない環境では apt-get install -y libreoffice-writer が要ります")

H2("３　再実行できないスクリプトが3本あります")
P("次の3本は、受領資料をセッション固有の場所から読むため再実行できません。"
  "実行してエラーになっても不備ではありません。")
for _n in NO_RERUN:
    BUL(_n)
NOTE("この3本が作った成果品は output/ に収録済みで、内容は確定しています。"
     "作り直しが要るときだけ、個票等の再提供を発注者に依頼します。")

H2("４　やってはいけないこと")
BUL("個票データ（個人情報）をリポジトリに置かない。集計値のみを収める")
BUL("担当者名・電話番号・メールアドレスを成果品に書かない"
    "（施設の管理者名を含む）")
BUL("他団体の計画書等を成果品に用いない。記載の借用も行わない。"
    "比較材料としてのみ用い、作業終了後に消去する")
BUL("二次情報（民間の情報サイト・業界紙）を出典に使わない。原典のみ")
BUL("計画素案の本文に確認事項の注記を書かない。"
    "本文の注記は、図表を読むために必要なもの"
    "（出典・単位・欠測・数値の性質・記号の意味）に限る")
BUL("確定しない値を推測して埋めない。"
    "［要協議］［要確認］［要内訳］のまま残す")


# ============================================================ 第3節
H1("第3節　点検の対象")

H2("１　自己点検を備えたスクリプト")
P("次のスクリプトは、自らの計算を別の経路から検算し、"
  "1件でも合わなければ終了コード1で終わります。"
  "点検する側は、実行して終了コード0であることを見ればよいことになります。")

CAP("自己点検を備えたスクリプト（%d本・計%d件）"
    % (len(SELFCHK), sum(SELFCHK.values())))
_rows = [[k, "%d件" % v] for k, v in sorted(SELFCHK.items(),
                                            key=lambda x: (-x[1], x[0]))]
_half = (len(_rows) + 1) // 2
_pair = []
for i in range(_half):
    a = _rows[i]
    b = _rows[i + _half] if i + _half < len(_rows) else ["", ""]
    _pair.append(a + b)
TBL(["スクリプト", "自己点検", "スクリプト", "自己点検"], _pair,
    [6.3, 1.6, 6.3, 1.6], size=8, center={1, 3})

H2("２　特に見ていただきたい成果品")
CAP("優先して見ていただきたい成果品")
_p1 = "第10期計画_サービス見込量算定_第1次概算.xlsx"
_p2 = "第10期計画_アンケート調査の集計分析報告書.xlsx"
_p3 = "第10期計画_計画素案の課題整理.xlsx"
TBL(["成果品", "規模", "なぜ優先するか"],
    [["計画素案（docx）", RP.draft_label(),
      "発注者・委員会・住民の目に触れるもの。"
      "未確定箇所が%d箇所（要協議%d・要確認%d・要内訳%d）残っている"
      % (sum(PH.values()), PH["要協議"], PH["要確認"], PH["要内訳"])],
     ["サービス見込量の算定 第1次概算",
      "%sシート" % (sheets_of(_p1) or "―"),
      "保険料の基礎になる。算定上の月額%s円・保険料基準額%s円は"
      "ここから出ている" % (yen(GEPPGAKU), yen(KIJUN_GAKU))],
     ["アンケート調査の集計分析報告書",
      "%sシート" % (sheets_of(_p2) or "―"),
      "計画素案 第2章・第3章の記述の裏づけ。"
      "見込量は算定スクリプトを読む形にしてある"],
     ["計画素案の課題整理",
      "%sシート" % (sheets_of(_p3) or "―"),
      "未確定箇所と確認事項の対応。"
      "決着しない場合の既定値がここに置いてある"],
     ["業務工程管理表", "確認事項%d件" % len(KAKUNIN),
      "確認事項の台帳。成果品の記述と備考が食い違っていないかを見る"]],
    [4.4, 2.4, 9.4], first_bold=True)
NOTE("成果品の一覧と区分（送付・条件付き・内部保管・対象外）は"
     "data_dispatch.py にあります。"
     "「送付」は発注者にお渡しするもの、「条件付き」は前提を付してお渡しするもの、"
     "「内部保管」は社内にとどめるものです。")


# ============================================================ 第4節
H1("第4節　点検箇所①　数値の出所")

H2("１　同じ数値が2か所に別々に書かれていないか")
P("この業務で実際に起きた誤りです。"
  "アンケート分析報告書の見込量が将来推計 第2段階のままで、"
  "計画素案 第6章（第1次概算）と別の答えが並んでいました。")
P("いまは、数値を持つ成果品が算定スクリプトを runpy で読む形にしてあります。"
  "算定を改めれば両方が追随します。"
  "**点検では、成果品に数値がじか書きされていないかを見ます。**")
CODE("grep -n '%s\\|%d\\|%s' build_*.py"
     % (yen(GEPPGAKU), GEPPGAKU, yen(SOU3)))
NOTE("算定スクリプト（build_mikomiryo_santei.py）以外の場所に"
     "これらの数値がじか書きされていたら、それは追随しない写しです。"
     "数値を持つ成果品は、算定スクリプトを runpy で読む形にしてあります"
     "（build_plan_draft.py・build_survey_report.py ほか）。")

H2("２　件数・シート数・段落数を固定値で書いていないか")
P("件数を固定値で書くと、成果品を改訂したときに管理表とずれます。"
  "令和8年9月11日の点検で、素案の段落・表・図の数が"
  "管理表と合っていないことが分かりました（実物806/130/36に対し管理表667/114/34）。")
CAP("実物から数えるための関数")
TBL(["数えるもの", "使う関数", "置いてある場所"],
    [["計画素案の段落・表・図（現時点）", "RP.draft_label()", "repo_paths.py"],
     ["計画素案の段落・表・図（過去の時点）", "RP.draft_label_at(基準日)",
      "repo_paths.py。凍結値"],
     ["成果品（xlsx）のシート数", "_sheets(ファイル名)",
      "build_process_control.py"],
     ["進捗（現時点）", "data_progress.PROGRESS", "data_progress.py"],
     ["進捗（過去の時点）", "G.snapshot(基準日)", "data_progress.py"],
     ["未確定箇所", "本書の ph_count()／build_soan_kadai.py",
      "計画素案の実物から数える"],
     ["確認事項", "CHECK を ast で読む", "build_process_control.py"]],
    [5.2, 5.2, 6.0], first_bold=True)

H2("３　過去の時点を記す資料が現時点の値を読んでいないか")
P("令和8年9月17日の点検で見つかった誤りです。"
  "業務進捗報告書 令和8年8月分が data_progress.PROGRESS と"
  "repo_paths.draft_label() を直接読んでいたため、"
  "再出力のたびに8月分の報告の進捗率と素案の規模が書き換わっていました"
  "（素案の規模は667→806→843→896段落と3回書き換わっていた）。")
P("**過去の時点を記す資料（月次の進捗報告書・議事録）は、"
  "必ず snapshot(基準日) と draft_label_at(基準日) を使います。**")
CODE("grep -n 'draft_label()\\|PROGRESS' build_monthly_*.py build_minutes_*.py")

H2("４　出所により値が違うものがあります")
P("同じものを指す数値が、出所により違う値になります。"
  "これは誤りではなく、数える対象が違うことによるものです。"
  "成果品で、どちらの値を使っているかが明示されているかを見ます。")
CAP("出所により値が異なるもの")
TBL(["項目", "出所A", "出所B", "違いの理由"],
    [["施設・居住系の月平均利用者数",
      "計画作成支援ツール（3町合計）527人",
      "年報（保険者単位）480.17人",
      "区域外の施設を利用する当連合の被保険者と、"
      "区域内の施設を利用する他保険者の被保険者の扱いが違う"],
     ["令和7年度の給付費",
      "3町のファイルの合計\n3,213,054,558円",
      "広域連合のファイル\n3,083,474,291円",
      "129,580,267円（4.20％）の差。令和6年度・令和8年度は円単位で一致する。"
      "抽出条件を照会中（確認事項No.87）"],
     ["要介護認定率（令和7年度）",
      "年報（年度末）21.826％",
      "見える化の総括表（9月末）21.307％",
      "時点の違い。分子はいずれも第1号被保険者の認定者で、"
      "第2号被保険者は含まない"],
     ["認知症対応型共同生活介護の定員",
      "北海道の名簿 7事業所108人",
      "見える化 5事業所99人",
      "くるみの郷（東川町）が公表システムに掲載がなく定員未確定"],
     ["施設の在籍者と受給者",
      "調査② 区域内の施設の在籍者",
      "年報 当連合の被保険者の受給者",
      "数える対象が違う。介護老人保健施設だけが在籍206人＞受給134.1人／月"]],
    [3.4, 3.6, 3.6, 5.8], first_bold=True)

H2("５　定員と読み違えやすい値が4つあります")
P("いずれも実際に取り違えたものです。")
BUL("data_shien_tool.SHISETSU は「定員」ではなく**月平均利用者数**")
BUL("data_shakai_shigen.SHIGEN の第5〜7要素は"
    "従業員総数・常勤・非常勤であって定員ではない")
BUL("data_kyufu_jisseki の「1人1月あたり利用回数」"
    "（認知症対応型共同生活介護の108はこれ。定員ではない）")
BUL("見える化M系列の「合計」（内訳を足すと二重になる）")


# ============================================================ 第5節
H1("第5節　点検箇所②　守るべき制約")

H2("１　個人情報")
P("値の形で探します。語の有無で探すと偽陽性が出ます。")
CODE("grep -rEn '0[0-9]{1,3}-[0-9]{2,4}-[0-9]{4}' output/ *.py")
CODE("grep -rEn '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+' output/ *.py")
BUL("担当者名・施設の管理者名が列見出しや本文に入っていないか")
BUL("個票データ（回答者単位の行）が収録されていないか")
NOTE("社内の作業ファイルに担当者名が列見出しとして1件あったため、"
     "そのファイル自体はリポジトリに収めず、"
     "集計値のみを data_member_work.py に収めています。")

H2("２　禁止表現（5件）")
P("発注者から示された5件です。"
  "点検方法を記述したシート（この5件を掲げている箇所）を除いて判定します。")
for _w in NG_WORDS:
    BUL("「%s」" % _w)
P("「〜と整合する」は、**根拠として用いている場合**が禁止です。"
  "令和8年9月17日に、確認事項No.84の"
  "「北海道の指定事業所一覧に掲載がなく休止と整合する」を"
  "「北海道の指定事業所一覧にも掲載がない」に改めました。")
CODE("grep -rn 'に由来する\\|と整合する\\|全国トップ級' *.py")
NOTE("自己点検を備えたスクリプトの多くは、"
     "出力したブックの全セルを走査して禁止表現を検出し、"
     "見つかれば終了コード1で終わります。"
     "実行して0で終われば、その成果品には禁止表現がありません。")

H2("３　転載を禁ずる旨の記載がある受領資料が2件あります")
CAP("転載禁止の記載がある受領資料")
TBL(["資料", "記載", "当方の扱い"],
    [["通いの場の課題解決に向けたマニュアル\n（厚生労働省・令和6年3月）",
      "「本マニュアルの内容、テキスト、画像等の無断転載・無断使用を"
      "固く禁じます」",
      "出典として参照するにとどめ、記述・図表を転載しない"],
     ["令和7年度北海道介護予防活動普及展開事業\n普及啓発セミナー資料\n"
      "（北海道保健福祉部福祉局高齢者保健福祉課）",
      "「テキスト、音声、音楽、画像、映像等を無断使用、複製、改編したり"
      "販売や貸与することを禁止します」",
      "成果品に用いず、記載の借用も行わない。"
      "個人情報を含むためリポジトリに格納せず作業領域にとどめる"]],
    [4.6, 6.4, 5.4], first_bold=True)

H2("４　他団体の資料")
P("他団体の計画書等は、成果品に用いず、記載の借用も行いません。"
  "比較材料としてのみ用い、作業終了後に消去します。"
  "団体名は協議の場に限り示します。")
NOTE("厚生労働省の年報（全保険者の集計表）は、他団体の計画書等とは別のものです。"
     "当連合の行を用いることに差し支えはありません。")


# ============================================================ 第6節
H1("第6節　点検箇所③　計画素案の本文")

H2("１　未確定箇所は残してあります")
P("現在%d箇所（要協議%d・要確認%d・要内訳%d）です。"
  "うち第6章第4節（施設の見込み）に58箇所（43％）が集中しており、"
  "日常生活圏域と必要利用定員総数（確認事項No.88）が決まれば解消します。"
  "**これらを推測で埋めないでください。**"
  % (sum(PH.values()), PH["要協議"], PH["要確認"], PH["要内訳"]))

H2("２　本文に確認事項の注記を書かない")
P("計画素案の本文に、作業上のやり取りを書かない方針です（発注者指示）。"
  "令和8年9月17日に受領した加筆版には"
  "「〜が欲しい」「もらえれば」「大雪のOKが出れば」といった"
  "「※（案）」の注記が15件入っていました。"
  "これらは本文に置かず、確認事項・資料提供依頼として登録します。")
P("本文に置いてよい注記は、図表を読むために必要なものに限ります"
  "（出典・単位・欠測・数値の性質・記号の意味）。")

H2("３　単位・時点・基準年度がそろっているか")
P("実際に誤りが出た箇所です。")
CAP("単位・時点で誤りが出た箇所")
TBL(["箇所", "誤り", "正しい扱い"],
    [["第6章第2節（4）利用回（日）数",
      "通所介護・地域密着型通所介護・通所リハビリテーション・"
      "認知症対応型通所介護の4件を「日」と書いていた",
      "年報様式1の7も総括表詳細（３）も「回」。"
      "算定の KAISU_MAP から単位を採る形に改めた"],
     ["第6章第2節　見込量の表頭",
      "「R8実績」と書いていた",
      "令和8年度の月報（4〜7月）から計算した実績見込み値であり"
      "年度の実績ではない。「R7実績」に改めた"],
     ["第6章第4節2　定員の照合",
      "第2節・第3節が令和7年度基準に変わったのに"
      "第4節2だけ令和8年度基準のままだった",
      "同じ論点に2つの答えが素案内に併存していた。"
      "算定の TEIIN_MAP・JISSEKI から引く形に改めた"],
     ["第2章第1節　人口の時点",
      "本文は住民基本台帳の各年10月1日現在で統一しているが、"
      "加筆版の図は令和8年1月1日現在",
      "時点をそろえる。そろえられない場合は図の注記で明示する"]],
    [3.6, 6.2, 6.6], first_bold=True)

H2("４　反映済みかどうかを語の有無だけで判定しない")
P("章節を限って判定します。"
  "「除雪」は第1章第7節の除雪率にも、"
  "「同規模保険者」は第3章第4節にも現れるため、"
  "語の有無だけで見ると偽陽性が出ます。"
  "build_survey_report.py の _draft_sections()／in_draft(章節, 語...) を"
  "使ってください。")

H2("５　法定記載事項")
P("介護保険法第117条第2項の各号に当たる記載があるかを見ます。")
BUL("第1号　サービス種類ごとの量の見込み・必要利用定員総数"
    "（利用回（日）数を含む）→ 第6章第2節・第4節")
BUL("第2号　地域支援事業の量の見込み → 第6章第3節2")
BUL("第3号・第4号　給付適正化 → 第5章 基本目標5（3）・資料5")
BUL("第3項　任意記載事項（認知症施策はここに置いている）"
    " → 第1章第2節・第5章 基本目標2（3）")


# ============================================================ 第7節
H1("第7節　手順")

H2("手順1　スクリプトを再実行する（第1層）")
P("自己点検を備えた%d本を順に実行し、終了コードが0であることを確かめます。"
  "1件でも合わなければ終了コード1で終わり、不適合の内容が表示されます。"
  % len(SELFCHK))
CODE("for f in build_*.py; do")
CODE("  case \"$f\" in %s) continue;; esac"
     % "|".join(NO_RERUN))
CODE("  python3 \"$f\" >/dev/null 2>&1 || echo \"不適合: $f\"")
CODE("done")
NOTE("再実行できない3本（第2節3）は除いています。"
     "自己点検を持たないスクリプトも実行されますが、"
     "書き出しに失敗すればやはり0以外で終わります。"
     "自己点検を備えた%d本は第3節1の表のとおりです。" % len(SELFCHK))

H2("手順2　自己点検が働いていることを確かめる")
P("点検が素通りしていないことを確かめます。"
  "データの値をわざと書き換えて、"
  "自己点検が不適合を出すことを見ます（変異試験）。"
  "確かめたら必ず元に戻します。")
CODE("git diff            # 元に戻したことを確かめる")

H2("手順3　docx を検証する")
P("OOXML は要素の順序が定められており、"
  "python-docx で生の要素を足すと順序を誤ることがあります。"
  "**作った docx は必ず検証します。**")
CODE("python3 /mnt/skills/public/docx/scripts/office/validate.py "
     "output/<ファイル名>.docx")
NOTE("実際に誤りが出たもの：w:pgNumType（cols より前）、"
     "w:updateFields（compat より前）、w:tabs（spacing・ind より前）、"
     "w:bookmarkStart（pPr の次）、w:tblCellMar（tblLook より前）、"
     "w:shd（tcPr では vAlign より前）。"
     "w:zoom は必須属性 w:percent を欠くため \"100\" を補います。")

H2("手順4　見出しを増減したら目次を作り直す")
CODE("python3 build_toc_pages.py")
NOTE("LibreOffice Writer が入っていない環境では"
     "apt-get install -y libreoffice-writer が要ります。")

H2("手順5　数値の出所を辿る（第2層）")
P("成果品の数値を1つ選び、受領資料まで遡ります。"
  "第4節1の grep で、算定スクリプト以外の場所に数値が"
  "じか書きされていないかを見ます。")

H2("手順6　守るべき制約を見る（第5節）")
P("個人情報は値の形で、禁止表現は点検方法の記述シートを除いて判定します。")

H2("手順7　計画素案の本文を読む（第3層）")
P("第6節の観点で読みます。未確定箇所は埋めません。")

H2("手順8　気づいたことを記録する")
P("第9節の様式に書き出してください。"
  "直してよいかどうかの判断は、こちらで行います。"
  "**その場で直さず、まず記録をお願いします。**"
  "数値の変更は複数の成果品に波及するためです。")


# ============================================================ 第8節
H1("第8節　陥りやすい誤り（この業務で実際に起きたもの）")

P("同じ形の誤りが残っていないかという見方が、最も早く不備に行き当たります。")

CAP("実際に起きた誤りと、それを防ぐ手立て")
TBL(["何が起きたか", "どう気づいたか", "手立て"],
    [["過去の時点を記す資料が現時点の値を読んでいた。"
      "業務進捗報告書 令和8年8月分の素案の規模が"
      "再出力のたびに667→806→843→896段落と書き換わっていた",
      "再レビューの際に、同じ資料の数値が前回と違うことに気づいた",
      "data_progress.snapshot(基準日) と"
      "repo_paths.draft_label_at(基準日) を新設し、"
      "令和8年8月31日の素案を実際に数えた値で凍結した"],
     ["同じ数値を2つの成果品に別々に書いていた。"
      "アンケート分析報告書の見込量が将来推計 第2段階のままで、"
      "計画素案 第6章と食い違っていた",
      "報告書を読み直した際に、素案と数値が違うことに気づいた",
      "算定スクリプトを runpy で読む形に改めた。"
      "算定を改めれば両方が追随する"],
     ["受領資料を主要値だけ取り込み、明細を取り込んでいなかった。"
      "按分による推定値を作り、明細の提供を求める確認事項を起票していた",
      "同じファイルの再送により、明細が最初から入っていたことが分かった",
      "確認事項を起こす前に、"
      "受領済みのファイルにその値がないかを確かめる"],
     ["差の原因を1つに決めてから確かめていなかった。"
      "見える化の認定率と当方の差を"
      "「分子に第2号被保険者を含むかの違い」と書いていた",
      "受領資料の数値で逆算したところ、"
      "当方の分子に第2号は入っていなかった",
      "当たらない原因を確認事項に書くと、発注者の確認も空振りになる。"
      "**逆算して確かめてから書く**"],
     ["件数・シート数を固定値で書いていた。"
      "素案の段落・表・図の数が管理表と合っていなかった"
      "（実物806/130/36に対し管理表667/114/34）",
      "ブランチ内資料の点検で指摘された",
      "実物から数える関数を用いる（第4節2の表）"],
     ["出力先を絶対パスで固定していた。"
      "別ブランチの作業ツリーから実行すると生成物がそこへ落ちる",
      "別の場所から実行したときに実際に起きた",
      "repo_paths.py の RP.ROOT に統一した。"
      "全スクリプトを別ディレクトリから実行して確認済み"],
     ["反映済みかどうかを語の有無だけで判定していた",
      "「除雪」が第1章第7節の除雪率にも現れ、偽陽性が出た",
      "章節を限って判定する（in_draft(章節, 語...)）"],
     ["docx を検証していなかった",
      "OOXML の要素順序の誤りを3件抱えていた",
      "作った docx は必ず validate.py で検証する"]],
    [5.6, 4.6, 6.2], size=8.5)


# ============================================================ 第9節
H1("第9節　点検の記録様式")

P("気づいたことを次の様式で記録してください。"
  "直してよいかどうかの判断はこちらで行います。")

CAP("点検記録の様式")
TBL(["No.", "対象（成果品・章節・シート）", "気づいたこと",
     "根拠（どの資料のどこ）", "重み", "点検者"],
    [["1", "", "", "", "", "〔　　〕"],
     ["2", "", "", "", "", "〔　　〕"],
     ["3", "", "", "", "", "〔　　〕"]],
    [1.0, 4.0, 5.0, 4.0, 1.4, 1.8], center={0, 4, 5})

P("重みは次の3つでお願いします。")
BUL("**重大**　数値が変わる、又は法定記載事項を欠く")
BUL("**中**　記述の当否に関わるが数値は変わらない")
BUL("**軽**　表記・体裁")

NOTE("「これは誤りではないか」と断定できなくてかまいません。"
     "「ここが分からない」「ここは出所が書いていない」も記録の対象です。"
     "読んで分からない箇所は、発注者にも分からない箇所です。")


# ============================================================ 第10節
H1("第10節　分担の案")

P("2〜3名で分けていただく場合の案です。"
  "第1層は誰か1名が通しで行い、第2層・第3層を分けるのが早いと考えます。")

CAP("分担の案")
TBL(["担当", "範囲", "所要の目安", "担当者"],
    [["A", "第1層すべて（スクリプトの再実行・docx の検証）と"
      "第4節1・2（数値のじか書き・固定値）",
      "半日", "〔　　　　　〕"],
     ["B", "計画素案 第1章〜第5章（第6節の観点）と"
      "第5節（個人情報・禁止表現・転載）",
      "半日", "〔　　　　　〕"],
     ["C", "計画素案 第6章と"
      "サービス見込量の算定 第1次概算（%sシート）"
      % (sheets_of(_p1) or "―"),
      "半日", "〔　　　　　〕"]],
    [1.2, 9.6, 2.2, 4.2], first_bold=True, center={0, 2})

H2("期限の目安")
P("直近の工程は、サービス見込量 第2次概算（令和8年10月30日）と"
  "予算編成用（令和8年11月13日）です。"
  "第1次概算（令和8年9月26日）は既に算定を終えており、"
  "提示できる状態にあります。"
  "**第2次概算までにご指摘をいただけると、算定に反映できます。**")

P("")
P("本書についてのお尋ねは、受託者までお願いします。", size=10)


# ============================================================ 出力
# 既定のテンプレートの w:zoom は必須属性 w:percent を欠くため補う
# （CLAUDE.md §4）。
_zoom = doc.settings.element.find(qn("w:zoom"))
if _zoom is not None and _zoom.get(qn("w:percent")) is None:
    _zoom.set(qn("w:percent"), "100")

os.makedirs(RP.OUTPUT, exist_ok=True)
doc.save(OUT)

_ng = [c for c in CHECKS if c[3] != "適合"]
print("書き出しました:", OUT)
print("ビルドスクリプト %d本／自己点検を備えるもの %d本・計%d件"
      % (len(SCRIPTS), len(SELFCHK), sum(SELFCHK.values())))
print("成果品 %d件（%s）"
      % (len(DP.DISPATCH),
         "・".join("%s%d" % (k, v)
                    for k, v in sorted(KUBUN.items(), key=lambda x: -x[1]))))
print("計画素案 %s／未確定箇所 %d箇所" % (RP.draft_label(), sum(PH.values())))
print("確認事項 %d件／全体進捗 %d％" % (len(KAKUNIN), G.overall_pct()))
print("算定上の月額 %s円／保険料基準額 %s円" % (yen(GEPPGAKU), yen(KIJUN_GAKU)))
print("自己点検 %d件：適合%d件・不適合%d件"
      % (len(CHECKS), len(CHECKS) - len(_ng), len(_ng)))
for c in _ng:
    print("  不適合:", c[0], c[1], c[2])
if _ng:
    sys.exit(1)
print("すべての点検に適合しました。")
