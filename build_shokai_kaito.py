# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画　計画策定の進捗と今後の確認事項.

令和8年10月2日　大雪地区広域連合からのご照会
  「実施済み調査の結果報告書について（中略）内容を確認いたしました。
    一方、報告書内には、介護職員数の重複、特養待機者数、未回答施設等、
    取扱いが未確定となっている事項も複数あります。
    また、これまで第9期計画の評価・検証、第10期計画骨子案、KPI等について
    協議を進めてきたところですが、今回の報告書を含め、現在、計画策定全体の
    どの段階まで進んでおり、今後、何をどの順序で確認・決定していくのかが、
    当方で十分整理できていない状況です。
    つきましては、今後の対応を整理するため、次の事項について簡潔に
    整理のうえ共有をお願いいたします。
      ・今回の報告書の計画策定全体における位置づけと、
        調査分析として今後修正・追加を予定している事項
      ・現時点で広域連合に確認・判断を求める事項
      ・広域連合又は構成3町で追加確認・資料提供が必要な事項
      ・今後の計画素案、サービス見込量・給付費・保険料推計等の
        作成・協議スケジュール
      ・次回協議までの貴社及び広域連合それぞれの対応事項
    なお、報告書に記載されている「実施済み3調査の受領点検と集計」等、
    未確定事項の確認主体、期限、反映先等を整理した管理資料がありましたら、
    併せて共有をお願いいたします。」

━━ 本資料の考え方 ━━

1 お尋ねの5項目に、第2節から第7節が1対1で対応する。
  第8節に、未確定事項をどの資料のどこで管理しているかを示す。

2 **新たな台帳を起こさない。** 未確定事項は業務工程管理表（確認事項一覧・
  資料提供依頼）を正とし、調査に固有のものは実施済み3調査の受領点検と集計
  （点検で見つかった事項・計画本文への反映方針）による。
  本資料はそれらを読んで束ねたものである。

3 件数が多いものは絞り込みの条件を明記する。
  未決の確認事項は160件あり、そのまま掲げると用をなさない。

━━ 数値の出どころ ━━

進捗は業務内容別の進捗、確認事項と資料提供依頼は業務工程管理表、
見込量・給付費・保険料はサービス見込量の算定、
調査の未確定事項は実施済み調査 結果報告書及び
実施済み3調査の受領点検と集計、工程は業務工程管理表の月別工程による。

構成
  第1節　本資料の位置づけ
  第2節　計画策定全体の進み方と結果報告書の位置づけ
  第3節　調査分析として今後修正・追加を予定している事項
  第4節　現時点でご確認・ご判断をお願いする事項
  第5節　広域連合又は構成3町で追加確認・資料提供が必要な事項
  第6節　今後の作成・協議スケジュール
  第7節　次回協議までのそれぞれの対応事項
  第8節　未確定事項の管理資料

出力
  output/第10期計画_計画策定の進捗と今後の確認事項.docx

自己点検で1件でも不適合があると作成を止める。
"""

import docx_fix
import ast
import contextlib
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

from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import data_kitei as DK                                   # noqa: E402
import data_gaisan1 as G1                                 # noqa: E402
import data_progress as DP                                # noqa: E402
import repo_paths as RP                                   # noqa: E402

if isinstance(sys.stdout, io.TextIOWrapper):
    sys.stdout.reconfigure(encoding="utf-8")

OUT = RP.ROOT + "/output/第10期計画_計画策定の進捗と今後の確認事項.docx"
HOKOKU = RP.ROOT + "/output/第10期計画_実施済み調査_結果報告書.docx"
TENKEN = RP.ROOT + "/output/第10期計画_実施済み3調査の受領点検と集計.xlsx"

# ご照会の日。事実であり、本資料を作り直しても変わらない。
SHOKAI_BI = "令和8年10月2日"
# 基準日は業務内容別の進捗の基準日に合わせる。
KIJUNBI = DP.KIJUNBI

TOWNS = ["東川町", "美瑛町", "東神楽町"]
FONT = "游ゴシック"
NAVY = RGBColor(0x1F, 0x38, 0x64)

CHECKS = []


def chk(no, naiyo, kekka, ok):
    CHECKS.append((no, naiyo, kekka, "適合" if ok else "不適合"))
    return ok


# ==================================================== 台帳と算定を読む
def _load(name):
    with contextlib.redirect_stdout(io.StringIO()):
        return runpy.run_path(RP.ROOT + "/" + name)


def _literal(fn, name):
    src = io.open(RP.ROOT + "/" + fn, encoding="utf-8").read()
    for node in ast.parse(src).body:
        if (isinstance(node, ast.Assign)
                and any(getattr(t, "id", None) == name
                        for t in node.targets)):
            return ast.literal_eval(node.value)
    raise RuntimeError("%s に %s が見つからない" % (fn, name))


# 確認事項（正の台帳）
# (0)No (1)業務内容 (2)工程 (3)表題 (4)内容 (5)止めている成果物
# (6)確認先 (7)状態 (8)期限 (9)回答
CHECK = _literal("build_process_control.py", "CHECK")
# 資料提供依頼
# (0)No (1)領域 (2)資料 (3)内容 (4)確定する主張 (5)入手先
# (6)希望時期 (7)優先度 (8)状態 (9)備考
LACK = _literal("build_process_control.py", "LACK")
# 月別の工程
# (0)時期 (1)内容 (2)業務 (3)成果物 (4)状況 (5)状態 … (9)備考
MONTH = _literal("build_process_control.py", "MONTH")
# 第2次概算に向けた段取り (時期, 事項, 内容, 主体)
DANDORI = _literal("build_gaisan2.py", "DANDORI")
# 既定値のとおり計画素案へ反映した事項 (No, 反映先, 内容)
SUSUMETA = _literal("build_ikenkokankai.py", "SUSUMETA")
# 据え置きの区分（Z＝確定済／A＝受託者で確定できる／B＝発注者・3町／C＝国）
KUBUN = _literal("build_mikomi_juryo_nashi.py", "KUBUN")
# 決着しない場合の当方の扱い（1か所から引く）
KITEI = DK.all_kitei()

_M = _load("build_mikomiryo_santei.py")
SUEOKI = _M["SUEOKI"]
G3 = _M["G_DO"]
GETSU = G3["月額"]
KIJUN_GAKU = int(round(GETSU / 100.0)) * 100
SOU3 = sum(_M["KYUFU3"])

KANRYO = ("完了", "了承済", "了承済（保管せず廃棄）", "代替により解消", "解決")
MIKETSU = [c for c in CHECK if c[7] not in KANRYO]
LACK_MACHI = [x for x in LACK
              if x[8] not in ("受領済", "完了", "解消", "代替により解消")
              and x[5] != "受託者"]

# 既定値のとおり反映し終えた確認事項の番号
HANEI_NO = {x[0] for x in SUSUMETA}


# ==================================================== 結果報告書を読む
def _dtxt(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def _tables(path):
    """(見出し行の連結, 行の一覧) を文書の並び順に返す。"""
    import docx as _dx
    d = _dx.Document(path)
    out = []
    for el in d.element.body.iter(qn("w:tbl")):
        rows = el.findall(qn("w:tr"))
        if not rows:
            continue
        cells = [[_dtxt(c).strip() for c in r.findall(qn("w:tc"))]
                 for r in rows]
        out.append(("".join(cells[0]), cells[1:]))
    return out


_HT = _tables(HOKOKU)


def _find_table(head_prefix):
    for h, rows in _HT:
        if h.startswith(head_prefix):
            return rows
    raise RuntimeError("結果報告書に「%s」の表が見つからない" % head_prefix)


def _genkai_n():
    """結果報告書 第8章第3節（1）に掲げた限界の件数を数える。"""
    import docx as _dx
    d = _dx.Document(HOKOKU)
    on, n = False, 0
    for p in d.paragraphs:
        t = p.text.strip()
        if t.startswith("（１） 本報告書の限界"):
            on = True
            continue
        if on and t.startswith("（２）"):
            break
        if on and re.match(r"^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫]", t):
            n += 1
    return n


GENKAI_N = _genkai_n()


def _bekkanri_n():
    """計画素案の別管理表の確認事項A・Bの件数を数える。"""
    p = RP.ROOT + "/output/第10期計画_計画素案の別管理表.xlsx"
    if not os.path.exists(p):
        return (None, None)
    wb = load_workbook(p, data_only=True)
    out = []
    for nm in ("03_確認事項A_計画の文言", "04_確認事項B_数値・計算根拠"):
        n = 0
        for row in wb[nm].iter_rows(values_only=True):
            v = str(row[0] or "").strip()
            if v and (v.isdigit()
                      or (v[0] in "AB" and v[1:].isdigit())):
                n += 1
        out.append(n)
    return tuple(out)


BEK_A, BEK_B = _bekkanri_n()


# 第8章第2節（1）調査結果とサービス見込量の関係（8段階）
ICHIZUKE = _find_table("段階用いる資料")
# 第8章第3節（2）取扱いを確定していない事項（9件）
MIKAKUTEI = _find_table("事項内容確定しないと決まらないこと")


# ==================================================== 受領点検を読む
_WB = load_workbook(TENKEN, data_only=True)


def _rows(sheet, ncol):
    ws = _WB[sheet]
    out = []
    for row in ws.iter_rows(values_only=True):
        if row[0] is None or not str(row[0]).strip().isdigit():
            continue
        out.append([("" if v is None else str(v).strip())
                    for v in row[:ncol]])
    return out


# 02_点検で見つかった事項
# No, 重要度, 調査, 対象, 内容, 取扱い・確認すること, 確認主体, 期限, 反映先
TENKEN_JIKO = _rows("02_点検で見つかった事項", 9)
JUDAI = [x for x in TENKEN_JIKO if x[1] == "重大"]
# 09_計画本文への反映方針
# No, 該当箇所, 反映する内容, 前提となる確認, 時期, 状態
HANEI_HOSHIN = _rows("09_計画本文への反映方針", 6)

# 計画素案の章節ごとのテキスト。
# 「反映を完了した」と書くものは、実際にその章節にあることを確かめる。
DRAFT_SEC = RP.draft_sections()
# 反映を完了したものと、その章節に現れるべき値
HANEI_KANRYO = {
    8: ("第2章第3節", ("住宅型有料老人ホーム", "223", "66", "50")),
    9: ("第6章第5節", ("57.0", "47.0", "病院・診療所")),
}


def _in_draft(section, words):
    t = DRAFT_SEC.get(section, "")
    return bool(t) and all(w in t for w in words)

# 確認主体の区分
_SHUTAI_JIGYOSHO = "提出元の事業所（発注者を経由）"
TENKEN_JIGYOSHO = [x for x in TENKEN_JIKO if x[6] == _SHUTAI_JIGYOSHO]
TENKEN_HACCHU = [x for x in TENKEN_JIKO
                 if x[6].startswith("発注者") and x[6] != _SHUTAI_JIGYOSHO]
TENKEN_JUTAKU = [x for x in TENKEN_JIKO if x[6].startswith("受託者")]


# ==================================================== 体裁
doc = Document()
_st = doc.styles["Normal"]
_st.font.name = FONT
_st.font.size = Pt(10.5)
_st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
_st.paragraph_format.space_after = Pt(4)
_st.paragraph_format.line_spacing = 1.15
_sec = doc.sections[0]
_sec.page_width, _sec.page_height = Cm(21.0), Cm(29.7)
_sec.top_margin = _sec.bottom_margin = Cm(1.8)
_sec.left_margin = _sec.right_margin = Cm(1.9)
TEXTW = 21.0 - 1.9 * 2


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
    for tag, v in (("top", 0.06), ("left", 0.12), ("bottom", 0.06),
                   ("right", 0.12)):
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


def _runs(p, text, size, bold=False, color=None):
    """** で囲んだ部分を太字にする（docx は強調記号を解釈しない）。"""
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


def P(text="", size=10.5, bold=False, align=None, after=4):
    p = doc.add_paragraph()
    _runs(p, text, size, bold)
    if align:
        p.alignment = align
    p.paragraph_format.space_after = Pt(after)
    return p


# 節の先頭で改ページする節。全ての節に入れると、直前の節の本文が
# あふれた1〜2行だけのページができる。実際の紙面で確かめて決める。
BRK = set(os.environ.get("BRK", "2,3,5,6,8").split(",")) - {""}
_SECNO = [0]


def H1(text, brk=None):
    _SECNO[0] += 1
    if brk is None:
        brk = str(_SECNO[0]) in BRK
    p = doc.add_paragraph()
    p.paragraph_format.page_break_before = brk
    p.paragraph_format.space_before = Pt(0 if brk else 14)
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


KEEP_MAX = 14


def TBL(head, rows, widths, size=9, center=None, first_bold=False,
        bold=None, keep=None):
    center = center or set()
    bold = bold or set()
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
                  bold=(first_bold and j == 0) or j in bold)
    if keep if keep is not None else (1 + len(rows) <= KEEP_MAX):
        for tr in t.rows[:-1]:
            for c in tr.cells:
                for pp in c.paragraphs:
                    pp.paragraph_format.keep_with_next = True
    return t


def _n(v, nd=0):
    return format(round(v, nd), ",.%df" % nd)


def _yen(v):
    return format(int(round(v)), ",")


def _pct(v):
    return "%.0f％" % (v * 100)


def _cut(v, n):
    """n 文字を超えるものを切る。文の途中で切れたまま載せない。

    n 文字までに句点があればそこで切り、無ければ「…」を付けて切る。
    """
    v = str(v)
    if len(v) <= n:
        return v
    i = v.rfind("。", 0, n + 1)
    if i >= int(n * 0.5):
        return v[:i + 1]
    return v[:n] + "…"


def _oki(s):
    """算定の表の文を写すときに、他の資料で意味を持たない語を落とす。

    算定の側の表で用いている言い回し（自己点検の番号、シートの番号、
    「本表」）は本資料では意味を持たないため落とす（CLAUDE.md §4）。
    """
    v = str(s)
    v = re.sub(r"（点検[0-9０-９・点検]*）", "", v)
    v = re.sub(r"[0-9０-９]+シートに", "サービス見込量の算定に", v)
    v = v.replace("本表は", "").replace("本表の", "").replace("本表", "本資料")
    return v.strip()


# ============================================================ 表紙
P("大雪地区広域連合　第10期介護保険事業計画", size=12,
  align=WD_ALIGN_PARAGRAPH.CENTER)
_p = P("計画策定の進捗と今後の確認事項", size=18, bold=True,
       align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
for _r in _p.runs:
    _r.font.color.rgb = NAVY
P("令和8年10月", size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
P("")
P("%s にいただいたご照会に対する回答です。基準日 %s。"
  % (SHOKAI_BI, KIJUNBI))
P("実施済み調査 結果報告書の計画策定全体における位置づけ、"
  "今後修正・追加を予定している事項、"
  "ご確認・ご判断をお願いする事項、"
  "追加確認・資料提供が必要な事項、"
  "今後の作成・協議スケジュール及び次回協議までのそれぞれの対応事項を"
  "お示しします。")
P("**未確定事項の管理は、新たな台帳を起こさず、"
  "業務工程管理表と実施済み3調査の受領点検と集計によっています。** "
  "いずれも既にお送りしているもので、確認主体・期限・反映先を"
  "項目ごとに掲げています。所在は第8節のとおりです。")

# ============================================================ 第1節
H1("第1節　本資料の位置づけ")
P("ご照会のご趣旨は、計画策定全体の中で現在どこまで進んでおり、"
  "今後、何をどの順序で確認・決定していくのかを共有したい、"
  "というものと受け止めております。"
  "お尋ねの5項目に、本資料の第2節から第7節が対応します。")

CAP("お尋ねの事項と本資料の対応")
TBL(["お尋ねの事項", "本資料"],
    [["今回の報告書の計画策定全体における位置づけ", "第2節"],
     ["調査分析として今後修正・追加を予定している事項", "第3節"],
     ["現時点で広域連合に確認・判断を求める事項", "第4節"],
     ["広域連合又は構成3町で追加確認・資料提供が必要な事項", "第5節"],
     ["今後の計画素案、サービス見込量・給付費・保険料推計等の"
      "作成・協議スケジュール", "第6節"],
     ["次回協議までの貴社及び広域連合それぞれの対応事項", "第7節"],
     ["未確定事項の確認主体・期限・反映先等を整理した管理資料",
      "第8節"]],
    [13.6, 3.6], center={1}, first_bold=True)

H2("この3点を先に申し上げます")
BUL("**結果報告書は、施策とKPIの根拠及び見込量の上振れ・下振れの"
    "検討に用いる基礎資料です。** "
    "サービス見込量そのものは介護保険事業状況報告（年報）と"
    "見える化システムの給付実績により算定しており、"
    "報告書の件数を需要量に換算してはいません。"
    "したがって、報告書で取扱いを確定していない事項が"
    "見込量・給付費・保険料の算定を止めることはありません（第2節）。")
BUL("**ご決定をお待ちしている事項は、すべて「決まらない場合の当方の扱い」を"
    "仮置きしています。** したがってサービス見込量 第2次概算"
    "（令和8年10月30日）はご決定を待たずに組めます。"
    "ただし仮置きは協議を進めるためのものであり、"
    "計画として確定する前にはご判断を要します（第4節）。")
BUL("**算定上の月額基準額は%s円、百円未満を四捨五入した保険料基準額は"
    "%s円**（第9期と同額）です。"
    "現在の置き方は、単価を令和7年度で固定し、"
    "施策による増加も給付の適正化による抑制も織り込んでいないもので、"
    "保険料が低く出る側に寄っています（第4節2）。"
    % (_n(GETSU, 2), _n(KIJUN_GAKU)))

# ============================================================ 第2節
H1("第2節　計画策定全体の進み方と結果報告書の位置づけ")

H2("1　計画策定の段階と現在の位置")
P("計画策定は、次の5つの段階で進めています。"
  "現在は③がおおむね終わり、④を進めながら⑤の手続のご決定を"
  "お願いする段階です。")
CAP("計画策定の段階と現在の位置")
TBL(["段階", "内容", "状況"],
    [["① 第9期計画の評価・検証",
      "19施策の評価、対計画比、代表KPIの基準値",
      "おおむね終了（第9期の施策・事業実績の一部が未受領）"],
     ["② 現状分析と調査分析",
      "人口・認定者・給付費の分析、4調査の集計・分析、地域の課題",
      "終了（**結果報告書はこの段階の成果です**）"],
     ["③ 将来推計",
      "人口・認定者数、サービス見込量、給付費、保険料",
      "第1次概算（令和8年9月26日）を提示済み。"
      "第2次概算（令和8年10月30日）に向けて算定中"],
     ["④ 計画素案の作成",
      "全6章と資料編。法定記載事項の網羅、制度改正の反映",
      "協議用素案と公表ベースの版を作成済み。"
      "未確定箇所はご決定・ご提供を待って埋める"],
     ["⑤ 意見聴取と確定",
      "構成3町との協議、策定委員会、意見公募手続、北海道の意見の聴取",
      "3町との協議を令和8年10月に行う。"
      "意見公募手続の時期と方法はご決定をお願いする事項"]],
    [3.6, 6.0, 7.6], first_bold=True)

H2("2　業務内容別の進捗")
_PR = [x for x in DP.PROGRESS if x[2] is not None]
ZENTAI = sum(x[2] for x in _PR) / len(_PR)
CAP("業務内容別の進捗（基準日 %s）" % KIJUNBI)
TBL(["業務内容", "進捗", "状態", "業務内容", "進捗", "状態"],
    [[DP.PROGRESS[i][0] + " " + DP.PROGRESS[i][1],
      _pct(DP.PROGRESS[i][2]) if DP.PROGRESS[i][2] is not None else "―",
      DP.PROGRESS[i][3],
      DP.PROGRESS[i + 6][0] + " " + DP.PROGRESS[i + 6][1],
      (_pct(DP.PROGRESS[i + 6][2])
       if DP.PROGRESS[i + 6][2] is not None else "―"),
      DP.PROGRESS[i + 6][3]] for i in range(6)],
    [4.6, 1.5, 1.5, 4.6, 1.5, 1.5], center={1, 2, 4, 5},
    first_bold=True)
NOTE("全体は、進捗を数値で表せる10業務の単純平均で%sです。"
     "（11）成果品の帰属及び（12）その他は数値で表していません。"
     % _pct(ZENTAI))

H2("3　結果報告書を計画のどこに用いるか")
P("結果報告書そのものに、サービス見込量の算定の各段階で"
  "用いるか用いないかを示しています（同報告書 第8章第2節（1））。"
  "**見込量の算定そのものには用いず、施策とKPIの根拠及び"
  "上振れ・下振れの検討に用います。**")
P("**認定者数の見込みは、最終的には地方創生総合戦略による"
  "総人口の推移から算定します。** "
  "アンケートの結果は、この算定の妥当性を確認するために用いるものであり、"
  "認定者数そのものを求めるために用いるものではありません。")
CAP("サービス見込量の算定と結果報告書の関係")
TBL(["算定の段階", "用いる資料", "結果報告書の位置づけ"],
    ICHIZUKE, [4.2, 6.4, 6.6], first_bold=True)
NOTE("認定者数は、総人口を地方創生総合戦略、"
     "年齢階級別の構成を住民基本台帳の実績の趨勢により推計した"
     "将来人口に、性別・年齢階級別・要介護度別の認定率を乗じて算定しています。"
     "見込量は、この認定者数に"
     "介護保険事業状況報告（年報・令和7年度）の要介護度別の"
     "サービス利用率を乗じて算定しています。"
     "報告書の件数（例えば「より適切と思われるサービス」）を"
     "需要量に換算してはいません。")

P("")
P("このため、報告書で取扱いを確定していない事項は、"
  "**計画本文の記述とKPIの基準値には影響しますが、"
  "見込量・給付費・保険料の算定は止めません。** "
  "ご照会でお挙げいただいた3件についても同じです。")
CAP("ご照会でお挙げいただいた3件の現在の扱い")
TBL(["事項", "現在の扱い", "確定しないと決まらないこと", "影響する範囲"],
    [["介護職員数の重複",
      "介護職員348人から361人の**範囲で記述**",
      "介護職員の総数",
      "計画素案 第2章第6節（介護人材）の記述。"
      "見込量・保険料には入らない"],
     ["特別養護老人ホームの待機者数",
      "延べ137人・重複を除くと82人の**範囲で記述**",
      "整備の必要量の判断",
      "計画素案 第6章第4節・第5節（整備の方針）。"
      "必要利用定員総数のご判断の材料"],
     ["居所変更実態調査の未回答13施設",
      "区域内31施設のうち18施設の回答として**範囲を明記して記述**",
      "入所前の居場所・退去先の集計の代表性",
      "計画素案 第2章第3節・第6章第5節（医療と介護の連携）の記述"]],
    [3.2, 4.4, 3.4, 6.2], first_bold=True)

# ============================================================ 第3節
H1("第3節　調査分析として今後修正・追加を予定している事項")

H2("1　修正・追加を予定している事項の全体")
P("調査分析として今後予定しているものは、次の3つに分かれます。"
  "いずれも新たな集計を起こすものではなく、"
  "**取扱いが決まった時点で確定値に置き換える**ものです。")
CAP("今後の修正・追加の区分")
TBL(["区分", "件数", "内容", "必要なこと"],
    [["① 確認の結果により数値が変わるもの",
      "%d件" % len(MIKAKUTEI),
      "結果報告書で範囲又は参考として記述している事項。"
      "確定した時点で確定値に置き換え、報告書と計画素案を改める",
      "提出元の事業所又は広域連合によるご確認"],
     ["② 計画本文への反映を予定しているもの",
      "%d件" % len(HANEI_HOSHIN),
      "調査結果を計画素案のどの箇所にどう反映するかを定めたもの。"
      "うち2件は前提となる確認を要しないため、"
      "**作業を継続し計画素案への反映を完了しています**",
      "①の確認（完了した2件を除く）"],
     ["③ 次期調査への申し送り",
      "3件",
      "調査票の設計に関するもの（所在地区の選択肢、"
      "「市内／市外」の定義、事業所の所在町の記入欄）",
      "第11期計画の調査設計時のご検討"]],
    [5.0, 1.4, 6.6, 4.2], center={1}, first_bold=True)
NOTE("調査票の企画・設計は発注者が行うものであり、③は受託者からの提案です。")

H2("2　取扱いを確定していない事項と確認主体・期限・反映先")
P("結果報告書 第8章第3節（2）に掲げた%d件について、"
  "確認主体・確認の期限及び反映先を併せて示します。"
  "確認主体と期限は、実施済み3調査の受領点検と集計"
  "（02 点検で見つかった事項）に掲げているものです。" % len(MIKAKUTEI))

# 未確定事項に、点検で見つかった事項の確認主体・期限・反映先を結び付ける
_KEY = [
    ("介護職員数の重複", "有料老人ホーム華"),
    ("採用者数・離職者数の重複", "有料老人ホーム華"),
    ("利用者票の回答の偏り", "利用者票の回答の偏り"),
    ("利用者の所在地区", "利用者の所在地区"),
    ("老健・通所リハの介護職員数", "老健・通所リハの介護職員数"),
    ("居所変更実態調査の未回答13施設", "特定施設の1施設が回答していない"),
    ("特別養護老人ホームの待機者", "待機者数がいずれも55人"),
    ("認知症対応型共同生活介護1事業所の定員", "くるみの郷の定員"),
    ("特定施設の入居者の要介護度構成", "特定施設の1施設が回答していない"),
]


def _shutai(jiko):
    """未確定事項の名から、点検で見つかった事項の確認主体・期限・反映先を得る。"""
    key = dict(_KEY).get(jiko)
    if key is None:
        return ("―", "―", "―")
    for x in TENKEN_JIKO:
        if key in x[3] or key in x[4]:
            return (x[6], x[7], x[8] or "―")
    return ("―", "―", "―")


_MK_ROWS = []
for row in MIKAKUTEI:
    s, k, h = _shutai(row[0])
    _MK_ROWS.append([row[0], row[3], s, k, h])
CAP("取扱いを確定していない事項の確認主体・期限・反映先")
TBL(["事項", "確定までの扱い", "確認主体", "確認の期限", "反映先"],
    _MK_ROWS, [4.0, 2.8, 4.0, 3.0, 3.4], first_bold=True)
NOTE("「範囲で記述」は最小値と最大値を併記するもの、"
     "「参考として記述」は本文に数値を載せず注記にとどめるもの、"
     "「使用しない」は確定するまで計画本文に載せないものです。"
     "「確認主体」が「提出元の事業所（発注者を経由）」であるものは、"
     "照会の文面を既に作成しており、"
     "未回答時の扱いのご決定（確認事項No.8）を待って発出します。")

H2("3　集計の方法を決めないと計画本文の数値が変わるもの（重大%d件）"
   % len(JUDAI))
P("点検で見つかった事項%d件のうち、集計の方法を決めないと"
  "計画本文に載せる数値が変わるものが%d件あります。"
  "**このうち確認事項No.8（未回答時の扱い）のご決定が、"
  "調査の確定値化を止めている唯一の事項です。**"
  % (len(TENKEN_JIKO), len(JUDAI)))
CAP("重大と判定した事項")
TBL(["調査", "対象", "決めること", "確認主体", "反映先"],
    [[x[2], _cut(x[3], 34), _cut(x[5], 54), x[6], x[8] or "―"]
     for x in JUDAI],
    [1.8, 3.8, 6.4, 3.2, 2.0], first_bold=False)

H2("4　確認を要しないため反映を完了したもの")
_CHAKUSHU = [x for x in HANEI_HOSHIN if x[5] == "反映済"]
P("計画本文への反映を予定している%d件のうち、"
  "前提となる確認を要しないものが%d件あります。"
  "**こちらは作業を継続し、計画素案への反映を完了しています。**"
  "ご決定をお待ちしている事項はありません。"
  % (len(HANEI_HOSHIN), len(_CHAKUSHU)))
CAP("確認を要しないため反映を完了したもの")
TBL(["計画素案の該当箇所", "反映した内容", "状況"],
    [[x[1], _cut(x[2], 130), "完了"] for x in _CHAKUSHU],
    [4.4, 10.6, 2.2], center={2}, first_bold=True)

# ============================================================ 第4節
H1("第4節　現時点でご確認・ご判断をお願いする事項")

H2("1　全体の件数と絞り込みの条件")
P("確認事項は全%d件、うち未決が%d件です。"
  "そのまま掲げると用をなさないため、本節では"
  "**保険料の月額を動かすもの**と"
  "**法定記載事項に直結するもの**に絞ってお示しします。"
  "全件は業務工程管理表（03 確認事項一覧）のとおりです。"
  % (len(CHECK), len(MIKETSU)))
CAP("確認事項の状況")
TBL(["区分", "件数", "内容"],
    [["全体", "%d件" % len(CHECK), "業務着手時から起票したもの"],
     ["決着済み", "%d件" % (len(CHECK) - len(MIKETSU)),
      "ご了承・ご回答をいただいたもの、代替により解消したもの"],
     ["未決", "%d件" % len(MIKETSU),
      "ご決定・ご確認をお待ちしているもの"],
     ["　うち扱いを仮置きしたもの", "%d件" % len(MIKETSU),
      "**全件に「決まらない場合の当方の扱い」を置いています**"],
     ["　うち計画素案へ反映済み", "%d件" % len(SUSUMETA),
      "仮置きのとおり計画素案に反映し、ご決定を待っているもの"]],
    [5.6, 2.0, 9.6], center={1}, first_bold=True)
NOTE("仮置きは協議を進めるためのものであり、ご決定に代わるものではありません。"
     "計画として確定する前にご判断を要します。")

H2("2　保険料の月額を動かすご判断")
P("算定の前提%d件のうち、発注者・構成3町のご判断による%d件と"
  "国の告示・公布による%d件を、月額への効きの大きい順に示します。"
  "**国の告示・公布によるものは催促のできないものです。** "
  "告示が出れば、ご提供をお待ちしている資料がなくても"
  "計画の値は確定します。"
  % (len(SUEOKI),
     sum(1 for v in KUBUN.values() if v == "B"),
     sum(1 for v in KUBUN.values() if v == "C")))

_YEN_RE = re.compile(r"([0-9,]+)\s*円")


def _eff(text):
    """「月額」に続いて現れる円のうち最大のもの。無ければ 0。"""
    v = []
    for m in re.finditer("月額", text or ""):
        seg = text[m.end():m.end() + 24]
        v += [int(x.replace(",", "")) for x in _YEN_RE.findall(seg)]
    return max(v) if v else 0


_SU = []
for i, s in enumerate(SUEOKI, start=1):
    kb = KUBUN.get(i)
    if kb not in ("B", "C"):
        continue
    _SU.append((_eff(s[5]), kb, s))
_SU.sort(key=lambda x: -x[0])
_KB_NA = {"B": "発注者・3町", "C": "国の告示・公布"}
CAP("保険料の月額を動かす前提（効きの大きい順）")
TBL(["算定の前提", "決めること", "どこから", "関係する確認事項",
     "月額への効き"],
    [[s[0], s[1], _KB_NA[kb],
      "―" if s[4] in ("―", "") else s[4], _oki(s[5])]
     for _e, kb, s in _SU],
    [3.2, 3.6, 2.2, 2.4, 5.8], center={2, 3}, first_bold=True,
    keep=False)
NOTE("「月額への効き」は、現在の置き方から動かした場合に"
     "算定上の月額基準額が何円動くかです。"
     "現在の置き方は、単価を令和7年度で固定し、施策による増加も"
     "給付の適正化による抑制も織り込まないものです。"
     "**上向きの要素を織り込んでいないため、"
     "現在の%s円は低く出る側に寄っています。**" % _n(GETSU))

H2("3　法定記載事項に直結するご判断")
P("介護保険法第117条が計画に定めることを求めている事項のうち、"
  "ご決定がないと確定できないものです。"
  "いずれも計画素案には仮置きのとおり記述し、"
  "確定しない箇所は［要協議］［要確認］として残しています。")
CAP("法定記載事項に直結するもの")
TBL(["事項", "確認事項", "現在の仮置き", "確認先"],
    [["日常生活圏域と必要利用定員総数（法第117条第2項第1号）",
      "No.88・No.84",
      "3町（3圏域）を基盤整備単位とし、"
      "必要利用定員総数は現定員の据え置き案を掲げて"
      "本文は［要協議］のまま残す", "発注者・3町"],
     ["混合型特定施設入居者生活介護に係る必要利用定員総数"
      "（同条第3項）",
      "No.113", "定めないこととする旨を記載する", "発注者"],
     ["介護予防・日常生活支援総合事業の量の見込み"
      "（同条第2項第2号）",
      "No.112", "令和6年度の利用者実人数を基礎として見込む", "発注者"],
     ["認知症施策推進計画の位置づけ（認知症基本法第13条）",
      "No.130・No.149",
      "構成3町がそれぞれ市町村計画を策定し、"
      "本計画は法第117条第3項により認知症施策を掲げる"
      "（案B。ご了承済み）", "発注者・3町"],
     ["被保険者の意見を反映させるために必要な措置（法第117条）",
      "No.33ほか",
      "意見公募手続・認知症の人及びその家族等からの意見の聴取・"
      "北海道の意見の聴取・公表の4件を記載し、"
      "時期と方法は［要協議］のまま残す", "発注者"],
     ["地域密着型サービスの人員・設備・運営の基準"
      "（法第78条の4第1項）",
      "No.95",
      "広域連合の条例を受領していないため、"
      "令和6年の省令改正への対応を確認できていない", "発注者"]],
    [5.0, 1.8, 7.6, 2.8], first_bold=True, keep=False)

H2("4　調査の確定値化を止めている1件")
P("**確認事項No.8（点検で見つかった事項の重大%d件の取扱い）** です。"
  "提出元の事業所へ照会するか、未回答のまま範囲で記述するかのご決定により、"
  "調査の確定値化と、結果報告書・集計分析報告書・計画素案の"
  "該当箇所の更新が動きます。照会の文面は既に作成しています。" % len(JUDAI))
_K8 = [c for c in CHECK if c[0] == 8]
if _K8:
    NOTE("決まらない場合の当方の扱い：%s" % (KITEI.get(8) or "―"))

# ============================================================ 第5節
H1("第5節　広域連合又は構成3町で追加確認・資料提供が必要な事項")

H2("1　資料のご提供をお願いしたいもの")
P("資料提供依頼は全%d件で、うち%d件が未受領です。"
  "優先度の高いものから示します。"
  "**算定そのものが止まるものはありません**が、"
  "届かない場合は現に確認できている値で確定し、"
  "出所と制約を計画本文に注記することになります。"
  % (len(LACK), len(LACK_MACHI)))
_ORD = {"最高": 0, "高": 1, "中": 2, "低": 3}
_LK = sorted(LACK_MACHI, key=lambda x: (_ORD.get(x[7], 9), str(x[0])))
CAP("資料のご提供をお願いしたいもの")
TBL(["優先", "資料", "何のために要るか", "入手先", "希望時期"],
    [[x[7], x[2], x[4], x[5], x[6]] for x in _LK],
    [1.2, 6.0, 5.4, 2.6, 2.0], center={0, 4}, keep=False)

H2("2　構成3町にお伺いする事項")
_KOBETSU = {}
for c in MIKETSU:
    for t in TOWNS:
        if t in c[6]:
            _KOBETSU.setdefault(t, []).append(c)
P("町ごとにお伺いする事項です（%s）。"
  "町ごとの個別協議（令和8年10月中旬から下旬）でお伺いする予定です。"
  % "／".join("%s%d件" % (t, len(_KOBETSU.get(t, []))) for t in TOWNS))
CAP("構成3町にお伺いする事項")
TBL(["町", "確認事項", "表題", "期限"],
    [[t, "No.%d" % c[0], c[3], c[8]]
     for t in TOWNS for c in _KOBETSU.get(t, [])],
    [1.8, 1.8, 11.0, 1.6], center={0, 1, 3}, first_bold=True, keep=False)
NOTE("3町共通の事項は、3町合同の意見交換会（令和8年10月上旬）の資料に"
     "掲げています。")

H2("3　調査の提出元の事業所へのご照会")
P("点検で見つかった事項%d件のうち、"
  "確認主体が提出元の事業所であるものが%d件です。"
  "事業所への照会は広域連合を経由してお願いしたいものです。"
  "照会の文面は作成済みで、未回答時の扱いのご決定（確認事項No.8）を"
  "待って発出します。"
  % (len(TENKEN_JIKO), len(TENKEN_JIGYOSHO)))
CAP("確認主体の区分")
TBL(["確認主体", "件数", "内容"],
    [["提出元の事業所（広域連合を経由）", "%d件" % len(TENKEN_JIGYOSHO),
      "記入の誤り・記入漏れ・重複の疑いの確認"],
     ["広域連合", "%d件" % len(TENKEN_HACCHU),
      "特別養護老人ホームの入所申込者名簿、未回答施設への照会、"
      "指定事業所台帳、利用者票の所在地区の記入形式"],
     ["受託者", "%d件" % len(TENKEN_JUTAKU),
      "記録にとどめるもの、次期の調査設計に向けた申し送り"]],
    [5.0, 1.6, 10.6], center={1}, first_bold=True)

# ============================================================ 第6節
H1("第6節　今後の作成・協議スケジュール")

H2("1　月別の工程")
def _tsuki(label):
    """「令和8年10月」から通算月（令和の年×12＋月）を返す。読めなければ None。"""
    m = re.match(r"令和(\d+)年(\d+)月", str(label))
    return (int(m.group(1)) * 12 + int(m.group(2))) if m else None


_KARA = _tsuki("令和8年10月")
_MN = [m for m in MONTH
       if _tsuki(m[0]) is not None and _tsuki(m[0]) >= _KARA]
CAP("令和8年10月以降の工程")
TBL(["時期", "内容", "成果物"],
    [[m[0], m[1],
      "／".join(x.lstrip("・") for x in str(m[3]).split("\n") if x.strip())]
     for m in _MN],
    [2.6, 6.6, 8.0], first_bold=True, keep=False)
NOTE("令和9年度の予算編成に用いる見込量は、"
     "美瑛町の予算編成の工程に合わせて前倒しし、"
     "令和8年9月26日（第1次概算）・令和8年10月30日（第2次概算）・"
     "令和8年11月13日（予算編成用）に提示します。"
     "契約期間は令和9年3月26日までです。")

H2("2　令和8年10月30日までの段取り")
P("サービス見込量 第2次概算（令和8年10月30日）までの段取りです。"
  "**主体の欄が「発注者」「3町」であるものが、"
  "広域連合又は構成3町にお願いしたい事項です。**")
CAP("第2次概算までの段取り")
TBL(["時期", "事項", "内容", "主体"],
    [[x[0], x[1], _oki(x[2]), x[3]] for x in DANDORI],
    [2.0, 3.6, 8.4, 3.2], center={0}, first_bold=False)

H2("3　見込量・給付費・保険料の提示の見通し")
CAP("見込量・給付費・保険料の提示")
TBL(["時期", "提示するもの", "前提"],
    [["令和8年9月26日", "第1次概算（提示済み）",
      "算定上の月額基準額6,436円"],
     ["令和8年10月30日", "第2次概算",
      "受託者の側で確定できるものを織り込んだもの。"
      "算定上の月額基準額%s円・保険料基準額%s円"
      % (_n(GETSU, 2), _n(KIJUN_GAKU))],
     ["令和8年11月13日", "予算編成用",
      "第2次概算を基礎とし、ご決定・ご提供をいただいた分を織り込む"],
     ["令和8年12月", "計画素案（確定版）・保険料算定表",
      "基金残高・収納率・所得段階別被保険者数の受領と、"
      "国の政令・告示が前提"]],
    [2.8, 4.8, 9.6], first_bold=True)
NOTE("第1次概算として提示した版は、提示した時点の値を残すため"
     "別に残しています。第2次概算の資料に両者の差分を掲げています。")

# ============================================================ 第7節
H1("第7節　次回協議までのそれぞれの対応事項")
P("第6節2の段取りを、主体ごとに並べ替えたものです。")

_BY = {"受託者": [], "発注者": [], "3町": []}
for x in DANDORI:
    for k in _BY:
        if k in x[3]:
            _BY[k].append(x)

H2("1　受託者（ビズアップ公共コンサルティング）の対応事項")
for x in _BY["受託者"]:
    BUL("**%s　%s**　%s" % (x[0], x[1], _oki(x[2])))
BUL("**随時　ご照会への回答**　"
    "本資料のほか、ご決定・ご提供をいただいた事項は"
    "業務工程管理表に反映し、次回の協議でお示しします。")

H2("2　広域連合の対応事項")
for x in _BY["発注者"]:
    BUL("**%s　%s**　%s" % (x[0], x[1], _oki(x[2])))
BUL("**10月上旬　確認事項No.8のご決定**　"
    "点検で見つかった事項の重大%d件について、"
    "提出元の事業所へ照会するか、未回答のまま範囲で記述するかのご決定。"
    % len(JUDAI))

H2("3　構成3町の対応事項")
for x in _BY["3町"]:
    BUL("**%s　%s**　%s" % (x[0], x[1], _oki(x[2])))
BUL("**10月中旬　町ごとの確認事項へのご回答**　"
    "%s。施設整備・サービス提供の方針は必要利用定員総数に直結します。"
    % "／".join("%s%d件" % (t, len(_KOBETSU.get(t, []))) for t in TOWNS))

# ============================================================ 第8節
H1("第8節　未確定事項の管理資料")
P("未確定事項は、**新たな台帳を起こさず、既にお送りしている資料で"
  "管理しています。** 所在は次のとおりです。"
  "確認主体・期限・反映先は、いずれも項目ごとに掲げています。")

CAP("未確定事項の管理資料の所在")
TBL(["資料", "該当箇所", "管理している内容", "件数"],
    [["第10期計画_業務工程管理表.xlsx",
      "03 確認事項一覧",
      "**業務全体の確認事項（正の台帳）。** "
      "表題・内容・止めている成果物・確認先・状態・期限・"
      "決着しない場合の当方の扱い",
      "全%d件／未決%d件" % (len(CHECK), len(MIKETSU))],
     ["同上", "06 資料提供依頼",
      "資料の名称・何のために要るか・入手先・希望時期・優先度・状態",
      "全%d件／未受領%d件" % (len(LACK), len(LACK_MACHI))],
     ["同上", "02 月別工程管理",
      "月ごとの作業内容・成果物・状況・状態",
      "%d区分" % len(MONTH)],
     ["第10期計画_実施済み3調査の受領点検と集計.xlsx",
      "02 点検で見つかった事項",
      "**調査に固有の未確定事項。** 重要度・調査・対象・内容・"
      "取扱い・**確認主体・確認の期限・反映先（計画素案）**",
      "%d件（うち重大%d件）" % (len(TENKEN_JIKO), len(JUDAI))],
     ["同上", "09 計画本文への反映方針",
      "計画素案の該当箇所・反映する内容・前提となる確認・時期・状態",
      "%d件" % len(HANEI_HOSHIN)],
     ["第10期計画_実施済み調査_結果報告書.docx",
      "第8章第3節",
      "本報告書の限界と、取扱いを確定していない事項の"
      "確定までの扱い（範囲で記述・参考として記述・使用しない）",
      "限界%d件／未確定%d件" % (GENKAI_N, len(MIKAKUTEI))],
     ["第10期計画_計画素案の別管理表.xlsx", "確認事項の管理",
      "計画素案の章節から引く索引。"
      "業務工程管理表の確認事項番号を注記している",
      "A%s件・B%s件" % (BEK_A, BEK_B)]],
    [5.2, 3.0, 7.0, 2.6], first_bold=True, keep=False)

H2("ご覧いただく順序のご提案")
BUL("**まず** 業務工程管理表 03 確認事項一覧の「期限」が"
    "令和8年10月のものをご覧ください。"
    "「回答」の欄に、決着しない場合の当方の扱いを記しています。")
BUL("**次に** 本資料 第4節2（保険料の月額を動かすご判断）と"
    "第4節3（法定記載事項に直結するご判断）をご覧ください。"
    "ご決定の影響が大きい順に並べています。")
BUL("**調査に固有のもの**は、実施済み3調査の受領点検と集計 "
    "02 点検で見つかった事項の「重要度」が「重大」のものをご覧ください。"
    "確認主体・期限・反映先を列として持っています。")

# ============================================================ 自己点検
_ALLP = []
for _el in doc.element.body.iter(qn("w:p")):
    _ALLP.append("".join(t.text or "" for t in _el.iter(qn("w:t"))))
_ALL = "".join(_ALLP)

chk(1, "お尋ねの5項目に対応する節があること",
    "第2節から第7節＋管理資料（第8節）",
    all(("第%d節" % i) in _ALL for i in range(1, 9)))
chk(2, "確認事項を業務工程管理表から読んでいること",
    "全%d件・未決%d件" % (len(CHECK), len(MIKETSU)),
    len(CHECK) > 100 and len(MIKETSU) > 0)
chk(3, "確認事項の番号に欠番・重複がないこと",
    "No.1〜No.%d" % max(c[0] for c in CHECK),
    sorted(c[0] for c in CHECK) == list(range(1, len(CHECK) + 1)))
_KN = [c[0] for c in MIKETSU if not KITEI.get(c[0])]
chk(4, "未決の全件に決まらない場合の扱いがあること",
    "扱いのないもの %s" % ("なし" if not _KN else str(_KN)), not _KN)
chk(5, "結果報告書の未確定事項を同報告書から読んでいること",
    "%d件" % len(MIKAKUTEI), len(MIKAKUTEI) >= 9)
chk(6, "未確定事項の全件に確認主体が結び付いていること",
    "確認主体が「―」のもの %d件"
    % sum(1 for x in _MK_ROWS if x[2] == "―"),
    all(x[2] != "―" for x in _MK_ROWS))
chk(7, "点検で見つかった事項を受領点検から読んでいること",
    "%d件（重大%d件）" % (len(TENKEN_JIKO), len(JUDAI)),
    len(TENKEN_JIKO) >= 30 and len(JUDAI) >= 1)
chk(8, "確認主体の区分の和が全件に一致すること",
    "事業所%d＋広域連合%d＋受託者%d＝%d"
    % (len(TENKEN_JIGYOSHO), len(TENKEN_HACCHU), len(TENKEN_JUTAKU),
       len(TENKEN_JIKO)),
    (len(TENKEN_JIGYOSHO) + len(TENKEN_HACCHU) + len(TENKEN_JUTAKU)
     == len(TENKEN_JIKO)))
chk(9, "計画本文への反映方針を受領点検から読んでいること",
    "%d件（うち確認を要しないもの%d件）"
    % (len(HANEI_HOSHIN), len(_CHAKUSHU)),
    len(HANEI_HOSHIN) > 0 and len(_CHAKUSHU) > 0)
_KANRYO_NG = [no for no, (s, w) in HANEI_KANRYO.items()
              if not _in_draft(s, w)]
chk(9.5, "反映を完了したとしたものが計画素案の該当する節にあること",
    "%d件（%s）"
    % (len(HANEI_KANRYO),
       "／".join("No.%d %s" % (no, s) for no, (s, _w)
                 in sorted(HANEI_KANRYO.items()))),
    DRAFT_SEC != {} and not _KANRYO_NG
    and {x[0] for x in _CHAKUSHU} == {str(no) for no in HANEI_KANRYO})
chk(10, "算定上の月額基準額が算定の値と一致すること",
    "%s円" % _n(GETSU), 6000 < GETSU < 7000 and _n(GETSU) in _ALL)
chk(11, "百円未満を四捨五入した基準額を示していること",
    "%s円" % _n(KIJUN_GAKU), _n(KIJUN_GAKU) in _ALL)
chk(12, "月額を動かす前提が据え置きの区分B・Cに一致すること",
    "B %d件・C %d件＝掲げた%d件"
    % (sum(1 for v in KUBUN.values() if v == "B"),
       sum(1 for v in KUBUN.values() if v == "C"), len(_SU)),
    len(_SU) == sum(1 for v in KUBUN.values() if v in ("B", "C")))
chk(13, "資料提供依頼を業務工程管理表から読んでいること",
    "全%d件・未受領%d件" % (len(LACK), len(LACK_MACHI)),
    len(LACK) > 0 and len(LACK_MACHI) > 0)
chk(14, "町ごとの確認事項が3町とも拾えていること",
    "／".join("%s%d件" % (t, len(_KOBETSU.get(t, []))) for t in TOWNS),
    all(len(_KOBETSU.get(t, [])) > 0 for t in TOWNS))
chk(15, "工程を業務工程管理表から読んでいること",
    "令和8年10月以降 %d区分" % len(_MN), len(_MN) >= 5)
chk(16, "段取りを第2次概算の組立てから読んでいること",
    "%d件（受託者%d・発注者%d・3町%d）"
    % (len(DANDORI), len(_BY["受託者"]), len(_BY["発注者"]),
       len(_BY["3町"])),
    len(DANDORI) > 0 and all(len(v) > 0 for v in _BY.values()))
chk(17, "仮置きにより計画素案へ反映した件数が一致すること",
    "%d件" % len(SUSUMETA),
    len(SUSUMETA) > 0 and ("%d件" % len(SUSUMETA)) in _ALL)
chk(18, "業務内容別の進捗の基準日が本資料の基準日と一致すること",
    "%s" % KIJUNBI, DP.KIJUNBI == KIJUNBI and KIJUNBI in _ALL)

# 守るべき制約
_NAIBU = ("固定値", "実物から", "スクリプト", "モジュール", "リポジトリ",
          "runpy", ".py", "再実行", "書き写して", "章節ごとに",
          "読んで判定", "終了コード", "内部管理用")
_H_NAIBU = [w for w in _NAIBU if w in _ALL]
chk(19, "受託者の内部の仕組み・作業経過の語がないこと",
    "当たった語 %s" % ("なし" if not _H_NAIBU else "／".join(_H_NAIBU)),
    not _H_NAIBU)
_KINSHI = ("に由来する", "と整合する", "1件も", "有意差がないため",
           "全国トップ級")
_H_KIN = [w for w in _KINSHI if w in _ALL]
chk(20, "禁止表現がないこと",
    "当たった語 %s" % ("なし" if not _H_KIN else "／".join(_H_KIN)),
    not _H_KIN)
_PI = [re.compile(r"0\d{1,4}-\d{1,4}-\d{3,4}"),
       re.compile(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}")]
_H_PI = [p.pattern for p in _PI if p.search(_ALL)]
chk(21, "電話番号・電子メールアドレスの形がないこと",
    "当たった形 %s" % ("なし" if not _H_PI else "／".join(_H_PI)),
    not _H_PI)
_OK_CH = re.compile(
    r"[ぁ-んァ-ヴ一-龥々ー０-９0-9A-Za-z、。・（）「」【】〔〕〜％"
    r"：；，．／±＋▲→①②③④⑤⑥⑦⑧⑨⑩―※＝×［］_…"
    r"\s　\-,.%()\[\]:;/'\"！？]")
_NG_CH = sorted({c for c in _ALL if not _OK_CH.match(c)})
chk(22, "許容する文字の外の文字がないこと",
    "当たった文字 %s" % ("なし" if not _NG_CH else "".join(_NG_CH)),
    not _NG_CH)
_HANKAKU = [w for w in ("+1", "-1", "+2", "-2") if w in _ALL]
chk(23, "差の表記に半角の符号を用いていないこと",
    "当たった表記 %s" % ("なし" if not _HANKAKU else "／".join(_HANKAKU)),
    not _HANKAKU)

chk(24, "結果報告書の限界の件数を同報告書から数えていること",
    "%d件" % GENKAI_N, GENKAI_N >= 10)
chk(25, "別管理表の件数を同表から数えていること",
    "A%s件・B%s件" % (BEK_A, BEK_B),
    BEK_A is not None and BEK_A > 0 and BEK_B > 0)

# ------------------------------------------------------------ 点検の結果
H1("（参考）　本資料の自己点検", brk=True)
P("本資料の数値と記述が、業務工程管理表・サービス見込量の算定及び"
  "実施済み調査 結果報告書と合っていることを機械で確かめた結果です。"
  "1件でも不適合があれば作成を止めます。")
CAP("自己点検の結果（%d件）" % len(CHECKS))
TBL(["No.", "確かめたこと", "結果", "判定"],
    [[str(n), n2, k, j] for n, n2, k, j in CHECKS],
    [1.2, 7.6, 6.4, 2.0], center={0, 3}, keep=False)

docx_fix.fix(doc)
doc.save(OUT)

_NG = [c for c in CHECKS if c[3] != "適合"]
print("出力 %s" % os.path.basename(OUT))
print("  節8＋参考1／表%d点／自己点検%d件（不適合%d件）"
      % (TBLNO[0], len(CHECKS), len(_NG)))
print("  確認事項 全%d件・未決%d件／資料提供依頼 未受領%d件"
      % (len(CHECK), len(MIKETSU), len(LACK_MACHI)))
print("  調査の未確定事項%d件／点検で見つかった事項%d件（重大%d件）"
      % (len(MIKAKUTEI), len(TENKEN_JIKO), len(JUDAI)))
print("  算定上の月額%s円／保険料基準額%s円"
      % (_n(GETSU, 2), _n(KIJUN_GAKU)))
for n, n2, k, j in CHECKS:
    if j != "適合":
        print("  不適合 %d %s : %s" % (n, n2, k))
if _NG:
    sys.exit(1)
