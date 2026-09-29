# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画　公表版の生成.

令和8年9月29日のご指示
  「協議用素案と合わせて、実際に公表ベースで使う用の素案を作成して下さい」

━━ 2つの版の切り分け ━━

  協議用素案　　打合せでご判断いただくための版。
                当方の分析（複数の置き方の対比、感度、推計の前提）を残し、
                未確定箇所は［要協議］［要確認］［要内訳］で示す。

  公表版　　　　住民・委員がお読みになるための版。
                協議のための分析を落とし、
                サービスの内容と用語の解説を加える。

**内容の正本は協議用素案である。** 本スクリプトは協議用素案（docx）を読み、
公表用に組み替えたものを別ファイルとして書き出す。
数値・本文を別に持たないため、協議用素案を改めれば公表版も追随する。

━━ 公表版で行うこと ━━

  1 表紙と柱を公表用に改める
  2 「協議」と判定した分析のまとまりを落とす
    （判定は build_soan_kisai_taihi.py の DOKUJI から読む。台帳を二重に持たない）
  3 章の見出しに帯を付け、節・項の前後に余白を置く（視認性）
  4 巻末に「介護サービスの内容」と「用語の解説」を加える
  5 目次を組み替え、ページ番号を組版結果から求める
  6 未確定箇所（［要協議］等）が残っている間は表紙に「体裁確認版」と記す

出力
  output/第10期介護保険事業計画_公表版.docx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import ast
import copy
import io
import json
import os
import re
import subprocess
import sys
import tempfile

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.table import Table
from docx.text.paragraph import Paragraph

import docx_fix
import repo_paths as RP
import data_service_desc as SD

if hasattr(sys.stdout, "buffer"):      # 他から読み込まれた場合は差し替えない
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

OUT = RP.ROOT + "/output/第10期介護保険事業計画_公表版.docx"
PAGEMAP = RP.ROOT + "/output/_toc_pages_public.json"
FONT = "游ゴシック"
TEXTW = 16.6
HEAD_SHD = "DCE6F1"
BAND = "1F3864"

CHECKS = []


def chk(no, naiyo, kekka, ok):
    CHECKS.append((no, naiyo, kekka, ok))
    return ok


# ============================================================ 判定を読む
def _assign(path, *names):
    src = open(path, encoding="utf-8").read()
    out = {}
    for node in ast.parse(src).body:
        if (isinstance(node, ast.Assign)
                and getattr(node.targets[0], "id", None) in names):
            out[node.targets[0].id] = ast.literal_eval(node.value)
    return out


_TAIHI = _assign(RP.ROOT + "/build_soan_kisai_taihi.py", "DOKUJI")
DOKUJI = _TAIHI["DOKUJI"]
# 公表版から落とすのは、判定が「協議」であるもの（当方の分析）。
OTOSU = {k[2] for k, v in DOKUJI.items() if v[1] == "協議"}

# SVC は内包表記で組み立てられているため、算定を読み込んで得る。
import contextlib as _ctx
import runpy as _runpy
_buf = io.StringIO()
with _ctx.redirect_stdout(_buf):
    _S = _runpy.run_path(RP.ROOT + "/build_mikomiryo_santei.py")
SVC = _S["SVC"]

PH = re.compile(r"［(要協議|要確認|要内訳)］")
SHO = re.compile(r"^(第\d+章|資料\d+|資料編)")
SETSU = re.compile(r"^(第\d+節|基本目標\d)")
LEADP = re.compile(r"^【(.+)】$")

# ============================================================ 読み込む
doc = Document(RP.DRAFT)
body = doc.element.body


def _paras():
    return [Paragraph(e, doc) for e in body.iterchildren()
            if e.tag.split("}")[-1] == "p"]


def _text(e):
    return Paragraph(e, doc).text.strip() if e.tag.split("}")[-1] == "p" else ""


def _is_toc(e):
    return e.tag.split("}")[-1] == "p" and "\t" in _text(e) \
        and e.find(qn("w:hyperlink")) is not None


# ============================================================ 1 表紙
_ALLTXT0 = "\n".join(Paragraph(e, doc).text for e in body.iterchildren()
                     if e.tag.split("}")[-1] == "p")
_PH_N = len(PH.findall(_ALLTXT0))
for _e in body.iterchildren():
    if _e.tag.split("}")[-1] == "tbl":
        for _r in Table(_e, doc).rows:
            for _c in _r.cells:
                _PH_N += len(PH.findall(_c.text))

# 計画として議決・決定を経る前であるため、表題は「（案）」とする。
HYOSHI = "第10期介護保険事業計画（案）"
# 未確定箇所（［要協議］等）が残る間は、公表できる版ではないことを表紙に示す。
SUBTITLE = "公表版" if _PH_N == 0 else "公表版（体裁確認用）"

_hyo_done = 0
for p in _paras()[:14]:
    t = p.text.strip()
    if "協議用素案" in t:
        for r in p.runs:
            r.text = ""
        p.runs[0].text = SUBTITLE if p.runs else ""
        _hyo_done += 1
    elif t.startswith("第10期介護保険事業計画") and len(t) < 30:
        for r in p.runs:
            r.text = ""
        p.runs[0].text = HYOSHI
        _hyo_done += 1

# ============================================================ 2 分析を落とす
def _block_of(idx, elems):
    """【　】の見出しから、次の見出し（【　】・章・節・項）までを返す。"""
    out = [elems[idx]]
    for j in range(idx + 1, len(elems)):
        e = elems[j]
        tag = e.tag.split("}")[-1]
        # 本文の段落と表だけを対象とする。
        # 本文末尾の w:sectPr まで巻き込むと節（ヘッダー・フッター・
        # ページ番号）が失われるため、段落・表以外で必ず止める。
        if tag not in ("p", "tbl"):
            break
        if tag == "p":
            t = _text(e)
            if LEADP.match(t) or SHO.match(t) or SETSU.match(t):
                break
            if Paragraph(e, doc).style.name.startswith("Heading"):
                break
            pPr = e.find(qn("w:pPr"))
            if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
                break          # 節の区切りを持つ段落は落とさない
        out.append(e)
    return out


_removed = []
while True:
    elems = list(body.iterchildren())
    hit = None
    for i, e in enumerate(elems):
        if e.tag.split("}")[-1] != "p":
            continue
        m = LEADP.match(_text(e))
        if m and m.group(1) in OTOSU:
            hit = (i, m.group(1))
            break
    if hit is None:
        break
    i, nm = hit
    for e in _block_of(i, elems):
        body.remove(e)
    _removed.append(nm)

# ============================================================ 3 視認性
def _band(p):
    """章の見出しに上下の罫線と地色を付ける。"""
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    for side, sz in (("top", "18"), ("bottom", "18")):
        b = OxmlElement("w:" + side)
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), sz)
        b.set(qn("w:space"), "4")
        b.set(qn("w:color"), BAND)
        bdr.append(b)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "F2F5FA")
    # pBdr は spacing・ind より前、shd は pBdr の次に置く。
    ref = pPr.find(qn("w:spacing"))
    if ref is None:
        pPr.append(bdr)
        pPr.append(shd)
    else:
        ref.addprevious(bdr)
        bdr.addnext(shd)


_band_n = 0
for p in _paras():
    if p.style.name == "Heading 1" and "\t" not in p.text:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(10)
        _band(p)
        _band_n += 1
    elif p.style.name == "Heading 2" and "\t" not in p.text:
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(6)
    elif p.style.name == "Heading 3" and "\t" not in p.text:
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)

# ============================================================ 4 巻末を加える
HEADINGS = []                     # (level, text, bookmark)
for p in _paras():
    if "\t" in p.text:
        continue
    st = p.style.name
    if st in ("Heading 1", "Heading 2"):
        bm = p._p.find(qn("w:bookmarkStart"))
        HEADINGS.append((1 if st == "Heading 1" else 2, p.text.strip(),
                         bm.get(qn("w:name")) if bm is not None else None))

_BMN = [1000]


def _bookmark(p, name):
    _BMN[0] += 1
    bid = str(_BMN[0])
    s = OxmlElement("w:bookmarkStart")
    s.set(qn("w:id"), bid)
    s.set(qn("w:name"), name)
    e = OxmlElement("w:bookmarkEnd")
    e.set(qn("w:id"), bid)
    pPr = p._p.find(qn("w:pPr"))
    if pPr is None:
        p._p.insert(0, s)
    else:
        pPr.addnext(s)
    p._p.append(e)


def _fmt(run, size, bold=False, color=None):
    run.font.size = Pt(size)
    run.font.name = FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    run.bold = bold
    if color:
        run.font.color.rgb = color


def AP(text, size=10.5, bold=False, align=None, after=4, before=0):
    p = doc.add_paragraph()
    _fmt(p.add_run(text), size, bold)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.line_spacing = 1.08
    if align is not None:
        p.alignment = align
    return p


def AH(text, level):
    p = doc.add_paragraph(style="Heading %d" % level)
    _fmt(p.add_run(text), 15 if level == 1 else 12.5)
    p.paragraph_format.space_before = Pt(14 if level == 1 else 12)
    p.paragraph_format.space_after = Pt(10 if level == 1 else 6)
    if level == 1:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _band(p)
    name = "_Pub_%03d" % (len(HEADINGS) + 1)
    HEADINGS.append((level, text, name))
    _bookmark(p, name)
    return p


def _shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    va = tcPr.find(qn("w:vAlign"))
    if va is None:
        tcPr.append(shd)
    else:
        va.addprevious(shd)


def _row_flags(row, header=False):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))
    if header:
        trPr.append(OxmlElement("w:tblHeader"))


def ATBL(head, rows, widths, size=9):
    t = doc.add_table(rows=1, cols=len(head))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, hh in enumerate(head):
        c = t.rows[0].cells[i]
        c.text = ""
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _fmt(c.paragraphs[0].add_run(hh), size, True)
        _shade(c, HEAD_SHD)
    _row_flags(t.rows[0], header=True)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            _fmt(cells[i].paragraphs[0].add_run("" if v is None else str(v)),
                 size)
        _row_flags(t.rows[-1])
    tot = sum(widths)
    if tot > TEXTW:
        widths = [w * TEXTW / tot for w in widths]
    for i, w in enumerate(widths):
        for row in t.rows:
            row.cells[i].width = Cm(w)
    return t


doc.add_page_break()
AH("参考　介護サービスの内容", 1)
AP("本計画に掲げる介護サービスの内容は次のとおりです。"
   "介護保険法及び国が定める基準によるものであり、"
   "実際に利用できるサービスは要介護度により異なります。")

_KU = [("施設サービス", "施設サービス"),
       ("居住系サービス", "居住系サービス"),
       ("在宅サービス", "居宅サービス・地域密着型サービス等")]
_naikou = []
for key, midashi in _KU:
    AH(midashi, 2)
    rows = []
    for s in SVC:
        if not s.startswith(key + " "):
            continue
        nm = s.split(" ", 1)[1]
        d = SD.DESC.get(s)
        if d is None:
            _naikou.append(s)
            continue
        rows.append([nm, d])
    ATBL(["サービスの種類", "内容"], rows, [5.0, 11.6])

doc.add_page_break()
AH("用語の解説", 1)
AP("本計画で用いる主な用語の意味は次のとおりです。"
   "五十音順ではなく、計画に現れる順に近い形で並べています。")
ATBL(["用語", "意味"], [[k, v] for k, v in SD.YOGO], [4.4, 12.2])

# ============================================================ 5 目次
def _fldchar(kind):
    e = OxmlElement("w:fldChar")
    e.set(qn("w:fldCharType"), kind)
    return e


def _instr(txt):
    e = OxmlElement("w:instrText")
    e.set(qn("xml:space"), "preserve")
    e.text = txt
    return e


try:
    _PAGES = json.load(open(PAGEMAP, encoding="utf-8"))
except Exception:
    _PAGES = {}


def _toc_line(level, text, name):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2 if level == 1 else 0)
    p.paragraph_format.line_spacing = 1.0
    tabs = p.paragraph_format.tab_stops
    from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
    tabs.add_tab_stop(Cm(TEXTW), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    hl = OxmlElement("w:hyperlink")
    hl.set(qn("w:anchor"), name or "")
    p._p.append(hl)

    def _run(txt=None, child=None, bold=False, size=10):
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        rf = OxmlElement("w:rFonts")
        for a in ("w:ascii", "w:hAnsi", "w:eastAsia"):
            rf.set(qn(a), FONT)
        rPr.append(rf)
        if bold:
            rPr.append(OxmlElement("w:b"))
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(size * 2)))
        rPr.append(sz)
        r.append(rPr)
        if txt is not None:
            t = OxmlElement("w:t")
            t.set(qn("xml:space"), "preserve")
            t.text = txt
            r.append(t)
        if child is not None:
            r.append(child)
        hl.append(r)
        return r

    _run(txt=("" if level == 1 else "　　") + text,
         bold=(level == 1), size=10.5 if level == 1 else 10)
    _run(child=OxmlElement("w:tab"))
    _run(child=_fldchar("begin"))
    _run(child=_instr(" PAGEREF %s \\h " % (name or "")))
    _run(child=_fldchar("separate"))
    _run(txt=str(_PAGES.get(name, "－")), bold=(level == 1))
    _run(child=_fldchar("end"))
    return p


# 既存の目次の行を取り除き、作り直す。
_old_toc = [e for e in body.iterchildren() if _is_toc(e)]
_anchor = _old_toc[0] if _old_toc else None
for e in _old_toc[1:]:
    body.remove(e)
_new = [_toc_line(lv, tx, nm) for lv, tx, nm in HEADINGS]
if _anchor is not None:
    for p in _new:
        _anchor.addprevious(p._p)
    body.remove(_anchor)

# ============================================================ 保存
_st = doc.settings.element
_uf = _st.find(qn("w:updateFields"))
if _uf is None:
    _uf = OxmlElement("w:updateFields")
    _after = None
    for _tag in ("w:compat", "w:docVars", "w:rsids", "w:themeFontLang"):
        _after = _st.find(qn(_tag))
        if _after is not None:
            break
    if _after is None:
        _st.append(_uf)
    else:
        _after.addprevious(_uf)
_uf.set(qn("w:val"), "true")
_zoom = _st.find(qn("w:zoom"))
if _zoom is not None and _zoom.get(qn("w:percent")) is None:
    _zoom.set(qn("w:percent"), "100")

docx_fix.fix(doc)
doc.save(OUT)

# ============================================================ 自己点検
_pub = Document(OUT)
_ptxt = "\n".join(p.text for p in _pub.paragraphs)
for _t in _pub.tables:
    for _r in _t.rows:
        for _c in _r.cells:
            _ptxt += "\n" + _c.text
_pleads = {m.group(1) for m in
           (LEADP.match(p.text.strip()) for p in _pub.paragraphs) if m}

chk(0, "節（ヘッダー・フッター・ページ番号）が失われていないこと",
    "協議用素案%d節／公表版%d節"
    % (len(Document(RP.DRAFT).sections), len(_pub.sections)),
    len(_pub.sections) == len(Document(RP.DRAFT).sections))
chk(1, "協議用素案を読めていること",
    "%d段落%d表" % (len(Document(RP.DRAFT).paragraphs),
                   len(Document(RP.DRAFT).tables)),
    len(Document(RP.DRAFT).paragraphs) > 500)
chk(2, "表紙を公表用に改めたこと",
    "差し替え%d箇所／「協議用素案」%d件" % (_hyo_done, _ptxt.count("協議用素案")),
    _hyo_done >= 1 and "協議用素案" not in _ptxt)
chk(3, "「協議」と判定した分析を落としたこと",
    "落とした%d件／判定%d件" % (len(_removed), len(OTOSU)),
    len(_removed) == len(OTOSU))
chk(4, "落とした見出しが公表版に残っていないこと",
    "、".join(sorted(OTOSU & _pleads)) or "残っていない",
    not (OTOSU & _pleads))
chk(5, "落としていない【　】の見出しは残っていること",
    "%d件" % len(_pleads), len(_pleads) >= 40)
chk(6, "章の見出しに帯を付けたこと", "%d件" % _band_n, _band_n >= 7)
chk(7, "サービスの内容が全区分そろっていること",
    "説明のない区分 %s" % ("、".join(_naikou) or "なし"), not _naikou)
chk(8, "サービスの内容の説明が本文にあること",
    "%d件" % sum(1 for s in SVC if SD.DESC.get(s, "") and
                 SD.DESC[s][:12] in _ptxt),
    sum(1 for s in SVC if SD.DESC.get(s, "") and SD.DESC[s][:12] in _ptxt)
    >= len(SVC) - 1)
chk(9, "用語の解説があること", "%d語" % len(SD.YOGO),
    len(SD.YOGO) >= 25 and all(k in _ptxt for k, _v in SD.YOGO))
chk(10, "目次を作り直したこと（章・節の数と一致）",
    "目次%d行／見出し%d件"
    % (sum(1 for p in _pub.paragraphs if "\t" in p.text), len(HEADINGS)),
    sum(1 for p in _pub.paragraphs if "\t" in p.text) == len(HEADINGS))
_ph_pub = len(PH.findall(_ptxt))
chk(11, "未確定箇所の数を数えていること",
    "協議用素案%d箇所／公表版%d箇所" % (_PH_N, _ph_pub), _ph_pub <= _PH_N)
chk(12, "未確定箇所が残る間は表紙に体裁確認用と記すこと",
    "体裁確認用の記載 %s" % ("あり" if "体裁確認用" in _ptxt else "なし"),
    (_ph_pub > 0) == ("体裁確認用" in _ptxt))
_NAIBU = ("固定値", "runpy", ".py", "再実行", "スクリプト", "章節ごとに",
          "書き写して", "判定しています", "受託者", "本素案", "協議用",
          "修正指示書", "確認事項No")
chk(13, "受託者の内部の仕組みの語が現れないこと",
    "、".join(w for w in _NAIBU if w in _ptxt) or "現れない",
    not any(w in _ptxt for w in _NAIBU))
_KIN = ("に由来する", "1件も", "全国トップ級", "有意差がないため")
chk(14, "禁止表現が現れないこと",
    "、".join(w for w in _KIN if w in _ptxt) or "現れない",
    not any(w in _ptxt for w in _KIN))
_TADAN = ("北塩原村", "浜田地区", "浜田圏域", "川崎町", "金ヶ崎町",
          "空知中部", "日高中部", "後志広域", "京極町")
chk(15, "他の団体の固有名称が現れないこと",
    "、".join(w for w in _TADAN if w in _ptxt) or "現れない",
    not any(w in _ptxt for w in _TADAN))
chk(16, "個人情報の形（電話番号・メールアドレス）が現れないこと",
    "%d件" % (len(re.findall(r"0\d{1,4}-\d{1,4}-\d{3,4}", _ptxt))
              + len(re.findall(r"[\w.+-]+@[\w.-]+\.\w+", _ptxt))),
    not re.search(r"0\d{1,4}-\d{1,4}-\d{3,4}", _ptxt)
    and not re.search(r"[\w.+-]+@[\w.-]+\.\w+", _ptxt))
_OK_CH = re.compile(r"[ぁ-んァ-ヴ一-龥々ー０-９0-9A-Za-z、。・（）「」〜％　 \n]")
_zatsu = sorted({ch for v in list(SD.DESC.values())
                 + [a + b for a, b in SD.YOGO]
                 for ch in v if not _OK_CH.match(ch)})
chk(17.5, "サービスの内容と用語の解説に日本語以外の文字が混じっていないこと",
    "".join(_zatsu) or "混じっていない", not _zatsu)
chk(17, "表の見出し行が各ページで繰り返される設定であること",
    "%d表" % sum(1 for t in _pub.tables
                 if t.rows[0]._tr.find(qn("w:trPr")) is not None
                 and t.rows[0]._tr.find(qn("w:trPr")).find(
                     qn("w:tblHeader")) is not None),
    sum(1 for t in _pub.tables
        if t.rows[0]._tr.find(qn("w:trPr")) is not None
        and t.rows[0]._tr.find(qn("w:trPr")).find(qn("w:tblHeader"))
        is not None) >= len(_pub.tables) - 2)

NG = [c for c in CHECKS if not c[3]]
print("保存 %s" % OUT)
print("段落%d／表%d／図%d"
      % (len(_pub.paragraphs), len(_pub.tables), len(_pub.inline_shapes)))
print("落とした分析 %d件：%s" % (len(_removed), "、".join(_removed)))
print("加えたもの サービスの内容%d区分・用語の解説%d語"
      % (len(SD.DESC), len(SD.YOGO)))
print("未確定箇所 %d箇所（0になるまで表紙は体裁確認用）" % _ph_pub)
print("自己点検 %d件　不適合 %d件" % (len(CHECKS), len(NG)))
for c in NG:
    print("  不適合 No.%s %s → %s" % (c[0], c[1], c[2]))
if NG:
    sys.exit(1)
