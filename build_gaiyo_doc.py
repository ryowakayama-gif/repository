# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画　概要版の原稿の生成.

仕様書４（6）が求める「概要版原稿」である。
別冊「第10期計画_概要版の構成案」で定めた頁割（A4判8頁）に沿って組む。

━━ 概要版の位置付け ━━

  概要版は住民への周知資料であると同時に、
  令和9年1月のパブリックコメントで意見を求めるための資料である。
  後者の性格が構成を規定するため、意見を求めたい論点6件を頁ごとに示し、
  本編の該当頁を各頁の末尾に記す。

━━ 本編（計画素案）との関係 ━━

  **内容の正本は計画素案である。** 本原稿は計画素案（docx）の表と本文、
  及びサービス見込量の算定を読んで組む。数値を別に持たないため、
  計画素案を改めれば概要版も追随する。

  抜粋の方針は構成案 05シートのとおりである。
    ① 本編の文をそのまま採らない（基本理念・基本目標は本編と同一の表現とする）
    ② 数値は本編と一致させる
    ③ ［要協議］［要確認］は載せない（保険料は試算値として載せる）
    ④ 本編の頁を示す
    ⑤ 3町の記述の分量を揃える
    ⑥ 禁止表現を用いない

  未決定の達成目標は「検討中」と記す（構成案 06シート 差分⑤）。
  意見の提出先・期限・問合せ先は〔　〕の記入枠とする
  （個人情報及び連絡先を収録しない方針による）。

━━ 2版構成 ━━

  パブリックコメント用（第1版）と最終版（第2版）の2版を作成する。
  本原稿は第1版である。計画素案に未確定箇所が残っている間は
  表題を「（素案）の概要」とし、保険料を試算値として示す。
  未確定箇所が0になったときだけ「の概要」となる。

出力
  output/第10期介護保険事業計画_概要版.docx
  output/figures_gaiyo/gaiyo_kenniki.png   圏域の構成図（新規作図）
  output/figures_gaiyo/gaiyo_taikei.png    施策の体系図（新規作図）

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import ast
import contextlib
import io
import json
import os
import re
import runpy
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import rcParams
from matplotlib.patches import FancyBboxPatch

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.table import Table
from docx.text.paragraph import Paragraph

import docx_fix
import repo_paths as RP
import data_service_desc as SD
import data_zuhyo as DZ

if isinstance(sys.stdout, io.TextIOWrapper):
    sys.stdout.reconfigure(encoding="utf-8")

OUT = RP.ROOT + "/output/第10期介護保険事業計画_概要版.docx"
FIGD = RP.ROOT + "/output/figures_gaiyo"
FIGSRC = RP.ROOT + "/output/figures"
FONT = "游ゴシック"
NAVY = "1F3864"
HEAD_SHD = "DCE6F1"
RON_SHD = "FFF2CC"           # 意見を求める論点の枠
TEXTW = 21.0 - 1.7 * 2       # 本文幅 17.6cm

CHECKS = []


def chk(no, naiyo, kekka, ok):
    CHECKS.append((no, naiyo, kekka, ok))
    return ok


def _literal(path, name):
    src = open(path, encoding="utf-8").read()
    for node in ast.parse(src).body:
        if (isinstance(node, ast.Assign)
                and getattr(node.targets[0], "id", None) == name):
            return ast.literal_eval(node.value)
    raise RuntimeError("%s が見つからない（%s）" % (name, path))


# ============================================================ 構成案を読む
# 頁割と論点は構成案が持つ。本原稿はそこから読み、設計を二重に持たない。
PAGES = _literal(RP.ROOT + "/build_gaiyo_plan.py", "PAGES")
PG = {r[0]: r for r in PAGES}


def ronten(key):
    """構成案の論点の欄から（番号付きの題, かっこ書き）を返す。"""
    t = str(PG[key][5]).replace("\n", "")
    if t in ("―", "-", ""):
        return None
    m = re.match(r"^(論点\d+)　?(.*?)(（(.*)）)?$", t)
    if not m:
        return (t, "")
    return ("%s　%s" % (m.group(1), m.group(2)), m.group(4) or "")


# ============================================================ 計画素案を読む
src = Document(RP.DRAFT)
_SHO = re.compile(r"^(第\d+章|資料\d+|資料編)")
_SETSU = re.compile(r"^(第\d+節|基本目標\d)")

SEC_P = {}        # 章節 -> 段落の文のならび
SEC_T = {}        # 章節 -> 表（行×セルの文）のならび
TOCPG = {}        # 章・節の見出し -> 本編の頁
_cur = ["", ""]


def _key():
    return (_cur[0] + _cur[1]).replace("　", "").replace(" ", "")


for _e in src.element.body.iterchildren():
    _tag = _e.tag.split("}")[-1]
    if _tag == "p":
        _t = Paragraph(_e, src).text.strip()
        if "\t" in _t:                      # 目次の行から本編の頁を読む
            _nm, _, _pg = _t.rpartition("\t")
            if _pg.strip().isdigit():
                TOCPG[_nm.replace("　", " ").strip()] = int(_pg)
            continue
        if _SHO.match(_t):
            _cur[0] = _t.split("　")[0].split(" ")[0]
            _cur[1] = ""
        if _SETSU.match(_t):
            _cur[1] = _t.split("　")[0].split(" ")[0]
        SEC_P.setdefault(_key(), []).append(_t)
    elif _tag == "tbl":
        _tb = Table(_e, src)
        SEC_T.setdefault(_key(), []).append(
            [[c.text.strip() for c in r.cells] for r in _tb.rows])


def tbl(key, *head):
    """章節と見出し行の語で表を1つ選ぶ。位置では選ばない。"""
    cand = [t for t in SEC_T.get(key, [])
            if all(any(w in c for c in t[0]) for w in head)]
    if not cand:
        raise RuntimeError("表が見つからない：%s %s" % (key, head))
    return cand[0]


def jiten(key, *words):
    """素案の資料の行から出所の時点を読む。概要版に時点を固定値で書かない。

    資料の行は「資料：…（住民基本台帳　令和5年10月1日現在）」の形であり、
    かっこの中が出所と時点である。かっこの前は本編の出所の書き方
    （第9期計画の章節）であるため概要版には用いない。
    """
    for ln in SEC_P.get(key, []):
        if not re.match(r"^(資料|出典|単位)[：:]", ln) and "出典：" not in ln:
            continue
        if all(w in ln for w in words):
            m = re.search(r"（(.+?)）", ln)
            if m:
                return m.group(1).strip()
    raise RuntimeError("資料の行が見つからない：%s %s" % (key, words))


def sho_pages(sho):
    """章の本編の頁範囲（開始頁, 次の章の開始頁−1）を返す。"""
    nums = sorted((v, k) for k, v in TOCPG.items() if k.startswith(sho))
    if not nums:
        return None
    start = nums[0][0]
    later = [v for k, v in TOCPG.items()
             if re.match(r"^第(\d+)章", k)
             and int(re.match(r"^第(\d+)章", k).group(1))
             > int(re.match(r"^第(\d+)章", sho).group(1))]
    end = (min(later) - 1) if later else max(TOCPG.values())
    return (start, max(start, end))


# 未確定箇所は計画素案の実物から数える（第1版か最終版かを分ける）
PH = re.compile(r"［(要協議|要確認|要内訳)］")
_ALL = "\n".join(p.text for p in src.paragraphs)
for _t in src.tables:
    for _r in _t.rows:
        for _c in _r.cells:
            _ALL += "\n" + _c.text
PH_N = len(PH.findall(_ALL))


def noph(s, kari="検討中"):
    """［要協議］等を含む欄を「検討中」に置き換える（方針③・差分⑤）。"""
    s = str(s).strip()
    return kari if PH.search(s) or s in ("", "―") else s


# ============================================================ 算定を読む
_buf = io.StringIO()
with contextlib.redirect_stdout(_buf):
    _S = runpy.run_path(RP.ROOT + "/build_mikomiryo_santei.py")
Y3L = _S["Y3L"]
G_DO = _S["G_DO"]
KYUFU3 = _S["KYUFU3"]
MIKOMI = _S["MIKOMI"]
JISSEKI = _S["JISSEKI"]
NIN = _S["NIN"]
HIHO = _S["HIHO"]
Y3 = _S["Y3"]
GETSU = G_DO["月額"]
KIJUN = int(round(GETSU / 100.0)) * 100          # 百円未満四捨五入
SUSEI = _S["susei_haba"]()                       # 単価の趨勢を織り込んだ幅
# 報酬改定率1％あたりの月額は算定から求める。
KAITEI_1 = _S["gaku_kaitei"](0.01)["月額"] - GETSU
# 基金1億円の取崩しの効きは据え置きの一覧の記述から数字を読む。
_kk = [x for x in _S["SUEOKI"] if "基金" in str(x[0])]
_m = re.search(r"1億円の取崩しで月額約▲([\d,]+)円", str(_kk[0]) if _kk else "")
KIKIN_100M = int(_m.group(1).replace(",", "")) if _m else None


def _n(v, d=0):
    return ("{:,.%df}" % d).format(v)


# ============================================================ 図
_ZU = {z["no"]: z["name"] for z in DZ.load()}


def zu(no):
    """図の番号からPNGの所在を返す。無ければ None（点検に出す）。"""
    nm = _ZU.get(no)
    if nm is None:
        return None
    p = os.path.join(FIGSRC, nm + ".png")
    return p if os.path.exists(p) else None


# ============================================================ 新規作図2点
rcParams["font.family"] = "IPAGothic"
rcParams["axes.unicode_minus"] = False
rcParams["savefig.dpi"] = 220
rcParams["savefig.bbox"] = "tight"
rcParams["savefig.pad_inches"] = 0.04
rcParams["axes.grid"] = False
os.makedirs(FIGD, exist_ok=True)

# 圏域の表（第1章第7節）から3町の面積・人口・高齢化率を読む
_KEN = tbl("第1章第7節", "日常生活圏域名", "面積")
_KHEAD = _KEN[0]


def _col(head, *words):
    for i, h in enumerate(head):
        if all(w in h for w in words):
            return i
    raise RuntimeError("列が見つからない：%s" % (words,))


def _nendo(head):
    """表頭の「R7（基準）」から「令和7年度」を作る。年度を固定値で書かない。"""
    m = re.search(r"R(\d+)", str(head))
    if not m:
        raise RuntimeError("年度を読めない表頭：%s" % (head,))
    return "令和%s年度" % m.group(1)


_C_NM, _C_AR = _col(_KHEAD, "圏域名"), _col(_KHEAD, "面積")
_C_JN, _C_KR = _col(_KHEAD, "人口"), _col(_KHEAD, "高齢化率")
TOWNS = []
for _r in _KEN[1:]:
    _nm = _r[_C_NM].strip()
    if _nm.endswith("町"):
        TOWNS.append((_nm, _r[_C_AR], _r[_C_JN], _r[_C_KR]))
_KEI = [r for r in _KEN[1:] if r[_C_NM].strip() in ("計", "合計")]
KEI = _KEI[0] if _KEI else None


def _box(ax, x, y, w, h, text, fc, ec, fs=8.5, bold=False):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.012",
                                linewidth=0.9, facecolor=fc, edgecolor=ec))
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center",
            fontsize=fs, fontweight="bold" if bold else "normal",
            linespacing=1.45)


def fig_kenniki():
    """圏域の構成図。区域の位置関係を示す地図ではない。"""
    fig, ax = plt.subplots(figsize=(7.4, 3.0))
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    _box(ax, 0.26, 0.74, 0.48, 0.20,
         "大雪地区広域連合（保険者）\n介護保険の運営・計画の策定",
         "#DCE6F1", "#1F3864", fs=9.5, bold=True)
    w = 0.29
    for i, (nm, ar, jn, kr) in enumerate(TOWNS):
        x = 0.025 + i * (w + 0.035)
        ax.annotate("", xy=(x + w / 2, 0.56), xytext=(0.50, 0.74),
                    arrowprops=dict(arrowstyle="-", linewidth=0.8,
                                    color="#1F3864"))
        _box(ax, x, 0.10, w, 0.46,
             "%s\n\n面積　%s km²\n人口　%s\n高齢化率　%s"
             % (nm, ar, jn, kr), "#FFFFFF", "#1F3864", fs=8.8)
    ax.text(0.5, 0.025, "それぞれの町を一つの生活圏域とします"
                        "（本編の用語では日常生活圏域）",
            ha="center", va="center", fontsize=8)
    p = os.path.join(FIGD, "gaiyo_kenniki.png")
    fig.savefig(p, facecolor="white")
    plt.close(fig)
    return p


# 基本理念・基本戦略・基本目標を計画素案から読む
_RINEN_T = tbl("第4章第1節", "基本理念")
RINEN = _RINEN_T[0][1].strip().split("\n")[0].strip()
_SENRYAKU = tbl("第4章第2節", "戦略", "目指す状態")
_MOKUHYO = tbl("第4章第2節", "区分", "基本目標", "目指す成果")
MOKUHYO = [(r[0].strip(), r[1].strip(), r[2].strip()) for r in _MOKUHYO[1:]]
SENRYAKU = [(r[0].strip(), r[1].strip()) for r in _SENRYAKU[1:]]


def _wrap(s, n):
    out, line = [], ""
    for ch in s:
        line += ch
        if len(line) >= n:
            out.append(line)
            line = ""
    if line:
        out.append(line)
    return "\n".join(out)


def fig_taikei():
    """施策の体系図。基本理念・基本戦略・基本目標を計画素案から組む。"""
    fig, ax = plt.subplots(figsize=(7.4, 4.3))
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    _box(ax, 0.02, 0.855, 0.96, 0.125, _wrap(RINEN, 46),
         "#DCE6F1", "#1F3864", fs=9.2, bold=True)
    ax.text(0.5, 0.825, "基本理念", ha="center", va="center", fontsize=8,
            color="#1F3864")
    n = len(SENRYAKU)
    w = 0.96 / n - 0.012
    for i, (nm, sta) in enumerate(SENRYAKU):
        x = 0.02 + i * (w + 0.012)
        _box(ax, x, 0.60, w, 0.185,
             "%s\n\n%s" % (_wrap(nm, 11), _wrap(sta, 13)),
             "#F2F5FA", "#1F3864", fs=7.6)
    ax.text(0.5, 0.575, "4つの基本戦略", ha="center", va="center",
            fontsize=8, color="#1F3864")
    m = len(MOKUHYO)
    w2 = 0.96 / m - 0.010
    for i, (ku, nm, sei) in enumerate(MOKUHYO):
        x = 0.02 + i * (w2 + 0.010)
        _box(ax, x, 0.07, w2, 0.455,
             "%s\n\n%s\n\n〈目指す成果〉\n%s"
             % (ku, _wrap(nm, 9), _wrap(sei, 9)),
             "#FFFFFF", "#1F3864", fs=7.4)
    ax.text(0.5, 0.025, "5つの基本目標", ha="center", va="center",
            fontsize=8, color="#1F3864")
    p = os.path.join(FIGD, "gaiyo_taikei.png")
    fig.savefig(p, facecolor="white")
    plt.close(fig)
    return p


FIG_KEN = fig_kenniki()
FIG_TAI = fig_taikei()

# ============================================================ 組み立ての部品
doc = Document()
for _st in ("Normal", "Heading 1", "Heading 2", "List Bullet"):
    _f = doc.styles[_st].font
    _f.name = FONT
    doc.styles[_st].element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
doc.styles["Normal"].font.size = Pt(10.5)

sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.top_margin = sec.bottom_margin = Cm(1.8)
sec.left_margin = sec.right_margin = Cm(1.7)

HEADINGS = []
_BM = [2000]


def _fmt(run, size, bold=False, color=None):
    run.font.size = Pt(size)
    run.font.name = FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    run.bold = bold
    if color:
        run.font.color.rgb = color


def P(text="", size=10.5, bold=False, align=None, after=4, before=0,
      indent=None, ls=1.15):
    p = doc.add_paragraph()
    _fmt(p.add_run(text), size, bold)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.line_spacing = ls
    if align is not None:
        p.alignment = align
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    return p


def BUL(text, size=10, after=2):
    p = doc.add_paragraph(style="List Bullet")
    _fmt(p.add_run(text), size)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.12
    p.paragraph_format.left_indent = Cm(0.6)
    return p


def _bookmark(p, name):
    _BM[0] += 1
    bid = str(_BM[0])
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


def H(text, size=14.5):
    """頁の見出し。上下の罫線と地色の帯を付ける。"""
    p = doc.add_paragraph(style="Heading 1")
    _fmt(p.add_run(text), size, True, RGBColor(0x1F, 0x38, 0x64))
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    for side in ("top", "bottom"):
        b = OxmlElement("w:" + side)
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "18")
        b.set(qn("w:space"), "4")
        b.set(qn("w:color"), NAVY)
        bdr.append(b)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "F2F5FA")
    ref = pPr.find(qn("w:spacing"))
    if ref is None:
        pPr.append(bdr)
        pPr.append(shd)
    else:
        ref.addprevious(bdr)
        bdr.addnext(shd)
    nm = "_Gai_%03d" % (len(HEADINGS) + 1)
    HEADINGS.append((text, nm))
    _bookmark(p, nm)
    return p


def H2(text, size=11.5):
    p = doc.add_paragraph()
    _fmt(p.add_run(text), size, True, RGBColor(0x1F, 0x38, 0x64))
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
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


def _rowflag(row, header=False):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))
    if header:
        trPr.append(OxmlElement("w:tblHeader"))


_TBL_N = [0]          # 見出し行を持つ表の数（1行1列の枠と区別する）


def TBL(head, rows, widths, size=8.8, bold_col=(), center=()):
    _TBL_N[0] += 1
    t = doc.add_table(rows=1, cols=len(head))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(head):
        c = t.rows[0].cells[i]
        c.text = ""
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.paragraphs[0].paragraph_format.space_after = Pt(1)
        _fmt(c.paragraphs[0].add_run(h), size, True)
        _shade(c, HEAD_SHD)
    _rowflag(t.rows[0], header=True)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            pp = cells[i].paragraphs[0]
            pp.paragraph_format.space_after = Pt(1)
            pp.paragraph_format.line_spacing = 1.06
            if i in center:
                pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _fmt(pp.add_run("" if v is None else str(v)), size,
                 i in bold_col)
        _rowflag(t.rows[-1])
    tot = sum(widths)
    if tot > TEXTW:
        widths = [w * TEXTW / tot for w in widths]
    for i, w in enumerate(widths):
        for row in t.rows:
            row.cells[i].width = Cm(w)
    return t


def NOTE(text, size=8):
    p = doc.add_paragraph()
    _fmt(p.add_run(text), size)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.08
    return p


def FIG(path, width=16.2):
    if path is None:
        return None
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(3)
    p.add_run().add_picture(path, width=Cm(width))
    return p


_RON = []


def RONTEN(key):
    """意見を求める論点を枠で囲んで置く。"""
    r = ronten(key)
    if r is None:
        return None
    _RON.append(r[0])
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = t.rows[0].cells[0]
    c.width = Cm(TEXTW)
    c.text = ""
    p0 = c.paragraphs[0]
    p0.paragraph_format.space_after = Pt(2)
    _fmt(p0.add_run("【ご意見をお寄せください】" + r[0]), 10, True)
    p1 = c.add_paragraph()
    p1.paragraph_format.space_after = Pt(0)
    p1.paragraph_format.line_spacing = 1.1
    _fmt(p1.add_run(r[1] if r[1] else "この内容についてご意見をお聞かせください。"),
         9.5)
    _shade(c, RON_SHD)
    _rowflag(t.rows[0])
    return t


def HONPEN(sho_list):
    """本編の該当頁を頁の末尾に記す（方針④）。"""
    out = []
    for s in sho_list:
        r = sho_pages(s)
        if r:
            out.append("%s　%d頁から%d頁" % (s, r[0], r[1]))
    if not out:
        return None
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _fmt(p.add_run("詳しくは本編をご覧ください（" + "／".join(out) + "）"), 8.5)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(0)
    return p


def BREAK():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def PAGENO():
    sec.footer.is_linked_to_previous = False
    p = sec.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in p.runs:
        r.text = ""
    r = p.add_run("－ ")
    _fmt(r, 9)
    fr = p.add_run()
    _fmt(fr, 9)
    b = OxmlElement("w:fldChar")
    b.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = " PAGE "
    e = OxmlElement("w:fldChar")
    e.set(qn("w:fldCharType"), "end")
    fr._r.append(b)
    fr._r.append(it)
    fr._r.append(e)
    r2 = p.add_run(" －")
    _fmt(r2, 9)


# ============================================================ p1 表紙
# 計画として議決を経る前であり、未確定箇所が残る間は「（素案）の概要」とする。
DAI = ("第10期介護保険事業計画（素案）の概要" if PH_N
       else "第10期介護保険事業計画の概要")
HAN = "パブリックコメント用" if PH_N else "最終版"

P("", after=30)
P("大雪地区広域連合", size=12, align=WD_ALIGN_PARAGRAPH.CENTER, after=6)
P(DAI, size=20, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=6)
P("令和9年度から令和11年度まで", size=12.5,
  align=WD_ALIGN_PARAGRAPH.CENTER, after=22)
P("基本理念", size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
P(RINEN, size=13, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=20,
  ls=1.3)
FIG(FIG_KEN, width=15.4)
P("東川町・美瑛町・東神楽町", size=11, align=WD_ALIGN_PARAGRAPH.CENTER,
  after=18)
_t = doc.add_table(rows=1, cols=1)
_t.style = "Table Grid"
_t.alignment = WD_TABLE_ALIGNMENT.CENTER
_c = _t.rows[0].cells[0]
_c.width = Cm(TEXTW)
_c.text = ""
_p = _c.paragraphs[0]
_p.paragraph_format.space_after = Pt(2)
_fmt(_p.add_run("この冊子は計画の概要をまとめたものです。"), 10, True)
_p2 = _c.add_paragraph()
_p2.paragraph_format.space_after = Pt(0)
_p2.paragraph_format.line_spacing = 1.12
_fmt(_p2.add_run(
    "計画の本編は大雪地区広域連合のホームページ及び構成3町の窓口で"
    "ご覧いただけます。この冊子は本編の内容を住民の皆様に向けて"
    "書き直したものであり、数値は本編と同じものです。"
    "保険料は現時点の試算であり、今後の制度改正により変わります。"), 9.5)
_rowflag(_t.rows[0])
PAGENO()

# 制度改正は、住民の皆様に直接関わるものを選び、住民の皆様に向けて
# 書き直す（構成案 05シート 方針①　本編の文をそのまま採らない）。
# 本編の改正事項の名称には要綱の項番号が、内容には条番号が入るため、
# いずれもそのままでは概要版に用いられない。
# 鍵は素案の改正事項の名称に含まれる語であり、素案の表に1件だけ
# 当たることを自己点検で確かめる（名称が変われば不適合になる）。
KAISEI = [
    ("要介護認定等の申請に関する手続の代行",
     "要介護認定の申請を代わりに行える事業者が増えます。"
     "有料老人ホームや介護老人保健施設などをご利用の方は、"
     "その施設にお願いできるようになります。"),
    ("電子資格確認",
     "マイナンバーカードで介護保険の資格を確かめられるようになります。"),
    ("介護予防・日常生活支援総合事業",
     "介護予防と住民同士の交流をあわせて行う通いの場の拠点を"
     "運営する事業が、市町村が行う事業として位置づけられます。"),
    ("特定福祉用具販売",
     "福祉用具を購入したときに支給される額の決め方が変わり、"
     "国が定める基準による額が上限となります。"),
]

# 第10期に引き継ぐ課題も、住民の皆様に向けて書き直す（方針①）。
# 本編の表は確認方法・反映先を含む当方の整理であり、そのままでは
# 用いられない。鍵は素案の「第9期の課題」の欄に含まれる語である。
KADAI = [
    ("代表KPIが4項目",
     "達成度をはかる指標が4項目しかありませんでした。"
     "第10期は16項目に増やします。"),
    ("年次評価がない",
     "毎年度の点検の手順が定まっていませんでした。"
     "第10期は毎年度算定して公表します。"),
    ("圏域平均で地域差が不明",
     "3町の平均でみていたため町ごとの違いが分かりませんでした。"
     "第10期は町ごとにみます。"),
]

# 用語は住民の皆様に向けて言い換える（構成案 06シート）。
# 長い語を先に置く（「重度要介護認定率」が「要介護認定率」に先に当たらない）。
IIKAE = (
    ("重度要介護認定率", "重度（要介護3以上）の認定を受けた方の割合"),
    ("要介護認定率", "介護が必要と認定された方の割合"),
    ("第1号被保険者", "65歳以上の方"),
    ("有意な改善又は水準の維持", "改善又は現在の水準の維持"),
)


def iikae(s, skip=None):
    """言い換えを当てる。skip はその語を見出しとする説明文に用いる。"""
    s = str(s)
    for a, b in IIKAE:
        if a != skip:
            s = s.replace(a, b)
    return s


# 基本理念及び基本目標の名称は本編と同一の表現とする（方針①の例外）。
# 言い換えられない語は、注で意味を示す。事業名・指標名も同じ扱いとする。
GLOSS = (
    ("レジリエンス", "災害や感染症が起きても必要なサービスを続けられる力"),
    ("PDCA", "計画・実施・点検・改善を繰り返す進め方"),
    ("BCP", "災害や感染症が起きても事業を続けるための計画"),
    ("アウトリーチ事業",
     "認定を受けながらサービスを利用していない方のお宅へ伺い、"
     "状況をお聞きする取組"),
    ("医療情報との突合・縦覧点検",
     "医療と介護の請求の内容を照らし合わせて誤りを見つける取組"),
    ("給付費の計画乖離率", "計画で見込んだ費用と実際の費用との差の割合"),
)
_GL = dict(GLOSS)


def gloss_note(*keys):
    """注に掲げる意味の文を作る。説明していない語が残らないようにする。"""
    return "、".join("%sは%s" % (k, _GL[k]) for k in keys) + "をいいます。"

# ============================================================ p2 趣旨
BREAK()
H("1　計画の趣旨と位置付け")
P("介護保険法は、市町村が3年を一期とする介護保険事業計画を定めることを"
  "求めています。大雪地区広域連合は、東川町・美瑛町・東神楽町が"
  "共同で介護保険を運営するために設けられた団体であり、"
  "3町に代わって計画を定め、保険料を決め、介護サービスの給付を行います。",
  size=10)
P("この計画は、令和9年度から令和11年度までの3年間に、"
  "どれだけの介護サービスが必要になるかを見込み、"
  "その費用と保険料を定めるものです。"
  "あわせて、令和22（2040）年を見据えて、"
  "必要なときに必要なサービスを受けられる仕組みを"
  "どう保っていくかを示します。", size=10)

H2("生活圏域（本編の用語では日常生活圏域）")
P("この計画では、東川町・美瑛町・東神楽町のそれぞれを"
  "一つの生活圏域とし、3つの圏域を設けます。"
  "3町は面積・人口・高齢化率が大きく異なるため、"
  "一つにまとめて扱わず、町ごとに状況をみて整備の方針を考えます。",
  size=10)
_rows = [[nm, ar, jn, kr] for nm, ar, jn, kr in TOWNS]
if KEI:
    _rows.append(["3町計", KEI[_C_AR], KEI[_C_JN], KEI[_C_KR]])
TBL(["生活圏域", "面積（km²）", "人口", "高齢化率"], _rows,
    [4.4, 4.0, 4.4, 4.0], size=9.5, bold_col=(0,), center=(1, 2, 3))
NOTE("資料：%s。高齢化率は65歳以上の方が人口に占める割合です。"
     "美瑛町には、このほかに統計や分析のための4つの地区を設けています。"
     % jiten("第1章第7節", "住民基本台帳"))

H2("第10期の主な制度改正")
P("介護保険法などの改正により、令和9年4月から次のように変わります。"
  "計画はこれらを織り込んでいます。改正の全体は本編に掲げています。",
  size=10)
_KAI = tbl("第1章第9節", "改正事項", "内容", "施行")
_KAI_NM = [r[0].strip() for r in _KAI[1:]]
for _k, _bun in KAISEI:
    BUL(_bun, size=9.5)

H2("計画の進め方")
P("計画は、策定のための会議でのご審議、住民の皆様からのご意見の募集"
  "（パブリックコメント）、広域連合議会を経て定めます。"
  "定めた後も、毎年度その達成状況を点検し、結果を公表します。", size=10)
RONTEN("p2")
HONPEN(["第1章"])

# ============================================================ p3 現状
BREAK()
H("2　大雪圏域の現状")
_NINT = tbl("第6章第1節", "区分", "R7（基準）")
_dic = {r[0].strip(): r for r in _NINT[1:]}
_c7 = _col(_NINT[0], "R7")
_NB7 = _nendo(_NINT[0][_c7])
P("%sの3町を合わせた人口は%s、65歳以上の方は%sです。"
  "介護が必要と認定された方は%sで、65歳以上の方に占める割合は%sです。"
  % (_NB7, _dic["総人口"][_c7] + "人", _dic["65歳以上"][_c7] + "人",
     _dic["認定者数"][_c7] + "人", _dic["認定率"][_c7]), size=10)
P("人口は減っていきますが、65歳以上の方の数はほぼ横ばいです。"
  "一方で85歳以上の方は%sから令和11年度に%sへ増えます。"
  "介護を必要とする方が増えるのは、この85歳以上の方の増加によるものです。"
  % (_dic["うち85歳以上"][_c7] + "人",
     _dic["うち85歳以上"][_col(_NINT[0], "R11")] + "人"), size=10)
NOTE("本頁の数値は%sの値であり、前頁の生活圏域ごとの表は%sの値です。"
     "時点が異なります。"
     % (_NB7, jiten("第1章第7節", "住民基本台帳")))
FIG(zu("図2"), width=10.2)

H2("3町でサービスの整い方が異なります")
_ZERO = tbl("第2章第3節", "サービス", "区域内の事業所数")
P("次の表は、お住まいの町を事業の実施地域としている事業所の数です。"
  "同じサービスでも、お住まいの町により選べる数が違います。", size=10)
_zr = [[r[0].strip()] + [r[i].strip() for i in range(1, min(5, len(r)))]
       for r in _ZERO[1:5]]
TBL([c.strip() for c in _ZERO[0][:5]], _zr,
    [5.4, 3.0, 3.0, 3.0, 3.0], size=9.5, bold_col=(0,),
    center=(1, 2, 3, 4))
NOTE("町の欄は、その町を運営規程の「通常の事業実施地域」としている"
     "事業所の数です。資料：北海道「介護保険事業所一覧」（%s）。"
     % jiten("第2章第3節", "介護保険事業所一覧"))
P("圏域内に事業所がないサービスもあります。"
  "その場合も圏域外の事業所のサービスを利用できますが、"
  "交通費などの実費をご負担いただくことがあります。", size=10)
FIG(zu("図24-2"), width=9.4)
RONTEN("p3")
HONPEN(["第2章"])

# ============================================================ p4 振り返り
BREAK()
H("3　第9期計画（令和6年度〜令和8年度）の振り返り")
_K9 = tbl("第3章第1節", "指標", "第9期目標")
_c_jis = _col(_K9[0], "実績")
_c_mok = _col(_K9[0], "第9期目標")
_c_hi = _col(_K9[0], "目標比")
P("第9期計画で掲げた目標について、どこまで進んだかを点検しました。"
  "計画の達成度をはかる指標は4項目でした。", size=10)
_k9rows = []
for _r in _K9[1:]:
    _nm = iikae(re.sub(r"^[①-⑩]\s*", "", _r[0].strip()))
    _k9rows.append([_nm, noph(_r[_c_mok]).split("／")[0],
                    noph(_r[_c_jis]),
                    noph(_r[_c_hi]).split("（")[0]])
TBL(["指標", "第9期の目標", "実績", "評価"], _k9rows,
    [4.6, 4.4, 3.2, 3.0], size=9.3, bold_col=(0,), center=(3,))
NOTE("介護が必要と認定された方の割合の上昇の幅は、北海道・全国を"
     "下回っています。フレイルは加齢により心身の働きが弱まった状態、"
     "通いの場は住民の皆様が主体となって運営する集いの場です。")
FIG(zu("図19"), width=10.6)

H2("国の評価でみた強みと弱み")
P("国は、市町村の取組を評価して交付金を交付しています。"
  "3町の評価をみると、要介護度の変化や健康寿命といった"
  "「結果」は全国平均を上回る一方、"
  "取組の記録や報告が弱いという傾向があります。"
  "第10期は、取組の内容を記録し報告する仕組みを整えます。", size=10)
FIG(zu("図26-2"), width=10.2)

H2("第10期に引き継ぐ課題")
_KADAI = tbl("第3章第4節", "第9期の課題", "第10期の修正")
_KADAI_NM = [r[0].strip() for r in _KADAI[1:]]
for _k, _bun in KADAI:
    BUL(_bun, size=9.5)
RONTEN("p4")
HONPEN(["第3章"])

# ============================================================ p5 理念・目標
BREAK()
H("4　基本理念と5つの基本目標")
P("第9期の「地域包括ケアシステムの深化・推進」を引き継ぎ、"
  "第10期は『本人の生活の継続』と"
  "『サービス提供体制の持続可能性』の両立を目指します。", size=10)
P(RINEN, size=12.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=8,
  ls=1.3)
FIG(FIG_TAI, width=16.2)
TBL(["区分", "基本目標", "目指す成果"],
    [[ku, nm, sei] for ku, nm, sei in MOKUHYO],
    [2.6, 6.6, 7.6], size=9.3, bold_col=(0, 1))
NOTE("基本理念及び基本目標の名称は、本編と同じ表現としています。"
     + gloss_note("レジリエンス", "PDCA", "BCP"))

H2("達成度のはかり方")
P("計画の達成度は、16項目の指標ではかります。"
  "うち12項目を本文に掲げ、4項目は補いの指標とします。"
  "毎年度その値を算定して点検し、結果を公表します。"
  "現時点で数値を定められない指標は、令和9年度に"
  "もとになる値を把握してから目標を置きます。", size=10)
RONTEN("p5")
HONPEN(["第4章"])

# ============================================================ p6 取組
BREAK()
H("5　第10期の主な取組")
P("5つの基本目標ごとに、主な取組と達成目標をお示しします。"
  "ここに掲げたのは主なものです。それぞれの基本目標の取組の全体と、"
  "達成目標のもとになる値は本編に掲げています。", size=10)
_JIGYO_N = 0
for _ku, _nm, _sei in MOKUHYO:
    _key5 = "第5章" + _ku
    H2("%s　%s" % (_ku, _nm))
    try:
        _jt = tbl(_key5, "事業名及び内容")
    except RuntimeError:
        _jt = None
    if _jt:
        for _r in _jt[1:4]:
            # 事業名のみを掲げる。本編の達成目標の欄は年度・適合基準日等を
            # 含み、かっこ書きに入れると読めなくなる（方針①）。
            BUL(iikae(_r[0].strip().split("：")[0].split("\n")[0]),
                size=9.3, after=1)
            _JIGYO_N += 1
    try:
        _dt = tbl(_key5, "令和11年度目標")
        _row = _dt[1]
        _c_shi = _col(_dt[0], "指標")
        _c_mk = _col(_dt[0], "令和11年度目標")
        # 指標名のかっこ書きは単位・算定の方法であり概要版には用いない。
        # 基準値も、票数・算定の基礎を含むため本編に委ねる。
        _shi = iikae(re.sub(r"（.*?）", "", _row[_c_shi].strip()))
        P("達成目標　%s　→　令和11年度 %s"
          % (_shi, iikae(noph(_row[_c_mk], "検討中"))),
          size=9.3, indent=0.6, after=2)
    except RuntimeError:
        pass
NOTE("「検討中」は、目標の値をこれから定めるものです。"
     "もとになる値を令和9年度に把握してから、令和10年度・令和11年度の"
     "目標を置きます。"
     + gloss_note("アウトリーチ事業", "医療情報との突合・縦覧点検",
                  "給付費の計画乖離率"))
RONTEN("p6")
HONPEN(["第5章"])

# ============================================================ p7 見込みと保険料
BREAK()
H("6　サービスの見込みと保険料")
H2("介護が必要と認定される方の見込み")
TBL(["区分"] + Y3L + ["令和22年度"],
    [[r[0].strip(), r[_col(_NINT[0], "R9")], r[_col(_NINT[0], "R10")],
      r[_col(_NINT[0], "R11")], r[_col(_NINT[0], "R22")]]
     for r in _NINT[1:]],
    [3.4, 3.1, 3.1, 3.1, 3.4], size=9.3, bold_col=(0,),
    center=(1, 2, 3, 4))

H2("主なサービスの見込み（1か月あたりの利用者数）")
_MAIN = ["施設サービス 介護老人福祉施設", "施設サービス 介護老人保健施設",
         "居住系サービス 特定施設入居者生活介護",
         "居住系サービス 認知症対応型共同生活介護",
         "在宅サービス 訪問介護", "在宅サービス 通所介護"]
# JISSEKI・MIKOMI は要介護度別の配列であるため、区分の計を取る。
_mr = []
for _s in _MAIN:
    if _s not in JISSEKI:
        continue
    _mr.append([_s.split(" ", 1)[1], "%.1f" % sum(JISSEKI[_s])]
               + ["%.1f" % sum(MIKOMI[y][_s]) for y in Y3])
TBL(["サービス", "令和7年度（実績）"] + Y3L, _mr,
    [5.6, 3.2, 2.8, 2.8, 2.8], size=9.3, bold_col=(0,),
    center=(1, 2, 3, 4))
NOTE("資料：介護保険事業状況報告（令和7年度）及び本計画の推計。"
     "ほかのサービスの見込みは本編 第6章第2節に掲げています。")

H2("サービスにかかる費用の見込み")
# 基準年度（令和7年度）の総給付費は計画素案 第6章第3節の表から読む。
_KYU = tbl("第6章第3節", "区分", "R7実績")
_kyu_r7 = [r for r in _KYU[1:] if r[0].strip().startswith("総給付費")][0]
KYUFU_R7 = float(_kyu_r7[_col(_KYU[0], "R7実績")].replace(",", "")) * 1e6
P("3年間にかかる介護サービスの費用（総給付費）は"
  "%s億円と見込んでいます（令和9年度%s億円・令和10年度%s億円・"
  "令和11年度%s億円）。令和7年度の%s億円から%.1f％増える見込みです。"
  % (_n(sum(KYUFU3) / 1e8, 1), _n(KYUFU3[0] / 1e8, 1),
     _n(KYUFU3[1] / 1e8, 1), _n(KYUFU3[2] / 1e8, 1),
     _n(KYUFU_R7 / 1e8, 1),
     (KYUFU3[2] / KYUFU_R7 - 1) * 100), size=10)

H2("介護保険料（試算）")
P("65歳以上の方にご負担いただく保険料の基準額は、"
  "現時点の試算で月額%s円程度と見込んでいます。"
  "第9期の%s円と同じ水準です。" % (_n(KIJUN), "6,400"), size=10, bold=True)
P("この試算は、令和7年度の実績をもとに、"
  "介護報酬の改定を見込まず、積立金（介護保険の積立金）を"
  "取り崩さないものとして計算したものです。"
  "次のことが決まると金額は変わります。", size=10)
BUL("介護報酬が改定されると上がります（1％の改定で月額およそ%s円）"
    % _n(KAITEI_1), size=9.5)
BUL("お一人が1か月に使うサービスの費用がこれまでと同じ割合で伸びると、"
    "月額%s円から%s円上がります" % (_n(SUSEI[0]), _n(SUSEI[1])), size=9.5)
BUL("積立金を取り崩すと下がります（1億円の取崩しで月額およそ%s円）"
    % _n(KIKIN_100M), size=9.5)
NOTE("保険料の基準額は、国が定める介護報酬の改定率、"
     "令和8年度末の積立金の残高及び政令の改正が確定したのちに、"
     "広域連合議会の議決を経て決まります。"
     "上の金額は確定したものではありません。")
RONTEN("p7")
HONPEN(["第6章"])

# ============================================================ p8 裏表紙
BREAK()
H("ご意見をお寄せください")
P("この計画（素案）について、住民の皆様のご意見を募集します。"
  "いただいたご意見は、計画をまとめる際の参考とし、"
  "ご意見の概要と広域連合の考え方を後日公表します。", size=10)
TBL(["項目", "内容"],
    [["募集の期間", "令和9年1月〔　〕日から〔　〕日まで（2週間）"],
     ["資料の閲覧", "大雪地区広域連合のホームページ、"
                    "広域連合事務局及び構成3町の担当窓口"],
     ["提出の方法", "ホームページの入力欄、電子メール、郵送又は持参"],
     ["提出先", "〔　　　　　　　　　　　　　　　　　　　　　〕"],
     ["お問合せ", "大雪地区広域連合〔　　　　　　　　　　　　　〕"]],
    [3.4, 13.0], size=9.5, bold_col=(0,))
NOTE("〔　〕は決定後に記入します。"
     "期間・提出先・問合せ先は、パブリックコメントの実施要領により定まります。")

H2("計画を定めた後の進め方")
P("計画は、定めて終わるものではありません。"
  "毎年度、指標の値を算定して達成状況を点検し、"
  "結果を会議にお示しして公表します。"
  "必要があれば、計画期間の途中でも見直します。", size=10)

H2("用語の説明")
_YOGO_PICK = ("要介護認定", "日常生活圏域", "地域包括ケアシステム",
              "介護予防・日常生活支援総合事業", "地域包括支援センター",
              "保険料基準額", "介護給付費準備基金",
              "保険者機能強化推進交付金・介護保険保険者努力支援交付金")
_yd = dict(SD.YOGO)
_ymiss = [k for k in _YOGO_PICK if k not in _yd]
if _ymiss:
    raise RuntimeError("用語の解説に見出し語がない：%s" % "、".join(_ymiss))
_yr = [[k, iikae(_yd[k], skip=k)] for k in _YOGO_PICK]
TBL(["用語", "意味"], _yr, [4.2, 12.2], size=9, bold_col=(0,))
NOTE("このほかの用語は本編の巻末に掲げています。")

# ============================================================ 保存
_st = doc.settings.element
_zoom = _st.find(qn("w:zoom"))
if _zoom is not None and _zoom.get(qn("w:percent")) is None:
    _zoom.set(qn("w:percent"), "100")
docx_fix.fix(doc)
doc.save(OUT)

# ============================================================ 自己点検
_out = Document(OUT)
_txt = "\n".join(p.text for p in _out.paragraphs)
for _t in _out.tables:
    for _r in _t.rows:
        for _c in _r.cells:
            _txt += "\n" + _c.text
_txt_n = _txt.replace("%", "％")

chk(1, "構成案の頁割を読めていること",
    "%d頁（p1からp%d）" % (len(PAGES), len(PAGES)), len(PAGES) == 8)
chk(2, "計画素案を読めていること",
    "%d段落%d表／章節%d件" % (len(src.paragraphs), len(src.tables), len(SEC_T)),
    len(src.paragraphs) > 500 and len(SEC_T) >= 30)
chk(3, "基本理念が計画素案と同一の表現であること",
    RINEN[:24] + "…", RINEN and RINEN in _ALL and RINEN in _txt)
chk(4, "5つの基本目標が計画素案と一致すること",
    "%d件：%s" % (len(MOKUHYO), "／".join(m[0] for m in MOKUHYO)),
    len(MOKUHYO) == 5 and all(m[1] in _txt for m in MOKUHYO))
chk(5, "3町の面積・人口・高齢化率を圏域の表から読めていること",
    "／".join("%s %s" % (t[0], t[3]) for t in TOWNS),
    len(TOWNS) == 3 and all(t[3].endswith("％") for t in TOWNS))
chk(6, "意見を求める論点が構成案の6件そろっていること",
    "%d件：%s" % (len(_RON), "／".join(r.split("　")[0] for r in _RON)),
    len(_RON) == 6)
chk(7, "本編の頁の案内を各頁に置いたこと（表紙・裏表紙を除く）",
    "%d件" % _txt.count("詳しくは本編をご覧ください"),
    _txt.count("詳しくは本編をご覧ください") == 6)
chk(8, "未確定箇所の記号を概要版に持ち込んでいないこと（方針③）",
    "計画素案%d箇所／概要版%d箇所" % (PH_N, len(PH.findall(_txt))),
    not PH.search(_txt))
chk(9, "未確定箇所が残る間は表題を「（素案）の概要」とすること",
    "表題 %s／版 %s" % (DAI, HAN),
    (PH_N > 0) == ("（素案）の概要" in _txt))
chk(10, "保険料を試算値として示し、確定でないことを書いていること",
    "基準額の試算 %s円" % _n(KIJUN),
    ("%s円" % _n(KIJUN)) in _txt
    and "確定したものではありません" in _txt and "試算" in _txt)
chk(11, "保険料の試算が算定の値と一致すること",
    "算定上の月額 %.2f円 → 百円未満四捨五入 %s円" % (GETSU, _n(KIJUN)),
    abs(KIJUN - round(GETSU / 100.0) * 100) < 1 and 6000 <= KIJUN <= 7000)
chk(12, "総給付費が算定の値と一致すること",
    "3か年計 %s円（%s億円）" % (_n(sum(KYUFU3)), _n(sum(KYUFU3) / 1e8, 1)),
    _n(sum(KYUFU3) / 1e8, 1) in _txt)
chk(13, "認定者数の見込みが計画素案の表と一致すること",
    "令和11年度 %s人" % _dic["認定者数"][_col(_NINT[0], "R11")],
    _dic["認定者数"][_col(_NINT[0], "R11")] in _txt)
chk(14, "主なサービスの見込みを掲げたこと",
    "%d区分" % len(_mr), len(_mr) >= 5)
chk(15, "主な取組を基本目標ごとに掲げたこと",
    "%d件（5目標）" % _JIGYO_N, _JIGYO_N >= 12)
chk(16, "新規作図2点を作成したこと",
    "圏域の構成図・施策の体系図",
    os.path.exists(FIG_KEN) and os.path.exists(FIG_TAI))
_figs = len(_out.inline_shapes)
chk(17, "本編の図を用いたこと（新規作図2点を含む）",
    "%d点" % _figs, _figs >= 6)
_miss = [n for n in ("図2", "図24-2", "図19", "図26-2") if zu(n) is None]
chk(18, "構成案が挙げた本編の図が所在すること",
    "所在しない図 %s" % ("、".join(_miss) or "なし"), not _miss)
chk(19, "連絡先を記入枠としたこと（個人情報及び連絡先を収録しない）",
    "記入枠 %d件" % _txt.count("〔"), _txt.count("〔") >= 4)
_NAIBU = ("固定値", "runpy", ".py", "再実行", "スクリプト", "章節ごとに",
          "書き写して", "判定しています", "受託者", "本素案", "協議用",
          "修正指示書", "確認事項No", "リポジトリ")
chk(20, "受託者の内部の仕組みの語が現れないこと",
    "、".join(w for w in _NAIBU if w in _txt) or "現れない",
    not any(w in _txt for w in _NAIBU))
_KIN = ("に由来する", "と整合する", "1件も", "全国トップ級",
        "有意差がないため")
chk(21, "禁止表現が現れないこと",
    "、".join(w for w in _KIN if w in _txt) or "現れない",
    not any(w in _txt for w in _KIN))
_TADAN = ("北塩原村", "浜田地区", "川崎町", "金ヶ崎町", "空知中部",
          "日高中部", "後志広域", "京極町", "小野町")
chk(22, "他の団体の固有名称が現れないこと",
    "、".join(w for w in _TADAN if w in _txt) or "現れない",
    not any(w in _txt for w in _TADAN))
chk(23, "個人情報の形（電話番号・メールアドレス）が現れないこと",
    "%d件" % (len(re.findall(r"0\d{1,4}-\d{1,4}-\d{3,4}", _txt))
              + len(re.findall(r"[\w.+-]+@[\w.-]+\.\w+", _txt))),
    not re.search(r"0\d{1,4}-\d{1,4}-\d{3,4}", _txt)
    and not re.search(r"[\w.+-]+@[\w.-]+\.\w+", _txt))
_OK_CH = re.compile(r"[ぁ-んァ-ヴ一-龥々ー０-９0-9A-Za-zＡ-Ｚａ-ｚ"
                    r"、。，．・（）「」『』〔〕【】〜～％!？:：/／"
                    r"＋－±×÷=＝<>＜＞\-–—―№㎞²°\s　→←↑↓"
                    r"①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
                    r"ⅠⅡⅢⅣⅤ\"'\(\)\[\]\.,;%&#*＊〈〉▲△○◇◎■□]")
_zatsu = sorted({ch for ch in _txt if not _OK_CH.match(ch)})
chk(24, "許容する文字の外の文字が混じっていないこと",
    "".join(_zatsu) or "混じっていない", not _zatsu)
# 見出し行を持つ表（TBL で作ったもの）はすべて繰返しの設定を持つ。
# 1行1列の枠（論点・表紙の囲み）は見出し行を持たないため対象外とする。
_hdr_n = sum(1 for t in _out.tables
             if t.rows[0]._tr.find(qn("w:trPr")) is not None
             and t.rows[0]._tr.find(qn("w:trPr")).find(
                 qn("w:tblHeader")) is not None)
chk(25, "見出し行を持つ表は各ページで見出しを繰り返す設定であること",
    "設定あり%d表／見出し行を持つ表%d表／枠%d件"
    % (_hdr_n, _TBL_N[0], len(_out.tables) - _TBL_N[0]),
    _hdr_n == _TBL_N[0] and _TBL_N[0] + len(_RON) + 1 == len(_out.tables))
chk(26, "用語の説明を置いたこと", "%d語" % len(_yr), len(_yr) >= 5)
# 制度改正の言い換えは、素案の改正事項の名称に1件だけ当たる鍵で選ぶ。
# 名称が変われば当たらなくなり不適合になる（当方の書き写しが残らない）。
_kmiss = [k for k, _ in KAISEI
          if sum(1 for nm in _KAI_NM if k in nm) != 1]
chk(27, "制度改正の言い換えが計画素案の改正事項に1件ずつ当たること",
    "%d件／当たらない又は複数に当たる鍵 %s"
    % (len(KAISEI), "、".join(_kmiss) or "なし"), not _kmiss)
_dmiss = [k for k, _ in KADAI
          if sum(1 for nm in _KADAI_NM if k in nm) != 1]
chk(27.5, "引き継ぐ課題の言い換えが計画素案の課題に1件ずつ当たること",
    "%d件／当たらない又は複数に当たる鍵 %s"
    % (len(KADAI), "、".join(_dmiss) or "なし"), not _dmiss)
# 本編の文をそのまま採ると条番号・要綱の項番号が住民向けの資料に出る
# （構成案 05シート 方針①）。値の形で探す。
_JOBUN = (re.findall(r"第\d+条", _txt) + re.findall(r"要綱", _txt)
          + re.findall(r"第\d+項", _txt) + re.findall(r"第\d+号", _txt))
chk(28, "条番号・要綱の項番号を概要版に持ち込んでいないこと（方針①）",
    "、".join(sorted(set(_JOBUN))) or "現れない", not _JOBUN)
# 用語は住民の皆様に向けて言い換える（構成案 06シート）。
# 用語の説明の表に掲げる語は、意味を示すものであるため対象から外す。
_IIKAE = ("第1号被保険者", "要介護認定率", "KPI", "地域差指数")
_ii = [w for w in _IIKAE
       if w in _txt and not any(w == k for k, _ in _yr)]
chk(28.5, "言い換えることとした用語が本文に現れないこと",
    "、".join(_ii) or "現れない", not _ii)
# 時点の違う人口を2か所に掲げるときは、その違いを注に書く
# （同じ項目に2つの数値を断りなく並べない）。
_J17 = jiten("第1章第7節", "住民基本台帳")
# 本編と同じ表現とした名称に含まれる語は、注で意味を示す。
# 語が現れているのに意味を示していなければ不適合になる。
_gmiss = [k for k, v in GLOSS if k in _txt and v not in _txt_n]
chk(28.7, "言い換えられない語の意味を注で示していること",
    "、".join(_gmiss) or "示していない語はない", not _gmiss)
chk(29, "時点の異なる人口を掲げるときはその時点を注に書いていること",
    "生活圏域の表 %s／現状 %s" % (_J17, _NB7),
    _J17 in _txt and _NB7 in _txt and "時点が異なります" in _txt)

NG = [c for c in CHECKS if not c[3]]
print("保存 %s" % OUT)
print("図 %s" % FIGD)
print("表題 %s（%s）" % (DAI, HAN))
print("段落%d／表%d／図%d"
      % (len(_out.paragraphs), len(_out.tables), _figs))
print("意見を求める論点 %d件" % len(_RON))
print("保険料 算定上の月額%.2f円 → 基準額の試算 %s円" % (GETSU, _n(KIJUN)))
print("総給付費（3か年計） %s円" % _n(sum(KYUFU3)))
print("自己点検 %d件　不適合 %d件" % (len(CHECKS), len(NG)))
for c in NG:
    print("  不適合 No.%s %s → %s" % (c[0], c[1], c[2]))
if NG:
    sys.exit(1)
