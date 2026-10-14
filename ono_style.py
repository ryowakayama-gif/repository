"""小野町 計画策定業務　成果品の体裁。

他町村の案件で用いている「策定委員会資料のデザインルール」に合わせたもの。
**規則のみを取り入れたものであり、他団体の数値は一切用いていない。**

合わせた項目
  用紙・余白　　A4縦　上下2.0cm・左右1.9cm　ヘッダ／フッタ1.25cm
  本文　　　　　游ゴシック 10.5pt　段落前後3pt
  見出し　　　　H1 15pt 太字 #2E74B5／H2 12.5pt 太字 #2E74B5／H3 11pt 太字 #1F4D78
  　　　　　　　**令和8年10月14日に、Wordの見出しスタイル（Heading 1〜3）を
  　　　　　　　当てるように改めた。** 見た目は変えていない。従前はすべて
  　　　　　　　「標準」スタイルで、大きさと色を直に指定していたため、
  　　　　　　　ナビゲーションウィンドウ・PDFのしおり・スクリーンリーダーの
  　　　　　　　いずれにも見出しが出なかった（外部レビューの指摘）。
  　　　　　　　run への直の指定は残してある。既存の点検が run の大きさで
  　　　　　　　見出しを判定しているため。
  表　　　　　　見出し行 塗り#1F3864・白文字8.5pt太字・中央／本文8.5pt

  **令和8年9月29日に、書体と文字の大きさを他団体の協議用素案に合わせた。**
  従前は BIZ UDPゴシック 12pt（表11pt）だったが、表の多い計画書では
  1行に入る文字数が少なく折り返しが増えるため、游ゴシック10.5pt（表8.5pt）
  とし、左右の余白を1.9cmに詰めて本文幅を17.2cmに広げた。
  表の塗り・罫線・強調の扱いは従前のまま。
  　　　　　　　罫線 上下#2E75B6(sz6)、左右・内側#BFBFBF(sz4)、セル余白左右10dxa
  　　　　　　　列幅はセル幅とグリッド（w:gridCol）の両方に書く
  　　　　　　　（一方だけではLibreOfficeで反映されない）
  出典　　　　　8pt　段落前2pt・後8pt
  図　　　　　　キャプションは図の上に【図N】、出典は図の下に8pt
  ヘッダ　　　　文書名を中央に
  フッタ　　　　ページ番号を「- n -」の形で中央に
  強調　　　　　本文・表とも `**…**` で囲んだ部分を太字にする

使い方
    import ono_style as S
    r = S.Report("小野町高齢者保健福祉計画・第10期介護保険事業計画")
    r.cover(["小野町高齢者保健福祉計画", "第10期介護保険事業計画"], ["素案", "令和8年9月"])
    r.toc_here(); r.body_here()
    r.h1("1　評価の枠組み")
    r.p("**強調したい部分**はこのように書く。")
    r.tbl([["区分", "値"], ["総給付費", "1,005,144千円"]])
    r.src("資料：介護保険事業状況報告（年報）様式3。")
    r.build_toc(); r.save("out.docx")
"""

import json
import os
import pathlib
import re

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import (WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT,
                            WD_TAB_LEADER)
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

FONT = "游ゴシック"
BODY_PT = 10.5
TBL_PT = 8.5
SRC_PT = 8.5
C_H1 = RGBColor(0x2E, 0x74, 0xB5)
C_H3 = RGBColor(0x1F, 0x4D, 0x78)
TEXTW = 17.2                      # 本文の幅（cm）＝21.0−1.9−1.9
SHORT_TBL = 12                    # 寡婦行の処理を掛ける表の行数（見出し行を含む）
#   当初7行までとしていたが、10行の表（成果指標の仕様の1件ずつの表）で
#   孤立行が出たため12行までに広げた。これより長い表は、見出し行が次頁の
#   頭で繰り返されるため、続きが1行でも読める。


def _set_font(run, f=FONT, sz=BODY_PT, b=False, color=None, italic=False):
    run.font.name = f
    run.font.size = Pt(sz)
    run.bold = b
    run.italic = italic
    rf = run._element.get_or_add_rPr().get_or_add_rFonts()
    for a in ("w:eastAsia", "w:ascii", "w:hAnsi"):
        rf.set(qn(a), f)
    if color is not None:
        run.font.color.rgb = color


def _emph(p, text, sz=BODY_PT, b=False, color=None, italic=False):
    """`**…**` で囲まれた部分を太字にして段落に流し込む。"""
    for seg in re.split(r"(\*\*.*?\*\*)", str(text)):
        if not seg:
            continue
        bold = seg.startswith("**") and seg.endswith("**")
        _set_font(p.add_run(seg[2:-2] if bold else seg),
                  sz=sz, b=bold or b, color=color, italic=italic)
    return p


def _fldchar(kind):
    e = OxmlElement("w:fldChar")
    e.set(qn("w:fldCharType"), kind)
    return e


# 見出しの体裁。(スタイル名, 大きさpt, 色, 段落前pt, 段落後pt)
# 見た目は従前のまま。変えたのは「Wordの見出しスタイルを当てる」ことだけである。
HEAD_SPEC = {
    1: ("Heading 1", 15,   C_H1, 12, 5),
    2: ("Heading 2", 12.5, C_H1,  9, 5),
    3: ("Heading 3", 11,   C_H3,  9, 4),
}


def _set_outline(style, level):
    """スタイルに見出しの階層（w:outlineLvl）を書く。

    これが無いと、Wordのナビゲーションウィンドウにも、
    PDFのしおりにも、スクリーンリーダーの見出し一覧にも出ない。
    """
    pPr = style.element.get_or_add_pPr()
    for old in pPr.findall(qn("w:outlineLvl")):
        pPr.remove(old)
    e = OxmlElement("w:outlineLvl")
    e.set(qn("w:val"), str(level - 1))          # 0起点
    pPr.append(e)


def _auto_widths(rows, total=TEXTW, minw=1.8):
    """列幅を中身の長さから決める。合計は本文幅にそろえる。"""
    ncol = max(len(r) for r in rows)
    need = []
    for c in range(ncol):
        longest = max((len(re.sub(r"\*\*", "", str(r[c])))
                       for r in rows if c < len(r)), default=1)
        need.append(max(longest, 2) ** 0.72)    # 長い列を広げすぎない
    scale = (total - minw * ncol) / sum(need)
    w = [minw + n * scale for n in need]
    return [x * total / sum(w) for x in w]


class Report:
    """1つの成果品（docx）。"""

    def __init__(self, title, body_pt=BODY_PT, landscape=False):
        self.title = title
        self.body_pt = body_pt
        self.doc = Document()
        self.headings = []            # (level, text, bookmark名)
        self.figno = 0
        self.tblno = 0          # 置いた表の総数（番号を振らないものも含む）
        self.capno = 0          # 【表N】の番号。題を付けた表だけ数える
        # 図表番号一覧（資料編）を組むための控え。(番号, 見出し) を順に積む。
        self.figlist = []
        self.tbllist = []
        self._toc_anchor = None
        self._landscape = landscape
        self._setup_section(self.doc.sections[0])
        st = self.doc.styles["Normal"]
        st.font.name = FONT
        st.font.size = Pt(body_pt)
        for a in ("w:eastAsia", "w:ascii", "w:hAnsi"):
            st.element.rPr.rFonts.set(qn(a), FONT)
        st.paragraph_format.space_before = Pt(3)
        st.paragraph_format.space_after = Pt(3)
        self._setup_heading_styles()

    # ------------------------------------------------------------ 見出しスタイル
    def _setup_heading_styles(self):
        """Wordの見出しスタイル（Heading 1〜3）を本書の体裁に合わせる。

        見た目は従前と同じにしたうえで、スタイルと w:outlineLvl を与える。
        これにより次が働くようになる。

          ・Wordのナビゲーションウィンドウでの章立ての移動
          ・PDFのしおり
          ・スクリーンリーダー・PDFのタグ付け（見出し構造）
          ・町が節を足したときに目次のフィールドが拾えること

        段落側でも従前どおり run に大きさ・色・太字を直に書いている。
        既存の点検（`build_ono_kosshi.soan_kosei`・`check_ono_taihi.midashi_check`）が
        run のフォントの大きさで見出しを判定しているため、そちらを壊さない。
        """
        for lv, (name, sz, color, before, after) in HEAD_SPEC.items():
            st = self.doc.styles[name]
            st.base_style = self.doc.styles["Normal"]
            st.quick_style = True
            st.hidden = False
            st.font.name = FONT
            st.font.size = Pt(sz)
            st.font.bold = True
            st.font.italic = False
            st.font.color.rgb = color
            rf = st.element.get_or_add_rPr().get_or_add_rFonts()
            for a in ("w:eastAsia", "w:ascii", "w:hAnsi"):
                rf.set(qn(a), FONT)
            pf = st.paragraph_format
            pf.space_before = Pt(before)
            pf.space_after = Pt(after)
            pf.keep_with_next = True
            pf.left_indent = Cm(0)
            pf.first_line_indent = Cm(0)
            _set_outline(st, lv)

    # ------------------------------------------------------------ 体裁
    def _setup_section(self, s, pageno=True):
        if self._landscape:
            s.page_width, s.page_height = Cm(29.7), Cm(21.0)
        else:
            s.page_width, s.page_height = Cm(21.0), Cm(29.7)
        s.left_margin = s.right_margin = Cm(1.9)
        s.top_margin = s.bottom_margin = Cm(2.0)
        s.header_distance = s.footer_distance = Cm(1.25)
        if s is not self.doc.sections[0]:
            s.header.is_linked_to_previous = False
            s.footer.is_linked_to_previous = False
        hp = s.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in list(hp.runs):
            r.text = ""
        _set_font(hp.add_run(self.title), sz=9)
        fp = s.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in list(fp.runs):
            r.text = ""
        if pageno:
            _set_font(fp.add_run("- "), sz=10)
            self._field(fp, " PAGE ")
            _set_font(fp.add_run(" -"), sz=10)
            for r in fp.runs:
                _set_font(r, sz=10)
        return s

    @staticmethod
    def _field(p, instr):
        for tag, extra in (("begin", None), (None, instr), ("end", None)):
            r = OxmlElement("w:r")
            if tag:
                e = OxmlElement("w:fldChar")
                e.set(qn("w:fldCharType"), tag)
            else:
                e = OxmlElement("w:instrText")
                e.set(qn("xml:space"), "preserve")
                e.text = extra
            r.append(e)
            p._p.append(r)

    # ------------------------------------------------------------ 本文
    def p(self, text="", sz=None, b=False, ind=0, align=None, color=None,
          before=3, after=3, italic=False):
        q = self.doc.add_paragraph()
        pf = q.paragraph_format
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        if ind:
            pf.left_indent = Cm(ind)
        if align is not None:
            q.alignment = align
        _emph(q, text, sz=sz or self.body_pt, b=b, color=color, italic=italic)
        return q

    def bul(self, text, ind=0.4):
        return self.p("・" + str(text), ind=ind, before=2, after=2)

    def src(self, text):
        return self.p(text, sz=SRC_PT, ind=0.2, before=2, after=8)

    def spacer(self, pt=6):
        q = self.doc.add_paragraph()
        q.paragraph_format.space_after = Pt(pt)
        return q

    def pagebreak(self):
        self.doc.add_page_break()

    # ------------------------------------------------------------ 見出し
    def _head(self, text, level, toc=True):
        """見出しを1つ置く。

        level … 見出しの階層（1〜3）。Heading 1〜3 のスタイルを当てる。
        toc   … 目次に載せるか。H3は従前どおり載せない（スタイルは当てる）。
        """
        name, sz, color, before, after = HEAD_SPEC[level]
        q = self.doc.add_paragraph(style=self.doc.styles[name])
        q.paragraph_format.space_before = Pt(before)
        q.paragraph_format.space_after = Pt(after)
        q.paragraph_format.keep_with_next = True
        _set_font(q.add_run(str(text).replace("**", "")), sz=sz, b=True,
                  color=color)
        if toc:
            self._register(q, str(text).replace("**", ""), level)
        return q

    def h1(self, text):
        return self._head(text, 1)

    def h2(self, text):
        return self._head(text, 2)

    def h3(self, text):
        return self._head(text, 3, toc=False)

    def part(self, text):
        """第1部・第2部のような大区切り。改ページして中央に置く。"""
        self.pagebreak()
        q = self.doc.add_paragraph(style=self.doc.styles["Heading 1"])
        q.alignment = WD_ALIGN_PARAGRAPH.CENTER
        q.paragraph_format.space_before = Pt(24)
        q.paragraph_format.space_after = Pt(18)
        _set_font(q.add_run(str(text).replace("**", "")), sz=18, b=True,
                  color=C_H1)
        self._register(q, str(text).replace("**", ""), 1)
        return q

    def cover(self, mains, subs):
        for t in mains:
            q = self.doc.add_paragraph()
            q.alignment = WD_ALIGN_PARAGRAPH.CENTER
            q.paragraph_format.space_before = Pt(18)
            q.paragraph_format.space_after = Pt(6)
            _set_font(q.add_run(t), sz=22, b=True, color=C_H1)
        for t in subs:
            q = self.doc.add_paragraph()
            q.alignment = WD_ALIGN_PARAGRAPH.CENTER
            q.paragraph_format.space_after = Pt(6)
            _set_font(q.add_run(t), sz=13)
        return self

    # ------------------------------------------------------------ 表
    @staticmethod
    def _border(tbl):
        pr = tbl.tblPr
        for old in pr.findall(qn("w:tblBorders")):
            pr.remove(old)
        b = OxmlElement("w:tblBorders")
        for tag, sz, col in (("top", "6", "2E75B6"), ("left", "4", "BFBFBF"),
                             ("bottom", "6", "2E75B6"), ("right", "4", "BFBFBF"),
                             ("insideH", "4", "BFBFBF"),
                             ("insideV", "4", "BFBFBF")):
            e = OxmlElement("w:" + tag)
            e.set(qn("w:val"), "single")
            e.set(qn("w:sz"), sz)
            e.set(qn("w:space"), "0")
            e.set(qn("w:color"), col)
            b.append(e)
        pr.append(b)
        m = OxmlElement("w:tblCellMar")
        for tag in ("left", "right"):
            e = OxmlElement("w:" + tag)
            e.set(qn("w:w"), "10")
            e.set(qn("w:type"), "dxa")
            m.append(e)
        pr.append(m)

    def tbl(self, rows, widths=None, fs=TBL_PT, header=True, right=(),
            caption=None):
        """rows[0] を見出し行として表を置く。

        widths を省くと中身の長さから決め、合計を本文幅にそろえる。
        right に列番号（0起点）を渡すとその列を右寄せにする。
        caption を渡すと表の上に【表N】を置き、図表番号一覧に控える。
        渡さなければ番号を振らない（従前どおり）。
        """
        if caption:
            self.capno += 1
            self.tbllist.append((self.capno, caption))
            q = self.doc.add_paragraph()
            q.paragraph_format.space_before = Pt(10)
            q.paragraph_format.space_after = Pt(2)
            q.paragraph_format.keep_with_next = True
            _set_font(q.add_run("【表%d】%s" % (self.capno, caption)),
                      sz=11, b=True, color=C_H3)
        rows = [list(r) for r in rows]
        ncol = max(len(r) for r in rows)
        for r in rows:
            r.extend([""] * (ncol - len(r)))
        t = self.doc.add_table(rows=len(rows), cols=ncol)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        self._border(t._tbl)
        for ri, row in enumerate(rows):
            for ci, v in enumerate(row):
                c = t.cell(ri, ci)
                para = c.paragraphs[0]
                for r0 in list(para.runs):
                    r0._element.getparent().remove(r0._element)
                para.paragraph_format.space_before = Pt(1)
                para.paragraph_format.space_after = Pt(1)
                head = header and ri == 0
                if head:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif ci in right:
                    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                _emph(para, str(v), sz=fs, b=head,
                      color=RGBColor(0xFF, 0xFF, 0xFF) if head else None)
                tcPr = c._tc.get_or_add_tcPr()
                if head:
                    sh = OxmlElement("w:shd")
                    sh.set(qn("w:val"), "clear")
                    sh.set(qn("w:fill"), "1F3864")
                    tcPr.append(sh)
                va = OxmlElement("w:vAlign")
                va.set(qn("w:val"), "center")
                tcPr.append(va)
        # 短い表の寡婦行の処理。
        #   表が頁をまたぐと、続きの頁に見出し行＋最後の1行だけが残ることがある。
        #   最後から2番目の行に keepNext を掛け、最後の行と一緒に次の頁へ送る。
        #   表そのものを分割させない形にすると頁が増えるため、こちらを採る。
        #   長い表は見出し行が次頁の頭で繰り返されるため、続きが1行でも読める。
        if 3 <= len(rows) <= SHORT_TBL:
            for c in t.rows[-2].cells:
                for para in c.paragraphs:
                    para.paragraph_format.keep_with_next = True
        w = widths or _auto_widths(rows)
        for ci, cw in enumerate(w):
            for r in t.rows:
                r.cells[ci].width = Cm(cw)
        grid = t._tbl.find(qn("w:tblGrid"))
        if grid is not None:
            for ci, gc in enumerate(grid.findall(qn("w:gridCol"))):
                if ci < len(w):
                    gc.set(qn("w:w"), str(int(Cm(w[ci]).twips)))
        self.doc.add_paragraph().paragraph_format.space_after = Pt(2)
        self.tblno += 1
        return t

    # ------------------------------------------------------------ 図
    def fig(self, path, caption, width=14.0, src=None):
        """図を1つ置く。キャプションは図の上、出典は図の下に置く。"""
        self.figno += 1
        self.figlist.append((self.figno, caption))
        q = self.doc.add_paragraph()
        q.paragraph_format.space_before = Pt(10)
        q.paragraph_format.space_after = Pt(2)
        q.paragraph_format.keep_with_next = True
        _set_font(q.add_run("【図%d】%s" % (self.figno, caption)), sz=11,
                  b=True, color=C_H3)
        r = self.doc.add_paragraph()
        r.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r.paragraph_format.space_before = Pt(0)
        r.paragraph_format.space_after = Pt(2 if src else 8)
        r.paragraph_format.keep_with_next = bool(src)
        r.add_run().add_picture(str(path), width=Cm(width))
        if src:
            self.src(src)
        return self.figno

    # ------------------------------------------------------------ 目次
    def _register(self, p, text, level):
        name = "_Toc_%03d" % (len(self.headings) + 1)
        self.headings.append((level, text, name))
        bid = str(len(self.headings) + 100)
        st = OxmlElement("w:bookmarkStart")
        st.set(qn("w:id"), bid)
        st.set(qn("w:name"), name)
        en = OxmlElement("w:bookmarkEnd")
        en.set(qn("w:id"), bid)
        pPr = p._p.find(qn("w:pPr"))
        if pPr is None:
            p._p.insert(0, st)
        else:
            pPr.addnext(st)
        p._p.append(en)
        return p

    def toc_here(self, title="目　次", field=False):
        """目次を置く位置を決める。

        field=False（既定）… 見出しの一覧を当方で組み、ページ番号だけを
            PAGEREF フィールドにする。**開いた瞬間から中身が見える**ため、
            Wordでフィールドを更新しなくても、PDFに変換しても目次が空にならない。
        field=True … Word の目次フィールド（TOC）を置く。町が節を足しても
            更新すれば拾えるが、**更新するまで中身は空である**。
            見出しスタイルを当てたことで、どちらも使える状態になっている。
        """
        self.pagebreak()
        q = self.doc.add_paragraph()
        q.alignment = WD_ALIGN_PARAGRAPH.CENTER
        q.paragraph_format.space_after = Pt(14)
        _set_font(q.add_run(title), sz=15, b=True, color=C_H1)
        r = self.doc.add_paragraph()
        r.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r.paragraph_format.space_after = Pt(10)
        _set_font(r.add_run(
            "各行をクリックすると本文へ移動します。"
            + ("ページ番号はWordで開いた時点で更新されます" if not field else
               "Wordで「フィールドの更新」を行うと作り直されます")), sz=9)
        if field:
            p = self.doc.add_paragraph()
            self._field(p, ' TOC \\o "1-2" \\h \\z \\u ')
            self._toc_anchor = None
            self._toc_field = True
            self._update_fields()
            return p
        self._toc_anchor = self.doc.add_paragraph()
        return self._toc_anchor

    def body_here(self):
        """本文の開始位置でセクションを切り、ページ番号を1から数え直す。"""
        s = self.doc.add_section(WD_SECTION.NEW_PAGE)
        self._setup_section(s)
        self._pgnum_start(s, 1)
        self._setup_section(self.doc.sections[0], pageno=False)
        return s

    @staticmethod
    def _pgnum_start(section, start):
        sp = section._sectPr
        e = sp.find(qn("w:pgNumType"))
        if e is None:
            e = OxmlElement("w:pgNumType")
            after = None
            for tag in ("w:cols", "w:formProt", "w:vAlign", "w:noEndnote",
                        "w:titlePg", "w:textDirection", "w:docGrid"):
                after = sp.find(qn(tag))
                if after is not None:
                    break
            if after is None:
                sp.append(e)
            else:
                after.addprevious(e)
        e.set(qn("w:start"), str(start))

    def _toc_run(self, parent, txt=None, child=None, bold=False, size=10.5):
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        f = OxmlElement("w:rFonts")
        for a in ("w:ascii", "w:hAnsi", "w:eastAsia"):
            f.set(qn(a), FONT)
        rPr.append(f)
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
        parent.append(r)
        return r

    def _toc_line(self, level, text, name, pages):
        q = self.doc.add_paragraph()
        pf = q.paragraph_format
        pf.space_before = Pt(1 if level == 1 else 0)
        pf.space_after = Pt(1 if level == 1 else 0)
        pf.line_spacing = 1.05
        pf.left_indent = Cm(0 if level == 1 else 0.9)
        pf.tab_stops.add_tab_stop(Cm(TEXTW), WD_TAB_ALIGNMENT.RIGHT,
                                  WD_TAB_LEADER.DOTS)
        hl = OxmlElement("w:hyperlink")
        hl.set(qn("w:anchor"), name)
        q._p.append(hl)
        self._toc_run(hl, txt=text, bold=(level == 1),
                      size=11 if level == 1 else 10)
        self._toc_run(hl, child=OxmlElement("w:tab"), size=10)
        self._toc_run(hl, child=_fldchar("begin"), size=10)
        it = OxmlElement("w:instrText")
        it.set(qn("xml:space"), "preserve")
        it.text = " PAGEREF %s \\h " % name
        self._toc_run(hl, child=it, size=10)
        self._toc_run(hl, child=_fldchar("separate"), size=10)
        self._toc_run(hl, txt=str(pages.get(name, "－")), bold=(level == 1),
                      size=10)
        self._toc_run(hl, child=_fldchar("end"), size=10)
        return q

    def _update_fields(self):
        stg = self.doc.settings.element
        uf = stg.find(qn("w:updateFields"))
        if uf is None:
            uf = OxmlElement("w:updateFields")
            after = None
            for tag in ("w:compat", "w:docVars", "w:rsids", "w:themeFontLang"):
                after = stg.find(qn(tag))
                if after is not None:
                    break
            if after is None:
                stg.append(uf)
            else:
                after.addprevious(uf)
        uf.set(qn("w:val"), "true")
        z = stg.find(qn("w:zoom"))
        if z is not None and z.get(qn("w:percent")) is None:
            z.set(qn("w:percent"), "100")

    def build_toc(self, pagemap=None):
        if self._toc_anchor is None:
            return
        pages = {}
        if pagemap and os.path.exists(pagemap):
            with open(pagemap, encoding="utf-8") as f:
                pages = json.load(f)
        lines = [self._toc_line(lv, tx, nm, pages)
                 for lv, tx, nm in self.headings]
        anchor = self._toc_anchor
        for q in lines:
            anchor._p.addprevious(q._p)
        anchor._p.getparent().remove(anchor._p)
        self._update_fields()

    # ------------------------------------------------------------ 保存
    def _set_lang(self, lang="ja-JP"):
        """既定の言語を日本語にする。

        PDFに変換したときの言語の指定になり、スクリーンリーダーが
        読み上げる言語を取り違えない。
        """
        st = self.doc.styles["Normal"]
        rPr = st.element.get_or_add_rPr()
        for old in rPr.findall(qn("w:lang")):
            rPr.remove(old)
        e = OxmlElement("w:lang")
        e.set(qn("w:val"), "en-US")
        e.set(qn("w:eastAsia"), lang)
        rPr.append(e)

    def save(self, path):
        p = pathlib.Path(path)
        p.parent.mkdir(parents=True, exist_ok=True)
        # 文書の題名はPDFのしおり・ウィンドウの題・読み上げに使われる
        self.doc.core_properties.title = self.title
        self.doc.core_properties.language = "ja-JP"
        self._set_lang()
        self.doc.save(p)
        return p
