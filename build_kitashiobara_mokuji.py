# -*- coding: utf-8 -*-
"""
北塩原村 第8期障がい福祉計画・第4期障がい児福祉計画
計画素案の目次整理・ページ番号付与ジェネレータ

入出力: output/北塩原村_計画素案.docx（上書き更新）

背景
  build_kitashiobara_kosshi_rev.py が原本（令和8年7月版）に章節を追加してきた結果、
  原本が持っていた目次（Word の目次フィールド）が本文と一致しなくなっていた。

    本文の見出し   94件（第1階層8・第2階層47・第3階層39）
    目次の項目     57件
    参照先が消えたブックマーク  3件（旧第3章の小見出し）
    残る54件も旧位置のまま      → 目次から飛ぶと違う場所に着く
    ページ番号は全項目「1」     → 一度も更新されていない

  本モジュールは目次を本文から作り直し、次を行う。

    1. 古い _Toc ブックマークを本文から除去する
    2. 各章（第1階層）の先頭で改ページする
    3. すべての見出しに新しいブックマークを付ける
    4. 目次を再生成する。各項目はブックマークへのハイパーリンクとし、
       ページ番号は PAGEREF フィールドで持たせる
    5. 本文のページ数を推定し、PAGEREF のキャッシュ値として書き込む
       （Word で開いた時点で数字が見え、フィールド更新で正確な値になる）
    6. 表紙・目次をノンブルの対象外とし、本文を1ページから始める
    7. Word で開いたときにフィールドが自動更新されるよう設定する

  目次に載せる階層は TOC_LEVELS で切り替える。現行計画（第7期）の目次が
  章＋節の2階層であったため、既定は2階層とする。
  3階層にすると94項目・約4ページとなり、本文60ページに対して過大になる。
"""

import copy
import math
import re

from docx import Document
from docx.oxml.ns import nsmap, qn
from docx.oxml import OxmlElement

SRC = "/home/user/repository/output/北塩原村_計画素案.docx"
OUT = SRC

# 目次に載せる見出しの階層（1=章のみ、2=章＋節、3=章＋節＋小見出し）
TOC_LEVELS = 2

# 目次から除外する見出し（作業用の記録であり計画の内容ではない）
TOC_EXCLUDE = ("本版での主な追加・修正",)

# 目次エントリのスタイル（原本が持つ Word 自動生成の目次スタイル）
TOC_STYLE = {1: "11", 2: "20", 3: "30"}
HYPERLINK_STYLE = "a5"
TOC_FONT = "BIZ UDPゴシック"
TAB_POS = 9628          # 右揃えタブ（ドットリーダー）の位置

# --- ページ推定のパラメータ（A4・余白2cm・行送り18pt・本文10.5pt）--------
LINES_PER_PAGE = 40     # (16838 - 1134*2) / 360
TWIPS_PER_CHAR = 210    # 10.5pt 全角1文字
TEXT_WIDTH = 11906 - 1134 * 2
CHARS_PER_LINE = TEXT_WIDTH // TWIPS_PER_CHAR    # 45
HEAD_LINES = {1: 3.0, 2: 2.2, 3: 1.8}            # 見出しは前後の空きを含む
TABLE_ROW_PAD = 0.35                             # 罫線・セル余白の分

BOOKMARK_BASE = 8100000


# ============================================================
# XML ヘルパー
# ============================================================
def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k.replace("_", ":")), str(v))
    return e


def _rpr(bold=False, style=None):
    rPr = OxmlElement("w:rPr")
    if style:
        rPr.append(_el("w:rStyle", w_val=style))
    f = _el("w:rFonts", w_ascii=TOC_FONT, w_eastAsia=TOC_FONT, w_hAnsi=TOC_FONT)
    rPr.append(f)
    if bold:
        rPr.append(OxmlElement("w:b"))
        rPr.append(OxmlElement("w:bCs"))
    rPr.append(OxmlElement("w:noProof"))
    return rPr


def _run(text=None, bold=False, style=None, tab=False):
    r = OxmlElement("w:r")
    r.append(_rpr(bold=bold, style=style))
    if tab:
        r.append(OxmlElement("w:tab"))
    if text is not None:
        t = OxmlElement("w:t")
        t.text = text
        t.set(qn("xml:space"), "preserve")
        r.append(t)
    return r


def _fld_run(kind, instr=None, bold=False):
    """fldChar（begin/separate/end）または instrText を持つ run を作る。"""
    r = OxmlElement("w:r")
    r.append(_rpr(bold=bold))
    if instr is not None:
        it = OxmlElement("w:instrText")
        it.text = instr
        it.set(qn("xml:space"), "preserve")
        r.append(it)
    else:
        r.append(_el("w:fldChar", w_fldCharType=kind))
    return r


def paragraph_text(p):
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def style_of(p):
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        return None
    s = pPr.find(qn("w:pStyle"))
    return s.get(qn("w:val")) if s is not None else None


def get_or_add_pPr(p):
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p.insert(0, pPr)
    return pPr


# ============================================================
# 1. 古いブックマークの除去
# ============================================================
def strip_old_bookmarks(doc):
    """原本が持つ _Toc ブックマークを本文から取り除く。

    追加した見出しには付いておらず、残っているものも旧位置を指しているため、
    作り直す前に一掃する。
    """
    body = doc.element.body
    removed_ids = set()
    for bs in list(body.iter(qn("w:bookmarkStart"))):
        name = bs.get(qn("w:name")) or ""
        if name.startswith("_Toc"):
            removed_ids.add(bs.get(qn("w:id")))
            bs.getparent().remove(bs)
    for be in list(body.iter(qn("w:bookmarkEnd"))):
        if be.get(qn("w:id")) in removed_ids:
            be.getparent().remove(be)
    return len(removed_ids)


# ============================================================
# 2. 章頭の改ページ
# ============================================================
def set_page_break_before(p):
    pPr = get_or_add_pPr(p)
    for old in pPr.findall(qn("w:pageBreakBefore")):
        pPr.remove(old)
    # pStyle の直後に置く（スキーマ上の順序）
    brk = OxmlElement("w:pageBreakBefore")
    ps = pPr.find(qn("w:pStyle"))
    if ps is not None:
        ps.addnext(brk)
    else:
        pPr.insert(0, brk)


# ============================================================
# 3. ページ数の推定
# ============================================================
def table_lines(tbl):
    """表の高さを行数で見積もる。列幅から1セルの折返し行数を求める。"""
    total = 0.0
    for tr in tbl.findall(qn("w:tr")):
        row_max = 1
        for tc in tr.findall(qn("w:tc")):
            tcPr = tc.find(qn("w:tcPr"))
            w = None
            if tcPr is not None:
                tcW = tcPr.find(qn("w:tcW"))
                if tcW is not None:
                    try:
                        w = int(tcW.get(qn("w:w")))
                    except (TypeError, ValueError):
                        w = None
            w = w or (TEXT_WIDTH // max(1, len(tr.findall(qn("w:tc")))))
            # 注記ボックスは9.5pt、表本文も9.5pt で組んでいる
            chars = max(1, w // 190)
            text = "".join(t.text or "" for t in tc.iter(qn("w:t")))
            row_max = max(row_max, math.ceil(len(text) / chars) if text else 1)
        total += row_max + TABLE_ROW_PAD
    return total


def estimate_pages(doc, body_start_index):
    """本文各ブロックのページ番号を推定する。

    戻り値は {ブロックの通し番号: ページ番号}。
    章（第1階層の見出し）で改ページする前提で積み上げる。
    """
    pages = {}
    page = 1
    used = 0.0
    for i, c in enumerate(doc.element.body.iterchildren()):
        if i < body_start_index:
            continue
        if c.tag == qn("w:p"):
            st = style_of(c)
            text = paragraph_text(c).strip()
            if st == "1":
                # 章頭で改ページ（本文の先頭章を除く）
                if not (page == 1 and used == 0):
                    page += 1
                    used = 0.0
                need = HEAD_LINES[1]
            elif st in ("2", "3"):
                need = HEAD_LINES[int(st)]
            else:
                need = max(1.0, math.ceil(len(text) / CHARS_PER_LINE)) if text else 1.0
        elif c.tag == qn("w:tbl"):
            need = table_lines(c)
        else:
            continue

        if used + need > LINES_PER_PAGE and used > 0:
            page += 1
            used = 0.0
        pages[i] = page
        used += need
    return pages


# ============================================================
# 4. 見出しの収集とブックマーク付与
# ============================================================
def collect_headings(doc, body_start_index, pages):
    """本文の見出しを順に集め、ブックマークを付ける。"""
    heads = []
    n = 0
    for i, c in enumerate(doc.element.body.iterchildren()):
        if c.tag != qn("w:p"):
            continue
        st = style_of(c)
        if st not in ("1", "2", "3"):
            continue
        text = paragraph_text(c).strip()
        if not text or text in TOC_EXCLUDE:
            continue
        if i < body_start_index:
            continue
        n += 1
        name = f"_Toc{BOOKMARK_BASE + n}"
        bid = str(BOOKMARK_BASE + n)
        # 見出し段落の内側を囲む（pPr の直後から末尾まで）
        pPr = c.find(qn("w:pPr"))
        bs = _el("w:bookmarkStart", w_id=bid, w_name=name)
        if pPr is not None:
            pPr.addnext(bs)
        else:
            c.insert(0, bs)
        c.append(_el("w:bookmarkEnd", w_id=bid))
        heads.append({
            "level": int(st),
            "text": text,
            "bookmark": name,
            "page": pages.get(i, 1),
            "index": i,
        })
    return heads


# ============================================================
# 5. 目次の再生成
# ============================================================
def build_toc_paragraph(h, first, levels):
    """目次の1行を作る。ハイパーリンク＋タブ＋PAGEREFフィールド。"""
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    pPr.append(_el("w:pStyle", w_val=TOC_STYLE.get(h["level"], "20")))
    tabs = OxmlElement("w:tabs")
    tabs.append(_el("w:tab", w_val="right", w_leader="dot", w_pos=TAB_POS))
    pPr.append(tabs)
    pPr.append(_rpr())
    p.append(pPr)

    if first:
        # 目次フィールドの開始。Word はこの命令で目次全体を作り直す
        p.append(_fld_run("begin"))
        p.append(_fld_run(None, instr=f'TOC \\h \\z \\u \\o "1-{levels}"'))
        p.append(_fld_run("separate"))

    bold = h["level"] == 1
    link = _el("w:hyperlink", w_anchor=h["bookmark"], w_history="1")
    link.append(_run(h["text"], bold=bold, style=HYPERLINK_STYLE))
    link.append(_run(tab=True, bold=bold))
    link.append(_fld_run("begin", bold=bold))
    link.append(_fld_run(None, instr=f' PAGEREF {h["bookmark"]} \\h ', bold=bold))
    link.append(_fld_run("separate", bold=bold))
    link.append(_run(str(h["page"]), bold=bold))   # キャッシュ値（更新前に見える数字）
    link.append(_fld_run("end", bold=bold))
    p.append(link)
    return p


def rebuild_toc(doc, heads, levels):
    """目次の内容管理コントロール（sdt）の中身を作り直す。"""
    body = doc.element.body
    sdt = None
    for c in body.iterchildren():
        if c.tag == qn("w:sdt"):
            sdt = c
            break
    if sdt is None:
        raise LookupError("目次の内容管理コントロールが見つかりません")

    content = sdt.find(qn("w:sdtContent"))
    for child in list(content):
        content.remove(child)

    entries = [h for h in heads if h["level"] <= levels]
    for k, h in enumerate(entries):
        content.append(build_toc_paragraph(h, first=(k == 0), levels=levels))

    # 目次フィールドの終了
    tail = OxmlElement("w:p")
    tailPr = OxmlElement("w:pPr")
    tailPr.append(_rpr())
    tail.append(tailPr)
    tail.append(_fld_run("end"))
    content.append(tail)
    return len(entries)


# ============================================================
# 6. セクション分割（表紙・目次はノンブルなし、本文は1から）
# ============================================================
def split_sections(doc, body_start_index):
    """本文の直前でセクションを区切り、本文のページ番号を1から始める。"""
    body = doc.element.body
    body_sect = body.find(qn("w:sectPr"))
    if body_sect is None:
        raise LookupError("本文末尾の sectPr が見つかりません")

    # 前付け用の sectPr（本文 sectPr の複製。既定フッターを空フッターに差し替える）
    front = copy.deepcopy(body_sect)
    for fr in front.findall(qn("w:footerReference")):
        if fr.get(qn("w:type")) == "default":
            fr.set(qn("r:id"), "rId19")      # footer2.xml（空）
    for hr in front.findall(qn("w:headerReference")):
        if hr.get(qn("w:type")) == "default":
            hr.set(qn("r:id"), "rId18")      # header2.xml

    blocks = list(body.iterchildren())
    anchor = None
    for i in range(body_start_index - 1, -1, -1):
        if blocks[i].tag == qn("w:p"):
            anchor = blocks[i]
            break
    if anchor is None:
        raise LookupError("前付けの最終段落が見つかりません")
    pPr = get_or_add_pPr(anchor)
    for old in pPr.findall(qn("w:sectPr")):
        pPr.remove(old)
    pPr.append(front)

    # 本文セクションは1ページ目から番号を振り直す
    for old in body_sect.findall(qn("w:pgNumType")):
        body_sect.remove(old)
    pgnum = _el("w:pgNumType", w_start="1")
    cols = body_sect.find(qn("w:cols"))
    if cols is not None:
        cols.addprevious(pgnum)
    else:
        body_sect.append(pgnum)
    # 本文は表紙ではないので titlePg を外す
    for tp in body_sect.findall(qn("w:titlePg")):
        body_sect.remove(tp)


# ============================================================
# 7. フィールドの自動更新
# ============================================================
def set_update_fields(doc):
    st = doc.settings.element
    for old in st.findall(qn("w:updateFields")):
        st.remove(old)
    st.insert(0, _el("w:updateFields", w_val="true"))


# ============================================================
def find_body_start(doc):
    """最初の章（第1階層の見出し）のブロック位置を返す。"""
    for i, c in enumerate(doc.element.body.iterchildren()):
        if c.tag == qn("w:p") and style_of(c) == "1":
            return i
    raise LookupError("第1階層の見出しが見つかりません")


def main():
    doc = Document(SRC)

    removed = strip_old_bookmarks(doc)
    body_start = find_body_start(doc)

    # 章頭で改ページ
    n_break = 0
    for c in doc.element.body.iterchildren():
        if c.tag == qn("w:p") and style_of(c) == "1" and paragraph_text(c).strip():
            set_page_break_before(c)
            n_break += 1

    pages = estimate_pages(doc, body_start)
    heads = collect_headings(doc, body_start, pages)
    n_toc = rebuild_toc(doc, heads, TOC_LEVELS)
    split_sections(doc, body_start)
    set_update_fields(doc)
    doc.save(OUT)

    total = max(pages.values()) if pages else 0
    lv = {1: 0, 2: 0, 3: 0}
    for h in heads:
        lv[h["level"]] += 1

    print(f"更新: {OUT}")
    print(f"  古いブックマークを除去: {removed}件")
    print(f"  章頭の改ページ: {n_break}件")
    print(f"  ブックマークを付与した見出し: {len(heads)}件"
          f"（第1階層{lv[1]}・第2階層{lv[2]}・第3階層{lv[3]}）")
    print(f"  目次に掲載: {n_toc}件（{TOC_LEVELS}階層まで）")
    print(f"  本文の推定ページ数: {total}ページ")
    print()
    print("  章別の推定ページ")
    for h in heads:
        if h["level"] == 1:
            print(f"    p.{h['page']:>3}  {h['text']}")
    return heads, pages, total


if __name__ == "__main__":
    main()
