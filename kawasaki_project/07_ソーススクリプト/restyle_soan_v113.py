# -*- coding: utf-8 -*-
"""
計画素案 v1.12 → v1.13（体裁の統一。大雪地区広域連合の第10期計画素案の様式に合わせる）

■ 直した不具合
  素案の網掛けはすべて <w:shd w:val="solid" w:fill="…"/> で指定されていた。
  w:val="solid" は「前景色で塗りつぶす」指定であり、前景色（w:color）を省略すると
  既定の auto＝黒になる。このため、意図した淡い水色・淡い緑ではなく
  黒ベタで印刷・表示されていた。章題・▌小見出し・●コールアウト・表の見出し行が
  これに該当する（表の見出し行は白文字のため、黒地に白文字となっていた）。

■ 合わせた様式（大雪地区広域連合 第10期介護保険事業計画 協議用素案 令和8年8月）
  ・網掛け・段落罫線を用いない
  ・文字色を用いない（すべて自動＝黒）
  ・章見出し15pt／節見出し12.5pt／小見出し11pt／本文10.5pt／注記8.5pt
  ・本文・表の本文は太字にしない（見出しと表の見出し行のみ太字）
  ・表は罫線のみ（見出し行は太字。網掛けなし）
  フォントは両文書とも游ゴシックで一致しているため変更していない。

配色を残したい場合は、w:val を "clear" に改めるだけでも黒ベタは解消する。
その版が必要な場合は本スクリプトの STRIP_SHADING を False にする。
"""
import re
import sys

import docx
from docx.oxml.ns import qn

sys.path.insert(0, "07_ソーススクリプト")
from fix_soan_v111 import set_text  # noqa: E402

SRC = "01_第10期_最新版成果品/川崎町_計画書素案_v1.12_地域マネジメント章新設版.docx"
DST = "01_第10期_最新版成果品/川崎町_計画書素案_v1.13_体裁統一版.docx"

STRIP_SHADING = True        # False にすると網掛けを残し、黒ベタだけを直す
FRONT_MATTER = 23           # 表紙・挨拶・目次（この番号未満は大きさを保つ）
MARU = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫"


# ------------------------------------------------------------------ 部品
def get_pPr(p, create=False):
    pr = p._p.find(qn("w:pPr"))
    if pr is None and create:
        pr = p._p.makeelement(qn("w:pPr"), {})
        p._p.insert(0, pr)
    return pr


def par_shading(p):
    pr = get_pPr(p)
    if pr is None:
        return None
    s = pr.find(qn("w:shd"))
    return s.get(qn("w:fill")) if s is not None else None


def par_jc(p):
    pr = get_pPr(p)
    if pr is None:
        return None
    j = pr.find(qn("w:jc"))
    return j.get(qn("w:val")) if j is not None else None


def first_sz(p):
    for r in p._p.findall(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        if rpr is not None:
            s = rpr.find(qn("w:sz"))
            if s is not None:
                return s.get(qn("w:val"))
    return None


def strip_decoration(pr):
    """段落の網掛けと罫線を外す。"""
    if pr is None:
        return
    for tag in ("w:shd", "w:pBdr"):
        e = pr.find(qn(tag))
        if e is not None:
            pr.remove(e)


def set_jc(pr, val):
    if pr is None:
        return
    j = pr.find(qn("w:jc"))
    if val is None:
        if j is not None:
            pr.remove(j)
        return
    if j is None:
        j = pr.makeelement(qn("w:jc"), {})
        pr.append(j)
    j.set(qn("w:val"), val)


def set_spacing(pr, before=None, after=None):
    if pr is None:
        return
    sp = pr.find(qn("w:spacing"))
    if sp is None:
        sp = pr.makeelement(qn("w:spacing"), {})
        pr.append(sp)
    if before is not None:
        sp.set(qn("w:before"), str(before))
    if after is not None:
        sp.set(qn("w:after"), str(after))


def format_runs(el, sz=None, bold=None, drop_color=True):
    """要素配下のすべての run に、大きさ・太字・文字色を適用する。"""
    for r in el.findall(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            rpr = r.makeelement(qn("w:rPr"), {})
            r.insert(0, rpr)
        if drop_color:
            for tag in ("w:color", "w:highlight"):
                e = rpr.find(qn(tag))
                if e is not None:
                    rpr.remove(e)
            sh = rpr.find(qn("w:shd"))
            if sh is not None:
                rpr.remove(sh)
        if sz is not None:
            for tag, val in (("w:sz", sz), ("w:szCs", sz)):
                e = rpr.find(qn(tag))
                if e is None:
                    e = rpr.makeelement(qn(tag), {})
                    rpr.append(e)
                e.set(qn("w:val"), str(val))
        if bold is not None:
            for tag in ("w:b", "w:bCs"):
                e = rpr.find(qn(tag))
                if bold:
                    if e is None:
                        e = rpr.makeelement(qn(tag), {})
                        rpr.append(e)
                    e.set(qn("w:val"), "1")
                elif e is not None:
                    rpr.remove(e)


# ------------------------------------------------------------------ 段落
def restyle_paragraphs(doc):
    ps = doc.paragraphs
    kinds = []
    prev = None
    for i, p in enumerate(ps):
        txt = p.text.strip()
        shd = par_shading(p)
        jc = par_jc(p)
        sz = first_sz(p)

        if i < FRONT_MATTER:
            kind = "front"
        elif re.fullmatch(r"第\s*\d+\s*章", txt):
            kind = "ch"
        elif prev == "ch" and txt:
            kind = "chtitle"
        elif re.match(r"^\d+-\d+　", txt):
            kind = "sec"
        elif txt.startswith("▌"):
            kind = "sub"
        elif txt.startswith("●"):
            kind = "call"
        elif txt.startswith("⚠"):
            kind = "warn"
        elif txt.startswith(("出典", "資料：")):
            kind = "note"
        elif txt.startswith("図") and jc == "center":
            kind = "fig"
        elif txt and txt[0] in MARU:
            kind = "num"
        elif txt.startswith("※"):
            kind = "note2"
        elif jc == "center" and txt:
            kind = "display"
        else:
            kind = "body"

        pr = get_pPr(p, create=bool(txt))
        if STRIP_SHADING:
            strip_decoration(pr)

        if kind == "front":
            format_runs(p._p, bold=(sz is not None and int(sz) > 21) or None)
        elif kind == "ch":
            format_runs(p._p, sz=30, bold=True)
            set_jc(pr, None)
            set_spacing(pr, before=240, after=0)
        elif kind == "chtitle":
            format_runs(p._p, sz=30, bold=True)
            set_jc(pr, None)
            set_spacing(pr, before=0, after=140)
        elif kind == "sec":
            format_runs(p._p, sz=25, bold=True)
            set_jc(pr, None)
            set_spacing(pr, before=200, after=100)
        elif kind == "sub":
            format_runs(p._p, sz=22, bold=True)
            set_spacing(pr, before=180, after=80)
        elif kind == "call":
            format_runs(p._p, sz=20, bold=True)
            set_spacing(pr, before=80, after=80)
        elif kind == "warn":
            format_runs(p._p, sz=19, bold=False)
            set_spacing(pr, before=80, after=80)
        elif kind == "note":
            format_runs(p._p, sz=17, bold=False)
            set_jc(pr, "right")
        elif kind == "note2":
            format_runs(p._p, sz=18, bold=False)
        elif kind == "fig":
            format_runs(p._p, sz=18, bold=True)
        elif kind == "num":
            format_runs(p._p, sz=21, bold=False)
        elif kind == "display":
            format_runs(p._p, sz=min(int(sz or 22), 26), bold=True)
        else:
            format_runs(p._p, sz=21, bold=False)

        kinds.append(kind)
        if txt:
            prev = kind
    return kinds


# ------------------------------------------------------------------ 表
def restyle_tables(doc):
    n_head = n_cell = 0
    for t in doc.tables:
        rows = t.rows
        head = (len(rows) >= 2 and len(t.columns) >= 2
                and all(c.text.strip() for c in rows[0].cells))
        for ri, row in enumerate(rows):
            for cell in row.cells:
                tcPr = cell._tc.find(qn("w:tcPr"))
                if STRIP_SHADING and tcPr is not None:
                    s = tcPr.find(qn("w:shd"))
                    if s is not None:
                        tcPr.remove(s)
                for p in cell.paragraphs:
                    strip_decoration(get_pPr(p))
                    format_runs(p._p, bold=(head and ri == 0))
                    n_cell += 1
        if head:
            n_head += 1
    return n_head, n_cell


# ------------------------------------------------------------------ 版数
REPLACE = {
    21: "※本素案Ver.1.13は、令和8年6〜7月に実施した2調査（一般高齢者576件・"
        "要支援要介護認定者142件）の結果、令和8年9月1日に町からご提供いただいた"
        "実績データ、保険者機能強化推進交付金等の評価結果、及び令和8年9月2日の"
        "第1回策定委員会でのご意見を反映した版です。"
        "本版では、網掛けが黒く印刷される不具合を是正し、"
        "体裁を印刷に適した様式に統一しました。"
        "第9章に掲載している保険料の概算試算は確定値ではなく仮試算であり、"
        "第3回策定委員会（令和9年1月）で確定します。",
}


def main():
    doc = docx.Document(SRC)
    ps = doc.paragraphs

    for i, t in REPLACE.items():
        set_text(ps[i], t)

    kinds = restyle_paragraphs(doc)
    n_head, n_cell = restyle_tables(doc)

    n_v = 0
    for p in doc.paragraphs:
        if "Ver.1.12" in p.text:
            set_text(p, p.text.replace("Ver.1.12", "Ver.1.13"))
            n_v += 1

    doc.save(DST)

    import collections
    c = collections.Counter(kinds)
    d2 = docx.Document(DST)
    print(f"保存：{DST}")
    print("  段落の種類別件数：" + "／".join(f"{k} {v}" for k, v in c.most_common()))
    print(f"  表 {len(doc.tables)}（うち見出し行あり {n_head}）／セル内段落 {n_cell}")
    print(f"  版数の表記 {n_v}箇所")
    print(f"  段落 {len(d2.paragraphs)}／表 {len(d2.tables)}／図 {len(d2.inline_shapes)}")

    # 検証
    n_shd = d2.element.body.xml.count('w:shd')
    n_bdr = d2.element.body.xml.count('w:pBdr')
    n_col = d2.element.body.xml.count('<w:color ')
    print(f"  残存：w:shd {n_shd}件／w:pBdr {n_bdr}件／w:color {n_col}件")


if __name__ == "__main__":
    main()
