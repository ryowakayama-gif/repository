#!/usr/bin/env python3
"""Rebuild word/document.xml from a chapter plan (plan.py) without touching formatting.

Every element that lands in the new body is either an untouched original block or a
clone of an original block with its text swapped, so run/paragraph properties are
preserved exactly.
"""
import html
import re
import shutil
import subprocess
import sys
from pathlib import Path

import docxblocks as db

WORK = Path(__file__).resolve().parent
ROOT = WORK.parent
SRC_DOCX = ROOT / "src" / "足寄町企画提案書_原本.docx"
UNPACKED = WORK / "unpacked"
SRC = UNPACKED / "word" / "document.xml"

if not SRC.exists():
    import zipfile
    with zipfile.ZipFile(SRC_DOCX) as z:
        z.extractall(UNPACKED)

FONT = ('<w:rFonts w:ascii="Yu Gothic" w:cs="Yu Gothic" w:eastAsia="Yu Gothic" '
        'w:hAnsi="Yu Gothic"/>')

PAGEBREAK = b'<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

_DATA = SRC.read_bytes()
HEAD, BLOCKS, TAIL = db.split_body(_DATA)
text_of = db.text_of


def esc(s):
    return html.escape(s, quote=False)


def chapter_heading(text, bookmark_id, bookmark_name):
    """Chapter heading paragraph (blue, bottom border, 14pt bold) with a bookmark."""
    return (
        f'<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:color="1F4E78" w:sz="6" '
        f'w:space="4"/></w:pBdr><w:spacing w:after="160" w:before="320"/></w:pPr>'
        f'<w:bookmarkStart w:id="{bookmark_id}" w:name="{bookmark_name}"/>'
        f'<w:r><w:rPr>{FONT}<w:b/><w:bCs/><w:color w:val="1F4E78"/><w:sz w:val="28"/>'
        f'<w:szCs w:val="28"/></w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r>'
        f'<w:bookmarkEnd w:id="{bookmark_id}"/></w:p>'
    ).encode()


def subheading(text):
    """（１）-level subheading: bold, 11pt, near-black."""
    return (
        f'<w:p><w:pPr><w:spacing w:after="100" w:before="220"/></w:pPr>'
        f'<w:r><w:rPr>{FONT}<w:b/><w:bCs/><w:color w:val="262626"/><w:sz w:val="22"/>'
        f'<w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'
    ).encode()


def body_para(text):
    """Plain body paragraph, cloned from the document's existing body style."""
    return (
        f'<w:p><w:pPr><w:spacing w:after="140" w:line="300"/><w:jc w:val="left"/></w:pPr>'
        f'<w:r><w:rPr>{FONT}<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
        f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'
    ).encode()


def caption(text):
    """■-style blue caption used above tables (cloned from the ■町民アンケート line)."""
    return (
        f'<w:p><w:pPr><w:spacing w:after="80" w:before="140"/></w:pPr>'
        f'<w:r><w:rPr>{FONT}<w:b/><w:bCs/><w:color w:val="1F4E78"/><w:sz w:val="20"/>'
        f'<w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'
    ).encode()


def bullet(text):
    """Hanging-indent bullet line (cloned from the ・数値目標… lines)."""
    return (
        f'<w:p><w:pPr><w:spacing w:after="60"/><w:ind w:left="340" w:hanging="220"/></w:pPr>'
        f'<w:r><w:rPr>{FONT}<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
        f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'
    ).encode()


def bookmark(block, bookmark_id, name):
    """Wrap an existing paragraph's content in a bookmark (for PAGEREF targets)."""
    if not block.startswith(b"<w:p>") and not block.startswith(b"<w:p "):
        raise ValueError("bookmark target must be a paragraph")
    start = f'<w:bookmarkStart w:id="{bookmark_id}" w:name="{name}"/>'.encode()
    end = f'<w:bookmarkEnd w:id="{bookmark_id}"/>'.encode()
    m = re.match(rb"(<w:p(?:\s[^>]*)?>)(<w:pPr>.*?</w:pPr>)?", block, re.S)
    insert_at = m.end()
    return block[:insert_at] + start + block[insert_at:-len(b"</w:p>")] + end + b"</w:p>"


def toc_line(label, bookmark_name, cached_page):
    """Manual-look TOC line whose page number is a self-updating PAGEREF field."""
    rpr = f'<w:rPr>{FONT}<w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>'
    return (
        f'<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="9026" w:leader="dot"/></w:tabs>'
        f'<w:spacing w:after="50" w:before="70"/></w:pPr>'
        f'<w:r>{rpr}<w:t xml:space="preserve">{esc(label)}</w:t></w:r>'
        f'<w:r>{rpr}<w:t xml:space="preserve">\t</w:t></w:r>'
        f'<w:r>{rpr}<w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r>{rpr}<w:instrText xml:space="preserve"> PAGEREF {bookmark_name} \\h </w:instrText></w:r>'
        f'<w:r>{rpr}<w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r>{rpr}<w:t>{cached_page}</w:t></w:r>'
        f'<w:r>{rpr}<w:fldChar w:fldCharType="end"/></w:r></w:p>'
    ).encode()


def toc_sub_line(label, bookmark_name, cached_page):
    """Indented second-level TOC line (used only for the two weighted chapters)."""
    rpr = f'<w:rPr>{FONT}<w:sz w:val="20"/><w:szCs w:val="20"/><w:color w:val="404040"/></w:rPr>'
    return (
        f'<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="9026" w:leader="dot"/></w:tabs>'
        f'<w:spacing w:after="0"/><w:ind w:firstLine="221"/></w:pPr>'
        f'<w:r>{rpr}<w:t xml:space="preserve">{esc(label)}</w:t></w:r>'
        f'<w:r>{rpr}<w:t xml:space="preserve">\t</w:t></w:r>'
        f'<w:r>{rpr}<w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r>{rpr}<w:instrText xml:space="preserve"> PAGEREF {bookmark_name} \\h </w:instrText></w:r>'
        f'<w:r>{rpr}<w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r>{rpr}<w:t>{cached_page}</w:t></w:r>'
        f'<w:r>{rpr}<w:fldChar w:fldCharType="end"/></w:r></w:p>'
    ).encode()


def retext(block, old, new):
    """Replace visible text inside a block, matching across run boundaries is not
    attempted -- old must sit inside a single <w:t>."""
    o = esc(old).encode()
    n = esc(new).encode()
    if o not in block:
        raise ValueError(f"text not found for replace: {old!r}")
    return block.replace(o, n)


def set_update_fields(settings_path):
    data = settings_path.read_bytes()
    if b"<w:updateFields" in data:
        return
    data = re.sub(rb"(<w:settings[^>]*>)", rb'\1<w:updateFields w:val="true"/>', data, count=1)
    settings_path.write_bytes(data)


def build(plan, out_docx):
    head, blocks, tail = HEAD, BLOCKS, TAIL

    new = []
    for item in plan:
        if isinstance(item, int):
            new.append(blocks[item])
        elif isinstance(item, bytes):
            new.append(item)
        elif isinstance(item, tuple):
            idx, old, rep = item
            new.append(retext(blocks[idx], old, rep))
        else:
            raise TypeError(type(item))
    out_xml = db.assemble(head, new, tail)

    build_dir = WORK / "build"
    if build_dir.exists():
        shutil.rmtree(build_dir)
    shutil.copytree(UNPACKED, build_dir)
    (build_dir / "word" / "document.xml").write_bytes(out_xml)
    set_update_fields(build_dir / "word" / "settings.xml")

    out = Path(out_docx)
    if out.exists():
        out.unlink()
    subprocess.run(["zip", "-Xrq", str(out.resolve()), "."], cwd=build_dir, check=True)
    return out


def page_map(pdf_path):
    """Map each rendered page number to its first line, for locating chapters."""
    from pypdf import PdfReader
    r = PdfReader(pdf_path)
    pages = []
    for p in r.pages:
        t = (p.extract_text() or "")
        lines = [l for l in t.split("\n") if l.strip()]
        pages.append(lines)
    return pages


if __name__ == "__main__":
    import plan as planmod
    out = build(planmod.build_plan(), sys.argv[1] if len(sys.argv) > 1 else "restructured.docx")
    print("wrote", out)
