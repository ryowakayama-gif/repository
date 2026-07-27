#!/usr/bin/env python3
"""目次構成の再編成メモを Word 文書として出力する。

提案書本体のパッケージ（styles.xml・フッター・セクション設定）をそのまま流用し、
document.xml だけを生成するので、書体・表の体裁が提案書と揃う。
"""
import html
import shutil
import subprocess
import sys
from pathlib import Path

import restructure as R

WORK = Path(__file__).resolve().parent
FONT = R.FONT
CONTENT_W = 9314


def esc(s):
    return html.escape(s, quote=False)


def _runs(text, base):
    """**強調** を太字ランに分けて組む。"""
    out = []
    for i, part in enumerate(text.split("**")):
        if not part:
            continue
        b = "<w:b/><w:bCs/>" if i % 2 else ""
        out.append(f'<w:r><w:rPr>{FONT}{b}{base}</w:rPr>'
                   f'<w:t xml:space="preserve">{esc(part)}</w:t></w:r>')
    return "".join(out)


def title(text, sub):
    return (
        f'<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:color="1F4E78" w:sz="10" '
        f'w:space="6"/></w:pBdr><w:spacing w:after="140"/></w:pPr>'
        f'<w:r><w:rPr>{FONT}<w:b/><w:bCs/><w:color w:val="1F4E78"/><w:sz w:val="32"/>'
        f'<w:szCs w:val="32"/></w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'
        f'<w:p><w:pPr><w:spacing w:after="320"/></w:pPr>'
        f'<w:r><w:rPr>{FONT}<w:color w:val="595959"/><w:sz w:val="19"/><w:szCs w:val="19"/>'
        f'</w:rPr><w:t xml:space="preserve">{esc(sub)}</w:t></w:r></w:p>'
    ).encode()


def h1(text):
    return (
        f'<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:color="1F4E78" w:sz="6" '
        f'w:space="4"/></w:pBdr><w:spacing w:after="160" w:before="360"/>'
        f'<w:keepNext/></w:pPr>'
        f'<w:r><w:rPr>{FONT}<w:b/><w:bCs/><w:color w:val="1F4E78"/><w:sz w:val="26"/>'
        f'<w:szCs w:val="26"/></w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'
    ).encode()


def h2(text):
    return (
        f'<w:p><w:pPr><w:spacing w:after="90" w:before="240"/><w:keepNext/></w:pPr>'
        f'<w:r><w:rPr>{FONT}<w:b/><w:bCs/><w:color w:val="262626"/><w:sz w:val="21"/>'
        f'<w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'
    ).encode()


def h3(text):
    return (
        f'<w:p><w:pPr><w:spacing w:after="70" w:before="160"/><w:keepNext/></w:pPr>'
        f'<w:r><w:rPr>{FONT}<w:b/><w:bCs/><w:color w:val="1F4E78"/><w:sz w:val="20"/>'
        f'<w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'
    ).encode()


def p(text):
    base = '<w:sz w:val="20"/><w:szCs w:val="20"/>'
    return (f'<w:p><w:pPr><w:spacing w:after="120" w:line="290"/><w:jc w:val="left"/></w:pPr>'
            f'{_runs(text, base)}</w:p>').encode()


def li(text, level=0):
    base = '<w:sz w:val="20"/><w:szCs w:val="20"/>'
    ind = 340 + level * 300
    return (f'<w:p><w:pPr><w:spacing w:after="60" w:line="280"/>'
            f'<w:ind w:left="{ind}" w:hanging="220"/></w:pPr>'
            f'{_runs("・" + text, base)}</w:p>').encode()


def num(n, text):
    base = '<w:sz w:val="20"/><w:szCs w:val="20"/>'
    return (f'<w:p><w:pPr><w:spacing w:after="60" w:line="280"/>'
            f'<w:ind w:left="420" w:hanging="300"/></w:pPr>'
            f'{_runs(f"{n}. " + text, base)}</w:p>').encode()


def note(text):
    base = '<w:i/><w:iCs/><w:color w:val="595959"/><w:sz w:val="18"/><w:szCs w:val="18"/>'
    return (f'<w:p><w:pPr><w:spacing w:after="140" w:before="40"/></w:pPr>'
            f'{_runs(text, base)}</w:p>').encode()


def _cell(text, w, header, align_center):
    shd = '<w:shd w:fill="44546A" w:val="clear"/>' if header else ""
    rpr = (f'{FONT}<w:b/><w:bCs/><w:color w:val="FFFFFF"/><w:sz w:val="18"/><w:szCs w:val="18"/>'
           if header else f'{FONT}<w:sz w:val="18"/><w:szCs w:val="18"/>')
    jc = '<w:jc w:val="center"/>' if (header or align_center) else ""
    borders = ('<w:tcBorders><w:top w:val="single" w:color="BFBFBF" w:sz="4"/>'
               '<w:left w:val="single" w:color="BFBFBF" w:sz="4"/>'
               '<w:bottom w:val="single" w:color="BFBFBF" w:sz="4"/>'
               '<w:right w:val="single" w:color="BFBFBF" w:sz="4"/></w:tcBorders>')
    mar = ('<w:tcMar><w:top w:type="dxa" w:w="60"/><w:left w:type="dxa" w:w="100"/>'
           '<w:bottom w:type="dxa" w:w="60"/><w:right w:type="dxa" w:w="100"/></w:tcMar>')
    body = "".join(
        f'<w:p><w:pPr>{jc}<w:spacing w:after="0" w:line="260"/></w:pPr>'
        f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc(line)}</w:t></w:r></w:p>'
        for line in (text.split("\n") if text else [""]))
    return (f'<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="{w}"/>{borders}{shd}{mar}'
            f'<w:vAlign w:val="center"/></w:tcPr>{body}</w:tc>')


def table(rows, widths, center_cols=()):
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    total = sum(widths)
    trs = []
    for i, row in enumerate(rows):
        hdr = '<w:trPr><w:tblHeader/></w:trPr>' if i == 0 else ""
        tcs = "".join(_cell(c, w, i == 0, j in center_cols)
                      for j, (c, w) in enumerate(zip(row, widths)))
        trs.append(f'<w:tr>{hdr}{tcs}</w:tr>')
    return (
        f'<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="{total}"/><w:tblBorders>'
        f'<w:top w:val="single" w:color="auto" w:sz="4"/>'
        f'<w:left w:val="single" w:color="auto" w:sz="4"/>'
        f'<w:bottom w:val="single" w:color="auto" w:sz="4"/>'
        f'<w:right w:val="single" w:color="auto" w:sz="4"/>'
        f'<w:insideH w:val="single" w:color="auto" w:sz="4"/>'
        f'<w:insideV w:val="single" w:color="auto" w:sz="4"/></w:tblBorders></w:tblPr>'
        f'<w:tblGrid>{grid}</w:tblGrid>{"".join(trs)}</w:tbl>'
        f'<w:p><w:pPr><w:spacing w:after="60" w:line="120"/></w:pPr></w:p>'
    ).encode()


def build(out_path):
    import content
    body = b"".join(content.BLOCKS)
    xml = R.HEAD + body + R.BLOCKS[201] + R.TAIL

    build_dir = WORK / "memo_build"
    if build_dir.exists():
        shutil.rmtree(build_dir)
    shutil.copytree(R.UNPACKED, build_dir)
    (build_dir / "word" / "document.xml").write_bytes(xml)

    out = Path(out_path)
    if out.exists():
        out.unlink()
    subprocess.run(["zip", "-Xrq", str(out.resolve()), "."], cwd=build_dir, check=True)
    return out


if __name__ == "__main__":
    print("wrote", build(sys.argv[1] if len(sys.argv) > 1 else "memo.docx"))
