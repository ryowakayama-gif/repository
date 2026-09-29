# -*- coding: utf-8 -*-
"""公表版の目次に入れるページ番号を、組版結果から求める.

build_toc_pages.py と同じ考え方を公表版に当てたものである。
公表版の目次はブックマークへのリンクとページ番号のフィールドで構成する。
Wordで開いた時点でフィールドは更新されるが、
更新前に表示する値を組版結果から求めておく。

値は output/_toc_pages_public.json に書き、build_plan_public.py が読む。

前提
  LibreOffice（soffice）と PyMuPDF が使えること。
  使えない場合、ページ番号は「－」と表示され、
  Wordで開いた時点で正しい値に更新される。
"""

import json
import os
import re
import subprocess
import sys
import tempfile

import pymupdf

BASE = os.path.dirname(os.path.abspath(__file__))
DOCX = os.path.join(BASE, "output", "第10期介護保険事業計画_公表版.docx")
PAGEMAP = os.path.join(BASE, "output", "_toc_pages_public.json")
BUILD = os.path.join(BASE, "build_plan_public.py")

FOOT = re.compile(r"－\s*(\d+)\s*－")


def norm(s):
    return re.sub(r"\s|　", "", s)


def build():
    subprocess.run([sys.executable, BUILD], check=True,
                   stdout=subprocess.DEVNULL)


def to_pdf(workdir):
    env = dict(os.environ, HOME=workdir)
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf",
         "--outdir", workdir, DOCX],
        check=True, env=env, stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL, timeout=1800)
    out = [f for f in os.listdir(workdir) if f.endswith(".pdf")]
    if not out:
        raise RuntimeError("PDFが作られなかった")
    return os.path.join(workdir, out[0])


def headings():
    """公表版の見出しの一覧（level, text, bookmark）を得る。"""
    import runpy
    import io
    import contextlib
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        g = runpy.run_path(BUILD)
    return g["HEADINGS"]


def pagemap(pdf, heads):
    d = pymupdf.open(pdf)
    text = [d[i].get_text() for i in range(d.page_count)]
    printed = {}
    for i, t in enumerate(text):
        lines = [x.strip() for x in t.splitlines() if x.strip()]
        if lines:
            m = FOOT.fullmatch(lines[-1].replace(" ", " "))
            if m:
                printed[i] = int(m.group(1))
    if not printed:
        raise RuntimeError("フッターのページ番号が読めなかった")
    body_from = min(printed)
    ntext = [norm(t) for t in text]
    res, cur = {}, body_from
    for level, txt, name in heads:
        key = norm(txt)
        hit = None
        for i in range(cur, len(ntext)):
            if key in ntext[i]:
                hit = i
                break
        if hit is None:
            hit = cur
        res[name] = printed.get(hit, printed[body_from])
        cur = hit
    return res


def main():
    if not os.path.exists(DOCX):
        build()
    heads = headings()
    prev = None
    for i in range(1, 6):
        build()
        with tempfile.TemporaryDirectory() as w:
            pdf = to_pdf(w)
            cur = pagemap(pdf, heads)
        print("%d回目　%d件　最終ページ %d" % (i, len(cur), max(cur.values())))
        if cur == prev:
            print("ページ番号が安定した")
            break
        prev = cur
        with open(PAGEMAP, "w", encoding="utf-8") as f:
            json.dump(cur, f, ensure_ascii=False, indent=1)
    build()
    print("saved: %s" % os.path.relpath(PAGEMAP, BASE))


if __name__ == "__main__":
    main()
