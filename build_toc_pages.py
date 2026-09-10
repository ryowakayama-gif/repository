# -*- coding: utf-8 -*-
"""計画素案の目次に入れるページ番号を、組版結果から求める.

計画素案の目次は、本文のブックマークへのリンクと
ページ番号のフィールド（PAGEREF）で構成している。
リンクはフィールドではないため更新しなくても機能するが、
ページ番号はフィールドであり、Wordで開いた時点で更新される。

本スクリプトは、更新前に表示する値（フィールドの結果として
書き込んでおく値）を組版結果から求めるものである。
値は output/_toc_pages.json に書き、build_plan_draft.py が読む。

目次の行数が変わると本文の開始ページも動くため、
ページ番号が変わらなくなるまで組版と算定を繰り返す。

前提
  LibreOffice（soffice）と PyMuPDF が使えること。
  いずれも使えない場合、目次のページ番号は「－」と表示され、
  Wordで開いた時点で正しい値に更新される。

出力
  output/_toc_pages.json   ブックマーク名 → ページ番号
"""

import json
import os
import re
import subprocess
import sys
import tempfile

import pymupdf

BASE = os.path.dirname(os.path.abspath(__file__))
DOCX = os.path.join(BASE, "output",
                    "第10期介護保険事業計画_協議用素案_令和8年8月.docx")
PAGEMAP = os.path.join(BASE, "output", "_toc_pages.json")
BUILD = os.path.join(BASE, "build_plan_draft.py")

FOOT = re.compile(r"－\s*(\d+)\s*－")


def norm(s):
    """空白（全角を含む）を取り除いて突き合わせに用いる。"""
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
        stderr=subprocess.DEVNULL, timeout=1200)
    out = [f for f in os.listdir(workdir) if f.endswith(".pdf")]
    if not out:
        raise RuntimeError("PDFが作られなかった")
    return os.path.join(workdir, out[0])


def headings():
    """build_plan_draft.py が作る見出しの一覧を得る。"""
    import runpy
    import io
    buf, old = io.StringIO(), sys.stdout
    sys.stdout = buf
    try:
        g = runpy.run_path(BUILD)
    finally:
        sys.stdout = old
    return g["HEADINGS"]


def pagemap(pdf, heads):
    d = pymupdf.open(pdf)
    text = [d[i].get_text() for i in range(d.page_count)]
    # 各ページのフッターに印刷されるページ番号を拾う。
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
        if hit is None:                     # 見つからないときは前の見出しと同じ
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
        print("%d回目　%d件　最終ページ %d"
              % (i, len(cur), max(cur.values())))
        if cur == prev:
            print("ページ番号が安定した")
            break
        prev = cur
        with open(PAGEMAP, "w", encoding="utf-8") as f:
            json.dump(cur, f, ensure_ascii=False, indent=1)
    build()
    print("saved: %s" % os.path.relpath(PAGEMAP, BASE))
    for level, txt, name in heads[:6]:
        print("  %s%s … %s" % ("  " * (level - 1), txt, prev.get(name)))


if __name__ == "__main__":
    main()
