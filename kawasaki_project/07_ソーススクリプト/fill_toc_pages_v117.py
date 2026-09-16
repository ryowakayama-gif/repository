# -*- coding: utf-8 -*-
"""計画素案の目次に頁番号を入れる。

Ver.1.16 までは目次の頁番号が「－」のままであった。
LibreOffice で PDF に変換して各頁のフッター（─ n ─）を読み、
目次の各行の見出しが最初に現れる頁を突き止めて書き込む。

見出しを増減したとき、及び表や図を入れ替えたときは、本スクリプトを
実行し直す（頁の割付けが動くため）。
"""
import os
import re
import subprocess
import sys
import tempfile

import docx
import pypdf

sys.path.insert(0, "07_ソーススクリプト")
from fix_soan_v111 import set_tc  # noqa: E402

DOCX = "01_第10期_最新版成果品/川崎町_計画書素案_v1.17_中長期推計反映版.docx"
FOOT = re.compile(r"[─－―‐\-]\s*(\d+)\s*[─－―‐\-]")


def norm(s):
    """突き合わせ用に、空白・区切り記号・括弧を取り除く。"""
    return re.sub(r"[\s　．.・,，、（）()【】〔〕]", "", s)


def to_pdf(path, workdir):
    env = dict(os.environ, HOME=workdir)
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf",
                    "--outdir", workdir, path],
                   check=True, env=env, timeout=1800,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    out = [f for f in os.listdir(workdir) if f.endswith(".pdf")]
    if not out:
        raise RuntimeError("PDFが作られなかった")
    return os.path.join(workdir, out[0])


def pages(pdf):
    """(印字されている頁番号, 正規化した本文) の並びを返す。"""
    rd = pypdf.PdfReader(pdf)
    out = []
    for i, pg in enumerate(rd.pages):
        txt = pg.extract_text() or ""
        m = FOOT.findall(txt)
        out.append((int(m[-1]) if m else None, norm(txt), i))
    return out


def main():
    with tempfile.TemporaryDirectory() as wd:
        pdf = to_pdf(os.path.abspath(DOCX), wd)
        pg = pages(pdf)

    doc = docx.Document(DOCX)
    toc = doc.tables[0]
    # 目次そのものが載っている頁は突き合わせの対象から外す
    body_from = 0
    for no, txt, i in pg:
        if norm("第1章") in txt and norm("計画の策定にあたって") in txt:
            body_from = i
    body_from = max(body_from, 1)

    def page_of(key, frm):
        for no, txt, i in pg:
            if i < frm:
                continue
            if key and key in txt:
                return no if no is not None else i + 1
        return None

    # 節の行を先に解く。章の行は、その章に属する節の最小頁とする
    # （章題の語は本文中にも現れるため、本文の語に引かれないようにする）。
    rows = [r for r in toc.rows if r.cells[1].text.strip()]
    found = [None] * len(rows)
    for n, row in enumerate(rows):
        if row.cells[0].text.strip():          # 章の行は後で解く
            continue
        ttl = row.cells[1].text.strip()
        frm = body_from if not ttl.startswith("ご挨拶") else 0
        found[n] = page_of(norm(ttl), frm)
    for n, row in enumerate(rows):
        if not row.cells[0].text.strip():
            continue
        kids = []
        for m in range(n + 1, len(rows)):
            if rows[m].cells[0].text.strip():
                break
            if found[m] is not None:
                kids.append(found[m])
        found[n] = min(kids) if kids else page_of(norm(row.cells[1].text), body_from)

    hit = miss = 0
    for n, row in enumerate(rows):
        if found[n] is None:
            miss += 1
            print("  未特定：", row.cells[1].text.strip())
            continue
        set_tc(row.cells[2]._tc, str(found[n]))
        hit += 1

    doc.save(DOCX)
    print("頁番号を入れました：", DOCX)
    print(f"目次 {len(rows)}行　記入 {hit}件　未特定 {miss}件")
    print(f"総頁数 {len(pg)}　本文の開始 {body_from + 1}頁目（PDF）")


if __name__ == "__main__":
    main()
