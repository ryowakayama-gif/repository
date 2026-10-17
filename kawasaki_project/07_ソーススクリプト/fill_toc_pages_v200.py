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

import sys as _s
DOCX = (_s.argv[1] if len(_s.argv) > 1 else
        "01_第10期_最新版成果品/川崎町_計画書素案_v2.0_施策体系書下ろし版.docx")
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
        """frm 頁以降で最初に現れる頁の (印字番号, PDF上の索引) を返す。"""
        for no, txt, i in pg:
            if i < frm:
                continue
            if key and key in txt:
                return (no if no is not None else i + 1), i
        return None, None

    # 節の行を先に解く。章の行は、その章に属する節の最小頁とする
    # （章題の語は本文中にも現れるため、本文の語に引かれないようにする）。
    rows = [r for r in toc.rows if r.cells[1].text.strip()]
    found = [None] * len(rows)
    # 節は計画書に現れる順に並んでいるため、探す位置を前の節の頁以降に限る。
    # そうしないと、節名と同じ語が本文中に先に現れたときにそちらを拾う
    # （例：3-3の表にある「第10章10-4（KPI管理表）」を10-4の頁としてしまう）。
    # ⚠ floor が body_from より前に戻ると、目次自身の頁を拾ってしまう。
    #    実際、Ver.2.3までは「ご挨拶（川崎町長）」の行が目次の頁で一致し、
    #    floor がその頁まで戻ったため、第1章〜第4章の24行がすべて
    #    目次の頁（3頁）になっていた。floor は body_from 未満に戻さない。
    floor = body_from
    for n, row in enumerate(rows):
        if row.cells[0].text.strip():          # 章の行は後で解く
            continue
        ttl = row.cells[1].text.strip()
        is_aisatsu = ttl.startswith("ご挨拶")
        frm = 0 if is_aisatsu else max(floor, body_from)
        no, idx = page_of(norm(ttl), frm)
        found[n] = no
        if idx is not None and not is_aisatsu:
            floor = max(idx, body_from)
    for n, row in enumerate(rows):
        if not row.cells[0].text.strip():
            continue
        kids = []
        for m in range(n + 1, len(rows)):
            if rows[m].cells[0].text.strip():
                break
            if found[m] is not None:
                kids.append(found[m])
        found[n] = (min(kids) if kids
                    else page_of(norm(row.cells[1].text), body_from)[0])
    # 目次自身の頁を拾っていないかを確かめる
    toc_pages = {no for no, txt, i in pg if i < body_from and no is not None}
    on_toc = [(rows[i].cells[1].text.strip(), found[i])
              for i in range(len(found))
              if found[i] is not None and found[i] in toc_pages
              and not rows[i].cells[1].text.strip().startswith("ご挨拶")]

    # 頁番号が戻っていないかを確かめる
    seq = [v for v in found if v is not None]
    back = [(rows[i].cells[1].text.strip(), found[i])
            for i in range(1, len(found))
            if found[i] is not None and found[i - 1] is not None
            and found[i] < found[i - 1]]

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
    if on_toc:
        print("  ⚠ 目次自身の頁を拾っている行があります：")
        for t, v in on_toc:
            print(f"      {t}　{v}")
    else:
        print("  目次自身の頁を拾っている行はありません。")
    if back:
        print("  ⚠ 頁番号が前の行より小さい行があります：")
        for t, v in back:
            print(f"     {t}　{v}")
    else:
        print("  頁番号は単調増加です。")


if __name__ == "__main__":
    main()
