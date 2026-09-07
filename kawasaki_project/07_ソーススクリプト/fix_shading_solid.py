# -*- coding: utf-8 -*-
"""
docx の網掛けが黒ベタになる不具合を直す（<w:shd w:val="solid"> → "clear"）

OOXML の <w:shd> は、w:val が塗りつぶしのパターン、w:fill が背景色、
w:color が前景（パターン）の色である。w:val="solid" は「前景色でべた塗り」を意味し、
w:color を省略すると既定の auto＝黒になるため、w:fill にどの色を指定しても
**黒ベタで描画される**。淡色の背景を敷きたい場合は w:val="clear" が正しい。

本スクリプトは、指定した docx の <w:shd> について
  w:val="solid" → w:val="clear"、w:color="auto" を補う
という最小限の修正を行う。書体・配色・レイアウトは一切変えない。

使い方：
    python3 07_ソーススクリプト/fix_shading_solid.py <docx> [<docx> ...]

計画素案は、これに加えて体裁そのものを統一したため restyle_soan_v113.py で処理している。
"""
import shutil
import sys
import zipfile

import docx
from docx.oxml.ns import qn


def fix(path):
    doc = docx.Document(path)
    n = 0
    for part in (doc.element.body,):
        for shd in part.iter(qn("w:shd")):
            if shd.get(qn("w:val")) == "solid":
                shd.set(qn("w:val"), "clear")
                if shd.get(qn("w:color")) is None:
                    shd.set(qn("w:color"), "auto")
                n += 1
    # ヘッダー・フッターも対象にする
    for sec in doc.sections:
        for hf in (sec.header, sec.footer, sec.first_page_header,
                   sec.first_page_footer, sec.even_page_header, sec.even_page_footer):
            if hf is None:
                continue
            for shd in hf._element.iter(qn("w:shd")):
                if shd.get(qn("w:val")) == "solid":
                    shd.set(qn("w:val"), "clear")
                    if shd.get(qn("w:color")) is None:
                        shd.set(qn("w:color"), "auto")
                    n += 1
    if n:
        shutil.copy2(path, path + ".bak")
        doc.save(path)
        zipfile.ZipFile(path).testzip()
    return n


def main(paths):
    for p in paths:
        n = fix(p)
        print(f"{n:5d}箇所を修正　{p}" if n else f"　　  修正なし　{p}")


if __name__ == "__main__":
    main(sys.argv[1:])
