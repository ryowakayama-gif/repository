# -*- coding: utf-8 -*-
"""カラー版と単色版の組を確かめる（verify_soan 点検84・verify_shiryo 点検26 が使う）。

仕様書5①は計画書をモノクロ（コピー）で作成するとしており、村のご指示
（令和8年10月23日）により、計画書の素案と委員会資料をカラー版と単色版の2つで
管理する。片方だけを組み直すと中身が食い違うため、次の3つを見る。
  ① 両版の本文の文字が同じこと
  ② 単色版の文字・塗りに色がないこと
  ③ 単色版の図に色がないこと
"""
import io
import os
import re
import zipfile

import paths as P


def _honbun(z):
    x = z.read("word/document.xml").decode("utf-8")
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", x))


def _iro(h):
    r, g, b = int(h[:2], 16), int(h[2:4], 16), int(h[4:], 16)
    return max(r, g, b) - min(r, g, b) > 8


def check(names):
    """names は output/ のファイル名（拡張子なし）。食い違いの一覧を返す"""
    bad = []
    for nm in names:
        fc = os.path.join(P.OUT, nm + ".docx")
        fm = os.path.join(P.OUT, nm + "_単色.docx")
        nai = [x for x, f in (("カラー版", fc), ("単色版", fm)) if not os.path.exists(f)]
        if nai:
            bad.append(f"{nm}の{'・'.join(nai)}がない")
            continue
        zc, zm = zipfile.ZipFile(fc), zipfile.ZipFile(fm)
        if _honbun(zc) != _honbun(zm):
            bad.append(f"{nm}のカラー版と単色版で本文が違う（片方だけ組み直した）")
        iro = set()
        for n in zm.namelist():
            if n.endswith(".xml"):
                iro |= {h for h in re.findall(r'(?:w:val|w:fill|w:color)="([0-9A-Fa-f]{6})"',
                                              zm.read(n).decode("utf-8", "ignore")) if _iro(h)}
        if iro:
            bad.append(f"{nm}_単色に色がある（{'・'.join(sorted(iro)[:3])}）")
        try:
            from PIL import Image
            import numpy as np
        except ImportError:
            continue
        for n in zm.namelist():
            if n.startswith("word/media/") and n.lower().endswith(".png"):
                a = np.asarray(Image.open(io.BytesIO(zm.read(n))).convert("RGB")).astype(int)
                if ((a.max(2) - a.min(2)) > 20).mean() > 0.001:
                    bad.append(f"{nm}_単色の図に色がある（{n}）")
                    break
    return bad
