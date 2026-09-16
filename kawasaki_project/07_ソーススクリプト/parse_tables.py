# -*- coding: utf-8 -*-
"""dump_docx_text.py の出力から表を取り出す。"""
import re

RE_T = re.compile(r"^--- TABLE (\d+) \((\d+)行 x (\d+)列\) ---$")
RE_E = re.compile(r"^--- /TABLE (\d+) ---$")
RE_P = re.compile(r"^\[P(\d+)\](\[図[^\]]*\])?<([^>]*)> (.*)$")


def parse(path):
    """[(表番号, キャプション, [[セル,...],...])] とテキスト段落を返す"""
    lines = open(path, encoding="utf-8").read().splitlines()
    tables, paras = [], []
    last_cap = ""
    i = 0
    while i < len(lines):
        m = RE_T.match(lines[i])
        if m:
            num = int(m.group(1))
            rows = []
            i += 1
            while i < len(lines) and not RE_E.match(lines[i]):
                rows.append([c.strip() for c in lines[i].split(" | ")])
                i += 1
            tables.append((num, last_cap, rows))
            i += 1
            continue
        mp = RE_P.match(lines[i])
        if mp:
            paras.append((int(mp.group(1)), mp.group(3), mp.group(4)))
            last_cap = mp.group(4)
        i += 1
    return tables, paras
