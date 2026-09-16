# -*- coding: utf-8 -*-
"""表キャプションが表の直前に置かれているかを検査する。"""
import re
import sys
RE_T = re.compile(r"^--- TABLE (\d+) ")
RE_P = re.compile(r"^\[P(\d+)\](\[図[^\]]*\])?<([^>]*)> (.*)$")
CAP = re.compile(r"^(表|付表)[0-9０-９]")
for path in sys.argv[1:]:
    lines = open(path, encoding="utf-8").read().splitlines()
    prev = None
    bad = []
    for i, l in enumerate(lines):
        m = RE_T.match(l)
        if m:
            # 直前の段落行
            j = i - 1
            while j >= 0 and lines[j].startswith("--- /TABLE"):
                j -= 1
            mp = RE_P.match(lines[j]) if j >= 0 else None
            txt = mp.group(4) if mp else ""
            if not CAP.match(txt):
                bad.append((m.group(1), txt[:60]))
    print(f"### {path}: キャプションが直前にない表 {len(bad)}件")
    for n, t in bad:
        print(f"   TABLE {n} ← 直前の段落「{t}」")
