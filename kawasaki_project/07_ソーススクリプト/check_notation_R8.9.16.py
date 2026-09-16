# -*- coding: utf-8 -*-
"""表記ゆれ・誤字の機械検査。"""
import re
import sys
from collections import Counter
from parse_tables import parse

PAIRS = [
    ("波ダッシュ", r"～", r"〜"),
    ("％の全半角", r"％", r"%"),
    ("括弧 n=（全半角）", r"（n[=＝]", r"\(n[=＝]"),
]
NUMS = [("要支援の数字", r"要支援[0-9]", r"要支援[０-９一二三四五１２３４５]"),
        ("要介護の数字", r"要介護[0-9]", r"要介護[０-９]")]

for path in sys.argv[1:]:
    tables, paras = parse(path)
    body = "\n".join(t for _, _, t in paras)
    tbl = "\n".join(" ".join(r) for _, _, rows in tables for r in rows)
    whole = body + "\n" + tbl
    print(f"##### {path}")
    for name, a, b in PAIRS:
        print(f"  {name:16s} 本文 {len(re.findall(a, body)):4d}/{len(re.findall(b, body)):4d}"
              f"   表 {len(re.findall(a, tbl)):4d}/{len(re.findall(b, tbl)):4d}")
    for name, a, b in NUMS:
        print(f"  {name:16s} 半角 {len(re.findall(a, whole)):4d}  全角 {len(re.findall(b, whole)):4d}")
    # 数え方の単位
    for w in ["人", "件", "名"]:
        print(f"  「{w}」を数えた回数（本文）: "
              f"{len(re.findall(r'[0-9０-９]' + w, body))}")
    print()
