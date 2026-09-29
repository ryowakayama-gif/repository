# -*- coding: utf-8 -*-
"""図表データ管理台帳（xlsx）の上で直された数値を、正本 data_zuhyo.py へ取り込む.

台帳は利用者がエクセルの上で数値を直す場所、data_zuhyo.py は数値の正本である。
台帳を作り直すと正本の値で上書きされるため、**直した値は先にここで取り込む。**

  python3 scripts/sync_zuhyo.py          差分を表示するだけ（書き換えない）
  python3 scripts/sync_zuhyo.py --write  data_zuhyo.py を書き換える

行や列の数が変わっている図は取り込まない（形が合わないため）。
"""
import io
import os
import re
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import data_zuhyo as DZ

SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data_zuhyo.py")


def diffs():
    base = {d["name"]: d for d in DZ.load(use_book=False).values()}
    got = DZ.read_book(DZ.BOOK)
    out = []
    for k, ser in got.items():
        if k not in base:
            continue
        b = base[k]["series"]
        if len(ser) != len(b):
            continue
        for i, (_n, v) in enumerate(ser):
            if len(v) != len(b[i][1]):
                continue
            for j, (nv, ov) in enumerate(zip(v, b[i][1])):
                if nv is None or ov is None:
                    continue
                if abs(float(nv) - float(ov)) > 1e-9:
                    out.append((k, b[i][0], base[k]["cats"][j], ov, nv, i, j))
    return out


def write(ds):
    """ZU の該当する値だけを書き換える。並びや体裁はそのまま残す。"""
    import ast
    import pprint
    zu = [dict(d) for d in DZ.ZU]
    idx = {d["name"]: d for d in zu}
    for k, _sn, _cat, _ov, nv, i, j in ds:
        ser = [list(x) for x in idx[k]["series"]]
        vals = list(ser[i][1])
        vals[j] = nv
        ser[i][1] = vals
        idx[k]["series"] = ser
    s = io.open(SRC, encoding="utf-8").read()
    head = s[: s.index("ZU = [")]
    tail = s[s.index("\n\n\ndef by_name():"):]
    body = pprint.pformat(zu, width=96, sort_dicts=False)
    io.open(SRC, "w", encoding="utf-8").write(head + "ZU = " + body + tail)


if __name__ == "__main__":
    ds = diffs()
    if not ds:
        print("台帳と正本の数値は一致しています。取り込むものはありません。")
        sys.exit(0)
    print("台帳の上で直された数値 %d件" % len(ds))
    for k, sn, cat, ov, nv, _i, _j in ds:
        print("  %-28s %-16s %-12s %s → %s" % (k, sn, cat, ov, nv))
    if "--write" in sys.argv:
        write(ds)
        print("\ndata_zuhyo.py に取り込みました。"
              "続けて build_figures.py → build_soan.py → build_soan_docx.js "
              "→ build_zuhyo_daicho.py を実行してください。")
    else:
        print("\n書き換えるには  python3 scripts/sync_zuhyo.py --write")
