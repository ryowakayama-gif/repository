# -*- coding: utf-8 -*-
"""照会票（束ごと）の中身を JSON に落とす。docx 側（build_irai_bundle_docx.js）がこれを読む。

   使い方：
     python3 scripts/build_irai_bundle.py            中身を書き出す
     node    scripts/build_irai_bundle_docx.js 1     束1 の docx を組む
"""
import io
import json
import os
import sys
sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths as P
import irai_bundle as IB


def build():
    out = {"meta": IB.META, "goannai": IB.GOANNAI, "bundles": []}
    for nm, ti, nerai in IB.bundles():
        b = {"no": nm, "title": ti, "nerai": nerai, "n": IB.count(nm), "parts": []}
        for lv, title, lead, rows in IB.items(nm):
            b["parts"].append({"lv": lv, "title": title, "lead": lead,
                               "rows": [list(r) for r in rows]})
        out["bundles"].append(b)
    return out


def main():
    out = build()
    f = P.build("irai_bundle.json")
    io.open(f, "w", encoding="utf-8").write(json.dumps(out, ensure_ascii=False, indent=1))
    print("JSON:", f)
    for b in out["bundles"]:
        print("  %s %s　%d件（%d部）" % (b["no"], b["title"], b["n"], len(b["parts"])))
    print("  計 %d件" % sum(b["n"] for b in out["bundles"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
