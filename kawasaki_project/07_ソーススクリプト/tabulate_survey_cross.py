# -*- coding: utf-8 -*-
"""アンケート個票から属性別クロス集計の正解値を作る（レビュー用）。

使い方
  python3 tabulate_survey_cross.py <xlsx> <出力txt> [属性1,属性2,...]

既定の属性は 性別／年齢階級／要介護状態区分 と、性別×年齢階級の合成属性。
"""
import sys
from decimal import Decimal, ROUND_HALF_UP
from collections import OrderedDict

from tabulate_survey_truth import MARK, load



def pct1(k, n):
    """百分率を小数第1位で四捨五入する（Python の round は偶数丸めになるため使わない）"""
    if not n:
        return None
    return float(Decimal(k * 100) / Decimal(n)
                 == 0 and 0 or (Decimal(k * 100) / Decimal(n)).quantize(
                     Decimal("0.1"), rounding=ROUND_HALF_UP))


def levels_of(ws, rows, cols):
    """cols（1つ又は複数の属性列）の値の組をラベル化して行を束ねる"""
    d = OrderedDict()
    for r in rows:
        vs = []
        for c in cols:
            v = ws.cell(r, c).value
            vs.append(str(v).strip() if v is not None else "（無回答）")
        d.setdefault("×".join(vs), []).append(r)
    return d


def cross(ws, q, rs):
    if q["ncol"]:
        n = sum(1 for r in rs if ws.cell(r, q["ncol"]).value not in (None, "", 0))
    else:
        n = sum(1 for r in rs if any(ws.cell(r, c).value in MARK for c in q["cols"]))
    out = []
    for o, c in zip(q["opts"], q["cols"]):
        k = sum(1 for r in rs if ws.cell(r, c).value in MARK)
        out.append((o, k, pct1(k, n) if n else None))
    return n, out


def main():
    src, dst = sys.argv[1], sys.argv[2]
    ws, qs, rows, attrs = load(src)

    if len(sys.argv) > 3:
        specs = [s.split("×") for s in sys.argv[3].split(",")]
    else:
        specs = [["性別"], ["年齢階級"], ["要介護状態区分"], ["性別", "年齢階級"]]

    lines = [f"クロス集計の正解値（個票 {len(rows)}件）", f"元ファイル：{src}", ""]
    for spec in specs:
        cols = [attrs[a] for a in spec if a in attrs]
        if len(cols) != len(spec):
            lines.append(f"【{'×'.join(spec)}】属性列が見つかりません")
            continue
        lv = levels_of(ws, rows, cols)
        lines.append(f"================ 属性：{'×'.join(spec)}"
                     f"（層 {len(lv)}／各層の件数 "
                     + "、".join(f"{k}={len(v)}" for k, v in lv.items()) + "）")
        for q in qs:
            lines.append(f"■ {q['group']}　{q['text'][:60]}")
            for name, rs in lv.items():
                n, cells = cross(ws, q, rs)
                body = "　".join(f"{o}:{k}({p}%)" for o, k, p in cells)
                lines.append(f"　［{name}］n={n}　{body}")
            lines.append("")
        lines.append("")
    with open(dst, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"{dst}: {len(lines)}行")


if __name__ == "__main__":
    main()
