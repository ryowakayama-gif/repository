# -*- coding: utf-8 -*-
"""資料編 付表2（属性別クロス集計）を個票から全数再集計して照合する。

あわせて、冒頭[P90]が宣言したマスク運用
  n<3        → 各欄「−」
  3<=n<10    → 割合を表示せず件数のみ
が表の側で守られているかを検査する。
"""
import re
import sys
from load_jittai import load, layers
from parse_tables import parse
from decimal import Decimal, ROUND_HALF_UP


def pct(cnt, n):
    """四捨五入して小数第1位まで"""
    if not n:
        return None
    return float(Decimal(str(cnt * 100 / n)).quantize(Decimal("0.1"),
                                                      rounding=ROUND_HALF_UP))

DOC = sys.argv[1] if len(sys.argv) > 1 else "N_shiryo.txt"

blocks, recs = load()
LAY = layers(recs)
BYKEY = {b["key"]: b for b in blocks}

# 付表2-n → 個票の設問キー（キャプションの設問名とサービス名から対応づける）
SUBMAP = {"訪問介護": "訪問介護", "訪問入浴介護": "訪問入浴", "訪問看護": "訪問看護",
          "訪問リハビリテーション": "訪問リハ", "通所介護": "通所介護",
          "通所リハビリテーション": "通所リハ", "夜間対応型訪問介護": "夜間対応型訪問",
          "定期巡回・随時対応型訪問介護看護": "定期巡回", "小規模多機能型居宅介護": "小多機",
          "看護小規模多機能型居宅介護": "看多機",
          "短期入所生活介護": "ショートステイ", "居宅療養管理指導": "居宅療養管理"}


def keyof(cap):
    m = re.search(r"([AB]問\d)", cap)
    if not m:
        return None
    q = m.group(1)
    if q != "A問8":
        return q
    for k in sorted(SUBMAP, key=len, reverse=True):
        if k in cap:
            return "A問8_" + SUBMAP[k]
    return None


NUM = re.compile(r"^(\d+)(?:⏎?\((\d+\.\d)%\))?$")
tables, paras = parse(DOC)
res = dict(ok=0, ng=[], cells=0, mask_ng=[], nrow=0)

for num, cap, rows in tables:
    if "付表2-" not in cap:
        continue
    key = keyof(cap)
    if key is None:
        res["ng"].append((cap, "設問を特定できない"))
        continue
    blk = BYKEY[key]
    head = rows[0]
    opts = head[3:]
    # 表側の選択肢ラベル → 個票の選択肢
    real = [o for c, o in blk["opts"]]
    idx = []
    for o in opts:
        o2 = o.replace("⏎", "").strip()
        cands = [i for i, rr in enumerate(real)
                 if rr.replace(" ", "") == o2.replace(" ", "")]
        idx.append(cands[0] if cands else None)
    if any(i is None for i in idx):
        res["ng"].append((cap, f"列見出しが個票の選択肢と一致しない: {opts}"))
        continue
    axis = None
    for row in rows[1:]:
        a, k = row[0], row[1]
        if a:
            axis = a
        if k == "－" or a == "全　体" or a == "全体":
            pool = recs
            label = "全体"
        else:
            if axis not in LAY:
                res["ng"].append((cap, f"属性軸不明: {axis}"))
                break
            hit = [g for nm, g in LAY[axis] if nm == k]
            if not hit:
                res["ng"].append((cap, f"区分不明: {axis}/{k}"))
                continue
            pool = hit[0]
            label = f"{axis}/{k}"
        ans = [r["ans"][key] for r in pool]
        valid = [s for s in ans if s]
        n_true = len(valid)
        res["nrow"] += 1
        try:
            n_rep = int(row[2])
        except ValueError:
            res["ng"].append((cap, f"{label} n欄が数値でない: {row[2]}"))
            continue
        if n_rep != n_true:
            res["ng"].append((cap, f"{label} n={n_rep} ≠ 再集計{n_true}"))
        # 各セル
        for j, o in enumerate(opts):
            cell = row[3 + j]
            cnt = sum(1 for s in valid if real[idx[j]] in s)
            res["cells"] += 1
            if cell == "−" or cell == "－" or cell == "-":
                if n_true >= 3:
                    res["mask_ng"].append(
                        (cap, f"{label} n={n_true}（3以上）なのに「−」: {o}"))
                continue
            m = NUM.match(cell.replace(" ", ""))
            if not m:
                res["ng"].append((cap, f"{label}／{o} セルが読めない: {cell!r}"))
                continue
            c_rep = int(m.group(1))
            p_rep = m.group(2)
            if c_rep != cnt:
                res["ng"].append(
                    (cap, f"{label}／{o} 件数{c_rep} ≠ 再集計{cnt}"))
            else:
                res["ok"] += 1
            if p_rep is None:
                if n_true >= 10:
                    res["mask_ng"].append(
                        (cap, f"{label} n={n_true}（10以上）なのに割合なし: {o}"))
            else:
                if n_true < 10:
                    res["mask_ng"].append(
                        (cap, f"{label} n={n_true}（10未満）なのに割合表示 {p_rep}%: {o}"))
                exp = pct(cnt, n_true)
                if exp is None or abs(float(p_rep) - exp) > 0.051:
                    res["ng"].append(
                        (cap, f"{label}／{o} 割合{p_rep}% ≠ 再計算"
                              f"{exp}%（{cnt}/{n_true}）"))

print(f"照合セル {res['cells']}／層行 {res['nrow']}／件数一致 {res['ok']}")
print(f"数値の不一致 {len(res['ng'])}件")
for cap, msg in res["ng"][:60]:
    print("  ×", cap[:34], "|", msg)
print(f"マスク運用の不整合 {len(res['mask_ng'])}件")
for cap, msg in res["mask_ng"][:60]:
    print("  △", cap[:34], "|", msg)
