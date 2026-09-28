# -*- coding: utf-8 -*-
"""集計仕様書の自己点検。

   分母と指標の定義が、集計が始まってから崩れることを防ぐ。
   不適合があれば終了コード1を返す。
"""
import os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from shukei_data import N, Z, DERIVED, CROSS, AXES, SHU_HO

MD = "/home/user/repository/docs/北塩原村_第10期/25_集計仕様書.md"
R = []
def chk(no, name, ok, detail=""):
    R.append((no, name, ok, detail))

md = open(MD, encoding="utf-8").read()

# 1 クロス表20件すべてに分母が書かれていること
nob = [c[0] for c in CROSS if not str(c[5]).strip() or str(c[5]).strip() == "―"]
chk(1, "クロス表に分母が書かれていること", len(CROSS) == 20 and not nob,
    f"分母が空 {nob}" if nob else f"20表すべてに分母を記載")

# 2 分母に「回収票の総数」を使っていないこと
bad2 = [c[0] for c in CROSS if "回収票" in str(c[5]) and "総数" in str(c[5])]
chk(2, "分母に回収票の総数を用いていないこと", not bad2,
    f"総数で割っている {bad2}" if bad2 else "20表とも設問または両設問の有効回答")

# 3 分母の書き方が「設問」「両設問」「名簿由来」のいずれかに揃っていること
KEY = ("両設問", "設問", "有効回答", "名簿由来")
bad3 = [c[0] for c in CROSS if not any(k in str(c[5]) for k in KEY)]
chk(3, "分母の書き方が揃っていること", not bad3,
    f"書き方が不揃い {bad3}" if bad3 else "20表とも定型の語で記載")

# 4 主指標・補足指標の整理があり、主に分母が書かれていること
nom = [x[0] for x in SHU_HO if "分母" not in x[1] and "データ" not in x[1] and "判定" not in x[1]]
chk(4, "主指標に出所と分母が書かれていること", len(SHU_HO) >= 5 and not nom,
    f"分母がない {nom}" if nom else f"{len(SHU_HO)}件すべてに出所を記載")

# 5 基本方針が11項目で、9〜11が md にあること
need5 = ["分母は設問ごとの有効回答に統一", "同じ指標を複数の設問から作らない", "表ごとに分母とn数を書く"]
miss5 = [x for x in need5 if x not in md]
chk(5, "基本方針9〜11が集計仕様書にあること", "集計の基本方針（11項目）" in md and not miss5,
    f"欠落 {miss5}" if miss5 else "11項目・方針9〜11を記載")

# 6 主指標・補足指標の表が md にもあること（xlsx と md の両方に置く）
miss6 = [x[0] for x in SHU_HO if x[0] not in md]
chk(6, "主指標・補足指標が集計仕様書にあること", not miss6,
    f"mdにない {miss6}" if miss6 else f"{len(SHU_HO)}件すべて記載")

# 7 在宅調査の要介護度が2区分に揃っていること（名簿に要支援がいないため）
bad7 = [c[0] for c in CROSS if c[4] == "在宅" and "I（3区分）" in (c[2], c[3])]
chk(7, "在宅調査の要介護度が2区分であること", not bad7,
    f"3区分のまま {bad7}" if bad7 else "在宅のクロスは2区分に統一")

# 8 n<30 の断りが残っていること
chk(8, "n<30を参考値とする断りがあること", "n<30" in md and "参考値" in md)

w = max(len(n) for _, n, _, _ in R)
print("■ 集計仕様書の自己点検")
ng = 0
for no, name, ok, detail in R:
    print(f"  {no:>3}  {name:<{w}}  {'適合' if ok else '不適合'}" + (f"  {detail}" if detail else ""))
    ng += 0 if ok else 1
print(f"\n  {len(R)}件のうち適合{len(R)-ng}件／不適合{ng}件")
sys.exit(1 if ng else 0)
