# -*- coding: utf-8 -*-
"""集計仕様書の自己点検。

   分母と指標の定義が、集計が始まってから崩れることを防ぐ。
   不適合があれば終了コード1を返す。
"""
import os as _os_p
import sys as _sys_p
_sys_p.path.insert(0, _os_p.path.dirname(_os_p.path.abspath(__file__)))
import paths as _P   # 置き場所はここで決める（じか書きしない）
import os, re, sys
sys.dont_write_bytecode = True   # 古いバイトコードで誤った結果が出ることを防ぐ
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from shukei_data import N, Z, DERIVED, CROSS, AXES, SHU_HO

MD = os.path.join(_P.BASE, "25_集計仕様書.md")
MD24 = os.path.join(_P.BASE, "24_アンケート調査報告書_骨子案.md")
R = []
def chk(no, name, ok, detail=""):
    R.append((no, name, ok, detail))

md = open(MD, encoding="utf-8").read()
md24 = open(MD24, encoding="utf-8").read()   # 報告書骨子案（doc24）

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

# ── ここから doc24（報告書骨子案）との照合 ────────────────
#    集計仕様書を正とし、報告書の書き方が食い違っていないことを確かめる。

# 9 分母の方針が報告書骨子にも書かれていること（方針9・11に対応）
need9 = [
    "構成比の分母は、その設問に回答した票の数（設問ごとの有効回答）です",
    "回収票の総数で割った値は用いません",
    "行・列の双方に回答した票の数です",
    "時点の異なる値を足して合計としません",
]
miss9 = [x[:16] for x in need9 if x not in md24]
chk(9, "分母の方針が報告書骨子にもあること", not miss9,
    f"欠落 {miss9}" if miss9 else "分母の定義・総数で割らない・両設問・時点を記載")

# 10 主指標・補足指標の5件が報告書骨子にも書かれていること（方針10）
miss10 = [x[0] for x in SHU_HO if x[0] not in md24]
chk(10, "主指標・補足指標が報告書骨子にもあること", not miss10,
    f"doc24にない {miss10}" if miss10 else f"{len(SHU_HO)}件すべて記載")

# 11 20表の分母が、集計仕様書の定義のまま報告書骨子に書かれていること
miss11 = [c[0] for c in CROSS if c[0] not in md24 or str(c[5]) not in md24]
chk(11, "20表の分母が報告書骨子と一致すること", not miss11,
    f"不一致 {miss11}" if miss11 else "X-01〜X-20の分母が doc24 §13 と一致")

# 12 在宅調査の要介護度が、報告書骨子でも2区分であること
bad12 = "要介護度（要支援1・2" in md24
chk(12, "報告書骨子の在宅の要介護度が2区分であること",
    not bad12 and "要介護1・2／要介護3〜5の2区分" in md24,
    "3区分の記載が残っている" if bad12 else "2区分に統一（名簿に要支援なし）")

# 13 問7の枝番が集計仕様書と一致していること（村独自3問は(9)(10)(11)）
old13 = [x for x in ("問7(1)〜(10)", "問7(8)(9)", "問7(8)(9)(10)") if x in md24]
q7 = [q for (_, q, e, *_r) in N if q == "問7"]
chk(13, "問7の枝番が集計仕様書と一致すること",
    not old13 and "問7(11)" in md24 and len(q7) == 11,
    f"旧番号が残っている {old13}" if old13 else "村独自3問は問7(9)(10)(11)・全11枝")

# 14 リスク判定の領域数が一致していること（R01〜R10＋R99）
rcode = [d[0] for d in DERIVED if d[0].startswith("R") and d[0] != "R99"]
miss14 = [c for c in rcode if c not in md24]
chk(14, "リスク判定が10領域＋R99で一致すること",
    len(rcode) == 10 and not miss14 and "R99" in md24 and "R01〜R11" not in md24,
    f"doc24にない {miss14}" if miss14 else "R01〜R10・R99が両文書で一致")

# 15 集計軸Iの区分が2区分に揃っていること（凡例・md・doc24）
axI = [a for a in AXES if a[0] == "I"][0]
chk(15, "集計軸Iの区分が2区分であること",
    "要支援" not in axI[2] and "要介護1・2／要介護3〜5（2区分）" in md,
    f"軸Iの区分 {axI[2]}")

# 16 中核の★表の数が、CROSS と両文書の記述で一致すること
star = [c[0] for c in CROSS if str(c[6]).startswith("★")]
hokyo = [c[0] for c in CROSS if not str(c[6]).startswith("★")]
say = f"★を付した{len(star)}表が分析編"
chk(16, "中核の★表の数が一致すること",
    say in md and f"**★を付した{len(star)}表が分析編（第Ⅱ部）の中核**" in md24
    and f"残る{len(hokyo)}表" in md,
    f"★{len(star)}表・補強{len(hokyo)}表")

w = max(len(n) for _, n, _, _ in R)
print("■ 集計仕様書の自己点検")
ng = 0
for no, name, ok, detail in R:
    print(f"  {no:>3}  {name:<{w}}  {'適合' if ok else '不適合'}" + (f"  {detail}" if detail else ""))
    ng += 0 if ok else 1
print(f"\n  {len(R)}件のうち適合{len(R)-ng}件／不適合{ng}件")
sys.exit(1 if ng else 0)
