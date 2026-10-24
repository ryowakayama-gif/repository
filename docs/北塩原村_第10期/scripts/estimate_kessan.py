# -*- coding: utf-8 -*-
"""介護保険特別会計の決算（令和5・6年度）と介護給付費準備基金の残高

【出どころ】
  村より受領した令和5年度・令和6年度の決算書（doc118）から、介護保険に関わる
  数値だけを書き写した data/原典_決算/ の3表。原本は版管理に入れていない。
  令和7年度の決算は受領待ちであり、受領したら同じ表に列を足す
  （00_受領待ち_令和7年度決算.md）。

【数を書き写さない】
  素案（4-6・5-5(2)・5-6・5-7・5-12）はこのモジュールの値を引く。

  python3 scripts/estimate_kessan.py
"""
import csv
import os
import sys

sys.dont_write_bytecode = True
import paths as P

D = P.data("原典_決算")
NENDO = ("令和5年度", "令和6年度")


def tokkai():
    """{(区分, 款, 項): {年度: 円}}"""
    out = {}
    with open(os.path.join(D, "介護保険特別会計_R5R6.csv"), encoding="utf-8") as f:
        for r in csv.DictReader(f):
            out[(r["区分"], r["款"], r["項"])] = {y: int(r[y]) for y in NENDO if r.get(y)}
    return out


def sai(kou, kan=None):
    """歳出の項（または款の計）の額 {年度: 円}"""
    t = tokkai()
    for (ku, ka, ko), v in t.items():
        if ku == "歳出" and ko == kou and (kan is None or ka == kan):
            return v
    raise KeyError(kou)


def hojo():
    """補助給付等（特定入所者・高額・高額医療合算・その他諸費）の計 {年度: 円}"""
    ks = ("3 特定入所者介護サービス等費", "4 高額介護サービス等費",
          "5 高額医療合算介護サービス等費", "6 その他諸費")
    return {y: sum(sai(k)[y] for k in ks) for y in NENDO}


def shohi():
    """サービス諸費（介護サービス等諸費＋介護予防サービス等諸費） {年度: 円}"""
    return {y: sai("1 介護サービス等諸費")[y] + sai("2 介護予防サービス等諸費")[y]
            for y in NENDO}


def kyufu():
    """保険給付費の計 {年度: 円}"""
    return sai("（計）", "2 保険給付費")


def chiiki():
    """地域支援事業費の計 {年度: 円}"""
    return sai("（計）", "3 地域支援事業費")


def kikin():
    """介護給付費準備基金の年度末残高 {時点: 円}"""
    with open(os.path.join(D, "介護給付費準備基金_残高.csv"), encoding="utf-8") as f:
        return {r["時点"]: int(r["残高（円）"]) for r in csv.DictReader(f)}


def rojin():
    """老人福祉費・社会福祉施設費の事業別支出 [(事業, {年度: 円}, 節, 備考)]"""
    with open(os.path.join(D, "老人福祉費_村単独事業_R5R6.csv"), encoding="utf-8") as f:
        return [(r["事業"], {y: int(r[f"{y} 支出済額（円）"]) for y in NENDO}, r["節"], r["備考"])
                for r in csv.DictReader(f)]


def main():
    for nm, f in (("サービス諸費", shohi), ("補助給付等", hojo), ("保険給付費", kyufu),
                  ("地域支援事業費", chiiki)):
        v = f()
        print(f"  {nm:<10}" + "".join(f"{v[y]:>14,}" for y in NENDO))
    print("  準備基金", kikin())


if __name__ == "__main__":
    main()
