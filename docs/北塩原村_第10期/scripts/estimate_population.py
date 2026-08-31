# -*- coding: utf-8 -*-
"""第10期 第1号被保険者数の推計（社人研推計＋実績乖離率補正）
   入力：P1（社人研 将来推計人口／各年10月1日）／B1・B2（介護保険事業状況報告／各年3月末）
   出力：令和9〜11年度末の前期・後期別第1号被保険者数と認定者数の参考試算"""

# ── P1 社人研推計（各年10月1日、単位：人）──
P1 = {2025:{"前期":457,"後期":480}, 2030:{"前期":360,"後期":551},
      2035:{"前期":302,"後期":541}, 2040:{"前期":246,"後期":519}}

# ── 実績（B2 前期・後期別第1号被保険者数、令和8年3月末）──
ACTUAL_2026 = {"前期":493, "後期":522}

# ── 年齢階級別認定率（令和8年3月末実績）──
#    前期＝(合計認定者214−75歳以上認定者182)/前期被保険者493
#    後期＝75歳以上認定者182/後期被保険者522
RATE = {"前期":32/493, "後期":182/522}


def interp(t, key):
    """P1を線形補間（tは西暦の小数、10月1日基準）"""
    ys = sorted(P1)
    for a, b in zip(ys, ys[1:]):
        if a <= t <= b:
            va, vb = P1[a][key], P1[b][key]
            return va + (vb - va) * (t - a) / (b - a)
    raise ValueError(f"範囲外: {t}")


def at_fiscal_year_end(march_year, key):
    """◯年3月末の値。10月1日基準に対し0.5年前倒しで換算する。"""
    return interp(march_year - 0.5, key)


# ── 乖離率の算定 ──
base = {k: at_fiscal_year_end(2026, k) for k in ("前期", "後期")}
ratio = {k: ACTUAL_2026[k] / base[k] for k in ("前期", "後期")}


def estimate(march_year):
    """◯年3月末（＝前年度末）の補正後 第1号被保険者数"""
    return {k: round(at_fiscal_year_end(march_year, k) * ratio[k]) for k in ("前期", "後期")}


if __name__ == "__main__":
    print("【乖離率】2026年3月末（令和7年度末）")
    for k in ("前期", "後期"):
        print(f"  {k}: P1補間 {base[k]:6.1f} / 実績 {ACTUAL_2026[k]:4d} → 乖離率 {ratio[k]:.4f}")

    print("\n【補正後の第1号被保険者数】")
    plan = [(2026, "令和7年度末（実績）"), (2028, "令和9年度末"),
            (2029, "令和10年度末"), (2030, "令和11年度末")]
    print(f"{'':<18}{'前期':>7}{'後期':>7}{'計':>7}{'後期割合':>9}")
    for y, label in plan:
        v = ACTUAL_2026 if y == 2026 else estimate(y)
        tot = v["前期"] + v["後期"]
        print(f"{label:<18}{v['前期']:7d}{v['後期']:7d}{tot:7d}{v['後期']/tot*100:8.1f}%")

    print(f"\n【認定率】前期 {RATE['前期']*100:.2f}% / 後期 {RATE['後期']*100:.2f}%"
          f"（{RATE['後期']/RATE['前期']:.1f}倍）")
    print("\n【認定者数の参考試算】認定率を令和7年度末水準で固定")
    print(f"{'':<18}{'前期':>7}{'後期':>7}{'計':>7}{'認定率':>8}")
    for y, label in plan:
        v = ACTUAL_2026 if y == 2026 else estimate(y)
        a, b = v["前期"] * RATE["前期"], v["後期"] * RATE["後期"]
        tot = v["前期"] + v["後期"]
        print(f"{label:<18}{a:7.1f}{b:7.1f}{a+b:7.0f}{(a+b)/tot*100:7.1f}%")
