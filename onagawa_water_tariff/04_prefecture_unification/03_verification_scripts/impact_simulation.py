"""
水産加工業への統一化影響 定量試算スクリプト
=====================================================

本スクリプトは、4団体の確定料金体系を用いて
以下の政策シナリオ別に水産加工業への影響を試算するものです。

【シナリオ】
1. 現行体系(基準)
2. 単純統一(4団体の加重平均的水準)
3. 塩竈モデル統一(生産用水用区分の全県導入)
4. 10年経過措置(初年度)
5. 15年経過措置(初年度)

【水産加工業者の規模分布モデル】
- 小規模: 25mm/200㎥/月 (家族経営加工場) - 構成比 5/19
- 中規模: 50mm/1000㎥/月 (一般加工会社)  - 構成比 8/19
- 大規模: 75mm/3000㎥/月 (冷凍・加工主要事業者) - 構成比 5/19
- 超大口: 100mm/8000㎥/月 (大手水産加工) - 構成比 1/19

【想定社数】
- 女川町: 19社(水産加工団地の構成)
- 気仙沼市: 120社
- 石巻広域: 180社
- 塩竈市: 90社
"""

import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__)))
from verify_all_rates import onagawa, kesennuma, ishinomaki, shiogama, shiogama_production


# ============================================================
# 単純統一の仮置き料金(4団体の加重平均的水準)
# ============================================================
def unified_avg(q, c):
    """統一料金の仮置き(4団体の加重平均的水準)"""
    bases = {13: 1050, 20: 1500, 25: 2300, 40: 4800,
             50: 9800, 75: 19000, 100: 37000, 150: 75000}
    basic = bases.get(c, 0)
    if c in (13, 20):
        excess = max(0, q - 10)
    else:
        excess = q
    fee = 0
    r = excess
    for w, u in [(10, 130), (30, 185), (50, 220), (400, 265), (500, 285), (float('inf'), 305)]:
        take = min(r, w)
        fee += take * u
        r -= take
        if r <= 0:
            break
    return int((basic + fee) * 1.10)


# ============================================================
# 塩竈モデル統一(生産用水用区分を全県導入した場合)
# ============================================================
def shiogama_model(q, c):
    """塩竈モデル: 産業用は105円/㎥(基本料金なし)、家庭用は標準体系"""
    if c in (13, 20):
        # 家庭用は統一標準料金
        return unified_avg(q, c)
    # 大口産業用は生産用水用適用
    return int(q * 105 * 1.10)


# ============================================================
# 経過措置(初年度の料金)
# ============================================================
def transitional(current_func, target_func, q, c, years, current_year=1):
    """経過措置: 現行料金 → 統一料金へ years 年で段階的移行"""
    cur = current_func(q, c)
    tgt = target_func(q, c)
    return int(cur + (tgt - cur) * (current_year / years))


# ============================================================
# 水産加工セグメント
# ============================================================
SEGMENTS = [
    ("小規模", 200, 25, 5),   # (名称, 使用量, 口径, 19社中の構成比)
    ("中規模", 1000, 50, 8),
    ("大規模", 3000, 75, 5),
    ("超大口", 8000, 100, 1),
]

TEAMS = [
    ("女川町", onagawa, 19),      # (名称, 現行料金関数, 社数)
    ("気仙沼市", kesennuma, 120),
    ("石巻広域", ishinomaki, 180),
    ("塩竈市", shiogama, 90),
]


# ============================================================
# 個別事業者への影響(単純統一)
# ============================================================
def scenario_impact_by_segment():
    """規模別の影響率を計算"""
    print("=" * 90)
    print("個別事業者への影響率(単純統一時)")
    print("=" * 90)
    print(f"{'区分':<12}{'使用条件':<18}{'女川町':<12}{'気仙沼市':<12}{'石巻広域':<12}{'塩竈市':<12}")
    print("-" * 90)

    scenarios = [
        ("一般家庭", 20, 20),
        ("小規模加工", 200, 25),
        ("中規模加工", 1000, 50),
        ("大規模加工", 3000, 75),
        ("超大口加工", 8000, 100),
    ]

    for name, q, c in scenarios:
        unified = unified_avg(q, c)
        on_change = (unified - onagawa(q, c)) / onagawa(q, c) * 100
        ke_change = (unified - kesennuma(q, c)) / kesennuma(q, c) * 100
        isn_change = (unified - ishinomaki(q, c)) / ishinomaki(q, c) * 100
        sh_change = (unified - shiogama(q, c)) / shiogama(q, c) * 100
        print(f"{name:<12}{c}mm/{q}㎥        {on_change:>+7.1f}%    {ke_change:>+7.1f}%    {isn_change:>+7.1f}%    {sh_change:>+7.1f}%")


# ============================================================
# 県全体への年間影響
# ============================================================
def total_annual_impact(target_func, label="統一"):
    """4団体の年間影響を集計"""
    def monthly_per_19(func):
        return sum(func(q, c) * cnt for _, q, c, cnt in SEGMENTS)

    print(f"\n=" * 90)
    print(f"県全体への年間影響({label}適用時)")
    print("=" * 90)
    print(f"{'団体':<10}{'社数':<8}{'現行年額':<15}{'{}後年額'.format(label):<15}{'差額':<20}{'変化率':<10}")
    print("-" * 90)

    total_cur = 0
    total_new = 0
    for name, func, cnt in TEAMS:
        cur_yr = monthly_per_19(func) * (cnt / 19) * 12
        new_yr = monthly_per_19(target_func) * (cnt / 19) * 12
        diff = new_yr - cur_yr
        total_cur += cur_yr
        total_new += new_yr
        sign = "+" if diff >= 0 else ""
        pct = diff / cur_yr * 100 if cur_yr > 0 else 0
        print(f"{name:<10}{cnt}社    {cur_yr / 1e8:>8.2f}億円    {new_yr / 1e8:>8.2f}億円    {sign}{diff / 1e8:>+7.2f}億円  ({sign}{pct:+.1f}%)")

    total_diff = total_new - total_cur
    total_pct = total_diff / total_cur * 100
    sign = "+" if total_diff >= 0 else ""
    print("-" * 90)
    print(f"{'4団体計':<10}{'409社':<8}{total_cur / 1e8:>8.2f}億円    {total_new / 1e8:>8.2f}億円    {sign}{total_diff / 1e8:>+7.2f}億円  ({sign}{total_pct:+.1f}%)")


# ============================================================
# 政策シナリオ別の比較(女川町中規模を例に)
# ============================================================
def policy_scenario_comparison():
    """女川町中規模加工業者への影響を政策シナリオ別に比較"""
    print("\n" + "=" * 90)
    print("女川町 中規模水産加工業者(50mm/1,000㎥/月)政策シナリオ別影響")
    print("=" * 90)

    q, c = 1000, 50
    current = onagawa(q, c)
    unified = unified_avg(q, c)
    shiogama_m = shiogama_production(q)
    trans_10 = current + (unified - current) / 10  # 10年経過措置初年度
    trans_15 = current + (unified - current) / 15  # 15年経過措置初年度

    scenarios = [
        ("現行(基準)", current),
        ("単純統一", unified),
        ("塩竈モデル(生産用水用)", shiogama_m),
        ("10年経過措置(初年度)", trans_10),
        ("15年経過措置(初年度)", trans_15),
    ]

    for label, val in scenarios:
        pct = (val - current) / current * 100
        sign = "+" if pct >= 0 else ""
        marker = " ← 推奨" if "塩竈モデル" in label else ""
        print(f"  {label:<30}: {val:>10,.0f}円  ({sign}{pct:.1f}%){marker}")


if __name__ == "__main__":
    scenario_impact_by_segment()
    total_annual_impact(unified_avg, "単純統一")
    total_annual_impact(shiogama_model, "塩竈モデル")
    policy_scenario_comparison()
