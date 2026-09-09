#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""料金改定案の比較表を作成する。

  案A : 現行体系を維持し、全区分の単価を一律に引き上げる
  案B : 同等の増収を確保しつつ体系を見直す
        ・基本使用料は据置（少量利用者への配慮）
        ・6〜10㎥の単価を引き上げ、11㎥での4.7倍の段差を緩和
        ・51㎥〜（大口）は全体改定率と同水準にとどめる

改定率は5%・10%・15%の3水準。経費回収率は決算統計32表ベースで
経営戦略のR16目標（公共40%・漁集26%）を既に上回っているため、
20%改定は検討対象から外している。

使い方:
    python3 compare_plans.py <【R7】使用料集計ブック.xlsx> [出力ディレクトリ]
"""
import csv
import io
import math
import os
import sys

import openpyxl

# --- 令和6年度 決算統計（32表・40表） ---------------------------------------
# 汚水処理費は維持管理費（汚水分）のみ。資本費は全額が「分流式下水道等に要する
# 経費」として基準内繰入金で措置されているため汚水処理費に含まれない。
R6 = {
    '公共下水道': dict(
        使用料=33652, 汚水処理費=67464, 有収水量=209289,
        維持管理費計=67794, 非汚水分=330, 非汚水分名称='水洗便所等普及費',
        減価償却費=122209, 支払利息=26042, 長期前受金戻入=60445,
        分流式=87801, 資本費その他=5,
        繰入_収益_基準額=88136, 繰入_収益_実額=220527),
    '漁業集落排水': dict(
        使用料=7229, 汚水処理費=19614, 有収水量=45402,
        維持管理費計=21932, 非汚水分=2318, 非汚水分名称='不明水処理費',
        減価償却費=35740, 支払利息=1798, 長期前受金戻入=28584,
        分流式=8949, 資本費その他=5,
        繰入_収益_基準額=11272, 繰入_収益_実額=16202),
}

# --- 料金体系 ----------------------------------------------------------------
# base  : 1か月の基本使用料（5㎥まで・税込）
# tiers : [(1か月の区分上限㎥, 1㎥あたり税込単価), ...] 末尾は (None, 単価)
CURRENT = dict(base=1108.8, tiers=[(10, 40.7), (50, 191.4), (None, 220.0)])

UPLIFTS = (0.05, 0.10, 0.15)

PLAN_A = {  # 現行の各単価を一律に引き上げ（0.1円単位に丸め）
    0.05: dict(base=1164.2, tiers=[(10, 42.7), (50, 201.0), (None, 231.0)]),
    0.10: dict(base=1219.7, tiers=[(10, 44.8), (50, 210.5), (None, 242.0)]),
    0.15: dict(base=1275.1, tiers=[(10, 46.8), (50, 220.1), (None, 253.0)]),
}
PLAN_B = {  # 基本使用料は据置。6〜10㎥を引き上げ、11㎥の段差を緩和。大口は全体率並み
    0.05: dict(base=1108.8, tiers=[(10, 60), (50, 200.7), (None, 231.0)]),
    0.10: dict(base=1108.8, tiers=[(10, 80), (50, 209.5), (None, 242.0)]),
    0.15: dict(base=1108.8, tiers=[(10, 90), (50, 225.1), (None, 253.0)]),
}
# 集計ブックが作成済みの案（現行体系を一律引上げ。厳密には+10.30%）
PLAN_1 = dict(base=1220, tiers=[(10, 45.5), (50, 211.5), (None, 242)])

# 月使用量帯（世帯像の目安つき）。境界は「超〜以下」
BANDS = [(0, 5, '月5㎥以下', '単身・高齢世帯など'),
         (5, 10, '月6〜10㎥', '2人世帯など'),
         (10, 20, '月11〜20㎥', '3〜4人世帯など'),
         (20, 30, '月21〜30㎥', '5人以上世帯など'),
         (30, 50, '月31〜50㎥', '小規模事業所など'),
         (50, 10 ** 9, '月51㎥〜', '大口（事業所・公共施設）')]


def charge2m(volume, base, tiers):
    """2か月分の水量に対する請求額（税込・1円未満切捨て）。

    隔月検針のため、月額表の区分境界を2倍して適用する（現行の賦課方法と同じ）。
    """
    amount = base * 2
    prev = 10                                   # 基本使用料の範囲（5㎥×2）
    for upper, rate in tiers:
        hi = volume if upper is None else min(volume, upper * 2)
        if hi > prev:
            amount += rate * (hi - prev)
        if upper is None:
            break
        prev = upper * 2
        if volume <= upper * 2:
            break
    return math.floor(amount)


def monthly(volume, base, tiers):
    """1か月分の使用料（税込・端数処理前）。モデルケース表示用。"""
    amount = base
    prev = 5
    for upper, rate in tiers:
        hi = volume if upper is None else min(volume, upper)
        if hi > prev:
            amount += rate * (hi - prev)
        if upper is None:
            break
        prev = upper
        if volume <= upper:
            break
    return amount


def rate_of(volume, base, tiers):
    """その使用量における単価（円/㎥）。負担の公平性をみるための指標。"""
    return monthly(volume, base, tiers) / volume


# --- R7調定データ（通常算定／基本料金帯／特殊算定） --------------------------
EXACT_VOLUME = {charge2m(v, **CURRENT): v for v in range(11, 4000)}
BASIC_AMOUNT = charge2m(10, **CURRENT)
NORMAL, BASIC, SPECIAL = '通常算定', '基本料金帯', '特殊算定'


def load_counts(workbook, sheet):
    records = []
    for row in workbook[sheet].iter_rows(min_row=3, values_only=True):
        if str(row[0]).strip() == '合計':
            continue
        amount, count = row[1], row[8]
        if not isinstance(count, (int, float)) or count == 0:
            continue
        count = int(count)
        if not isinstance(amount, (int, float)):
            records.append((SPECIAL, None, count, int(row[9] or 0)))
            continue
        amount = int(amount)
        if amount == BASIC_AMOUNT:
            records.append((BASIC, None, count, amount * count))
        elif amount in EXACT_VOLUME:
            records.append((NORMAL, EXACT_VOLUME[amount], count, amount * count))
        else:
            records.append((SPECIAL, None, count, amount * count))
    return records


def revenue(records, plan, ratio):
    """特殊算定（日割・異動等）は本来の水量が不明なため現行調定額×改定率で扱う。"""
    total = 0
    for kind, volume, count, current in records:
        if kind == NORMAL:
            total += charge2m(volume, **plan) * count
        elif kind == BASIC:
            total += charge2m(10, **plan) * count
        else:
            total += round(current * ratio)
    return total


def recovery(business, uplift):
    """改定後の経費回収率と繰入金。

    汚水処理費はR6水準で据置。基準内繰入金は、繰出基準の使用料対象資本費
    （＝使用料収入−汚水維持管理費、負なら0）がゼロのままのため変化しない。
    水洗便所等普及費（漁集は不明水処理費）は現行額が継続する前提。
    """
    d = R6[business]
    fee = round(d['使用料'] * (1 + uplift))
    capital_covered = max(0, fee - d['汚水処理費'])          # 使用料対象資本費
    base_in = d['繰入_収益_基準額'] - capital_covered
    return dict(使用料=fee, 経費回収率=fee / d['汚水処理費'] * 100,
                使用料対象資本費=capital_covered, 基準内繰入金=base_in,
                増収額=fee - d['使用料'],
                基準外削減余地=d['繰入_収益_実額'] - d['繰入_収益_基準額'] - (fee - d['使用料']),
                汚水処理原価=d['汚水処理費'] / d['有収水量'] * 1000,
                使用料単価=fee / d['有収水量'] * 1000)


def write(path, rows):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with io.open(path, 'w', encoding='utf-8-sig', newline='') as fh:
        csv.writer(fh).writerows(rows)
    print('  ->', os.path.basename(path), '(%d行)' % len(rows))


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    workbook = openpyxl.load_workbook(sys.argv[1], data_only=True)
    out = sys.argv[2] if len(sys.argv) > 2 else '.'
    gyo = load_counts(workbook, 'R7漁集 (件数)')
    kou = load_counts(workbook, 'R7公共 (件数) ')
    cur_g = sum(a for _, _, _, a in gyo)
    cur_k = sum(a for _, _, _, a in kou)

    # === 決算統計の確認 =====================================================
    rows = [['#', '令和6年度決算統計 32表・40表。資本費は全額が「分流式下水道等に要する経費」'
                  'として基準内繰入金で措置され、汚水処理費は維持管理費（汚水分）のみとなる'],
            ['項目', '公共下水道', '漁業集落排水', '出典・算式']]
    for key, label, src in [
            ('維持管理費計', '維持管理費 合計', '32表(43)'),
            ('汚水処理費', 'うち汚水処理費', '32表(44)'),
            ('非汚水分', 'うち非汚水分', '32表(47)水洗便所等普及費／(48)不明水処理費'),
            ('減価償却費', '減価償却費', '32表(58)'),
            ('支払利息', '支払利息', '32表(51)'),
            ('長期前受金戻入', '長期前受金戻入分', '32表(2-5)'),
            ('分流式', '分流式下水道等に要する経費', '32表(2-12)'),
            ('資本費その他', '資本費のその他', '32表(2-13)'),
            ('繰入_収益_基準額', '40表 基準内繰入金（収益的）基準額', '40表(12)'),
            ('繰入_収益_実額', '40表 繰入金（収益的）実繰入額', '40表(13)'),
            ('使用料', '下水道使用料', '32表(2-24)'),
            ('有収水量', '年間有収水量(㎥)', '10表(52)')]:
        rows.append([label, R6['公共下水道'][key], R6['漁業集落排水'][key], src])
    for label, fn, src in [
            ('［検証］減価償却費−長期前受金戻入＋支払利息',
             lambda d: d['減価償却費'] - d['長期前受金戻入'] + d['支払利息'], '資本費'),
            ('［検証］分流式＋その他', lambda d: d['分流式'] + d['資本費その他'],
             '上と一致すれば資本費全額が基準内繰入金で措置されている'),
            ('［検証］非汚水分＋分流式＋その他',
             lambda d: d['非汚水分'] + d['分流式'] + d['資本費その他'],
             '40表 基準内繰入金 基準額と一致する'),
            ('経費回収率(%)', lambda d: round(d['使用料'] / d['汚水処理費'] * 100, 1),
             '使用料 ÷ 汚水処理費'),
            ('汚水処理原価(円/㎥)', lambda d: round(d['汚水処理費'] / d['有収水量'] * 1000, 1),
             '汚水処理費 ÷ 有収水量'),
            ('使用料単価(円/㎥)', lambda d: round(d['使用料'] / d['有収水量'] * 1000, 1),
             '使用料 ÷ 有収水量')]:
        rows.append([label, fn(R6['公共下水道']), fn(R6['漁業集落排水']), src])
    write(os.path.join(out, '14_R6決算統計_32表40表.csv'), rows)

    # === 料金表の比較 =======================================================
    rows = [['#', '1か月あたり・税込。基本使用料は5㎥まで。'
                  '案Aは全区分を一律に引き上げ、案Bは基本使用料を据え置いて6〜10㎥を引き上げる'],
            ['料金案', '基本使用料(5㎥まで)', '6〜10㎥', '11〜50㎥', '51㎥〜',
             '11㎥での単価の跳ね上がり', '基本使用料の増減率(%)']]

    def tariff_row(label, plan):
        rate_6, rate_11, rate_51 = (plan['tiers'][0][1], plan['tiers'][1][1], plan['tiers'][2][1])
        return [label, plan['base'], rate_6, rate_11, rate_51,
                '%.2f倍' % (rate_11 / rate_6),
                '%+.1f' % (plan['base'] / CURRENT['base'] * 100 - 100)]

    rows.append(tariff_row('現行', CURRENT))
    for uplift in UPLIFTS:
        rows.append(tariff_row('案A %+.0f%%（一律改定）' % (uplift * 100), PLAN_A[uplift]))
    for uplift in UPLIFTS:
        rows.append(tariff_row('案B %+.0f%%（体系見直し）' % (uplift * 100), PLAN_B[uplift]))
    rows.append(tariff_row('（参考）集計ブック案1', PLAN_1))
    write(os.path.join(out, '15_料金表の比較.csv'), rows)

    # === 増収と経費回収率の比較 =============================================
    rows = [['#', '増収額はR7奇数月6回調定に各案を当てはめた再計算（税込）。'
                  '経費回収率・繰入金はR6決算統計に増収率を乗じたもので、汚水処理費はR6水準で据置'],
            ['#', '水洗便所等普及費330千円（漁集は不明水処理費2,318千円）は現行額が継続する前提。'
                  '繰出基準の使用料対象資本費がゼロのままのため基準内繰入金は全案で不変'],
            ['料金案', '漁集 調定額(円)', '公共 調定額(円)', '合計(円)', '増収額(円)', '増収率(%)',
             '公共 経費回収率(%)', '漁集 経費回収率(%)',
             '公共 基準内繰入金(千円)', '漁集 基準内繰入金(千円)',
             '基準外繰入金の削減余地(千円)']]
    base_total = cur_g + cur_k
    plans = [('現行', CURRENT, 0.0)]
    for uplift in UPLIFTS:
        plans.append(('案A %+.0f%%（一律改定）' % (uplift * 100), PLAN_A[uplift], uplift))
    for uplift in UPLIFTS:
        plans.append(('案B %+.0f%%（体系見直し）' % (uplift * 100), PLAN_B[uplift], uplift))
    plans.append(('（参考）集計ブック案1', PLAN_1, 0.10))
    for label, plan, uplift in plans:
        g = revenue(gyo, plan, 1 + uplift)
        k = revenue(kou, plan, 1 + uplift)
        rk = recovery('公共下水道', k / cur_k - 1)
        rg = recovery('漁業集落排水', g / cur_g - 1)
        rows.append([label, g, k, g + k, g + k - base_total,
                     '%.2f' % ((g + k) / base_total * 100 - 100),
                     '%.1f' % rk['経費回収率'], '%.1f' % rg['経費回収率'],
                     rk['基準内繰入金'], rg['基準内繰入金'],
                     rk['増収額'] + rg['増収額']])
    write(os.path.join(out, '16_増収と経費回収率の比較.csv'), rows)

    # === 世帯への影響度 =====================================================
    rows = [['#', '一般汚水。1か月あたりの使用料（税込）。調定は隔月のため実際の請求額はこの2倍'],
            ['#', '案Aは全区分を一律に引き上げるため、どの使用量帯もほぼ同じ率で上がる。'
                  '案Bは基本使用料を据え置くため月5㎥以下は据置となり、その分を6〜50㎥が負担する'],
            ['月使用量(㎥)', '世帯像の目安', '現行(円)']]
    for uplift in UPLIFTS:
        rows[-1] += ['案A%+.0f%% 月額' % (uplift * 100), '案A%+.0f%% 増減額' % (uplift * 100),
                     '案A%+.0f%% 増減率(%%)' % (uplift * 100)]
    for uplift in UPLIFTS:
        rows[-1] += ['案B%+.0f%% 月額' % (uplift * 100), '案B%+.0f%% 増減額' % (uplift * 100),
                     '案B%+.0f%% 増減率(%%)' % (uplift * 100)]
    examples = [(3, '単身（節水）'), (5, '単身・高齢世帯'), (8, '2人世帯'), (10, '2人世帯'),
                (13, '3人世帯'), (15, '3〜4人世帯'), (20, '4人世帯'), (25, '5人世帯'),
                (30, '5人以上世帯'), (40, '小規模事業所'), (50, '小規模事業所'),
                (100, '大口事業所'), (200, '大口事業所')]
    for volume, who in examples:
        current = monthly(volume, **CURRENT)
        row = [volume, who, round(current, 1)]
        for plans_by_rate in (PLAN_A, PLAN_B):
            for uplift in UPLIFTS:
                after = monthly(volume, **plans_by_rate[uplift])
                row += [round(after, 1), '%+.1f' % (after - current),
                        '%+.2f' % ((after / current - 1) * 100)]
        rows.append(row)
    write(os.path.join(out, '17_世帯への影響度.csv'), rows)

    # === 使用量帯別の構成と影響 =============================================
    def band_of(kind, volume):
        v = 5 if kind == BASIC else volume / 2      # 2か月水量→月使用量
        for low, high, name, who in BANDS:
            if v <= high:
                return name
        return BANDS[-1][2]

    rows = [['#', '月使用量帯ごとの件数・調定額の構成と、各案での負担増加率（帯の代表値で算定）'],
            ['#', '大口（月51㎥〜）は件数がごく僅かで、増収への寄与も限られる'],
            ['事業', '月使用量帯', '世帯像の目安', '件数', '件数割合(%)', '調定額(円)', '金額割合(%)',
             '代表値(㎥)', '案A+10% 増減率(%)', '案B+10% 増減率(%)']]
    rep = {'月5㎥以下': 5, '月6〜10㎥': 8, '月11〜20㎥': 15, '月21〜30㎥': 25,
           '月31〜50㎥': 40, '月51㎥〜': 100}
    for records, label, total in [(kou, '公共下水道', cur_k), (gyo, '漁業集落排水', cur_g)]:
        agg = {name: [0, 0] for _, _, name, _ in BANDS}
        special = [0, 0]
        for kind, volume, count, amount in records:
            if kind == SPECIAL:
                special[0] += count
                special[1] += amount
                continue
            name = band_of(kind, volume)
            agg[name][0] += count
            agg[name][1] += amount
        count_all = sum(v[0] for v in agg.values()) + special[0]
        for _, _, name, who in BANDS:
            count, amount = agg[name]
            v = rep[name]
            cur_m = monthly(v, **CURRENT)
            rows.append([label, name, who, count, '%.1f' % (count / count_all * 100),
                         amount, '%.1f' % (amount / total * 100), v,
                         '%+.2f' % ((monthly(v, **PLAN_A[0.10]) / cur_m - 1) * 100),
                         '%+.2f' % ((monthly(v, **PLAN_B[0.10]) / cur_m - 1) * 100)])
        rows.append([label, '特殊算定（日割・異動等）', '—', special[0],
                     '%.1f' % (special[0] / count_all * 100), special[1],
                     '%.1f' % (special[1] / total * 100), '—', '—', '—'])
    write(os.path.join(out, '18_使用量帯別の構成と影響.csv'), rows)

    # === 設計レバーの感度 ===================================================
    # 料金体系の設計は「基本使用料をどれだけ据え置くか」と「大口をどこまで抑えるか」
    # の2つのレバーで決まる。どちらも中間層（月11〜50㎥）の負担に跳ね返る。
    coef = {'件数': 0, '6〜10㎥帯': 0, '11〜50㎥帯': 0, '51㎥〜帯': 0, '特殊': 0}
    for records in (gyo, kou):
        for kind, volume, count, amount in records:
            if kind == SPECIAL:
                coef['特殊'] += amount
                continue
            coef['件数'] += count
            if kind == BASIC:
                continue
            coef['6〜10㎥帯'] += max(0, min(volume, 20) - 10) * count
            coef['11〜50㎥帯'] += max(0, min(volume, 100) - 20) * count
            coef['51㎥〜帯'] += max(0, volume - 100) * count
    total_current = cur_g + cur_k

    def solve_middle(uplift, base, rate_6, rate_51):
        """基本使用料・6〜10㎥・51㎥〜を決めたとき、目標増収に必要な11〜50㎥の単価。"""
        need = (total_current * uplift - coef['特殊'] * uplift
                - (base - CURRENT['base']) * 2 * coef['件数']
                - (rate_6 - 40.7) * coef['6〜10㎥帯']
                - (rate_51 - 220.0) * coef['51㎥〜帯'])
        return round(191.4 + need / coef['11〜50㎥帯'], 1)

    rows = [['#', '全体+10%を確保する前提で、設計レバーを動かしたときの負担分布'],
            ['#', '基本使用料を据え置くほど少量利用者は軽くなるが、その分を中間層が負担する。'
                  '大口を抑えるほど中間層の負担が増えるが、大口は件数が僅少のため影響は限定的'],
            ['設計', '基本使用料', '6〜10㎥', '11〜50㎥（逆算）', '51㎥〜',
             '月5㎥', '月10㎥', '月15㎥', '月20㎥', '月30㎥', '月50㎥', '月100㎥']]
    levers = [('案A：全区分を一律+10%', 1219.7, 44.8, 242.0),
              ('案B：基本据置・大口は全体率並み', 1108.8, 80, 242.0),
              ('（変種）基本据置・大口も据置', 1108.8, 80, 220.0),
              ('（変種）基本を全体率の半分だけ引上げ', 1164.2, 80, 242.0),
              ('（変種）基本据置・6〜10㎥は控えめ', 1108.8, 60, 242.0)]
    for label, base, rate_6, rate_51 in levers:
        rate_mid = solve_middle(0.10, base, rate_6, rate_51)
        plan = dict(base=base, tiers=[(10, rate_6), (50, rate_mid), (None, rate_51)])
        row = [label, base, rate_6, rate_mid, rate_51]
        for v in (5, 10, 15, 20, 30, 50, 100):
            row.append('%+.1f' % ((monthly(v, **plan) / monthly(v, **CURRENT) - 1) * 100))
        rows.append(row)
    write(os.path.join(out, '19_設計レバーの感度.csv'), rows)
    return 0


if __name__ == '__main__':
    sys.exit(main())
