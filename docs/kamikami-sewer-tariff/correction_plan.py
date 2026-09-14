#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""6〜10㎥区分の是正案のエビデンスを作成する。

説明資料「下水道使用料の改定について ― 6〜10㎥区分の是正を含む料金体系の見直し案 ―」
に載せた数値は、すべてこのスクリプトで再現できる。

使い方:
    python3 correction_plan.py <【R7】使用料集計ブック.xlsx> [出力ディレクトリ]
"""
import csv
import io
import math
import os
import sys

import openpyxl

# --- 現行使用料体系（1か月・税込） -------------------------------------------
CURRENT = dict(base=1108.8, tiers=[(10, 40.7), (50, 191.4), (None, 220.0)])
MID_RATE = 191.4          # 11〜50㎥の単価。是正の到達点
LOW_RATE = 40.7           # 6〜10㎥の現行単価


def charge2m(volume, base, tiers):
    """2か月分の水量に対する請求額（税込・1円未満切捨て）。"""
    amount = base * 2
    prev = 10
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


def coefficients(records_list):
    """料金は各単価について線形。区分ごとの係数（2か月水量の合計）を集計する。"""
    c = dict(件数=0, 帯6_10=0, 帯11_50=0, 帯51=0, 特殊=0)
    for records in records_list:
        for kind, volume, count, amount in records:
            if kind == SPECIAL:
                c['特殊'] += amount
                continue
            c['件数'] += count
            if kind == BASIC:
                continue
            c['帯6_10'] += max(0, min(volume, 20) - 10) * count
            c['帯11_50'] += max(0, min(volume, 100) - 20) * count
            c['帯51'] += max(0, volume - 100) * count
    return c


def revenue(records, plan, ratio):
    total = 0
    for kind, volume, count, current in records:
        if kind == NORMAL:
            total += charge2m(volume, **plan) * count
        elif kind == BASIC:
            total += charge2m(10, **plan) * count
        else:
            total += round(current * ratio)
    return total


def write(path, rows):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with io.open(path, 'w', encoding='utf-8-sig', newline='') as fh:
        csv.writer(fh).writerows(rows)
    print('  ->', os.path.basename(path), '(%d行)' % len(rows))


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
    out = sys.argv[2] if len(sys.argv) > 2 else '.'
    gyo = load_counts(wb, 'R7漁集 (件数)')
    kou = load_counts(wb, 'R7公共 (件数) ')
    cur_g = sum(a for _, _, _, a in gyo)
    cur_k = sum(a for _, _, _, a in kou)
    total_current = cur_g + cur_k
    co = coefficients([gyo, kou])

    def solve_mid(uplift, base, rate_6, rate_51):
        """基本使用料・6〜10㎥・51㎥〜を決めたとき、目標増収に必要な11〜50㎥の単価。"""
        need = (total_current * uplift - co['特殊'] * uplift
                - (base - CURRENT['base']) * 2 * co['件数']
                - (rate_6 - LOW_RATE) * co['帯6_10']
                - (rate_51 - 220.0) * co['帯51'])
        return round(MID_RATE + need / co['帯11_50'], 1)

    def plan(base, rate_6, rate_mid, rate_51):
        return dict(base=base, tiers=[(10, rate_6), (50, rate_mid), (None, rate_51)])

    def correction_ratio(rate_6):
        return (rate_6 - LOW_RATE) / (MID_RATE - LOW_RATE) * 100

    # === 20 是正余地の算定 ==================================================
    gap = (MID_RATE - LOW_RATE) * co['帯6_10']
    write(os.path.join(out, '20_是正余地の算定.csv'), [
        ['#', '6〜10㎥を11〜50㎥と同じ単価に揃えた場合の増収余地（2事業計・税込）'],
        ['#', '料金は各単価について線形のため、単価差×該当水量で増収額が求まる'],
        ['項目', '数値', '算式・備考'],
        ['6〜10㎥の従量水量（2事業計）', co['帯6_10'], '2か月分の合計。調定の金額から逆算'],
        ['11〜50㎥の従量水量', co['帯11_50'], '同上'],
        ['51㎥〜の従量水量', co['帯51'], '同上'],
        ['調定件数（水量を推計できる分）', co['件数'], '基本料金部分が乗じられる件数'],
        ['特殊算定の調定額', co['特殊'], '日割・異動等。改定率を乗じて扱う'],
        [],
        ['6〜10㎥の現行単価', LOW_RATE, '円/㎥（税込）'],
        ['11〜50㎥の現行単価', MID_RATE, '円/㎥（税込）'],
        ['単価差', round(MID_RATE - LOW_RATE, 1), '191.4 − 40.7'],
        ['増収額', round(gap), '%s㎥ × %s円' % (format(co['帯6_10'], ','), round(MID_RATE - LOW_RATE, 1))],
        ['現行調定額', total_current, '漁集 %s ＋ 公共 %s' % (format(cur_g, ','), format(cur_k, ','))],
        ['現行比', '%.1f%%' % (gap / total_current * 100), '増収額 ÷ 現行調定額'],
        ['段差（11〜50㎥ ÷ 6〜10㎥）', '%.2f倍' % (MID_RATE / LOW_RATE), '191.4 ÷ 40.7'],
    ])

    # === 21 是正水準の比較 ==================================================
    LEVELS = [('是正なし', LOW_RATE), ('弱い是正', 60), ('中程度の是正（推奨）', 80),
              ('強い是正', 110), ('（参考）完全是正', MID_RATE)]
    rows = [['#', '全体+10%を確保する前提で、6〜10㎥をどこまで引き上げるかを比較'],
            ['#', '基本使用料は1,108.8円で据置、51㎥〜は242.0円（+10%）で固定。'
                  '11〜50㎥は目標増収から逆算'],
            ['是正水準', '6〜10㎥(円/㎥)', '是正率(%)', '11〜50㎥(円/㎥)', '51㎥〜(円/㎥)',
             '段差(倍)', '月5㎥', '月8㎥', '月10㎥', '月15㎥', '月20㎥', '月30㎥', '月50㎥', '月100㎥']]
    for label, rate_6 in LEVELS:
        rate_mid = solve_mid(0.10, CURRENT['base'], rate_6, 242.0)
        p = plan(CURRENT['base'], rate_6, rate_mid, 242.0)
        row = [label, rate_6, '%.1f' % correction_ratio(rate_6), rate_mid, 242.0,
               '%.2f' % (rate_mid / rate_6)]
        for v in (5, 8, 10, 15, 20, 30, 50, 100):
            row.append('%+.1f' % ((monthly(v, **p) / monthly(v, **CURRENT) - 1) * 100))
        rows.append(row)
    write(os.path.join(out, '21_是正水準の比較.csv'), rows)

    # === 22 推奨案の料金表と増収 ============================================
    RECOMMENDED = {0.05: 60, 0.10: 80, 0.15: 90}
    R6 = {'公共下水道': dict(fee=33652, cost=67464), '漁業集落排水': dict(fee=7229, cost=19614)}
    rows = [['#', '推奨案＝基本使用料は据置、6〜10㎥を引き上げ、51㎥〜は全体改定率と同水準'],
            ['#', '増収額はR7年度奇数月6回調定に当てはめた再計算（税込）。'
                  '経費回収率はR6決算統計32表に増収率を乗じたもので、汚水処理費はR6水準で据置'],
            ['料金案', '基本使用料', '6〜10㎥', '11〜50㎥', '51㎥〜', '是正率(%)',
             '漁集 調定額(円)', '公共 調定額(円)', '合計(円)', '増収額(円)', '増収率(%)',
             '公共 経費回収率(%)', '漁集 経費回収率(%)']]
    rows.append(['現行', CURRENT['base'], LOW_RATE, MID_RATE, 220.0, '0.0',
                 cur_g, cur_k, total_current, 0, '0.00',
                 '%.1f' % (R6['公共下水道']['fee'] / R6['公共下水道']['cost'] * 100),
                 '%.1f' % (R6['漁業集落排水']['fee'] / R6['漁業集落排水']['cost'] * 100)])
    for uplift, rate_6 in sorted(RECOMMENDED.items()):
        rate_51 = round(220.0 * (1 + uplift), 1)
        rate_mid = solve_mid(uplift, CURRENT['base'], rate_6, rate_51)
        p = plan(CURRENT['base'], rate_6, rate_mid, rate_51)
        g, k = revenue(gyo, p, 1 + uplift), revenue(kou, p, 1 + uplift)
        rk = R6['公共下水道']['fee'] * (k / cur_k) / R6['公共下水道']['cost'] * 100
        rg = R6['漁業集落排水']['fee'] * (g / cur_g) / R6['漁業集落排水']['cost'] * 100
        rows.append(['推奨案 %+.0f%%' % (uplift * 100), '%s（据置）' % CURRENT['base'],
                     rate_6, rate_mid, rate_51, '%.1f' % correction_ratio(rate_6),
                     g, k, g + k, g + k - total_current,
                     '%.2f' % ((g + k) / total_current * 100 - 100), '%.1f' % rk, '%.1f' % rg])
    write(os.path.join(out, '22_推奨案の料金表と増収.csv'), rows)

    # === 23 世帯への影響 ====================================================
    rec = plan(CURRENT['base'], 80, solve_mid(0.10, CURRENT['base'], 80, 242.0), 242.0)
    stage2 = plan(CURRENT['base'], 120, solve_mid(0.20, CURRENT['base'], 120, 264.0), 264.0)
    rows = [['#', '一般汚水・税込。調定は隔月のため、請求額は月額の2倍'],
            ['月使用量(㎥)', '世帯像の目安', '現行 月額', '現行 2か月請求',
             '推奨案+10% 月額', '推奨案+10% 2か月請求', '差額(2か月)', '増減率(%)',
             '第2段階 月額', '第2段階 2か月請求', '現行比 増減率(%)']]
    for v, who in [(3, '単身（節水）'), (5, '単身・高齢世帯'), (8, '2人世帯'), (10, '2人世帯'),
                   (13, '3人世帯'), (15, '3〜4人世帯'), (20, '4人世帯'), (25, '5人世帯'),
                   (30, '5人以上世帯'), (40, '小規模事業所'), (50, '小規模事業所'),
                   (100, '大口事業所'), (200, '大口事業所')]:
        c_m, c_2 = monthly(v, **CURRENT), charge2m(2 * v, **CURRENT)
        r_m, r_2 = monthly(v, **rec), charge2m(2 * v, **rec)
        s_m, s_2 = monthly(v, **stage2), charge2m(2 * v, **stage2)
        rows.append([v, who, round(c_m, 1), c_2, round(r_m, 1), r_2, r_2 - c_2,
                     '%+.1f' % ((r_m / c_m - 1) * 100), round(s_m, 1), s_2,
                     '%+.1f' % ((s_m / c_m - 1) * 100)])
    write(os.path.join(out, '23_世帯への影響.csv'), rows)

    # === 24 段階的是正 ======================================================
    rows = [['#', '経営戦略が織り込む令和8年度+10%・令和13年度累計+20%に合わせた段階的是正'],
            ['段階', '基本使用料', '6〜10㎥', '是正率(%)', '11〜50㎥', '51㎥〜',
             '段差(倍)', '現行比増収率(%)', '増収額(円)']]
    rows.append(['現行', CURRENT['base'], LOW_RATE, '0.0', MID_RATE, 220.0,
                 '%.2f' % (MID_RATE / LOW_RATE), '0.00', 0])
    for label, uplift, rate_6 in [('第1段階（令和8年度）', 0.10, 80),
                                  ('第2段階（令和13年度）', 0.20, 120)]:
        rate_51 = round(220.0 * (1 + uplift), 1)
        rate_mid = solve_mid(uplift, CURRENT['base'], rate_6, rate_51)
        p = plan(CURRENT['base'], rate_6, rate_mid, rate_51)
        g, k = revenue(gyo, p, 1 + uplift), revenue(kou, p, 1 + uplift)
        rows.append([label, '%s（据置）' % CURRENT['base'], rate_6,
                     '%.1f' % correction_ratio(rate_6), rate_mid, rate_51,
                     '%.2f' % (rate_mid / rate_6),
                     '%.2f' % ((g + k) / total_current * 100 - 100), g + k - total_current])
    write(os.path.join(out, '24_段階的是正.csv'), rows)
    return 0


if __name__ == '__main__':
    sys.exit(main())
