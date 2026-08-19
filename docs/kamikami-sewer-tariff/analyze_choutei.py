#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""階上町 下水道使用料 改定検討用 集計スクリプト.

使い方:
    python3 analyze_choutei.py <集計ブック.xlsx> [調定明細.xlsx ...]

  第1引数 : 【R7】使用料集計ブック（「R7漁集 (件数)」「R7公共 (件数) 」シートを持つもの）
  第2引数以降（任意）: 調定簿明細のブック。現行料金体系との整合を検証する。

調定Excelは個人情報を含むためリポジトリには含めていない。
"""
import math
import sys

import openpyxl

# --- 現行使用料体系（税抜・1か月あたり） -----------------------------------
# 基本使用料 5㎥まで 1,008円 / 6〜10㎥ 37円 / 11〜50㎥ 174円 / 51㎥〜 200円
# 調定は隔月検針のため、2か月分の水量 V に対して区分境界を2倍(10/20/100㎥)して適用する。
CURRENT = dict(base=2016, r1=37, r2=174, r3=200)

# 新料金案1（税込・1か月あたり）を2か月分の水量 V に直接適用できる形にしたもの
# 基本 1,220円 / 6〜10㎥ 45.5円 / 11〜50㎥ 211.5円 / 51㎥〜 242円
PLAN1 = dict(base=2440, r1=45.5, r2=211.5, r3=242)


def charge(volume, base, r1, r2, r3):
    """2か月分の使用水量から料金を計算する（税抜ベースの積み上げ）。"""
    amount = base
    if volume > 10:
        amount += r1 * min(volume - 10, 10)
    if volume > 20:
        amount += r2 * min(volume - 20, 80)
    if volume > 100:
        amount += r3 * (volume - 100)
    return amount


def tax_included(amount):
    """消費税10%を加算し、1円未満を切り捨てる。"""
    return math.floor(amount * 1.1)


def current_amount(volume):
    return tax_included(charge(volume, **CURRENT))


def plan1_amount(volume):
    # PLAN1 は税込単価のため、消費税の再乗算は行わない
    return math.floor(charge(volume, **PLAN1))


def volume_index():
    """現行料金額 → 使用水量 の逆引き表。"""
    index = {}
    for volume in range(1, 4000):
        index.setdefault(current_amount(volume), volume)
    return index


def load_counts(worksheet, index):
    """（件数）シートを (水量, 件数, 現行調定額) のリストに変換する。"""
    records = []
    for row in worksheet.iter_rows(min_row=3, values_only=True):
        if str(row[0]).strip() == '合計':
            continue
        amount, count = row[1], row[8]
        if not isinstance(count, (int, float)) or count == 0:
            continue
        if not isinstance(amount, (int, float)):
            # 「～2,216」区分（日割・休止等）。金額は集計ブックの計算値をそのまま使う
            records.append((None, int(count), int(row[9] or 0)))
            continue
        amount = int(amount)
        if amount in index:
            volume = index[amount]
        elif amount <= 2217:
            volume = 10
        else:
            volume = index[max(k for k in index if k <= amount)]
        records.append((volume, int(count), amount * int(count)))
    return records


def revenue(records, amount_of, ratio):
    """料金体系を当てはめた年間調定額。水量不明分は ratio 倍で近似する。"""
    total = 0
    for volume, count, current in records:
        if volume is None:
            total += round(current * ratio)
        else:
            total += amount_of(volume) * count
    return total


def uniform_plan(ratio):
    """一律 ratio 倍の料金案（税込単価）。"""
    params = dict(
        base=math.floor(2217 * ratio),
        r1=round(40.7 * ratio, 1),
        r2=round(191.4 * ratio, 1),
        r3=round(220.0 * ratio, 1),
    )
    return lambda volume: math.floor(charge(volume, **params))


def report_summary(path):
    workbook = openpyxl.load_workbook(path, data_only=True)
    index = volume_index()
    grand = {}
    for sheet, label in [('R7漁集 (件数)', '漁業集落排水'), ('R7公共 (件数) ', '公共下水道')]:
        records = load_counts(workbook[sheet], index)
        count = sum(c for _, c, _ in records)
        current = sum(a for _, _, a in records)
        volume = sum((v or 10) * c for v, c, _ in records)
        print('=== %s' % label)
        print('  調定件数        : %s 件（隔月6回）' % format(count, ','))
        print('  調定額（税込）  : %s 円' % format(current, ','))
        print('  推計年間水量    : %s ㎥（料金からの逆算）' % format(volume, ','))
        print('  実効単価        : %.1f 円/㎥（税込）' % (current / volume))
        print('  1件平均         : %.0f 円 / %.1f ㎥（2か月）' % (current / count, volume / count))
        plan = revenue(records, plan1_amount, 1.10)
        print('  案1（約+10%%）   : %s 円 (%+.2f%%, 増収 %s 円)'
              % (format(plan, ','), (plan / current - 1) * 100, format(plan - current, ',')))
        for ratio in (1.15, 1.20):
            total = revenue(records, uniform_plan(ratio), ratio)
            print('  参考 一律%+.0f%%   : %s 円 (%+.2f%%, 増収 %s 円)'
                  % ((ratio - 1) * 100, format(total, ','), (total / current - 1) * 100,
                     format(total - current, ',')))
        grand[label] = (current, plan)

    current_total = sum(v[0] for v in grand.values())
    plan_total = sum(v[1] for v in grand.values())
    print('=== 2事業合計')
    print('  現行 %s 円 → 案1 %s 円（増収 %s 円 / %+.2f%%）'
          % (format(current_total, ','), format(plan_total, ','),
             format(plan_total - current_total, ','), (plan_total / current_total - 1) * 100))


def verify_tariff(paths):
    """調定明細の実額が現行料金体系で再現できるか検証する。"""
    index = volume_index()
    for path in paths:
        workbook = openpyxl.load_workbook(path, data_only=True)
        matched = under_base = mismatched = 0
        for worksheet in workbook.worksheets:
            for row in worksheet.iter_rows(values_only=True):
                if not (row[1] and str(row[1]).startswith('階上') and row[2]):
                    continue
                if not isinstance(row[6], (int, float)):
                    continue
                amount = int(row[6])
                if amount in index:
                    matched += 1
                elif amount < 2217:
                    under_base += 1
                else:
                    mismatched += 1
        total = matched + under_base + mismatched
        print('=== 料金体系の検証: %s' % path)
        print('  対象 %s 件 / 体系と一致 %s 件 (%.1f%%) / 基本料金未満（日割等） %s 件 / 不一致 %s 件'
              % (format(total, ','), format(matched, ','), matched / total * 100,
                 format(under_base, ','), format(mismatched, ',')))


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    report_summary(sys.argv[1])
    if len(sys.argv) > 2:
        verify_tariff(sys.argv[2:])
    return 0


if __name__ == '__main__':
    sys.exit(main())
