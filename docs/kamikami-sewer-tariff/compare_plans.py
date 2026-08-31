#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""料金改定案の比較表を作成する。

  案A : 現行体系を維持し、各単価を一律に引き上げる（経営戦略ケース①準拠）
  案B : 同等の増収を確保しつつ体系を見直す
        （6〜10㎥の単価を引き上げ、11㎥での4.7倍の段差を緩和し、区分を1つ増やす）

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

PLAN_A = {  # 現行体系を一律に引き上げ（+10%は集計ブックの案1と同一）
    0.10: dict(base=1220, tiers=[(10, 45.5), (50, 211.5), (None, 242)]),
    0.15: dict(base=1274.5, tiers=[(10, 46.8), (50, 220.1), (None, 253)]),
    0.20: dict(base=1330, tiers=[(10, 48.8), (50, 229.7), (None, 264)]),
}
PLAN_B = {  # 体系見直し（区分を4つに、6〜10㎥を引上げ、11㎥の段差を緩和）
    0.10: dict(base=1145, tiers=[(10, 110), (30, 180), (50, 205), (None, 240)]),
    0.15: dict(base=1190, tiers=[(10, 115), (30, 190), (50, 215), (None, 250)]),
    0.20: dict(base=1235, tiers=[(10, 120), (30, 200), (50, 225), (None, 260)]),
}


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
    rows = [['#', '1か月あたり・税込。基本使用料は5㎥まで。案Bは区分を4つに増やし、'
                  '6〜10㎥の単価を引き上げて11㎥での段差を緩和している'],
            ['料金案', '基本使用料(5㎥まで)', '6〜10㎥', '11〜30㎥', '31〜50㎥', '51㎥〜',
             '11㎥での単価の跳ね上がり']]
    rows.append(['現行', 1108.8, 40.7, 191.4, 191.4, 220.0, '4.70倍'])
    for uplift in (0.10, 0.15, 0.20):
        a = PLAN_A[uplift]
        rows.append(['案A %+.0f%%（体系維持）' % (uplift * 100), a['base'], a['tiers'][0][1],
                     a['tiers'][1][1], a['tiers'][1][1], a['tiers'][2][1],
                     '%.2f倍' % (a['tiers'][1][1] / a['tiers'][0][1])])
    for uplift in (0.10, 0.15, 0.20):
        b = PLAN_B[uplift]
        rows.append(['案B %+.0f%%（体系見直し）' % (uplift * 100), b['base'], b['tiers'][0][1],
                     b['tiers'][1][1], b['tiers'][2][1], b['tiers'][3][1],
                     '%.2f倍' % (b['tiers'][1][1] / b['tiers'][0][1])])
    write(os.path.join(out, '15_料金表の比較.csv'), rows)

    # === 増収と経費回収率の比較 =============================================
    rows = [['#', '増収額はR7奇数月6回調定に各案を当てはめた再計算（税込）。'
                  '経費回収率・繰入金はR6決算統計に増収率を乗じたもので、汚水処理費はR6水準で据置'],
            ['#', '水洗便所等普及費330千円（漁集は不明水処理費2,318千円）は現行額が継続する前提'],
            ['料金案', '漁集 調定額(円)', '公共 調定額(円)', '合計(円)', '増収額(円)', '増収率(%)',
             '公共 経費回収率(%)', '漁集 経費回収率(%)',
             '公共 使用料対象資本費(千円)', '公共 基準内繰入金(千円)',
             '漁集 基準内繰入金(千円)', '2事業 増収額(千円・税抜換算)']]
    base_total = cur_g + cur_k
    plans = [('現行', CURRENT, 0.0)]
    for uplift in (0.10, 0.15, 0.20):
        plans.append(('案A %+.0f%%（体系維持）' % (uplift * 100), PLAN_A[uplift], uplift))
    for uplift in (0.10, 0.15, 0.20):
        plans.append(('案B %+.0f%%（体系見直し）' % (uplift * 100), PLAN_B[uplift], uplift))
    for label, plan, uplift in plans:
        g = revenue(gyo, plan, 1 + uplift)
        k = revenue(kou, plan, 1 + uplift)
        u_g = g / cur_g - 1
        u_k = k / cur_k - 1
        rk = recovery('公共下水道', u_k)
        rg = recovery('漁業集落排水', u_g)
        rows.append([label, g, k, g + k, g + k - base_total,
                     '%.2f' % ((g + k) / base_total * 100 - 100),
                     '%.1f' % rk['経費回収率'], '%.1f' % rg['経費回収率'],
                     rk['使用料対象資本費'], rk['基準内繰入金'], rg['基準内繰入金'],
                     rk['増収額'] + rg['増収額']])
    write(os.path.join(out, '16_増収と経費回収率の比較.csv'), rows)

    # === モデルケースと単価の公平性 =========================================
    rows = [['#', '一般汚水。1か月あたりの使用料（税込）と、その使用量における単価（円/㎥）'],
            ['#', '現在は月10㎥の層の単価が最も低く（131.2円/㎥）、月5㎥の層（221.8円/㎥）を下回っている'],
            ['月使用量(㎥)', '現行(円)', '案A+10%(円)', '案B+10%(円)',
             '現行との差 案A', '現行との差 案B', '増減率 案A(%)', '増減率 案B(%)',
             '現行 単価(円/㎥)', '案A+10% 単価', '案B+10% 単価']]
    for v in (5, 8, 10, 12, 15, 20, 25, 30, 40, 50, 60, 80, 100):
        c = monthly(v, **CURRENT)
        a = monthly(v, **PLAN_A[0.10])
        b = monthly(v, **PLAN_B[0.10])
        rows.append([v, round(c, 1), round(a, 1), round(b, 1),
                     '%+.1f' % (a - c), '%+.1f' % (b - c),
                     '%+.2f' % ((a / c - 1) * 100), '%+.2f' % ((b / c - 1) * 100),
                     '%.1f' % (c / v), '%.1f' % (a / v), '%.1f' % (b / v)])
    write(os.path.join(out, '17_モデルケースと単価の公平性.csv'), rows)
    return 0


if __name__ == '__main__':
    sys.exit(main())
