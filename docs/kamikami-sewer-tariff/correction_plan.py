#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""6〜10㎥区分の単価差縮小案のエビデンスを作成する（v2）。

説明資料「下水道使用料の改定について ― 6〜10㎥区分の単価差縮小を含む料金体系の見直し ―」
に載せた数値は、すべてこのスクリプトで再現できる。

v2での変更（外部レビューの指摘反映）
  1. 条例別表は税抜のため、料金案を「税抜の整数単価」で設計するよう変更した。
     税込表示は税抜単価から導く（条例第17条：税抜額に消費税を加算し1円未満切捨て）。
  2. 「推奨案」を「中心案（事務局たたき台）」に改めた。
  3. 「是正率」を「単価差縮小率（現行11〜50㎥単価174円への到達率）」に改めた。
  4. 6〜10㎥を174円まで引き上げた案は、11〜50㎥がそれを下回る逆転体系になるため
     「完全是正」ではなく「参考」として扱う。
  5. 増収額は平年度ベースである旨を明示した（施行日未定のため初年度額は別途）。

使い方:
    python3 correction_plan.py <【R7】使用料集計ブック.xlsx> [出力ディレクトリ]
"""
import csv
import io
import math
import os
import sys

import openpyxl

# --- 現行の条例別表（税抜・1か月あたり） -------------------------------------
# 階上町公共下水道条例／漁業集落排水処理施設条例の別表。いずれも同額。
BASE_NET = 1008        # 基本使用料（5㎥まで）
LOW_NET = 37           # 6〜10㎥（1㎥につき）
MID_NET = 174          # 11〜50㎥（1㎥につき）
HIGH_NET = 200         # 51㎥〜（1㎥につき）
CURRENT = (BASE_NET, LOW_NET, MID_NET, HIGH_NET)

# 料金案（税抜整数）。11〜50㎥は目標増収率から逆算して整数に丸めたもの
PLANS = {
    '中心案（平均増収率 約10%）': (1008, 73, 190, 220),
    '参考 平均増収率 約5%': (1008, 55, 182, 210),
    '参考 平均増収率 約15%': (1008, 82, 205, 230),
}
STAGE1 = (1008, 73, 190, 220)
STAGE2 = (1008, 109, 207, 240)

# 令和6年度決算統計（32表）。参考経費回収率の算定に用いる
R6 = {'公共下水道': dict(fee=33652, cost=67464), '漁業集落排水': dict(fee=7229, cost=19614)}


def net_2months(volume, base, rate_low, rate_mid, rate_high):
    """2か月分の税抜額。隔月検針のため月額表の区分境界を2倍して適用する。"""
    amount = base * 2
    if volume > 10:
        amount += rate_low * min(volume - 10, 10)
    if volume > 20:
        amount += rate_mid * min(volume - 20, 80)
    if volume > 100:
        amount += rate_high * (volume - 100)
    return amount


def bill(volume, *rates):
    """2か月分の請求額（税込）。条例第17条により1円未満を切り捨てる。"""
    return math.floor(net_2months(volume, *rates) * 1.1)


def monthly_net(volume, base, rate_low, rate_mid, rate_high):
    """1か月あたりの税抜使用料。モデルケース表示用。"""
    amount = base
    if volume > 5:
        amount += rate_low * min(volume - 5, 5)
    if volume > 10:
        amount += rate_mid * min(volume - 10, 40)
    if volume > 50:
        amount += rate_high * (volume - 50)
    return amount


EXACT_VOLUME = {bill(v, *CURRENT): v for v in range(11, 4000)}
BASIC_AMOUNT = bill(10, *CURRENT)
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


def revenue(records, rates, ratio):
    """料金案を当てはめた年間調定額。特殊算定は現行額×改定率で扱う。"""
    total = 0
    for kind, volume, count, current in records:
        if kind == NORMAL:
            total += bill(volume, *rates) * count
        elif kind == BASIC:
            total += bill(10, *rates) * count
        else:
            total += round(current * ratio)
    return total


def metered_volume(records):
    """基本水量（2か月10㎥）を除く従量課金対象水量。"""
    low = mid = high = 0
    for kind, volume, count, _ in records:
        if kind in (SPECIAL, BASIC):
            continue
        low += max(0, min(volume, 20) - 10) * count
        mid += max(0, min(volume, 100) - 20) * count
        high += max(0, volume - 100) * count
    return low, mid, high


def reduction_ratio(rate_low):
    """単価差縮小率＝現行11〜50㎥単価（174円）への到達率。"""
    return (rate_low - LOW_NET) / (MID_NET - LOW_NET) * 100


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
    cur_g, cur_k = sum(a for _, _, _, a in gyo), sum(a for _, _, _, a in kou)
    total_current = cur_g + cur_k

    def total(rates, ratio):
        return revenue(gyo, rates, ratio) + revenue(kou, rates, ratio)

    # === 20 単価差の縮小余地 ================================================
    low_g, mid_g, high_g = metered_volume(gyo)
    low_k, mid_k, high_k = metered_volume(kou)
    low, mid, high = low_g + low_k, mid_g + mid_k, high_g + high_k
    levelled = total((BASE_NET, MID_NET, MID_NET, HIGH_NET), 1.0)
    write(os.path.join(out, '20_単価差の縮小余地.csv'), [
        ['#', '6〜10㎥を現行の11〜50㎥単価（税抜174円）まで引き上げた場合の増収余地'],
        ['#', '「従量課金対象水量」は、基本使用料に含まれる水量（2か月10㎥）を除いた水量。'
              '全汚水量ではない'],
        ['項目', '2事業計', '公共下水道', '漁業集落排水', '備考'],
        ['従量課金対象水量 6〜10㎥帯(㎥)', low, low_k, low_g, '2か月分の合計'],
        ['従量課金対象水量 11〜50㎥帯(㎥)', mid, mid_k, mid_g, ''],
        ['従量課金対象水量 51㎥〜帯(㎥)', high, high_k, high_g, ''],
        ['従量課金対象水量 計(㎥)', low + mid + high, low_k + mid_k + high_k, low_g + mid_g + high_g,
         '基本水量を除く'],
        ['6〜10㎥帯の水量シェア(%)', '%.1f' % (low / (low + mid + high) * 100),
         '%.1f' % (low_k / (low_k + mid_k + high_k) * 100),
         '%.1f' % (low_g / (low_g + mid_g + high_g) * 100), '従量課金対象水量に占める割合'],
        [],
        ['6〜10㎥の現行単価（税抜）', LOW_NET, '', '', '税込 %.1f円' % (LOW_NET * 1.1)],
        ['11〜50㎥の現行単価（税抜）', MID_NET, '', '', '税込 %.1f円' % (MID_NET * 1.1)],
        ['単価の開き（倍）', '%.2f' % (MID_NET / LOW_NET), '', '', '174 ÷ 37'],
        ['現行調定額（税込）', total_current, cur_k, cur_g, 'R7年度奇数月6回調定'],
        ['6〜10㎥を174円に揃えた場合の調定額', levelled, '', '', '他の区分は現行のまま'],
        ['増収余地（税込）', levelled - total_current, '', '', '平年度ベース'],
        ['現行比(%)', '%.2f' % ((levelled / total_current - 1) * 100), '', '', ''],
    ])

    # === 21 単価差縮小の水準別比較 ==========================================
    LEVELS = [('縮小なし', 37), ('弱い縮小', 55), ('中心案', 73), ('強い縮小', 100),
              ('参考：174円まで引上げ', 174)]
    rows = [['#', '調定額ベースの平均増収率が約10%になる前提で、'
                  '6〜10㎥をどこまで引き上げるかを比較（税抜単価）'],
            ['#', '基本使用料は1,008円で据置、51㎥〜は220円（+10.0%）で固定。'
                  '11〜50㎥は目標増収率から逆算して整数に丸めた'],
            ['#', '「参考：174円まで引上げ」は11〜50㎥が6〜10㎥を下回る逆転体系になるため、'
                  '実施案ではなく比較の端点として示す'],
            ['水準', '6〜10㎥(税抜)', '6〜10㎥(税込)', '6〜10㎥の改定率(%)',
             '単価差縮小率(%)', '11〜50㎥(税抜)', '11〜50㎥の改定率(%)', '51㎥〜(税抜)',
             '平均増収率(%)', '月5㎥', '月8㎥', '月10㎥', '月15㎥', '月20㎥', '月30㎥', '月50㎥', '月100㎥']]
    for label, rate_low in LEVELS:
        best = None
        for rate_mid in range(120, 301):
            t = total((BASE_NET, rate_low, rate_mid, 220), 1.10)
            gap = abs(t / total_current - 1 - 0.10)
            if best is None or gap < best[0]:
                best = (gap, rate_mid, t)
        _, rate_mid, t = best
        rates = (BASE_NET, rate_low, rate_mid, 220)
        row = [label, rate_low, '%.1f' % (rate_low * 1.1), '%+.1f' % ((rate_low / LOW_NET - 1) * 100),
               '%.1f' % reduction_ratio(rate_low), rate_mid,
               '%+.1f' % ((rate_mid / MID_NET - 1) * 100), 220,
               '%+.2f' % ((t / total_current - 1) * 100)]
        for v in (5, 8, 10, 15, 20, 30, 50, 100):
            row.append('%+.1f' % ((monthly_net(v, *rates) / monthly_net(v, *CURRENT) - 1) * 100))
        rows.append(row)
    write(os.path.join(out, '21_単価差縮小の水準別比較.csv'), rows)

    # === 22 料金案と平年度増収 ==============================================
    rows = [['#', '料金案は条例別表と同じ税抜単価で設計し、税込は条例第17条により'
                  '税抜額に消費税を加算して1円未満を切り捨てて算定'],
            ['#', '増収額は平年度ベース（R7年度の1年分相当の調定分布に新料金を当てたもの）。'
                  '施行日が未定のため初年度影響額は別途算定が必要'],
            ['#', '経費回収率はR6決算の使用料・汚水処理費を固定した静態試算の参考値'],
            ['料金案', '基本使用料(税抜)', '6〜10㎥(税抜)', '11〜50㎥(税抜)', '51㎥〜(税抜)',
             '基本(税込)', '6〜10㎥(税込)', '11〜50㎥(税込)', '51㎥〜(税込)',
             '6〜10㎥の改定率(%)', '11〜50㎥の改定率(%)', '51㎥〜の改定率(%)', '単価差縮小率(%)',
             '漁集 調定額(円)', '公共 調定額(円)', '合計(円)', '平年度増収額(円)', '平均増収率(%)',
             '参考経費回収率 公共(%)', '参考経費回収率 漁集(%)']]
    rows.append(['現行'] + list(CURRENT) + ['%.1f' % (r * 1.1) for r in CURRENT] +
                ['0.0', '0.0', '0.0', '0.0', cur_g, cur_k, total_current, 0, '0.00',
                 '%.1f' % (R6['公共下水道']['fee'] / R6['公共下水道']['cost'] * 100),
                 '%.1f' % (R6['漁業集落排水']['fee'] / R6['漁業集落排水']['cost'] * 100)])
    for label, rates in list(PLANS.items()) + [('第2段階（累計 約20%）', STAGE2)]:
        ratio = 1 + (0.20 if '20' in label else float(label.split('約')[1].rstrip('%）')) / 100)
        g, k = revenue(gyo, rates, ratio), revenue(kou, rates, ratio)
        rk = R6['公共下水道']['fee'] * (k / cur_k) / R6['公共下水道']['cost'] * 100
        rg = R6['漁業集落排水']['fee'] * (g / cur_g) / R6['漁業集落排水']['cost'] * 100
        rows.append([label] + list(rates) + ['%.1f' % (r * 1.1) for r in rates] +
                    ['%+.1f' % ((rates[1] / LOW_NET - 1) * 100),
                     '%+.1f' % ((rates[2] / MID_NET - 1) * 100),
                     '%+.1f' % ((rates[3] / HIGH_NET - 1) * 100),
                     '%.1f' % reduction_ratio(rates[1]),
                     g, k, g + k, g + k - total_current,
                     '%+.2f' % ((g + k) / total_current * 100 - 100), '%.1f' % rk, '%.1f' % rg])
    write(os.path.join(out, '22_料金案と平年度増収.csv'), rows)

    # === 23 世帯への影響 ====================================================
    rows = [['#', '一般汚水・税込。調定は隔月のため、請求額は月額の2倍'],
            ['#', '公衆浴場汚水は別単価（基本1,008円＋6㎥〜57円／いずれも税抜）だが、'
                  'R7調定データ上は適用者を確認できないため試算対象外'],
            ['月使用量(㎥)', '世帯像の目安', '現行 月額(税込)', '現行 2か月請求',
             '中心案 月額(税込)', '中心案 2か月請求', '差額(2か月)', '増減率(%)',
             '第2段階 月額(税込)', '第2段階 2か月請求', '現行比 増減率(%)']]
    for v, who in [(3, '単身（節水）'), (5, '単身・高齢世帯'), (8, '2人世帯'), (10, '2人世帯'),
                   (13, '3人世帯'), (15, '3〜4人世帯'), (20, '4人世帯'), (25, '5人世帯'),
                   (30, '5人以上世帯'), (40, '小規模事業所'), (50, '小規模事業所'),
                   (100, '大口事業所'), (200, '大口事業所')]:
        c_m, c_2 = monthly_net(v, *CURRENT) * 1.1, bill(2 * v, *CURRENT)
        r_m, r_2 = monthly_net(v, *STAGE1) * 1.1, bill(2 * v, *STAGE1)
        s_m, s_2 = monthly_net(v, *STAGE2) * 1.1, bill(2 * v, *STAGE2)
        rows.append([v, who, round(c_m, 1), c_2, round(r_m, 1), r_2, r_2 - c_2,
                     '%+.1f' % ((r_m / c_m - 1) * 100), round(s_m, 1), s_2,
                     '%+.1f' % ((s_m / c_m - 1) * 100)])
    write(os.path.join(out, '23_世帯への影響.csv'), rows)

    # === 24 段階的な縮小 ====================================================
    rows = [['#', '経営戦略が見込む使用料収入の水準（令和8年度 約+10%、令和13年度 累計約+20%）に'
                  '合わせて単価差を段階的に縮小する場合'],
            ['#', '経営戦略が検証しているのは使用料収入の増加であり、区分間の負担配分までを'
                  '決定しているわけではない'],
            ['段階', '基本(税抜)', '6〜10㎥(税抜)', '6〜10㎥の改定率(%)', '単価差縮小率(%)',
             '11〜50㎥(税抜)', '51㎥〜(税抜)', '単価の開き(倍)', '平均増収率(%)', '平年度増収額(円)']]
    rows.append(['現行'] + list(CURRENT[:2]) + ['0.0', '0.0'] + list(CURRENT[2:]) +
                ['%.2f' % (MID_NET / LOW_NET), '0.00', 0])
    for label, rates, ratio in [('第1段階（令和8年度）', STAGE1, 1.10),
                                ('第2段階（令和13年度）', STAGE2, 1.20)]:
        t = total(rates, ratio)
        rows.append([label, rates[0], rates[1], '%+.1f' % ((rates[1] / LOW_NET - 1) * 100),
                     '%.1f' % reduction_ratio(rates[1]), rates[2], rates[3],
                     '%.2f' % (rates[2] / rates[1]),
                     '%+.2f' % ((t / total_current - 1) * 100), t - total_current])
    write(os.path.join(out, '24_段階的な縮小.csv'), rows)

    # === 25 条例改正に向けた整理 ============================================
    write(os.path.join(out, '25_条例改正に向けた整理.csv'), [
        ['#', '階上町公共下水道条例／漁業集落排水処理施設条例の別表は税抜で定められており、'
              '第17条により税抜額に消費税及び地方消費税を加算して1円未満を切り捨てる'],
        ['#', '改正対象は2条例（公共下水道・漁業集落排水）。いずれも別表は同額'],
        ['区分', '現行(税抜)', '中心案(税抜)', '第2段階(税抜)',
         '現行(税込・参考)', '中心案(税込・参考)', '第2段階(税込・参考)'],
        ['基本使用料（5㎥まで）', BASE_NET, STAGE1[0], STAGE2[0],
         '%.1f' % (BASE_NET * 1.1), '%.1f' % (STAGE1[0] * 1.1), '%.1f' % (STAGE2[0] * 1.1)],
        ['6〜10㎥（1㎥につき）', LOW_NET, STAGE1[1], STAGE2[1],
         '%.1f' % (LOW_NET * 1.1), '%.1f' % (STAGE1[1] * 1.1), '%.1f' % (STAGE2[1] * 1.1)],
        ['11〜50㎥（1㎥につき）', MID_NET, STAGE1[2], STAGE2[2],
         '%.1f' % (MID_NET * 1.1), '%.1f' % (STAGE1[2] * 1.1), '%.1f' % (STAGE2[2] * 1.1)],
        ['51㎥〜（1㎥につき）', HIGH_NET, STAGE1[3], STAGE2[3],
         '%.1f' % (HIGH_NET * 1.1), '%.1f' % (STAGE1[3] * 1.1), '%.1f' % (STAGE2[3] * 1.1)],
        [],
        ['公衆浴場汚水 基本使用料（5㎥まで）', BASE_NET, '未設定', '未設定',
         '%.1f' % (BASE_NET * 1.1), '', ''],
        ['公衆浴場汚水 6㎥〜（1㎥につき）', 57, '未設定', '未設定', '62.7', '', ''],
        ['#', '公衆浴場汚水はR7調定データ上は適用者を確認できない。'
              '据置とするか一般汚水と同率で改定するかを別途決定する必要がある'],
    ])
    return 0


if __name__ == '__main__':
    sys.exit(main())
