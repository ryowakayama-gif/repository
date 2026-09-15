# -*- coding: utf-8 -*-
"""地域支援事業費の見込み（第1次概算）

【考え方】
  地域支援事業費は、事業の対象がだれかによって延ばす基礎を変える。
    介護予防・生活支援サービス事業費 … 対象は要支援認定者と事業対象者 → 認定者数の伸び
    一般介護予防事業費               … 対象は第1号被保険者全体       → 第1号被保険者数の伸び
    包括的支援事業・任意事業費       … 対象は第1号被保険者全体       → 第1号被保険者数の伸び
  基準は最新の決算である令和5年度。令和6年度・令和7年度の決算の提供を受けた時点で
  基準を置き直す。

【保険料への効き】
  地域支援事業費は②第1号被保険者負担分相当額の算定基礎に含まれる（×α）。
  総合事業費は③調整交付金相当額の算定基礎に含まれる（×5％）とともに、
  調整交付金見込額の算定基礎にもなる。estimate_premium.py の構造による。
"""
import csv, os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, 'data')

# ── 令和5年度決算（介護保険特別会計）──────────────────
R5 = {'介護予防・生活支援サービス事業': 14_637_340,
      '一般介護予防事業':               1_825_335,
      '包括的支援事業・任意事業':      23_423_318}

# ── 延ばす基礎（素案5-2・5-3。いずれも年度末の値）──────────
#    令和5年度末＝令和6年3月末の実績（「見える化」B2-a・B4-a）
NINTEI = {'基準': 200, '令和9年度': 222, '令和10年度': 226, '令和11年度': 230}
INSURED = {'基準': 1_004, '令和9年度': 1_002, '令和10年度': 997, '令和11年度': 990}

YEARS = ['令和9年度', '令和10年度', '令和11年度']
BASIS = {'介護予防・生活支援サービス事業': ('認定者数', NINTEI),
         '一般介護予防事業':               ('第1号被保険者数', INSURED),
         '包括的支援事業・任意事業':       ('第1号被保険者数', INSURED)}

# ── 保険料の算定の前提（estimate_premium.py と同じ）────────
ALPHA = 0.23
KOFU_RATE = 0.0410          # 調整交付金の見込交付割合（第9期から復元）
DENOM = 34_201              # 補正後被保険者数×12か月×予定収納率（人・月）


def main():
    out = []
    p = out.append
    p('■ 延ばす基礎（令和5年度末＝100）')
    p(f"  {'基礎':<18}{'令和5年度末':>12}" + ''.join(f'{y:>12}' for y in YEARS))
    for lab, d in [('認定者数', NINTEI), ('第1号被保険者数', INSURED)]:
        p(f"  {lab:<18}{d['基準']:>10}人" + ''.join(f"{d[y]:>8}人({d[y]/d['基準']*100:>5.1f})" for y in YEARS))

    p('')
    p('■ 地域支援事業費の見込み（単位：千円）')
    p(f"  {'区分':<28}{'R5決算':>11}" + ''.join(f'{y:>12}' for y in YEARS))
    est = {}
    for k, v in R5.items():
        lab, d = BASIS[k]
        est[k] = [v * d[y] / d['基準'] for y in YEARS]
        p(f"  {k:<28}{v/1000:>10,.0f}" + ''.join(f'{e/1000:>11,.0f}' for e in est[k]))
    tot_r5 = sum(R5.values())
    tot = [sum(est[k][i] for k in R5) for i in range(3)]
    p(f"  {'合計':<28}{tot_r5/1000:>10,.0f}" + ''.join(f'{t/1000:>11,.0f}' for t in tot))
    p(f"  {'うち総合事業費':<28}"
      f"{(R5['介護予防・生活支援サービス事業']+R5['一般介護予防事業'])/1000:>10,.0f}"
      + ''.join(f"{(est['介護予防・生活支援サービス事業'][i]+est['一般介護予防事業'][i])/1000:>11,.0f}"
                for i in range(3)))

    p('')
    p('■ 3か年計と、令和5年度決算を据え置いた場合との差')
    three = sum(tot)
    flat = tot_r5 * 3
    p(f'  積み上げ　　　　　　　{three/1000:>12,.0f}千円')
    p(f'  令和5年度決算×3年　　{flat/1000:>12,.0f}千円')
    p(f'  差　　　　　　　　　　{(three-flat)/1000:>+12,.0f}千円')

    sogo3 = sum(est['介護予防・生活支援サービス事業'][i] + est['一般介護予防事業'][i] for i in range(3))
    sogo_flat = (R5['介護予防・生活支援サービス事業'] + R5['一般介護予防事業']) * 3
    d2 = (three - flat) * ALPHA
    d3 = (sogo3 - sogo_flat) * 0.05
    dk = (sogo3 - sogo_flat) * KOFU_RATE
    d4 = d2 + d3 - dk
    p('')
    p('■ 保険料基準額への効き')
    p(f'  ②第1号被保険者負担分相当額の増　{(three-flat)/1000:>10,.0f}千円 × α{ALPHA:.0%} ＝ {d2/1000:>8,.0f}千円')
    p(f'  ③調整交付金相当額の増　　　　　　{(sogo3-sogo_flat)/1000:>10,.0f}千円 × 5％ ＝ {d3/1000:>8,.0f}千円')
    p(f'  　調整交付金見込額の増　　　　　　同上 × {KOFU_RATE:.2%} ＝ {dk/1000:>8,.0f}千円')
    p(f'  ④保険料収納必要額の増　　　　　　　　　　　　　　　　 ＝ {d4/1000:>8,.0f}千円')
    p(f'  → 月額 {d4/DENOM:+.0f}円')

    print('\n'.join(out))

    path = os.path.join(DATA, '第10期_地域支援事業費見込み.csv')
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['区分', '延ばす基礎', 'R5決算(千円)'] + [y + '(千円)' for y in YEARS])
        for k, v in R5.items():
            w.writerow([k, BASIS[k][0], round(v / 1000)] + [round(e / 1000) for e in est[k]])
        w.writerow(['合計', '', round(tot_r5 / 1000)] + [round(t / 1000) for t in tot])
    print(f'\n保存: {os.path.relpath(path, BASE)}')


if __name__ == '__main__':
    main()
