# -*- coding: utf-8 -*-
"""第10期 要支援・要介護認定者数の推計
   人口：厚労省配布「第10期将来推計用人口推計」（住基ベース）＋前期・後期別補正係数
   認定率：見える化 B4-a（全体）・B4-d（75歳以上）から前期・後期別に分解
   出力：data/第10期_要介護度別認定者数推計.csv
"""
import csv, os, collections

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TIDY = os.path.join(BASE, 'data', 'mieruka_tidy.csv')
OUT  = os.path.join(BASE, 'data', '第10期_要介護度別認定者数推計.csv')

KAIGO = ['要支援１','要支援２','経過的要介護','要介護１','要介護２','要介護３','要介護４','要介護５']
KAIGO_D = ['要支援1','要支援2','経過的要介護','要介護1','要介護2','要介護3','要介護4','要介護5']

# 第1号被保険者数の実績（介護保険事業状況報告。国配布の補正シートより）
JISSEKI = {'前期': 486, '後期': 526}          # 令和7年12月末
# 配布データ（前期・後期別補正後・各年1月1日時点）
SUIKEI = {
 '令和7年': (463, 498), '令和8年': (444, 509), '令和9年': (425, 521),
 '令和10年': (406, 533), '令和11年': (387, 545), '令和12年': (371, 547),
 # 中長期推計（基本指針で2040年を見据えた推計が求められる）
 '令和17年': (309, 535), '令和22年': (283, 488),
}
# B案（令和8年の実績に水準を合わせるリベース係数）
REBASE = {'前期': JISSEKI['前期'] / SUIKEI['令和8年'][0],
          '後期': JISSEKI['後期'] / SUIKEI['令和8年'][1]}

def load():
    rows = [r for r in csv.DictReader(open(TIDY, encoding='utf-8-sig'))
            if '北塩原' in r['region']]
    def get(code, year):
        return {r['indicator']: float(r['value'])
                for r in rows if r['code'] == code and r['year'] == year and r['value'] not in ('', None)}
    return get

def certification_rates(get, years=('2024', '2025', '2026')):
    """前期・後期別の要介護度別認定率（3か年平均）を返す。
       後期＝B4-d（75歳以上）／前期＝B4-a（全体）−B4-d を前期被保険者数で割る。"""
    # 各年の被保険者数（全体・75歳以上）は認定者数÷認定率から復元する
    zen_n, kou_n = collections.defaultdict(list), collections.defaultdict(list)
    for y in years:
        a, d = get('B4-a', y), get('B4-d', y)
        if not a or not d:
            continue
        # 分母の復元：合計認定者数 ÷ 合計認定率(%) × 100
        tot_all = a.get('合計認定者数'); rate_all = a.get('合計認定率')
        tot_75 = d.get('合計認定者数');  rate_75 = d.get('合計認定率')
        if not all((tot_all, rate_all, tot_75, rate_75)):
            continue
        pop_all = tot_all / rate_all * 100
        pop_75 = tot_75 / rate_75 * 100
        pop_zen = pop_all - pop_75
        if pop_zen <= 0:
            continue
        for k_a, k_d in zip(KAIGO, KAIGO_D):
            n_all = a.get(f'認定者数（{k_a}）', 0.0)
            n_75 = d.get(f'認定者数（{k_d}）', 0.0)
            kou_n[k_a].append(n_75 / pop_75)
            zen_n[k_a].append(max(n_all - n_75, 0.0) / pop_zen)
    avg = lambda xs: sum(xs) / len(xs) if xs else 0.0
    return ({k: avg(v) for k, v in zen_n.items()},
            {k: avg(v) for k, v in kou_n.items()})

def project(rate_zen, rate_kou, case):
    rows = []
    for year, (z, k) in SUIKEI.items():
        if case == 'B':
            z = z * REBASE['前期']; k = k * REBASE['後期']
        rec = {'案': case, '年': year, '前期高齢者': round(z), '後期高齢者': round(k),
               '第1号被保険者計': round(z + k)}
        tot = 0.0
        for kg in KAIGO:
            v = z * rate_zen.get(kg, 0.0) + k * rate_kou.get(kg, 0.0)
            rec[kg] = round(v); tot += v
        rec['認定者計'] = round(tot)
        rec['認定率(%)'] = round(tot / (z + k) * 100, 1)
        rows.append(rec)
    return rows

RATE_BASES = {'令和8年': ('2026',), '3か年平均': ('2024', '2025', '2026')}

if __name__ == '__main__':
    get = load()
    rz, rk = certification_rates(get, RATE_BASES['令和8年'])
    print('■ 要介護度別認定率（3か年平均・％）')
    print(f"{'区分':<8}{'前期(65-74)':>12}{'後期(75+)':>12}")
    for kg in KAIGO:
        print(f'{kg:<8}{rz.get(kg,0)*100:>11.2f}%{rk.get(kg,0)*100:>11.2f}%')
    print(f"{'合計':<8}{sum(rz.values())*100:>11.2f}%{sum(rk.values())*100:>11.2f}%")
    print()
    all_rows = []
    for base, yrs in RATE_BASES.items():
        a, b = certification_rates(get, yrs)
        for case in ('A', 'B'):
            for r in project(a, b, case):
                r['認定率の基準'] = base
                all_rows.append(r)
    fields = ['認定率の基準', '案', '年', '前期高齢者', '後期高齢者', '第1号被保険者計'] + KAIGO + ['認定者計', '認定率(%)']
    with open(OUT, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=fields); w.writeheader(); w.writerows(all_rows)
    for case in ('A', 'B'):
        print(f'■ {case}案（認定率の基準：令和8年）')
        print(f"{'年':<8}{'第1号':>7}{'認定者計':>8}{'認定率':>7}  " + ''.join(f'{k:>7}' for k in KAIGO))
        for r in all_rows:
            if r['案'] != case or r['認定率の基準'] != '令和8年': continue
            print(f"{r['年']:<8}{r['第1号被保険者計']:>7}{r['認定者計']:>8}{r['認定率(%)']:>6}%  "
                  + ''.join(f"{r[k]:>7}" for k in KAIGO))
        print()
    print('保存:', OUT)
