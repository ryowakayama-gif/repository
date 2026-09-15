# -*- coding: utf-8 -*-
"""見込量の自己点検（国の推計ツールのワーニングチェックに相当するもの）

【国の判定】
  変化率（％）＝（当該年度の利用者数 － 前年度の利用者数）÷ 前年度の利用者数 × 100
  閾値 ＝ Q1 － 2×(Q3－Q1) 〜 Q3 ＋ 2×(Q3－Q1)
  Q1・Q3は全国の保険者の変化率の四分位。ツールの出力に併載されるため、
  受領前に本村の値を全国の分布と照らして判定することはできない。

【本稿で行うこと】
  全国の分布が分からなくても、本村の推計値が年度間でどれだけ動くかは計算できる。
  動きの大きい項目をあらかじめ洗い出し、理由を用意しておく。
  受給者が1人から数人のサービスでは1人の増減が変化率を大きく動かすため、
  変化の実数（人・月）も併せて示す。
"""
import csv, os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, 'data')

YEARS = ['令和9年度', '令和10年度', '令和11年度']
ACT = '令和7年度実績'
FLAG = 10.0        # 本稿で「動きが大きい」とみなす変化率（％）


def fnum(v):
    try:
        return float(str(v).replace(',', ''))
    except (TypeError, ValueError):
        return None


def rate(a, b):
    """bからaへの変化率（％）。bが0のときは None を返す。"""
    if a is None or b is None:
        return None
    if b == 0:
        return None if a == 0 else float('inf')
    return (a - b) / b * 100.0


def main():
    out = []
    p = out.append

    # ── 1 区分別（認定者数・在宅・居住系・施設）────────────
    p('■ 1　国の判定の対象となる3区分と認定者数')
    kubun = {}
    with open(os.path.join(DATA, '第10期_サービス区分別受給者数推計.csv'),
              encoding='utf-8-sig') as f:
        for r in csv.DictReader(f):
            if r['要介護度'] != '計':
                continue
            kubun[r['区分']] = {'令和7年度実績': fnum(r['令和7年度実績'])} | \
                               {y: fnum(r[y]) for y in YEARS}
    nintei = {'令和7年度実績': 214.0, '令和9年度': 214.0,
              '令和10年度': 217.0, '令和11年度': 220.0}   # 5-3（令和7年度実績＝令和8年3月末）
    p(f"  {'区分':<16}{'R7実績':>9}{'R9':>9}{'R10':>9}{'R11':>9}"
      f"{'R7→R9':>10}{'R9→R10':>10}{'R10→R11':>10}")
    series = [('要支援・要介護認定者数', nintei)] + list(kubun.items())
    for name, d in series:
        v = [d.get(ACT)] + [d.get(y) for y in YEARS]
        rs = [rate(v[1], v[0]), rate(v[2], v[1]), rate(v[3], v[2])]
        p(f"  {name:<16}" + ''.join(f'{x:>8.1f}' if x is not None else f"{'―':>9}" for x in v)
          + ''.join(f'{x:>9.1f}%' if x is not None and x != float('inf') else f"{'―':>10}" for x in rs))
    p('  ※R7→R9は2年分の変化です。計画期間の初年度は令和9年度であるため、')
    p('　　国のツールでは令和8年度の値を挟んで判定されます。令和8年度は年度の途中であり、')
    p('　　確定した実績がないため本稿では令和7年度実績から直接比較しています。')

    # ── 2 サービス種類別 ───────────────────────
    p('')
    p('■ 2　サービス種類別の受給者数の変化率')
    rows = []
    with open(os.path.join(DATA, '第10期_サービス種類別見込量.csv'), encoding='utf-8-sig') as f:
        for r in csv.DictReader(f):
            act = fnum(r.get('令和7年度実績(人/月)'))
            est = [fnum(r.get(y + '(人/月)')) for y in YEARS]
            rows.append((r['区分'], r['サービス種類'], act, est))

    p(f"  {'サービス':<24}{'R7実績':>8}{'R9':>7}{'R10':>7}{'R11':>7}"
      f"{'R7→R9':>10}{'R9→R10':>9}{'R10→R11':>9}  判定")
    flagged = []
    for kb, name, act, est in rows:
        v = [act] + est
        rs = [rate(v[1], v[0]), rate(v[2], v[1]), rate(v[3], v[2])]
        big = [x for x in rs if x is not None and (x == float('inf') or abs(x) > FLAG)]
        mark = '★' if big else ''
        if big:
            flagged.append((kb, name, act, est, rs))
        def f1(x):
            if x is None:
                return f"{'―':>9}"
            if x == float('inf'):
                return f"{'新規':>9}"
            return f'{x:>8.1f}%'
        p(f"  {name:<24}" + ''.join(f'{x:>7.1f}' if x is not None else f"{'―':>8}" for x in v)
          + f1(rs[0]) + f1(rs[1]) + f1(rs[2]) + f'  {mark}')

    p('')
    p(f'■ 3　動きが大きい項目（いずれかの年で変化率±{FLAG:.0f}%超）')
    if not flagged:
        p('  該当なし')
    for kb, name, act, est, rs in flagged:
        d = (est[0] or 0) - (act or 0)
        p(f'  ・{name}（{kb}）')
        p(f'　　令和7年度 {act:.1f}人／月 → 令和9年度 {est[0]:.1f}人／月（{d:+.1f}人）')
        p(f'　　変化の実数が{abs(d):.1f}人であり、受給者の少なさが変化率を大きくしている。')

    p('')
    p('■ 4　理由の整理')
    p('  受給率を3か年平均で固定しているため、令和7年度の単年度の値が3か年平均から')
    p('  離れているサービスほど、令和7年度実績から令和9年度見込への変化率が大きくなる。')
    p('  年度間（令和9→10→11）の変化率は認定者数の変化率（年1.4%程度）に収まり、')
    p('  閾値を外れる可能性は低い。')

    path = os.path.join(DATA, '第10期_見込量の自己点検.csv')
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['区分', 'サービス種類', 'R7実績(人/月)', 'R9(人/月)', 'R10(人/月)',
                    'R11(人/月)', 'R7→R9(%)', 'R9→R10(%)', 'R10→R11(%)', '判定'])
        for kb, name, act, est in rows:
            v = [act] + est
            rs = [rate(v[1], v[0]), rate(v[2], v[1]), rate(v[3], v[2])]
            big = any(x is not None and (x == float('inf') or abs(x) > FLAG) for x in rs)
            w.writerow([kb, name, act] + est +
                       [('' if x is None else ('新規' if x == float('inf') else round(x, 1))) for x in rs] +
                       ['★' if big else ''])
    print('\n'.join(out))
    print(f'\n保存: {os.path.relpath(path, BASE)}')


if __name__ == '__main__':
    main()
