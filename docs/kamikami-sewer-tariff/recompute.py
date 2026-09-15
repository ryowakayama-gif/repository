#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""令和7年度 調定データの区分・水量・収入の確定モデル。

条例（別表・第17条／第19条・第18条第3項/第4項／第20条第3項/第4項）に忠実に実装し、
集計ブックの水量欄と算定式を照合して各行の検針区分（毎月／隔月）を判定する。

入力は 00_入力データ/ の3ファイル（いずれも集計値で個人情報を含まない）。
    R7_使用料集計ブック_公共.csv
    R7_使用料集計ブック_漁集.csv
    R7_日割実額.csv
場所はスクリプトからの相対で解決する。別の場所に置く場合は環境変数
KAMIKAMI_DATA でディレクトリを指定する。

    $ python3 recompute.py            # 自己診断（主要な確定値を表示）
"""
import csv
import io
import math
import os

HERE = os.path.dirname(os.path.abspath(__file__))


def data_dir():
    """入力データのディレクトリを解決する。"""
    env = os.environ.get('KAMIKAMI_DATA')
    if env:
        return env
    for rel in ('00_入力データ', '../00_入力データ', '../../00_入力データ', '.'):
        d = os.path.normpath(os.path.join(HERE, rel))
        if os.path.exists(os.path.join(d, 'R7_日割実額.csv')):
            return d
    raise SystemExit(
        '入力データが見つかりません。00_入力データ/ を配置するか、'
        '環境変数 KAMIKAMI_DATA でディレクトリを指定してください。')


def _rows(fn):
    with io.open(os.path.join(data_dir(), fn), encoding='utf-8-sig') as f:
        for r in csv.reader(f):
            if r and not r[0].startswith('#'):
                yield r


CURRENT = (1008, 37, 174, 200)          # 条例別表（税抜・1か月）


def month_net(v, b=1008, r1=37, r2=174, r3=200):
    """1か月・水量v㎥の税抜額（5㎥まで基本／5超10まで／10超50まで／50超）"""
    a = b
    if v > 5:
        a += r1 * (min(v, 10) - 5)
    if v > 10:
        a += r2 * (min(v, 50) - 10)
    if v > 50:
        a += r3 * (v - 50)
    return a


def bill(v, months, *rates):
    """各月均等按分（第18条第3項）＋1円未満切捨て・2か月一括（第17条ただし書）"""
    return math.floor(months * month_net(v, *(rates or CURRENT)) * 1.1)


def hiwari():
    """基本水量未満の日割額の実測値 {事業: (件数, 合計額, 平均額)}。
    平均額はCSVの表示値（小数1位）ではなく 合計額÷件数 で再計算する。"""
    out = {}
    for r in _rows('R7_日割実額.csv'):
        if r[0] in ('事業',):
            continue
        n, amt = int(r[2]), float(r[3])
        out[r[0]] = (n, amt, amt / n)
    return out


MAIN6 = [1, 3, 5, 7, 9, 11]             # 主検針月（5,7,9,11,1,3月）の列位置
ALL12 = list(range(12))


def parse(biz, cols=ALL12):
    """(種別, 月あたり水量, 月数, 件数, 現行額) に分解。
    種別 N=通常算定 / B=基本水量内 / SPH=基本水量未満の日割 / SP=その他（異動等）

    SPH は第18条第4項（漁集は第20条第4項）の日割であり、基本使用料を日数で
    按分した額である。基本使用料を据え置く料金案では改定後も変わらない。"""
    fn = {'公共下水道': 'R7_使用料集計ブック_公共.csv',
          '漁業集落排水': 'R7_使用料集計ブック_漁集.csv'}[biz]
    _, _, avg = hiwari()[biz if biz in hiwari() else '公共下水道']
    out = []
    for r in _rows(fn):
        if r[0] == '使用料（㎥）':
            continue
        vol = int(r[0]) if r[0] not in ('', None) else None
        amt = int(r[1]) if r[1] not in ('', None) else None
        cnt = sum(int(r[2 + i]) for i in cols if r[2 + i] not in ('', None))
        if not cnt:
            continue
        if amt is None:                      # 「〜2,216」＝基本水量未満の日割
            out.append(('SPH', None, None, cnt, avg * cnt))
            continue
        if vol is not None:
            # 水量欄の意味は行によって異なるため3通りを照合する
            for v, m in ((vol / 2, 2), (vol, 1), (vol, 2)):
                if bill(v, m) == amt:
                    out.append(('N', v, m, cnt, amt * cnt))
                    break
            else:
                out.append(('SP', None, None, cnt, amt * cnt))
            continue
        out.append(('B', None, 2, cnt, amt * cnt) if amt == bill(5, 2)
                   else ('SP', None, None, cnt, amt * cnt))
    return out


def apply_gyo_hiwari(recs):
    """漁集は集計ブックが日割を基本使用料2,217円に切り上げて基本料金帯に含めている
    ため、実測の件数・金額で計上し直す。"""
    n, amt, _ = hiwari()['漁業集落排水']
    out, left = [], n
    for k, v, m, c, cur in recs:
        if k == 'B' and left:
            take = min(c, left)
            left -= take
            if c - take:
                out.append((k, v, m, c - take, cur * (c - take) / c))
        else:
            out.append((k, v, m, c, cur))
    out.append(('SPH', None, None, n, float(amt)))
    return out


def load():
    """確定データ（公共・漁集）。漁集は日割の実額補正を適用する。"""
    return parse('公共下水道'), apply_gyo_hiwari(parse('漁業集落排水'))


def decompose(recs):
    """区分別の収入（税込）と従量課金対象水量。

    各記録の成分（基本・6〜10㎥・11〜50㎥・51㎥〜）は税抜額に1.1を乗じた
    未丸め値だが、実際の請求額は1円未満を切り捨てた額である。区分別収入の
    合計が調定額の実額と一致するよう、記録ごとに切捨て差を成分へ比例配分する。"""
    inc = dict(base=0.0, low=0.0, mid=0.0, high=0.0, hiwari=0.0, sp=0.0)
    vol = dict(low=0.0, mid=0.0, high=0.0)
    for k, v, m, c, cur in recs:
        if k == 'SPH':
            inc['hiwari'] += cur
            continue
        if k == 'SP':
            inc['sp'] += cur
            continue
        if k == 'B':
            inc['base'] += bill(5, 2) * c          # 2,217円（切捨て後）
            continue
        part = {'base': 1008 * m}
        for key, q, rate in (('low',  max(0.0, min(v, 10) - 5),  37),
                             ('mid',  max(0.0, min(v, 50) - 10), 174),
                             ('high', max(0.0, v - 50),          200)):
            vol[key] += q * m * c
            part[key] = rate * q * m
        net = sum(part.values())
        f = bill(v, m) / (net * 1.1) if net else 0.0   # 切捨て差の比例配分
        for key, x in part.items():
            inc[key] += x * 1.1 * f * c
    return inc, vol


def revenue(recs, plan, sp_ratio):
    """料金案planでの調定額。SPH（基本水量未満の日割）は基本使用料を据え置く
    限り改定の影響を受けないため据置とし、SP（異動等）のみ平均改定率で連動させる。"""
    return sum(bill(v, m, *plan) * c if k == 'N'
               else bill(5, 2, *plan) * c if k == 'B'
               else cur if k == 'SPH'
               else cur * sp_ratio
               for k, v, m, c, cur in recs)


def solve(groups, plan, iters=80):
    """異動等（SP）は平均改定率で連動させる（不動点）。日割（SPH）は据置。"""
    base = sum(sum(x[4] for x in g) for g in groups)
    r = 1.0
    for _ in range(iters):
        r = sum(revenue(g, plan, r) for g in groups) / base
    return sum(revenue(g, plan, r) for g in groups), r, base


if __name__ == '__main__':
    K, G = load()
    IK, VK = decompose(K)
    IG, VG = decompose(G)
    I = {k: IK[k] + IG[k] for k in IK}
    V = {k: VK[k] + VG[k] for k in VK}
    base = sum(x[4] for x in K) + sum(x[4] for x in G)
    cnt = sum(x[3] for x in K) + sum(x[3] for x in G)
    met = I['low'] + I['mid'] + I['high']
    comp = sum(I.values())
    tv = sum(V.values())
    tot, rate, _ = solve([K, G], (1008, 73, 191, 220))
    print('入力ディレクトリ : %s' % data_dir())
    print('調定件数         : %s 件' % format(cnt, ','))
    print('調定額（税込）   : %s 円' % format(round(base), ','))
    print('区分別収入の合計 : %s 円（実額との差 %s 円）'
          % (format(round(comp), ','), format(round(comp - base), ',')))
    print('従量収入         : %s 円' % format(round(met), ','))
    print('従量課金対象水量 : %s ㎥' % format(round(tv), ','))
    print('6〜10㎥ 従量収入比 %.1f%% / 水量比 %.1f%%'
          % (I['low'] / met * 100, V['low'] / tv * 100))
    print('B案 平年度増収  : %s 円（%+.2f%%）'
          % (format(round(tot - base), ','), (rate - 1) * 100))
