#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""料金体系の設計レバー（基本使用料・基本水量・料金区分）を変えた場合の試算。

recompute.py の確定モデル（令和7年度 年間全件 9,220レコード）をそのまま使い、
料金表の側だけを一般化する。条例の算定規定（第18条第3項の各月均等按分、
第17条ただし書の2か月一括・1円未満切捨て）は recompute.py と同一である。

料金表は (基本使用料, 基本水量, ((上限, 単価), ...)) で表す。上限 None は無制限。
現行 = (1008, 5, ((10, 37), (50, 174), (None, 200)))

    $ python3 structure.py        # 自己診断（現行料金表で recompute と一致するか）
"""
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import recompute as RC

CURRENT = (1008, 5, ((10, 37), (50, 174), (None, 200)))


def net_month(v, plan):
    """1か月・水量v㎥の税抜額。"""
    base, q0, tiers = plan
    a, prev = base, q0
    for upper, rate in tiers:
        if v <= prev:
            break
        hi = v if upper is None else min(v, upper)
        a += rate * (hi - prev)
        prev = hi
    return a


def bill(v, months, plan):
    """各月均等按分（第18条第3項）＋2か月一括・1円未満切捨て（第17条ただし書）。"""
    return math.floor(months * net_month(v, plan) * 1.1)


def base_bill(plan, months=2):
    """基本水量内（水量を特定できない記録）の請求額。"""
    return math.floor(months * plan[0] * 1.1)


def revenue(recs, plan, sp_ratio, base_ratio, b_assume=None):
    """料金表planでの調定額（税込）。

    SPH（基本水量未満の日割）は基本使用料の日数按分であるため、基本使用料の
    改定率 base_ratio に連動する。SP（異動等）は水量不明のため平均改定率で連動。
    B（基本水量内）は水量を特定できないため、基本水量を下げる場合のみ
    b_assume（0〜現行基本水量の想定水量）が必要になる。"""
    t = 0.0
    for k, v, m, c, cur in recs:
        if k == 'N':
            t += bill(v, m, plan) * c
        elif k == 'B':
            if plan[1] >= CURRENT[1]:
                t += base_bill(plan) * c          # 従来どおり基本水量内
            else:
                if b_assume is None:
                    raise ValueError('基本水量を下げる場合は b_assume が必要')
                t += bill(b_assume, 2, plan) * c
        elif k == 'SPH':
            t += cur * base_ratio
        else:
            t += cur * sp_ratio
    return t


def solve(groups, plan, b_assume=None, iters=200):
    """平均改定率とSP（異動等）の連動率を一致させる不動点。

    戻り値 (改定後調定額, 平均改定率, 現行調定額)。"""
    cur = sum(x[4] for g in groups for x in g)
    br = plan[0] / CURRENT[0]
    r = 1.0
    for _ in range(iters):
        t = sum(revenue(g, plan, r, br, b_assume) for g in groups)
        nr = t / cur
        if abs(nr - r) < 1e-12:
            break
        r = nr
    return t, r, cur


def summary(groups, plan, b_assume=None):
    """(増収額, 平均増収率%, 経費回収率%) を返す。経費回収率は49.5%からの比例。"""
    t, r, cur = solve(groups, plan, b_assume)
    return t - cur, (r - 1) * 100, 49.5 * r


def solve_rate(groups, plan_fn, target, lo, hi, b_assume=None):
    """plan_fn(x) が返す料金表で平均増収率が target に最も近い整数 x を探す。"""
    best = None
    for x in range(lo, hi + 1):
        _, r, _ = solve(groups, plan_fn(x), b_assume)
        g = abs(r - 1 - target)
        if best is None or g < best[0]:
            best = (g, x)
    return best[1]


def bands(recs, plan, edges):
    """月使用量帯ごとの件数・現行額・改定後額。edges は上限のリスト。"""
    out = {e: [0, 0.0, 0.0] for e in edges}
    out['基本水量内'] = [0, 0.0, 0.0]
    out['日割'] = [0, 0.0, 0.0]
    out['異動等'] = [0, 0.0, 0.0]
    return out


def main():
    RK, RG = RC.load()
    groups = [RK, RG]
    bad = [v for v in range(0, 400) if abs(net_month(v / 2, CURRENT) - RC.month_net(v / 2)) > 1e-9]
    print('現行料金表が recompute と一致 :', '一致' if not bad else '不一致 %s' % bad[:5])
    t, r, cur = solve(groups, CURRENT)
    print('現行料金表での再現           : %s 円（差 %+.0f 円）' % (format(round(t), ','), t - cur))
    B = (1008, 5, ((10, 73), (50, 191), (None, 220)))
    inc, rate, rec = summary(groups, B)
    print('B案の再現                    : 増収 %s 円（%+.2f%%）・経費回収率 %.1f%%'
          % (format(round(inc), ','), rate, rec))


if __name__ == '__main__':
    main()
