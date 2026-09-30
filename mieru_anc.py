# -*- coding: utf-8 -*-
"""見える化「総人口と被保険者数の設定」へ入れる案Cの値（男女別6階級）。

案Cは65歳以上を65〜74歳・75〜84歳・85歳以上の3区分で推計している。
画面は6階級×男女を求めるため、3区分の合計を保ったまま、
画面が現に表示している補正データの同じ年の構成比で分ける
（階級ごとに合計を保つ最大剰余法）。

令和6年度・令和7年度は実績であるため画面の値をそのまま用いる。

本モジュールは受領点検と入力値の一覧の双方から読む。
同じ配分を2か所に書かないために切り出したものである。
"""

import os
import runpy
import sys
from decimal import Decimal, ROUND_HALF_UP

import data_mieru_jinko as J
import repo_paths as RP

JITSU = ("R6", "R7")                     # 実績年。画面の値をそのまま用いる
PAIR3 = {"65-74": ["65-69", "70-74"], "75-84": ["75-79", "80-84"],
         "85+": ["85-89", "90+"]}


def r0(x):
    return int(Decimal(str(x)).quantize(Decimal("1"), ROUND_HALF_UP))


def saidai(total, w):
    """最大剰余法。w の比で total（整数）に配分し、合計を保つ。"""
    s = sum(w)
    raw = [total * x / s for x in w]
    base = [int(x) for x in raw]
    order = sorted(range(len(w)), key=lambda i: -(raw[i] - base[i]))
    for j in range(total - sum(base)):
        base[order[j % len(order)]] += 1
    return base


def _projection():
    old = sys.stdout
    try:
        sys.stdout = open(os.devnull, "w")
        return runpy.run_path(os.path.join(RP.ROOT, "build_projection.py"))
    finally:
        try:
            sys.stdout.close()
        except Exception:
            pass
        sys.stdout = old


_P = _projection()
pop_tot, pop_juki, CL3 = _P["pop_tot"], _P["pop_juki"], _P["CL"]


def build():
    """{年度: {総人口, 男{階級}, 女{階級}, 出所}} を返す。"""
    anc = {}
    for y in J.YEARS:
        if y in JITSU:
            anc[y] = {"総人口": J.SOJINKO[y],
                      "男": {a: J.dan(a, y) for a in J.AGE},
                      "女": {a: J.jo(a, y) for a in J.AGE},
                      "出所": "画面（実績）"}
            continue
        s = J.SEIREKI[y]
        c3 = {c: pop_juki(c, s) for c in CL3}
        t65 = r0(sum(c3.values()))
        a6 = {}
        for c in CL3:
            x, z = PAIR3[c]
            w = [J.kei(x, y), J.kei(z, y)]
            a6[x] = c3[c] * w[0] / sum(w)
            a6[z] = c3[c] * w[1] / sum(w)
        ints = dict(zip(J.AGE, saidai(t65, [a6[a] for a in J.AGE])))
        m, f = {}, {}
        for a in J.AGE:
            m[a], f[a] = saidai(ints[a], [J.dan(a, y), J.jo(a, y)])
        anc[y] = {"総人口": r0(pop_tot(s)), "男": m, "女": f, "出所": "案C"}
    return anc


ANC = build()


def ichigo(y):
    """第1号被保険者数（画面が自動計算する値）。"""
    return sum(ANC[y]["男"].values()) + sum(ANC[y]["女"].values())


def kouki(y):
    """75歳以上（後期高齢者）。"""
    return sum(ANC[y]["男"][a] + ANC[y]["女"][a] for a in J.AGE[2:])
