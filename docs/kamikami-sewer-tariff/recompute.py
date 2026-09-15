#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""令和7年度 調定データの区分・水量・収入の確定モデル（v3）。

条例（別表・第17条・第18条第3項/第4項）に忠実に実装し、集計ブックの水量欄と
算定式を照合して各行の検針区分（毎月／隔月）を判定する。v2までは2か月の算定式
のみで照合していたため、毎月検針の記録が「特殊算定」に混入していた。
"""
import math, statistics, openpyxl

BOOK   = '90c4df88-__R7________.xlsx'     # 【R7】使用料集計ブック
DETAIL = '6ff71ec0-R8.3_.xlsx'            # 令和7年度3月分 調定簿明細表（公共）
DETAIL_G = '37ccc0e0-R7___.xlsx'          # 令和7年度 調定簿明細表（漁業集落排水・5月〜3月）
CURRENT = (1008, 37, 174, 200)            # 条例別表（税抜・1か月）

def month_net(v, b=1008, r1=37, r2=174, r3=200):
    """1か月・水量v㎥の税抜額（別表：5㎥まで基本／5超10まで／10超50まで／50超）"""
    a = b
    if v > 5:  a += r1 * (min(v, 10) - 5)
    if v > 10: a += r2 * (min(v, 50) - 10)
    if v > 50: a += r3 * (v - 50)
    return a

def bill(v, months, *rates):
    """第18条第3項（各月均等按分）＋第17条ただし書（1円未満切捨て・一括）"""
    return math.floor(months * month_net(v, *(rates or CURRENT)) * 1.1)

MAIN6 = [3, 5, 7, 9, 11, 13]              # 主検針月（5,7,9,11,1,3月）
ALL12 = list(range(2, 14))                # 4月〜3月

def hiwari_mean():
    """基本水量未満の日割額（第18条第4項）の実測平均。集計ブックは上限2,216円で
    評価しているため、明細から実額を採る。"""
    ws = openpyxl.load_workbook(DETAIL, data_only=True)['Sheet1']
    v = [r[6] for r in ws.iter_rows(values_only=True)
         if len(r) > 6 and r[1] == '階上公共'
         and isinstance(r[6], (int, float)) and r[6] < bill(5, 2)]
    return statistics.mean(v)

def parse(sheet, cols=ALL12, hiwari=None):
    """(種別, 月あたり水量, 月数, 件数, 現行額) に分解。
    種別 N=通常算定 / B=基本水量内 / SP=日割・異動等"""
    hiwari = hiwari_mean() if hiwari is None else hiwari
    wb = openpyxl.load_workbook(BOOK, data_only=True)
    out = []
    for r in wb[sheet].iter_rows(min_row=3, values_only=True):
        if str(r[0]).strip() == '合計':
            continue
        amt, vol = r[1], r[0]
        cnt = sum(r[i] for i in cols if isinstance(r[i], (int, float)))
        if not cnt:
            continue
        cnt = int(cnt)
        if not isinstance(amt, (int, float)):        # 「〜2,216」＝日割
            out.append(('SP', None, None, cnt, hiwari * cnt)); continue
        amt = int(amt)
        if isinstance(vol, (int, float)):
            vol = int(vol)
            # 水量欄の意味は行によって異なるため3通りを照合する
            for v, m in ((vol / 2, 2), (vol, 1), (vol, 2)):
                if bill(v, m) == amt:
                    out.append(('N', v, m, cnt, amt * cnt)); break
            else:
                out.append(('SP', None, None, cnt, amt * cnt))
            continue
        out.append(('B', None, 2, cnt, amt * cnt) if amt == bill(5, 2)
                   else ('SP', None, None, cnt, amt * cnt))
    return out

def decompose(recs):
    """区分別の収入（税込・未丸めの成分で按分）と従量課金対象水量"""
    inc = dict(base=0.0, low=0.0, mid=0.0, high=0.0, sp=0.0)
    vol = dict(low=0.0, mid=0.0, high=0.0)
    for k, v, m, c, cur in recs:
        if k == 'SP': inc['sp'] += cur; continue
        if k == 'B':  inc['base'] += 1008 * 2 * 1.1 * c; continue
        inc['base'] += 1008 * m * 1.1 * c
        for key, q, rate in (('low',  max(0.0, min(v, 10) - 5),  37),
                             ('mid',  max(0.0, min(v, 50) - 10), 174),
                             ('high', max(0.0, v - 50),          200)):
            vol[key] += q * m * c
            inc[key] += rate * 1.1 * q * m * c
    return inc, vol

def revenue(recs, plan, sp_ratio):
    return sum(bill(v, m, *plan) * c if k == 'N'
               else bill(5, 2, *plan) * c if k == 'B'
               else cur * sp_ratio
               for k, v, m, c, cur in recs)

def solve(groups, plan, iters=80):
    """日割・異動等は平均改定率で連動させる（不動点）"""
    base = sum(sum(x[4] for x in g) for g in groups)
    r = 1.0
    for _ in range(iters):
        r = sum(revenue(g, plan, r) for g in groups) / base
    return sum(revenue(g, plan, r) for g in groups), r, base


def gyo_hiwari():
    """漁集の日割（基本水量未満）の実額。集計ブックは日割を基本使用料2,217円に
    切り上げて基本料金帯の行に含めているため、明細から実額を採って補正する。"""
    wb = openpyxl.load_workbook(DETAIL_G, data_only=True)
    v = [r[6] for sn in wb.sheetnames for r in wb[sn].iter_rows(values_only=True)
         if len(r) > 6 and r[1] == '階上漁集排'
         and isinstance(r[6], (int, float)) and r[6] < bill(5, 2)]
    return len(v), sum(v)

def apply_gyo_hiwari(recs, n=None, amt=None):
    """基本料金帯から日割n件を抜き、実額amtの日割（SP）として計上し直す。"""
    if n is None:
        n, amt = gyo_hiwari()
    out, left = [], n
    for k, v, m, c, cur in recs:
        if k == 'B' and left:
            take = min(c, left); left -= take
            if c - take:
                out.append((k, v, m, c - take, cur * (c - take) / c))
        else:
            out.append((k, v, m, c, cur))
    out.append(('SP', None, None, n, float(amt)))
    return out

def load():
    """確定データ（公共・漁集）。漁集は日割の実額補正を適用する。"""
    return parse('R7公共'), apply_gyo_hiwari(parse('R7漁集'))
