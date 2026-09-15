# -*- coding: utf-8 -*-
"""介護給付費と予防給付費の分離

【背景】
  計画の様式は介護給付費と予防給付費を分けて記載することを求めている。
  「見える化」システムのサービス種類別の給付費（D17系）は両者を合算した値
  のみであるため、素案5-6ではこの2行を空欄としていた。
  しかし、決算（D48-b）には「介護サービス等諸費」と「介護予防サービス等諸費」が
  区分して収録されており、受給者数は要介護度別に把握できる。
  この2つを突き合わせれば、要支援・要介護それぞれの1人あたり給付月額を復元でき、
  第10期の見込量に当てはめて分離することができる。

【手順】
  ① 決算の予防給付費 ÷（要支援1・2の受給者数×12）＝ 予防給付の1人あたり給付月額
  ② 決算の介護給付費 ÷（要介護1〜5の受給者数×12）＝ 介護給付の1人あたり給付月額
  ③ ①②を3か年平均で固定し、5-3の認定者数推計に基づく受給者数の見込みに乗じる
  ④ ③で得た予防給付費の構成比を、5-4のサービス諸費に適用して分離する
     （合計を5-4・5-6と一致させるため、比だけを用いる）

【検証】
  ③の方法で令和3〜5年度の給付費を復元し、決算と比べる。
"""
import csv, os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, 'data')

YOSHIEN = ['要支援1', '要支援2']
YOKAIGO = ['要介護1', '要介護2', '要介護3', '要介護4', '要介護5']
KUBUN = [('D4', '在宅サービス'), ('D3', '居住系サービス'), ('D2', '施設サービス')]

# 決算の年度ラベル（D48-bは「◯年3月末」表記＝前年度の決算）
KESSAN_YEAR = {'令和4年3月末': '2021', '令和5年3月末': '2022', '令和6年3月末': '2023'}
YLAB = {'2021': '令和3年度', '2022': '令和4年度', '2023': '令和5年度'}
ACT = ['2021', '2022', '2023']

# 5-4のサービス諸費の見込み（estimate_services.py／素案5-6）
SVC_EST = {'令和9年度': 263_470, '令和10年度': 263_505, '令和11年度': 262_488,
           '令和12年度': 262_488, '令和17年度': 240_893, '令和22年度': 220_004}
EST = list(SVC_EST)


def load_mieruka():
    idx = {}
    with open(os.path.join(DATA, 'mieruka_tidy.csv'), encoding='utf-8-sig') as f:
        for r in csv.DictReader(f):
            if r['region'] != '北塩原村':
                continue
            idx.setdefault((r['code'], r['indicator'], r['year']), r['value'])
            idx.setdefault((r['code'], r['indicator'], '期間:' + r['period']), r['value'])
    return idx


def fnum(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def recipients(idx, year, degrees):
    """年度延べ受給者数の合計（在宅＋居住系＋施設）を12で割った月平均"""
    tot = 0.0
    for code, knd in KUBUN:
        for d in degrees:
            v = fnum(idx.get((code, f'受給者数（{knd}）（{d}）', year)))
            if v:
                tot += v
    return tot / 12.0


def kessan(idx, label, item):
    return fnum(idx.get(('D48-b', item, '期間:' + label)))


def load_est():
    """区分別受給者数推計から、要支援・要介護の月平均受給者数を取り出す"""
    path = os.path.join(DATA, '第10期_サービス区分別受給者数推計.csv')
    pre = {y: 0.0 for y in EST}
    care = {y: 0.0 for y in EST}
    with open(path, encoding='utf-8-sig') as f:
        for r in csv.DictReader(f):
            d = r['要介護度']
            if d == '計':
                continue
            tgt = pre if d in YOSHIEN else care
            for y in EST:
                v = fnum(r.get(y))
                if v:
                    tgt[y] += v
    return pre, care


def main():
    idx = load_mieruka()
    out = []
    p = out.append

    # ── ①② 決算から1人あたり給付月額を復元 ──────────────
    p('■ 1　決算から復元した1人あたり給付月額')
    p(f"  {'年度':<8}{'予防給付費':>12}{'要支援':>9}{'予防単価':>11}"
      f"{'介護給付費':>14}{'要介護':>9}{'介護単価':>11}")
    pu, cu = [], []
    hist = {}
    for lab, y in KESSAN_YEAR.items():
        yobo = kessan(idx, lab, '介護予防サービス等諸費')
        kaigo = kessan(idx, lab, '介護サービス等諸費')
        np_, nc = recipients(idx, y, YOSHIEN), recipients(idx, y, YOKAIGO)
        if not (yobo and kaigo and np_ and nc):
            continue
        up, uc = yobo / (np_ * 12), kaigo / (nc * 12)
        pu.append(up); cu.append(uc)
        hist[y] = (yobo, kaigo, np_, nc)
        p(f"  {YLAB[y]:<8}{yobo/1000:>11,.0f}千{np_:>8.1f}人{up:>10,.0f}円"
          f"{kaigo/1000:>13,.0f}千{nc:>8.1f}人{uc:>10,.0f}円")
    up_avg, uc_avg = sum(pu) / len(pu), sum(cu) / len(cu)
    p('')
    p(f'  3か年平均　予防 {up_avg:,.0f}円／月　介護 {uc_avg:,.0f}円／月　比 1：{uc_avg/up_avg:.1f}')
    p('  ※決算は「見える化」システムのD48-b系により令和5年度まで収録。')
    p('　　単価の3か年平均は令和3〜5年度による（5-4の受給率は令和5〜7年度の3か年平均）。')

    # ── 検証 ─────────────────────────────
    p('')
    p('■ 2　復元した単価による令和3〜5年度の再現（決算との比較）')
    p(f"  {'年度':<8}{'区分':<6}{'復元':>12}{'決算':>12}{'差':>10}")
    for y in ACT:
        if y not in hist:
            continue
        yobo, kaigo, np_, nc = hist[y]
        rp, rc = np_ * up_avg * 12, nc * uc_avg * 12
        p(f"  {YLAB[y]:<8}{'予防':<6}{rp/1000:>11,.0f}千{yobo/1000:>11,.0f}千{(rp-yobo)/yobo*100:>9.1f}%")
        p(f"  {'':<8}{'介護':<6}{rc/1000:>11,.0f}千{kaigo/1000:>11,.0f}千{(rc-kaigo)/kaigo*100:>9.1f}%")

    # ── ③ 第10期の見込みに当てはめる ──────────────────
    pre, care = load_est()
    p('')
    p('■ 3　第10期の見込みへの当てはめ')
    p(f"  {'年度':<9}{'要支援':>8}{'要介護':>8}{'予防給付費':>12}{'介護給付費':>12}{'計':>12}{'予防の割合':>10}")
    ratio = {}
    for y in EST:
        rp = pre[y] * up_avg * 12 / 1000
        rc = care[y] * uc_avg * 12 / 1000
        ratio[y] = rp / (rp + rc)
        p(f"  {y:<9}{pre[y]:>7.1f}人{care[y]:>7.1f}人{rp:>11,.0f}千{rc:>11,.0f}千"
          f"{rp+rc:>11,.0f}千{ratio[y]*100:>9.2f}%")

    # ── ④ 5-4のサービス諸費に構成比を適用 ────────────────
    p('')
    p('■ 4　素案5-4のサービス諸費に構成比を適用（採用値）')
    p('  合計を5-4・5-6と一致させるため、③で得た比のみを用いる。')
    p(f"  {'年度':<9}{'サービス諸費':>13}{'予防給付費':>12}{'介護給付費':>12}{'③との差':>11}")
    rows = []
    for y in EST:
        svc = SVC_EST[y]
        rp = svc * ratio[y]
        rc = svc - rp
        raw = pre[y] * up_avg * 12 / 1000 + care[y] * uc_avg * 12 / 1000
        p(f"  {y:<9}{svc:>12,}千{rp:>11,.0f}千{rc:>11,.0f}千{(svc-raw)/raw*100:>10.1f}%")
        rows.append((y, svc, rp, rc, ratio[y]))
    tot3 = sum(r[1] for r in rows[:3])
    pre3 = sum(r[2] for r in rows[:3])
    p('')
    p(f'  第10期3か年計　サービス諸費 {tot3:,}千円　'
      f'予防給付費 {pre3:,.0f}千円　介護給付費 {tot3-pre3:,.0f}千円')
    p(f'  予防給付費の割合は令和5年度決算の3.67%から第10期は{pre3/tot3*100:.2f}%へ上がる。')
    p('  要支援1の認定者が6年間で3倍になったことが給付費の構成に現れている。')

    path = os.path.join(DATA, '第10期_介護給付費_予防給付費の分離.csv')
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['年度', 'サービス諸費(千円)', '介護給付費(千円)', '予防給付費(千円)',
                    '予防の割合(%)', '要支援受給者(人/月)', '要介護受給者(人/月)'])
        for y, svc, rp, rc, rt in rows:
            w.writerow([y, svc, round(rc), round(rp), round(rt * 100, 2),
                        round(pre[y], 1), round(care[y], 1)])
    print('\n'.join(out))
    print(f'\n保存: {os.path.relpath(path, BASE)}')


if __name__ == '__main__':
    main()
