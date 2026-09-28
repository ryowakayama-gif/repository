# -*- coding: utf-8 -*-
"""北塩原村 第10期　需要に対する供給（担い手の枠）の試算

【なぜ算定するか】
  国の基本指針は、介護人材の必要数の見通しを踏まえた計画の策定を求めている。
  本村は介護サービス事業所の従事者数の実績を持たないため、必要数を直接には
  算定できない。そこで、担い手の規模が生産年齢人口に比例するという仮定のもとで
  「令和7年の担い手規模で支えられる量」を置き、需要（認定者数）と突き合わせる。

【方法】（他案件で用いられていた考え方を本村のデータに当てたもの）
  1. 基準年（令和7年）の認定者数を、その年の担い手が支えている量とみなす
  2. 各年の供給の枠 ＝ 基準年の量 × （各年の生産年齢人口 ÷ 基準年の生産年齢人口）
  3. 重度から順に充てるものとして、要介護1以上の需要を差し引く
  4. 残りが「要支援の方に回せる量」となる

【限界】
  ・担い手の規模が生産年齢人口に比例するという仮定による
  ・本村は給付費の58.7％が村外の事業所によるものであり、その分の担い手は
    他市町村の生産年齢人口に依存する。ここでは同じ率で細ると置いている
  ・生産性向上・テクノロジーの導入・広域での確保は織り込んでいない
  ・重度から順に充てるという置き方であり、実際の配分を示すものではない
"""
import csv
import os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, 'data')

# 生産年齢人口（15〜64歳）。素案2-2・5-11の将来人口による
SEISAN = {'令和7年': 1192, '令和12年': 1008, '令和17年': 888, '令和22年': 770}

# 認定者数（B案の推計。素案5-3）。(要支援1, 要支援2, 計)
NINTEI = {
    '令和7年':  (50, 29, 208),
    '令和12年': (53, 30, 220),
    '令和17年': (51, 29, 212),
    '令和22年': (46, 26, 193),
}

BASE_YEAR = '令和7年'


def run():
    zen, kou, tot = NINTEI[BASE_YEAR]
    base_pop = SEISAN[BASE_YEAR]
    rows = []
    for y in ('令和7年', '令和12年', '令和17年', '令和22年'):
        s1, s2, total = NINTEI[y]
        yousien = s1 + s2                      # 要支援1・2の需要
        chuudo = total - yousien               # 要介護1以上の需要
        waku = tot * SEISAN[y] / base_pop      # 供給の枠
        nokori = waku - chuudo                 # 要支援に回せる量
        rows.append((y, SEISAN[y], total, yousien, chuudo,
                     round(waku), round(nokori),
                     round(nokori / yousien * 100, 1) if yousien else None))
    return rows


def main():
    rows = run()
    w = 12
    print('■ 需要に対する供給（担い手の枠）')
    print(f'  基準年 {BASE_YEAR}：認定者数 {NINTEI[BASE_YEAR][2]}人／生産年齢人口 {SEISAN[BASE_YEAR]}人')
    print()
    print(f'  {"年":<8}{"生産年齢":>{w}}{"認定者数":>{w}}{"要支援":>{w}}'
          f'{"要介護1以上":>{w}}{"供給の枠":>{w}}{"要支援に回る":>{w}}{"充足率":>{w}}')
    for y, pop, total, yousien, chuudo, waku, nokori, ratio in rows:
        print(f'  {y:<8}{pop:>{w},}{total:>{w},}{yousien:>{w},}'
              f'{chuudo:>{w},}{waku:>{w},}{nokori:>{w},}'
              + (f'{ratio:>{w-1}.1f}%' if ratio is not None else f'{"―":>{w}}'))
    print()
    base = rows[0]
    last = rows[-1]
    print(f'  要支援の方に回せる量は、{BASE_YEAR}の{base[6]}人分から'
          f'令和22年の{last[6]}人分へ{base[6] - last[6]}人分減ります。')
    print('  供給の枠が縮む一方で要介護1以上の需要はほぼ横ばいであるため、'
          '軽度の方に回る量が先に細ります。')

    # 受給者1人を支える生産年齢人口
    print()
    print('■ 受給者1人を支える生産年齢人口（参考）')
    JUKYU = {'令和7年': 168.5, '令和12年': 175.8, '令和17年': 169.4, '令和22年': 154.2}
    for y in ('令和7年', '令和12年', '令和17年', '令和22年'):
        print(f'  {y:<8}受給者{JUKYU[y]:>7.1f}人／月　生産年齢{SEISAN[y]:>6,}人　'
              f'1人あたり {SEISAN[y] / JUKYU[y]:.2f}人')

    path = os.path.join(DATA, '第10期_需要に対する供給.csv')
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        wr = csv.writer(f)
        wr.writerow(['年', '生産年齢人口', '認定者数', '要支援1・2', '要介護1以上',
                     '供給の枠', '要支援に回せる量', '要支援の充足率(%)'])
        for r in rows:
            wr.writerow(r)
        wr.writerow([])
        wr.writerow(['基準年', BASE_YEAR])
        wr.writerow(['受給者1人あたり生産年齢人口'])
        for y in ('令和7年', '令和12年', '令和17年', '令和22年'):
            wr.writerow([y, JUKYU[y], SEISAN[y], round(SEISAN[y] / JUKYU[y], 2)])
    print(f'\n保存: {os.path.relpath(path, BASE)}')


if __name__ == '__main__':
    main()
