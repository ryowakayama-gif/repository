# -*- coding: utf-8 -*-
"""中長期（令和17年度・令和22年度）の保険料基準額の見通し

【背景】
  社会福祉法等の一部を改正する法律（令和8年法律第51号）により、
  中長期的なサービスの見通しを計画に定めることが必須となった。
  素案5-11は見込量・給付費まで算定済みで、保険料のみが空欄であった。

【算定の構造】（estimate_premium.py の5式を単年度に当てはめる）
  ① 標準給付費見込額 ＝ サービス諸費 ＋ 補助給付等
  ② 第1号被保険者負担分相当額 ＝（① ＋ 地域支援事業費）× α
  ③ 調整交付金相当額 ＝（① ＋ 総合事業費）× 5％
  ④ 保険料収納必要額 ＝ ② ＋ ③ － 調整交付金見込額
       見込交付割合 ＝（α ＋ 0.05）－ α × F × G
       ⇒ ④ ＝ α ×〔（① ＋ 地域支援事業費）＋（① ＋ 総合事業費）×（F×G－1）〕
  ⑤ 保険料基準額 ＝ ④ ÷ 予定収納率 ÷（被保険者数 × G）÷ 12

【延ばす基礎】
  補助給付等　　 令和11年度は素案5-6と同じ令和5年度決算額を置き、
                 令和17・22年度はサービス諸費に比例させる
  地域支援事業費 第1号被保険者数に比例
  総合事業費　　 第1号被保険者数に比例
  いずれも令和5年度決算を基準とする。

【留意】
  報酬改定・制度改正・供給制約を織り込まない参考値である。
  αは3年ごとに政令で定められるため、据え置いた場合と傾向が続いた場合の2案を示す。
"""
import csv, os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, 'data')

# ── 令和5年度決算（介護保険特別会計）──────────────────
SVC_R5 = 258_848_000                     # サービス諸費
HOJO_R5 = 14_495_976 + 6_768_762 + 0 + 252_189      # 補助給付等 21,516,927
CHIIKI_R5 = 39_885_993                   # 地域支援事業費
SOGO_R5 = 14_637_340 + 1_825_335         # 総合事業費 16,462,675
INS_R5 = 1_004                           # 第1号被保険者数（令和6年3月末）

# ── 素案5-11の見通し（estimate_services.py・estimate_certified.py による）──
YEARS = ['令和11年度', '令和17年度', '令和22年度']
SVC = {'令和11年度': 262_488_000, '令和17年度': 240_893_000, '令和22年度': 220_004_000}
INS = {'令和11年度': 987, '令和17年度': 891, '令和22年度': 814}

# ── 保険料の係数（第9期計画から復元。estimate_premium.py と同じ）──
FG = 1.03428          # 後期高齢者加入割合補正係数 × 所得段階別加入割合補正係数
G = 0.95928           # 所得段階別加入割合補正係数
SHUNO = 0.995         # 予定保険料収納率

# ── αの2案 ────────────────────────────
#   ケース1：第9期の23％を据え置く
#   ケース2：第1期17％から第9期23％まで3年ごとに1ポイント上がってきた傾向が続く
#            第10期（令和9〜11年度）24％／第12期（令和15〜17年度）26％／
#            第14期（令和21〜23年度）28％
ALPHA = {'据え置き': {'令和11年度': 0.23, '令和17年度': 0.23, '令和22年度': 0.23},
         '傾向が続く': {'令和11年度': 0.24, '令和17年度': 0.26, '令和22年度': 0.28}}


def parts(y):
    """年度yの ①・地域支援事業費・総合事業費・分母 を返す"""
    svc = SVC[y]
    # 補助給付等は素案5-6と同じく令和5年度決算を令和11年度に置き、
    # 中長期はサービス諸費に比例させる（5-6の令和11年度と一致させるため）
    hojo = HOJO_R5 * svc / SVC['令和11年度']
    std = svc + hojo                                   # ①
    chiiki = CHIIKI_R5 * INS[y] / INS_R5
    sogo = SOGO_R5 * INS[y] / INS_R5
    denom = INS[y] * G * 12 * SHUNO
    return std, chiiki, sogo, denom


def premium(y, alpha):
    std, chiiki, sogo, denom = parts(y)
    need = alpha * ((std + chiiki) + (std + sogo) * (FG - 1))     # ④
    return need, need / denom


def main():
    out = []
    p = out.append

    p('■ 算定の基礎（単位：千円）')
    p(f"  {'区分':<22}" + ''.join(f'{y:>14}' for y in YEARS))
    rows = {}
    for y in YEARS:
        rows[y] = parts(y)
    for i, lab in enumerate(['サービス諸費', '　補助給付等', '①標準給付費見込額',
                             '地域支援事業費', '　うち総合事業費']):
        if lab == 'サービス諸費':
            vals = [SVC[y] for y in YEARS]
        elif lab == '　補助給付等':
            vals = [rows[y][0] - SVC[y] for y in YEARS]
        elif lab == '①標準給付費見込額':
            vals = [rows[y][0] for y in YEARS]
        elif lab == '地域支援事業費':
            vals = [rows[y][1] for y in YEARS]
        else:
            vals = [rows[y][2] for y in YEARS]
        p(f"  {lab:<22}" + ''.join(f'{v/1000:>13,.0f}' for v in vals))
    p(f"  {'第1号被保険者数':<22}" + ''.join(f'{INS[y]:>12,}人' for y in YEARS))
    p(f"  {'分母（人・月）':<22}" + ''.join(f'{rows[y][3]:>13,.0f}' for y in YEARS))

    res = {}
    for case, am in ALPHA.items():
        p('')
        p(f'■ ケース：αを{case}')
        p(f"  {'区分':<22}" + ''.join(f'{y:>14}' for y in YEARS))
        p(f"  {'第1号被保険者負担割合':<22}" + ''.join(f'{am[y]*100:>12.0f}%' for y in YEARS))
        needs, prems = [], []
        for y in YEARS:
            n, pr = premium(y, am[y])
            needs.append(n); prems.append(pr)
        p(f"  {'④保険料収納必要額':<22}" + ''.join(f'{n/1000:>13,.0f}' for n in needs))
        p(f"  {'⑤保険料基準額（月額）':<22}" + ''.join(f'{pr:>12,.0f}円' for pr in prems))
        res[case] = prems
        p(f'  令和11年度→令和22年度　{prems[0]:,.0f}円 → {prems[2]:,.0f}円'
          f'（{(prems[2]/prems[0]-1)*100:+.1f}%）')

    p('')
    p('■ 読み取り')
    a = res['据え置き']
    b = res['傾向が続く']
    p(f'  αを据え置くと、サービス諸費が16.2%減っても保険料基準額は'
      f'{a[0]:,.0f}円→{a[2]:,.0f}円（{(a[2]/a[0]-1)*100:+.1f}%）でほぼ動かない。')
    p(f'  給付費の減少と被保険者数の減少がほぼ打ち消し合うためである。')
    p(f'  αが3年ごとに1ポイント上がると{b[0]:,.0f}円→{b[2]:,.0f}円'
      f'（{(b[2]/b[0]-1)*100:+.1f}%）となり、令和22年度の差は{b[2]-a[2]:,.0f}円。')
    p(f'  中長期の保険料を決めるのは給付費ではなくαである。')
    p('  参考：第9期計画は令和22年度の基準月額を8,700円と見込んでいた。')

    print('\n'.join(out))

    path = os.path.join(DATA, '第10期_中長期保険料見通し.csv')
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['区分'] + YEARS)
        w.writerow(['サービス諸費(千円)'] + [round(SVC[y] / 1000) for y in YEARS])
        w.writerow(['標準給付費見込額(千円)'] + [round(rows[y][0] / 1000) for y in YEARS])
        w.writerow(['地域支援事業費(千円)'] + [round(rows[y][1] / 1000) for y in YEARS])
        w.writerow(['第1号被保険者数(人)'] + [INS[y] for y in YEARS])
        for case, am in ALPHA.items():
            w.writerow([f'α（{case}）'] + [f'{am[y]*100:.0f}%' for y in YEARS])
            w.writerow([f'保険料基準額（{case}・円）'] + [round(v) for v in res[case]])
    print(f'\n保存: {os.path.relpath(path, BASE)}')


if __name__ == '__main__':
    main()
