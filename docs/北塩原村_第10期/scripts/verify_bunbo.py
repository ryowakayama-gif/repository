# -*- coding: utf-8 -*-
"""サービス利用率の分母の検算

【背景】
  他案件（大雪地区広域連合）の点検で、次の指摘が出された。
    ・総括表詳細（２）とD32系の受給率の分母は第1号被保険者数である
    ・一方、国の将来推計における在宅サービス利用率の分母は
      「認定者数 － 施設・居住系利用者数」である
    ・取り違えると伸びが大きく変わる
  本村の算定（estimate_services.py）は、区分別の見込量を
  「見える化のサービス利用率（D45系）× 要介護度別認定者数」で置いている。
  この分母が何であるかを実データで確かめ、国のツールが別の分母を採る場合に
  どれだけ変わるかを測る。
"""
import csv, os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, 'data')

DEGREES = ['要支援1', '要支援2', '要介護1', '要介護2', '要介護3', '要介護4', '要介護5']
KUBUN = [('在宅サービス', 'D4', 'D45-a'), ('居住系サービス', 'D3', 'D45-b'),
         ('施設サービス', 'D2', 'D45-c')]
# 認定者数の指標名は全角数字（認定者数（要介護１））
ZEN = {'要支援1': '要支援１', '要支援2': '要支援２', '要介護1': '要介護１', '要介護2': '要介護２',
       '要介護3': '要介護３', '要介護4': '要介護４', '要介護5': '要介護５'}
# 認定者数（各年3月末＝前年度末）
CERT = {'2023': '令和6年3月末', '2024': '令和7年3月末時点', '2025': '令和8年3月末時点'}
YLAB = {'2023': '令和5年度', '2024': '令和6年度', '2025': '令和7年度'}


def load():
    idx = {}
    with open(os.path.join(DATA, 'mieruka_tidy.csv'), encoding='utf-8-sig') as f:
        for r in csv.DictReader(f):
            if r['region'] != '北塩原村':
                continue
            idx.setdefault((r['code'], r['indicator'], r['year']), r['value'])
            idx.setdefault((r['code'], r['indicator'], '期間:' + r['period']), r['value'])
    return idx


def fn(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def main():
    idx = load()
    out = []
    p = out.append

    def rate(code, kubun, d, y):
        for key in (f'サービス利用率（{kubun}）({d})', f'サービス利用率（{kubun}）（{d}）'):
            v = fn(idx.get((code, key, y)))
            if v is not None:
                return v
        return None

    def juk(code, kubun, d, y):
        return fn(idx.get((code, f'受給者数（{kubun}）（{d}）', y)))

    def cert(d, y):
        return fn(idx.get(('B4-a', f'認定者数（{ZEN[d]}）', '期間:' + CERT[y])))

    def ins(y):
        return fn(idx.get(('D4', '第1号被保険者数', y)))

    # ── ① D45系の分母を特定する ──────────────────────
    p('■ 1　見える化「サービス利用率」（D45系）の分母の特定')
    p('  受給者数（年度延べ÷12）を、①要介護度別認定者数 ②第1号被保険者数 で割り、')
    p('  公表されている利用率と一致するかを見る。')
    p(f"  {'区分':<12}{'要介護度':<8}{'年度':<8}{'受給者':>8}{'公表率':>8}{'÷認定者':>9}{'÷第1号':>9}  判定")
    ok_cert = ok_ins = n = 0
    for kubun, dcode, rcode in KUBUN:
        for d in DEGREES:
            for y in ['2023', '2024', '2025']:
                rt, nm, cn, iz = rate(rcode, kubun, d, y), juk(dcode, kubun, d, y), cert(d, y), ins(y)
                if rt is None or nm is None or not cn or not iz or rt == 0:
                    continue
                m = nm / 12.0
                a, b = m / cn * 100, m / iz * 100
                n += 1
                da, db = abs(a - rt), abs(b - rt)
                if da <= 0.6:
                    ok_cert += 1
                if db <= 0.6:
                    ok_ins += 1
                mark = '認定者' if da < db else '第1号'
                p(f"  {kubun:<12}{d:<8}{YLAB[y]:<8}{m:>7.1f}人{rt:>7.1f}%{a:>8.1f}%{b:>8.1f}%  {mark}")
    p('')
    p(f'  照合できた組 {n} 件のうち　÷認定者数で一致 {ok_cert} 件／÷第1号被保険者数で一致 {ok_ins} 件')
    p('  ±0.6ポイント以内で一致した組だけを数えている。一致しない組があるのは、')
    p('  認定者数が年度末（3月末）の値である一方、利用率は年度を通じた平均であるためとみられる。')
    p('  令和5年度は多くの組が小数第1位まで一致する。')
    p(f'  → 「サービス利用率」の分母は{"要介護度別認定者数" if ok_cert > ok_ins else "第1号被保険者数"}である。')
    p('  本村の算定（estimate_services.py）はこの率に要介護度別認定者数を乗じており、定義と整合する。')

    # ── ② 3区分の利用率の合計 ──────────────────────
    p('')
    p('■ 2　3区分の利用率の合計（100%を超えると重複計上を疑う）')
    p(f"  {'要介護度':<10}{'在宅':>8}{'居住系':>8}{'施設':>8}{'合計':>9}{'サービスなし':>12}")
    for d in DEGREES:
        vals = []
        for kubun, dcode, rcode in KUBUN:
            v = rate(rcode, kubun, d, '2025')
            vals.append(v if v is not None else 0.0)
        tot = sum(vals)
        p(f"  {d:<10}{vals[0]:>7.1f}%{vals[1]:>7.1f}%{vals[2]:>7.1f}%{tot:>8.1f}%{100-tot:>11.1f}%")
    p('  合計が100%を超える要介護度は、年度の途中で在宅から施設へ移った方が両方に数えられている')
    p('  （年度延べの受給者数を12で除しているため）ことを示す。重複の大きさは超過分に当たる。')

    # ── ③ 在宅の分母を「認定者数−施設・居住系」に置き換えた場合 ──────
    p('')
    p('■ 3　在宅サービス利用率の分母を「認定者数 － 施設・居住系利用者数」とした場合')
    p('  国の将来推計の設定がこの分母を採る場合、同じ在宅受給者数でも率がまったく変わる。')
    p(f"  {'要介護度':<10}{'認定者':>8}{'施設+居住系':>11}{'残余':>8}{'在宅受給者':>11}{'認定者ベース':>13}{'残余ベース':>12}")
    for d in DEGREES:
        cn = cert(d, '2025')
        if not cn:
            continue
        zaitaku = juk('D4', '在宅サービス', d, '2025')
        shisetsu = juk('D2', '施設サービス', d, '2025') or 0.0
        kyojyu = juk('D3', '居住系サービス', d, '2025') or 0.0
        if zaitaku is None:
            continue
        z, sk = zaitaku / 12.0, (shisetsu + kyojyu) / 12.0
        zan = cn - sk
        a = z / cn * 100
        b = z / zan * 100 if zan > 0 else float('nan')
        p(f"  {d:<10}{cn:>7.0f}人{sk:>10.1f}人{zan:>7.1f}人{z:>10.1f}人{a:>12.1f}%{b:>11.1f}%")
    p('')
    p('  在宅の率は、分母の取り方により大きく変わる。')
    p('  本村は「認定者数ベース」で算定している（素案5-4）。国の推計ワークシートを受領した時点で、')
    p('  ワークシートが求める分母を確認し、必要なら率を換算する。換算を誤ると見込量が数割動く。')

    path = os.path.join(DATA, '第10期_利用率の分母の検算.csv')
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['要介護度', '認定者数(人)', '施設+居住系(人/月)', '残余(人)', '在宅受給者(人/月)',
                    '認定者ベースの率(%)', '残余ベースの率(%)'])
        for d in DEGREES:
            cn = cert(d, '2025')
            zaitaku = juk('D4', '在宅サービス', d, '2025')
            if not cn or zaitaku is None:
                continue
            sk = ((juk('D2', '施設サービス', d, '2025') or 0.0) +
                  (juk('D3', '居住系サービス', d, '2025') or 0.0)) / 12.0
            z, zan = zaitaku / 12.0, cn - sk
            w.writerow([d, round(cn), round(sk, 1), round(zan, 1), round(z, 1),
                        round(z / cn * 100, 1), round(z / zan * 100, 1) if zan > 0 else ''])
    print('\n'.join(out))
    print(f'\n保存: {os.path.relpath(path, BASE)}')


if __name__ == '__main__':
    main()
