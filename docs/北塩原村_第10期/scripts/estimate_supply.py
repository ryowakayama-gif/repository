# -*- coding: utf-8 -*-
"""供給の制約（村内定員に対する到達率）の点検

【背景】
  他案件（大雪地区広域連合）の点検で、見込量に定員の上限を課すかどうかが
  論点として整理された。そこでは次の2点が確かめられている。
    ・定員到達率を年度別に算定し、超過が生じる年度を特定する
    ・介護保険事業状況報告（年報）は施設の所在地を問わず当該保険者の
      被保険者を数えるため、区域内の定員は上限ではない。
      到達率が100％を超えることは誤りではない

【本村の事情】
  村内にあるのは認知症対応型共同生活介護（定員27人）と通所介護（定員50人）の
  2種類だけで、施設サービスは村内にない。多くを村外の事業所で利用している。
  したがって「村内定員に対する到達率」は供給の上限ではなく、
  村内で受け止められている割合を示す指標として読む。
"""
import csv, os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, 'data')

# 村内の定員（平成30年度以降変わらず。素案2-5）
TEIIN = {'認知症対応型共同生活介護': 27, '通所介護': 50}
YEARS = ['令和9年度', '令和10年度', '令和11年度', '令和12年度', '令和17年度', '令和22年度']
ACT = ['令和5年度', '令和6年度', '令和7年度']


def fn(v):
    try:
        return float(str(v).replace(',', ''))
    except (TypeError, ValueError):
        return None


def main():
    rows = {}
    # 受給者数を持つサービス（居宅・地域密着型）
    with open(os.path.join(DATA, '第10期_サービス種類別見込量.csv'), encoding='utf-8-sig') as f:
        for r in csv.DictReader(f):
            rows[r['サービス種類']] = r
    # 給付費のみ把握できるサービス（居住系・施設）は「人/月・参考」の列名
    with open(os.path.join(DATA, '第10期_給付費のみ把握サービス.csv'), encoding='utf-8-sig') as f:
        for r in csv.DictReader(f):
            d = dict(r)
            for k in list(d):
                if k.endswith('(人/月・参考)'):
                    d[k.replace('(人/月・参考)', '(人/月)')] = d[k]
                if k.endswith('実績(人/月・参考)'):
                    d[k.replace('実績(人/月・参考)', '実績(人/月)')] = d[k]
            rows.setdefault(r['サービス種類'], d)

    out = []
    p = out.append
    p('■ 村内定員に対する到達率')
    p('  受給者数（人／月）÷ 村内定員。100％を超えても誤りではない（村外利用を含むため）。')
    hdr = f"  {'サービス':<26}{'定員':>6}" + ''.join(f'{y[:5]:>11}' for y in ACT + YEARS[:3])
    p(hdr)
    tbl = []
    for name, cap in TEIIN.items():
        r = rows.get(name)
        if not r:
            p(f'  {name}：見込量の行が見つかりません')
            continue
        vals = []
        for y in ACT:
            vals.append(fn(r.get(y + '実績(人/月)')))
        for y in YEARS[:3]:
            vals.append(fn(r.get(y + '(人/月)')))
        p(f"  {name:<26}{cap:>5}人" + ''.join(
            f'{v/cap*100:>10.1f}%' if v is not None else f"{'―':>11}" for v in vals))
        tbl.append((name, cap, vals))
    p('')
    p('■ 読み取り')
    for name, cap, vals in tbl:
        act = [v for v in vals[:3] if v is not None]
        est = [v for v in vals[3:] if v is not None]
        if not act or not est:
            continue
        p(f'  ・{name}（定員{cap}人）')
        p(f'　　令和7年度 {act[-1]:.1f}人／月＝到達率{act[-1]/cap*100:.1f}％'
          f'　→　令和11年度 {est[-1]:.1f}人／月＝到達率{est[-1]/cap*100:.1f}％')
        if max(est) > cap:
            p(f'　　計画期間に定員を超過します（最大{max(est):.1f}人／月）。')
        else:
            p(f'　　計画期間に定員の超過は生じません（最大{max(est):.1f}人／月・'
              f'到達率{max(est)/cap*100:.1f}％）。頭打ちを課しても見込量は動きません。')
    p('')
    p('  本村は施設サービスが村内になく、認知症対応型共同生活介護と通所介護の2種類のみが')
    p('  村内で提供されています。到達率は供給の上限ではなく、村内で受け止められている')
    p('  割合を示すものとして読みます。到達率が上がることは、村外への流出が減ることを意味します。')

    path = os.path.join(DATA, '第10期_村内定員の到達率.csv')
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['サービス', '村内定員'] + [y + '(人/月)' for y in ACT + YEARS[:3]]
                   + [y + 'の到達率(%)' for y in ACT + YEARS[:3]])
        for name, cap, vals in tbl:
            w.writerow([name, cap] + [('' if v is None else round(v, 1)) for v in vals]
                       + [('' if v is None else round(v / cap * 100, 1)) for v in vals])
    print('\n'.join(out))
    print(f'\n保存: {os.path.relpath(path, BASE)}')


if __name__ == '__main__':
    main()
