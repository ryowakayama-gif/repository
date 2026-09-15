# -*- coding: utf-8 -*-
"""単価の趨勢と、推計の不確かさが保険料に及ぼす影響

【背景】
  素案5-4の見込量は、受給率・1人あたり利用量・1人あたり給付月額（単価）を
  いずれも令和5〜7年度の3か年平均で固定している。これは計画の標準的な立て方だが、
  ・単価が趨勢的に上がっていれば給付費は過小になる
  ・サービス別の伸びの偏りは平均に吸収される
  ・推計そのものに誤差がある
  という3つの影響が残る。本稿はその大きさを実測する。

【単価指数】
  サービスごとの単価の変化を、前年度の給付費構成比で加重して1本の指数にまとめる
  （連鎖ラスパイレス型）。
      P(t)/P(t-1) ＝ Σ w(s,t-1) × p(s,t)/p(s,t-1)
      w(s,t-1) ＝ q(s,t-1)·p(s,t-1) ÷ Σ q·p
  加重をとることで、受給者が1〜2人のサービスの単価の振れが指数を動かさないようにする。

【保険料への換算】
  限界率23.79％、分母34,201人・月（estimate_premium.py）による。
  標準給付費見込額854,013,781円の1％＝8,540,138円＝月額59円。
"""
import csv, os, statistics

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, 'data')

# 素案5-4(1)〜(3)で見込量を立てる13サービス（D32＝受給者、D17＝単価）
SVC = [('訪問介護', 'D32-a', 'D17-a'), ('訪問入浴介護', 'D32-b', 'D17-b'),
       ('訪問看護', 'D32-c', 'D17-c'), ('訪問リハビリテーション', 'D32-d', 'D17-d'),
       ('居宅療養管理指導', 'D32-e', 'D17-e'), ('通所介護', 'D32-f', 'D17-f'),
       ('通所リハビリテーション', 'D32-g', 'D17-g'), ('短期入所生活介護', 'D32-h', 'D17-h'),
       ('短期入所療養介護', 'D32-i', 'D17-i'), ('福祉用具貸与', 'D32-j', 'D17-j'),
       ('特定施設入居者生活介護', 'D32-k', 'D17-k'),
       ('認知症対応型共同生活介護', 'D32-q', 'D17-q'),
       ('地域密着型通所介護', 'D32-s', 'D17-t')]

YEARS = [str(y) for y in range(2014, 2026)]      # H26〜R7
YLAB = {'2014': 'H26', '2015': 'H27', '2016': 'H28', '2017': 'H29', '2018': 'H30',
        '2019': 'R元', '2020': 'R2', '2021': 'R3', '2022': 'R4', '2023': 'R5',
        '2024': 'R6', '2025': 'R7'}

# 保険料への換算（estimate_premium.py による）
YEN_PER_PCT = 59.0          # 標準給付費が1％動いたときの保険料基準額（月額・円）


def load():
    idx = {}
    with open(os.path.join(DATA, 'mieruka_tidy.csv'), encoding='utf-8-sig') as f:
        for r in csv.DictReader(f):
            if r['region'] != '北塩原村':
                continue
            idx.setdefault((r['code'], r['indicator'], r['year']), r['value'])
    return idx


def num(idx, c, i, y):
    try:
        return float(idx.get((c, i, y)))
    except (TypeError, ValueError):
        return None


def rate(idx, code, y, key='合計'):
    for (c, i, yy), v in idx.items():
        if c != code or yy != y:
            continue
        n = i.replace('（', '(').replace('）', ')').replace(' ', '')
        if n == key or n.endswith('(' + key + ')'):
            try:
                return float(v)
            except (TypeError, ValueError):
                return None
    return None


def juk(idx, code, y):
    """年度延べ受給者数。系列がなければ受給率×第1号被保険者数×12で復元する。"""
    for (c, i, yy), v in idx.items():
        if c == code and yy == y and i.startswith('受給者数') and i.endswith('（合計）'):
            try:
                return float(v)
            except (TypeError, ValueError):
                return None
    rt = rate(idx, code, y)
    ins = num(idx, 'D4', '第1号被保険者数', y)
    return rt / 100 * ins * 12 if rt is not None and ins else None


def price(idx, code, name, y):
    return num(idx, code, f'受給者1人あたり給付月額（{name}）', y)


def main():
    idx = load()
    out = []
    p = out.append

    # ── 1. 連鎖ラスパイレス型の単価指数 ────────────────────
    p('■ 1　単価の趨勢（連鎖ラスパイレス型の指数）')
    p(f"  {'年度':<6}{'指数':>9}{'前年比':>9}{'採用ｻｰﾋﾞｽ数':>12}")
    index, chain = {YEARS[0]: 100.0}, []
    for prev, cur in zip(YEARS[:-1], YEARS[1:]):
        num_, den, n = 0.0, 0.0, 0
        for name, c32, c17 in SVC:
            q0, p0, p1 = juk(idx, c32, prev), price(idx, c17, name, prev), price(idx, c17, name, cur)
            if not (q0 and p0 and p1) or q0 <= 0 or p0 <= 0:
                continue
            w = q0 / 12 * p0                      # 前年度の月あたり給付費＝ウェイト
            num_ += w * (p1 / p0)
            den += w
            n += 1
        if den <= 0:
            index[cur] = index[prev]
            continue
        r = num_ / den
        index[cur] = index[prev] * r
        chain.append((cur, r - 1, n))
        p(f"  {YLAB[cur]:<6}{index[cur]:>9.1f}{(r-1)*100:>8.2f}%{n:>10}社")
    y0, y1 = YEARS[0], YEARS[-1]
    span = len(YEARS) - 1
    g_all = (index[y1] / index[y0]) ** (1 / span) - 1
    # 直近6年（R元→R7）は報酬改定2回（R3・R6）を含む
    g_6 = (index['2025'] / index['2019']) ** (1 / 6) - 1
    p('')
    p(f'  {YLAB[y0]}→{YLAB[y1]}（{span}年）の年率　　{g_all*100:+.2f}%　累積 {index[y1]-100:+.1f}%')
    p(f'  R元→R7（6年）の年率　　　{g_6*100:+.2f}%　累積 {index["2025"]/index["2019"]*100-100:+.1f}%')
    yr = [c[1] for c in chain]
    p(f'  年次の変化率　　最小 {min(yr)*100:+.2f}%　最大 {max(yr)*100:+.2f}%　中央値 {statistics.median(yr)*100:+.2f}%')
    p('  ※介護報酬改定は平成27年▲2.27%、平成30年＋0.54%、令和元年＋2.13%（消費税等）、')
    p('　　令和3年＋0.70%、令和6年＋1.59%。指数はこれに利用の構成の変化が加わった値。')

    # ── 2. 3か年平均で固定することの影響 ──────────────────
    p('')
    p('■ 2　単価を令和5〜7年度の3か年平均で固定することの影響')
    p('  3か年平均は中央の令和6年度の水準にほぼ等しい。計画期間の各年度は令和6年度から')
    p('  3年・4年・5年後であるため、単価が趨勢どおり上がるとすれば給付費は次のとおり過小になる。')
    p(f"  {'年度':<8}{'経過年数':>9}{'過小率':>10}{'月額換算':>10}")
    tot = 0.0
    for lab, k in [('令和9年度', 3), ('令和10年度', 4), ('令和11年度', 5)]:
        u = (1 + g_6) ** k - 1
        tot += u
        p(f"  {lab:<8}{k:>8}年{u*100:>9.2f}%{u*100*YEN_PER_PCT:>9.0f}円")
    avg = tot / 3
    p(f"  {'3か年平均':<8}{'':>9}{avg*100:>9.2f}%{avg*100*YEN_PER_PCT:>9.0f}円")
    p(f'  → 直近6年の年率{g_6*100:+.2f}%が続く場合、保険料基準額は月額約{avg*100*YEN_PER_PCT:.0f}円の過小となる。')
    p(f'  → 全期間の年率{g_all*100:+.2f}%を用いる場合は月額約{((1+g_all)**4-1)*100*YEN_PER_PCT:.0f}円。')

    # ── 3. サービス別の伸びの偏り ────────────────────
    p('')
    p('■ 3　サービス別の受給者数の伸びの偏り（令和2→令和7年度・5年）')
    p(f"  {'サービス':<24}{'R2':>8}{'R7':>8}{'年率':>9}")
    gs = []
    for name, c32, c17 in SVC:
        q0, q1 = juk(idx, c32, '2020'), juk(idx, c32, '2025')
        if not (q0 and q1) or q0 <= 0:
            continue
        g = (q1 / q0) ** (1 / 5) - 1
        gs.append((name, g))
        p(f"  {name:<24}{q0/12:>7.1f}人{q1/12:>7.1f}人{g*100:>8.2f}%")
    n0, n1 = num(idx, 'B4-a', '合計認定者数', '2020'), num(idx, 'B4-a', '合計認定者数', '2025')
    gn = (n1 / n0) ** (1 / 5) - 1
    p(f"  {'（参考）認定者数':<24}{n0:>7.0f}人{n1:>7.0f}人{gn*100:>8.2f}%")
    lo, hi = min(g for _, g in gs), max(g for _, g in gs)
    p('')
    p(f'  サービス別の年率は {lo*100:+.2f}% から {hi*100:+.2f}% まで {(hi-lo)*100:.0f}ポイント開いている。')
    p(f'  受給率を3か年平均で固定することは、すべてのサービスを認定者数の年率{gn*100:+.2f}%で')
    p('  一律に延ばすことに等しく、この68ポイントの偏りは見込量に反映されない。')

    # ── 4. 総額の予測誤差 ──────────────────────
    p('')
    p('■ 4　一律に延ばす方法で3年先を予測したときの総額の誤差')
    p('  起点を変えて、11サービスの給付費の合計を3年先まで予測し、実績と比べる。')
    p(f"  {'起点':<6}{'予測年':<8}{'予測':>12}{'実績':>12}{'誤差':>9}")
    CASES = [('2019', '2022'), ('2020', '2023'), ('2021', '2024'), ('2022', '2025')]
    errs = []
    for base, target in CASES:
        win = str(int(base) - 5)
        a, b, c = (num(idx, 'B4-a', '合計認定者数', y) for y in (win, base, target))
        if not (a and b and c):
            continue
        g = (b / a) ** (1 / 5) - 1
        pred = act = 0.0
        for name, c32, c17 in SVC:
            qb, qt = juk(idx, c32, base), juk(idx, c32, target)
            pb = price(idx, c17, name, base)
            pt = price(idx, c17, name, target)
            if not (qb and qt and pb and pt):
                continue
            pred += qb / 12 * (1 + g) ** 3 * pb * 12      # 受給者は一律に延ばし、単価は据え置き
            act += qt / 12 * pt * 12
        if act <= 0:
            continue
        e = pred / act - 1
        errs.append(e)
        p(f"  {YLAB[base]:<6}{YLAB[target]:<8}{pred/1e6:>10.1f}百万{act/1e6:>10.1f}百万{e*100:>8.1f}%")
    mean = statistics.fmean(errs)
    mabs = statistics.fmean(abs(e) for e in errs)
    p('')
    p(f'  誤差の平均 {mean*100:+.1f}%　絶対値の平均 {mabs*100:.1f}%　幅 {min(errs)*100:+.1f}%〜{max(errs)*100:+.1f}%')
    p(f'  絶対値の平均を保険料に換算すると月額約 {mabs*100*YEN_PER_PCT:.0f}円。')
    p('  第9期計画が見込んだ基金取崩24,000千円の効き（月額702円）と同じ桁であり、')
    p('  推計の精度そのものが保険料の水準を左右することを示している。')

    print('\n'.join(out))

    path = os.path.join(DATA, '第10期_単価の趨勢と推計誤差.csv')
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['区分', '項目', '値'])
        w.writerow(['単価指数', f'{YLAB[y0]}→{YLAB[y1]}の年率(%)', round(g_all * 100, 2)])
        w.writerow(['単価指数', 'R元→R7の年率(%)', round(g_6 * 100, 2)])
        for y, r, n in chain:
            w.writerow(['単価指数', f'{YLAB[y]}の前年比(%)', round(r * 100, 2)])
        w.writerow(['3か年平均固定', '給付費の過小率(3か年平均・%)', round(avg * 100, 2)])
        w.writerow(['3か年平均固定', '保険料月額換算(円)', round(avg * 100 * YEN_PER_PCT)])
        for name, g in gs:
            w.writerow(['受給者数の年率', name + '(%)', round(g * 100, 2)])
        w.writerow(['受給者数の年率', '認定者数(%)', round(gn * 100, 2)])
        w.writerow(['総額の予測誤差', '平均(%)', round(mean * 100, 1)])
        w.writerow(['総額の予測誤差', '絶対値の平均(%)', round(mabs * 100, 1)])
        w.writerow(['総額の予測誤差', '保険料月額換算(円)', round(mabs * 100 * YEN_PER_PCT)])
    print(f'\n保存: {os.path.relpath(path, BASE)}')


if __name__ == '__main__':
    main()
