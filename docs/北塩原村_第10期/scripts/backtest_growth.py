# -*- coding: utf-8 -*-
"""サービス別の伸びの偏りを見込量にどれだけ反映するかの検証

【問題】
  本村の見込量は受給率を3か年平均で固定している。これは実質的に、
  すべてのサービスを認定者数の伸び（年+2.62%）で一律に延ばすことに等しい。
  しかしサービス別の実際の伸びは年率で −13.12％から+54.88％まで68ポイント開いている。
  一律で置くと、伸びているサービスは過小に、縮んでいるサービスは過大になる。

【補正の式】
    g'ₛ ＝ gₙ ＋ φ × clip(gₛ − gₙ, ±c)
      gₛ  サービスsの実測の年率
      gₙ  全体（認定者数）の年率
      φ   反映割合（0＝一律、1＝実測をそのまま）
      c   一律からの乖離の上限

【検証】
  起点年を変えて3年先を予測し、実績との誤差を比べる。
  受給者数が少ないサービスでは1人の増減が年率を大きく動かすため、
  誤差は給付費で重み付けした加重平均絶対誤差（WMAPE）で評価する。
"""
import csv, os, statistics

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, 'data')

SVC = [('訪問介護', 'D32-a', 'D17-a'), ('訪問入浴介護', 'D32-b', 'D17-b'),
       ('訪問看護', 'D32-c', 'D17-c'), ('訪問リハビリテーション', 'D32-d', 'D17-d'),
       ('居宅療養管理指導', 'D32-e', 'D17-e'), ('通所介護', 'D32-f', 'D17-f'),
       ('通所リハビリテーション', 'D32-g', 'D17-g'), ('短期入所生活介護', 'D32-h', 'D17-h'),
       ('短期入所療養介護', 'D32-i', 'D17-i'), ('福祉用具貸与', 'D32-j', 'D17-j'),
       ('地域密着型通所介護', 'D32-s', 'D17-t')]


def load():
    rows = [r for r in csv.DictReader(open(os.path.join(DATA, 'mieruka_tidy.csv'),
                                           encoding='utf-8-sig')) if r['region'] == '北塩原村']
    idx = {}
    for r in rows:
        idx.setdefault((r['code'], r['indicator'], r['year']), r['value'])
    return idx


def num(idx, c, i, y):
    v = idx.get((c, i, y))
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def rate(idx, code, y, key):
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
    """サービスの年度延べ受給者数。受給者数の系列がなければ受給率から復元する。"""
    for (c, i, yy), v in idx.items():
        if c == code and yy == y and i.startswith('受給者数') and i.endswith('（合計）'):
            try:
                return float(v)
            except (TypeError, ValueError):
                return None
    rt = rate(idx, code, y, '合計')
    ins = num(idx, 'D4', '第1号被保険者数', y)
    return rt / 100 * ins * 12 if rt is not None and ins else None


def clip(x, c):
    return max(-c, min(c, x))


def main():
    idx = load()
    # 起点（基準年）と予測年の組。窓は5年、予測は3年先。
    CASES = [('2019', '2022'), ('2020', '2023'), ('2021', '2024'), ('2022', '2025')]
    GRID = [(phi, c) for phi in (0.0, 0.25, 0.5, 0.75, 1.0)
            for c in (0.03, 0.05, 0.10, 0.20)]

    results = {}
    detail = []
    for phi, c in GRID:
        errs, wts = [], []
        for base, target in CASES:
            win = str(int(base) - 5)      # 窓5年の起点
            gn0 = num(idx, 'B4-a', '合計認定者数', win)
            gn1 = num(idx, 'B4-a', '合計認定者数', base)
            gn2 = num(idx, 'B4-a', '合計認定者数', target)
            if not (gn0 and gn1 and gn2):
                continue
            gn = (gn1 / gn0) ** (1 / 5) - 1          # 全体の年率（窓5年）
            for name, code, pcode in SVC:
                q_win, q_base, q_tgt = juk(idx, code, win), juk(idx, code, base), juk(idx, code, target)
                if not (q_win and q_base and q_tgt) or q_win <= 0 or q_base <= 0:
                    continue
                gs = (q_base / q_win) ** (1 / 5) - 1
                g = gn + phi * clip(gs - gn, c)
                pred = q_base * (1 + g) ** 3
                price = num(idx, pcode, f'受給者1人あたり給付月額（{name}）', base) or 0
                w = q_tgt / 12 * price                # 給付費で重み付け
                if w <= 0:
                    continue
                errs.append(abs(pred - q_tgt) / q_tgt * w)
                wts.append(w)
        if wts:
            results[(phi, c)] = sum(errs) / sum(wts)

    order = sorted(results.items(), key=lambda kv: kv[1])
    print('■ バックテスト（起点4通り×11サービス、3年先を予測）')
    print('  評価は給付費で重み付けした加重平均絶対誤差（WMAPE）')
    print(f"  {'φ':>5}{'上限c':>8}{'WMAPE':>10}")
    for (phi, c), e in order:
        mark = ' ←最小' if (phi, c) == order[0][0] else (' （一律）' if phi == 0 else '')
        print(f"  {phi:>5.2f}{c*100:>7.0f}%{e*100:>9.1f}%{mark}")

    best = order[0][0]
    flat = results[(0.0, 0.03)]
    print()
    print(f"  一律（φ=0）のWMAPE {flat*100:.1f}%　最小 {order[0][1]*100:.1f}%（φ={best[0]}・c={best[1]*100:.0f}%）")
    if order[0][1] >= flat * 0.98:
        print('  → 改善が2%未満。本村の規模では一律に対する優位が認められない。')
    else:
        print(f"  → {(1-order[0][1]/flat)*100:.1f}%の改善。")

    path = os.path.join(DATA, '第10期_伸びの反映割合のバックテスト.csv')
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['反映割合φ', '乖離の上限c(%)', 'WMAPE(%)'])
        for (phi, c), e in order:
            w.writerow([phi, round(c * 100), round(e * 100, 2)])
    print(f'\n保存: {os.path.relpath(path, BASE)}')


if __name__ == '__main__':
    main()
