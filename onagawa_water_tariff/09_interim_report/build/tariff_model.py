# -*- coding: utf-8 -*-
"""料金改定パターンの算定（中間報告書用）"""

METER = {13:55, 20:77, 25:110, 40:220, 50:880, 75:1100, 100:1320, 150:3300}

def current(vol, dia):
    """現行（用途別・逓減制・税込）"""
    base = 990 if vol <= 5 else 1210
    ex = 0
    if vol > 10:
        ex += min(vol, 100) * 0 + (min(vol,100) - 10) * 121
        if vol > 100: ex += (vol - 100) * 110
    return base + ex + METER[dia]

# 検証
for v, d, exp in [(20,20,2497),(200,25,23210),(1000,50,111980),(3000,75,332200),(8000,100,882420)]:
    got = current(v, d)
    assert got == exp, f'{v}㎥/{d}mm: {got} != {exp}'
print('現行料金体系の再現：全件一致 ✓')

# ============================================================
# 水量ゾーン（月次調定サンプル）
# ============================================================
ZONES = [
    ('① 0㎥',          0.0, 297, 13),
    ('② 1〜5㎥',        2.9, 513, 13),
    ('③ 6〜10㎥',       8.0, 590, 13),
    ('④ 11〜20㎥',     14.9, 856, 20),
    ('⑤ 21〜50㎥',     28.7, 672, 20),
    ('⑥ 51〜100㎥',    71.6,  55, 25),
    ('⑦ 101〜500㎥',  197.5,  55, 25),
    ('⑧ 501㎥超',    1030.5,  18, 50),
]
SAMPLE_MONTH = sum(c * current(v, d) for _, v, c, d in ZONES)
ACTUAL_YEAR = 118_169_000          # 現行料金収入（年・円）R8〜R12平均
SCALE = ACTUAL_YEAR / (SAMPLE_MONTH * 12)
print(f'ゾーン模型の月次収入 {SAMPLE_MONTH:,.0f}円 → 年 {SAMPLE_MONTH*12:,.0f}円')
print(f'実際の年間料金収入 {ACTUAL_YEAR:,}円　補正係数 {SCALE:.4f}')

# ============================================================
# 水産加工業19社モデル
# ============================================================
FISH = [('小規模（家族経営）',25,200,5), ('中規模（一般加工）',50,1000,8),
        ('大規模（冷凍・主要）',75,3000,5), ('超大口（大手）',100,8000,1)]

# ============================================================
# 確定した前提での財政指標
# ============================================================
# R02対応：既存資産の償却費を財政シミュの実計上額に接続し直した（+41,318千円）
#   R6末取得資産 1,822,881 ＋ R7取得資産 32,456 ＝ 残すべき既存分 1,855,337
#   ＋ 調査票に基づく新規分 147,642 ＝ 2,002,979千円（従前 1,961,661）
TOTAL_COST_5Y   = 1_175_810        # 総括原価5年計（千円）※基準内繰入金を控除した仮定計算
TOTAL_COST_MAX  = 1_448_807        # 　参考：基準内繰入金を控除しない場合
KIJUNNAI_5Y     =   272_997        # 基準内繰入金（他会計補助金）5年計（千円）
REV_5Y          =   590_845        # 給水収益5年計（千円）
YUSHU_5Y        = 4_316_000        # 有収水量5年計（㎥）
MAINT_5Y        =   719_494        # 維持管理費5年計（千円）
DEP_5Y          = 2_002_979        # 減価償却費5年計（千円）※R02対応後
INT_5Y          =   157_345        # 支払利息5年計（既往50,585＋新規106,760）
AM_YEAR         =    26_889        # 資産維持費（年・千円）

kyusui_genka  = TOTAL_COST_5Y * 1000 / YUSHU_5Y
# 供給単価は3種類あり、混用できない
TANKA_R7_INC  = 128.2                              # 令和7年度実績・税込調定単価
TANKA_R7_EX   = 116.6                              # 令和7年度実績・税抜（消費税相当額11,948,075円を除く）
kyoukyu_tanka = REV_5Y / 5 * 1000 / (YUSHU_5Y / 5)  # 令和8〜12年度の推計・税抜
kaishu_wariai = kyoukyu_tanka / kyusui_genka        # 総括原価回収割合（税抜どうしの比）
rev_year      = REV_5Y / 5
shortfall     = (TOTAL_COST_5Y - REV_5Y) / 5
ratio_full    = TOTAL_COST_5Y / REV_5Y
# 基準内繰入金を総括原価から控除する以上、キャッシュ不足額でも収入側に算入する
cash_short    = rev_year + KIJUNNAI_5Y/5 - (MAINT_5Y/5 + INT_5Y/5 + AM_YEAR)

print(f'''
【確定した前提】
  総括原価5年計      {TOTAL_COST_5Y:,}千円
  給水原価           {kyusui_genka:.1f}円/㎥
  供給単価（R8〜R12・税抜） {kyoukyu_tanka:.1f}円/㎥　総括原価回収割合 {kaishu_wariai:.1%}
  　（参考）R7実績 税込{TANKA_R7_INC}円/㎥・税抜{TANKA_R7_EX}円/㎥　※将来推計と同じ指標として比較しない
  現行料金収入       {rev_year:,.0f}千円/年
  年間不足額         {shortfall:,.0f}千円/年
  必要改定率（総括原価回収） {ratio_full:.2f}倍
  　（参考）基準内繰入金を控除しない場合 {TOTAL_COST_MAX/REV_5Y:.2f}倍
  キャッシュ不足額（定義③） {cash_short:,.0f}千円/年
  　→ 半減に必要な増収額   {-cash_short/2:,.0f}千円/年（現行収入比 {-cash_short/2/rev_year:+.1%}）
''')

# ============================================================
# パターン1：現行体系維持（一律改定）
# ============================================================
def p1(vol, dia, k):
    """現行体系の基本料金・超過料金・メーター使用料に一律 k を乗じる"""
    return round(current(vol, dia) * k)

# ============================================================
# パターン3-A：口径別＋段階逓増（塩竈市の体系比率を参考に水準を設定）
# ============================================================
BASE_A = {13:770, 20:1430, 25:2310, 40:5500, 50:11000, 75:22000, 100:41800}  # 塩竈市×1.1（税込）
TIER_A = [(10, 91.3), (20, 192.5), (50, 258.5), (100, 280.5), (None, 324.5)]

def p3a(vol, dia, m):
    amt = BASE_A[dia] * m
    prev = 0
    for upper, rate in TIER_A:
        if upper is None:
            amt += max(0, vol - prev) * rate * m; break
        amt += max(0, min(vol, upper) - prev) * rate * m
        prev = upper
        if vol <= upper: break
    return round(amt)

def p3b(vol, dia, m, ind=False):
    """3-B：産業用特例（月500㎥以上の指定業種）は基本料金なし・115円/㎥"""
    if ind and vol >= 500: return round(vol * 115)
    return p3a(vol, dia, m)

# ============================================================
# 収入の算定
# ============================================================
def revenue(fn):
    """ゾーン模型から年間料金収入（円）を推計"""
    return sum(c * fn(v, d) for _, v, c, d in ZONES) * 12 * SCALE

CUR_REV = revenue(lambda v,d: current(v,d))
TARGET  = TOTAL_COST_5Y / 5 * 1000          # 総括原価 年平均（円）

# パターン1の改定率
K_FULL = ratio_full                          # 2.38倍（総括原価100%回収）
K_HALF = 1 + (-cash_short/2) / rev_year      # 赤字半減

# パターン3-Aの水準（総括原価100%回収に合わせる）
lo, hi = 0.1, 10.0
for _ in range(60):
    mid = (lo+hi)/2
    if revenue(lambda v,d: p3a(v,d,mid)) < TARGET: lo = mid
    else: hi = mid
M_A = (lo+hi)/2

print(f'''
【水準の設定】
  現行の年間料金収入（模型）   {CUR_REV:,.0f}円
  総括原価（年平均）           {TARGET:,.0f}円
  パターン1 必要改定率         {K_FULL:.3f}倍
  パターン2 赤字半減の改定率   {K_HALF:.3f}倍
  パターン3-A 体系の水準係数   {M_A:.4f}（塩竈市体系比×{M_A:.3f}）
  　→ 3-Aの年間収入           {revenue(lambda v,d: p3a(v,d,M_A)):,.0f}円
''')

# ============================================================
# 影響の算定
# ============================================================
PATTERNS = [
 ('現行',                                  lambda v,d: current(v,d)),
 (f'P1-A 現行体系×{K_FULL:.2f}',            lambda v,d: p1(v,d,K_FULL)),
 ('P1-B 現行体系×1.50',                    lambda v,d: p1(v,d,1.50)),
 (f'P2 赤字半減×{K_HALF:.3f}',              lambda v,d: p1(v,d,K_HALF)),
 ('P3-A 口径別逓増',                        lambda v,d: p3a(v,d,M_A)),
]

print('■ 水量ゾーン別 月額料金（円）と現行比')
hdr = f"{'ゾーン':13s}{'水量':>7s}{'件数':>6s}"+''.join(f'{n.split()[0]:>12s}' for n,_ in PATTERNS)
print(hdr); print('-'*len(hdr))
for name, v, c, d in ZONES:
    vals=[fn(v,d) for _,fn in PATTERNS]
    print(f"{name:13s}{v:>7.1f}{c:>6d}"+''.join(f'{x:>12,}' for x in vals))
print(' '*26+''.join(f'{"":>12s}' for _ in PATTERNS))
print(f"{'現行比':13s}{'':>7s}{'':>6s}"+''.join(f'{"":>12s}' for _ in PATTERNS))
for name, v, c, d in ZONES:
    base=current(v,d); vals=[fn(v,d)/base for _,fn in PATTERNS]
    print(f"{name:13s}{v:>7.1f}{c:>6d}"+''.join(f'{x:>11.2f}倍' for x in vals))

print()
print('■ 標準家庭（20mm・月20㎥）の月額')
for n,fn in PATTERNS:
    print(f'  {n:22s} {fn(20,20):>8,}円  （現行比 {fn(20,20)/current(20,20):.2f}倍・{fn(20,20)-current(20,20):+,}円）')

print()
print('■ 水産加工業19社モデル 月額（円）')
hdr=f"{'区分':16s}{'口径':>6s}{'水量':>7s}{'社数':>5s}"+''.join(f'{n.split()[0]:>12s}' for n,_ in PATTERNS)+f"{'P3-B産業特例':>14s}"
print(hdr); print('-'*len(hdr))
for nm, d, v, c in FISH:
    vals=[fn(v,d) for _,fn in PATTERNS]
    b=p3b(v,d,M_A,ind=True)
    print(f"{nm:16s}{d:>6d}{v:>7d}{c:>5d}"+''.join(f'{x:>12,}' for x in vals)+f"{b:>14,}")
print(f"{'現行比':16s}{'':>6s}{'':>7s}{'':>5s}"+''.join(f'{"":>12s}' for _ in PATTERNS))
for nm, d, v, c in FISH:
    base=current(v,d); vals=[fn(v,d)/base for _,fn in PATTERNS]
    b=p3b(v,d,M_A,ind=True)/base
    print(f"{nm:16s}{d:>6d}{v:>7d}{c:>5d}"+''.join(f'{x:>11.2f}倍' for x in vals)+f"{b:>13.2f}倍")

# 19社の合計額・産業特例による減収額は算定しない。
# 口径別の実績と母集団が接続しないため（100mmは1社の仮定が口径全体実績の約5.71倍、
# 水量ゾーン別試算の501㎥超は年約233千㎥に対し19社の500㎥以上だけで年372千㎥）。
# 使用者別の調定明細を受領したうえで、同一母集団に各案を適用し直して算定する。

print()
print('■ 年間料金収入と総括原価回収割合（税抜・水量ゾーン別試算による推計）')
for n,fn in PATTERNS:
    r=revenue(fn)
    print(f'  {n:26s} {r/1e6:>8.1f}百万円  供給単価 {r/(YUSHU_5Y/5):>6.1f}円/㎥  回収割合 {r/(YUSHU_5Y/5)/kyusui_genka:>6.1%}')
print('  P3-B 産業特例適用          算定していない（母集団が接続しないため）')

# ============================================================
# 特殊用途区分（湯屋用・船舶用）への影響
# ============================================================
def yuya(vol):      # 湯屋用：基本36,300円（500㎥含む）＋超過110円/㎥
    return 36300 + max(0, vol-500)*110
def sempaku(vol):   # 船舶用（直接）：基本2,420円（10㎥含む）＋242円/㎥
    return 2420 + max(0, vol-10)*242

print()
print('■ 特殊用途区分への影響（月額・円）')
print(f"{'区分':18s}{'水量':>7s}{'現行':>10s}{'P1-A':>11s}{'P2':>10s}{'P3-A（区分廃止）':>18s}")
for nm, f, vols, dia in [('湯屋用', yuya, [500, 800], 40), ('船舶用（直接）', sempaku, [30, 100], 20)]:
    for v in vols:
        cur = f(v)
        print(f"{nm:18s}{v:>7d}{cur:>10,}{round(cur*K_FULL):>11,}{round(cur*K_HALF):>10,}{p3a(v,dia,M_A):>18,}")

print()
print('■ 段階改定の例（パターン2を第1次とする2段階）')
steps=[('第1次（R9.4適用開始）', K_HALF, '料金不足額の半減'),
       ('第2次（R13〜R17）', K_FULL, '総括原価の回収')]
for nm,k,aim in steps:
    r=revenue(lambda v,d: p1(v,d,k))
    print(f'  {nm:24s} 改定率{k:.3f}倍  収入{r/1e6:6.1f}百万円  回収割合{r/(YUSHU_5Y/5)/kyusui_genka:6.1%}  標準家庭{p1(20,20,k):,}円  （{aim}）')

# ============================================================
# 結果の保存
# ============================================================
import json, os
res = dict(
  総括原価5年計=TOTAL_COST_5Y, 給水原価=round(kyusui_genka,1),
  供給単価_R8toR12_税抜=round(kyoukyu_tanka,1),
  総括原価回収割合=round(kaishu_wariai,3),
  供給単価_R7実績_税込=TANKA_R7_INC, 供給単価_R7実績_税抜=TANKA_R7_EX,
  現行料金収入年=round(rev_year), 年間不足額=round(shortfall),
  必要改定率=round(ratio_full,3), キャッシュ不足額=round(cash_short),
  赤字半減改定率=round(K_HALF,3), P3A水準係数=round(M_A,4),
  ゾーン=[{'名称':n,'水量':v,'件数':c,'口径':d,
          '現行':current(v,d),'P1A':p1(v,d,K_FULL),'P1B':p1(v,d,1.5),
          'P2':p1(v,d,K_HALF),'P3A':p3a(v,d,M_A)} for n,v,c,d in ZONES],
  水産加工=[{'区分':n,'口径':d,'水量':v,'社数':c,
            '現行':current(v,d),'P1A':p1(v,d,K_FULL),'P1B':p1(v,d,1.5),
            'P2':p1(v,d,K_HALF),'P3A':p3a(v,d,M_A),'P3B':p3b(v,d,M_A,ind=True)} for n,d,v,c in FISH],
  基本料金3A={k:round(v*M_A) for k,v in BASE_A.items()},
  従量料金3A=[(u, round(r*M_A,1)) for u,r in TIER_A],
)
with open(os.path.join(os.path.dirname(__file__),'figures.json'),'w',encoding='utf-8') as f:
    json.dump(res,f,ensure_ascii=False,indent=1)
print('\n算定結果を figures.json に保存しました')
print('\n■ パターン3-A 料金表（税込・円）')
print('  口径別基本料金:', '／'.join(f'{k}mm {round(v*M_A):,}' for k,v in BASE_A.items()))
print('  従量料金:', '／'.join(f'{"〜"+str(u)+"㎥" if u else "101㎥超"} {r*M_A:.1f}' for u,r in TIER_A))
