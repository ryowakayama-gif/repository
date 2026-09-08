# -*- coding: utf-8 -*-
"""
建設改良費の改定が総括原価・必要改定幅に与える影響の試算

  建設改良費 → ①減価償却費 ②長期前受金戻入益 ③支払利息 ④企業債償還金 → 総括原価 → 必要改定率

【旧計画】01_source_evidence/mgmt_strategy_simulation_sheet.xlsx
         【成行】財政シミュ r86+r90（建設改良費）, r74〜r77（財源）, 新規取得シート（償却）
【新計画】06_capital_plan_update/construction_cost_survey_R8toR17.xlsx
         （女川町回答の建設改良費調査票。No.3 の繰入金は数式誤りを補正）

すべて千円単位。可変パラメータは RATES と DEP_SCENARIOS。
"""

YRS = ['R8','R9','R10','R11','R12','R13','R14','R15','R16','R17']
N = len(YRS)

# ============================================================
# 可変パラメータ
# ============================================================
# 借入予定利率。経営戦略は「新規債返済スケジュール」C3 = 2.1%（加重平均借入利率）を採用。
# 調査票の「利率(%)」欄は4事業とも未入力のため、2〜3%で振って感応度をみる。
RATES = [0.020, 0.021, 0.025, 0.030]
RATE_LABELS = {0.020:'2.0%', 0.021:'2.1%（戦略踏襲）', 0.025:'2.5%', 0.030:'3.0%'}
BASE_RATE = 0.021

# 償却年数シナリオ（事業No → 年数）
#   経営戦略は新規取得資産を機械・電気も含めて一律38年で償却している（新規取得シート G10:G58 が全て38）。
#   調査票は4事業とも40年。No.3 鷲神高度処理は機械・電気設備であり法定耐用年数はこれより短い。
DEP_SCENARIOS = {
    '経営戦略準拠 一律38年':      {1:38, 2:38, 3:38, 4:38},
    '調査票どおり 一律40年':      {1:40, 2:40, 3:40, 4:40},
    'No.3を機械電気20年':         {1:40, 2:40, 3:20, 4:40},
    'No.3を機械電気15年':         {1:40, 2:40, 3:15, 4:40},
}
BASE_DEP = '経営戦略準拠 一律38年'

DEP_OLD   = 38          # 旧計画の償却年数（経営戦略の実際の設定）
YUSHU_5Y  = 4_316_000   # R8〜R12 有収水量（㎥）
REV_5Y    = 590_845     # R8〜R12 給水収益 5年計（千円）
ASSET_AVG = 896_312     # 対象資産（期首期末平均・千円）
AM_RATE   = 0.03        # 資産維持率（業務仕様書の正式ケース）

# 総括原価の現行前提（01_rate_reform_simulation_MAIN.xlsx 02_総括原価）
MAINT   = [131_926,133_182,134_453,135_742,137_024]   # 維持管理費
DEP_B   = [395_267,414_923,412_863,407_153,402_234]   # 減価償却費
INT_OLD = [ 10_790, 10_530, 10_172,  9_805,  9_288]   # 既往債利息
INT_NEW = [  9_238, 18_507, 26_641, 34_316, 38_700]   # 新規債利息
KOUJU   = [-94_000,-92_000,-90_000,-88_000,-86_000]   # 長期前受金戻入（控除）
HOJO    = [-3_000]*5                                  # 他会計補助金等（控除）

# ============================================================
# 旧計画（現行の財政シミュレーション）
# ============================================================
OLD = {
    '整備費':     [1_030_000, 950_000, 500_000, 330_000, 310_000, 311_000, 371_000, 294_000, 230_000, 200_000],
    '国庫補助金':  [   63_330,  50_000,  50_000,  50_000,  50_000,       0,       0,       0,       0,       0],
    '他会計補助金':[    3_000,   3_000,   3_000,   3_000,   3_000,   3_000,   3_000,   3_000,   3_000,   3_000],
    '出資金':     [  515_000, 495_000,  60_000,  40_000,  20_000,  20_000,  20_000,  20_000,  20_000,  20_000],
    '企業債':     [  448_670, 402_000, 387_000, 237_000, 237_000, 287_000, 347_000, 267_000, 187_000, 157_000],
}
OLD_REDEEM, OLD_GRACE = 40, 0   # 40年元利均等・借入年度の翌年度から償還

# ============================================================
# 新計画（建設改良費調査票 No.1〜No.4）
# ============================================================
PROJECTS = [
    {'no':1,'name':'老朽管布設替工事','asset':'その他（管路）','redeem':25,'grace':3,
     'cost':[180_000,150_000,150_000,100_000,100_000,100_000,100_000,100_000,100_000,100_000],
     '国費':[ 80_000, 80_000, 80_000,      0,      0,      0,      0,      0,      0,      0],
     '企業債':[100_000, 70_000, 70_000,100_000,100_000,100_000,100_000,100_000,100_000,100_000],
     '繰入金':[0]*10, 'その他':[0]*10},
    {'no':2,'name':'水道施設耐震化工事','asset':'配水池','redeem':40,'grace':3,
     'cost':[ 12_000, 35_000, 60_000,220_000,220_000,220_000, 60_000,150_000,150_000,150_000],
     '国費':[  6_000, 17_500, 15_000, 55_000, 55_000, 55_000, 15_000, 37_500, 37_500, 37_500],
     '企業債':[      0,      0, 45_000,165_000,165_000,165_000, 45_000,112_500,112_500,112_500],
     '繰入金':[0]*10, 'その他':[6_000,17_500,0,0,0,0,0,0,0,0]},
    {'no':3,'name':'鷲神高度処理設備整備工事','asset':'浄水場（機械・電気）','redeem':40,'grace':3,
     'cost':[353_400,534_600,0,0,0,0,0,0,0,0],
     '国費':[0]*10,
     '企業債':[176_700,267_300,0,0,0,0,0,0,0,0],
     # ★調査票の繰入金セルは数式誤り。整備費−企業債＝50%を繰入金として補正
     '繰入金':[176_700,267_300,0,0,0,0,0,0,0,0], 'その他':[0]*10},
    {'no':4,'name':'江島海底送水管本復旧工事','asset':'（未入力）','redeem':40,'grace':3,
     'cost':[ 20_000,150_000,0,0,0,0,0,0,0,0],
     '国費':[0]*10,
     '企業債':[ 15_000,112_500,0,0,0,0,0,0,0,0],
     '繰入金':[0]*10, 'その他':[5_000,37_500,0,0,0,0,0,0,0,0]},
]
NEW = {k2: [sum(p[k1][i] for p in PROJECTS) for i in range(N)]
       for k1,k2 in [('cost','整備費'),('国費','国費'),('企業債','企業債'),('繰入金','繰入金'),('その他','その他')]}

# ============================================================
# 計算関数
# ============================================================
def annuity(principal, rate, years):
    return principal * rate / (1 - (1 + rate) ** -years)

def bond_profile(draws, rate, years, grace, horizon=N):
    """借入額の系列から、各年度の（元金償還額, 支払利息, 年度末残高）を返す。
    借入年度は残高のみ。grace 年は据置（利息のみ）。以後、元利均等で償還。"""
    prin=[0.0]*horizon; intr=[0.0]*horizon; bal=[0.0]*horizon
    for s, draw in enumerate(draws):
        if draw <= 0: continue
        pay = annuity(draw, rate, years); out = draw
        for t in range(s, horizon):
            if t == s:
                bal[t] += out; continue
            i = out * rate
            if t - s <= grace:
                intr[t] += i; bal[t] += out
            else:
                p = min(pay - i, out); out -= p
                prin[t] += p; intr[t] += i; bal[t] += out
    return prin, intr, bal

def straight_line(invest, life, horizon=N):
    """取得年度の翌年度から満額償却（経営戦略「新規取得」シートと同じ扱い）"""
    return [sum(v/life for v in invest[:t]) for t in range(horizon)]

def new_plan(rate, dep_map):
    """新計画の（減価償却費, 長期前受金戻入, 支払利息, 元金償還, 残高）を返す"""
    dep=[0.0]*N; ch=[0.0]*N; pr=[0.0]*N; it=[0.0]*N; bl=[0.0]*N
    for p in PROJECTS:
        life = dep_map[p['no']]
        for i,v in enumerate(straight_line(p['cost'], life)): dep[i]+=v
        for i,v in enumerate(straight_line(p['国費'], life)): ch[i]+=v
        a,b,c = bond_profile(p['企業債'], rate, p['redeem'], p['grace'])
        for i in range(N): pr[i]+=a[i]; it[i]+=b[i]; bl[i]+=c[i]
    return dep, ch, pr, it, bl

def old_plan(rate):
    dep = straight_line(OLD['整備費'], DEP_OLD)
    src = [OLD['国庫補助金'][i]+OLD['他会計補助金'][i] for i in range(N)]
    ch  = straight_line(src, DEP_OLD)
    pr, it, bl = bond_profile(OLD['企業債'], rate, OLD_REDEEM, OLD_GRACE)
    return dep, ch, pr, it, bl

def total_cost_5y(rate, dep_map):
    """R8〜R12 の総括原価5年計（千円）と内訳を返す。
    旧計画は経営戦略が定めた利率2.1%で固定し、新計画の借入利率のみを可変とする
    （旧計画は「現行シミュがそう置いている」という所与の比較基準であるため）。"""
    o_dep,o_ch,_,o_it,_ = old_plan(BASE_RATE)
    n_dep,n_ch,_,n_it,_ = new_plan(rate, dep_map)
    am = round(ASSET_AVG * AM_RATE)
    base = sum(MAINT)+sum(DEP_B)+sum(INT_OLD)+sum(INT_NEW)+am*5+sum(KOUJU)+sum(HOJO)
    d_dep = sum(n_dep[i]-o_dep[i] for i in range(5))
    d_it  = sum(n_it[i]-o_it[i] for i in range(5))
    d_ch  = sum(n_ch[i]-o_ch[i] for i in range(5))
    return base, base + d_dep + d_it - d_ch, d_dep, d_it, d_ch

def metrics(cost5):
    g = cost5*1000/YUSHU_5Y
    return g, 128.2/g, (cost5-REV_5Y)/5, cost5/REV_5Y

# ============================================================
# 出力
# ============================================================
def row(label, vals, tot=True, w=26):
    s = label.ljust(w) + ''.join(f'{v:>10,.0f}' for v in vals)
    print(s + (f'{sum(vals):>13,.0f}' if tot else ''))

def header(title):
    print('\n'+'='*152); print(title); print('='*152)
    print(' '*26 + ''.join(y.rjust(10) for y in YRS) + '合計'.rjust(11))

print(f"""
建設改良費の改定 → 必要改定幅への影響
  基準ケース：借入利率 {RATE_LABELS[BASE_RATE]} ／ 償却年数 {BASE_DEP} ／ 資産維持率 {AM_RATE:.0%}""")

header('【1】建設改良費と財源構成の比較（千円）')
row('旧 整備費', OLD['整備費']); row('新 整備費', NEW['整備費'])
row('  差引', [NEW['整備費'][i]-OLD['整備費'][i] for i in range(N)])
print('-'*152)
row('旧 企業債', OLD['企業債']); row('新 企業債', NEW['企業債'])
row('  差引', [NEW['企業債'][i]-OLD['企業債'][i] for i in range(N)])
print('-'*152)
row('旧 国庫補助金', OLD['国庫補助金']);   row('新 国費', NEW['国費'])
row('旧 出資金', OLD['出資金']);          row('新 繰入金', NEW['繰入金'])
row('旧 他会計補助金', OLD['他会計補助金']); row('新 その他', NEW['その他'])

print('\n'+'='*152)
print('【2】経営戦略における償却年数の設定（確認結果）')
print('='*152)
print("""
  経営戦略試算シート「新規取得」シートを確認した結果、R7以降に取得する資産の償却年数は
  【機械・電気設備を区分せず、一律38年】で設定されている。

    ・G10:G58（取得年度 R7〜R55 の49行）が すべて 38
    ・原水及び浄水施設費（D列＝浄水場・機械電気設備を含む）と
      給水及び配水施設費（E列＝管路）を同じ38年で償却している
    ・長期前受金の収益化も同じ38年（H3行 =【成行】財政シミュ!$V$77/38）
    ・償却開始は取得年度の翌年度から（J10 の帳簿価額ロールフォワード）

  なお「新規取得」シートB4・B8の注記は【更新のタイミング】の前提であり、償却年数ではない。

    ①構造物及び設備：耐用年数が20年未満の資産について、従前、大規模な修繕が1件程度しか
      発生していないため、法定耐用年数×1.5のタイミングで、それ以外は法定用年数の1.2の
      タイミングで、試算時点の帳簿価額にデフレーターを加味した金額で試算
    ②管路：総務省提供資産ツールに基づき、法定耐用年数×1.2のタイミングで、基準単価×数量で試算

  → 「耐用年数が20年未満の資産」の存在は認識されているが、償却計算には反映されていない。
     BS（【BS】決算数値入力）では「機械及び装置」を独立科目として持っているため、
     既往資産は区分されている一方、将来取得分だけが一律38年になっている。

  ★注意：支払利息の算定根拠欄（【成行】財政シミュ r61）には「利率の平均値（1.79％）」とあるが、
    実際の計算は「新規債返済スケジュール」C3＝2.1％で行われており、報告書案の記載も2.1％。
    根拠欄の1.79％は更新漏れとみられる。""")

o_dep,o_ch,o_pr,o_it,o_bl = old_plan(BASE_RATE)
n_dep,n_ch,n_pr,n_it,n_bl = new_plan(BASE_RATE, DEP_SCENARIOS[BASE_DEP])

header(f'【3】年度別の影響（千円／年）　利率{RATE_LABELS[BASE_RATE]}・{BASE_DEP}')
row('旧 減価償却費(新規分)', o_dep, False); row('新 減価償却費(新規分)', n_dep, False)
row('  ① 差引', [n_dep[i]-o_dep[i] for i in range(N)], False)
print('-'*152)
row('旧 長期前受金戻入', o_ch, False); row('新 長期前受金戻入', n_ch, False)
row('  ② 差引', [n_ch[i]-o_ch[i] for i in range(N)], False)
print('-'*152)
row('旧 支払利息(新規債)', o_it, False); row('新 支払利息(新規債)', n_it, False)
row('  ③ 差引', [n_it[i]-o_it[i] for i in range(N)], False)
print('-'*152)
row('★総括原価への影響 ①-②+③', [(n_dep[i]-o_dep[i])-(n_ch[i]-o_ch[i])+(n_it[i]-o_it[i]) for i in range(N)], False)
print('-'*152)
row('旧 企業債償還金(新規債)', o_pr, False); row('新 企業債償還金(新規債)', n_pr, False)
row('  ④ 差引', [n_pr[i]-o_pr[i] for i in range(N)], False)
row('  企業債残高 差引', [n_bl[i]-o_bl[i] for i in range(N)], False)

# ============================================================
# 感応度マトリクス
# ============================================================
print('\n'+'='*152)
print('【4】利率 × 償却年数 の感応度（R8〜R12・資産維持率3%）')
print('='*152)

am = round(ASSET_AVG*AM_RATE)
for metric, fmt, title in [
        (0,'{:>13,.0f}','総括原価 5年計（千円）'),
        (1,'{:>13.1f}','給水原価（円/㎥）'),
        (2,'{:>13,.0f}','年間不足額（千円/年）'),
        (3,'{:>13.2f}','必要改定率（総括原価÷現行収入・倍）')]:
    print(f'\n  ◆ {title}')
    print('  ' + '償却年数シナリオ'.ljust(26) + ''.join(RATE_LABELS[r].rjust(14) for r in RATES))
    print('  ' + '-'*82)
    for dname, dmap in DEP_SCENARIOS.items():
        cells=[]
        for r in RATES:
            base, after, *_ = total_cost_5y(r, dmap)
            g, rec, short, ratio = metrics(after)
            cells.append([after, g, short, ratio][metric])
        mark = '★' if dname == BASE_DEP else '  '
        print(f'  {mark}{dname:24s}' + ''.join(fmt.format(c)+' ' for c in cells))
    b = [sum(MAINT)+sum(DEP_B)+sum(INT_OLD)+sum(INT_NEW)+am*5+sum(KOUJU)+sum(HOJO)]*1
    g, rec, short, ratio = metrics(b[0])
    print('  ' + '-'*82)
    print(f'  {"（参考）旧計画":24s}  ' + fmt.format([b[0], g, short, ratio][metric]).strip())

print(f'''
  ◆ 新計画の新規債 支払利息（R8〜R12 5年計・千円）※利率のみを変えた場合の絶対額
  {"":26s}''' + ''.join(RATE_LABELS[r].rjust(14) for r in RATES))
print('  ' + '-'*82)
_ints = []
for r in RATES:
    _,_,_,it,_ = new_plan(r, DEP_SCENARIOS[BASE_DEP])
    _ints.append(sum(it[:5]))
print('  ' + '新計画 R8〜R12'.ljust(24) + ''.join(f'{v:>13,.0f} ' for v in _ints))
_,_,_,oit,_ = old_plan(BASE_RATE)
print('  ' + '-'*82)
print(f'  （参考）旧計画 R8〜R12（利率2.1%固定）  {sum(oit[:5]):,.0f}')
print(f'  → 利率が2.0%→3.0%に振れると新計画の支払利息は{_ints[0]:,.0f}→{_ints[-1]:,.0f}千円（+{_ints[-1]-_ints[0]:,.0f}千円）')

# ============================================================
# 北上川導水管が調査票に計上されていない件の感応度
# ============================================================
import copy
_ORIG = PROJECTS[:]
PROJECTS = _ORIG + [{'no':99,'name':'北上川導水管（調査票に未計上・旧計画から復元）',
    'asset':'その他（導水管）','redeem':40,'grace':3,
    'cost':[0,670_000,0,0,0,0,0,0,0,0], '国費':[0]*10,
    '企業債':[0,502_500,0,0,0,0,0,0,0,0], '繰入金':[0]*10,
    'その他':[0,167_500,0,0,0,0,0,0,0,0]}]
print(f'''
  ◆ 北上川導水管6.7億円を復元した場合の給水原価（円/㎥）
    旧計画R9の原水及び浄水施設費670,000千円。企業債75%と仮定。償却年数は各シナリオに従う。
  {"":26s}''' + ''.join(RATE_LABELS[r].rjust(14) for r in RATES))
print('  ' + '-'*82)
for dname, dmap in DEP_SCENARIOS.items():
    dmap = dict(dmap); dmap[99] = dmap[1]
    cells=[]
    for r in RATES:
        _, after, *_ = total_cost_5y(r, dmap)
        cells.append(metrics(after)[0])
    mark = '★' if dname == BASE_DEP else '  '
    print(f'  {mark}{dname:24s}' + ''.join(f'{c:>13.1f} ' for c in cells))
print('  ' + '-'*82)
print('  （参考）旧計画                   591.3')
PROJECTS = _ORIG

# ============================================================
# 基準ケースのまとめ
# ============================================================
base, after, d_dep, d_it, d_ch = total_cost_5y(BASE_RATE, DEP_SCENARIOS[BASE_DEP])
g0,r0,s0,x0 = metrics(base); g1,r1,s1,x1 = metrics(after)
print('\n'+'='*152)
print(f'【5】基準ケースのまとめ（利率{RATE_LABELS[BASE_RATE]}・{BASE_DEP}・資産維持率{AM_RATE:.0%}）')
print('='*152)
print(f'  {"":34s}{"旧計画":>16s}{"新調査票":>16s}{"差引":>14s}')
print('  '+'-'*80)
print(f'  {"総括原価 5年計（千円）":32s}{base:>16,.0f}{after:>16,.0f}{after-base:>14,.0f}')
print(f'  {"　うち 減価償却費の増減":32s}{"":>16s}{"":>16s}{d_dep:>14,.0f}')
print(f'  {"　うち 支払利息の増減":32s}{"":>16s}{"":>16s}{d_it:>14,.0f}')
print(f'  {"　うち 長期前受金戻入の増減":32s}{"":>16s}{"":>16s}{-d_ch:>14,.0f}')
print(f'  {"給水原価（円/㎥）":32s}{g0:>16,.1f}{g1:>16,.1f}{g1-g0:>14,.1f}')
print(f'  {"料金回収率（現行128.2円）":32s}{r0:>15.1%}{r1:>16.1%}')
print(f'  {"年間不足額（千円/年）":32s}{s0:>16,.0f}{s1:>16,.0f}{s1-s0:>14,.0f}')
print(f'  {"必要改定率（総括原価÷現行収入）":32s}{x0:>15.2f}倍{x1:>15.2f}倍{x1-x0:>13.2f}倍')
c0 = REV_5Y/5 - (sum(MAINT)/5 + (sum(INT_OLD)+sum(INT_NEW))/5 + am)
c1 = c0 - d_it/5
print(f'  {"キャッシュ不足額 定義③（千円/年）":32s}{c0:>16,.0f}{c1:>16,.0f}{c1-c0:>14,.0f}')
print(f'  {"　→ 赤字半減の目標増収額":32s}{-c0/2:>16,.0f}{-c1/2:>16,.0f}{(c0-c1)/2:>14,.0f}')
