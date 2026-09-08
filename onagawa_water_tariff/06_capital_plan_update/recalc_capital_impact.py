# -*- coding: utf-8 -*-
"""
建設改良費の改定が総括原価・必要改定幅に与える影響の試算

  建設改良費 → ①減価償却費 ②長期前受金戻入益 ③支払利息 ④企業債償還金 → 総括原価 → 必要改定率

【旧計画】01_source_evidence/mgmt_strategy_simulation_sheet.xlsx
         【成行】財政シミュ r86+r90（建設改良費）, r74〜r77（財源）
【新計画】06_capital_plan_update/construction_cost_survey_R8toR17.xlsx
         （女川町回答の建設改良費調査票。No.3 の繰入金は数式誤りを補正）

すべて千円単位。
"""

YRS = ['R8','R9','R10','R11','R12','R13','R14','R15','R16','R17']
N = len(YRS)

# ============================================================
# 前提パラメータ（経営戦略試算シートの設定に準拠）
# ============================================================
RATE      = 0.021   # 加重平均借入利率。※調査票の「利率(%)」欄は未入力のため経営戦略の値を仮置き
DEP_YEARS = 40      # 償却年数。調査票の4事業とも40年
YUSHU_5Y  = 4_316_000        # R8〜R12 有収水量（㎥）
REV_5Y    = 590_845          # R8〜R12 給水収益 5年計（千円）

# ============================================================
# 旧計画（現行の財政シミュレーション）
# ============================================================
OLD = {
    '整備費':    [1_030_000, 950_000, 500_000, 330_000, 310_000, 311_000, 371_000, 294_000, 230_000, 200_000],
    '国庫補助金': [   63_330,  50_000,  50_000,  50_000,  50_000,       0,       0,       0,       0,       0],
    '他会計補助金':[    3_000,   3_000,   3_000,   3_000,   3_000,   3_000,   3_000,   3_000,   3_000,   3_000],
    '出資金':    [  515_000, 495_000,  60_000,  40_000,  20_000,  20_000,  20_000,  20_000,  20_000,  20_000],
    '企業債':    [  448_670, 402_000, 387_000, 237_000, 237_000, 287_000, 347_000, 267_000, 187_000, 157_000],
}

# ============================================================
# 新計画（建設改良費調査票 No.1〜No.4）
# ============================================================
PROJECTS = [
    {'no':1,'name':'老朽管布設替工事','asset':'その他','dep':40,'redeem':25,'grace':3,
     'cost':[180_000,150_000,150_000,100_000,100_000,100_000,100_000,100_000,100_000,100_000],
     '国費':[ 80_000, 80_000, 80_000,      0,      0,      0,      0,      0,      0,      0],
     '企業債':[100_000, 70_000, 70_000,100_000,100_000,100_000,100_000,100_000,100_000,100_000],
     '繰入金':[0]*10, 'その他':[0]*10},
    {'no':2,'name':'水道施設耐震化工事','asset':'配水池','dep':40,'redeem':40,'grace':3,
     'cost':[ 12_000, 35_000, 60_000,220_000,220_000,220_000, 60_000,150_000,150_000,150_000],
     '国費':[  6_000, 17_500, 15_000, 55_000, 55_000, 55_000, 15_000, 37_500, 37_500, 37_500],
     '企業債':[      0,      0, 45_000,165_000,165_000,165_000, 45_000,112_500,112_500,112_500],
     '繰入金':[0]*10, 'その他':[6_000,17_500,0,0,0,0,0,0,0,0]},
    {'no':3,'name':'鷲神高度処理設備整備工事','asset':'浄水場','dep':40,'redeem':40,'grace':3,
     'cost':[353_400,534_600,0,0,0,0,0,0,0,0],
     '国費':[0]*10,
     '企業債':[176_700,267_300,0,0,0,0,0,0,0,0],
     # ★調査票の繰入金セルは数式誤り（下記 BUG 参照）。整備費−企業債＝50%を繰入金として補正
     '繰入金':[176_700,267_300,0,0,0,0,0,0,0,0],
     'その他':[0]*10},
    {'no':4,'name':'江島海底送水管本復旧工事','asset':'（未入力）','dep':40,'redeem':40,'grace':3,
     'cost':[ 20_000,150_000,0,0,0,0,0,0,0,0],
     '国費':[0]*10,
     '企業債':[ 15_000,112_500,0,0,0,0,0,0,0,0],
     '繰入金':[0]*10, 'その他':[5_000,37_500,0,0,0,0,0,0,0,0]},
]

def agg(key):
    return [sum(p[key][i] for p in PROJECTS) for i in range(N)]

NEW = {'整備費': agg('cost'), '国費': agg('国費'), '企業債': agg('企業債'),
       '繰入金': agg('繰入金'), 'その他': agg('その他')}

# ============================================================
# 計算関数
# ============================================================
def annuity(principal, rate, years):
    """元利均等の年賦金"""
    return principal * rate / (1 - (1 + rate) ** -years)

def bond_profile(draws, rate, years, grace, horizon):
    """借入額の系列から、各年度の（元金償還額, 支払利息, 年度末残高）を返す。
    grace=0 のとき借入年度の翌年度から償還開始（現行財政シミュの方式）。
    grace=g のとき g 年据置（据置期間中は利息のみ）。"""
    prin = [0.0]*horizon; intr = [0.0]*horizon; bal = [0.0]*horizon
    for s, draw in enumerate(draws):
        if draw <= 0: continue
        pay = annuity(draw, rate, years)
        outstanding = draw
        for t in range(s, horizon):
            if t == s:                       # 借入年度：残高のみ計上
                bal[t] += outstanding; continue
            k = t - s                        # 借入からの経過年数
            if k <= grace:                   # 据置期間：利息のみ
                i = outstanding * rate
                intr[t] += i; bal[t] += outstanding
            else:
                i = outstanding * rate
                p = min(pay - i, outstanding)
                outstanding -= p
                prin[t] += p; intr[t] += i; bal[t] += outstanding
    return prin, intr, bal

def cumulative_dep(invest, dep_years, horizon):
    """取得年度の翌年度から満額償却（簡便法）"""
    return [sum(v/dep_years for v in invest[:t]) for t in range(horizon)]

# ============================================================
# 旧計画
# ============================================================
old_dep  = cumulative_dep(OLD['整備費'], DEP_YEARS, N)
old_chouki_src = [OLD['国庫補助金'][i] + OLD['他会計補助金'][i] for i in range(N)]
old_chouki = cumulative_dep(old_chouki_src, DEP_YEARS, N)
old_prin, old_int, old_bal = bond_profile(OLD['企業債'], RATE, 40, 0, N)

# ============================================================
# 新計画（事業ごとの償還条件で計算）
# ============================================================
new_dep = cumulative_dep(NEW['整備費'], DEP_YEARS, N)
new_chouki = cumulative_dep(NEW['国費'], DEP_YEARS, N)   # 国費のみ長期前受金の対象とする
new_prin = [0.0]*N; new_int = [0.0]*N; new_bal = [0.0]*N
for p in PROJECTS:
    pr, it, bl = bond_profile(p['企業債'], RATE, p['redeem'], p['grace'], N)
    for i in range(N):
        new_prin[i] += pr[i]; new_int[i] += it[i]; new_bal[i] += bl[i]

# ============================================================
# 出力
# ============================================================
def row(label, vals, tot=True):
    s = label.ljust(26) + ''.join(f'{v:>10,.0f}' for v in vals)
    if tot: s += f'{sum(vals):>13,.0f}'
    print(s)

def header(title):
    print('\n' + '='*156); print(title); print('='*156)
    print(' '*26 + ''.join(y.rjust(10) for y in YRS) + '合計'.rjust(11))

header('【1】建設改良費と財源構成の比較（千円）')
row('旧 整備費', OLD['整備費']);            row('新 整備費', NEW['整備費'])
row('  差引', [NEW['整備費'][i]-OLD['整備費'][i] for i in range(N)])
print('-'*156)
row('旧 企業債', OLD['企業債']);             row('新 企業債', NEW['企業債'])
row('  差引', [NEW['企業債'][i]-OLD['企業債'][i] for i in range(N)])
print('-'*156)
row('旧 国庫補助金', OLD['国庫補助金']);      row('新 国費', NEW['国費'])
row('旧 出資金', OLD['出資金']);             row('新 繰入金', NEW['繰入金'])
row('旧 他会計補助金', OLD['他会計補助金']);  row('新 その他', NEW['その他'])

header('【2】収益的収支への影響（千円／年）')
row('旧 減価償却費(新規分)', old_dep, False);  row('新 減価償却費(新規分)', new_dep, False)
row('  ① 差引', [new_dep[i]-old_dep[i] for i in range(N)], False)
print('-'*156)
row('旧 長期前受金戻入', old_chouki, False);   row('新 長期前受金戻入', new_chouki, False)
row('  ② 差引', [new_chouki[i]-old_chouki[i] for i in range(N)], False)
print('-'*156)
row('旧 支払利息(新規債)', old_int, False);    row('新 支払利息(新規債)', new_int, False)
row('  ③ 差引', [new_int[i]-old_int[i] for i in range(N)], False)
print('-'*156)
net = [(new_dep[i]-old_dep[i]) - (new_chouki[i]-old_chouki[i]) + (new_int[i]-old_int[i]) for i in range(N)]
row('★総括原価への影響 ①-②+③', net, False)

header('【3】資本的収支への影響（千円／年）')
row('旧 企業債償還金(新規債)', old_prin, False); row('新 企業債償還金(新規債)', new_prin, False)
row('  ④ 差引', [new_prin[i]-old_prin[i] for i in range(N)], False)
print('-'*156)
row('旧 企業債残高(新規債)', old_bal, False);    row('新 企業債残高(新規債)', new_bal, False)
row('  差引', [new_bal[i]-old_bal[i] for i in range(N)], False)

# ============================================================
# 総括原価・必要改定率への反映（R8〜R12）
# ============================================================
MAINT  = [131_926,133_182,134_453,135_742,137_024]   # 維持管理費
DEP_B  = [395_267,414_923,412_863,407_153,402_234]   # 減価償却費（現行前提）
INT_OLD= [ 10_790, 10_530, 10_172,  9_805,  9_288]   # 既往債利息
INT_NEW= [  9_238, 18_507, 26_641, 34_316, 38_700]   # 新規債利息（現行前提）
KOUJU  = [-94_000,-92_000,-90_000,-88_000,-86_000]   # 長期前受金戻入（控除）
HOJO   = [-3_000]*5                                  # 他会計補助金等（控除）
ASSET_AVG = 896_312                                  # 対象資産（期首期末平均）

header('【4】総括原価・必要改定率への反映（R8〜R12・資産維持率3%）')
print()
for rate_am in (0.03,):
    am = round(ASSET_AVG * rate_am)
    base = sum(MAINT)+sum(DEP_B)+sum(INT_OLD)+sum(INT_NEW)+am*5+sum(KOUJU)+sum(HOJO)
    d_dep    = sum(new_dep[i]-old_dep[i] for i in range(5))
    d_int    = sum(new_int[i]-old_int[i] for i in range(5))
    d_chouki = sum(new_chouki[i]-old_chouki[i] for i in range(5))
    after = base + d_dep + d_int - d_chouki
    print(f'  {"":32s}{"改定前":>16s}{"改定後":>16s}{"差引":>14s}')
    print('  ' + '-'*78)
    print(f'  {"総括原価 5年計（千円）":30s}{base:>16,.0f}{after:>16,.0f}{after-base:>14,.0f}')
    print(f'  {"　うち 減価償却費の増減":30s}{"":>16s}{"":>16s}{d_dep:>14,.0f}')
    print(f'  {"　うち 支払利息の増減":30s}{"":>16s}{"":>16s}{d_int:>14,.0f}')
    print(f'  {"　うち 長期前受金戻入の増減":30s}{"":>16s}{"":>16s}{-d_chouki:>14,.0f}')
    g0, g1 = base*1000/YUSHU_5Y, after*1000/YUSHU_5Y
    print(f'  {"給水原価（円/㎥）":30s}{g0:>16,.1f}{g1:>16,.1f}{g1-g0:>14,.1f}')
    print(f'  {"料金回収率（現行128.2円）":30s}{128.2/g0:>15.1%}{128.2/g1:>16.1%}')
    s0, s1 = (base-REV_5Y)/5, (after-REV_5Y)/5
    print(f'  {"年間不足額（千円/年）":30s}{s0:>16,.0f}{s1:>16,.0f}{s1-s0:>14,.0f}')
    r0, r1 = base/REV_5Y, after/REV_5Y
    print(f'  {"必要改定率（総括原価/現行収入）":30s}{r0:>15.2f}倍{r1:>15.2f}倍{r1-r0:>13.2f}倍')
    print()
    # キャッシュ不足額（定義③）
    c0 = REV_5Y/5 - (sum(MAINT)/5 + (sum(INT_OLD)+sum(INT_NEW))/5 + am)
    c1 = REV_5Y/5 - (sum(MAINT)/5 + (sum(INT_OLD)+sum(INT_NEW))/5 + d_int/5 + am)
    print(f'  {"キャッシュ不足額 定義③（千円/年）":30s}{c0:>16,.0f}{c1:>16,.0f}{c1-c0:>14,.0f}')
    print(f'  {"　→ 赤字半減の目標増収額":30s}{-c0/2:>16,.0f}{-c1/2:>16,.0f}{(c0-c1)/2:>14,.0f}')


# ============================================================
# 感応度分析：結論を左右する2つの前提
# ============================================================
def totalcost_5y(invest, kokuhi, bonds_by_project, am_rate=0.03):
    """R8〜R12 の総括原価5年計を返す"""
    am = round(ASSET_AVG * am_rate)
    dep = cumulative_dep(invest, DEP_YEARS, N)
    return dep, kokuhi, bonds_by_project

def scenario(label, projects, extra_note=''):
    dep_t = [0.0]*N
    for p in projects:
        d = cumulative_dep(p['cost'], p['dep'], N)
        for i in range(N): dep_t[i] += d[i]
    kok = [sum(p['国費'][i] for p in projects) for i in range(N)]
    ch  = cumulative_dep(kok, DEP_YEARS, N)
    pr_t=[0.0]*N; in_t=[0.0]*N
    for p in projects:
        pr,it,_ = bond_profile(p['企業債'], RATE, p['redeem'], p['grace'], N)
        for i in range(N): pr_t[i]+=pr[i]; in_t[i]+=it[i]
    am = round(ASSET_AVG*0.03)
    base = sum(MAINT)+sum(DEP_B)+sum(INT_OLD)+sum(INT_NEW)+am*5+sum(KOUJU)+sum(HOJO)
    d_dep = sum(dep_t[i]-old_dep[i] for i in range(5))
    d_int = sum(in_t[i]-old_int[i] for i in range(5))
    d_ch  = sum(ch[i]-old_chouki[i] for i in range(5))
    after = base + d_dep + d_int - d_ch
    g = after*1000/YUSHU_5Y
    short = (after-REV_5Y)/5
    inv10 = sum(sum(p['cost'][i] for p in projects) for i in range(N))
    print(f'  {label:38s}{inv10:>13,.0f}{after:>14,.0f}{g:>11.1f}{short:>14,.0f}{after/REV_5Y:>10.2f}倍')
    if extra_note: print(f'  {"":38s}{extra_note}')

print('\n' + '='*156)
print('【5】感応度分析（R8〜R12・資産維持率3%）')
print('='*156)
print(f'  {"シナリオ":38s}{"建設改良費":>13s}{"総括原価5年計":>14s}{"給水原価":>11s}{"年間不足額":>14s}{"必要改定率":>12s}')
print(f'  {"":38s}{"R8〜R17":>13s}{"(千円)":>14s}{"(円/㎥)":>11s}{"(千円/年)":>14s}')
print('  '+'-'*152)

# 基準：旧計画
am = round(ASSET_AVG*0.03)
base = sum(MAINT)+sum(DEP_B)+sum(INT_OLD)+sum(INT_NEW)+am*5+sum(KOUJU)+sum(HOJO)
print(f'  {"① 旧計画（現行の財政シミュ）":38s}{sum(OLD["整備費"]):>13,.0f}{base:>14,.0f}'
      f'{base*1000/YUSHU_5Y:>11.1f}{(base-REV_5Y)/5:>14,.0f}{base/REV_5Y:>10.2f}倍')

import copy
scenario('② 新調査票（調査票の入力どおり）', PROJECTS)

# 感応度A：鷲神高度処理設備の償却年数を20年に
pa = copy.deepcopy(PROJECTS); pa[2]['dep']=20
scenario('③ ②＋鷲神高度処理を20年償却', pa,
         '※機械・電気設備の法定耐用年数は15〜20年。調査票の入力値40年は建物・構築物の年数')

# 感応度B：北上川導水管が調査票から漏れている場合（旧計画のR9 原水浄水施設費 670,000 を追加）
pb = copy.deepcopy(PROJECTS)
pb.append({'no':99,'name':'北上川導水管（調査票に未計上・旧計画から復元）','asset':'その他','dep':40,'redeem':40,'grace':3,
           'cost':[0,670_000,0,0,0,0,0,0,0,0], '国費':[0]*10,
           '企業債':[0,502_500,0,0,0,0,0,0,0,0], '繰入金':[0]*10,
           'その他':[0,167_500,0,0,0,0,0,0,0,0]})
scenario('④ ②＋北上川導水管6.7億円を復元', pb,
         '※旧計画R9の原水及び浄水施設費670,000千円。企業債75%と仮定')

pc = copy.deepcopy(pb); pc[2]['dep']=20
scenario('⑤ ③＋④（両方を織り込んだ場合）', pc)

print()
print('  ▶ ②を採る場合、必要改定幅は旧計画より小さくなる（給水原価 591.3→569.5円/㎥）。')
print('  ▶ ただし③④のいずれか一方でも該当すると、必要改定幅は逆に旧計画以上となる。')
print('  ▶ したがって、償却年数と北上川導水管の取扱いを女川町に確認するまで、改定幅は確定できない。')
