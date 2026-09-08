# -*- coding: utf-8 -*-
"""前提条件切替式 総括原価計算表 の作成（作業用ファイル）"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl
from openpyxl.worksheet.datavalidation import DataValidation

OUT = os.path.join(os.path.dirname(__file__), '..', '09_assumption_switch_cost_calc.xlsx')
YRS = ['R8','R9','R10','R11','R12','R13','R14','R15','R16','R17']; N = 10

NAVY,BLUE,RED,ORANGE,GREEN,YELLOW,GREY = '1F3864','2E75B6','C00000','C55A11','375623','FFF2CC','F2F2F2'
THIN = Side(border_style='thin', color='BFBFBF'); BD = Border(THIN,THIN,THIN,THIN)
def F(**k): return Font(name='Arial', **k)

# ===================== 計画データ =====================
OLD_PLAN = [dict(name='経営戦略の建設改良費', kind='管路',
    cost=[1_030_000,950_000,500_000,330_000,310_000,311_000,371_000,294_000,230_000,200_000],
    kokuhi=[66_330,53_000,53_000,53_000,53_000,3_000,3_000,3_000,3_000,3_000],
    bond=[448_670,402_000,387_000,237_000,237_000,287_000,347_000,267_000,187_000,157_000],
    rd_s=(40,0), rd_q=(40,3))]

SURVEY = [
 dict(name='No.1 老朽管布設替工事', kind='管路',
      cost=[180_000,150_000,150_000,100_000,100_000,100_000,100_000,100_000,100_000,100_000],
      kokuhi=[80_000,80_000,80_000,0,0,0,0,0,0,0],
      bond=[100_000,70_000,70_000,100_000,100_000,100_000,100_000,100_000,100_000,100_000],
      rd_s=(40,0), rd_q=(25,3)),
 dict(name='No.2 水道施設耐震化工事', kind='管路',
      cost=[12_000,35_000,60_000,220_000,220_000,220_000,60_000,150_000,150_000,150_000],
      kokuhi=[6_000,17_500,15_000,55_000,55_000,55_000,15_000,37_500,37_500,37_500],
      bond=[0,0,45_000,165_000,165_000,165_000,45_000,112_500,112_500,112_500],
      rd_s=(40,0), rd_q=(40,3)),
 dict(name='No.3 鷲神高度処理設備整備工事', kind='機電',
      cost=[353_400,534_600,0,0,0,0,0,0,0,0], kokuhi=[0]*10,
      bond=[176_700,267_300,0,0,0,0,0,0,0,0], rd_s=(40,0), rd_q=(40,3)),
 dict(name='No.4 江島海底送水管本復旧工事', kind='管路',
      cost=[20_000,150_000,0,0,0,0,0,0,0,0], kokuhi=[0]*10,
      bond=[15_000,112_500,0,0,0,0,0,0,0,0], rd_s=(40,0), rd_q=(40,3)),
]
# 経営戦略の更新投資額の内訳（【成行】財政シミュ r119〜r124）では「取・導水管」は全年度ゼロであり、
# R9の670,000千円は「構造物及び設備」。北上川導水管は経営戦略の投資計画に含まれていない。
# 一方、経営戦略には量水器費106,000千円（R8〜R17）があるが、調査票は委託・工事のみを対象と
# しているため計上されていない。第3の計画はこれを戻した場合。
METER = dict(name='量水器費（調査票の対象外・経営戦略から復元）', kind='管路',
      cost=[20_000,20_000,20_000,0,0,1_000,1_000,4_000,20_000,20_000], kokuhi=[0]*10,
      bond=[20_000,20_000,20_000,0,0,1_000,1_000,4_000,20_000,20_000], 繰入金=[0]*10,
      その他=[0]*10, rd_s=(40,0), rd_q=(40,3))

PLANS = {'経営戦略の計画': OLD_PLAN, '建設改良費調査票': SURVEY, '調査票＋量水器費': SURVEY+[METER]}

# ===================== 計算 =====================
def annuity(p, r, y): return p*r/(1-(1+r)**-y)

def bonds(draws, rate, years, grace):
    prin=[0.0]*N; intr=[0.0]*N
    for s,d in enumerate(draws):
        if d<=0: continue
        pay=annuity(d,rate,years); out=d
        for t in range(s+1,N):
            i=out*rate
            if t-s<=grace: intr[t]+=i
            else:
                pr=min(pay-i,out); out-=pr; prin[t]+=pr; intr[t]+=i
    return prin, intr

def sl(vals, life): return [sum(v/life for v in vals[:t]) for t in range(N)]

PIPE_LIVES=[35,38,40]; MECH_LIVES=[15,20,25,30,38,40]
RATES=[round(2.0+0.1*i,1) for i in range(11)]
DEP_METHODS={'損益計算書ベース':'A','固定資産台帳ベース':'B'}
REDEEM={'経営戦略方式（40年・据置なし）':'s','調査票方式（25/40年・据置3年）':'q'}

dep_rows=[]     # [計画, 償却方式, 管路年数, 機電年数, 減価償却5年計, 長期前受金(新規)5年計]
for pn, blocks in PLANS.items():
    for mname, m in DEP_METHODS.items():
        for pl in PIPE_LIVES:
            for ml in MECH_LIVES:
                dep=[0.0]*N; ch=[0.0]*N
                for b in blocks:
                    life = pl if b['kind']=='管路' else ml
                    basis = b['cost'] if m=='A' else [b['cost'][i]-b['kokuhi'][i] for i in range(N)]
                    for i,v in enumerate(sl(basis, life)): dep[i]+=v
                    if m=='A':
                        for i,v in enumerate(sl(b['kokuhi'], life)): ch[i]+=v
                dep_rows.append([pn, mname, pl, ml, sum(dep[:5]), sum(ch[:5])])

int_rows=[]     # [計画, 利率, 償還方式, 支払利息5年計, 元金償還5年計]
for pn, blocks in PLANS.items():
    for rt in RATES:
        for rname, rk in REDEEM.items():
            pr=[0.0]*N; it=[0.0]*N
            for b in blocks:
                y,g = b['rd_s'] if rk=='s' else b['rd_q']
                a,c = bonds(b['bond'], rt/100, y, g)
                for i in range(N): pr[i]+=a[i]; it[i]+=c[i]
            int_rows.append([pn, rt, rname, sum(it[:5]), sum(pr[:5])])

# 旧計画・経営戦略前提での新規取得分（既往資産分を切り出すために使用）
_base = next(r for r in dep_rows if r[0]=='経営戦略の計画' and r[1].startswith('損益') and r[2]==38 and r[3]==38)
OLD_NEW_DEP = _base[4]

# ===================== 固定値 =====================
FIX = {
 '維持管理費': 672_327,
 '減価償却費（財政シミュ R8〜R12 計上額）': 2_032_440,
 '支払利息（既往債）': 50_585,
 '支払利息（R7以前の新規債）': 44_630,
 '長期前受金戻入益（経営戦略・既往資産分）': 1_542_259,
 '長期前受金戻入益（成果品の推計値）': 450_000,
 '他会計補助金（経営戦略の見込額）': 272_997,
 '補助金（経営戦略の見込額）': 419_000,
 '他会計補助金等（成果品の推計値）': 15_000,
 'その他営業収益・受取利息・雑収益': 26_494,
 '雑支出': 36_696,
 '動力費・薬品費の単価補正額': 47_167,
 '固定資産台帳ベースの減価償却費（年額）': 32_188,
 '対象資産（期首 R8.4.1）': 1_042_625,
 '対象資産（期首期末平均）': 896_312,
 '有収水量 R8〜R12（㎥）': 4_316_000,
 '給水収益 R8〜R12': 590_845,
 '現行供給単価（円/㎥）': 128.2,
}

# ===================== 書式ヘルパ =====================
def title(ws, text, sub='', width=12):
    ws['A1']=text; ws['A1'].font=F(size=14,bold=True,color='FFFFFF')
    ws['A1'].fill=PatternFill('solid',fgColor=NAVY)
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=width)
    if sub:
        ws['A2']=sub; ws['A2'].font=F(size=9,color='404040')
        ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=width)
    ws.row_dimensions[1].height=24

def band(ws, row, text, width=12):
    c=ws.cell(row,1,text); c.font=F(size=11,bold=True,color='FFFFFF')
    c.fill=PatternFill('solid',fgColor=BLUE)
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=width)

def head(ws, row, labels, col=1):
    for i,t in enumerate(labels):
        c=ws.cell(row,col+i,t); c.font=F(size=9,bold=True,color='FFFFFF')
        c.fill=PatternFill('solid',fgColor=ORANGE); c.border=BD
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)

def inp(ws, coord):
    c=ws[coord]; c.fill=PatternFill('solid',fgColor=YELLOW)
    c.font=F(size=10,bold=True,color=RED); c.border=BD
    c.alignment=Alignment(horizontal='center')

wb=Workbook()

# ===================== 01_設定 =====================
ws=wb.active; ws.title='01_設定'
title(ws,'女川町水道事業　総括原価　前提条件の切替表',
      '黄色のセルが選択箇所です。経営戦略の前提と、今回の再試算の前提を項目ごとに切り替えられます。'
      '選択を変えると「02_総括原価」の計算結果が更新されます。')
for col,w in zip('ABCDEFGHIJ',[4,34,26,26,26,17,52,3,3,3]): ws.column_dimensions[col].width=w

r=4; band(ws,r,'【A】建設改良費と資本費の前提'); r+=1
head(ws,r,['','項目','選択（黄色セル）','経営戦略の設定','再試算の候補','','説明']); r+=1
ROWS_A=[
 ('A1','建設改良費','建設改良費調査票','経営戦略の計画','調査票 ／ 調査票＋量水器費',
  'R8〜R17の整備費。経営戦略4,526,000千円、調査票3,515,000千円。'),
 ('A2','新規取得資産の耐用年数（管路・構築物）',38,'38年','35年 ／ 40年（調査票の入力値）',
  '経営戦略は機械・電気を区分せず一律38年。調査票は4事業とも40年。'),
 ('A3','新規取得資産の耐用年数（機械・電気）',38,'38年（区分なし）','15年 ／ 20年 ／ 25年 ／ 30年 ／ 40年',
  '鷲神高度処理設備が該当。経営戦略は38年、法定耐用年数はこれより短い。'),
 ('A4','企業債の借入利率（％）',2.1,'2.1％','2.0〜3.0％から選択',
  '調査票の利率欄は4事業とも未入力。経営戦略の加重平均借入利率は2.1％。'),
 ('A5','企業債の償還条件','経営戦略方式（40年・据置なし）','40年・元利均等・借入翌年度から','調査票方式（25/40年・据置3年）',
  '調査票はNo.1が25年、他は40年、据置3年。'),
]
for k,name,sel,stg,alt,note in ROWS_A:
    ws.cell(r,1,k).font=F(size=9,color='808080'); ws.cell(r,2,name).font=F(size=10)
    ws.cell(r,3,sel); inp(ws,f'C{r}')
    ws.cell(r,4,stg).font=F(size=9,color='404040'); ws.cell(r,5,alt).font=F(size=9,color='404040')
    ws.cell(r,7,note).font=F(size=9,color='595959')
    for c in (2,4,5,7): ws.cell(r,c).border=BD; ws.cell(r,c).alignment=Alignment(vertical='center',wrap_text=True)
    r+=1
ROW_A={k:6+i for i,(k,*_ ) in enumerate(ROWS_A)}

r+=1; band(ws,r,'【B】資産維持費の前提'); r+=1
head(ws,r,['','項目','選択（黄色セル）','経営戦略の設定','再試算の候補','','説明']); r+=1
ROWS_B=[
 ('B1','資産維持率（％）',3.0,'1.0％','3.0％（業務仕様書の正式ケース）',
  '経営戦略は「従前それほど大きな修繕がないため1.0％と仮定」。仕様書は3％の算入を求めている。'),
 ('B2','対象資産の取り方','期首期末平均','期首の帳簿価額 1,042,625千円','期首期末平均 896,312千円',
  '経営戦略は期首のみ。日本水道協会の算定要領は期首期末平均。'),
 ('B3','資産維持費の年数',5,'4年','5年（R8〜R12）',
  '経営戦略の料金算定期間は4年。本件の算定期間はR8〜R12の5年。'),
]
for k,name,sel,stg,alt,note in ROWS_B:
    ws.cell(r,1,k).font=F(size=9,color='808080'); ws.cell(r,2,name).font=F(size=10)
    ws.cell(r,3,sel); inp(ws,f'C{r}')
    ws.cell(r,4,stg).font=F(size=9,color='404040'); ws.cell(r,5,alt).font=F(size=9,color='404040')
    ws.cell(r,7,note).font=F(size=9,color='595959')
    for c in (2,4,5,7): ws.cell(r,c).border=BD; ws.cell(r,c).alignment=Alignment(vertical='center',wrap_text=True)
    r+=1
ROW_B={k:r-len(ROWS_B)+i for i,(k,*_ ) in enumerate(ROWS_B)}

r+=1; band(ws,r,'【C】減価償却費と控除項目の前提　★総括原価への影響が最も大きい区分'); r+=1
head(ws,r,['','項目','選択（黄色セル）','経営戦略の設定','再試算の候補','','説明']); r+=1
ROWS_C=[
 ('C1','減価償却費の算入方法','損益計算書ベース','固定資産台帳ベース','損益計算書ベース',
  '経営戦略は補助財源で取得した資産の償却費を原価に入れず、長期前受金戻入も控除しない。'
  '成果品は償却費を全額入れ、長期前受金戻入を控除する。両者は本来同じ結果になるべきもの。'),
 ('C2','長期前受金戻入益の控除額','成果品の推計値','実額（1,542,259千円・既往資産分）','実額 ／ 推計値 ／ 控除しない',
  '★経営戦略の実額は成果品の推計値の3.4倍。震災復興交付金で取得した資産が多いため。'
  '「減価償却費の算入方法」で固定資産台帳ベースを選ぶとこの項目は自動的に無効になる。'),
 ('C3','他会計補助金等の控除額','控除しない','他会計補助金 272,997千円','他会計補助金＋補助金 691,997千円 ／ 推計値 ／ 控除しない',
  '★基準内繰入金は控除、基準外繰入金（赤字補填）は控除しない。経営戦略報告書案3-8⑹は「国が定める繰出基準以外の'
  '繰出金によって収入不足を補填している」と述べ、4-1⑷も繰入金を「現金部分の赤字補填」で算定としているため、'
  '全額を基準外とみて控除しないのが本文に整合する。'),
 ('C4','その他営業収益・受取利息・雑収益の控除','控除する','控除する（26,494千円）','控除しない',
  '算定要領では営業外収益として控除対象。成果品では控除されていない。'),
 ('C5','雑支出の算入','算入する','算入していない','算入する（36,696千円）／ 算入しない',
  '営業外費用。経営戦略の料金算定でも成果品でも算入されていない。'),
 ('C6','動力費・薬品費の単価','実績単価に補正','経営戦略の値のまま','実績単価に補正（＋47,167千円）',
  '★経営戦略の単価算定は、1年分の費用を10年分の有収水量で割っており、実績の1/9〜1/14の水準。'),
]
for k,name,sel,stg,alt,note in ROWS_C:
    ws.cell(r,1,k).font=F(size=9,color='808080'); ws.cell(r,2,name).font=F(size=10)
    ws.cell(r,3,sel); inp(ws,f'C{r}')
    ws.cell(r,4,stg).font=F(size=9,color='404040'); ws.cell(r,5,alt).font=F(size=9,color='404040')
    ws.cell(r,7,note).font=F(size=9,color='595959')
    for c in (2,4,5,7): ws.cell(r,c).border=BD; ws.cell(r,c).alignment=Alignment(vertical='center',wrap_text=True)
    ws.row_dimensions[r].height=30
    r+=1
ROW_C={k:r-len(ROWS_C)+i for i,(k,*_ ) in enumerate(ROWS_C)}

r+=1
ws.cell(r,2,'▼ 一括設定の目安').font=F(size=10,bold=True,color=NAVY)
r+=1
for t in ['「経営戦略に合わせる」…A1=経営戦略の計画／A2=38／A3=38／A4=2.1／A5=経営戦略方式／B1=1.0／B2=期首の帳簿価額／B3=4／C1=固定資産台帳ベース／C3=他会計補助金 272,997千円／C4=控除する／C5=算入しない／C6=経営戦略の値のまま',
          '「今回の再試算」…A1=建設改良費調査票／A2=38／A3=38／A4=2.1／A5=経営戦略方式／B1=3.0／B2=期首期末平均／B3=5／C1=損益計算書ベース／C2=実額／C3=他会計補助金 272,997千円／C4=控除する／C5=算入する／C6=実績単価に補正']:
    ws.cell(r,2,t).font=F(size=9,color='595959')
    ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=7); r+=1

DVS=[
 (f'C{ROW_A["A1"]}','"経営戦略の計画,建設改良費調査票,調査票＋量水器費"'),
 (f'C{ROW_A["A2"]}','"35,38,40"'),
 (f'C{ROW_A["A3"]}','"15,20,25,30,38,40"'),
 (f'C{ROW_A["A4"]}','"2.0,2.1,2.2,2.3,2.4,2.5,2.6,2.7,2.8,2.9,3.0"'),
 (f'C{ROW_A["A5"]}','"経営戦略方式（40年・据置なし）,調査票方式（25/40年・据置3年）"'),
 (f'C{ROW_B["B1"]}','"1.0,1.5,2.0,2.5,3.0"'),
 (f'C{ROW_B["B2"]}','"期首の帳簿価額,期首期末平均"'),
 (f'C{ROW_B["B3"]}','"4,5"'),
 (f'C{ROW_C["C1"]}','"損益計算書ベース,固定資産台帳ベース"'),
 (f'C{ROW_C["C2"]}','"経営戦略の実額,成果品の推計値,控除しない"'),
 (f'C{ROW_C["C3"]}','"他会計補助金のみ,他会計補助金＋補助金,成果品の推計値,控除しない"'),
 (f'C{ROW_C["C4"]}','"控除する,控除しない"'),
 (f'C{ROW_C["C5"]}','"算入する,算入しない"'),
 (f'C{ROW_C["C6"]}','"経営戦略の値のまま,実績単価に補正"'),
]
for coord, formula in DVS:
    dv=DataValidation(type='list', formula1=formula, allow_blank=False); ws.add_data_validation(dv); dv.add(ws[coord])
for coord, val in [(f'C{ROW_A["A1"]}','経営戦略の計画'),
                   (f'C{ROW_C["C2"]}','成果品の推計値'),
                   (f'C{ROW_C["C3"]}','成果品の推計値'),
                   (f'C{ROW_C["C4"]}','控除しない'),
                   (f'C{ROW_C["C5"]}','算入しない'),
                   (f'C{ROW_C["C6"]}','経営戦略の値のまま')]:
    ws[coord]=val
n=ROW_A['A1']-3
c=ws.cell(n,2,'初期値は現在の成果品（01_rate_reform_simulation_MAIN.xlsx）の前提を再現しています。'
              'この状態で総括原価は2,552,199千円、給水原価591.3円/㎥、必要改定率4.32倍になります。')
c.font=F(size=10,bold=True,color=RED)
ws.merge_cells(start_row=n,start_column=2,end_row=n,end_column=7)

S='01_設定'
A1=f"'{S}'!$C${ROW_A['A1']}"; A2=f"'{S}'!$C${ROW_A['A2']}"; A3=f"'{S}'!$C${ROW_A['A3']}"
A4=f"'{S}'!$C${ROW_A['A4']}"; A5=f"'{S}'!$C${ROW_A['A5']}"
B1=f"'{S}'!$C${ROW_B['B1']}"; B2=f"'{S}'!$C${ROW_B['B2']}"; B3=f"'{S}'!$C${ROW_B['B3']}"
C1=f"'{S}'!$C${ROW_C['C1']}"; C2=f"'{S}'!$C${ROW_C['C2']}"; C3=f"'{S}'!$C${ROW_C['C3']}"
C4=f"'{S}'!$C${ROW_C['C4']}"; C5=f"'{S}'!$C${ROW_C['C5']}"; C6=f"'{S}'!$C${ROW_C['C6']}"

# ===================== 05_計算用データ =====================
wd=wb.create_sheet('05_計算用データ')
title(wd,'計算用データ（参照専用）','「01_設定」の選択に応じて「02_総括原価」が参照します。直接編集しないでください。',9)
for col,w in zip('ABCDEFGHI',[24,34,12,12,18,18,14,20,18]): wd.column_dimensions[col].width=w
wd['A4']='【減価償却費・長期前受金（新規取得分）R8〜R12 5年計・千円】'; wd['A4'].font=F(bold=True,color=NAVY)
head(wd,5,['計画','減価償却費の算入方法','管路年数','機電年数','減価償却費','長期前受金(新規)','照合キー'])
rr=6
for pn,mn,pl,ml,dp,ch in dep_rows:
    for i,v in enumerate([pn,mn,pl,ml,round(dp),round(ch)]):
        c=wd.cell(rr,i+1,v); c.font=F(size=9); c.border=BD
        if i>=4: c.number_format='#,##0'
    wd.cell(rr,7,f'{pn}|{mn}|{pl}|{ml}').font=F(size=8,color='808080'); rr+=1
DEP_TOP, DEP_END = 6, rr-1

r2=rr+2
wd.cell(r2,1,'【支払利息・元金償還（新規債）R8〜R12 5年計・千円】').font=F(bold=True,color=NAVY)
head(wd,r2+1,['計画','借入利率(％)','償還条件','支払利息','元金償還','','照合キー'])
rr=r2+2
for pn,rt,rn,it,pr in int_rows:
    for i,v in enumerate([pn,rt,rn,round(it),round(pr)]):
        c=wd.cell(rr,i+1,v); c.font=F(size=9); c.border=BD
        if i>=3: c.number_format='#,##0'
    wd.cell(rr,7,f'{pn}|{rt}|{rn}').font=F(size=8,color='808080'); rr+=1
INT_TOP, INT_END = r2+2, rr-1

D=lambda col: f"INDEX('05_計算用データ'!${col}${DEP_TOP}:${col}${DEP_END},MATCH({A1}&\"|\"&{C1}&\"|\"&{A2}&\"|\"&{A3},'05_計算用データ'!$G${DEP_TOP}:$G${DEP_END},0))"
I=lambda col: f"INDEX('05_計算用データ'!${col}${INT_TOP}:${col}${INT_END},MATCH({A1}&\"|\"&TEXT({A4},\"0.0\")&\"|\"&{A5},'05_計算用データ'!$G${INT_TOP}:$G${INT_END},0))"

# ===================== 02_総括原価 =====================
wc=wb.create_sheet('02_総括原価',1)
title(wc,'女川町水道事業　総括原価の算定　R8〜R12（5年間）',
      '「01_設定」の選択に基づいて計算しています。金額はすべて千円。',9)
for col,w in zip('ABCDEFGHI',[4,40,20,20,14,3,64,3,3]): wc.column_dimensions[col].width=w

r=4; band(wc,r,'【1】適用中の前提条件',9); r+=1
for lbl, ref in [('建設改良費',A1),('耐用年数（管路・構築物）',A2),('耐用年数（機械・電気）',A3),
                 ('借入利率（％）',A4),('償還条件',A5),('資産維持率（％）',B1),('対象資産の取り方',B2),
                 ('資産維持費の年数',B3),('減価償却費の算入方法',C1),('長期前受金戻入益の控除',C2),
                 ('他会計補助金等の控除',C3),('その他収益の控除',C4),('雑支出',C5),('動力費・薬品費',C6)]:
    wc.cell(r,2,lbl).font=F(size=9,color='404040'); wc.cell(r,2).border=BD
    c=wc.cell(r,3,f'={ref}'); c.font=F(size=9,bold=True,color=NAVY); c.border=BD
    wc.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4); r+=1

r+=1; band(wc,r,'【2】総括原価の算定',9); R0=r; r+=1
head(wc,r,['','費目','金額（千円）','','','','算定の内容']); r+=1
def line(row, no, name, formula, note, bold=False, color=None):
    wc.cell(row,1,no).font=F(size=9,color='808080')
    c=wc.cell(row,2,name); c.font=F(size=10,bold=bold,color=color or '000000'); c.border=BD
    v=wc.cell(row,3,formula); v.font=F(size=10,bold=bold,color=color or '000000')
    v.number_format='#,##0'; v.border=BD
    n=wc.cell(row,7,note); n.font=F(size=9,color='595959'); n.alignment=Alignment(wrap_text=True,vertical='center')
    wc.row_dimensions[row].height=26

RM=FIX['維持管理費']; RD=FIX['減価償却費（財政シミュ R8〜R12 計上額）']
r_m=r; line(r,'Ⅰ','維持管理費',
  f'={RM}+IF({C6}="実績単価に補正",{FIX["動力費・薬品費の単価補正額"]},0)',
  '財政シミュの営業費用−減価償却費。動力費・薬品費の単価をR1〜R5実績で置き換えると＋47,167千円。'); r+=1
r_d=r; line(r,'Ⅱ','減価償却費',
  f'=IF({C1}="固定資産台帳ベース",{FIX["固定資産台帳ベースの減価償却費（年額）"]}*5+{D("E")},'
  f'{RD}-{round(OLD_NEW_DEP)}+{D("E")})',
  '損益計算書ベースは財政シミュ計上額から旧計画の新規取得分を除き、選択した計画の償却費を加算。'
  '固定資産台帳ベースは既往資産の年32,188千円×5年に、補助財源を除いた新規取得分を加算。'); r+=1
r_i1=r; line(r,'Ⅲ','支払利息（既往債）', f'={FIX["支払利息（既往債）"]}', '財政シミュの既往債利息。'); r+=1
r_i2=r; line(r,'Ⅳ','支払利息（新規債）', f'={FIX["支払利息（R7以前の新規債）"]}+{I("D")}',
  'R7以前の借入分44,630千円は固定。R8以降の借入分を選択した計画・借入利率・償還条件で算定。'); r+=1
r_z=r; line(r,'Ⅴ','雑支出', f'=IF({C5}="算入する",{FIX["雑支出"]},0)', '営業外費用。'); r+=1
r_am=r; line(r,'Ⅵ','資産維持費',
  f'=ROUND(IF({B2}="期首の帳簿価額",{FIX["対象資産（期首 R8.4.1）"]},{FIX["対象資産（期首期末平均）"]})*{B1}/100,0)*{B3}',
  '対象資産×資産維持率×年数。'); r+=1
r_c1=r; line(r,'Ⅶ','控除：長期前受金戻入益',
  f'=-IF({C1}="固定資産台帳ベース",0,'
  f'IF({C2}="経営戦略の実額",{FIX["長期前受金戻入益（経営戦略・既往資産分）"]}+{D("F")},'
  f'IF({C2}="成果品の推計値",{FIX["長期前受金戻入益（成果品の推計値）"]},0)))',
  '固定資産台帳ベースを選んだ場合は二重控除になるため自動的にゼロ。'); r+=1
r_c2=r; line(r,'Ⅷ','控除：他会計補助金等',
  f'=-IF({C3}="他会計補助金のみ",{FIX["他会計補助金（経営戦略の見込額）"]},'
  f'IF({C3}="他会計補助金＋補助金",{FIX["他会計補助金（経営戦略の見込額）"]+FIX["補助金（経営戦略の見込額）"]},'
  f'IF({C3}="成果品の推計値",{FIX["他会計補助金等（成果品の推計値）"]},0)))',
  '基準内繰入金は控除、基準外繰入金（赤字補填）は控除しないのが原則。'); r+=1
r_c3=r; line(r,'Ⅸ','控除：その他営業収益・受取利息・雑収益',
  f'=-IF({C4}="控除する",{FIX["その他営業収益・受取利息・雑収益"]},0)', '算定要領上の控除項目。'); r+=1
r_t=r; line(r,'','総括原価　合計', f'=SUM(C{r_m}:C{r_c3})', '', bold=True, color=NAVY)
wc.cell(r,3).fill=PatternFill('solid',fgColor='DDEBF7'); r+=2

band(wc,r,'【3】給水原価・料金回収率・必要改定幅',9); r+=1
head(wc,r,['','指標','算定値','','','','算定式']); r+=1
YU=FIX['有収水量 R8〜R12（㎥）']; RV=FIX['給水収益 R8〜R12']; UP=FIX['現行供給単価（円/㎥）']
def mline(row,name,formula,fmt,note):
    c=wc.cell(row,2,name); c.font=F(size=10); c.border=BD
    v=wc.cell(row,3,formula); v.font=F(size=11,bold=True,color=NAVY); v.number_format=fmt; v.border=BD
    n=wc.cell(row,7,note); n.font=F(size=9,color='595959')
mline(r,'有収水量 R8〜R12（㎥）',f'={YU}','#,##0','財政シミュの水需要予測。'); r+=1
mline(r,'給水収益 R8〜R12（千円・現行料金）',f'={RV}','#,##0','財政シミュの給水収益。'); r+=1
r_g=r; mline(r,'給水原価（円/㎥）',f'=C{r_t}*1000/{YU}','#,##0.0','総括原価×1,000÷有収水量。'); r+=1
mline(r,'現行供給単価（円/㎥）',f'={UP}','#,##0.0','R7実績 131,429千円÷1,025千㎥。'); r+=1
mline(r,'料金回収率',f'={UP}/C{r_g}','0.0%','現行供給単価÷給水原価。'); r+=1
r_s=r; mline(r,'年間不足額（千円/年）',f'=(C{r_t}-{RV})/5','#,##0','（総括原価−給水収益）÷5年。'); r+=1
mline(r,'必要改定率（総括原価÷現行収入）',f'=C{r_t}/{RV}','#,##0.00','1.00倍なら現行料金で原価を回収できる水準。'); r+=1
mline(r,'標準世帯（20㎥/月）月額への換算（円）',f'=ROUND(2395*C{r_t}/{RV},0)','#,##0','現行2,395円×必要改定率。参考値。'); r+=2
wc.cell(r,2,'※ この表は開いたときに再計算されます。数値が空欄に見える場合は、Excelで開き直してください。').font=F(size=9,color=RED)

# ===================== 評価関数（06_主要ケース用） =====================
def look_dep(plan, meth, pl, ml):
    for a,b,c,d,e,f in dep_rows:
        if a==plan and b==meth and c==pl and d==ml: return e,f
    raise KeyError
def look_int(plan, rate, rname):
    for a,b,c,d,e in int_rows:
        if a==plan and abs(b-rate)<1e-9 and c==rname: return d,e
    raise KeyError

def evaluate(plan, pipe, mech, rate, redeem, am_rate, am_base, am_yrs,
             dep_meth, chouki, hojo, other, zatsu, power):
    dep_new, ch_new = look_dep(plan, dep_meth, pipe, mech)
    it_new, _ = look_int(plan, rate, redeem)
    maint = FIX['維持管理費'] + (FIX['動力費・薬品費の単価補正額'] if power=='補正' else 0)
    if dep_meth.startswith('固定資産台帳'):
        dep = FIX['固定資産台帳ベースの減価償却費（年額）']*5 + dep_new
    else:
        dep = FIX['減価償却費（財政シミュ R8〜R12 計上額）'] - OLD_NEW_DEP + dep_new
    asset = FIX['対象資産（期首 R8.4.1）'] if am_base=='期首' else FIX['対象資産（期首期末平均）']
    am = round(asset*am_rate/100)*am_yrs
    if dep_meth.startswith('固定資産台帳'): c_ch = 0
    elif chouki=='実額': c_ch = FIX['長期前受金戻入益（経営戦略・既往資産分）'] + ch_new
    elif chouki=='推計': c_ch = FIX['長期前受金戻入益（成果品の推計値）']
    else: c_ch = 0
    c_h = {'他会計補助金':FIX['他会計補助金（経営戦略の見込額）'],
           '他会計補助金＋補助金':FIX['他会計補助金（経営戦略の見込額）']+FIX['補助金（経営戦略の見込額）'],
           '推計':FIX['他会計補助金等（成果品の推計値）'], 'なし':0}[hojo]
    c_o = FIX['その他営業収益・受取利息・雑収益'] if other=='控除する' else 0
    zt  = FIX['雑支出'] if zatsu=='算入する' else 0
    total = (maint + dep + FIX['支払利息（既往債）'] + FIX['支払利息（R7以前の新規債）'] + it_new
             + zt + am - c_ch - c_h - c_o)
    g = total*1000/FIX['有収水量 R8〜R12（㎥）']
    return dict(total=total, g=g, rec=FIX['現行供給単価（円/㎥）']/g,
                short=(total-FIX['給水収益 R8〜R12'])/5, ratio=total/FIX['給水収益 R8〜R12'],
                maint=maint, dep=dep, it=it_new, am=am, c_ch=c_ch, c_h=c_h, c_o=c_o, zt=zt)

STG = dict(plan='経営戦略の計画', pipe=38, mech=38, rate=2.1, redeem='経営戦略方式（40年・据置なし）',
           am_rate=1.0, am_base='期首', am_yrs=4, dep_meth='固定資産台帳ベース',
           chouki='実額', hojo='他会計補助金', other='控除する', zatsu='算入しない', power='戦略のまま')
RE  = dict(plan='建設改良費調査票', pipe=38, mech=38, rate=2.1, redeem='経営戦略方式（40年・据置なし）',
           am_rate=3.0, am_base='平均', am_yrs=5, dep_meth='損益計算書ベース',
           chouki='実額', hojo='他会計補助金', other='控除する', zatsu='算入する', power='補正')
CUR = dict(RE, chouki='推計', hojo='推計', other='控除しない', zatsu='算入しない', power='戦略のまま')

REC = dict(RE, hojo='なし')   # 繰入金は基準外とみて控除しない（報告書案の記載に整合）

CASES = [
 ('① 経営戦略の前提をそのまま使う', STG),
 ('② 成果品（現行）の前提', dict(CUR, plan='経営戦略の計画')),
 ('③ ②に建設改良費調査票だけ反映', dict(CUR, plan='建設改良費調査票')),
 ('④ ③＋長期前受金戻入を実額に', dict(CUR, plan='建設改良費調査票', chouki='実額')),
 ('⑤ ④＋動力費・薬品費と雑支出を補正（繰入金は控除しない）', REC),
 ('⑥ ⑤＋他会計補助金272,997を控除（基準内とみる場合）', dict(REC, hojo='他会計補助金')),
 ('⑦ ⑤＋他会計補助金＋補助金691,997を控除', dict(REC, hojo='他会計補助金＋補助金')),
 ('⑧ ⑤＋機械・電気を20年償却', dict(REC, mech=20)),
 ('⑨ ⑤＋量水器費を戻す', dict(REC, plan='調査票＋量水器費')),
 ('⑩ ⑤＋借入利率3.0％', dict(REC, rate=3.0)),
 ('⑪ ⑤で減価償却を固定資産台帳ベースに', dict(REC, dep_meth='固定資産台帳ベース')),
]

# ===================== 06_主要ケース =====================
wk=wb.create_sheet('06_主要ケース')
title(wk,'主要ケースの計算結果一覧','「01_設定」で選択できる組み合わせのうち、代表的なものを並べたものです。金額は千円。',10)
for col,w in zip('ABCDEFGHIJ',[4,40,17,15,13,13,13,13,13,13]): wk.column_dimensions[col].width=w
r=4; head(wk,r,['','ケース','総括原価5年計','給水原価(円/㎥)','料金回収率','年間不足額','必要改定率','維持管理費','減価償却費','控除計']); r+=1
for nm, kw in CASES:
    x=evaluate(**kw)
    wk.cell(r,2,nm).font=F(size=10,bold=nm.startswith(('①','②','⑤')))
    for i,(v,fmt) in enumerate([(x['total'],'#,##0'),(x['g'],'#,##0.0'),(x['rec'],'0.0%'),
                                (x['short'],'#,##0'),(x['ratio'],'#,##0.00'),
                                (x['maint'],'#,##0'),(x['dep'],'#,##0'),
                                (-(x['c_ch']+x['c_h']+x['c_o']),'#,##0')]):
        c=wk.cell(r,3+i,round(v,1) if isinstance(v,float) else v); c.number_format=fmt
        c.font=F(size=10); c.border=BD
    wk.cell(r,2).border=BD; r+=1
r+=1
for t in ['▼ 読み取り方',
  '・①と②の差が、経営戦略と成果品の前提の違いによるものです。建設改良費の改定より、控除項目の扱いの方が影響が大きくなっています。',
  '・②→④で給水原価が大きく下がるのは、長期前受金戻入益を推計値450,000千円から実額1,542,259千円に置き換えたためです。',
  '・⑤は繰入金を控除しない前提です。経営戦略報告書案が「繰出基準以外の繰出金で補填している」と述べているためで、'
  '基準内と認められる分があれば⑥⑦の水準まで下がります。',
  '・⑪は減価償却費を経営戦略と同じ固定資産台帳ベースに揃えたもので、①に近づきます。',
  '・⑧⑨⑩は、確定していない前提（機械電気の耐用年数・量水器費・借入利率）が振れた場合の幅を示しています。']:
    wk.cell(r,2,t).font=F(size=9,bold=t.startswith('▼'),color=NAVY if t.startswith('▼') else '595959')
    wk.merge_cells(start_row=r,start_column=2,end_row=r,end_column=10); r+=1

# ===================== 03_差異一覧 =====================
wv=wb.create_sheet('03_差異一覧',2)
title(wv,'経営戦略と再試算の差異一覧','双方の設定が異なる箇所と、確認が必要な箇所をまとめたものです。金額は千円。',8)
for col,w in zip('ABCDEFGH',[4,30,30,30,16,3,58,3]): wv.column_dimensions[col].width=w
r=4; band(wv,r,'【A】前提条件の差異（「01_設定」で切り替えられる項目）',8); r+=1
head(wv,r,['','項目','経営戦略','再試算（成果品）','影響の大きさ','','内容']); r+=1
DIFF_A=[
 ('料金算定期間','R6〜R9の4年間','R8〜R12の5年間','大',
  '経営戦略の期間原価集計は年度ラベルがH31・R2・R3・R4のままだが、参照先の数値はR6〜R9。ラベルの更新漏れ。'),
 ('減価償却費の算入方法','固定資産台帳ベース 128,750（4年）','損益計算書ベース 2,032,440（5年）','特大',
  '経営戦略は補助財源で取得した資産の償却費を原価に算入していない。成果品は全額算入し長期前受金戻入を控除する方式。'),
 ('長期前受金戻入益の控除','実額 1,542,259（R8〜R12・既往資産分）','推計 450,000','特大',
  '成果品の推計値は実額の約29％。震災復興交付金で取得した資産が多く、償却費の8割超が戻入で相殺される構造。'),
 ('他会計補助金等の控除','他会計補助金 272,997（R8〜R12）','推計 15,000','大',
  '経営戦略の算定式は「減価償却費＋人件費＋支払利息」。基準内・基準外の区分が未確定。'),
 ('補助金の控除','419,000（R8〜R12・営業外収益）','算入なし','大','性格の確認が必要。'),
 ('その他営業収益・受取利息・雑収益','控除する 26,494','控除していない','小','算定要領上の控除項目。'),
 ('雑支出','算入していない 36,696','算入していない','小','営業外費用。双方とも未算入。'),
 ('資産維持率','1.0％','3.0％','中','経営戦略は「従前それほど大きな修繕がない」ことを理由に1.0％と仮定。仕様書は3％を求めている。'),
 ('資産維持費の対象資産','期首 1,042,625','期首期末平均 896,312','中','算定要領は期首期末平均。'),
 ('資産維持費の年数','4年','5年','中','料金算定期間の設定に連動。'),
 ('新規取得資産の耐用年数','一律38年（機械・電気の区分なし）','調査票40年','中',
  '鷲神高度処理設備は機械・電気設備。法定耐用年数はこれより短い。'),
 ('企業債の借入利率','2.1％','調査票は未入力','小','経営戦略の根拠欄には「1.79％」と別の数値が残っている。'),
 ('企業債の償還条件','40年・元利均等・借入翌年度から','調査票は25／40年・据置3年','小',''),
 ('建設改良費','4,526,000（R8〜R17）','3,515,000（調査票）','中','差 1,011,000。施設区分別では構造物及び設備▲475,000、送水管▲230,000、配水本管▲200,000、量水器費▲106,000。'),
 ('動力費・薬品費の単価','1年分の費用を10年分の有収水量で除している','経営戦略の値をそのまま使用','中',
  '実績の1/9〜1/14の水準。年約7,200の過小計上。'),
 ('有収水量 R8〜R12','4,315,366㎥','4,316,000㎥','なし','丸めの差のみ。'),
 ('給水収益 R8〜R12','590,845','590,845','なし','一致。'),
 ('維持管理費 R8〜R12','672,328','672,327','なし','一致。'),
 ('支払利息（既往債）R8〜R12','50,585','50,585','なし','一致。'),
]
for i,(nm,stg,re_,imp,note) in enumerate(DIFF_A,1):
    wv.cell(r,1,i).font=F(size=9,color='808080')
    for cn,v in [(2,nm),(3,stg),(4,re_),(5,imp),(7,note)]:
        c=wv.cell(r,cn,v); c.font=F(size=9, bold=(cn==5 and imp in('特大','大')),
                                    color=(RED if cn==5 and imp in('特大','大') else '000000'))
        c.border=BD; c.alignment=Alignment(wrap_text=True,vertical='center')
    wv.row_dimensions[r].height=28; r+=1

r+=1; band(wv,r,'【B】経営戦略試算シートで確認が必要な箇所',8); r+=1
head(wv,r,['','箇所','内容','','','','対応']); r+=1
DIFF_B=[
 ('料金算定－要領 総括原価シート','固定費の配賦から料金表の算定まで、参照切れにより計算結果が出ない状態。',
  '有収水量合計と口径別調定件数の参照先を復旧する。'),
 ('料金算定－現行×改定率 総括原価シート','口径別基本料金・水量料金の算定結果がすべて参照切れ。',
  '同上。'),
 ('原水及び浄水費の動力費','1年分の平均費用を10年分の有収水量合計で除して単価を算定しており、'
  'さらに単価を整数に丸めているため1円/㎥になっている。計上額は有収水量と同額。',
  '単価を1年あたりの有収水量で算定し直す。実績ベースでは約4.7円/㎥。'),
 ('原水及び浄水費の薬品費','同じ理由で0.258円/㎥。実績ベースでは約2.3円/㎥。','同上。'),
 ('配水及び給水費の動力費','同じ理由で単価が0.25円/㎥となり、整数への丸めで0円になっている。'
  '実績では年2,159千円が発生している。','同上。丸めの位置も見直す。'),
 ('期間原価集計の年度ラベル','K4〜N4が「H31・R2・R3・R4」のままだが、参照先の数値はR6〜R9。',
  'ラベルを実態に合わせる。'),
 ('支払利息の算定根拠欄','「利率の平均値（1.79％）」と記載されているが、実際の計算は2.1％。',
  '根拠欄を2.1％に更新する。'),
 ('新規取得資産の耐用年数','機械・電気設備を区分せず一律38年。貸借対照表では「機械及び装置」を独立科目として持っている。',
  '将来取得分についても資産区分ごとの耐用年数を設定するか、一律とする理由を明記する。'),
]
for i,(pl,cont,act) in enumerate(DIFF_B,1):
    wv.cell(r,1,i).font=F(size=9,color='808080')
    c=wv.cell(r,2,pl); c.font=F(size=9,bold=True); c.border=BD; c.alignment=Alignment(wrap_text=True,vertical='center')
    c=wv.cell(r,3,cont); c.font=F(size=9); c.border=BD; c.alignment=Alignment(wrap_text=True,vertical='center')
    wv.merge_cells(start_row=r,start_column=3,end_row=r,end_column=5)
    c=wv.cell(r,7,act); c.font=F(size=9,color=GREEN); c.alignment=Alignment(wrap_text=True,vertical='center')
    wv.row_dimensions[r].height=34; r+=1

r+=1; band(wv,r,'【C】建設改良費調査票で確認が必要な箇所',8); r+=1
head(wv,r,['','箇所','内容','','','','対応']); r+=1
DIFF_C=[
 ('No.3 鷲神高度処理設備の財源構成割合','割合(%)欄に金額444,000が入力されているため、繰入金が約3.9兆円になる。'
  '総整備費は正しいままなので気づきにくい。','割合欄に50を入力する。'),
 ('企業債の利率','4事業とも未入力。','借入予定利率を記入する。'),
 ('No.4 江島海底送水管','資産種別・更新新規の別・補助単独の別がいずれも未入力。','記入する。'),
 ('量水器費','経営戦略はR8〜R17で106,000を計上しているが、調査票は委託・工事のみを対象としているため計上されていない。','量水器の更新費を別途見込む必要があるか確認する。'),
 ('出資金と繰入金','経営戦略は出資金1,230,000、調査票は繰入金444,000。差786,000。','一般会計負担の方針を確認する。'),
 ('対象期間','入力方法シートはR9〜R17だが、様式と回答はR8を含む。','R8分の網羅性を確認する。'),
 ('償却年数','4事業とも40年。鷲神高度処理設備は機械・電気設備。','資産区分に応じた年数か確認する。'),
]
for i,(pl,cont,act) in enumerate(DIFF_C,1):
    wv.cell(r,1,i).font=F(size=9,color='808080')
    c=wv.cell(r,2,pl); c.font=F(size=9,bold=True); c.border=BD; c.alignment=Alignment(wrap_text=True,vertical='center')
    c=wv.cell(r,3,cont); c.font=F(size=9); c.border=BD; c.alignment=Alignment(wrap_text=True,vertical='center')
    wv.merge_cells(start_row=r,start_column=3,end_row=r,end_column=5)
    c=wv.cell(r,7,act); c.font=F(size=9,color=GREEN); c.alignment=Alignment(wrap_text=True,vertical='center')
    wv.row_dimensions[r].height=30; r+=1

# ===================== 04_建設改良費 =====================
wp=wb.create_sheet('04_建設改良費',3)
title(wp,'建設改良費の新旧比較（R8〜R17）','金額は千円。',13)
for col in 'ABCDEFGHIJKLM': wp.column_dimensions[col].width=13
wp.column_dimensions['A'].width=34
r=4
for pname, blocks in PLANS.items():
    band(wp,r,f'【{pname}】',13); r+=1
    head(wp,r,['事業名']+YRS+['合計']); r+=1
    tot=[0]*N
    for b in blocks:
        wp.cell(r,1,b['name']).font=F(size=9); wp.cell(r,1).border=BD
        for i,v in enumerate(b['cost']):
            c=wp.cell(r,2+i,v); c.number_format='#,##0'; c.font=F(size=9); c.border=BD; tot[i]+=v
        c=wp.cell(r,12,sum(b['cost'])); c.number_format='#,##0'; c.font=F(size=9,bold=True); c.border=BD
        r+=1
    wp.cell(r,1,'計').font=F(size=9,bold=True,color=NAVY); wp.cell(r,1).border=BD
    for i,v in enumerate(tot):
        c=wp.cell(r,2+i,v); c.number_format='#,##0'; c.font=F(size=9,bold=True,color=NAVY); c.border=BD
        c.fill=PatternFill('solid',fgColor='DDEBF7')
    c=wp.cell(r,12,sum(tot)); c.number_format='#,##0'; c.font=F(size=10,bold=True,color=NAVY); c.border=BD
    c.fill=PatternFill('solid',fgColor='DDEBF7'); r+=3


# ===================== 07_参照元 =====================
wr=wb.create_sheet('07_参照元')
title(wr,'主要な数値の参照元','各数値がどのファイル・シート・行から来ているかの対応表です。',8)
for col,w in zip('ABCDEFGH',[4,30,34,30,3,3,60,3]): wr.column_dimensions[col].width=w
r=4; band(wr,r,'【A】控除項目',8); r+=1
head(wr,r,['','項目','参照元','値','','','確認結果']); r+=1
SRC_A=[
 ('長期前受金戻入益（既往資産分）','経営戦略試算シート【成行】財政シミュ r17',
  'R8 329,946／R9 324,870／R10 311,865／R11 299,986／R12 290,154（5年計1,542,259）',
  '全年度が直接入力の固定値。根拠欄は「固定資産台帳データに基づく算定額」。'
  '固定資産台帳そのものは同ブックに含まれていない。実績（【PL】決算数値入力 r26）はR3 336,382／R4 419,320／R5 336,409で、'
  '見込額はこの水準と整合している。'),
 ('長期前受金戻入益（新規取得分）','【成行】財政シミュ r18 ＝ 新規取得!P5',
  'R8 0／R9 1,667／R10 2,982／R11 4,298／R12 5,614',
  '国庫補助金を38年で収益化。新規取得シートH3の式は 財政シミュ!$V$77/38。'),
 ('成果品の推計値','01_rate_reform_simulation_MAIN.xlsx 02_総括原価 C13:G13',
  'R8 ▲94,000／R9 ▲92,000／R10 ▲90,000／R11 ▲88,000／R12 ▲86,000（5年計450,000）',
  '注記は「（推計）非キャッシュ控除。減価償却費の約22%を相殺」。'
  '経営戦略の見込額はR8で減価償却費の83%、5年平均で77%にあたり、推計値は約3.4分の1。'),
 ('他会計補助金（収益的収入）','【成行】財政シミュ r13 ＝ r44＋r59',
  'R8 38,617／R9 47,830／R10 55,812／R11 63,330／R12 67,408（5年計272,997）',
  '総係費の人件費（r44）と支払利息（r59）の合計。根拠欄は「現状の算定式（＝減価償却費＋人件費＋支払利息）」'
  'と書かれているが、式に減価償却費は入っていない。'),
 ('補助金（収益的収入）','【成行】財政シミュ r14',
  'R8 59,000／R9〜R12 各90,000（5年計419,000）',
  '直接入力の固定値で根拠欄なし。料金改定後シートでは同じ行が「既往債の支払利息（r60）」に置き換わっており、'
  '2つのシートで中身が食い違っている。'),
 ('他会計補助金（資本的収入）','【成行】財政シミュ r76',
  'R8〜R17 各3,000',
  '直接入力の固定値。根拠欄は「R6以降はゼロ」とあり、値と矛盾している。'),
 ('他会計補助（元金償還）','【成行】財政シミュ (料金改定後) r79 ＝ r97',
  'R8 30,882／R9 44,652／R10 38,509／R11 47,923／R12 53,577',
  '★資本的収入の最下段。R6末の既往債の元金償還額と同額で、企業債元金償還に対する繰出。'
  '繰出基準に該当する典型例だが、資本的収支のため収益的収支の総括原価には入らない。'
  '成行シートには計上されていない（同じ行は「固定資産売却代金・計上しない」）。'),
 ('出資金（資本的収入）','【成行】財政シミュ r75',
  'R8 515,000（＝65,000＋450,000）／R9 495,000／R10 60,000／R11 40,000／R12以降 20,000',
  '直接入力の固定値。根拠欄は「予定額」。'),
]
for i,(nm,src,val,note) in enumerate(SRC_A,1):
    wr.cell(r,1,i).font=F(size=9,color='808080')
    for cn,v in [(2,nm),(3,src),(4,val),(7,note)]:
        c=wr.cell(r,cn,v); c.font=F(size=9,bold=(cn==2)); c.border=BD
        c.alignment=Alignment(wrap_text=True,vertical='center')
    wr.row_dimensions[r].height=44; r+=1

r+=1; band(wr,r,'【B】経営戦略報告書案の記載',8); r+=1
head(wr,r,['','箇所','記載内容','','','','読み取れること']); r+=1
SRC_B=[
 ('3-8 ⑹料金回収率','数値が低く、国が定める繰出基準以外の繰出金によって収入不足を補填している状況で、'
  '今後も施設更新の予定がありますが、料金改定を行うのか施設更新（投資規模）を先送り又は縮小するのか判断が求められます。',
  '現在の繰入金の相当部分が基準外（赤字補填）であると本文が明言している。'),
 ('4-1 ⑷一般会計繰入金','過去の一般会計からの繰入手法（現金部分の赤字補填）により算定しました。',
  '財政シミュの繰入金は赤字補填額として算定されており、繰出基準に基づく積上げではない。'),
 ('4-1 ⑹支払利息','企業債残高に地方債の加重平均借入利率（2.1％）を乗じて試算しました。',
  '借入利率は2.1％。ただし財政シミュr61の根拠欄には「1.79％」と別の数値が残っている。'),
 ('4-1 ⑺企業債償還金','建設改良費から補助金収入を控除した残額を企業債で充当し、借入年度の翌年度から40年にわたり'
  '元利均等で返済すると仮定して試算しました。','据置期間は設定されていない。調査票は据置3年。'),
 ('4-5 ⑵②企業債','新たに高度処理施設の整備のほか、老朽管布設替えや水道施設耐震化など将来の更新投資に伴い、'
  '企業債の元利償還金の額が増加する見込みです。','調査票のNo.1〜No.3と対応している。北上川導水管への言及はない。'),
 ('4-5 ⑵③繰入金','企業債償還及び施設改良費等を水道料金等収入だけで賄うことが困難な状況にあることから、'
  '一般会計から一定額を繰入を行う計画となっています。国の定める繰出基準内の繰入金はもとより、'
  '不採算地域となる出島地区・江島地区に対する費用負担について、一般会計側との協議を継続していきます。',
  '基準内繰入は既にあり、それに加えて離島分の基準外繰入を求めていく方針。'),
 ('5 ⑶一般会計繰入金','離半島地区での施設整備には、国の定める基準内繰入金のほかに、地域的事情によるものとして'
  '基準外繰入金により整備することが必要となるので、一般会計での負担ができないか検討・協議を行っていきます。',
  '基準内と基準外を明確に区別して記述している。現時点では基準外の負担は未確定。'),
 ('2-2 現状','今後は、本町から24km離れた一級河川・北上川（石巻市）からの導水管や東日本大震災後に未更新の管路、'
  '施設の更新による災害危機対策が求められています。',
  '★北上川導水管は「求められている課題」として書かれているだけで、投資計画には計上されていない。'
  '更新投資額の内訳（r119）の「取・導水管」は全年度ゼロ。'),
]
for i,(pl,cont,note) in enumerate(SRC_B,1):
    wr.cell(r,1,i).font=F(size=9,color='808080')
    c=wr.cell(r,2,pl); c.font=F(size=9,bold=True); c.border=BD; c.alignment=Alignment(wrap_text=True,vertical='center')
    c=wr.cell(r,3,cont); c.font=F(size=9); c.border=BD; c.alignment=Alignment(wrap_text=True,vertical='center')
    wr.merge_cells(start_row=r,start_column=3,end_row=r,end_column=5)
    c=wr.cell(r,7,note); c.font=F(size=9,color=GREEN); c.alignment=Alignment(wrap_text=True,vertical='center')
    wr.row_dimensions[r].height=46; r+=1

r+=1; band(wr,r,'【C】建設改良費の施設区分別対比（R8〜R17・千円）',8); r+=1
head(wr,r,['','施設区分','経営戦略の計画','建設改良費調査票','差引','','対応する事業']); r+=1
SRC_C=[('取・導水管',0,0,'―'),('送水管',400_000,170_000,'No.4 江島海底送水管本復旧工事'),
       ('配水本管',1_380_000,1_180_000,'No.1 老朽管布設替工事'),('配水支管',0,0,'―'),
       ('構造物及び設備',2_640_000,2_165_000,'No.2 水道施設耐震化（1,277,000）＋No.3 鷲神高度処理（888,000）'),
       ('量水器費',106_000,0,'調査票の対象外')]
for i,(nm,o,n,proj) in enumerate(SRC_C,1):
    wr.cell(r,1,i).font=F(size=9,color='808080')
    wr.cell(r,2,nm).font=F(size=9,bold=True); wr.cell(r,2).border=BD
    for cn,v in [(3,o),(4,n),(5,n-o)]:
        c=wr.cell(r,cn,v); c.number_format='#,##0'; c.font=F(size=9,color=(RED if cn==5 and n-o<0 else '000000')); c.border=BD
    wr.cell(r,7,proj).font=F(size=9,color='595959'); r+=1
wr.cell(r,2,'計').font=F(size=10,bold=True,color=NAVY); wr.cell(r,2).border=BD
for cn,v in [(3,4_526_000),(4,3_515_000),(5,-1_011_000)]:
    c=wr.cell(r,cn,v); c.number_format='#,##0'; c.font=F(size=10,bold=True,color=NAVY); c.border=BD
    c.fill=PatternFill('solid',fgColor='DDEBF7')
r+=2
wr.cell(r,2,'※ 経営戦略の内訳は【成行】財政シミュ r119〜r124（更新投資額）。量水器費は同 r93。').font=F(size=9,color='595959')
wr.merge_cells(start_row=r,start_column=2,end_row=r,end_column=7)

for s in wb.worksheets:
    s.sheet_view.showGridLines=False
    if s.title=='05_計算用データ': s.sheet_state='hidden'
wb.save(OUT)
print('作成:', os.path.abspath(OUT))
print('減価償却表', len(dep_rows), '行 ／ 利息表', len(int_rows), '行')
print()
print(f'{"ケース":42s}{"総括原価":>13s}{"給水原価":>11s}{"回収率":>9s}{"年間不足額":>13s}{"必要改定率":>11s}')
print('-'*100)
for nm,kw in CASES:
    x=evaluate(**kw)
    print(f'{nm:40s}{x["total"]:>13,.0f}{x["g"]:>11.1f}{x["rec"]:>9.1%}{x["short"]:>13,.0f}{x["ratio"]:>10.2f}倍')
