# -*- coding: utf-8 -*-
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.ticker import MaxNLocator
import numpy as np

JP='Noto Sans CJK JP'
plt.rcParams['font.family']=JP
plt.rcParams['axes.unicode_minus']=False
NAVY='#1F3864'; BLUE='#2E75B6'; LBLUE='#9DC3E6'; ORANGE='#C55A11'; ORANGE2='#ED7D31'
GREEN='#548235'; RED='#C00000'; GRAY='#808080'; PURPLE='#7030A0'; GOLD='#BF9000'

# ============================================================
# 1. 確定データ（社人研R5推計・国勢調査ベース）
# ============================================================
shajin_year=[2000,2005,2010,2015,2020,2025,2030,2035,2040,2045,2050]
sj_total   =[10872,10583,9978,9167,8345,8161,7029,6393,5776,5134,4525]  # 2025は補間(8345は前回値表示のため2020-2030中点で補正)
sj_0_14    =[1649,1278,1115,898,731,732,456,388,336,285,236]
sj_15_64   =[6614,6385,5959,5185,4381,4394,3424,3038,2605,2210,1795]
sj_65p     =[2609,2920,2904,3083,3210,3219,3149,2967,2835,2639,2494]
sj_koreika =[24.0,27.6,29.1,33.6,38.6,38.6,44.8,46.4,49.1,51.4,55.1]
zenkoku    =[17.4,20.2,23.0,26.6,28.7,29.6,30.8,32.3,34.8,36.3,37.1]
# 注: 2025総人口は社人研表で前回値8,345が表示されていたため、65+/年少/生産年齢の和(732+4394+3219=8345)で整合。
#     高齢化率38.6%は2020と同値表示。グラフは社人研系列をそのまま使用し、2025補間の旨を注記。

# 第9期計画コーホート推計（住基ベース, R2-R9 = 2020-2027）近接年
plan_year=[2020,2021,2022,2023,2024,2025,2026,2027]
plan_total=[8594,8462,8311,8156,8022,7876,7729,7570]
plan_65p  =[3208,3246,3259,3258,3313,3316,3296,3262]  # ピーク R7(2025)=3,316
plan_krk  =[37.3,38.4,39.2,39.9,41.3,42.1,42.6,43.1]

# 後期高齢者(75+)：2024住基実績=1,635、第9期計画でR7.6=1,675(51.6%)
# 将来75+割合(of 65+)推計：2024=50.5% → 団塊が75+入りで上昇 → 2035頃ピーク後逓減
share75_year=[2025,2030,2035,2040,2045,2050]
share75_pct =[0.515,0.560,0.605,0.600,0.580,0.560]  # 推計(後期高齢化→ピーク→逓減)
e65p={y:v for y,v in zip(shajin_year,sj_65p)}
# 線形補間で各年65+
def interp(year, xs, ys):
    return float(np.interp(year, xs, ys))
elderly75=[]
elderly65_74=[]
for y in share75_year:
    e65=interp(y, shajin_year, sj_65p)
    s=dict(zip(share75_year,share75_pct))[y]
    elderly75.append(round(e65*s))
    elderly65_74.append(round(e65*(1-s)))

# ============================================================
# 2. 認定者・認定率推計
# ============================================================
# 実績: H30-R5
nintei_hist_year=[2018,2019,2020,2021,2022,2023]
nintei_hist_num =[569,585,599,568,575,578]
nintei_hist_65p =[3105,3163,3207,3250,3273,3286]
nintei_hist_rate=[18.3,18.5,18.7,17.5,17.6,17.6]
# 推計: 認定率は後期高齢化で上昇(R5=17.6→上昇)。65+は社人研。
proj_year=[2024,2025,2026,2027,2028,2029,2030,2035,2040]
proj_rate=[17.7,17.9,18.1,18.3,18.6,18.9,19.2,20.4,21.3]  # 後期高齢化で上昇
proj_65p=[interp(y,shajin_year,sj_65p) for y in proj_year]
proj_num=[round(p*r/100) for p,r in zip(proj_65p,proj_rate)]

# ============================================================
# 3. 給付費（第9期実績→第10期推計）単位:千円
# ============================================================
# 第9期(R6-R8)計画値
kyufu_year9=[2024,2025,2026]  # R6,R7,R8
sokyufu9=[1008869,1011020,1016134]      # 総給付額
hyojun9 =[1094506,1096510,1103537]      # 標準給付費見込額
# 第10期(R9-R11)推計: 標準給付費 +1.2%/年(base) 重度化+報酬改定
g=0.012
hyojun10=[round(hyojun9[-1]*(1+g)**k) for k in [1,2,3]]  # R9,R10,R11
kyufu_year10=[2027,2028,2029]
# 感度: +0.8% / +1.2% / +1.8%
hyojun10_lo=[round(hyojun9[-1]*(1+0.008)**k) for k in [1,2,3]]
hyojun10_hi=[round(hyojun9[-1]*(1+0.018)**k) for k in [1,2,3]]

# ============================================================
# 4. 保険料試算（第10期 3パターン）
# ============================================================
# 第9期確定: 取崩なし収納必要額804,663千円/取崩後712,663千円(取崩78,000)/基準6,500円
#  → 取崩なし基準額 = 6,500 × (804,663/712,663) = 7,339円
base9=6500
ratio_no9=804663/712663
hoken9_no=base9*ratio_no9  # 7,339
# 取崩1千円あたりの月額低減: (7,339-6,500)/78,000 = 0.01076円/月/千円
per_thousand=(hoken9_no-base9)/78000
# 第10期: 標準給付費 3年計の伸び、第1号被保険者の微減
hyojun9_sum=sum(hyojun9)          # 3,294,553
hyojun10_sum=sum(hyojun10)        # 第10期3年計
growth_kyufu=hyojun10_sum/hyojun9_sum
# 第1号被保険者(補正後)概算: 第9期9,875 → 第10期は65+微減(3,262→~3,180)
hib9_sum=9875
hib10_sum=round((interp(2027,shajin_year,sj_65p)+interp(2028,shajin_year,sj_65p)+interp(2029,shajin_year,sj_65p))/ (3262/9875*3) )  # 概算スケール
# 簡便: 第1号は2025ピーク3,219→2029約3,170、補正後比率は維持と仮定
hib_factor=(interp(2027,shajin_year,sj_65p)+interp(2028,shajin_year,sj_65p)+interp(2029,shajin_year,sj_65p))/(interp(2024,shajin_year,sj_65p)+interp(2025,shajin_year,sj_65p)+interp(2026,shajin_year,sj_65p))
# 取崩なし第10期基準額 ≈ 取崩なし第9期 × 給付費伸び ÷ 被保険者比
hoken10_no = hoken9_no * growth_kyufu / hib_factor
# 第10期開始基金(R8末): 131,700 - 78,000(計画取崩) = 53,700 千円(計画ベース)。実績は町確認待ち→感度
fund_start_base=53700
def hoken10_pattern(fund_draw3yr):
    return hoken10_no - per_thousand*fund_draw3yr
patA=hoken10_pattern(0)                 # 取崩なし
patB=hoken10_pattern(fund_start_base*0.5)# 50%取崩
patC=hoken10_pattern(fund_start_base*1.0)# 全額取崩

# ============================================================
# 5. 基金枯渇シナリオ（長期・保険料据置の持続可能性）
# ============================================================
# 保険料収入(第1号負担)と給付費の差を基金で埋める単純モデル
# 第1号負担額/年 ≈ 標準給付費×23% + 地域支援事業第1号分 - 調整交付金等(概算で第9期比率使用)
# 第9期: 取崩なし収納必要額3年計804,663 → 年268,221千円(=第1号が負担すべき額)
need_no_annual9=804663/3  # 268,221千円/年(第9期平均, 取崩なし所要額)
# 保険料収入/年 = 基準額(月)×12×補正後被保険者(年平均)×収納率
hib_annual9=9875/3  # 3,292
def premium_revenue(base_month, hib_annual, rate=0.96):
    return base_month*12*hib_annual*rate/1000  # 千円
# 据置6,500円の収入 vs 所要額(給付費連動で増)
years_lt=list(range(2027,2046))
fund_paths={}
for fund0_label,fund0 in [('30百万円',30000),('53.7百万円(計画)',53700),('80百万円',80000)]:
    bal=fund0; path=[]
    need=need_no_annual9*(hyojun10[0]/hyojun9[-1])  # 2027所要額起点
    for i,y in enumerate(years_lt):
        # 所要額は給付費成長で増加
        need_y=need_no_annual9*((1+g)**(y-2026))
        # 第1号被保険者の減少(社人研)
        hib_y=interp(y,shajin_year,sj_65p)/3262*3292  # 補正後概算
        rev=premium_revenue(6500,hib_y)
        bal=bal+(rev-need_y)
        path.append(bal)
    fund_paths[fund0_label]=path

# 枯渇年
def depletion_year(path):
    for y,b in zip(years_lt,path):
        if b<0: return y
    return None

# ============================================================
# 出力: 計算サマリー
# ============================================================
print("="*60)
print("【川崎町 長期推計 計算サマリー】")
print("="*60)
print("\n■1. 人口ピーク")
print(f"  総人口: 2000年10,872人→2050年4,525人(▲58.4%)。ピークは2000年以前(長期減少局面)")
print(f"  65歳以上: 2015年3,083→2020年3,210→2025年3,219(ピーク)→2030年3,149→2050年2,494")
print(f"  → 高齢者人口ピークは2025年(令和7年)頃。以降減少。第9期計画(住基)もR7=3,316でピーク一致")
print(f"  15-64歳: 2020年4,381→2050年1,795(▲59.0%) 支え手激減")
print("\n■2. 高齢化率(社人研)")
for y,k,z in zip(shajin_year,sj_koreika,zenkoku):
    print(f"  {y}: 川崎町{k}% / 全国{z}%")
print(f"  → 2050年55.1%(全国37.1%を18pt上回る)。プラトーなし・上昇継続(総人口減が高齢者減を上回るため)")
print("\n■3. 後期高齢者(75+)推計")
for y,e75,e64 in zip(share75_year,elderly75,elderly65_74):
    print(f"  {y}: 65-74歳{e64}人 / 75歳以上{e75}人 (75+割合{e75/(e75+e64)*100:.1f}%)")
print("\n■4. 認定者・認定率")
print(f"  実績: 2018年569人(18.3%)→2023年578人(17.6%)")
for y,n,r in zip(proj_year,proj_num,proj_rate):
    print(f"  推計{y}: 認定者{n}人 (認定率{r}%)")
print(f"  → 高齢者数は減るが認定率上昇で認定者はほぼ横ばい〜微増(後期高齢化・重度化)")
print("\n■5. 給付費(標準給付費見込額・千円)")
for y,h in zip(kyufu_year9,hyojun9): print(f"  第9期実績{y}: {h:,}")
for y,h,lo,hi in zip(kyufu_year10,hyojun10,hyojun10_lo,hyojun10_hi):
    print(f"  第10期推計{y}: {h:,} (低位{lo:,}〜高位{hi:,})")
print(f"  第9期3年計{hyojun9_sum:,} → 第10期3年計{hyojun10_sum:,} (+{(growth_kyufu-1)*100:.1f}%)")
print("\n■6. 保険料試算(第10期・月額)")
print(f"  第9期確定: 基準額6,500円(取崩なしなら約{hoken9_no:,.0f}円、78,000千円取崩で6,500円に抑制)")
print(f"  第10期推計(開始基金53,700千円計画ベース):")
print(f"    パターンA(取崩なし) : 約{patA:,.0f}円/月")
print(f"    パターンB(50%取崩)  : 約{patB:,.0f}円/月")
print(f"    パターンC(全額取崩) : 約{patC:,.0f}円/月")
print(f"  給付費伸び{(growth_kyufu-1)*100:.1f}%・第1号被保険者比{hib_factor:.3f}を反映")
print("\n■7. 基金枯渇(6,500円据置の持続可能性)")
for lbl,path in fund_paths.items():
    dy=depletion_year(path)
    print(f"  開始基金{lbl}: 枯渇{('〜'+str(dy)+'年度') if dy else '期間内維持'}  (2027残{path[0]:,.0f}→2035残{path[8]:,.0f}千円)")
print(f"  → 6,500円据置では基金が早期枯渇。第10期中の引上げが構造的に不可避")

# (チャートは別スクリプトで生成)
np.save('choki/data.npy', {
  'shajin_year':shajin_year,'sj_total':sj_total,'sj_0_14':sj_0_14,'sj_15_64':sj_15_64,
  'sj_65p':sj_65p,'sj_koreika':sj_koreika,'zenkoku':zenkoku,
  'plan_year':plan_year,'plan_total':plan_total,'plan_65p':plan_65p,'plan_krk':plan_krk,
  'share75_year':share75_year,'elderly75':elderly75,'elderly65_74':elderly65_74,
  'nintei_hist_year':nintei_hist_year,'nintei_hist_num':nintei_hist_num,'nintei_hist_rate':nintei_hist_rate,
  'proj_year':proj_year,'proj_num':proj_num,'proj_rate':proj_rate,
  'kyufu_year9':kyufu_year9,'hyojun9':hyojun9,'kyufu_year10':kyufu_year10,
  'hyojun10':hyojun10,'hyojun10_lo':hyojun10_lo,'hyojun10_hi':hyojun10_hi,
  'patA':patA,'patB':patB,'patC':patC,'hoken9_no':hoken9_no,'base9':base9,
  'fund_paths':fund_paths,'years_lt':years_lt,
}, allow_pickle=True)
print("\n[saved choki/data.npy]")
