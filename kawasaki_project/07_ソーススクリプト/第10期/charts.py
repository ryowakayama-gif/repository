# -*- coding: utf-8 -*-
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
plt.rcParams['font.family']='Noto Sans CJK JP'
plt.rcParams['axes.unicode_minus']=False
plt.rcParams['savefig.dpi']=150
NAVY='#1F3864'; BLUE='#2E75B6'; LBLUE='#9DC3E6'; ORANGE='#C55A11'; ORANGE2='#ED7D31'
GREEN='#548235'; RED='#C00000'; GRAY='#808080'; PURPLE='#7030A0'; GOLD='#BF9000'; TEAL='#2F8F83'

D=np.load('choki/data.npy',allow_pickle=True).item()
sy=D['shajin_year']

# ---- Chart 1: 人口の長期推移(年齢3区分・棒)+総人口(線) ----
fig,ax=plt.subplots(figsize=(8.6,4.5))
w=3.2
y014=np.array(D['sj_0_14']); y1564=np.array(D['sj_15_64']); y65=np.array(D['sj_65p'])
ax.bar(sy,y014,w,label='0-14歳',color=LBLUE)
ax.bar(sy,y1564,w,bottom=y014,label='15-64歳',color=BLUE)
ax.bar(sy,y65,w,bottom=y014+y1564,label='65歳以上',color=ORANGE)
tot=y014+y1564+y65
ax.plot(sy,tot,'-o',color=NAVY,lw=2,ms=4,label='総人口')
ax.axvline(2025,color=RED,ls='--',lw=1,alpha=0.6)
ax.annotate('高齢者人口ピーク\n2025年 3,219人',xy=(2025,3219),xytext=(2031,7600),
  fontsize=9,color=RED,ha='left',arrowprops=dict(arrowstyle='->',color=RED))
ax.annotate('総人口▲58%\n10,872→4,525人',xy=(2050,4525),xytext=(2038,9200),fontsize=9,color=NAVY,ha='left')
ax.set_ylabel('人口（人）'); ax.set_title('川崎町の人口の長期推移（社人研R5推計）',fontweight='bold',color=NAVY)
ax.legend(loc='upper right',fontsize=8,ncol=2); ax.set_ylim(0,12500); ax.grid(axis='y',alpha=0.3)
ax.set_xticks(sy); ax.set_xticklabels([str(y) for y in sy],fontsize=8)
plt.tight_layout(); plt.savefig('choki/c1_population.png'); plt.close()

# ---- Chart 2: 高齢化率の長期推移(川崎町vs全国) ----
fig,ax=plt.subplots(figsize=(8.6,4.3))
ax.plot(sy,D['sj_koreika'],'-o',color=ORANGE,lw=2.5,ms=5,label='川崎町')
ax.plot(sy,D['zenkoku'],'-s',color=GRAY,lw=1.8,ms=4,label='全国平均')
ax.fill_between(sy,D['sj_koreika'],D['zenkoku'],color=ORANGE,alpha=0.10)
ax.axvspan(2027,2029,color=BLUE,alpha=0.10)
ax.annotate('第10期\n(R9-11)',xy=(2028,30),fontsize=8,color=BLUE,ha='center')
ax.annotate('55.1%',xy=(2050,55.1),xytext=(2044,53),fontsize=10,color=ORANGE,fontweight='bold')
ax.annotate('全国37.1%',xy=(2050,37.1),xytext=(2043,33),fontsize=9,color=GRAY)
ax.text(2009,46,'プラトーなし・上昇継続\n(総人口減＞高齢者減)',fontsize=8.5,color=RED,
  bbox=dict(boxstyle='round',fc='#FCE4D6',ec=ORANGE,alpha=0.8))
ax.set_ylabel('高齢化率（%）'); ax.set_title('川崎町の高齢化率の長期推移（社人研R5推計）',fontweight='bold',color=NAVY)
ax.legend(loc='lower right',fontsize=9); ax.set_ylim(15,60); ax.grid(alpha=0.3)
ax.set_xticks(sy); ax.set_xticklabels([str(y) for y in sy],fontsize=8)
plt.tight_layout(); plt.savefig('choki/c2_koreika.png'); plt.close()

# ---- Chart 3: 後期高齢化(65-74 vs 75+) ----
fig,ax=plt.subplots(figsize=(8.6,4.3))
sh_y=D['share75_year']; e64=np.array(D['elderly65_74']); e75=np.array(D['elderly75'])
ax.bar(sh_y,e64,3.2,label='65-74歳（前期）',color=LBLUE)
ax.bar(sh_y,e75,3.2,bottom=e64,label='75歳以上（後期）',color=PURPLE)
share=[a/(a+b)*100 for a,b in zip(e75,e64)]
ax2=ax.twinx()
ax2.plot(sh_y,share,'-o',color=RED,lw=2,ms=5,label='75歳以上割合')
ax2.set_ylabel('75歳以上が高齢者に占める割合（%）',color=RED); ax2.set_ylim(40,70)
ax2.tick_params(axis='y',colors=RED)
for x,s in zip(sh_y,share): ax2.annotate(f'{s:.0f}%',xy=(x,s),xytext=(x,s+1.5),fontsize=8,color=RED,ha='center')
ax.set_ylabel('高齢者人口（人）'); ax.set_title('川崎町の高齢者人口の後期高齢化（推計）',fontweight='bold',color=NAVY)
ax.legend(loc='upper left',fontsize=8); ax.set_ylim(0,3600); ax.grid(axis='y',alpha=0.3)
ax.set_xticks(sh_y)
plt.tight_layout(); plt.savefig('choki/c3_kouki.png'); plt.close()

# ---- Chart 4: 認定者数・認定率 ----
fig,ax=plt.subplots(figsize=(8.6,4.3))
hy=D['nintei_hist_year']; hn=D['nintei_hist_num']
py=D['proj_year']; pn=D['proj_num']; pr=D['proj_rate']
ax.bar(hy,hn,0.8,color=BLUE,label='認定者数（実績）')
ax.bar(py,pn,0.8,color=ORANGE2,alpha=0.85,label='認定者数（推計）')
ax2=ax.twinx()
ax2.plot(hy,D['nintei_hist_rate'],'-o',color=NAVY,lw=2,ms=4,label='認定率（実績）')
ax2.plot(py,pr,'--o',color=RED,lw=2,ms=4,label='認定率（推計）')
ax2.set_ylabel('認定率（%）'); ax2.set_ylim(14,24)
ax.set_ylabel('認定者数（人）'); ax.set_ylim(0,800)
ax.set_title('川崎町の要支援・要介護認定者数と認定率の推計',fontweight='bold',color=NAVY)
ax.annotate('高齢者数は減るが\n認定率上昇で認定者は横ばい〜微増',xy=(2032,560),fontsize=8.5,color=RED,
  bbox=dict(boxstyle='round',fc='#FCE4D6',ec=ORANGE,alpha=0.8))
l1,la1=ax.get_legend_handles_labels(); l2,la2=ax2.get_legend_handles_labels()
ax.legend(l1+l2,la1+la2,loc='upper left',fontsize=7.5,ncol=2); ax.grid(axis='y',alpha=0.3)
plt.tight_layout(); plt.savefig('choki/c4_nintei.png'); plt.close()

# ---- Chart 5: 給付費(第9期実績→第10期推計) ----
fig,ax=plt.subplots(figsize=(8.6,4.3))
y9=D['kyufu_year9']; h9=[v/1000 for v in D['hyojun9']]
y10=D['kyufu_year10']; h10=[v/1000 for v in D['hyojun10']]
h10lo=[v/1000 for v in D['hyojun10_lo']]; h10hi=[v/1000 for v in D['hyojun10_hi']]
ax.bar(y9,h9,0.7,color=BLUE,label='第9期（計画値）')
ax.bar(y10,h10,0.7,color=ORANGE,label='第10期（推計・中位）')
ax.errorbar(y10,h10,yerr=[np.array(h10)-np.array(h10lo),np.array(h10hi)-np.array(h10)],
  fmt='none',ecolor=RED,capsize=4,lw=1.2,label='低位〜高位')
for x,v in zip(y9,h9): ax.annotate(f'{v:,.0f}',xy=(x,v),xytext=(x,v+8),fontsize=7.5,ha='center',color=NAVY)
for x,v in zip(y10,h10): ax.annotate(f'{v:,.0f}',xy=(x,v),xytext=(x,v+18),fontsize=7.5,ha='center',color=ORANGE)
ax.set_ylabel('標準給付費見込額（百万円）'); ax.set_ylim(1000,1200)
ax.set_title('川崎町の介護給付費の推移と第10期推計',fontweight='bold',color=NAVY)
ax.annotate('第10期3年計+2.9%\n重度化＋報酬改定',xy=(2028,1160),fontsize=8.5,color=RED,ha='center')
ax.legend(loc='upper left',fontsize=8); ax.grid(axis='y',alpha=0.3)
ax.set_xticks(list(y9)+list(y10))
plt.tight_layout(); plt.savefig('choki/c5_kyufu.png'); plt.close()

# ---- Chart 6: 基金枯渇シナリオ + 保険料3パターン ----
fig,(axL,axR)=plt.subplots(1,2,figsize=(9.4,4.3),gridspec_kw={'width_ratios':[1.35,1]})
years_lt=D['years_lt']
colors={'30百万円':GRAY,'53.7百万円(計画)':BLUE,'80百万円':GREEN}
for lbl,path in D['fund_paths'].items():
    p=[max(v,0) for v in path]  # 0でクリップ
    # 枯渇後は0
    dep=None
    for i,v in enumerate(path):
        if v<0: dep=i; break
    if dep is not None:
        p=[max(path[i],0) if i<=dep else 0 for i in range(len(path))]
    axL.plot(years_lt,[v/1000 for v in p],'-o',color=colors[lbl],lw=2,ms=3,label=f'開始基金{lbl}')
axL.axhline(0,color=RED,lw=1)
axL.axvspan(2027,2029,color=BLUE,alpha=0.10)
axL.annotate('第10期',xy=(2028,axL.get_ylim()[1]*0.6 if False else 60),fontsize=8,color=BLUE,ha='center')
axL.set_ylabel('基金残高（百万円）'); axL.set_xlabel('年度')
axL.set_title('基金残高シナリオ（6,500円据置の場合）',fontweight='bold',color=NAVY,fontsize=11)
axL.legend(loc='upper right',fontsize=7.5); axL.grid(alpha=0.3); axL.set_ylim(0,90)
axL.text(2030,40,'6,500円据置では\n第10期中に枯渇',fontsize=8.5,color=RED,
  bbox=dict(boxstyle='round',fc='#FCE4D6',ec=ORANGE,alpha=0.85))

# 保険料3パターン
pats=['第9期\n実績','第10期\nA取崩なし','第10期\nB 50%','第10期\nC全額']
vals=[D['base9'],D['patA'],D['patB'],D['patC']]
cols=[GRAY,RED,ORANGE,GREEN]
bars=axR.bar(pats,vals,color=cols,alpha=0.9)
for b,v in zip(bars,vals): axR.annotate(f'{v:,.0f}円',xy=(b.get_x()+b.get_width()/2,v),xytext=(0,3),
  textcoords='offset points',ha='center',fontsize=8.5,fontweight='bold')
axR.axhline(D['base9'],color=GRAY,ls=':',lw=1)
axR.set_ylabel('保険料基準額（月額・円）'); axR.set_ylim(0,8400)
axR.set_title('第10期 保険料試算',fontweight='bold',color=NAVY,fontsize=11)
axR.tick_params(axis='x',labelsize=7.5); axR.grid(axis='y',alpha=0.3)
plt.tight_layout(); plt.savefig('choki/c6_fund_premium.png'); plt.close()

print("charts generated:")
import os
for f in sorted(os.listdir('choki')):
    if f.endswith('.png'): print(f"  {f} ({os.path.getsize('choki/'+f)//1024}KB)")
