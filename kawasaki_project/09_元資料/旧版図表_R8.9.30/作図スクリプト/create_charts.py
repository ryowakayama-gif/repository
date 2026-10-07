"""
川崎町第10期計画素案用 図表生成スクリプト
6種類の主要グラフを生成：
1. chart_population.png  - 第1号被保険者の年齢階級別推移（R3→R7）
2. chart_recipient.png   - サービス受給者構成（円グラフ）
3. chart_benefit.png     - 給付費のサービス種類別構成（横棒）
4. chart_income.png      - 所得段階別人口分布
5. chart_premium.png     - 保険料推移（第8期→第9期→第10期予測）
6. chart_aging.png       - 高齢化率比較（川崎町・宮城県・全国）
"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np

# 日本語フォント設定
JP_FONT = fm.FontProperties(fname='/usr/share/fonts/opentype/noto/NotoSansCJK-Medium.ttc', size=10)
JP_FONT_S = fm.FontProperties(fname='/usr/share/fonts/opentype/noto/NotoSansCJK-Medium.ttc', size=8)
JP_FONT_L = fm.FontProperties(fname='/usr/share/fonts/opentype/noto/NotoSansCJK-Medium.ttc', size=12)
JP_FONT_T = fm.FontProperties(fname='/usr/share/fonts/opentype/noto/NotoSansCJK-Medium.ttc', size=14)
plt.rcParams['axes.unicode_minus'] = False

# 計画書のカラーパレット
NAVY = "#1F3864"
BLUE = "#2F5597"
LBLUE = "#DAE3F3"
ORANGE = "#ED7D31"
LORANGE = "#FCE4D6"
GREEN = "#548235"
LGREEN = "#E2EFDA"
RED = "#C00000"
PURPLE = "#7030A0"
GRAY = "#808080"

# ===========================================================
# 1. 第1号被保険者の年齢階級別推移
# ===========================================================
fig, ax = plt.subplots(figsize=(8, 5), dpi=150)
labels = ['65〜75歳未満\n（前期高齢者）', '75〜85歳未満', '85歳以上']
r3_vals = [1745, 905, 605]
r7_vals = [1569, 1070, 605]
x = np.arange(len(labels))
width = 0.35

bars1 = ax.bar(x - width/2, r3_vals, width, label='令和3年度末', color=BLUE, edgecolor='white', linewidth=1.5)
bars2 = ax.bar(x + width/2, r7_vals, width, label='令和7年6月', color=ORANGE, edgecolor='white', linewidth=1.5)

for bars in [bars1, bars2]:
    for b in bars:
        h = b.get_height()
        ax.text(b.get_x()+b.get_width()/2, h+30, f'{h:,}人', ha='center', va='bottom',
                fontproperties=JP_FONT_S, fontsize=9)

ax.set_xticks(x)
ax.set_xticklabels(labels, fontproperties=JP_FONT)
ax.set_ylabel('人数（人）', fontproperties=JP_FONT)
ax.set_title('図2-1　第1号被保険者の年齢階級別推移（R3年度末→R7.6月）', fontproperties=JP_FONT_T, color=NAVY, pad=15)
ax.set_ylim(0, 2200)
ax.legend(prop=JP_FONT, loc='upper right', framealpha=0.95)
ax.grid(axis='y', linestyle=':', alpha=0.4)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

# 増減率の矢印・コメント
ax.annotate('△10.1%', xy=(0, 1700), xytext=(0, 2050),
            arrowprops=dict(arrowstyle='->', color=RED, lw=2),
            ha='center', fontproperties=JP_FONT, fontsize=11, color=RED, fontweight='bold')
ax.annotate('+18.2%', xy=(1, 1000), xytext=(1, 1500),
            arrowprops=dict(arrowstyle='->', color=GREEN, lw=2),
            ha='center', fontproperties=JP_FONT, fontsize=11, color=GREEN, fontweight='bold')
ax.annotate('±0.0%', xy=(2, 605), xytext=(2, 900),
            arrowprops=dict(arrowstyle='->', color=GRAY, lw=2),
            ha='center', fontproperties=JP_FONT, fontsize=11, color=GRAY, fontweight='bold')

plt.tight_layout()
plt.savefig('/home/claude/kawasaki_work/chart_population.png', dpi=150, bbox_inches='tight', facecolor='white')
plt.close()
print("✓ chart_population.png")

# ===========================================================
# 2. サービス受給者構成（円グラフ・R7.6時点）
# ===========================================================
fig, ax = plt.subplots(figsize=(7, 6), dpi=150)
sizes = [276, 56, 134]
labels_pie = ['居宅サービス\n276人 (59.2%)', '地域密着型\n56人 (12.0%)', '施設サービス\n134人 (28.8%)']
colors_pie = [BLUE, GREEN, ORANGE]
explode = (0, 0.04, 0.04)
wedges, texts = ax.pie(sizes, labels=labels_pie, colors=colors_pie, 
                       startangle=90, explode=explode,
                       textprops={'fontproperties': JP_FONT, 'fontsize': 11},
                       wedgeprops={'edgecolor': 'white', 'linewidth': 2})

# 中央に合計を表示（白背景の円を先に描画）
from matplotlib.patches import Circle
ax.add_patch(Circle((0, 0), 0.32, facecolor='white', edgecolor=NAVY, linewidth=1.5, zorder=10))
ax.text(0, 0, '計466人\n（重複あり）', ha='center', va='center', 
        fontproperties=JP_FONT_L, fontsize=13, color=NAVY, fontweight='bold', zorder=11)

ax.set_title('図2-2　サービス受給者の区分別構成（令和7年6月時点）', 
             fontproperties=JP_FONT_T, color=NAVY, pad=15)

plt.tight_layout()
plt.savefig('/home/claude/kawasaki_work/chart_recipient.png', dpi=150, bbox_inches='tight', facecolor='white')
plt.close()
print("✓ chart_recipient.png")

# ===========================================================
# 3. 給付費のサービス種類別構成（横棒・R3年度）
# ===========================================================
fig, ax = plt.subplots(figsize=(9, 5.5), dpi=150)
services = ['居宅サービス\n（在宅）', '地域密着型\nサービス', '施設サービス\n（特養・老健）']
values_oku = [3.66, 1.53, 4.42]  # 億円
colors_bar = [BLUE, GREEN, ORANGE]

bars = ax.barh(services, values_oku, color=colors_bar, edgecolor='white', linewidth=2)
for i, (b, v) in enumerate(zip(bars, values_oku)):
    pct = v / sum(values_oku) * 100
    ax.text(v + 0.08, b.get_y()+b.get_height()/2, f'{v:.2f}億円  ({pct:.1f}%)',
            va='center', fontproperties=JP_FONT, fontsize=11, fontweight='bold')

ax.set_xlabel('給付費（億円）', fontproperties=JP_FONT)
ax.set_yticklabels(services, fontproperties=JP_FONT)
ax.set_title('図2-3　サービス区分別の年間給付費（令和3年度実績・総額9.6億円）',
             fontproperties=JP_FONT_T, color=NAVY, pad=15)
ax.set_xlim(0, 5.5)
ax.grid(axis='x', linestyle=':', alpha=0.4)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.invert_yaxis()

plt.tight_layout()
plt.savefig('/home/claude/kawasaki_work/chart_benefit.png', dpi=150, bbox_inches='tight', facecolor='white')
plt.close()
print("✓ chart_benefit.png")

# ===========================================================
# 4. 所得段階別人口分布
# ===========================================================
fig, ax = plt.subplots(figsize=(10, 5.5), dpi=150)
dankai = ['第1段階', '第2段階', '第3段階', '第4段階', '第5段階\n(基準)', '第6段階', '第7段階', '第8段階', '第9段階']
people = [431, 273, 275, 466, 697, 449, 385, 173, 106]

# 段階別の色分け（非課税層=緑系、本人非課税=青系、課税層=オレンジ系）
colors_inc = [LGREEN, LGREEN, LGREEN, LBLUE, ORANGE, LORANGE, LORANGE, LORANGE, LORANGE]
edge_colors = [GREEN, GREEN, GREEN, BLUE, NAVY, ORANGE, ORANGE, ORANGE, ORANGE]

bars = ax.bar(dankai, people, color=colors_inc, edgecolor=edge_colors, linewidth=2)
for b, n in zip(bars, people):
    ax.text(b.get_x()+b.get_width()/2, n+15, f'{n}', ha='center', va='bottom',
            fontproperties=JP_FONT_S, fontsize=10, fontweight='bold')

ax.set_ylabel('人数（人）', fontproperties=JP_FONT)
ax.set_xticklabels(dankai, fontproperties=JP_FONT_S, fontsize=9)
ax.set_title('図2-4　所得段階別第1号被保険者の構成（令和3年度末・計3,255人）',
             fontproperties=JP_FONT_T, color=NAVY, pad=15)
ax.set_ylim(0, 800)
ax.grid(axis='y', linestyle=':', alpha=0.4)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

# 凡例
from matplotlib.patches import Patch
legend_elements = [
    Patch(facecolor=LGREEN, edgecolor=GREEN, label='非課税層（30.1%）'),
    Patch(facecolor=LBLUE, edgecolor=BLUE, label='本人非課税（14.3%）'),
    Patch(facecolor=ORANGE, edgecolor=NAVY, label='第5段階（基準・21.4%）'),
    Patch(facecolor=LORANGE, edgecolor=ORANGE, label='本人課税層（34.2%）'),
]
ax.legend(handles=legend_elements, prop=JP_FONT_S, loc='upper right', framealpha=0.95)

plt.tight_layout()
plt.savefig('/home/claude/kawasaki_work/chart_income.png', dpi=150, bbox_inches='tight', facecolor='white')
plt.close()
print("✓ chart_income.png")

# ===========================================================
# 5. 保険料推移（第8期→第9期→第10期試算3パターン）
# ===========================================================
fig, ax = plt.subplots(figsize=(9, 5.5), dpi=150)

periods = ['第8期\n(R3〜R5)', '第9期\n(R6〜R8)', '第10期\n(R9〜R11)\n試算A\n基金なし', '第10期\n(R9〜R11)\n試算B\n基金50%', '第10期\n(R9〜R11)\n試算C\n基金全額']
# 仮置きの試算値（実際の試算結果反映後Ver2.0で更新）
prices = [6380, 6500, None, None, None]  # 第10期は試算後

x = np.arange(len(periods))
# 確定分（実線）
ax.plot([0, 1], [6380, 6500], color=NAVY, linewidth=3, marker='o', markersize=12, 
        markerfacecolor=NAVY, markeredgecolor='white', markeredgewidth=2)
# 試算（破線・点線）
ax.plot([1, 2], [6500, 6650], color=GRAY, linewidth=2, linestyle='--', marker='s', markersize=10,
        markerfacecolor=LBLUE, markeredgecolor=GRAY, markeredgewidth=2)
ax.plot([1, 3], [6500, 6500], color=GRAY, linewidth=2, linestyle='--', marker='s', markersize=10,
        markerfacecolor=LORANGE, markeredgecolor=GRAY, markeredgewidth=2)
ax.plot([1, 4], [6500, 6350], color=GRAY, linewidth=2, linestyle='--', marker='s', markersize=10,
        markerfacecolor=LGREEN, markeredgecolor=GRAY, markeredgewidth=2)

# データラベル
ax.text(0, 6380-100, '6,380円', ha='center', va='top', fontproperties=JP_FONT, fontsize=11, fontweight='bold', color=NAVY)
ax.text(1, 6500+100, '6,500円', ha='center', va='bottom', fontproperties=JP_FONT, fontsize=11, fontweight='bold', color=NAVY)
ax.text(2, 6650+100, '【試算例】\n約6,650円', ha='center', va='bottom', fontproperties=JP_FONT_S, fontsize=9, color=GRAY, style='italic')
ax.text(3, 6500+100, '【試算例】\n約6,500円', ha='center', va='bottom', fontproperties=JP_FONT_S, fontsize=9, color=GRAY, style='italic')
ax.text(4, 6350-100, '【試算例】\n約6,350円', ha='center', va='top', fontproperties=JP_FONT_S, fontsize=9, color=GRAY, style='italic')

ax.set_xticks(x)
ax.set_xticklabels(periods, fontproperties=JP_FONT_S, fontsize=9)
ax.set_ylabel('月額基準額（円）', fontproperties=JP_FONT)
ax.set_title('図7-1　介護保険料月額基準額の推移（試算は仮置き・委員会協議後確定）',
             fontproperties=JP_FONT_T, color=NAVY, pad=15)
ax.set_ylim(6000, 7000)
ax.grid(axis='y', linestyle=':', alpha=0.4)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

# 凡例
ax.text(0.02, 0.97, '※第10期は基金取崩方針により3パターン試算\n　実数値はVer.2.0で確定',
        transform=ax.transAxes, fontproperties=JP_FONT_S, fontsize=9, color=ORANGE,
        verticalalignment='top', style='italic',
        bbox=dict(boxstyle='round,pad=0.5', facecolor='#FFF2CC', edgecolor=ORANGE, linewidth=1))

plt.tight_layout()
plt.savefig('/home/claude/kawasaki_work/chart_premium.png', dpi=150, bbox_inches='tight', facecolor='white')
plt.close()
print("✓ chart_premium.png")

# ===========================================================
# 6. 高齢化率比較（川崎町・宮城県・全国）
# ===========================================================
fig, ax = plt.subplots(figsize=(8, 5), dpi=150)

regions = ['川崎町', '宮城県平均', '全国平均']
aging_rates = [41.4, 28.5, 29.1]
colors_age = [ORANGE, BLUE, GRAY]

bars = ax.barh(regions, aging_rates, color=colors_age, edgecolor='white', linewidth=2)
for b, v in zip(bars, aging_rates):
    ax.text(v + 0.5, b.get_y()+b.get_height()/2, f'{v}%',
            va='center', fontproperties=JP_FONT, fontsize=13, fontweight='bold')

ax.set_xlabel('高齢化率（%）', fontproperties=JP_FONT)
ax.set_yticklabels(regions, fontproperties=JP_FONT_L, fontsize=12)
ax.set_title('図2-5　高齢化率比較（令和7年3月31日時点）',
             fontproperties=JP_FONT_T, color=NAVY, pad=15)
ax.set_xlim(0, 50)
ax.grid(axis='x', linestyle=':', alpha=0.4)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.invert_yaxis()

# 注釈
ax.text(0.98, 0.05, '川崎町は宮城県内35市町村中5位の高齢化率',
        transform=ax.transAxes, ha='right', fontproperties=JP_FONT_S, fontsize=10, color=ORANGE,
        style='italic', fontweight='bold',
        bbox=dict(boxstyle='round,pad=0.5', facecolor=LORANGE, edgecolor=ORANGE, linewidth=1))

plt.tight_layout()
plt.savefig('/home/claude/kawasaki_work/chart_aging.png', dpi=150, bbox_inches='tight', facecolor='white')
plt.close()
print("✓ chart_aging.png")

print("\n=== 全6図表の生成完了 ===")
import os
for f in ['chart_population.png','chart_recipient.png','chart_benefit.png',
          'chart_income.png','chart_premium.png','chart_aging.png']:
    path = f'/home/claude/kawasaki_work/{f}'
    if os.path.exists(path):
        sz = os.path.getsize(path)
        print(f"  {f}: {sz:,} bytes")
