"""
統合報告書用グラフ生成スクリプト
=====================================================

本スクリプトは、宮城県水道料金体系統一化検討報告書に掲載する
5種類のグラフを生成するものです。

生成される画像:
1. report_chart1_4teams.png       - 4団体料金水準比較
2. report_chart2_shiogama_model.png - 塩竈生産用水用 vs 女川町逓減制
3. report_chart3_scenarios.png    - 政策シナリオ別影響
4. report_chart4_recommendation.png - 塩竈モデル適用時の各団体影響
5. report_chart5_roadmap.png      - 実施ロードマップ

【フォント】
Noto Sans CJK JP または ipagothic を自動検出。
Ubuntu/Debian環境では以下でインストール可能:
  apt-get install fonts-noto-cjk fonts-ipafont-gothic
"""
import os
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '03_verification_scripts'))
import matplotlib.pyplot as plt
import matplotlib
import matplotlib.font_manager as fm
import numpy as np

from verify_all_rates import (onagawa, kesennuma, ishinomaki, shiogama, shiogama_production)

# 日本語フォント自動検出
jp_font_path = None
for f in fm.findSystemFonts():
    if 'NotoSansCJK' in f or 'ipag' in f.lower():
        jp_font_path = f
        break
if jp_font_path:
    matplotlib.rcParams['font.family'] = fm.FontProperties(fname=jp_font_path).get_name()
matplotlib.rcParams['axes.unicode_minus'] = False

outdir = os.path.join(os.path.dirname(__file__), 'charts')
os.makedirs(outdir, exist_ok=True)

# ============================================================
# Chart 1: 4団体料金水準比較
# ============================================================
def chart1_4teams():
    usage_labels = ['一般家庭\n20mm/20㎥', '小規模\n25mm/200㎥', '中規模\n50mm/1,000㎥',
                     '大規模\n75mm/3,000㎥', '超大口\n100mm/8,000㎥']
    on_vals = [onagawa(20, 20), onagawa(200, 25), onagawa(1000, 50),
                onagawa(3000, 75), onagawa(8000, 100)]
    ke_vals = [kesennuma(20, 20), kesennuma(200, 25), kesennuma(1000, 50),
                kesennuma(3000, 75), kesennuma(8000, 100)]
    isn_vals = [ishinomaki(20, 20), ishinomaki(200, 25), ishinomaki(1000, 50),
                ishinomaki(3000, 75), ishinomaki(8000, 100)]
    sh_vals = [shiogama(20, 20), shiogama(200, 25), shiogama(1000, 50),
                shiogama(3000, 75), shiogama(8000, 100)]

    fig, ax = plt.subplots(figsize=(13, 6.5))
    x = np.arange(len(usage_labels))
    w = 0.2
    ax.bar(x - 1.5 * w, on_vals, w, label='女川町', color='#70AD47')
    ax.bar(x - 0.5 * w, ke_vals, w, label='気仙沼市', color='#4472C4')
    ax.bar(x + 0.5 * w, isn_vals, w, label='石巻広域', color='#ED7D31')
    ax.bar(x + 1.5 * w, sh_vals, w, label='塩竈市(一般用)', color='#7030A0')
    ax.set_yscale('log')
    ax.set_ylabel('月額料金(円・対数軸)', fontsize=11)
    ax.set_title('県内臨海4団体 料金水準比較(条例・早見表逆算による確定値)', fontsize=13, fontweight='bold')
    ax.set_xticks(x)
    ax.set_xticklabels(usage_labels, fontsize=10)
    ax.legend(loc='upper left', fontsize=11)
    ax.grid(True, axis='y', alpha=0.3)
    plt.tight_layout()
    plt.savefig(f'{outdir}/report_chart1_4teams.png', dpi=140, bbox_inches='tight')
    plt.close()


# ============================================================
# Chart 2: 塩竈生産用水用の産業用優遇効果
# ============================================================
def chart2_shiogama_model():
    labels = ['中規模\n50mm/1,000㎥', '大規模\n75mm/3,000㎥', '超大口\n100mm/8,000㎥']
    on_prod = [onagawa(1000, 50), onagawa(3000, 75), onagawa(8000, 100)]
    sh_normal = [shiogama(1000, 50), shiogama(3000, 75), shiogama(8000, 100)]
    sh_prod = [shiogama_production(1000), shiogama_production(3000), shiogama_production(8000)]

    fig, ax = plt.subplots(figsize=(12, 6.5))
    x = np.arange(len(labels))
    w = 0.26
    ax.bar(x - w, on_prod, w, color='#70AD47', label='女川町(一般用・逓減制)')
    ax.bar(x, sh_prod, w, color='#9BC2E6', label='塩竈市(生産用水用)')
    ax.bar(x + w, sh_normal, w, color='#7030A0', label='塩竈市(一般用)')
    ax.set_ylabel('月額料金(円)', fontsize=11)
    ax.set_title('決定的発見 ― 塩竈市『生産用水用』は女川町逓減制とほぼ同水準', fontsize=13, fontweight='bold')
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=10)
    ax.legend(fontsize=11)
    ax.grid(True, axis='y', alpha=0.3)
    for i, (o, p, n) in enumerate(zip(on_prod, sh_prod, sh_normal)):
        ax.text(i - w, o + 50000, f'{o:,}', ha='center', fontsize=9)
        ax.text(i, p + 50000, f'{p:,}', ha='center', fontsize=9)
        ax.text(i + w, n + 50000, f'{n:,}', ha='center', fontsize=9, color='#4D1A7A')
    plt.tight_layout()
    plt.savefig(f'{outdir}/report_chart2_shiogama_model.png', dpi=140, bbox_inches='tight')
    plt.close()


# ============================================================
# Chart 3: 政策シナリオ別影響比較
# ============================================================
def chart3_scenarios():
    scenarios = ['現行', '単純統一\n(加重平均)', '塩竈モデル\n(生産用水用)',
                 '10年経過措置\n(初年度)', '15年経過措置\n(初年度)']
    current = 111980
    unified_simple = 327120
    shiogama_m = 115500
    trans_10 = current + (unified_simple - current) / 10
    trans_15 = current + (unified_simple - current) / 15

    vals = [current, unified_simple, shiogama_m, trans_10, trans_15]
    changes = [0, (unified_simple - current) / current * 100,
                (shiogama_m - current) / current * 100,
                (trans_10 - current) / current * 100,
                (trans_15 - current) / current * 100]

    fig, ax = plt.subplots(figsize=(13, 6.5))
    colors = ['#70AD47', '#C00000', '#548235', '#ED7D31', '#FFC000']
    bars = ax.bar(scenarios, vals, color=colors, width=0.6)
    ax.set_ylabel('月額料金(円)', fontsize=12)
    ax.set_title('女川町中規模水産加工業者(50mm/1,000㎥)― 政策シナリオ別影響比較',
                 fontsize=14, fontweight='bold')
    ax.grid(True, axis='y', alpha=0.3)
    ax.axhline(y=current, color='green', linestyle='--', alpha=0.5)

    for bar, val, chg in zip(bars, vals, changes):
        h = bar.get_height()
        label = f'{val:,.0f}円' + (f'\n({chg:+.0f}%)' if chg != 0 else '\n(基準)')
        col = 'black' if abs(chg) < 30 else ('darkorange' if chg < 100 else 'darkred')
        ax.text(bar.get_x() + bar.get_width() / 2, h + 8000, label,
                ha='center', fontsize=10, fontweight='bold', color=col)
    ax.set_ylim(0, max(vals) * 1.20)
    plt.tight_layout()
    plt.savefig(f'{outdir}/report_chart3_scenarios.png', dpi=140, bbox_inches='tight')
    plt.close()


# ============================================================
# Chart 4: 塩竈モデル統一時の各団体影響
# ============================================================
def chart4_recommendation():
    teams_jp = ['女川町\n(19社)', '気仙沼市\n(120社)', '石巻広域\n(180社)', '塩竈市\n(90社)']
    current_yr = [0.43, 8.34, 12.99, 5.96]

    def shiogama_model_yr(cnt):
        monthly_19 = 60500 * 5 + 115500 * 8 + 346500 * 5 + 924000 * 1
        return monthly_19 * 12 * (cnt / 19) / 1e8

    unified_yr = [shiogama_model_yr(19), shiogama_model_yr(120),
                   shiogama_model_yr(180), shiogama_model_yr(90)]

    fig, ax = plt.subplots(figsize=(12, 6))
    x = np.arange(len(teams_jp))
    w = 0.35
    ax.bar(x - w / 2, current_yr, w, color='#4472C4', label='現行体系(年額)')
    ax.bar(x + w / 2, unified_yr, w, color='#70AD47', label='塩竈モデル統一適用後(年額)')
    ax.set_ylabel('年額水道料金(億円)', fontsize=11)
    ax.set_title('推奨政策の効果 ― 塩竈モデル統一適用時の各団体影響(年額)',
                 fontsize=13, fontweight='bold')
    ax.set_xticks(x)
    ax.set_xticklabels(teams_jp, fontsize=10)
    ax.legend(fontsize=11)
    ax.grid(True, axis='y', alpha=0.3)
    for i, (c, u) in enumerate(zip(current_yr, unified_yr)):
        ax.text(i - w / 2, c + 0.2, f'{c:.2f}', ha='center', fontsize=10)
        ax.text(i + w / 2, u + 0.2, f'{u:.2f}', ha='center', fontsize=10)
        pct = (u - c) / c * 100 if c != 0 else 0
        sign = '+' if pct >= 0 else ''
        col = '#C00000' if pct > 10 else ('#548235' if pct < -1 else '#000000')
        ax.text(i, max(c, u) + 1.5, f'{sign}{pct:.1f}%', ha='center',
                fontsize=11, fontweight='bold', color=col)
    plt.tight_layout()
    plt.savefig(f'{outdir}/report_chart4_recommendation.png', dpi=140, bbox_inches='tight')
    plt.close()


# ============================================================
# Chart 5: 実施ロードマップ
# ============================================================
def chart5_roadmap():
    fig, ax = plt.subplots(figsize=(13, 5.5))
    ax.axis('off')

    milestones = [
        (2026, '基礎検討\n(本資料)', '#4472C4'),
        (2027, '広域化推進プラン\n条例案策定', '#ED7D31'),
        (2028, '統一体系決定\n塩竈モデル採用', '#70AD47'),
        (2030, '統一体系施行\n経過措置開始', '#C00000'),
        (2035, '中間レビュー', '#FFC000'),
        (2040, '経過措置終了\n完全統一', '#548235'),
    ]

    for year, label, color in milestones:
        x_pos = (year - 2025) / 15
        ax.plot([x_pos], [0.4], 'o', markersize=20, color=color,
                markeredgecolor='black', markeredgewidth=1.5,
                transform=ax.transAxes)
        ax.text(x_pos, 0.25, f'{year}年', ha='center', fontsize=11,
                fontweight='bold', transform=ax.transAxes)
        ax.text(x_pos, 0.55, label, ha='center', fontsize=10,
                transform=ax.transAxes)

    ax.annotate('', xy=(0.97, 0.4), xytext=(0.03, 0.4),
                arrowprops=dict(arrowstyle='->', color='black', lw=2),
                transform=ax.transAxes)

    ax.text(0.5, 0.85, '宮城県水道料金体系統一化 ― 推奨ロードマップ',
            ha='center', fontsize=14, fontweight='bold', transform=ax.transAxes)
    ax.text(0.5, 0.75, '塩竈モデル(生産用水用区分)導入 + 10年経過措置',
            ha='center', fontsize=11, color='#555555', transform=ax.transAxes)

    plt.tight_layout()
    plt.savefig(f'{outdir}/report_chart5_roadmap.png', dpi=140, bbox_inches='tight')
    plt.close()


if __name__ == "__main__":
    print("グラフ生成を開始します...")
    chart1_4teams()
    print("  ✓ chart1_4teams.png")
    chart2_shiogama_model()
    print("  ✓ chart2_shiogama_model.png")
    chart3_scenarios()
    print("  ✓ chart3_scenarios.png")
    chart4_recommendation()
    print("  ✓ chart4_recommendation.png")
    chart5_roadmap()
    print("  ✓ chart5_roadmap.png")
    print(f"\n全てのグラフを {outdir} に生成しました。")
