# -*- coding: utf-8 -*-
"""策定委員会資料の図。第1回策定委員会資料の図の体裁に合わせる。

　体裁の規則（第1回資料の図1〜16から読み取ったもの）
　　配色　　強調・最大 #C00000／主 #2E75B6／淡 #BDD7EE／参照・基準 #ED7D31
　　　　　　減少・対象外 #A6A6A6／数値ラベル #1F3864（太字）
　　書体　　IPAGothic（本文のBIZ UDPゴシックに近いゴシック体）
　　形　　　順位・比較は横棒、推移は折れ線または縦棒
　　規則　　①値ラベルを必ず付す　②軸ラベルに単位を（）で入れる
　　　　　　③基準値は破線＋ラベル　④凡例は系列が2つ以上のときのみ
　　　　　　⑤背景は白、枠線あり、目盛線は用いない（横棒のみ縦の目盛線を薄く）
　　出力　　PNG・dpi200・幅はA4本文幅（15.5cm）に収まる比率
"""
import os

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager

FONT = 'IPAGothic'
plt.rcParams['font.family'] = FONT
plt.rcParams['axes.unicode_minus'] = False

RED = '#C00000'      # 強調・最大
BLUE = '#2E75B6'     # 主
PALE = '#BDD7EE'     # 淡
ORANGE = '#ED7D31'   # 参照・基準
GREY = '#A6A6A6'     # 減少・対象外
NAVY = '#1F3864'     # 数値ラベル

OUT = 'out/fig'
os.makedirs(OUT, exist_ok=True)


def _save(fig, name):
    p = '%s/%s.png' % (OUT, name)
    fig.savefig(p, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    return p


def _lab(ax, x, y, t, color=NAVY, ha='left', va='center', sz=10):
    ax.text(x, y, t, color=color, ha=ha, va=va, fontsize=sz, fontweight='bold')


# ── 図1　交付金の合計得点の推移 ──────────────────────────────
def f_suii():
    yr = ['令和6年度', '令和7年度', '令和8年度']
    kg = [393, 463, 434]
    iw = [399.4, 406.2, 406.3]
    nat = [422.3, 435.0, 455.1]
    fig, ax = plt.subplots(figsize=(7.2, 3.6))
    ax.plot(yr, nat, '-o', color=GREY, lw=2, ms=6, label='全国平均')
    ax.plot(yr, iw, '-s', color=ORANGE, lw=2, ms=6, label='岩手県平均')
    ax.plot(yr, kg, '-o', color=RED, lw=3, ms=8, label='金ケ崎町')
    for i, v in enumerate(kg):
        _lab(ax, i, v + 14, '%d' % v, RED, ha='center')
    for i, v in enumerate(nat):
        _lab(ax, i, v + 12, '%.1f' % v, GREY, ha='center', sz=9)
    for i, v in enumerate(iw):
        _lab(ax, i, v - 20, '%.1f' % v, ORANGE, ha='center', sz=9)
    ax.set_ylim(360, 490)
    ax.set_ylabel('合計得点（800点満点）')
    ax.legend(loc='lower right', frameon=False, fontsize=9, ncol=3)
    return _save(fig, 'f1_suii')


# ── 図2　県内33保険者の分布と本町の位置 ──────────────────────
def f_bunpu():
    lab = ['県内最大', '全国平均', '県内第3四分位', '金ケ崎町', '岩手県平均',
           '県内中央値', '県内第1四分位', '県内最小']
    val = [615, 455.1, 452.0, 434, 406.3, 390.0, 347.5, 305]
    col = [PALE, GREY, PALE, RED, ORANGE, BLUE, PALE, PALE]
    y = range(len(lab))
    fig, ax = plt.subplots(figsize=(7.0, 3.4))
    ax.barh(list(y), val, height=0.58, color=col)
    for i, v in enumerate(val):
        t = '%d' % v if float(v).is_integer() else '%.1f' % v
        _lab(ax, v + 6, i, t, col[i] if col[i] in (RED, ORANGE, GREY) else NAVY, sz=10)
    _lab(ax, 440, 3.0, '', RED)
    ax.set_yticks(list(y)); ax.set_yticklabels(lab, fontsize=10)
    ax.invert_yaxis(); ax.set_xlim(0, 700)
    ax.set_xlabel('令和8年度 交付金の合計得点（800点満点）')
    ax.xaxis.grid(True, color='#E6E6E6', lw=0.8); ax.set_axisbelow(True)
    ax.text(695, 3.0, '県内14位／33', color=RED, ha='right', va='center',
            fontsize=10, fontweight='bold')
    return _save(fig, 'f2_bunpu')


# ── 図3　目標別の比較 ────────────────────────────────────
def f_mokuhyo():
    lab = ['推進Ⅰ 持続可能な地域', '推進Ⅱ 公正・公平な給付\n（介護給付適正化）',
           '推進Ⅲ 介護人材の確保', '推進Ⅳ 成果指標群',
           '支援Ⅰ 介護予防／日常生活支援', '支援Ⅱ 認知症総合支援',
           '支援Ⅲ 在宅医療・在宅介護連携']
    kg = [70, 56, 46, 55, 54, 44, 54]
    iw = [54.1, 52.0, 45.2, 48.9, 53.8, 49.3, 54.0]
    nat = [62.3, 69.2, 50.8, 47.8, 57.8, 51.1, 68.3]
    y = range(len(lab))
    fig, ax = plt.subplots(figsize=(7.4, 5.2))
    h = 0.26
    ax.barh([i + h for i in y], kg, height=h,
            color=[RED if k < w else BLUE for k, w in zip(kg, iw)], label='金ケ崎町')
    ax.barh(list(y), iw, height=h, color=ORANGE, label='岩手県平均')
    ax.barh([i - h for i in y], nat, height=h, color=GREY, label='全国平均')
    for i, v in enumerate(kg):
        _lab(ax, v + 1, i + h, '%d' % v, RED if v < iw[i] else NAVY, sz=9)
    for i, v in enumerate(iw):
        _lab(ax, v + 1, i, '%.1f' % v, ORANGE, sz=8)
    for i, v in enumerate(nat):
        _lab(ax, v + 1, i - h, '%.1f' % v, GREY, sz=8)
    ax.set_yticks(list(y)); ax.set_yticklabels(lab, fontsize=9)
    ax.invert_yaxis()
    ax.set_xlim(0, 88)
    ax.set_xlabel('得点（各目標100点満点）')
    ax.xaxis.grid(True, color='#E6E6E6', lw=0.8); ax.set_axisbelow(True)
    hd, lb = ax.get_legend_handles_labels()
    ax.legend(hd[::-1], lb[::-1], loc='lower right', frameon=False, fontsize=9,
              bbox_to_anchor=(1.0, -0.02))
    return _save(fig, 'f3_mokuhyo')


# ── 図4　認知症総合支援100点の内訳 ──────────────────────────
def f_ninchi():
    lab = ['認知症サポーター数', '早期診断・早期対応の体制構築',
           '認知症地域支援推進員の業務の状況', 'サポーター等を活用した\n地域支援体制の構築',
           'ステップアップ講座修了者数', '難聴高齢者の\n早期発見・早期介入']
    hai = [12, 19, 12, 25, 12, 20]
    kg = [12, 14, 6, 12, 0, 0]
    nat = [4.8, 15.8, 6.0, 15.2, 2.5, 6.9]
    y = range(len(lab))
    fig, ax = plt.subplots(figsize=(7.4, 4.4))
    ax.barh(list(y), hai, height=0.62, color='#F2F2F2', edgecolor='#D9D9D9', label='配点')
    ax.barh([i + 0.15 for i in y], kg, height=0.3,
            color=[RED if k < n else BLUE for k, n in zip(kg, nat)], label='金ケ崎町')
    ax.barh([i - 0.17 for i in y], nat, height=0.3, color=GREY, label='全国平均')
    for i, v in enumerate(kg):
        _lab(ax, v + 0.4, i + 0.15, '%d' % v, RED if v < nat[i] else NAVY, sz=9)
    for i, v in enumerate(nat):
        _lab(ax, v + 0.4, i - 0.17, '%.1f' % v, GREY, sz=8)
    for i, v in enumerate(hai):
        _lab(ax, v + 0.4, i - 0.44, '配点%d' % v, '#808080', sz=8)
    ax.set_yticks(list(y)); ax.set_yticklabels(lab, fontsize=9)
    ax.invert_yaxis(); ax.set_xlim(0, 29)
    ax.set_xlabel('得点（介護保険保険者努力支援交付金 目標Ⅱ　計100点）')
    ax.xaxis.grid(True, color='#E6E6E6', lw=0.8); ax.set_axisbelow(True)
    hd, lb = ax.get_legend_handles_labels()
    ax.legend(hd[::-1], lb[::-1], loc='upper right', frameon=False, fontsize=9)
    return _save(fig, 'f4_ninchi')


# ── 図5　難聴20点の内訳 ─────────────────────────────────
def f_nancho():
    lab = ['ア　普及啓発', 'イ　早期発見', 'ウ　早期介入', 'エ　フォローアップ']
    nat = [3.44, 2.44, 0.73, 0.24]
    ken = [12, 12, 0, 0]          # 県内33保険者のうち得点している数
    y = range(len(lab))
    fig, ax = plt.subplots(figsize=(7.2, 2.9))
    ax.barh(list(y), [5] * 4, height=0.6, color='#F2F2F2', edgecolor='#D9D9D9',
            )
    ax.barh(list(y), nat, height=0.34, color=GREY)
    ax.text(4.92, -0.52, '配点 各5点', color='#808080', ha='right', fontsize=8)
    ax.text(3.40, -0.52, '全国平均', color=GREY, ha='right', fontsize=8)
    for i, v in enumerate(nat):
        _lab(ax, v + 0.08, i, '%.2f' % v, GREY, sz=9)
    for i in y:
        ax.plot([0.02], [i], marker='|', ms=16, color=RED, mew=3)
        _lab(ax, 5.25, i, '本町 0点　／　県内で得点 %d保険者' % ken[i],
             RED if ken[i] else '#808080', sz=9)
    ax.set_yticks(list(y)); ax.set_yticklabels(lab, fontsize=10)
    ax.invert_yaxis(); ax.set_xlim(0, 9.6)
    ax.set_xticks([0, 1, 2, 3, 4, 5])
    ax.set_xlabel('得点（難聴高齢者の早期発見・早期介入　計20点）')
    ax.set_ylim(3.62, -0.78)
    return _save(fig, 'f5_nancho')


# ── 図6　要介護認定率の推移 ───────────────────────────────
def f_nintei():
    yr = ['Ｒ4/3末', 'Ｒ5/3末', 'Ｒ6/3末', 'Ｒ7/3末', 'Ｒ8/3末']
    kg = [16.7, 16.1, 16.4, 16.9, 16.9]
    iw = [19.4, 19.3, 19.5, 19.7, 20.0]
    nat = [18.9, 19.0, 19.4, 19.7, 20.2]
    fig, ax = plt.subplots(figsize=(7.2, 3.3))
    ax.plot(yr, nat, '-o', color=GREY, lw=2, ms=5, label='全国')
    ax.plot(yr, iw, '-s', color=ORANGE, lw=2, ms=5, label='岩手県')
    ax.plot(yr, kg, '-o', color=RED, lw=3, ms=7, label='金ケ崎町')
    _lab(ax, 4.05, nat[-1], '%.1f％' % nat[-1], GREY, sz=9)
    _lab(ax, 4.05, iw[-1] - 0.25, '%.1f％' % iw[-1], ORANGE, sz=9)
    _lab(ax, 4.05, kg[-1], '%.1f％' % kg[-1], RED, sz=10)
    ax.set_ylim(15.0, 21.5); ax.set_xlim(-0.3, 4.9)
    ax.set_ylabel('要支援・要介護認定率（％）')
    ax.legend(loc='center left', frameon=False, fontsize=9)
    return _save(fig, 'f6_nintei')


# ── 図7　要介護1の重度化率の推移 ────────────────────────────
def f_judoka():
    yr = ['令和2年度', '令和4年度', '令和6年度', '令和7年度']
    v = [24.1, 41.5, 45.2, 57.0]
    fig, ax = plt.subplots(figsize=(6.4, 3.2))
    ax.bar(yr, v, width=0.52, color=[PALE, BLUE, BLUE, RED])
    for i, y in enumerate(v):
        ax.text(i, y + 1.2, '%.1f％' % y, color=RED if i == 3 else NAVY,
                ha='center', fontsize=10, fontweight='bold')
    ax.axhline(50, color=ORANGE, ls='--', lw=1.6)
    ax.text(3.45, 51, '更新の半数', color=ORANGE, ha='right', fontsize=9)
    ax.set_ylim(0, 66)
    ax.set_ylabel('要介護1の更新時に重度化した割合（％）')
    return _save(fig, 'f7_judoka')


# ── 図8　地域支援事業費（総合事業費）の見込み ────────────────────
def f_chiiki():
    yr = ['第9期\n各年度', '令和9年度', '令和10年度', '令和11年度']
    v = [45104, 45742, 46329, 46967]
    fig, ax = plt.subplots(figsize=(6.6, 3.2))
    ax.bar(yr, v, width=0.5, color=[GREY, PALE, BLUE, BLUE])
    for i, y in enumerate(v):
        ax.text(i, y + 90, '{:,}'.format(y), color=NAVY if i else GREY,
                ha='center', fontsize=10, fontweight='bold')
    ax.annotate('', xy=(3, 46967), xytext=(0, 45104),
                arrowprops=dict(arrowstyle='->', color=RED, lw=1.6))
    ax.text(1.5, 46700, '3年間で＋1,863千円（＋4.13％）', color=RED,
            ha='center', fontsize=9, fontweight='bold')
    ax.set_ylim(44400, 47600)
    ax.set_ylabel('介護予防・日常生活支援総合事業費（千円）')
    return _save(fig, 'f8_chiiki')


ALL = [f_suii, f_bunpu, f_mokuhyo, f_ninchi, f_nancho, f_nintei, f_judoka, f_chiiki]

if __name__ == '__main__':
    for f in ALL:
        print(f())
