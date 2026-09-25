# -*- coding: utf-8 -*-
"""計画素案の図。規則は figs.py（策定委員会資料と共通）による。

　配色・書体・保存・ラベルの規則は figs.py から取り込み、ここでは図の中身だけを書く。
"""
import matplotlib.pyplot as plt

from figs import RED, BLUE, PALE, ORANGE, GREY, NAVY, _save, _lab, LEG, LEGF


# ══════════════ 第2章 ══════════════
def s1_jinko():
    """人口の推移と将来推計"""
    yr = ['Ｒ6', 'Ｒ7', 'Ｒ8', 'Ｒ9', 'Ｒ10', 'Ｒ11']
    wk = [10263, 10138, 10020, 9899, 9781, 9660]
    a = [2134, 2114, 2072, 2030, 1989, 1947]
    b = [2633, 2652, 2688, 2726, 2761, 2799]
    rate = [31.7, 32.0, 32.2, 32.5, 32.7, 32.9]
    fig, (ax, ax2) = plt.subplots(2, 1, figsize=(7.2, 4.2), sharex=True,
                                  gridspec_kw={'height_ratios': [3, 1],
                                               'hspace': 0.12})
    ax.bar(yr, wk, width=0.56, color='#E8E8E8', label='0〜64歳')
    ax.bar(yr, a, width=0.56, bottom=wk, color=PALE, label='65〜74歳')
    ax.bar(yr, b, width=0.56, bottom=[w + x for w, x in zip(wk, a)],
           color=BLUE, label='75歳以上')
    for i, (w, x, y) in enumerate(zip(wk, a, b)):
        ax.text(i, w + x + y + 240, '{:,}'.format(w + x + y), color=NAVY,
                ha='center', fontsize=9, fontweight='bold')
        ax.text(i, w + x + y / 2, '{:,}'.format(y), color='white',
                ha='center', va='center', fontsize=9, fontweight='bold')
        ax.text(i, w + x / 2, '{:,}'.format(x), color=NAVY,
                ha='center', va='center', fontsize=9)
    ax.set_ylim(0, 17600); ax.set_ylabel('人口（人）')
    LEGF(fig, ax, ncol=3)
    ax2.plot(yr, rate, '-o', color=RED, lw=2.4, ms=6)
    for i, v in enumerate(rate):
        ax2.text(i, v + 0.16, '%.1f％' % v, color=RED, ha='center',
                 fontsize=9, fontweight='bold')
    ax2.set_ylim(31.2, 33.6); ax2.set_ylabel('高齢化率（％）', fontsize=9)
    return _save(fig, 's1_jinko')


def s2_uchiwake():
    """高齢者人口の内訳の変化"""
    yr = ['Ｒ6', 'Ｒ7', 'Ｒ8', 'Ｒ9', 'Ｒ10', 'Ｒ11']
    a = [2134, 2114, 2072, 2030, 1989, 1947]
    b = [1571, 1578, 1630, 1684, 1735, 1789]
    c = [1062, 1074, 1058, 1042, 1026, 1010]
    fig, ax = plt.subplots(figsize=(7.2, 3.4))
    ax.bar(yr, a, width=0.56, color=PALE, label='65〜74歳')
    ax.bar(yr, b, width=0.56, bottom=a, color=BLUE, label='75〜84歳')
    ax.bar(yr, c, width=0.56, bottom=[x + y for x, y in zip(a, b)],
           color=RED, label='85歳以上')
    for i in range(6):
        ax.text(i, a[i] / 2, '{:,}'.format(a[i]), ha='center', va='center',
                color=NAVY, fontsize=9)
        ax.text(i, a[i] + b[i] / 2, '{:,}'.format(b[i]), ha='center', va='center',
                color='white', fontsize=9, fontweight='bold')
        ax.text(i, a[i] + b[i] + c[i] / 2, '{:,}'.format(c[i]), ha='center',
                va='center', color='white', fontsize=9)
    ax.set_ylim(0, 5600); ax.set_ylabel('第1号被保険者数（人）')
    LEG(ax, ncol=3)
    return _save(fig, 's2_uchiwake')


def s3_gyoseiku():
    """行政区別の高齢化率　―　区の数と高齢者の数は一致しない"""
    fig, ax = plt.subplots(figsize=(7.0, 2.2))
    rows = ['行政区（47区）', '高齢者（4,662人）']
    hi = [53.2, 42.2]          # 高齢化率40％以上の区
    lo = [46.8, 57.8]
    ax.barh(rows, hi, height=0.5, color=RED, label='高齢化率40％以上の区（25区）')
    ax.barh(rows, lo, height=0.5, left=hi, color=PALE, label='その他の区（22区）')
    for i, v in enumerate(hi):
        ax.text(v / 2, i, '%.1f％' % v, color='white', ha='center', va='center',
                fontsize=10, fontweight='bold')
        ax.text(v + lo[i] / 2, i, '%.1f％' % lo[i], color=NAVY, ha='center',
                va='center', fontsize=10, fontweight='bold')
    ax.invert_yaxis(); ax.set_xlim(0, 100)
    ax.set_xlabel('構成比（％）')
    LEG(ax, ncol=2, gap=0.42)
    return _save(fig, 's3_gyoseiku')


def s4_nintei_suii():
    """要介護度別認定者数の推移"""
    yr = ['Ｒ4/3末', 'Ｒ5/3末', 'Ｒ6/3末', 'Ｒ7/3末', 'Ｒ8/3末']
    seg = [('要支援1', [86, 92, 87, 93, 89], '#E8E8E8'),
           ('要支援2', [90, 103, 96, 85, 90], PALE),
           ('要介護1', [157, 154, 157, 154, 139], '#9DC3E6'),
           ('要介護2', [148, 140, 159, 184, 184], RED),
           ('要介護3', [119, 118, 123, 120, 119], BLUE),
           ('要介護4', [110, 86, 95, 99, 103], '#1F4E79'),
           ('要介護5', [85, 66, 60, 58, 66], GREY)]
    fig, ax = plt.subplots(figsize=(7.2, 3.8))
    bottom = [0] * 5
    for name, v, c in seg:
        ax.bar(yr, v, width=0.5, bottom=bottom, color=c, label=name)
        if name in ('要介護2',):
            for i in range(5):
                ax.text(i, bottom[i] + v[i] / 2, str(v[i]), ha='center',
                        va='center', color='white', fontsize=9, fontweight='bold')
        bottom = [b + x for b, x in zip(bottom, v)]
    for i, t in enumerate(bottom):
        ax.text(i, t + 14, '計%d人' % t, ha='center', color=NAVY,
                fontsize=9, fontweight='bold')
    ax.set_ylim(0, 960); ax.set_ylabel('認定者数（人）')
    LEG(ax, ncol=7, gap=0.10, fontsize=8)
    return _save(fig, 's4_nintei_suii')


def s5_ninchi_jiritsu():
    """認知症高齢者の自立度別の増減（令和2年4月末→令和6年10月末）"""
    lab = ['Ｍ', 'Ⅳ', 'Ⅲb', 'Ⅲa', 'Ⅱb', 'Ⅱa', 'Ⅰ', '自立']
    d = [5, -3, 13, -6, 93, -9, 34, 7]
    col = [BLUE if v > 0 else GREY for v in d]
    col[4] = RED
    fig, ax = plt.subplots(figsize=(6.8, 3.2))
    ax.barh(lab, d, height=0.56, color=col)
    for i, v in enumerate(d):
        ax.text(v + (2 if v >= 0 else -2), i, '%+d人' % v,
                color=RED if i == 4 else NAVY, ha='left' if v >= 0 else 'right',
                va='center', fontsize=9, fontweight='bold')
    ax.axvline(0, color='#808080', lw=1)
    ax.set_xlim(-22, 110)
    ax.set_xlabel('4年半の増減（人）')
    ax.text(100, 4, '増加分93人はすべてⅡb', color=RED, ha='right', va='bottom',
            fontsize=9, fontweight='bold')
    return _save(fig, 's5_ninchi_jiritsu')


def s6_kyufu():
    """サービス区分別の給付費の増減（令和6年度→令和7年度）"""
    lab = ['居宅サービス', '居宅介護支援等', '施設サービス', '地域密着型サービス']
    v = [10.6, 4.5, 1.0, 0.6]
    amt = [44128, 2678, 3715, 2094]
    col = [RED, BLUE, PALE, PALE]
    fig, ax = plt.subplots(figsize=(7.0, 2.8))
    ax.barh(lab, v, height=0.52, color=col)
    for i, x in enumerate(v):
        _lab(ax, x + 0.25, i, '＋%.1f％（＋%s千円）' % (x, '{:,}'.format(amt[i])),
             RED if i == 0 else NAVY, sz=9)
    ax.axvline(4.4, color=ORANGE, ls='--', lw=1.6)
    ax.text(4.5, -0.72, '保険給付費 全体＋4.4％', color=ORANGE, fontsize=9)
    ax.invert_yaxis(); ax.set_xlim(0, 20)
    ax.set_xlabel('給付費の増減率（％）')
    ax.xaxis.grid(True, color='#E6E6E6', lw=0.8); ax.set_axisbelow(True)
    return _save(fig, 's6_kyufu')


def s7_suikei():
    """計画期間の認定者数の推計"""
    yr = ['Ｒ8（基準）', 'Ｒ9', 'Ｒ10', 'Ｒ11']
    seg = [('要支援1・2', [179, 184, 186, 185], PALE),
           ('要介護1', [139, 142, 141, 143], '#9DC3E6'),
           ('要介護2', [184, 185, 184, 187], RED),
           ('要介護3以上', [288, 289, 292, 291], BLUE)]
    fig, ax = plt.subplots(figsize=(6.8, 3.4))
    bottom = [0] * 4
    for name, v, c in seg:
        ax.bar(yr, v, width=0.46, bottom=bottom, color=c, label=name)
        bottom = [b + x for b, x in zip(bottom, v)]
    for i, t in enumerate(bottom):
        ax.text(i, t + 12, '%d人' % t, ha='center', color=NAVY,
                fontsize=10, fontweight='bold')
    ax.set_ylim(0, 950); ax.set_ylabel('認定者数（第1号被保険者・人）')
    LEG(ax, ncol=4, gap=0.12)
    return _save(fig, 's7_suikei')


def s8_choki():
    """長期推計　―　令和12年度を100とした指数"""
    yr = ['Ｒ12', 'Ｒ15', 'Ｒ17', 'Ｒ22', 'Ｒ27', 'Ｒ30', 'Ｒ32']
    ser = [('第1号被保険者', [4740, 4691, 4659, 4786, 4823, 4751, 4703], GREY, '-o'),
           ('75歳以上', [2835, 2857, 2871, 2875, 2792, 2847, 2883], ORANGE, '-s'),
           ('85歳以上', [994, 1058, 1101, 1257, 1261, 1242, 1229], BLUE, '-^'),
           ('認定者数', [823, 842, 850, 915, 931, 919, 909], RED, '-o')]
    xs = [12, 15, 17, 22, 27, 30, 32]
    fig, ax = plt.subplots(figsize=(7.2, 3.6))
    for name, v, c, m in ser:
        idx = [x / v[0] * 100 for x in v]
        ax.plot(xs, idx, m, color=c, lw=2.6 if c == RED else 2, ms=6, label=name)
        _lab(ax, 32.4, idx[-1], '%.0f' % idx[-1], c, sz=9)
    ax.axhline(100, color='#BFBFBF', lw=1, ls='--')
    ax.set_xticks(xs); ax.set_xticklabels(yr)
    ax.set_xlim(11.2, 34.0); ax.set_ylim(92, 132)
    ax.set_ylabel('令和12年度＝100とした指数')
    LEG(ax, ncol=4, gap=0.14)
    return _save(fig, 's8_choki')


def s9_hokenryo():
    """介護保険料の見通し（低位・中位・高位）"""
    lab = ['第9期\n（実績）', '第10期\n（令和9〜11年度）', '令和15年度', '令和30年度']
    lo = [4900, 4996, 6300, 5930]
    mid = [4900, 5536, 6850, 7500]
    hi = [4900, 5781, 7550, 9910]
    fig, ax = plt.subplots(figsize=(6.8, 3.6))
    for i in range(4):
        if hi[i] > lo[i]:
            ax.plot([i, i], [lo[i], hi[i]], color=PALE, lw=16, solid_capstyle='butt')
            ax.text(i + 0.22, hi[i], '{:,}'.format(hi[i]), color=GREY,
                    fontsize=8, va='center')
            ax.text(i + 0.22, lo[i], '{:,}'.format(lo[i]), color=GREY,
                    fontsize=8, va='center')
        ax.plot([i], [mid[i]], marker='o', ms=10, color=RED if i else GREY)
        ax.text(i, mid[i] + 330, '{:,}円'.format(mid[i]),
                color=RED if i else GREY, ha='center', fontsize=10, fontweight='bold')
    ax.set_xticks(range(4)); ax.set_xticklabels(lab, fontsize=9)
    ax.set_xlim(-0.5, 3.7); ax.set_ylim(4000, 10800)
    ax.set_ylabel('第1号保険料基準額（月額・円）')
    ax.text(3.65, 9910, '高位', color=GREY, fontsize=8, ha='left', va='center')
    ax.text(3.65, 5930, '低位', color=GREY, fontsize=8, ha='left', va='center')
    return _save(fig, 's9_hokenryo')


# ══════════════ 第3章 ══════════════
def s10_seika():
    """第9期の成果指標　―　目標に対する達成状況"""
    lab = ['⑤配食サービス参画事業者数', '⑦生活支援サポーター養成（累計）',
           '⑧介護職員就職支援助成金（累計）', '⑥認知症サポーター養成（累計）',
           '④シルバー人材センター受託件数', '②オレンジカフェ箇所数']
    v = [150.0, 103.8, 90.3, 89.2, 76.4, 74.6]
    col = [BLUE if x >= 100 else RED for x in v]      # 達成＝青／未達＝赤
    fig, ax = plt.subplots(figsize=(7.2, 3.0))
    ax.barh(lab, v, height=0.54, color=col)
    for i, x in enumerate(v):
        _lab(ax, x + 1.5, i, '%.1f％' % x, col[i], sz=10)
    ax.axvline(100, color='#808080', ls='--', lw=1.6)
    ax.text(101, -0.78, '目標値', color='#808080', fontsize=9)
    ax.invert_yaxis(); ax.set_xlim(0, 175)
    ax.set_xlabel('令和7年度実績の目標比（％）')
    ax.xaxis.grid(True, color='#E6E6E6', lw=0.8); ax.set_axisbelow(True)
    return _save(fig, 's10_seika')


def s11_chiiki():
    """全国との地域差（地域差指数・令和5年）"""
    lab = ['第1号被保険者1人あたり\n給付月額', '要介護認定率', '受給率',
           '受給者1人あたり給付月額']
    kg = [0.97, 0.80, 0.85, 0.87]
    iw = [1.06, 0.95, 1.02, 1.04]
    y = range(len(lab))
    fig, ax = plt.subplots(figsize=(7.0, 3.0))
    ax.barh([i + 0.17 for i in y], kg, height=0.32, color=RED, label='金ケ崎町')
    ax.barh([i - 0.17 for i in y], iw, height=0.32, color=ORANGE, label='岩手県')
    for i, x in enumerate(kg):
        _lab(ax, x + 0.012, i + 0.17, '%.2f' % x, RED, sz=9)
    for i, x in enumerate(iw):
        _lab(ax, x + 0.012, i - 0.17, '%.2f' % x, ORANGE, sz=9)
    ax.axvline(1.00, color=GREY, ls='--', lw=1.6)
    ax.text(1.012, -0.80, '全国＝1.00', color=GREY, fontsize=9)
    ax.set_yticks(list(y)); ax.set_yticklabels(lab, fontsize=9)
    ax.invert_yaxis(); ax.set_xlim(0, 1.32)
    ax.set_xlabel('全国を1.00とした指数（性・年齢調整後）')
    LEG(ax, ncol=2, gap=0.24)
    return _save(fig, 's11_chiiki')


# ══════════════ 第4章 ══════════════
def s12_risk():
    """生活機能リスクの該当率"""
    lab = ['社会的役割の低下', 'うつ傾向', '認知機能の低下', '知的能動性の低下',
           '転倒リスク', '口腔機能の低下', '閉じこもり傾向', '運動器機能の低下',
           'ＩＡＤＬの低下', '低栄養傾向']
    v = [55.6, 46.6, 45.4, 37.5, 33.0, 25.8, 22.9, 12.2, 8.7, 0.7]
    col = [RED, RED, RED] + [BLUE] * 4 + [PALE] * 3
    fig, ax = plt.subplots(figsize=(7.0, 3.6))
    ax.barh(lab, v, height=0.58, color=col)
    for i, x in enumerate(v):
        _lab(ax, x + 0.9, i, '%.1f％' % x, RED if i < 3 else NAVY, sz=9)
    ax.invert_yaxis(); ax.set_xlim(0, 68)
    ax.set_xlabel('該当率（％・無回答を除いた分母）')
    ax.xaxis.grid(True, color='#E6E6E6', lw=0.8); ax.set_axisbelow(True)
    return _save(fig, 's12_risk')


def s13_fukugo():
    """複合リスクの重複数"""
    fig, ax = plt.subplots(figsize=(7.0, 1.7))
    seg = [('リスク0個', 21.9, 104, '#E8E8E8'), ('リスク1個', 24.9, 118, PALE),
           ('リスク2個', 22.6, 107, BLUE), ('リスク3個以上', 30.6, 145, RED)]
    left = 0
    for name, p, nnn, c in seg:
        ax.barh([0], [p], left=left, height=0.5, color=c)
        ax.text(left + p / 2, 0, '%s\n%.1f％（%d人）' % (name, p, nnn),
                ha='center', va='center', fontsize=9, fontweight='bold',
                color='white' if c in (BLUE, RED) else NAVY)
        left += p
    ax.set_yticks([]); ax.set_xlim(0, 100)
    ax.set_xlabel('回答者に占める割合（％・n=474）')
    for s in ('left', 'right', 'top'):
        ax.spines[s].set_visible(False)
    return _save(fig, 's13_fukugo')


def s14_shisaku():
    """住民が今後重点を置くべきと考える施策"""
    lab = ['病気・介護・認知症にならないための予防対策',
           '高齢者の外出を支援する移動手段の確保',
           'ホームヘルパーなどの在宅サービスの充実',
           '特別養護老人ホームなどの施設サービスの整備',
           '生活相談窓口の整備・充実']
    v = [49.4, 49.0, 37.4, 37.1, 31.9]
    col = [RED, RED, BLUE, BLUE, PALE]
    fig, ax = plt.subplots(figsize=(7.2, 2.6))
    ax.barh(lab, v, height=0.56, color=col)
    for i, x in enumerate(v):
        _lab(ax, x + 0.8, i, '%.1f％' % x, RED if i < 2 else NAVY, sz=10)
    ax.invert_yaxis(); ax.set_xlim(0, 62)
    ax.set_xlabel('回答率（％・複数回答 n=474）')
    ax.xaxis.grid(True, color='#E6E6E6', lw=0.8); ax.set_axisbelow(True)
    return _save(fig, 's14_shisaku')


# ══════════════ 第7章 ══════════════
def s15_seikatsu():
    """訪問介護員の年齢別に見た提供時間の構成比"""
    age = ['30歳代', '40歳代', '50歳代', '60歳代', '70歳以上']
    sin = [10.0, 14.6, 23.0, 32.5, 19.9]
    sei = [4.0, 10.6, 14.3, 31.0, 40.4]
    y = range(len(age))
    fig, ax = plt.subplots(figsize=(7.0, 2.9))
    ax.barh([i + 0.18 for i in y], sin, height=0.34, color=BLUE, label='身体介護')
    ax.barh([i - 0.18 for i in y], sei, height=0.34, color=RED, label='生活援助')
    for i, x in enumerate(sin):
        _lab(ax, x + 0.6, i + 0.18, '%.1f％' % x, BLUE, sz=9)
    for i, x in enumerate(sei):
        _lab(ax, x + 0.6, i - 0.18, '%.1f％' % x, RED, sz=9)
    ax.set_yticks(list(y)); ax.set_yticklabels(age, fontsize=10)
    ax.invert_yaxis(); ax.set_xlim(0, 52)
    ax.set_xlabel('提供時間の構成比（％）')
    ax.text(51, 4.55, '60歳以上　身体介護52.4％／生活援助71.4％', color=RED,
            ha='right', fontsize=9, fontweight='bold')
    LEG(ax, ncol=2, gap=0.24)
    return _save(fig, 's15_seikatsu')


def s16_jinzai():
    """介護職員の採用・離職とサービス系統別の昨年比"""
    lab = ['全サービス系統', '訪問系', '通所系', '施設・居住系']
    sai = [41, 11, 5, 25]
    ri = [49, 12, 4, 33]
    hi = [97.1, 98.7, 103.1, 95.1]
    y = range(len(lab))
    fig, (ax, ax2) = plt.subplots(1, 2, figsize=(7.6, 2.8),
                                  gridspec_kw={'width_ratios': [1.1, 1]})
    ax.barh([i + 0.18 for i in y], sai, height=0.34, color=BLUE, label='採用')
    ax.barh([i - 0.18 for i in y], ri, height=0.34, color=RED, label='離職')
    for i, x in enumerate(sai):
        _lab(ax, x + 0.8, i + 0.18, '%d人' % x, NAVY, sz=9)
    for i, x in enumerate(ri):
        _lab(ax, x + 0.8, i - 0.18, '%d人' % x, RED, sz=9)
    ax.set_yticks(list(y)); ax.set_yticklabels(lab, fontsize=9)
    ax.invert_yaxis(); ax.set_xlim(0, 62)
    ax.set_xlabel('令和7年度の人数（人）')
    LEG(ax, ncol=2, gap=0.26)
    ax2.barh(list(y), hi, height=0.5, color=PALE)
    for i, x in enumerate(hi):
        _lab(ax2, x + 0.4, i, '%.1f％%s' % (x, '　▲' if x < 100 else ''), NAVY, sz=9)
    ax2.axvline(100, color=ORANGE, ls='--', lw=1.6)
    ax2.text(109.6, -0.74, '▲は前年を下回る系統', color=NAVY, ha='right', fontsize=8)
    ax2.set_yticks(list(y)); ax2.set_yticklabels([])
    ax2.invert_yaxis(); ax2.set_xlim(88, 110)
    ax2.set_xlabel('職員数の昨年比（％）')
    return _save(fig, 's16_jinzai')


# ══════════════ 第8章 ══════════════
def s17_kanno():
    """保険料の水準を動かす論点（月額・円）"""
    lab = ['準備基金を9,000万円取り崩す', '後期高齢者加入割合補正係数',
           '令和8年度の見込量の置き方', '第1号被保険者負担割合（23→24％）',
           '地域密着型特養の稼働率が100％に回復', '令和9年度介護報酬改定率＋1.59％',
           'グループホームをワークシートどおり', '介護予防通所リハをワークシートどおり',
           '令和8年度の認定者数が813人であった場合', '訪問介護をワークシートどおり',
           '地域支援事業費を第9期と同額に据え置く']
    v = [-540, 262, 245, 241, 89, 84, -74, 58, 50, -22, -5]
    kind = ['町', '国', '事', '国', '事', '国', '事', '事', '事', '事', '町']
    col = {'町': RED, '国': ORANGE, '事': BLUE}
    fig, ax = plt.subplots(figsize=(7.4, 3.8))
    ax.barh(lab, v, height=0.56, color=[col[k] for k in kind])
    for i, x in enumerate(v):
        ax.text(x + (9 if x >= 0 else -9), i, '%+d円' % x, color=col[kind[i]],
                ha='left' if x >= 0 else 'right', va='center',
                fontsize=9, fontweight='bold')
    ax.axvline(0, color='#808080', lw=1)
    ax.invert_yaxis(); ax.set_xlim(-660, 400)
    ax.set_xlabel('基準額（月額）に与える影響（円）')
    hs = [plt.Rectangle((0, 0), 1, 1, color=col[k]) for k in ('町', '国', '事')]
    LEG(ax, hs, ['町のご判断', '国が決める', '事実の確認'], ncol=3, gap=0.14)
    return _save(fig, 's17_kanno')


ALL = [s1_jinko, s2_uchiwake, s3_gyoseiku, s4_nintei_suii, s5_ninchi_jiritsu,
       s6_kyufu, s7_suikei, s8_choki, s9_hokenryo, s10_seika, s11_chiiki,
       s12_risk, s13_fukugo, s14_shisaku, s15_seikatsu, s16_jinzai, s17_kanno]

if __name__ == '__main__':
    import sys
    import figs
    for f in ALL:
        print(f())
    if figs.ISSUES:
        print('\n図の指摘　%d件' % len(figs.ISSUES))
        for x in sorted(set(figs.ISSUES)):
            print('  ・' + x)
        sys.exit(1)
