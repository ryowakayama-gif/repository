# -*- coding: utf-8 -*-
"""計画案の概要版（住民向け・A4 16ページ）の図。

　配色・書体・保存の規則は figs.py による。
　計画書の図26点のうち概要版に載せるのは5点であり、本ファイルは
　新たに作る3点を置く（パブリックコメント・住民説明会 資料 第1部の２）。

　　g1_taikei    施策の体系図（p10-11）
　　g2_kyufu     3年間の総事業費の構成（p14）
　　g3_zaigen    介護保険の財源と介護保険料（p15）

　g3 の基準額は介護給付費準備基金の取崩額により確定する。
　取崩額が決まった時点で作り直す。
"""
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

from figs import RED, BLUE, PALE, ORANGE, GREY, NAVY, _save


def _box(ax, x, y, w, h, text, face, edge, tc='white', sz=9, bold=True):
    ax.add_patch(FancyBboxPatch((x, y), w, h,
                                boxstyle='round,pad=0.012,rounding_size=0.02',
                                facecolor=face, edgecolor=edge, linewidth=1.0))
    ax.text(x + w / 2, y + h / 2, text, ha='center', va='center',
            color=tc, fontsize=sz, fontweight='bold' if bold else 'normal',
            linespacing=1.35)


def _arrow(ax, x1, y1, x2, y2, color=GREY):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle='-|>',
                                 mutation_scale=11, color=color, linewidth=1.1,
                                 shrinkA=0, shrinkB=0))


# ══════════════ 図Ａ　施策の体系図 ══════════════
HOSHIN = [
    ('基本方針１\n健康づくり・\n介護予防の推進', [
        ('(1) 高齢者の社会参加の支援', ''),
        ('(2) 健康づくり・介護予防の推進', ''),
        ('(3) 要介護1・2層の重度化防止', '重点1'),
        ('(4) 介護予防の評価の仕組みの構築', '重点3'),
        ('(5) 生活援助の担い手の確保', '重点2・4'),
    ]),
    ('基本方針２\n安心して暮らせる\n地域づくり', [
        ('(1) 在宅生活の支援', ''),
        ('(2) 高齢者の住まいかたの支援', ''),
        ('(3) 高齢者の権利擁護の支援', ''),
        ('(4) 移動支援と外出機会の確保', '重点2'),
    ]),
    ('基本方針３\n認知症施策の推進\n【認知症施策推進計画】', [
        ('(1) 認知症の人を支える地域環境づくり', ''),
        ('(2) 認知症の人や家族への支援体制の強化', ''),
        ('(3) 認知症の予防・早期発見と「新しい認知症観」の普及', ''),
    ]),
    ('基本方針４\n地域包括ケア\nシステムの深化', [
        ('(1) 多職種（医療・介護など）の連携', ''),
        ('(2) 地域包括支援センターの機能強化', ''),
        ('(3) 多様な生活支援サービスの推進', ''),
        ('(4) 家族介護者への支援', '重点2'),
    ]),
    ('基本方針５\n介護保険制度の\n円滑な運営', [
        ('(1) 介護保険事業の推進', ''),
        ('(2) 情報提供体制の充実', ''),
        ('(3) 災害や感染症対策に係る体制整備', ''),
        ('(4) 介護給付適正化の推進', '法定'),
    ]),
]


def g1_taikei():
    """施策の体系図　―　基本理念・5つの基本方針・20の施策"""
    n = sum(len(v) for _, v in HOSHIN)
    assert n == 20, n
    fig, ax = plt.subplots(figsize=(7.4, 7.6))
    ax.set_xlim(0, 10); ax.set_ylim(-0.45, 10.6); ax.axis('off')

    # 基本理念
    _box(ax, 0.3, 9.7, 9.4, 0.72,
         '基本理念　支え合い　健やかに　自分らしく暮らせるまち　かねがさき',
         NAVY, NAVY, sz=11)

    top, bot = 9.35, 0.5
    cur = top
    gap_h = 0.18
    for hi, (name, shisaku) in enumerate(HOSHIN):
        h = len(shisaku) * 0.42
        y = cur - h
        _box(ax, 0.3, y, 2.5, h, name, BLUE, BLUE, sz=9)
        for si, (nm_s, mark) in enumerate(shisaku):
            sy = cur - 0.42 * (si + 1)
            juten = mark.startswith('重点')
            face = 'white' if juten else '#F2F2F2'
            edge = RED if juten else GREY
            tc = RED if juten else NAVY
            _box(ax, 3.1, sy + 0.04, 5.3, 0.34, nm_s, face, edge, tc=tc,
                 sz=8.2, bold=juten)
            if mark:
                _box(ax, 8.55, sy + 0.06, 1.15, 0.30, mark,
                     RED if juten else GREY, RED if juten else GREY,
                     tc='white', sz=8.0)
            _arrow(ax, 2.8, sy + 0.21, 3.1, sy + 0.21,
                   RED if juten else GREY)
        cur = y - gap_h
    assert cur > bot - 0.5, cur

    ax.text(9.7, -0.38, '「重点」は計画書が重点課題として掲げた施策、'
            '「法定」は介護保険法が定める記載事項',
            ha='right', va='bottom', color=NAVY, fontsize=8.2,
            fontweight='bold')
    return _save(fig, 'g1_taikei')


# ══════════════ 図Ｂ　総事業費の構成 ══════════════
def g2_kyufu():
    """3年間の総事業費の構成（令和9〜11年度）"""
    item = ['介護給付費', 'その他給付費', '地域支援事業費', '予防給付費']
    val = [3709668, 233984, 199237, 49872]      # 千円
    col = [BLUE, PALE, ORANGE, GREY]
    tot = sum(val)
    assert tot == 4192761, tot

    fig, ax = plt.subplots(figsize=(7.2, 3.3))
    left = 0
    for v, c, nm in zip(val, col, item):
        ax.barh(0, v, left=left, height=0.52, color=c,
                edgecolor='white', linewidth=1.2)
        left += v
    ax.set_xlim(0, tot * 1.005); ax.set_ylim(-1.45, 0.95)
    ax.axis('off')

    # 大きい1つは帯の中、小さい3つは帯の下に並べる
    ax.text(val[0] / 2, 0, '介護給付費\n37.1億円（%.1f％）' % (val[0] / tot * 100),
            ha='center', va='center', color='white', fontsize=10,
            fontweight='bold', linespacing=1.3)
    y = -0.62
    for i in (1, 2, 3):
        ax.add_patch(plt.Rectangle((0, y - 0.07), tot * 0.016, 0.15,
                                   color=col[i], clip_on=False))
        ax.text(tot * 0.026, y,
                '%s　%.1f億円（%.1f％）' % (item[i], val[i] / 1e5, val[i] / tot * 100),
                ha='left', va='center', color=NAVY, fontsize=9.4,
                fontweight='bold')
        y -= 0.30
    ax.text(0, 0.62, '3年間（令和9〜11年度）の総事業費　41億9,276万円',
            ha='left', va='bottom', color=NAVY, fontsize=11, fontweight='bold')
    return _save(fig, 'g2_kyufu')


# ══════════════ 図Ｃ　財源と介護保険料 ══════════════
def g3_zaigen():
    """介護保険の財源と介護保険料"""
    fig, ax = plt.subplots(figsize=(7.2, 5.0))
    ax.set_xlim(0, 10); ax.set_ylim(0, 10); ax.axis('off')

    # 上段　総事業費と財源の割合
    ax.text(0.3, 9.6, '３年間の総事業費　41億9,276万円', color=NAVY,
            fontsize=11, fontweight='bold', va='bottom')
    seg = [('第1号被保険者\n（65歳以上）の\n保険料　23％', 23, BLUE, 'white'),
           ('第2号被保険者\n（40〜64歳）の\n保険料　27％', 27, PALE, NAVY),
           ('国　25％', 25, '#D9D9D9', NAVY),
           ('県\n12.5％', 12.5, '#E6E6E6', NAVY),
           ('町\n12.5％', 12.5, '#F2F2F2', NAVY)]
    x = 0.3
    w_all = 9.4
    for nm, p, c, tc in seg:
        w = w_all * p / 100
        _box(ax, x, 8.2, w, 1.25, nm, c, GREY if c != BLUE else BLUE,
             tc=tc, sz=8.4)
        x += w
    ax.text(9.7, 8.0, '保険料で50％、公費で50％を負担する', color=GREY,
            fontsize=8.6, va='top', ha='right')

    # 中段　第1号被保険者の保険料の算定
    _arrow(ax, 1.4, 8.1, 1.4, 7.2, BLUE)
    rows = [('第1号被保険者負担分相当額\n（総事業費 × 23％）', '9億6,433万円'),
            ('＋　調整交付金相当額', '2億0,663万円'),
            ('−　調整交付金見込額', '2億8,396万円'),
            ('＝　保険料収納必要額', '8億8,700万円')]
    y = 6.6
    for nm, v in rows:
        hl = nm.startswith('＝')
        _box(ax, 0.3, y, 5.4, 0.78, nm, 'white' if hl else '#F2F2F2',
             RED if hl else GREY, tc=RED if hl else NAVY, sz=9)
        ax.text(6.0, y + 0.39, v, color=RED if hl else NAVY,
                fontsize=10.5 if hl else 10, fontweight='bold', va='center')
        y -= 0.92

    # 下段　基準額
    _arrow(ax, 3.0, 3.80, 3.0, 3.12, RED)
    _box(ax, 0.3, 2.05, 9.4, 1.0,
         '保険料収納必要額 8億8,700万円 ÷ 166,663人・月'
         '\n＝　基準額（月額）　5,322円',
         RED, RED, sz=11)
    ax.text(0.3, 1.82,
            '分母は、所得段階別加入割合補正後被保険者数 13,920.56人 × 12か月 × 予定保険料収納率 99.77％',
            color=GREY, fontsize=7.4, va='top')
    _box(ax, 0.3, 0.62, 9.4, 0.78,
         '介護給付費準備基金を取り崩すと、この額から引き下げられる',
         'white', ORANGE, tc=ORANGE, sz=9.4)
    ax.text(0.3, 0.38, '取崩額は町のご判断により決まる', color=GREY,
            fontsize=8.2, va='top')
    return _save(fig, 'g3_zaigen')


ALL = [g1_taikei, g2_kyufu, g3_zaigen]

if __name__ == '__main__':
    for f in ALL:
        print(f())
