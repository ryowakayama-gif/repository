# -*- coding: utf-8 -*-
"""第10期計画 第2章 図表の作成
   成果品はモノクロ印刷（仕様書）のため、色相ではなく
   明度差＋線種＋マーカー形状＋ハッチングの二重符号化で系列を識別する。
   出力: output/figures/*.png（300dpi）"""
import os as _os_p
import sys as _sys_p
_sys_p.path.insert(0, _os_p.path.dirname(_os_p.path.abspath(__file__)))
import paths as _P   # 置き場所はここで決める（じか書きしない）
import sys
sys.dont_write_bytecode = True   # 古い .pyc で古い成果品ができるのを防ぐ
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

OUT = _P.FIGURES
os.makedirs(OUT, exist_ok=True)

# 数値の正本。台帳（xlsx）が置かれていればその値が優先される
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import data_zuhyo as DZ
_Z = DZ.load()


def V(name, series):
    """図の1系列の値。data_zuhyo.py（または台帳）から引く"""
    for n, v in _Z[name]["series"]:
        if n == series:
            return list(v)
    raise KeyError("%s に系列 %s がない" % (name, series))

plt.rcParams.update({
    "font.family": "IPAGothic",
    "font.size": 10,
    "axes.titlesize": 12,
    "axes.labelsize": 10,
    "axes.edgecolor": "#808080",
    "axes.linewidth": 0.8,
    "axes.grid": True,
    "grid.color": "#D9D9D9",
    "grid.linewidth": 0.6,
    "axes.axisbelow": True,
    "figure.facecolor": "white",
    "axes.facecolor": "white",
    "legend.frameon": False,
    "xtick.color": "#404040",
    "ytick.color": "#404040",
    "axes.labelcolor": "#262626",
    "text.color": "#262626",
})

# ── 明度ランプ（モノクロ印刷で確実に分離する4段階＋補助）──
K = {"d": "#1A1A1A", "m": "#595959", "l": "#A6A6A6", "xl": "#D9D9D9", "w": "#FFFFFF"}
# 系列スタイル（村＝実線・丸／県＝破線・四角／全国＝点線・三角）
S_MURA = dict(color=K["d"], ls="-",  marker="o", ms=6, lw=2.2, zorder=5)
S_KEN  = dict(color=K["m"], ls="--", marker="s", ms=5.5, lw=1.8, zorder=4)
S_ZEN  = dict(color=K["l"], ls=":",  marker="^", ms=6, lw=1.8, zorder=3)

def save(fig, name):
    fig.tight_layout(pad=0.6)
    p = os.path.join(OUT, name)
    fig.savefig(p, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("  ", name)

def style_ax(ax, ylab=None, xlab=None):
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="x", visible=False)
    if ylab: ax.set_ylabel(ylab)
    if xlab: ax.set_xlabel(xlab)

def label_last(ax, xs, ys, text, dy=0, color=K["d"], ha="left"):
    ax.annotate(text, (xs[-1], ys[-1]), xytext=(6, dy), textcoords="offset points",
                fontsize=9, color=color, va="center", ha=ha, fontweight="bold")

# ══════════ 図2-1 人口ピラミッド（2010／2025）══════════
def fig_2_1():
    labels = ["15歳未満", "15〜39歳", "40〜64歳", "65〜74歳", "75歳以上"]
    y2010 = V("fig2-1_人口構成比較", "平成22年（2010年）")
    y2025 = V("fig2-1_人口構成比較", "令和7年（2025年）")
    fig, ax = plt.subplots(figsize=(7.2, 3.6))
    idx = range(len(labels))
    h = 0.38
    b1 = ax.barh([i + h/2 for i in idx], y2010, height=h, color=K["l"],
                 edgecolor=K["m"], linewidth=0.8, label="平成22年（2010年）")
    b2 = ax.barh([i - h/2 for i in idx], y2025, height=h, color=K["d"],
                 edgecolor=K["d"], linewidth=0.8, label="令和7年（2025年）",
                 hatch="///")
    for bars in (b1, b2):
        for b in bars:
            ax.annotate(f"{int(b.get_width()):,}", (b.get_width(), b.get_y()+b.get_height()/2),
                        xytext=(4, 0), textcoords="offset points", va="center", fontsize=8.5)
    ax.set_yticks(list(idx)); ax.set_yticklabels(labels)
    ax.set_xlim(0, 1400)
    style_ax(ax, xlab="人口（人）")
    ax.grid(axis="x", visible=True); ax.grid(axis="y", visible=False)
    ax.legend(loc="lower right", fontsize=9)
    ax.set_title("年齢階級別人口の比較（平成22年・令和7年）", loc="left", pad=10)
    save(fig, "fig2-1_人口構成比較.png")

# ══════════ 図2-2 高齢化率の推移（村・県・全国）══════════
def fig_2_2():
    x = [2010, 2015, 2020, 2025, 2030, 2035, 2040, 2045, 2050]
    mura = V("fig2-2_高齢化率推移", "北塩原村")
    ken = V("fig2-2_高齢化率推移", "福島県")
    zen = V("fig2-2_高齢化率推移", "全国")
    fig, ax = plt.subplots(figsize=(7.2, 3.8))
    ax.axvspan(2027, 2029, color="#F2F2F2", zorder=0)
    ax.annotate("本計画期間\n（令和9〜11年度）", (2028, 21.2), ha="center", fontsize=8.5,
                color=K["m"], va="bottom")
    ax.plot(x, mura, label="北塩原村", **S_MURA)
    ax.plot(x, ken,  label="福島県",   **S_KEN)
    ax.plot(x, zen,  label="全国",     **S_ZEN)
    label_last(ax, x, mura, "49.2%")
    label_last(ax, x, ken,  "44.2%", color=K["m"])
    label_last(ax, x, zen,  "37.1%", color="#808080")
    ax.set_xlim(2008, 2056); ax.set_ylim(20, 54)
    ax.set_xticks(x); ax.set_xticklabels([str(v) for v in x], fontsize=9)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{v:.0f}%"))
    style_ax(ax, ylab="高齢化率")
    ax.legend(loc="upper left", fontsize=9, ncols=3)
    ax.set_title("高齢化率の推移と将来推計（村・福島県・全国）", loc="left", pad=10)
    save(fig, "fig2-2_高齢化率推移.png")

# ══════════ 図2-3 将来推計人口（年齢3区分）══════════
def fig_2_3():
    x = [2010, 2015, 2020, 2025, 2030, 2035, 2040, 2045, 2050]
    u15 = V("fig2-3_将来推計人口", "15歳未満")
    prod = V("fig2-3_将来推計人口", "15〜64歳")
    o65 = V("fig2-3_将来推計人口", "65歳以上")
    fig, ax = plt.subplots(figsize=(7.2, 3.8))
    ax.stackplot(x, u15, prod, o65,
                 labels=["15歳未満", "生産年齢人口（15〜64歳）", "高齢者人口（65歳以上）"],
                 colors=[K["xl"], K["l"], K["d"]],
                 edgecolor="white", linewidth=2,
                 hatch=[None, "..", "///"])
    tot = V("fig2-3_将来推計人口", "総人口")
    for xi, ti in zip(x, tot):
        ax.annotate(f"{ti:,}", (xi, ti), xytext=(0, 5), textcoords="offset points",
                    ha="center", fontsize=8.5, color=K["d"])
    ax.set_xlim(2010, 2050); ax.set_ylim(0, 3600)
    ax.set_xticks(x); ax.set_xticklabels([str(v) for v in x], fontsize=9)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{int(v):,}"))
    style_ax(ax, ylab="人口（人）")
    ax.legend(loc="upper right", fontsize=9)
    ax.set_title("将来推計人口（年齢3区分）", loc="left", pad=10)
    ax.annotate("※ 令和2年は年齢不詳1人を含むため、年齢区分の合計と総人口が1人異なります。",
                (0, -0.16), xycoords="axes fraction", fontsize=8, color=K["m"])
    save(fig, "fig2-3_将来推計人口.png")

# ══════════ 図2-4 前期／後期高齢者の推移 ══════════
def fig_2_4():
    x = [2010, 2015, 2020, 2025, 2030, 2035, 2040, 2045, 2050]
    zen_ki = V("fig2-4_前期後期高齢者", "65〜74歳")
    kou_ki = V("fig2-4_前期後期高齢者", "75歳以上")
    fig, ax = plt.subplots(figsize=(7.2, 3.8))
    ax.axvspan(2027, 2029, color="#F2F2F2", zorder=0)
    ax.annotate("本計画期間", (2028, 168), ha="center", fontsize=8.5, color=K["m"], va="bottom")
    ax.plot(x, kou_ki, label="75歳以上", color=K["d"], ls="-", marker="o", ms=6, lw=2.2, zorder=5)
    ax.plot(x, zen_ki, label="65〜74歳", color=K["m"], ls="--", marker="s", ms=5.5, lw=1.8, zorder=4)
    ax.annotate("ピーク 551人\n（令和12年）", (2030, 551), xytext=(0, 18),
                textcoords="offset points", ha="center", fontsize=9, fontweight="bold",
                color=K["d"])
    ax.plot([2030], [551], marker="o", ms=11, mfc="none", mec=K["d"], mew=1.8, zorder=6)
    label_last(ax, x, kou_ki, "398人")
    label_last(ax, x, zen_ki, "221人", color=K["m"])
    ax.set_xlim(2008, 2056); ax.set_ylim(150, 640)
    ax.set_xticks(x); ax.set_xticklabels([str(v) for v in x], fontsize=9)
    style_ax(ax, ylab="人口（人）")
    ax.legend(loc="lower left", fontsize=9)
    ax.set_title("前期高齢者と後期高齢者の推移（75歳以上は令和12年がピーク）",
                 loc="left", pad=10)
    save(fig, "fig2-4_前期後期高齢者.png")

if __name__ == "__main__":
    print("第2章 図の作成:")
    fig_2_1(); fig_2_2(); fig_2_3(); fig_2_4()

# ══════════ 図2-5 認定率の推移（村・県・全国）══════════
def fig_2_5():
    lab = ["令和2年\n3月末", "令和3年\n3月末", "令和4年\n3月末", "令和5年\n3月末",
           "令和6年\n3月末", "令和7年\n3月末", "令和8年\n3月末"]
    x = list(range(len(lab)))
    mura = V("fig2-5_認定率推移", "北塩原村")
    ken = V("fig2-5_認定率推移", "福島県")
    zen = V("fig2-5_認定率推移", "全国")
    fig, ax = plt.subplots(figsize=(7.2, 3.8))
    ax.plot(x, mura, label="北塩原村", **S_MURA)
    ax.plot(x, ken,  label="福島県",   **S_KEN)
    ax.plot(x, zen,  label="全国",     **S_ZEN)
    ax.axvline(4, color=K["m"], lw=0.9, ls="-", zorder=1)
    ax.annotate("令和6年3月末に\n県・全国を上回る", (4, 17.2), ha="center", va="bottom",
                fontsize=8.5, color=K["d"], fontweight="bold")
    label_last(ax, x, mura, "21.1%")
    label_last(ax, x, ken,  "19.8%", dy=-9, color=K["m"])
    label_last(ax, x, zen,  "20.2%", dy=7, color="#808080")
    ax.set_xlim(-0.4, 6.9); ax.set_ylim(16.5, 22)
    ax.set_xticks(x); ax.set_xticklabels(lab, fontsize=8.5)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{v:.0f}%"))
    style_ax(ax, ylab="認定率")
    ax.legend(loc="upper left", fontsize=9, ncols=3)
    ax.set_title("要支援・要介護認定率の推移（村・福島県・全国）", loc="left", pad=10)
    save(fig, "fig2-5_認定率推移.png")

# ══════════ 性・年齢調整済み認定率の推移 ══════════
def fig_2_21():
    """全国の人口構成に合わせて年齢構成の違いを取り除いた認定率。

    粗い認定率（図2-5）は高齢化が進むだけでも上がる。調整済みの認定率で
    なお上がっているかどうかが、自立支援・重度化防止の成果を表す。
    国の評価指標（支援Ⅰ）が求める分析にあたる。
    """
    lab = ["令和2年\n3月末", "令和3年\n3月末", "令和4年\n3月末", "令和5年\n3月末",
           "令和6年\n3月末", "令和7年\n3月末", "令和8年\n3月末"]
    x = list(range(len(lab)))
    tot = V("fig2-21_調整済み認定率", "合計")
    kei = V("fig2-21_調整済み認定率", "軽度（要支援1〜要介護2）")
    ni2 = V("fig2-21_調整済み認定率", "要介護2以上")
    juu = V("fig2-21_調整済み認定率", "重度（要介護3以上）")
    fig, ax = plt.subplots(figsize=(7.2, 4.0))
    ax.plot(x, tot, label="合計", **S_MURA)
    ax.plot(x, kei, label="軽度（要支援1〜要介護2）",
            color=K["m"], ls="--", marker="s", ms=5.5, lw=1.8, zorder=4)
    ax.plot(x, ni2, label="要介護2以上（国の成果指標）",
            color=K["d"], ls="-.", marker="D", ms=5, lw=1.6, zorder=3)
    ax.plot(x, juu, label="重度（要介護3以上）", **S_ZEN)
    label_last(ax, x, tot, f"{tot[-1]:.1f}%")
    label_last(ax, x, kei, f"{kei[-1]:.1f}%", color=K["m"])
    label_last(ax, x, ni2, f"{ni2[-1]:.1f}%", dy=7, color=K["d"])
    label_last(ax, x, juu, f"{juu[-1]:.1f}%", dy=-9, color="#808080")
    ax.set_xlim(-0.4, 6.9)
    ax.set_ylim(4, 26)
    ax.set_xticks(x); ax.set_xticklabels(lab, fontsize=8.5)
    # 目盛りは5刻みに固定する。matplotlib に任せると 17.5 や 22.5 が置かれ、
    # 整数に丸めた表示（18%・22%）が実際の目盛りと食い違う
    ax.set_yticks([5, 10, 15, 20, 25])
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{v:.0f}%"))
    style_ax(ax, ylab="性・年齢調整済み認定率")
    ax.legend(loc="upper left", fontsize=8.5, ncols=2)
    ax.set_title("性・年齢調整済み認定率の推移（要介護度の区分別）", loc="left", pad=10)
    save(fig, "fig2-21_調整済み認定率.png")

# ══════════ 図2-6 要介護度別認定者数の推移（積上げ）══════════
def fig_2_6():
    lab = ["令和2年\n3月末", "令和3年\n3月末", "令和4年\n3月末", "令和5年\n3月末",
           "令和6年\n3月末", "令和7年\n3月末", "令和8年\n3月末"]
    x = list(range(len(lab)))
    # 数値は正本から。色とハッチングは体裁なのでここに持つ
    STYLE = [("要支援1", "#1A1A1A", "///"), ("要支援2", "#404040", "\\\\\\"),
             ("要介護1", "#666666", "..."), ("要介護2", "#8C8C8C", "xxx"),
             ("要介護3", "#A6A6A6", "---"), ("要介護4", "#C4C4C4", "|||"),
             ("要介護5", "#E0E0E0", None)]
    series = [(nm, V("fig2-6_要介護度別認定者数", nm), col, hc)
              for nm, col, hc in STYLE]
    fig, ax = plt.subplots(figsize=(7.4, 4.2))
    bottom = [0]*len(x)
    for name, vals, col, hatch in series:
        ax.bar(x, vals, bottom=bottom, width=0.62, label=name, color=col,
               edgecolor="white", linewidth=1.6, hatch=hatch)
        bottom = [b+v for b, v in zip(bottom, vals)]
    for xi, ti in zip(x, bottom):
        ax.annotate(f"{ti}人", (xi, ti), xytext=(0, 5), textcoords="offset points",
                    ha="center", fontsize=9, color=K["d"], fontweight="bold")
    ax.set_xticks(x); ax.set_xticklabels(lab, fontsize=8.5)
    ax.set_ylim(0, 250); ax.set_xlim(-0.6, 6.6)
    style_ax(ax, ylab="認定者数（人）")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.14), ncols=7, fontsize=8.5,
              handlelength=1.4, columnspacing=1.0)
    ax.set_title("要介護度別 認定者数の推移", loc="left", pad=10)
    save(fig, "fig2-6_要介護度別認定者数.png")

# ══════════ 図2-7 要支援1の推移【重点】══════════
def fig_2_7():
    lab = ["令和2年", "令和3年", "令和4年", "令和5年", "令和6年", "令和7年", "令和8年"]
    x = list(range(len(lab)))
    ninsu = V("fig2-7_要支援1", "要支援1")
    fig, ax = plt.subplots(figsize=(7.2, 3.6))
    bars = ax.bar(x, ninsu, width=0.58, color=K["l"], edgecolor=K["d"], linewidth=1.0)
    bars[-1].set_color(K["d"]); bars[-1].set_hatch("///")
    bars[0].set_color(K["xl"])
    for b, v in zip(bars, ninsu):
        ax.annotate(f"{v}", (b.get_x()+b.get_width()/2, v), xytext=(0, 4),
                    textcoords="offset points", ha="center", fontsize=10,
                    fontweight="bold", color=K["d"])
    ax.annotate("", xy=(6, 55), xytext=(0, 21),
                arrowprops=dict(arrowstyle="->", color=K["d"], lw=1.6,
                                connectionstyle="arc3,rad=-0.18"))
    ax.annotate("6年間で3.0倍", (3, 52), ha="center", fontsize=11, fontweight="bold",
                color=K["d"])
    ax.set_xticks(x); ax.set_xticklabels([f"{l}\n3月末" for l in lab], fontsize=8.5)
    ax.set_ylim(0, 62)
    style_ax(ax, ylab="認定者数（人）")
    ax.set_title("要支援1の認定者数の推移【重点】", loc="left", pad=10)
    save(fig, "fig2-7_要支援1.png")

# ══════════ 図2-8 新規認定者に占める要支援1の割合 ══════════
def fig_2_8():
    lab = ["令和元年度", "令和2年度", "令和3年度", "令和4年度", "令和5年度", "令和6年度"]
    x = list(range(len(lab)))
    ratio = V("fig2-8_新規認定要支援1", "要支援1の割合")
    fig, ax = plt.subplots(figsize=(7.2, 3.4))
    ax.plot(x, ratio, color=K["d"], ls="-", marker="o", ms=7, lw=2.2)
    ax.fill_between(x, 0, ratio, color=K["xl"], alpha=0.6, zorder=0)
    for xi, v in zip(x, ratio):
        ax.annotate(f"{v}%", (xi, v), xytext=(0, 8), textcoords="offset points",
                    ha="center", fontsize=9.5, fontweight="bold", color=K["d"])
    ax.axhline(40, color=K["m"], lw=0.9, ls="--")
    ax.annotate("4割", (5.35, 40), fontsize=9, color=K["m"], va="center")
    ax.set_xticks(x); ax.set_xticklabels(lab, fontsize=9)
    ax.set_ylim(0, 50); ax.set_xlim(-0.35, 5.6)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{v:.0f}%"))
    style_ax(ax, ylab="要支援1の割合")
    ax.set_title("新規認定者に占める要支援1の割合", loc="left", pad=10)
    save(fig, "fig2-8_新規認定要支援1.png")

# ══════════ 図2-9 1人1月あたり費用額（村・県・全国）══════════
def fig_2_9():
    lab = ["平成29", "平成30", "令和元", "令和2", "令和3", "令和4", "令和5", "令和6", "令和7"]
    x = list(range(len(lab)))
    mura = V("fig2-9_1人あたり費用額", "北塩原村")
    ken = V("fig2-9_1人あたり費用額", "福島県")
    zen = V("fig2-9_1人あたり費用額", "全国")
    fig, ax = plt.subplots(figsize=(7.2, 3.8))
    ax.plot(x, mura, label="北塩原村", **S_MURA)
    ax.plot(x, ken,  label="福島県",   **S_KEN)
    ax.plot(x, zen,  label="全国",     **S_ZEN)
    ax.axvline(2, color=K["m"], lw=0.9)
    ax.annotate("令和元年度に\n県・全国を下回る", (2, 28300), ha="center", va="top",
                fontsize=8.5, color=K["d"], fontweight="bold")
    label_last(ax, x, mura, "24,696円")
    label_last(ax, x, ken,  "26,848円", dy=-9, color=K["m"])
    label_last(ax, x, zen,  "27,815円", dy=8, color="#808080")
    ax.set_xlim(-0.4, 9.6); ax.set_ylim(19000, 29500)
    ax.set_xticks(x); ax.set_xticklabels([f"{l}\n年度" for l in lab], fontsize=8.5)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{int(v):,}"))
    style_ax(ax, ylab="1人1月あたり費用額（円）")
    ax.legend(loc="lower left", fontsize=9, ncols=3)
    ax.set_title("第1号被保険者1人1月あたり費用額（村・福島県・全国）", loc="left", pad=10)
    save(fig, "fig2-9_1人あたり費用額.png")

# ══════════ 図2-10 費用額の内訳推移 ══════════
def fig_2_10():
    lab = ["平成29", "平成30", "令和元", "令和2", "令和3", "令和4", "令和5", "令和6", "令和7"]
    x = list(range(len(lab)))
    zaitaku = V("fig2-10_費用額内訳", "在宅サービス")
    kyoju = V("fig2-10_費用額内訳", "居住系サービス")
    shisetsu = V("fig2-10_費用額内訳", "施設サービス")
    fig, ax = plt.subplots(figsize=(7.4, 4.0))
    w = 0.26
    ax.bar([i-w for i in x], zaitaku, width=w, label="在宅サービス",
           color=K["d"], edgecolor="white", linewidth=1.2, hatch="///")
    ax.bar(x, kyoju, width=w, label="居住系サービス",
           color=K["m"], edgecolor="white", linewidth=1.2, hatch="...")
    ax.bar([i+w for i in x], shisetsu, width=w, label="施設サービス",
           color=K["l"], edgecolor=K["m"], linewidth=0.8)
    ax.annotate("+48.2%", (8+w-0.02, 91.0), xytext=(0, 6), textcoords="offset points",
                ha="center", fontsize=9, fontweight="bold", color=K["m"])
    ax.annotate("−10.0%", (8+w, 120.2), xytext=(0, 6), textcoords="offset points",
                ha="center", fontsize=9, fontweight="bold", color=K["m"])
    ax.set_xticks(x); ax.set_xticklabels([f"{l}\n年度" for l in lab], fontsize=8.5)
    ax.set_ylim(0, 175)
    style_ax(ax, ylab="費用額（百万円）")
    ax.legend(loc="upper left", fontsize=9, ncols=3)
    ax.set_title("サービス区分別 費用額の推移", loc="left", pad=10)
    save(fig, "fig2-10_費用額内訳.png")

# ══════════ 図2-12 1人あたり定員の推移 ══════════
def fig_2_12():
    lab = ["H27", "H28", "H29", "H30", "R元", "R2", "R3", "R4", "R5", "R6", "R7"]
    x = list(range(len(lab)))
    kyoju = V("fig2-12_1人あたり定員", "居住系（認知症GH）")
    tusho = V("fig2-12_1人あたり定員", "通所系（通所介護）")
    fig, ax = plt.subplots(figsize=(7.2, 3.6))
    ax.plot(x, tusho, label="通所系（通所介護）", color=K["d"], ls="-", marker="o", ms=6, lw=2.2)
    ax.plot(x, kyoju, label="居住系（認知症GH）", color=K["m"], ls="--", marker="s", ms=5.5, lw=1.8)
    for xi, ys, c in [(4, tusho, K["d"]), (4, kyoju, K["m"])]:
        ax.plot([xi], [ys[xi]], marker="o", ms=11, mfc="none", mec=c, mew=1.6)
    ax.annotate("ピーク（令和元年度）", (4, 0.30), ha="center", fontsize=8.5,
                color=K["d"], fontweight="bold")
    label_last(ax, x, tusho, "0.234")
    label_last(ax, x, kyoju, "0.126", color=K["m"])
    ax.set_xticks(x); ax.set_xticklabels([f"{l}\n年度" for l in lab], fontsize=8.5)
    ax.set_xlim(-0.4, 11.4); ax.set_ylim(0.07, 0.32)
    style_ax(ax, ylab="認定者1人あたり定員")
    ax.legend(loc="lower left", fontsize=9)
    ax.set_title("要支援・要介護者1人あたり定員の推移（定員は平成30年度以降据置）",
                 loc="left", pad=10)
    save(fig, "fig2-12_1人あたり定員.png")

# ══════════ 図2-13 地域支援事業費の推移（内訳）══════════
def fig_2_13():
    lab = ["平成29", "平成30", "令和元", "令和2", "令和3", "令和4", "令和5"]
    x = list(range(len(lab)))
    ippan = V("fig2-13_地域支援事業費", "一般介護予防事業費")
    sogo = V("fig2-13_地域支援事業費", "介護予防・生活支援サービス事業費")
    hokat = V("fig2-13_地域支援事業費", "包括的支援事業費")
    sonota = V("fig2-13_地域支援事業費", "その他")
    fig, ax = plt.subplots(figsize=(7.4, 4.0))
    b = [0]*len(x)
    for name, vals, col, h in [("一般介護予防事業", ippan, K["d"], "///"),
                               ("介護予防・生活支援サービス事業", sogo, K["m"], "..."),
                               ("包括的支援事業・任意事業", hokat, K["l"], None),
                               ("その他", sonota, K["xl"], "xx")]:
        ax.bar(x, vals, bottom=b, width=0.6, label=name, color=col,
               edgecolor="white", linewidth=1.6, hatch=h)
        b = [p+v for p, v in zip(b, vals)]
    for xi, ti in zip(x, b):
        ax.annotate(f"{ti:.1f}", (xi, ti), xytext=(0, 5), textcoords="offset points",
                    ha="center", fontsize=9, fontweight="bold", color=K["d"])
    ax.set_xticks(x); ax.set_xticklabels([f"{l}\n年度" for l in lab], fontsize=8.5)
    ax.set_ylim(0, 50)
    style_ax(ax, ylab="事業費（百万円）")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.13), ncols=2, fontsize=8.5)
    ax.set_title("地域支援事業費の推移（内訳別）", loc="left", pad=10)
    save(fig, "fig2-13_地域支援事業費.png")

# ══════════ 図2-15 給付費・地域支援事業費・保険料収入（指数）══════════
def fig_2_15():
    lab = ["H26", "H27", "H28", "H29", "H30", "R元", "R2", "R3", "R4", "R5"]
    x = list(range(len(lab)))
    kyufu = V("fig2-15_財政指数", "保険給付費")
    chiiki = V("fig2-15_財政指数", "地域支援事業費")
    hoken = V("fig2-15_財政指数", "保険料収入")
    idx = lambda a: [v/a[0]*100 for v in a]
    fig, ax = plt.subplots(figsize=(7.2, 3.8))
    ax.axhline(100, color=K["l"], lw=1.0, ls="-", zorder=1)
    ax.plot(x, idx(chiiki), label="地域支援事業費", color=K["d"], ls="-", marker="o", ms=6, lw=2.2)
    ax.plot(x, idx(hoken),  label="保険料収入",     color=K["m"], ls="--", marker="s", ms=5.5, lw=1.8)
    ax.plot(x, idx(kyufu),  label="保険給付費",     color=K["l"], ls=":", marker="^", ms=6, lw=2.0)
    label_last(ax, x, idx(chiiki), "256")
    label_last(ax, x, idx(hoken),  "142", color=K["m"])
    label_last(ax, x, idx(kyufu),  "100", color="#808080")
    ax.set_xticks(x); ax.set_xticklabels([f"{l}\n年度" for l in lab], fontsize=8.5)
    ax.set_xlim(-0.35, 9.9); ax.set_ylim(80, 280)
    style_ax(ax, ylab="平成26年度＝100")
    ax.legend(loc="upper left", fontsize=9)
    ax.set_title("保険給付費・地域支援事業費・保険料収入の推移（平成26年度＝100）",
                 loc="left", pad=10)
    save(fig, "fig2-15_財政指数.png")

# ══════════ 図2-16 保険料基準額（村・県・全国）══════════
def fig_2_16():
    lab = ["第7期\n（H30〜R2）", "第8期\n（R3〜R5）", "第9期\n（R6〜R8）"]
    x = list(range(len(lab)))
    mura = V("fig2-16_保険料比較", "北塩原村")
    ken = V("fig2-16_保険料比較", "福島県平均")
    zen = V("fig2-16_保険料比較", "全国平均")
    fig, ax = plt.subplots(figsize=(7.0, 3.6))
    w = 0.26
    b1 = ax.bar([i-w for i in x], mura, width=w, label="北塩原村",
                color=K["d"], edgecolor="white", linewidth=1.2, hatch="///")
    b2 = ax.bar(x, ken, width=w, label="福島県平均",
                color=K["m"], edgecolor="white", linewidth=1.2, hatch="...")
    b3 = ax.bar([i+w for i in x], zen, width=w, label="全国平均",
                color=K["l"], edgecolor=K["m"], linewidth=0.8)
    for bars in (b1, b2, b3):
        for b in bars:
            ax.annotate(f"{int(b.get_height()):,}", (b.get_x()+b.get_width()/2, b.get_height()),
                        xytext=(0, 3), textcoords="offset points", ha="center", fontsize=9)
    ax.annotate("県平均\n+360円", (2-w, 6700), xytext=(-30, 26), textcoords="offset points",
                ha="center", fontsize=9.5, fontweight="bold", color=K["d"],
                arrowprops=dict(arrowstyle="->", color=K["d"], lw=1.2))
    ax.set_xticks(x); ax.set_xticklabels(lab, fontsize=9)
    ax.set_ylim(5000, 7300)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{int(v):,}"))
    style_ax(ax, ylab="保険料基準額（円／月）")
    ax.legend(loc="upper left", fontsize=9, ncols=3)
    ax.set_title("介護保険料基準額の推移（村・福島県・全国）", loc="left", pad=10)
    save(fig, "fig2-16_保険料比較.png")

# ══════════ 図2-17 保険料基準額と必要保険料額 ══════════
def fig_2_17():
    lab = ["H30", "R元", "R2", "R3", "R4", "R5", "R6", "R7"]
    x = list(range(len(lab)))
    kijun = V("fig2-17_保険料と必要額", "保険料基準額")
    hitsu = V("fig2-17_保険料と必要額", "必要保険料額")
    fig, ax = plt.subplots(figsize=(7.2, 3.8))
    ax.fill_between(x, kijun, hitsu, where=[k >= h for k, h in zip(kijun, hitsu)],
                    color=K["xl"], alpha=0.9, interpolate=True, label="余剰", zorder=1)
    ax.step(x, kijun, where="mid", color=K["d"], lw=2.4, label="保険料基準額", zorder=4)
    ax.plot(x, hitsu, color=K["m"], ls="--", marker="s", ms=6, lw=1.8,
            label="必要保険料額", zorder=5)
    for xi, k, h in zip(x, kijun, hitsu):
        d = k - h
        if d > 0:
            ypos = (k + h) / 2 if d >= 200 else k + 60
            ax.annotate(f"+{d}", (xi, ypos), ha="center", va="bottom" if d < 200 else "center",
                        fontsize=8.5, color=K["d"], fontweight="bold")
        elif d < 0:
            ax.annotate(f"{d}", (xi, h + 40), ha="center", va="bottom",
                        fontsize=8.5, color=K["m"])
    ax.set_xticks(x); ax.set_xticklabels([f"{l}\n年度" for l in lab], fontsize=8.5)
    ax.set_ylim(5200, 7100)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{int(v):,}"))
    style_ax(ax, ylab="金額（円／月）")
    ax.legend(loc="upper left", fontsize=9, ncols=3)
    ax.set_title("保険料基準額と必要保険料額の対比（第8期以降は余剰が継続）",
                 loc="left", pad=10)
    save(fig, "fig2-17_保険料と必要額.png")

# ══════════ 図2-18 交付金 指標群別得点 ══════════
def fig_2_18():
    """指標群別得点（令和8年度）と全国平均の対比。"""
    grp = ["推進\n目標Ⅰ", "推進\n目標Ⅱ", "推進\n目標Ⅲ", "推進\n目標Ⅳ",
           "支援\n目標Ⅰ", "支援\n目標Ⅱ", "支援\n目標Ⅲ", "支援\n目標Ⅳ"]
    mura = V("fig2-18_交付金指標群", "北塩原村")
    zenkoku = V("fig2-18_交付金指標群", "全国平均")
    x = list(range(len(grp)))
    fig, ax = plt.subplots(figsize=(7.4, 3.8))
    w = 0.36
    b1 = ax.bar([i-w/2 for i in x], mura, width=w, label="北塩原村",
                color=K["d"], edgecolor="white", linewidth=1.2, hatch="///")
    b2 = ax.bar([i+w/2 for i in x], zenkoku, width=w, label="全国平均",
                color=K["l"], edgecolor=K["m"], linewidth=0.8)
    for b in b1:
        ax.annotate(f"{int(b.get_height())}", (b.get_x()+b.get_width()/2, b.get_height()),
                    xytext=(0, 3), textcoords="offset points", ha="center",
                    fontsize=8.5, fontweight="bold")
    ax.annotate("支援 目標Ⅱは全国平均を19点下回る", (4.55, 87), ha="center", fontsize=9,
                fontweight="bold", color=K["d"])
    ax.annotate("", xy=(4.9, 34), xytext=(4.55, 84),
                arrowprops=dict(arrowstyle="->", color=K["d"], lw=1.2,
                                connectionstyle="arc3,rad=-0.15"))
    ax.set_xticks(x); ax.set_xticklabels(grp, fontsize=8.5)
    ax.set_ylim(0, 100)
    style_ax(ax, ylab="得点（各100点満点）")
    ax.legend(loc="upper right", fontsize=8.5, ncol=2, bbox_to_anchor=(1.0, 1.02))
    ax.set_title("交付金 目標別得点と全国平均の対比（令和8年度）", loc="left", pad=10)
    save(fig, "fig2-18_交付金指標群.png")


# ══════════ 図2-19 交付金 合計得点の推移 ══════════
def fig_2_19():
    lab = ["令和6年度", "令和7年度", "令和8年度"]
    x = list(range(len(lab)))
    mura = V("fig2-19_交付金得点推移", "北塩原村")
    zenkoku = V("fig2-19_交付金得点推移", "全国平均")
    ken = V("fig2-19_交付金得点推移", "福島県平均")
    fig, ax = plt.subplots(figsize=(6.6, 3.6))
    ax.plot(x, mura, label="北塩原村", **S_MURA)
    ax.plot(x, ken, label="福島県平均", **S_KEN)
    ax.plot(x, zenkoku, label="全国平均", **S_ZEN)
    for xi, v in zip(x, mura):
        ax.annotate(f"{v}点", (xi, v), xytext=(0, -18), textcoords="offset points",
                    ha="center", fontsize=9, fontweight="bold")
    # 令和6年度の村と全国平均の差を縦の矢印で示す
    ax.annotate("", xy=(0.06, 295), xytext=(0.06, 422.4),
                arrowprops=dict(arrowstyle="<->", color=K["d"], lw=1.1))
    ax.annotate("全国平均との差\n−127点", (0.12, 358), ha="left", va="center",
                fontsize=8.5, fontweight="bold", color=K["d"])
    ax.annotate("−8点", (2, 447), xytext=(10, -4), textcoords="offset points",
                ha="left", va="center", fontsize=8.5, fontweight="bold", color=K["d"])
    ax.set_xticks(x); ax.set_xticklabels(lab, fontsize=9)
    ax.set_xlim(-0.35, 2.75); ax.set_ylim(250, 520)
    style_ax(ax, ylab="合計得点（800点満点）")
    ax.legend(loc="lower right", fontsize=9)
    ax.set_title("交付金 合計得点の推移（村・福島県平均・全国平均）", loc="left", pad=10)
    save(fig, "fig2-19_交付金得点推移.png")


# ══════════ 図2-20 通いの場の状況 ══════════
def fig_2_20():
    """令和6年度 介護予防・日常生活支援総合事業実施状況調査による。"""
    _v = V("fig2-20_通いの場", "値")
    kasho, sanka = _v[0:2], _v[2:4]
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(7.2, 3.4))
    # 左：箇所数
    x = [0, 1]
    ax1.bar(x, kasho, width=0.5, color=[K["l"], K["d"]],
            edgecolor=K["m"], linewidth=0.9, hatch=["", "///"])
    for xi, v in zip(x, kasho):
        ax1.annotate(f"{v}か所", (xi, v), xytext=(0, 4), textcoords="offset points",
                     ha="center", fontsize=10, fontweight="bold")
    ax1.set_xticks(x); ax1.set_xticklabels(["全体", "週1回以上\n開催"], fontsize=9)
    ax1.set_ylim(0, 28)
    style_ax(ax1, ylab="箇所数")
    ax1.set_title("箇所数", loc="left", fontsize=10.5, pad=6)
    # 右：参加者数と参加率
    ax2.bar(x, sanka, width=0.5, color=[K["l"], K["d"]],
            edgecolor=K["m"], linewidth=0.9, hatch=["", "///"])
    for xi, v, r in zip(x, sanka, ["27.8%", "12.2%"]):
        ax2.annotate(f"{v}人\n（参加率{r}）", (xi, v), xytext=(0, 4),
                     textcoords="offset points", ha="center", fontsize=9.5, fontweight="bold")
    ax2.set_xticks(x); ax2.set_xticklabels(["全体", "週1回以上\n開催"], fontsize=9)
    ax2.set_ylim(0, 350)
    style_ax(ax2, ylab="参加者数（人）")
    ax2.set_title("参加者数", loc="left", fontsize=10.5, pad=6)
    fig.suptitle("住民主体の通いの場の状況（令和6年度）　※参加率の分母は65歳以上人口1,009人",
                 x=0.02, ha="left", fontsize=12, y=1.02)
    save(fig, "fig2-20_通いの場.png")


# ══════════ 図2-11 在宅・施設居住系の1人あたり給付月額 ══════════
def fig_2_11():
    lab = ["H29", "H30", "R元", "R2", "R3", "R4", "R5", "R6", "R7"]
    x = list(range(len(lab)))
    zaitaku = V("fig2-11_在宅施設別給付月額", "在宅サービス")
    shisetsu = V("fig2-11_在宅施設別給付月額", "施設・居住系サービス")
    fig, ax = plt.subplots(figsize=(7.2, 3.6))
    ax.plot(x, shisetsu, label="施設・居住系サービス", color=K["d"], ls="-",
            marker="o", ms=6, lw=2.2)
    ax.plot(x, zaitaku, label="在宅サービス", color=K["m"], ls="--",
            marker="s", ms=5.5, lw=1.8)
    ax.plot([5], [5597], marker="o", ms=11, mfc="none", mec=K["m"], mew=1.6)
    ax.annotate("最低 5,597円\n（令和4年度）", (5, 5597), xytext=(0, -36),
                textcoords="offset points", ha="center", fontsize=8.5, color=K["m"])
    label_last(ax, x, shisetsu, "15,651円")
    label_last(ax, x, zaitaku, "7,154円", color=K["m"])
    ax.set_xticks(x); ax.set_xticklabels([f"{l}\n年度" for l in lab], fontsize=8.5)
    ax.set_xlim(-0.4, 9.7); ax.set_ylim(3500, 18500)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{int(v):,}"))
    style_ax(ax, ylab="1人1月あたり給付月額（円）")
    ax.legend(loc="center left", fontsize=9)
    ax.set_title("在宅・施設居住系別 第1号被保険者1人1月あたり給付月額",
                 loc="left", pad=10)
    save(fig, "fig2-11_在宅施設別給付月額.png")

# ══════════ 図2-14 サービス区分別の変化（H29→R7）══════════
def fig_2_14():
    names = ["訪問系", "通所系", "短期入所", "福祉用具・住宅改修",
             "居宅介護支援", "居住系", "施設系"]
    h29 = V("fig2-14_サービス区分別変化", "平成29年度")
    r7 = V("fig2-14_サービス区分別変化", "令和7年度")
    y = list(range(len(names)))[::-1]
    fig, ax = plt.subplots(figsize=(7.4, 4.2))
    h = 0.36
    b1 = ax.barh([i + h/2 for i in y], h29, height=h, color=K["l"],
                 edgecolor=K["m"], linewidth=0.8, label="平成29年度")
    b2 = ax.barh([i - h/2 for i in y], r7, height=h, color=K["d"],
                 edgecolor=K["d"], linewidth=0.8, label="令和7年度", hatch="///")
    for bars in (b1, b2):
        for b in bars:
            ax.annotate(f"{int(b.get_width()):,}",
                        (b.get_width(), b.get_y()+b.get_height()/2),
                        xytext=(4, 0), textcoords="offset points", va="center", fontsize=8.5)
    for i, a, b in zip(y, h29, r7):
        g = (b/a - 1) * 100
        ax.annotate(f"{g:+.1f}%", (12300, i), fontsize=9.5, va="center",
                    fontweight="bold", color=K["d"] if abs(g) >= 40 else K["m"])
    ax.set_yticks(y); ax.set_yticklabels(names, fontsize=9.5)
    ax.set_xlim(0, 14200)
    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{int(v):,}"))
    style_ax(ax, xlab="1人1月あたり給付月額（円）")
    ax.grid(axis="x", visible=True); ax.grid(axis="y", visible=False)
    ax.legend(loc="upper right", fontsize=9, bbox_to_anchor=(1.0, 0.99))
    ax.set_title("サービス区分別 1人1月あたり給付月額の変化（平成29年度→令和7年度）",
                 loc="left", pad=10)
    save(fig, "fig2-14_サービス区分別変化.png")



# ══════════ 追加図（doc49 §3-1）══════════
# いずれも算定の結果を読んで描く。固定値を書き写さない。

# ── 受給率の内訳（在宅・居住系・施設別）／2-4 ──
def fig_jukyuritsu_uchiwake():
    import csv as _csv, io as _io, collections as _c
    rows = list(_csv.reader(_io.open(
        os.path.join(_P.DATA, "mieruka_tidy.csv"),
        encoding="utf-8")))
    d = _c.defaultdict(dict)
    for r in rows[1:]:
        if r[4].startswith("合計受給率") and r[6] == "北塩原村":
            d[r[4]][r[7]] = float(r[10])
    # 見える化の期の表記を年度の順に並べる
    ORDER = ["H26", "H27", "H28", "H29", "H30", "R元", "R2", "R3", "R4", "R5",
             "R6（R7/2月サービス提供分まで）", "R7（R8/2月サービス提供分まで）"]
    lab = ["H26", "H27", "H28", "H29", "H30", "R元", "R2", "R3", "R4", "R5", "R6", "R7"]
    zai = [d["合計受給率（在宅サービス）"][k] for k in ORDER]
    kyo = [d["合計受給率（居住系サービス）"][k] for k in ORDER]
    shi = [d["合計受給率（施設サービス）"][k] for k in ORDER]
    x = list(range(len(lab)))
    fig, ax = plt.subplots(figsize=(7.4, 4.0))
    b1 = ax.bar(x, zai, color=K["xl"], edgecolor=K["m"], linewidth=0.8,
                label="在宅サービス")
    b2 = ax.bar(x, kyo, bottom=zai, color=K["m"], edgecolor=K["d"], linewidth=0.8,
                label="居住系サービス", hatch="///")
    b3 = ax.bar(x, shi, bottom=[a + b for a, b in zip(zai, kyo)],
                color=K["d"], edgecolor=K["d"], linewidth=0.8, label="施設サービス")
    for xi, a, b, c in zip(x, zai, kyo, shi):
        ax.annotate(f"{a+b+c:.1f}", (xi, a + b + c), xytext=(0, 3),
                    textcoords="offset points", ha="center", fontsize=8.5,
                    fontweight="bold", color=K["d"])
    # 在宅が底だった時期（平成30年度〜令和2年度）に印をつける
    ax.annotate("在宅の底（H30〜R2）", xy=(5, zai[5]), xytext=(5, 3.4),
                ha="center", fontsize=8.5, color=K["d"], fontweight="bold",
                arrowprops=dict(arrowstyle="-", color=K["m"], lw=0.9,
                                shrinkA=2, shrinkB=2))
    ax.set_xticks(x); ax.set_xticklabels([f"{l}\n年度" for l in lab], fontsize=8.5)
    ax.set_ylim(0, 18.6)
    style_ax(ax, ylab="受給率（％）")
    ax.legend(loc="upper center", fontsize=9, ncols=3, bbox_to_anchor=(0.5, -0.16))
    ax.set_title("受給率の内訳（在宅・居住系・施設別）　施設が縮み在宅が戻している",
                 loc="left", pad=10)
    save(fig, "fig_受給率の内訳.png")


# ── 事業所数からみた供給体制の位置（偏差値）／2-5 ──
def fig_jigyosho_ichi():
    import csv as _csv, io as _io, collections as _c
    rows = list(_csv.reader(_io.open(
        os.path.join(_P.DATA, "mieruka_tidy.csv"),
        encoding="utf-8")))
    d = _c.defaultdict(dict)
    for r in rows[1:]:
        if r[2] == "δ1-b" and "事業所数（偏差値）" in r[4]:
            d[r[4]][r[8]] = float(r[10])
    yrs = ["2019", "2022", "2025"]
    lab = ["令和元年", "令和4年", "令和7年"]
    fuku = [d["人口10万人あたり居宅（福祉系）サービス事業所数（偏差値）"][y] for y in yrs]
    iryo = [d["人口10万人あたり居宅（医療系）サービス事業所数（偏差値）"][y] for y in yrs]
    kyotu = [d["人口10万人あたり居宅介護支援事業所数（偏差値）"][y] for y in yrs]
    x = list(range(len(lab)))
    fig, ax = plt.subplots(figsize=(7.2, 3.8))
    ax.axhline(50, color=K["m"], ls="-", lw=1.0, zorder=1)
    ax.annotate("全国平均（偏差値50）", (2.32, 50), fontsize=8.5, color=K["m"],
                va="center", ha="left")
    ax.plot(x, kyotu, label="居宅介護支援事業所", **S_MURA)
    ax.plot(x, fuku, label="居宅（福祉系）サービス事業所", **S_KEN)
    ax.plot(x, iryo, label="居宅（医療系）サービス事業所", **S_ZEN)
    label_last(ax, x, kyotu, f"{kyotu[-1]:.1f}")
    label_last(ax, x, fuku, f"{fuku[-1]:.1f}", color=K["m"])
    label_last(ax, x, iryo, f"{iryo[-1]:.1f}", color=K["l"])
    ax.annotate("実数は3時点とも0\n（村内に医療系の事業所なし）", (0.5, 31.5),
                fontsize=8.5, color=K["d"], fontweight="bold", ha="left")
    ax.set_xticks(x); ax.set_xticklabels(lab, fontsize=9.5)
    ax.set_xlim(-0.3, 2.9); ax.set_ylim(28, 74)
    style_ax(ax, ylab="全国の市町村の中での偏差値")
    ax.legend(loc="upper left", fontsize=9)
    ax.set_title("事業所数からみた供給体制の位置　医療系は全国平均を大きく下回る",
                 loc="left", pad=10)
    save(fig, "fig_事業所数の位置.png")


# ── 需要に対する供給（担い手の枠）／5-4(12) ──
def fig_juyo_kyokyu():
    import sys as _s, os as _o
    _s.path.insert(0, _o.path.dirname(_o.path.abspath(__file__)))
    import estimate_jinzai as _J
    rows = _J.run()
    lab = [r[0] for r in rows]
    waku = [r[5] for r in rows]         # 担い手の枠
    juudo = [r[4] for r in rows]        # 要介護1以上（先に充てる量）
    maware = [r[6] for r in rows]       # 要支援に回せる量
    hitsuyo = [r[3] for r in rows]      # 要支援の必要量
    juuritsu = [r[7] for r in rows]     # 充足率
    x = list(range(len(lab)))
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(7.6, 3.9),
                                   gridspec_kw={"width_ratios": [1, 1], "wspace": 0.32})

    # 左：担い手の枠の内訳（重度から充て、残りが軽度に回る）
    ax1.bar(x, juudo, color=K["d"], edgecolor=K["d"], linewidth=0.8,
            label="要介護1以上に充てる")
    ax1.bar(x, maware, bottom=juudo, color=K["xl"], edgecolor=K["m"], linewidth=0.8,
            label="要支援に回せる", hatch="///")
    for xi, j_, m, w in zip(x, juudo, maware, waku):
        ax1.annotate(f"{w}", (xi, w), xytext=(0, 3), textcoords="offset points",
                     ha="center", fontsize=8.5, fontweight="bold", color=K["d"])
        if m >= 20:
            ax1.annotate(f"{m}", (xi, j_ + m / 2), ha="center", va="center",
                         fontsize=8.5, color=K["d"])
    ax1.set_xticks(x); ax1.set_xticklabels(lab, fontsize=8.5, rotation=30, ha="right")
    ax1.set_ylim(0, 245)
    style_ax(ax1, ylab="人")
    ax1.legend(loc="upper center", fontsize=8, ncols=1,
               bbox_to_anchor=(0.5, -0.24))
    ax1.set_title("担い手の枠の内訳", loc="left", fontsize=10.5, pad=8)

    # 右：要支援の必要量と回せる量の対比（ここが充足率）
    h = 0.38
    ax2.bar([i - h/2 for i in x], hitsuyo, width=h, color=K["xl"],
            edgecolor=K["m"], linewidth=0.8, label="要支援の必要量")
    ax2.bar([i + h/2 for i in x], maware, width=h, color=K["d"],
            edgecolor=K["d"], linewidth=0.8, label="要支援に回せる量")
    for xi, hh, m, r in zip(x, hitsuyo, maware, juuritsu):
        ax2.annotate(f"{hh}", (xi - h/2, hh), xytext=(0, 3), textcoords="offset points",
                     ha="center", fontsize=8.5, color=K["m"])
        ax2.annotate(f"{m}", (xi + h/2, m), xytext=(0, 3), textcoords="offset points",
                     ha="center", fontsize=8.5, fontweight="bold", color=K["d"])
        # 充足率は専用の帯に置く（棒や凡例と重ねない）
        ax2.annotate(f"{r:.1f}%", (xi, 106), ha="center", va="center", fontsize=9,
                     fontweight="bold", color=K["d"] if r < 50 else K["m"])
    ax2.axhline(98, color=K["xl"], lw=0.8, zorder=1)
    ax2.set_xticks(x); ax2.set_xticklabels(lab, fontsize=8.5, rotation=30, ha="right")
    ax2.set_ylim(0, 116)
    style_ax(ax2, ylab="人")
    ax2.legend(loc="upper center", fontsize=8, ncols=2,
               bbox_to_anchor=(0.5, -0.24))
    ax2.set_title("要支援の必要量と回せる量（上段は充足率）", loc="left",
                  fontsize=10.5, pad=8)

    fig.suptitle("需要に対する供給（担い手の枠）　要支援に回せる量が細っていく",
                 x=0.012, ha="left", fontsize=12, y=1.04)
    save(fig, "fig_需要に対する供給.png")


# ── 算定上の保険料月額と条例上の基準額／5-7 ──
def fig_hokenryo_jorei():
    import csv as _csv, io as _io
    rows = list(_csv.reader(_io.open(
        os.path.join(_P.DATA, "第10期_保険料パターン.csv"),
        encoding="utf-8")))[1:]
    K9 = 6700                            # 第9期の条例基準額
    pats = [(r[0], r[1], r[2], int(r[4])) for r in rows]
    g1 = [p for p in pats if p[1] == "取崩なし"]
    g2 = [p for p in pats if p[1] != "取崩なし"]
    lo = min(p[3] for p in pats); hi = max(p[3] for p in pats)
    fig, ax = plt.subplots(figsize=(7.6, 4.3))
    ax.axvline(K9, color=K["d"], ls="-", lw=2.0, zorder=2)
    ys, ylab = [], []
    for i, (p, mk, cl, nm) in enumerate([
            (g2, "o", K["d"], "基金を24,000千円取り崩す"),
            (g1, "s", K["m"], "基金を取り崩さない")]):
        base = i * 7
        for j, (futan, _kikin, suijun, tsuki) in enumerate(sorted(p, key=lambda t: t[3])):
            y = base + j + 1
            ys.append(y); ylab.append(f"　{futan}・{suijun}")
            # 条例上の基準額の置き方は estimate_kikin を正本とする。
            # ここで四捨五入をじか書きしていたため、第9期の実例（算定上
            # 6,760.76円に対し条例6,700円＝百円未満切捨て）と食い違っていた。
            import estimate_kikin as _EK
            jorei = _EK.jorei(tsuki)
            ax.plot([tsuki], [y], marker=mk, ms=7, color=cl, zorder=5)
            # 文字は白地を敷く。第9期の縦線が文字を切らないようにするため
            ax.annotate(f"{tsuki:,} → 条例 {jorei:,}", (tsuki, y), xytext=(9, 0),
                        textcoords="offset points", va="center", ha="left",
                        fontsize=8.5, color=K["d"], zorder=6,
                        bbox=dict(boxstyle="square,pad=0.12", fc="white",
                                  ec="none"))
        ys.append(base); ylab.append(nm)      # 群の見出しを軸ラベルとして置く
    ax.set_yticks(ys); ax.set_yticklabels(ylab, fontsize=8.5)
    for t, l in zip(ax.get_yticklabels(), ylab):
        if not l.startswith("　"):
            t.set_fontweight("bold"); t.set_fontsize(9)
    ax.set_xlim(lo - 150, hi + 620); ax.set_ylim(-0.9, 13.4)
    ax.invert_yaxis()
    ax.annotate(f"第9期の条例基準額 {K9:,}円", (K9, -0.75), fontsize=9,
                color=K["d"], fontweight="bold", ha="center", va="bottom")
    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{int(v):,}"))
    style_ax(ax, xlab="算定上の保険料月額（円）／矢印の右は百円未満を切り捨てた条例上の基準額")
    ax.grid(axis="x", visible=True); ax.grid(axis="y", visible=False)
    ax.set_title(f"算定上の月額と条例上の基準額　12パターンの幅は{hi-lo:,}円",
                 loc="left", pad=24)
    save(fig, "fig_保険料と条例基準額.png")



# ── 給付適正化の取組別得点の推移（推進交付金 目標Ⅱ）／5-8 ──
#    交付金の評価が求める「ケアマネジメントの質の分析」の裏づけ。
#    数値は受領した交付金の明細（配点・本村・全国平均）から取組ごとに足して出す。
#    じか書きしない。明細を差し替えれば図も変わる。
def fig_tekiseika_tokuten():
    import csv as _csv, io as _io, collections as _c
    rows = list(_csv.DictReader(_io.open(
        os.path.join(_P.DATA, "交付金評価指標_3か年.csv"), encoding="utf-8-sig")))

    def _f(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return None

    grp = [("給付費適正化\n方策の策定", "1 / 給付費適正化方策の策定状況"),
           ("給付費適正化\n事業の取組", "2 / 給付費適正化事業の取組状況"),
           ("ケアプラン\n点検", "1 / ケアプラン点検の実施状況"),
           ("医療情報との\n突合", "2 / 医療情報との突合の実施状況")]
    yrs = ["令和6年度", "令和7年度", "令和8年度"]
    hon = _c.defaultdict(lambda: _c.defaultdict(float))
    zen = _c.defaultdict(lambda: _c.defaultdict(float))
    hai = _c.defaultdict(float)
    for r in rows:
        if r["交付金"] != "推進交付金" or r["目標"] != "目標Ⅱ" or "計" in r["指標"]:
            continue
        for nm, pre in grp:
            if not r["指標"].startswith(pre):
                continue
            if _f(r["北塩原村"]) is not None:
                hon[nm][r["年度"]] += _f(r["北塩原村"])
            if _f(r["全国平均"]) is not None:
                zen[nm][r["年度"]] += _f(r["全国平均"])
            if r["年度"] == "令和8年度" and _f(r["配点"]) is not None:
                hai[nm] += _f(r["配点"])

    lab = [nm for nm, _ in grp]
    x = list(range(len(lab)))
    w = 0.21
    fig, ax = plt.subplots(figsize=(7.4, 3.9))
    STY = [dict(color=K["xl"], edgecolor=K["m"], linewidth=0.8),
           dict(color=K["l"], edgecolor=K["m"], linewidth=0.8),
           dict(color=K["d"], edgecolor="white", linewidth=1.2, hatch="///")]
    for i, (per, st) in enumerate(zip(yrs, STY)):
        v = [hon[nm][per] for nm in lab]
        b = ax.bar([xi + (i - 1.5) * w for xi in x], v, width=w, label=per, **st)
        if per == "令和8年度":
            for bb in b:
                ax.annotate("%d" % bb.get_height(),
                            (bb.get_x() + bb.get_width() / 2, bb.get_height()),
                            xytext=(0, 3), textcoords="offset points", ha="center",
                            fontsize=8.5, fontweight="bold")
    v8 = [zen[nm]["令和8年度"] for nm in lab]
    ax.bar([xi + 1.5 * w for xi in x], v8, width=w, label="全国平均（令和8年度）",
           color="#FFFFFF", edgecolor=K["m"], linewidth=0.9)
    for xi, nm in zip(x, lab):
        ax.plot([xi - 2 * w, xi + 2 * w], [hai[nm]] * 2,
                color=K["m"], lw=0.9, ls=":")
    ax.annotate("点線は配点（満点）", (-0.42, 37.6), fontsize=8,
                color=K["m"], ha="left")
    # 値が0の年度は棒が見えないため、0であることを書き添える
    for i, per in enumerate(yrs):
        for xi, nm in zip(x, lab):
            if hon[nm][per] == 0:
                ax.annotate("0", (xi + (i - 1.5) * w, 0.3), ha="center",
                            va="bottom", fontsize=8, color=K["m"])
    ax.annotate("令和8年度に8点を取得\n（4段階のうち2段階）", (2.05, 19.5), ha="center",
                fontsize=8.5, color=K["d"], fontweight="bold")
    ax.annotate("", xy=(2.2, 9.3), xytext=(2.05, 18.0),
                arrowprops=dict(arrowstyle="->", color=K["d"], lw=1.2,
                                connectionstyle="arc3,rad=-0.15"))
    ax.annotate("令和8年度に20点→6点", (1.15, 23.0), ha="center",
                fontsize=8.5, color=K["d"], fontweight="bold")
    ax.annotate("", xy=(1.2, 7.2), xytext=(1.15, 21.6),
                arrowprops=dict(arrowstyle="->", color=K["d"], lw=1.2,
                                connectionstyle="arc3,rad=0.12"))
    ax.set_xticks(x); ax.set_xticklabels(lab, fontsize=8.5)
    ax.set_ylim(0, 44)
    style_ax(ax, ylab="得点（点）")
    ax.legend(loc="upper center", fontsize=8, ncol=4, bbox_to_anchor=(0.5, 1.03),
              frameon=False, columnspacing=1.0, handlelength=1.4)
    ax.set_title("給付適正化の取組別得点の推移（推進交付金 目標Ⅱ）", loc="left", pad=10)
    save(fig, "fig_給付適正化の得点.png")


# ── 医療と介護の連携に関する加算の算定率と全国の中での位置／2-5 ──
#    交付金の評価が求める「在宅医療・介護連携の状況」の裏づけ。
def fig_renkei_kasan():
    import csv as _csv, io as _io
    rows = list(_csv.reader(_io.open(
        os.path.join(_P.DATA, "mieruka_tidy.csv"), encoding="utf-8")))
    d = {}
    for r in rows[1:]:
        if r[2] == "δ1-c" and r[6] == "北塩原村":
            d[r[4]] = None if r[10] in ("-", "") else float(r[10])
    KASAN = [("協力医療機関連携加算\n（施設・居住系）",
              "協力医療機関連携加算算定率（施設・居住系）"),
             ("入院時情報連携加算・退院退所加算\n（居宅介護支援）",
              "入院時情報連携加算、退院・退所加算算定率（居宅介護支援）"),
             ("認知症（専門ケア）加算\n（通所系・多機能系・施設居住系）",
              "認知症（専門ケア）加算算定率（通所系、多機能系、施設・居住系）"),
             ("看取り介護加算・ターミナルケア加算\n（施設・居住系）",
              "看取り介護加算、ターミナルケア加算算定率（施設・居住系）")]
    lab = [nm for nm, _ in KASAN]
    hen = [d[k + "（偏差値）"] for _, k in KASAN]
    ritu = [d[k] for _, k in KASAN]
    y = list(range(len(lab)))
    fig, ax = plt.subplots(figsize=(7.4, 3.6))
    ax.axvline(50, color=K["m"], lw=1.0)
    cols = [K["d"] if h >= 50 else K["l"] for h in hen]
    ax.barh(y, hen, height=0.54, color=cols, edgecolor=K["m"], linewidth=0.8)
    for yi, h, r in zip(y, hen, ritu):
        ax.annotate("偏差値 %.1f　（算定率 %.2f%%）" % (h, r), (h, yi), xytext=(6, 0),
                    textcoords="offset points", va="center", fontsize=8.5,
                    fontweight="bold", color=K["d"])
    ax.set_yticks(y); ax.set_yticklabels(lab, fontsize=8.5)
    ax.set_xlim(0, 76); ax.set_ylim(-0.7, len(lab) - 0.3)
    ax.invert_yaxis()
    ax.annotate("全国平均（偏差値50）", (50, -0.66), fontsize=8.5, color=K["m"],
                ha="center", va="bottom")
    ax.annotate("緊急時訪問看護加算は、村内に訪問看護事業所がないため値がない",
                (0.8, 3.46), fontsize=8, color=K["m"], ha="left", va="top")
    style_ax(ax, xlab="全国の市町村の中での偏差値")
    ax.grid(axis="x", visible=True); ax.grid(axis="y", visible=False)
    ax.set_title("医療と介護の連携に関する加算の算定率と全国の中での位置（令和7年）",
                 loc="left", pad=20)
    save(fig, "fig_連携の加算.png")


if __name__ == "__main__":
    fig_2_5(); fig_2_6(); fig_2_7(); fig_2_8(); fig_2_21()
    fig_2_9(); fig_2_10(); fig_2_11(); fig_2_12()
    fig_2_13(); fig_2_14(); fig_2_15(); fig_2_16()
    fig_2_17(); fig_2_18(); fig_2_19(); fig_2_20()
    # 追加図（doc49 §3-1）
    fig_jukyuritsu_uchiwake(); fig_jigyosho_ichi()
    fig_juyo_kyokyu(); fig_hokenryo_jorei()
    # 追加図（doc61 §4）交付金の評価が求める分析の空白2件に対応
    fig_tekiseika_tokuten(); fig_renkei_kasan()
