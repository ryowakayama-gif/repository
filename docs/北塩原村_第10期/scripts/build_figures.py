# -*- coding: utf-8 -*-
"""第10期計画 第2章 図表の作成
   成果品はモノクロ印刷（仕様書）のため、色相ではなく
   明度差＋線種＋マーカー形状＋ハッチングの二重符号化で系列を識別する。
   出力: output/figures/*.png（300dpi）"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

OUT = "/home/user/repository/output/figures"
os.makedirs(OUT, exist_ok=True)

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
    y2010 = [415, 713, 1166, 345, 546]
    y2025 = [186, 451, 741, 457, 480]
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
    ax.set_title("図2-1　年齢階級別人口の比較（平成22年・令和7年）", loc="left", pad=10)
    save(fig, "fig2-1_人口構成比較.png")

# ══════════ 図2-2 高齢化率の推移（村・県・全国）══════════
def fig_2_2():
    x = [2010, 2015, 2020, 2025, 2030, 2035, 2040, 2045, 2050]
    mura = [28.0, 31.8, 37.1, 40.5, 43.9, 45.2, 46.3, 47.8, 49.2]
    ken  = [24.9, 28.3, 31.2, 34.2, 36.1, 37.7, 40.3, 42.5, 44.2]
    zen  = [22.8, 26.3, 28.0, 29.6, 30.8, 32.3, 34.8, 36.3, 37.1]
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
    ax.set_title("図2-2　高齢化率の推移と将来推計（村・福島県・全国）", loc="left", pad=10)
    save(fig, "fig2-2_高齢化率推移.png")

# ══════════ 図2-3 将来推計人口（年齢3区分）══════════
def fig_2_3():
    x = [2010, 2015, 2020, 2025, 2030, 2035, 2040, 2045, 2050]
    u15  = [415, 324, 258, 186, 156, 132, 119, 104, 86]
    prod = [1879, 1608, 1349, 1192, 1008, 888, 770, 651, 554]
    o65  = [891, 899, 948, 937, 911, 843, 765, 692, 619]
    fig, ax = plt.subplots(figsize=(7.2, 3.8))
    ax.stackplot(x, u15, prod, o65,
                 labels=["15歳未満", "生産年齢人口（15〜64歳）", "高齢者人口（65歳以上）"],
                 colors=[K["xl"], K["l"], K["d"]],
                 edgecolor="white", linewidth=2,
                 hatch=[None, "..", "///"])
    tot = [3185, 2831, 2556, 2315, 2075, 1863, 1654, 1447, 1259]  # 出典の総人口
    for xi, ti in zip(x, tot):
        ax.annotate(f"{ti:,}", (xi, ti), xytext=(0, 5), textcoords="offset points",
                    ha="center", fontsize=8.5, color=K["d"])
    ax.set_xlim(2010, 2050); ax.set_ylim(0, 3600)
    ax.set_xticks(x); ax.set_xticklabels([str(v) for v in x], fontsize=9)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{int(v):,}"))
    style_ax(ax, ylab="人口（人）")
    ax.legend(loc="upper right", fontsize=9)
    ax.set_title("図2-3　将来推計人口（年齢3区分）", loc="left", pad=10)
    ax.annotate("※ 令和2年は年齢不詳1人を含むため、年齢区分の合計と総人口が1人異なります。",
                (0, -0.16), xycoords="axes fraction", fontsize=8, color=K["m"])
    save(fig, "fig2-3_将来推計人口.png")

# ══════════ 図2-4 前期／後期高齢者の推移 ══════════
def fig_2_4():
    x = [2010, 2015, 2020, 2025, 2030, 2035, 2040, 2045, 2050]
    zen_ki = [345, 399, 507, 457, 360, 302, 246, 235, 221]
    kou_ki = [546, 500, 441, 480, 551, 541, 519, 457, 398]
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
    ax.set_title("図2-4　前期高齢者と後期高齢者の推移（75歳以上は令和12年がピーク）",
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
    mura = [17.4, 18.2, 18.4, 18.8, 19.9, 19.6, 21.1]
    ken  = [19.2, 19.3, 19.3, 19.2, 19.3, 19.5, 19.8]
    zen  = [18.4, 18.7, 18.9, 19.0, 19.4, 19.7, 20.2]
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
    ax.set_title("図2-5　要支援・要介護認定率の推移（村・福島県・全国）", loc="left", pad=10)
    save(fig, "fig2-5_認定率推移.png")

# ══════════ 図2-6 要介護度別認定者数の推移（積上げ）══════════
def fig_2_6():
    lab = ["令和2年\n3月末", "令和3年\n3月末", "令和4年\n3月末", "令和5年\n3月末",
           "令和6年\n3月末", "令和7年\n3月末", "令和8年\n3月末"]
    x = list(range(len(lab)))
    series = [
        ("要支援1", [17, 21, 29, 28, 38, 39, 51], "#1A1A1A", "///"),
        ("要支援2", [20, 26, 33, 34, 31, 27, 29], "#404040", "\\\\\\"),
        ("要介護1", [43, 45, 34, 38, 51, 47, 49], "#666666", "..."),
        ("要介護2", [32, 31, 32, 30, 27, 29, 26], "#8C8C8C", "xxx"),
        ("要介護3", [25, 28, 27, 25, 20, 20, 25], "#A6A6A6", "---"),
        ("要介護4", [19, 22, 21, 28, 20, 19, 19], "#C4C4C4", "|||"),
        ("要介護5", [18, 13, 11, 8, 13, 17, 15], "#E0E0E0", None),
    ]
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
    ax.set_title("図2-6　要介護度別 認定者数の推移", loc="left", pad=10)
    save(fig, "fig2-6_要介護度別認定者数.png")

# ══════════ 図2-7 要支援1の推移【重点】══════════
def fig_2_7():
    lab = ["令和2年", "令和3年", "令和4年", "令和5年", "令和6年", "令和7年", "令和8年"]
    x = list(range(len(lab)))
    ninsu = [17, 21, 29, 28, 38, 39, 51]
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
    ax.set_title("図2-7　要支援1の認定者数の推移【重点】", loc="left", pad=10)
    save(fig, "fig2-7_要支援1.png")

# ══════════ 図2-8 新規認定者に占める要支援1の割合 ══════════
def fig_2_8():
    lab = ["令和元年度", "令和2年度", "令和3年度", "令和4年度", "令和5年度", "令和6年度"]
    x = list(range(len(lab)))
    ratio = [24.1, 21.7, 28.6, 35.3, 41.5, 40.6]
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
    ax.set_title("図2-8　新規認定者に占める要支援1の割合", loc="left", pad=10)
    save(fig, "fig2-8_新規認定要支援1.png")

# ══════════ 図2-9 1人1月あたり費用額（村・県・全国）══════════
def fig_2_9():
    lab = ["平成29", "平成30", "令和元", "令和2", "令和3", "令和4", "令和5", "令和6", "令和7"]
    x = list(range(len(lab)))
    mura = [25039, 25606, 23624, 22072, 21041, 21563, 23008, 22751, 24696]
    ken  = [24056, 24449, 24819, 25213, 25426, 25477, 25871, 26518, 26848]
    zen  = [23238, 23499, 24106, 24567, 25137, 25471, 26229, 27147, 27815]
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
    ax.set_title("図2-9　第1号被保険者1人1月あたり費用額（村・福島県・全国）", loc="left", pad=10)
    save(fig, "fig2-9_1人あたり費用額.png")

# ══════════ 図2-10 費用額の内訳推移 ══════════
def fig_2_10():
    lab = ["平成29", "平成30", "令和元", "令和2", "令和3", "令和4", "令和5", "令和6", "令和7"]
    x = list(range(len(lab)))
    zaitaku = [103.5, 94.7, 81.5, 85.0, 87.3, 74.4, 75.5, 84.5, 95.4]
    kyoju   = [61.4, 63.1, 69.9, 74.0, 64.6, 69.5, 76.4, 81.4, 91.0]
    shisetsu= [133.6, 150.3, 138.0, 124.2, 118.0, 130.0, 134.8, 120.8, 120.2]
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
    ax.set_title("図2-10　サービス区分別 費用額の推移", loc="left", pad=10)
    save(fig, "fig2-10_費用額内訳.png")

# ══════════ 図2-12 1人あたり定員の推移 ══════════
def fig_2_12():
    lab = ["H27", "H28", "H29", "H30", "R元", "R2", "R3", "R4", "R5", "R6", "R7"]
    x = list(range(len(lab)))
    kyoju = [0.092, 0.134, 0.140, 0.148, 0.155, 0.145, 0.144, 0.141, 0.135, 0.136, 0.126]
    tusho = [0.204, 0.223, 0.233, 0.273, 0.287, 0.269, 0.267, 0.262, 0.250, 0.253, 0.234]
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
    ax.set_title("図2-12　要支援・要介護者1人あたり定員の推移（定員は平成30年度以降据置）",
                 loc="left", pad=10)
    save(fig, "fig2-12_1人あたり定員.png")

# ══════════ 図2-13 地域支援事業費の推移（内訳）══════════
def fig_2_13():
    lab = ["平成29", "平成30", "令和元", "令和2", "令和3", "令和4", "令和5"]
    x = list(range(len(lab)))
    ippan = [1.71, 1.61, 1.77, 1.34, 1.09, 1.78, 1.83]
    sogo  = [8.12, 8.48, 9.48, 10.73, 10.80, 14.18, 14.64]
    hokat = [13.50, 13.46, 13.89, 13.66, 23.33, 23.79, 23.42]
    sonota= [5.40, 9.92, 10.54, 11.73, 0, 0, 0]
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
    ax.set_title("図2-13　地域支援事業費の推移（内訳別）", loc="left", pad=10)
    save(fig, "fig2-13_地域支援事業費.png")

# ══════════ 図2-15 給付費・地域支援事業費・保険料収入（指数）══════════
def fig_2_15():
    lab = ["H26", "H27", "H28", "H29", "H30", "R元", "R2", "R3", "R4", "R5"]
    x = list(range(len(lab)))
    kyufu = [279774, 284396, 281570, 290913, 302371, 286563, 279049, 266191, 271863, 280366]
    chiiki= [15580, 16246, 22780, 28727, 33471, 35671, 37451, 35216, 39754, 39886]
    hoken = [48998, 55142, 55416, 56268, 65179, 64863, 64824, 69796, 69558, 69539]
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
    ax.set_title("図2-15　保険給付費・地域支援事業費・保険料収入の推移（平成26年度＝100）",
                 loc="left", pad=10)
    save(fig, "fig2-15_財政指数.png")

# ══════════ 図2-16 保険料基準額（村・県・全国）══════════
def fig_2_16():
    lab = ["第7期\n（H30〜R2）", "第8期\n（R3〜R5）", "第9期\n（R6〜R8）"]
    x = list(range(len(lab)))
    mura = [5900, 6300, 6700]; ken = [6061, 6108, 6340]; zen = [5784, 6014, 6225]
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
    ax.set_title("図2-16　介護保険料基準額の推移（村・福島県・全国）", loc="left", pad=10)
    save(fig, "fig2-16_保険料比較.png")

# ══════════ 図2-17 保険料基準額と必要保険料額 ══════════
def fig_2_17():
    lab = ["H30", "R元", "R2", "R3", "R4", "R5", "R6", "R7"]
    x = list(range(len(lab)))
    kijun = [5900, 5900, 5900, 6300, 6300, 6300, 6700, 6700]
    hitsu = [6012, 6011, 5830, 5451, 6007, 6335, 5991, 6501]
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
    ax.set_title("図2-17　保険料基準額と必要保険料額の対比（第8期以降は余剰が継続）",
                 loc="left", pad=10)
    save(fig, "fig2-17_保険料と必要額.png")

# ══════════ 図2-18 交付金 指標群別得点 ══════════
def fig_2_18():
    grp = ["取組・体制\n指標群", "活動指標群\n（アウトプット）", "成果指標群\n（アウトカム）"]
    x = list(range(len(grp)))
    suishin = [50, 9, 65]; shien = [77, 29, 65]
    fig, ax = plt.subplots(figsize=(7.0, 3.6))
    w = 0.34
    b1 = ax.bar([i-w/2 for i in x], suishin, width=w, label="保険者機能強化推進交付金",
                color=K["d"], edgecolor="white", linewidth=1.2, hatch="///")
    b2 = ax.bar([i+w/2 for i in x], shien, width=w, label="介護保険保険者努力支援交付金",
                color=K["l"], edgecolor=K["m"], linewidth=0.8)
    for bars in (b1, b2):
        for b in bars:
            ax.annotate(f"{int(b.get_height())}点", (b.get_x()+b.get_width()/2, b.get_height()),
                        xytext=(0, 3), textcoords="offset points", ha="center",
                        fontsize=9.5, fontweight="bold")
    ax.annotate("活動量が\n突出して低い", (1, 45), ha="center", fontsize=10,
                fontweight="bold", color=K["d"])
    ax.set_xticks(x); ax.set_xticklabels(grp, fontsize=9)
    ax.set_ylim(0, 92)
    style_ax(ax, ylab="得点（点）")
    ax.legend(loc="upper right", fontsize=8.5)
    ax.set_title("図2-18　保険者機能強化推進交付金等 指標群別得点（令和5年度）", loc="left", pad=10)
    save(fig, "fig2-18_交付金指標群.png")


# ══════════ 図2-11 在宅・施設居住系の1人あたり給付月額 ══════════
def fig_2_11():
    lab = ["H29", "H30", "R元", "R2", "R3", "R4", "R5", "R6", "R7"]
    x = list(range(len(lab)))
    zaitaku = [8157, 7334, 6198, 6324, 6527, 5597, 5737, 6329, 7154]
    shisetsu = [15127, 16286, 15561, 14531, 13422, 14729, 15747, 14915, 15651]
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
    ax.set_title("図2-11　在宅・施設居住系別 第1号被保険者1人1月あたり給付月額",
                 loc="left", pad=10)
    save(fig, "fig2-11_在宅施設別給付月額.png")

# ══════════ 図2-14 サービス区分別の変化（H29→R7）══════════
def fig_2_14():
    names = ["訪問系", "通所系", "短期入所", "福祉用具・住宅改修",
             "居宅介護支援", "居住系", "施設系"]
    h29 = [1481, 4166, 756, 628, 1127, 4740, 10387]
    r7  = [728, 3851, 593, 879, 1103, 6748, 8903]
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
    ax.set_title("図2-14　サービス区分別 1人1月あたり給付月額の変化（平成29年度→令和7年度）",
                 loc="left", pad=10)
    save(fig, "fig2-14_サービス区分別変化.png")

if __name__ == "__main__":
    fig_2_5(); fig_2_6(); fig_2_7(); fig_2_8()
    fig_2_9(); fig_2_10(); fig_2_11(); fig_2_12()
    fig_2_13(); fig_2_14(); fig_2_15(); fig_2_16()
    fig_2_17(); fig_2_18()
