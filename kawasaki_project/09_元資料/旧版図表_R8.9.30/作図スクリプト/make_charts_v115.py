# -*- coding: utf-8 -*-
"""素案 Ver.1.15 用の図の差替え（図2-1・図2-2）を生成する。

  図2-1　第1号被保険者の年齢階級別推移　令和3年度末→令和7年6月 → 令和3年度末→令和7年度末
  図2-2　サービス受給者の区分別構成　令和7年6月時点 → 令和7年度の月平均
"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import Circle

FONT = next(f for f in ("/usr/share/fonts/opentype/noto/NotoSansCJK-Medium.ttc",
                        "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf",
                        "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf")
            if os.path.exists(f))
JP = fm.FontProperties(fname=FONT, size=10)
JP_S = fm.FontProperties(fname=FONT, size=8)
JP_L = fm.FontProperties(fname=FONT, size=12)
JP_T = fm.FontProperties(fname=FONT, size=14)
plt.rcParams["axes.unicode_minus"] = False
NAVY, BLUE, ORANGE, GREEN = "#1F3864", "#2F5597", "#ED7D31", "#548235"
RED, GRAY = "#C00000", "#808080"
OUT = "08_図表"


def fig21():
    fig, ax = plt.subplots(figsize=(8, 5), dpi=150)
    labels = ["65〜75歳未満\n（前期高齢者）", "75〜85歳未満", "85歳以上"]
    r3 = [1745, 905, 605]
    r7 = [1498, 1142, 600]
    x = np.arange(len(labels)); w = 0.35
    b1 = ax.bar(x - w / 2, r3, w, label="令和3年度末", color=BLUE,
                edgecolor="white", linewidth=1.5)
    b2 = ax.bar(x + w / 2, r7, w, label="令和7年度末", color=ORANGE,
                edgecolor="white", linewidth=1.5)
    for bars in (b1, b2):
        for b in bars:
            h = b.get_height()
            ax.text(b.get_x() + b.get_width() / 2, h + 30, f"{h:,}人",
                    ha="center", va="bottom", fontproperties=JP_S, fontsize=9)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontproperties=JP)
    ax.set_ylabel("人数（人）", fontproperties=JP)
    ax.set_title("図2-1　第1号被保険者の年齢階級別推移（令和3年度末→令和7年度末）",
                 fontproperties=JP_T, color=NAVY, pad=15)
    ax.set_ylim(0, 2200)
    ax.legend(prop=JP, loc="upper right", framealpha=0.95)
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for xi, y0, y1, txt, col in ((0, 1600, 2050, "△14.2%", RED),
                                 (1, 1180, 1600, "+26.2%", GREEN),
                                 (2, 640, 950, "△0.8%", GRAY)):
        ax.annotate(txt, xy=(xi, y0), xytext=(xi, y1),
                    arrowprops=dict(arrowstyle="->", color=col, lw=2),
                    ha="center", fontproperties=JP, fontsize=11, color=col,
                    fontweight="bold")
    plt.tight_layout()
    p = f"{OUT}/chart_population_v115.png"
    plt.savefig(p, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()
    return p


def fig22(kyo, chi, shi):
    tot = kyo + chi + shi
    fig, ax = plt.subplots(figsize=(7, 6), dpi=150)
    sizes = [kyo, chi, shi]
    labels = [f"居宅サービス\n{kyo:,.1f}人 ({kyo / tot * 100:.1f}%)",
              f"地域密着型\n{chi:,.1f}人 ({chi / tot * 100:.1f}%)",
              f"施設サービス\n{shi:,.1f}人 ({shi / tot * 100:.1f}%)"]
    ax.pie(sizes, labels=labels, colors=[BLUE, GREEN, ORANGE], startangle=90,
           explode=(0, 0.04, 0.04),
           textprops={"fontproperties": JP, "fontsize": 11},
           wedgeprops={"edgecolor": "white", "linewidth": 2})
    ax.add_patch(Circle((0, 0), 0.32, facecolor="white", edgecolor=NAVY,
                        linewidth=1.5, zorder=10))
    ax.text(0, 0, f"計{tot:,.1f}人\n（重複あり）", ha="center", va="center",
            fontproperties=JP_L, fontsize=13, color=NAVY, fontweight="bold",
            zorder=11)
    ax.set_title("図2-2　サービス受給者の区分別構成（令和7年度の月平均）",
                 fontproperties=JP_T, color=NAVY, pad=15)
    plt.tight_layout()
    p = f"{OUT}/chart_recipient_v115.png"
    plt.savefig(p, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()
    return p


if __name__ == "__main__":
    print(fig21())
    print(fig22(261.8, 53.9, 134.3))
