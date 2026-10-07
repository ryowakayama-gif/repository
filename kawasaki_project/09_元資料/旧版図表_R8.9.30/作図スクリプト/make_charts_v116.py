# -*- coding: utf-8 -*-
"""素案 Ver.1.16 用の図の差替え（図2-4・図9-1）。

  図2-4　所得段階別第1号被保険者の構成　令和3年度末9段階 → 令和7年度末13段階
  図9-1　介護保険料月額基準額の推移　所得段階別被保険者数の反映後の値へ
"""
import os, sys
import matplotlib
matplotlib.use("Agg")
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
from matplotlib.patches import Patch

sys.path.insert(0, "07_ソーススクリプト")
from make_charts_v114 import fig91  # noqa: E402

FONT = next(f for f in ("/usr/share/fonts/opentype/noto/NotoSansCJK-Medium.ttc",
                        "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf",
                        "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf")
            if os.path.exists(f))
JP = fm.FontProperties(fname=FONT, size=10)
JP_S = fm.FontProperties(fname=FONT, size=8)
JP_T = fm.FontProperties(fname=FONT, size=14)
plt.rcParams["axes.unicode_minus"] = False
NAVY, BLUE, LBLUE = "#1F3864", "#2F5597", "#DAE3F3"
ORANGE, LORANGE, GREEN, LGREEN = "#ED7D31", "#FCE4D6", "#548235", "#E2EFDA"

R7 = [383, 298, 305, 343, 679, 410, 427, 232, 85, 26, 17, 7, 28]
TOT = sum(R7)


def fig24():
    fig, ax = plt.subplots(figsize=(11, 5.5), dpi=150)
    lab = ["第1\n段階", "第2\n段階", "第3\n段階", "第4\n段階", "第5段階\n(基準)",
           "第6\n段階", "第7\n段階", "第8\n段階", "第9\n段階", "第10\n段階",
           "第11\n段階", "第12\n段階", "第13\n段階"]
    fc = [LGREEN] * 3 + [LBLUE] + [ORANGE] + [LORANGE] * 8
    ec = [GREEN] * 3 + [BLUE] + [NAVY] + [ORANGE] * 8
    bars = ax.bar(lab, R7, color=fc, edgecolor=ec, linewidth=2)
    for b, n in zip(bars, R7):
        ax.text(b.get_x() + b.get_width() / 2, n + 12, f"{n}", ha="center",
                va="bottom", fontproperties=JP_S, fontsize=10, fontweight="bold")
    ax.set_ylabel("人数（人）", fontproperties=JP)
    ax.set_xticks(range(len(lab)))
    ax.set_xticklabels(lab, fontproperties=JP_S, fontsize=9)
    ax.set_title(f"図2-4　所得段階別第1号被保険者の構成（令和7年度末・計{TOT:,}人）",
                 fontproperties=JP_T, color=NAVY, pad=15)
    ax.set_ylim(0, 800)
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    hikazei = sum(R7[:3]) / TOT * 100
    honnin = R7[3] / TOT * 100
    kijun = R7[4] / TOT * 100
    kazei = sum(R7[5:]) / TOT * 100
    ax.legend(handles=[
        Patch(facecolor=LGREEN, edgecolor=GREEN, label=f"非課税層（{hikazei:.1f}%）"),
        Patch(facecolor=LBLUE, edgecolor=BLUE, label=f"本人非課税（{honnin:.1f}%）"),
        Patch(facecolor=ORANGE, edgecolor=NAVY, label=f"第5段階（基準・{kijun:.1f}%）"),
        Patch(facecolor=LORANGE, edgecolor=ORANGE, label=f"本人課税層（{kazei:.1f}%）"),
    ], prop=JP_S, loc="upper right", framealpha=0.95)
    plt.tight_layout()
    p = "08_図表/chart_income_v116.png"
    plt.savefig(p, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()
    return p


if __name__ == "__main__":
    print(fig24())
    import shutil
    print(fig91(6822, 6170, 5518))
    shutil.move("08_図表/chart_premium_v114.png", "08_図表/chart_premium_v116.png")
