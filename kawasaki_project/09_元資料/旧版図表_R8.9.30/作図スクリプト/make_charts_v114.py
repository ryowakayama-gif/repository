# -*- coding: utf-8 -*-
"""素案 Ver.1.14 用の図の差替え（図2-3・図9-1）を生成する。

  図2-3　サービス区分別の年間給付費　令和3年度実績 → 令和7年度実績
  図9-1　介護保険料月額基準額の推移　仮試算A/B/C を確定算定値に差替え
"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import numpy as np

FONT = next(f for f in ("/usr/share/fonts/opentype/noto/NotoSansCJK-Medium.ttc",
                        "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf",
                        "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf")
            if os.path.exists(f))
JP = fm.FontProperties(fname=FONT, size=10)
JP_S = fm.FontProperties(fname=FONT, size=8)
JP_T = fm.FontProperties(fname=FONT, size=14)
plt.rcParams["axes.unicode_minus"] = False

NAVY, BLUE, LBLUE = "#1F3864", "#2F5597", "#DAE3F3"
ORANGE, LORANGE = "#ED7D31", "#FCE4D6"
GREEN, LGREEN, GRAY = "#548235", "#E2EFDA", "#808080"
OUT = "08_図表"


# ─────────────────────────── 図2-3
def fig23():
    fig, ax = plt.subplots(figsize=(9, 5.5), dpi=150)
    services = ["居宅サービス\n（在宅・居宅介護支援を含む）",
                "地域密着型\nサービス", "施設サービス\n（特養・老健）"]
    v = [3.2268, 1.8219, 4.9335]                      # 億円
    bars = ax.barh(services, v, color=[BLUE, GREEN, ORANGE],
                   edgecolor="white", linewidth=2)
    for b, x in zip(bars, v):
        ax.text(x + 0.08, b.get_y() + b.get_height() / 2,
                f"{x:.2f}億円  ({x / sum(v) * 100:.1f}%)", va="center",
                fontproperties=JP, fontsize=11, fontweight="bold")
    ax.set_xlabel("給付費（億円）", fontproperties=JP)
    ax.set_yticks(range(len(services)))
    ax.set_yticklabels(services, fontproperties=JP)
    ax.set_title("図2-3　サービス区分別の年間給付費（令和7年度実績・総額9.98億円）",
                 fontproperties=JP_T, color=NAVY, pad=15)
    ax.set_xlim(0, 6.0)
    ax.grid(axis="x", linestyle=":", alpha=0.4)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.invert_yaxis()
    plt.tight_layout()
    p = f"{OUT}/chart_benefit_v114.png"
    plt.savefig(p, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()
    return p


# ─────────────────────────── 図9-1
def fig91(A, B, C):
    K8, K9 = 6_380, 6_500
    fig, ax = plt.subplots(figsize=(9, 5.5), dpi=150)
    periods = ["第8期\n(R3〜R5)", "第9期\n(R6〜R8)",
               "第10期\n(R9〜R11)\n試算A\n基金取崩なし",
               "第10期\n(R9〜R11)\n試算B\n基金50%取崩",
               "第10期\n(R9〜R11)\n試算C\n基金全額取崩"]
    ax.plot([0, 1], [K8, K9], color=NAVY, linewidth=3, marker="o", markersize=12,
            markerfacecolor=NAVY, markeredgecolor="white", markeredgewidth=2)
    for xi, y, fc in ((2, A, LBLUE), (3, B, LORANGE), (4, C, LGREEN)):
        ax.plot([1, xi], [K9, y], color=GRAY, linewidth=2, linestyle="--",
                marker="s", markersize=10, markerfacecolor=fc,
                markeredgecolor=GRAY, markeredgewidth=2)
    ax.text(0, K8 - 130, f"{K8:,}円", ha="center", va="top", fontproperties=JP,
            fontsize=11, fontweight="bold", color=NAVY)
    ax.text(1, K9 - 130, f"{K9:,}円", ha="center", va="top", fontproperties=JP,
            fontsize=11, fontweight="bold", color=NAVY)
    for xi, y in ((2, A), (3, B), (4, C)):
        d = y - K9
        sign = f"+{d:,}円" if d >= 0 else f"▲{-d:,}円"
        ax.text(xi, y + 130, f"{y:,}円\n（第9期比 {sign}）", ha="center",
                va="bottom", fontproperties=JP_S, fontsize=9, color=GRAY)
    ax.set_xticks(range(5))
    ax.set_xticklabels(periods, fontproperties=JP_S, fontsize=9)
    ax.set_ylabel("月額基準額（円）", fontproperties=JP)
    ax.set_title("図9-1　介護保険料月額基準額の推移（第10期は令和8年9月時点の算定）",
                 fontproperties=JP_T, color=NAVY, pad=15)
    ax.set_ylim(5_300, 7_900)
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.text(0.02, 0.24,
            "※第10期は令和7年度（完結年度）の実績を基準とし、要介護度別の認定者数の伸び\n"
            "　のみを乗じて算定した。利用率・1人1月あたり利用回（日）数・単価は据置き。\n"
            "　所得段階別第1号被保険者数（第10〜13段階）の確認により、さらに下がる余地がある。",
            transform=ax.transAxes, fontproperties=JP_S, fontsize=9, color=ORANGE,
            verticalalignment="top",
            bbox=dict(boxstyle="round,pad=0.5", facecolor="#FFF2CC",
                      edgecolor=ORANGE, linewidth=1))
    plt.tight_layout()
    p = f"{OUT}/chart_premium_v114.png"
    plt.savefig(p, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()
    return p


if __name__ == "__main__":
    import sys
    A = int(sys.argv[1]) if len(sys.argv) > 1 else 7174
    B = int(sys.argv[2]) if len(sys.argv) > 2 else 6488
    C = int(sys.argv[3]) if len(sys.argv) > 3 else 5803
    print(fig23())
    print(fig91(A, B, C))
