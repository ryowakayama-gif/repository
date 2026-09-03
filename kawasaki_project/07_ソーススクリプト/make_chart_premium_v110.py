# -*- coding: utf-8 -*-
"""
図7-1（介護保険料月額基準額の推移）の差替え用グラフを生成する。

素案 Ver.1.9 までの図7-1 は第10期を「約6,650円／約6,500円／約6,350円」という
仮置きの例示値で描いており、令和8年9月2日の概算試算（A 7,703円・B 7,023円・
C 6,344円）と整合しない。Ver.1.10 で本文・表に試算値を掲載するのに合わせ、
同じ配色・体裁のまま数値を試算値に差し替える。

出力：08_図表/chart_premium_v110.png
"""
import matplotlib

matplotlib.use("Agg")
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import numpy as np

import os

# 原本の図は Noto Sans CJK Medium で作成されている。当環境に同フォントがない場合は
# IPAゴシックで代替する（字形は異なるが、体裁・配色は原本に合わせている）。
FONT = next(f for f in ("/usr/share/fonts/opentype/noto/NotoSansCJK-Medium.ttc",
                        "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf")
            if os.path.exists(f))
JP = fm.FontProperties(fname=FONT, size=10)
JP_S = fm.FontProperties(fname=FONT, size=8)
JP_T = fm.FontProperties(fname=FONT, size=14)
plt.rcParams["axes.unicode_minus"] = False

NAVY, LBLUE, ORANGE, LORANGE = "#1F3864", "#DAE3F3", "#ED7D31", "#FCE4D6"
LGREEN, GRAY = "#E2EFDA", "#808080"

OUT = "08_図表/chart_premium_v110.png"

# 令和8年9月2日時点の概算試算（中位ケース・calc_hokenryo_R8.9.2.py）
K8, K9 = 6_380, 6_500
A, B, C = 7_703, 7_023, 6_344


def main():
    fig, ax = plt.subplots(figsize=(9, 5.5), dpi=150)

    periods = ["第8期\n(R3〜R5)", "第9期\n(R6〜R8)",
               "第10期\n(R9〜R11)\n仮試算A\n基金取崩なし",
               "第10期\n(R9〜R11)\n仮試算B\n基金50%取崩",
               "第10期\n(R9〜R11)\n仮試算C\n基金全額取崩"]
    x = np.arange(len(periods))

    # 確定分（実線）
    ax.plot([0, 1], [K8, K9], color=NAVY, linewidth=3, marker="o", markersize=12,
            markerfacecolor=NAVY, markeredgecolor="white", markeredgewidth=2)
    # 仮試算（破線）
    for xi, y, fc in ((2, A, LBLUE), (3, B, LORANGE), (4, C, LGREEN)):
        ax.plot([1, xi], [K9, y], color=GRAY, linewidth=2, linestyle="--",
                marker="s", markersize=10, markerfacecolor=fc,
                markeredgecolor=GRAY, markeredgewidth=2)

    ax.text(0, K8 - 130, f"{K8:,}円", ha="center", va="top",
            fontproperties=JP, fontsize=11, fontweight="bold", color=NAVY)
    ax.text(1, K9 - 130, f"{K9:,}円", ha="center", va="top",
            fontproperties=JP, fontsize=11, fontweight="bold", color=NAVY)
    for xi, y, d in ((2, A, A - K9), (3, B, B - K9), (4, C, C - K9)):
        sign = f"+{d:,}円" if d >= 0 else f"▲{-d:,}円"
        ax.text(xi, y + 130, f"【仮試算】\n{y:,}円\n（第9期比 {sign}）",
                ha="center", va="bottom", fontproperties=JP_S, fontsize=9,
                color=GRAY, style="italic")

    ax.set_xticks(x)
    ax.set_xticklabels(periods, fontproperties=JP_S, fontsize=9)
    ax.set_ylabel("月額基準額（円）", fontproperties=JP)
    ax.set_title("図7-1　介護保険料月額基準額の推移（第10期は令和8年9月時点の仮試算）",
                 fontproperties=JP_T, color=NAVY, pad=15)
    ax.set_ylim(6_000, 8_400)
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    ax.text(0.02, 0.97,
            "※第10期は令和8年9月2日時点の資料による仮試算（中位ケース）。\n"
            "　国保連データ・所得段階別被保険者数・調整交付金の交付実績・\n"
            "　令和9年度の介護報酬改定の確定を経て、第3回策定委員会で確定する。",
            transform=ax.transAxes, fontproperties=JP_S, fontsize=9, color=ORANGE,
            verticalalignment="top", style="italic",
            bbox=dict(boxstyle="round,pad=0.5", facecolor="#FFF2CC",
                      edgecolor=ORANGE, linewidth=1))

    plt.tight_layout()
    plt.savefig(OUT, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()
    print(f"保存：{OUT}")


if __name__ == "__main__":
    main()
