# -*- coding: utf-8 -*-
"""計画本文に差し込む図を `data_zuhyo.py` の数値から作り直す

図の数値・表題・出典はすべて `data_zuhyo.py` にあり、本スクリプトは形だけを持つ。
数値を直すときは `data_zuhyo.py`（又は図表データ管理台帳の xlsx）を直し、
本スクリプトを実行し直す。

  出力　08_図表/fig2-1_nenrei.png ほか6点

これを置く前は、図の数値が create_charts.py・make_charts_v114/v115/v116.py の
4つに散らばっており、図9-1 が素案Ver.2.3の是正（6,170→6,143円・5,518→5,464円）に
追随していなかった。また画像に焼き込んだ図番号が素案Ver.2.5の採番と
食い違っていた。
"""
import os
import sys

import matplotlib

matplotlib.use("Agg")
import matplotlib.font_manager as fm  # noqa: E402
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.patches import Patch  # noqa: E402

sys.path.insert(0, "07_ソーススクリプト")
from data_zuhyo import ZU  # noqa: E402

FONT = next(f for f in ("/usr/share/fonts/opentype/noto/NotoSansCJK-Medium.ttc",
                        "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf",
                        "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf")
            if os.path.exists(f))
JP = fm.FontProperties(fname=FONT, size=10)
JP_S = fm.FontProperties(fname=FONT, size=8)
JP_M = fm.FontProperties(fname=FONT, size=9)
JP_L = fm.FontProperties(fname=FONT, size=12)
JP_T = fm.FontProperties(fname=FONT, size=14)
plt.rcParams["axes.unicode_minus"] = False

NAVY, BLUE, LBLUE = "#1F3864", "#2F5597", "#DAE3F3"
ORANGE, LORANGE = "#ED7D31", "#FCE4D6"
GREEN, LGREEN = "#548235", "#E2EFDA"
RED, GRAY = "#C00000", "#808080"
OUT = "08_図表"


def title(ax, d):
    ax.set_title(f"{d['no']}　{d['title']}", fontproperties=JP_T,
                 color=NAVY, pad=15)


def save(fig, d):
    os.makedirs(OUT, exist_ok=True)
    p = os.path.join(OUT, d["png"])
    fig.savefig(p, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return p


def fig21(d):
    fig, ax = plt.subplots(figsize=(8, 5), dpi=150)
    lab = [r[0].replace("（", "\n（") for r in d["rows"]]
    a = [r[1] for r in d["rows"]]
    b = [r[2] for r in d["rows"]]
    x = range(len(lab))
    w = 0.36
    ax.bar([i - w / 2 for i in x], a, w, color=NAVY, label=d["head"][1])
    ax.bar([i + w / 2 for i in x], b, w, color=ORANGE, label=d["head"][2])
    for i, (u, v) in enumerate(zip(a, b)):
        ax.text(i - w / 2, u + 20, f"{u:,}人", ha="center",
                fontproperties=JP_S)
        ax.text(i + w / 2, v + 20, f"{v:,}人", ha="center",
                fontproperties=JP_S)
        r = (v - u) / u * 100
        col = RED if r < 0 else GREEN
        ax.annotate(f"{r:+.1f}%", xy=(i, max(u, v) + 210), ha="center",
                    fontproperties=JP_S, color=col, fontweight="bold")
    ax.set_xticks(list(x))
    ax.set_xticklabels(lab, fontproperties=JP)
    ax.set_ylabel("人数（人）", fontproperties=JP)
    ax.set_ylim(0, max(max(a), max(b)) * 1.28)
    ax.legend(prop=JP, frameon=True)
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    title(ax, d)
    return save(fig, d)


def fig22(d):
    fig, ax = plt.subplots(figsize=(8, 4.2), dpi=150)
    lab = [r[0] for r in d["rows"]]
    v = [r[1] for r in d["rows"]]
    bars = ax.barh(lab, v, color=[ORANGE, BLUE, GRAY],
                   edgecolor="white", linewidth=2)
    for b, x in zip(bars, v):
        ax.text(x + 0.6, b.get_y() + b.get_height() / 2, f"{x}%",
                va="center", fontproperties=JP_L, fontweight="bold")
    ax.set_xlabel("高齢化率（％）", fontproperties=JP)
    ax.set_yticks(range(len(lab)))
    ax.set_yticklabels(lab, fontproperties=JP_L)
    ax.set_xlim(0, 50)
    ax.grid(axis="x", linestyle=":", alpha=0.4)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.invert_yaxis()
    if d["note"]:
        ax.text(0.98, 0.06, d["note"], transform=ax.transAxes, ha="right",
                fontproperties=JP_M, color=ORANGE, fontweight="bold",
                bbox=dict(boxstyle="round,pad=0.5", facecolor=LORANGE,
                          edgecolor=ORANGE, linewidth=1))
    title(ax, d)
    return save(fig, d)


def fig23(d):
    fig, ax = plt.subplots(figsize=(9, 4.6), dpi=150)
    lab = [r[0].replace("（基準）", "\n（基準）") for r in d["rows"]]
    v = [r[1] for r in d["rows"]]
    tot = sum(v)
    fc = [LGREEN] * 3 + [LBLUE] + [ORANGE] + [LORANGE] * 8
    ec = [GREEN] * 3 + [BLUE] + [NAVY] + [ORANGE] * 8
    ax.bar(range(len(v)), v, color=fc, edgecolor=ec, linewidth=1.5)
    for i, x in enumerate(v):
        ax.text(i, x + 12, f"{x}", ha="center", fontproperties=JP_S)
    ax.set_xticks(range(len(lab)))
    ax.set_xticklabels([s.replace("第", "第\n") if len(s) > 4 else s
                        for s in lab], fontproperties=JP_S)
    ax.set_ylabel("人数（人）", fontproperties=JP)
    ax.set_ylim(0, max(v) * 1.22)
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    hikazei = sum(v[:3])
    ax.legend(handles=[
        Patch(facecolor=LGREEN, edgecolor=GREEN,
              label=f"非課税層（第1〜3段階）{hikazei}人・"
                    f"{hikazei / tot * 100:.1f}%"),
        Patch(facecolor=LBLUE, edgecolor=BLUE, label="本人非課税（第4段階）"),
        Patch(facecolor=ORANGE, edgecolor=NAVY, label="基準（第5段階）"),
        Patch(facecolor=LORANGE, edgecolor=ORANGE, label="課税層（第6〜13段階）"),
    ], prop=JP_S, loc="upper right", frameon=True)
    title(ax, d)
    return save(fig, d)


def fig24(d):
    fig, ax = plt.subplots(figsize=(7.4, 5.2), dpi=150)
    v = [r[1] for r in d["rows"]]
    tot = sum(v)
    lab = [f"{r[0]}\n{r[1]:,.1f}人 ({r[1] / tot * 100:.1f}%)"
           for r in d["rows"]]
    w, _ = ax.pie(v, labels=None, colors=[BLUE, GREEN, ORANGE],
                  startangle=90, counterclock=False,
                  wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2))
    ax.legend(w, lab, prop=JP, loc="center left",
              bbox_to_anchor=(0.98, 0.5), frameon=False)
    ax.text(0, 0.06, f"{tot:,.1f}人", ha="center", va="center",
            fontproperties=JP_T, color=NAVY)
    ax.text(0, -0.16, "延べ（重複あり）", ha="center", va="center",
            fontproperties=JP_S, color=GRAY)
    title(ax, d)
    return save(fig, d)


def fig25(d):
    fig, ax = plt.subplots(figsize=(8, 5), dpi=150)
    lab = [r[0] for r in d["rows"]]
    v = [r[1] for r in d["rows"]]
    tot = sum(v)
    bars = ax.bar(range(len(v)), v, color=[BLUE, GREEN, ORANGE],
                  edgecolor="white", linewidth=2, width=0.56)
    for b, x in zip(bars, v):
        ax.text(b.get_x() + b.get_width() / 2, x + 0.08,
                f"{x:.2f}億円\n({x / tot * 100:.1f}%)", ha="center",
                fontproperties=JP_S)
    ax.set_xticks(range(len(lab)))
    ax.set_xticklabels(lab, fontproperties=JP)
    ax.set_ylabel("給付費（億円）", fontproperties=JP)
    ax.set_ylim(0, max(v) * 1.28)
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    title(ax, d)
    return save(fig, d)


def fig91(d):
    fig, ax = plt.subplots(figsize=(9, 5.4), dpi=150)
    rows = d["rows"]
    lab = ["第8期\n(R3〜R5)", "第9期\n(R6〜R8)",
           "第10期\n(R9〜R11)\n算定Ａ\n基金取崩なし",
           "第10期\n(R9〜R11)\n算定Ｂ\n基金50%取崩",
           "第10期\n(R9〜R11)\n算定Ｃ\n基金全額取崩"]
    v = [r[1] for r in rows]
    ax.plot([0, 1], v[:2], "-", color=NAVY, linewidth=2.6, zorder=3)
    ax.plot(0, v[0], "o", color=NAVY, markersize=11, zorder=4)
    ax.plot(1, v[1], "s", color=LGREEN, markeredgecolor=GREEN,
            markeredgewidth=1.8, markersize=13, zorder=4)
    face = [None, None, LBLUE, LORANGE, LGREEN]
    edge = [None, None, BLUE, ORANGE, GREEN]
    for i in (2, 3, 4):
        ax.plot([1, i], [v[1], v[i]], "--", color=GRAY, linewidth=2, zorder=2)
        ax.plot(i, v[i], "s", color=face[i], markeredgecolor=edge[i],
                markeredgewidth=1.8, markersize=13, zorder=4)
        sa = rows[i][2]
        mark = "＋" if sa > 0 else "▲"
        ax.annotate(f"{v[i]:,}円\n（第9期比 {mark}{abs(sa):,}円）",
                    xy=(i, v[i]), xytext=(0, 26), textcoords="offset points",
                    ha="center", fontproperties=JP_M, color=GRAY)
    ax.annotate(f"{v[0]:,}円", xy=(0, v[0]), xytext=(0, -30),
                textcoords="offset points", ha="center",
                fontproperties=JP_L, color=NAVY)
    ax.annotate(f"{v[1]:,}円", xy=(1, v[1]), xytext=(0, -30),
                textcoords="offset points", ha="center",
                fontproperties=JP_L, color=NAVY)
    ax.set_xticks(range(len(lab)))
    ax.set_xticklabels(lab, fontproperties=JP_S)
    ax.set_ylabel("月額基準額（円）", fontproperties=JP)
    ax.set_xlim(-0.5, 4.5)
    ax.set_ylim(min(v) - 320, max(v) + 780)
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    if d["note"]:
        ax.text(0.02, 0.05, d["note"], transform=ax.transAxes, ha="left",
                va="bottom", fontproperties=JP_S, color=ORANGE,
                bbox=dict(boxstyle="round,pad=0.5", facecolor="#FFF7EF",
                          edgecolor=ORANGE, linewidth=1))
    title(ax, d)
    return save(fig, d)


DRAW = {"図2-1": fig21, "図2-2": fig22, "図2-3": fig23,
        "図2-4": fig24, "図2-5": fig25, "図9-1": fig91}


def main():
    made = []
    for d in ZU:
        if d["no"] not in DRAW:
            raise SystemExit(f"{d['no']} の作図がない")
        made.append(DRAW[d["no"]](d))
    print(f"図を{len(made)}点作りました")
    for p in made:
        print("  ", p, f"{os.path.getsize(p):,}バイト")
    # 自己点検　定義の数と作った画像の数
    ok = len(made) == len(ZU) and all(os.path.exists(p) for p in made)
    print("  定義", len(ZU), "／画像", len(made), "→", "適合" if ok else "不適合")
    if not ok:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
