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


def title_fig(fig, d):
    """縦に複数段ある図は、表題を図全体の上に置く。"""
    fig.suptitle(f"{d['no']}　{d['title']}", fontproperties=JP_T,
                 color=NAVY, y=0.985)


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


def fig25(d):
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


def fig26(d):
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



def fig24(d):
    """要介護認定者数と認定率の推移（縦3段・横軸だけを共有する）

    人・人・％の3つを1つの軸に重ねない（2軸の図にしない）ため、
    縦に3段へ分ける。
    """
    fig, axes = plt.subplots(3, 1, figsize=(8.4, 7.6), dpi=150,
                             sharex=True,
                             gridspec_kw=dict(height_ratios=[1, 1, 1.05],
                                              hspace=0.18))
    rows = d["rows"]
    lab = [r[0].replace("令和", "令和\n") for r in rows]
    hi = [r[1] for r in rows]          # 第1号被保険者数
    nin = [r[2] for r in rows]         # 認定者数
    ritsu = [r[3] for r in rows]       # 認定率
    kubun = [r[4] for r in rows]
    x = list(range(len(rows)))
    # 実績と見込みを塗り分ける（見込みは薄く）
    face_h = [BLUE if k == "実績" else LBLUE for k in kubun]
    face_n = [ORANGE if k == "実績" else LORANGE for k in kubun]

    ax = axes[0]
    ax.bar(x, hi, 0.58, color=face_h, edgecolor=BLUE, linewidth=1.4)
    for i, v in enumerate(hi):
        ax.text(i, v + 30, f"{v:,}", ha="center", fontproperties=JP_S)
    ax.set_ylabel("第1号被保険者数\n（人）", fontproperties=JP_M)
    ax.set_ylim(0, max(hi) * 1.30)
    ax.annotate(f"令和6→11年度　▲{hi[0] - hi[-1]:,}人"
                f"（▲{(hi[0] - hi[-1]) / hi[0] * 100:.1f}％）",
                xy=(0.99, 0.95), xycoords="axes fraction", ha="right",
                va="top", fontproperties=JP_M, color=BLUE, fontweight="bold")

    ax = axes[1]
    ax.bar(x, nin, 0.58, color=face_n, edgecolor=ORANGE, linewidth=1.4)
    for i, v in enumerate(nin):
        ax.text(i, v + 6, f"{v:,}", ha="center", fontproperties=JP_S)
    ax.set_ylabel("認定者数（第1号）\n（人）", fontproperties=JP_M)
    ax.set_ylim(0, max(nin) * 1.32)
    ax.annotate(f"令和6→11年度　＋{nin[-1] - nin[0]:,}人"
                f"（＋{(nin[-1] - nin[0]) / nin[0] * 100:.1f}％）",
                xy=(0.99, 0.95), xycoords="axes fraction", ha="right",
                va="top", fontproperties=JP_M, color=ORANGE, fontweight="bold")

    ax = axes[2]
    ax.plot(x, ritsu, "-", color=NAVY, linewidth=2.4, zorder=3)
    for i, v in enumerate(ritsu):
        mk = "o" if kubun[i] == "実績" else "s"
        fc = NAVY if kubun[i] == "実績" else "white"
        ax.plot(i, v, mk, color=fc, markeredgecolor=NAVY,
                markeredgewidth=1.8, markersize=10, zorder=4)
        ax.text(i, v + 0.13, f"{v:.1f}%", ha="center", fontproperties=JP_S,
                color=NAVY)
    ax.set_ylabel("認定率\n（％）", fontproperties=JP_M)
    ax.set_ylim(min(ritsu) - 0.7, max(ritsu) + 0.7)
    ax.annotate(f"令和6→11年度　＋{ritsu[-1] - ritsu[0]:.1f}ポイント",
                xy=(0.99, 0.08), xycoords="axes fraction", ha="right",
                fontproperties=JP_M, color=NAVY, fontweight="bold")
    ax.set_xticks(x)
    ax.set_xticklabels([f"{a}\n（{b}）" for a, b in zip(lab, kubun)],
                       fontproperties=JP_S)

    for ax in axes:
        ax.grid(axis="y", linestyle=":", alpha=0.4)
        for sp in ("top", "right"):
            ax.spines[sp].set_visible(False)
    axes[0].legend(handles=[
        Patch(facecolor=BLUE, edgecolor=BLUE, label="実績"),
        Patch(facecolor=LBLUE, edgecolor=BLUE, label="実績見込み・見込み"),
    ], prop=JP_S, loc="lower right", frameon=True, ncol=2)
    if d["note"]:
        fig.text(0.015, 0.002, d["note"], ha="left", va="bottom",
                 fontproperties=JP_S, color=GRAY)
    title_fig(fig, d)
    fig.subplots_adjust(top=0.93, bottom=0.17)
    return save(fig, d)


def fig31(d):
    """交付金の得点の3か年推移と指標群別の内訳

    積み上げの区分どうしが見分けにくくならないよう、
    白い2ptの隙間を置き、各区分に数値を直接書き入れる。
    """
    fig, ax = plt.subplots(figsize=(8.6, 5.6), dpi=150)
    rows = d["rows"]
    lab = [r[0] for r in rows]
    taisei = [r[1] for r in rows]
    katsudo = [r[2] for r in rows]
    seika = [r[3] for r in rows]
    goukei = [r[4] for r in rows]
    zenkoku = [r[5] for r in rows]
    x = list(range(len(rows)))
    w = 0.46
    series = [("体制・取組指標群（満点380点）", taisei, BLUE),
              ("活動指標群（満点220点）", katsudo, ORANGE),
              ("成果指標群（満点200点）", seika, GREEN)]
    bottom = [0] * len(rows)
    for name, vals, col in series:
        ax.bar(x, vals, w, bottom=bottom, color=col,
               edgecolor="white", linewidth=2, label=name)
        for i, v in enumerate(vals):
            if v >= 25:        # 狭い区分には書き入れない（外に出す）
                ax.text(i, bottom[i] + v / 2, f"{v}", ha="center",
                        va="center", fontproperties=JP_M, color="white",
                        fontweight="bold")
            else:
                ax.annotate(f"{v}", xy=(i + w / 2, bottom[i] + v / 2),
                            xytext=(10, 0), textcoords="offset points",
                            va="center", fontproperties=JP_S, color=col,
                            fontweight="bold")
        bottom = [b + v for b, v in zip(bottom, vals)]
    for i, v in enumerate(goukei):
        ax.text(i, v + 26, f"{v}点", ha="center", fontproperties=JP_L,
                color=NAVY, fontweight="bold")
    ax.plot(x, zenkoku, "--", color=GRAY, linewidth=2, zorder=5,
            label="全国平均（合計）")
    # 値は最後の年度にだけ添える（すべての点に数を書かない）
    for i, v in enumerate(zenkoku):
        ax.plot(i, v, "D", color="white", markeredgecolor=GRAY,
                markeredgewidth=1.8, markersize=9, zorder=6)
    ax.annotate(f"全国平均 {zenkoku[-1]:.1f}点",
                xy=(len(zenkoku) - 1, zenkoku[-1]), xytext=(10, -48),
                textcoords="offset points", ha="left", fontproperties=JP_S,
                color=GRAY, zorder=7,
                arrowprops=dict(arrowstyle="-", color=GRAY, linewidth=0.8),
                bbox=dict(boxstyle="round,pad=0.3", facecolor="white",
                          edgecolor=GRAY, linewidth=0.8))
    ax.set_xticks(x)
    ax.set_xticklabels(lab, fontproperties=JP)
    ax.set_ylabel("得点（点）", fontproperties=JP)
    ax.set_ylim(0, 640)
    ax.set_xlim(-0.6, len(rows) - 0.25)
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    ax.legend(prop=JP_S, loc="upper center", frameon=True, ncol=4,
              bbox_to_anchor=(0.5, -0.10))
    if d["note"]:
        fig.text(0.015, 0.002, d["note"], ha="left", va="bottom",
                 fontproperties=JP_S, color=GRAY)
    title(ax, d)
    fig.subplots_adjust(bottom=0.27)
    return save(fig, d)


def fig61(d):
    """認知症施策のKPIの3層構造（概念図）

    数値の軸を持たない。プロセス（下）→アウトプット→アウトカム（上）の
    連なりを下から積み上げ、層ごとに指標を並べる。
    層の高さは指標の数から求める（見出し分＋1指標あたりの行の高さ）。
    """
    import collections
    from matplotlib.patches import FancyArrow, FancyBboxPatch

    g = collections.OrderedDict()
    for layer, mok, name, joukyou in d["rows"]:
        g.setdefault(layer, []).append((mok, name, joukyou))
    layers = list(g.items())          # プロセス → アウトプット → アウトカム
    FILL = {"プロセス（活動量）": "#DAE3F3",
            "アウトプット（成果物）": "#BDD0EC",
            "アウトカム（住民の変化）": "#9DC3E6"}
    MARK = {"あり": "●", "未確定": "○", "新設": "◆"}

    HEAD, LINE, GAP, Y0 = 0.80, 0.46, 0.30, 1.05
    # 層の高さの合計から図の高さを決める（上の層がはみ出さないように）
    H = Y0 + sum(HEAD + LINE * len(v) for _, v in layers) \
        + GAP * (len(layers) - 1) + 0.45
    fig, ax = plt.subplots(figsize=(9.6, 0.72 * H), dpi=150)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, H)
    ax.axis("off")
    y = Y0
    for layer, items in layers:       # 下から上へ積む
        hh = HEAD + LINE * len(items)
        ax.add_patch(FancyBboxPatch(
            (1.60, y), 8.00, hh,
            boxstyle="round,pad=0.03,rounding_size=0.10",
            facecolor=FILL[layer], edgecolor=BLUE, linewidth=1.6, zorder=2))
        ax.text(1.85, y + hh - 0.22, f"{layer}　{len(items)}指標",
                fontproperties=JP_L, color=NAVY, fontweight="bold",
                va="top", zorder=3)
        for j, (mok, name, joukyou) in enumerate(items):
            ty = y + hh - HEAD - j * LINE - 0.06
            ax.text(2.00, ty, f"{MARK[joukyou]} {name}", fontproperties=JP_M,
                    color="#17365D", va="top", zorder=3)
            ax.text(9.40, ty, mok, fontproperties=JP_S, color=GRAY,
                    ha="right", va="top", zorder=3)
        y += hh + GAP

    top = y - GAP
    ax.add_patch(FancyArrow(0.80, Y0, 0, top - Y0, width=0.30,
                            head_width=0.76, head_length=0.50,
                            length_includes_head=True, facecolor=LORANGE,
                            edgecolor=ORANGE, linewidth=1.4, zorder=1))
    ax.text(0.80, (Y0 + top) / 2, "施 策 の 効 果", fontproperties=JP_M,
            color=ORANGE, rotation=90, ha="center", va="center",
            fontweight="bold", zorder=3)
    n = collections.Counter(r[3] for r in d["rows"])
    ax.text(1.60, 0.58,
            f"現状値　● あり {n['あり']}指標／○ 未確定 {n['未確定']}指標／"
            f"◆ 第10期に新設 {n['新設']}指標",
            fontproperties=JP_M, color=NAVY, va="center")
    ax.text(1.60, 0.20,
            "各行の右端は、国の認知症施策推進基本計画が掲げる重点目標"
            "（1〜3）との対応",
            fontproperties=JP_S, color=GRAY, va="center")
    ax.set_title(f"{d['no']}　{d['title']}", fontproperties=JP_T,
                 color=NAVY, pad=12)
    return save(fig, d)


DRAW = {"図2-1": fig21, "図2-2": fig22, "図2-3": fig23,
        "図2-4": fig24, "図2-5": fig25, "図2-6": fig26,
        "図3-1": fig31, "図6-1": fig61, "図9-1": fig91}


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
