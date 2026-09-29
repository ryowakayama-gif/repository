# -*- coding: utf-8 -*-
"""素案本文に差し込む図のPNG生成スクリプト.

図表集（第10期計画_図表集_白黒.xlsx）は編集可能な原本として維持し、
本スクリプトは計画本文へ差し込むための画像を生成する。

作図の様式は第9期計画（前回計画）に合わせている。
  1 図の表題は【　】で囲み、図の上に中央揃えで置く
  2 凡例は図の下に中央揃えで横並びに置く
  3 目盛線は横罫のみ。作図領域は細い実線で囲む
  4 白黒印刷を前提とし、濃淡・ハッチング・線種・マーカーで系列を区別する
  5 広域連合と構成3町を対比する図は、広域連合を棒グラフ、3町を折れ線で表す
    （前回計画14頁「高齢化率の推移」と同じ形式）
  6 出典は図の下に右寄せで「資料：〜」の1行を置く（本文側で付す）

図番号は図表集のシート見出し（A1）と一致させる。
"""

import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import rcParams
from matplotlib.ticker import MaxNLocator, FuncFormatter
import repo_paths as RP

rcParams["font.family"] = "IPAGothic"
rcParams["axes.unicode_minus"] = False
rcParams["figure.dpi"] = 220
rcParams["savefig.dpi"] = 220
rcParams["savefig.bbox"] = "tight"
rcParams["savefig.pad_inches"] = 0.05
rcParams["axes.grid"] = True
rcParams["axes.axisbelow"] = True
rcParams["grid.color"] = "#BFBFBF"
rcParams["grid.linewidth"] = 0.5
rcParams["axes.edgecolor"] = "#000000"
rcParams["axes.linewidth"] = 0.8
rcParams["font.size"] = 8.5
rcParams["legend.handlelength"] = 1.6
rcParams["legend.handleheight"] = 0.8
rcParams["legend.columnspacing"] = 1.4

OUT = RP.ROOT + "/output/figures"
os.makedirs(OUT, exist_ok=True)

# 前回計画の配色（白黒印刷前提のグレースケール）
G_DARK, G_MID, G_LIGHT, G_PALE = "#595959", "#A6A6A6", "#D9D9D9", "#F2F2F2"
GRAYS = [G_DARK, G_MID, G_LIGHT, G_PALE, "#7F7F7F", "#BFBFBF"]
HATCH = ["", "", "", "", "///", "..."]
# 折れ線（3町）の様式。前回計画14頁に合わせる
LINE_LABEL_UP = [True, True, False]   # 東川町・美瑛町は線の上、東神楽町は線の下
LINES = [dict(color="black", linestyle="-", marker="s", markerfacecolor="black"),
         dict(color="black", linestyle=":", marker="^", markerfacecolor="black"),
         dict(color="black", linestyle="--", marker="o", markerfacecolor="white")]

_saved = []
PCT = FuncFormatter(lambda v, _: "%.1f%%" % v)


def _title(ax, title):
    if title:
        ax.set_title("【" + title + "】", fontsize=10.5, pad=10, fontweight="normal")


def _grid(ax, axis="y"):
    ax.grid(axis=axis, color="#BFBFBF", linewidth=0.5)
    ax.grid(axis="x" if axis == "y" else "y", visible=False)
    for sp in ax.spines.values():
        sp.set_visible(True)
        sp.set_color("#000000")
        sp.set_linewidth(0.8)


def _legend(ax, ncol, y=-0.16):
    ax.legend(fontsize=8, ncol=ncol, frameon=False, loc="upper center",
              bbox_to_anchor=(0.5, y), handletextpad=0.5)


def _fin(fig, name):
    p = os.path.join(OUT, name + ".png")
    fig.savefig(p, facecolor="white")
    plt.close(fig)
    _saved.append(name)
    return p


def _box(ax, x, y, txt, fs=7, dy=0):
    """前回計画と同じ、白地・細枠の値ラベル。"""
    ax.annotate(txt, (x, y), textcoords="offset points", xytext=(0, dy),
                ha="center", va="center", fontsize=fs,
                bbox=dict(boxstyle="square,pad=0.20", fc="white", ec="black", lw=0.5))


# ------------------------------------------------------------------ 作図関数
def line(name, title, xs, series, ylabel, figsize=(6.6, 3.2), ylim=None,
         labelfmt=None, labelidx=None, legend_ncol=3, yint=False, pct=False):
    fig, ax = plt.subplots(figsize=figsize)
    for i, (lab, ys) in enumerate(series):
        st = LINES[i % len(LINES)] if len(series) <= 3 else {}
        ax.plot(xs, ys, label=lab, linewidth=1.3, markersize=4,
                markeredgecolor="black", markeredgewidth=0.9,
                color=st.get("color", "black"),
                linestyle=st.get("linestyle", ["-", ":", "--", "-.", (0, (5, 1, 1, 1))][i % 5]),
                marker=st.get("marker", ["s", "^", "o", "D", "v", "P"][i % 6]),
                markerfacecolor=st.get("markerfacecolor",
                                       ["black", "black", "white", "white", "black", "white"][i % 6]))
        if labelfmt is not None and (labelidx is None or i in labelidx):
            for x, y in zip(xs, ys):
                if y is None:
                    continue
                ax.annotate(labelfmt % y, (x, y), textcoords="offset points",
                            xytext=(0, 6), ha="center", fontsize=7)
    ax.set_ylabel(ylabel, fontsize=8.5)
    if ylim:
        ax.set_ylim(*ylim)
    if yint:
        ax.yaxis.set_major_locator(MaxNLocator(integer=True))
    if pct:
        ax.yaxis.set_major_formatter(PCT)
    ax.tick_params(labelsize=8)
    _grid(ax)
    _legend(ax, legend_ncol, -0.15)
    _title(ax, title)
    return _fin(fig, name)


def bars(name, title, cats, series, ylabel, figsize=(6.6, 3.2), rot=0,
         labelfmt=None, legend=True, ylim=None, legend_ncol=3, pct=False,
         xticklabels=None, hline=None, hlabel=None, labelbold=False):
    fig, ax = plt.subplots(figsize=figsize)
    n = len(series)
    w = 0.78 / n
    xs = range(len(cats))
    for i, (lab, ys) in enumerate(series):
        pos = [x - 0.39 + w * (i + 0.5) for x in xs]
        vals = [float("nan") if y is None else y for y in ys]
        ax.bar(pos, vals, width=w * 0.9, label=lab, color=GRAYS[i % len(GRAYS)],
               edgecolor="black", linewidth=0.6, hatch=HATCH[i % len(HATCH)])
        if labelfmt:
            for x, y in zip(pos, ys):
                if y is None:
                    continue
                ax.annotate(labelfmt % y, (x, y), textcoords="offset points",
                            xytext=(0, (3 if labelbold else 2.5) if y >= 0 else -9),
                            ha="center", fontsize=9 if labelbold else 7,
                            fontweight="bold" if labelbold else "normal")
    ax.set_xticks(list(xs))
    ax.set_xticklabels(xticklabels or cats, rotation=rot,
                       ha="right" if rot else "center", fontsize=8 if not xticklabels else 8.5)
    ax.set_ylabel(ylabel, fontsize=8.5)
    if ylim:
        ax.set_ylim(*ylim)
    if pct:
        ax.yaxis.set_major_formatter(PCT)
    ax.tick_params(axis="y", labelsize=8)
    if hline is not None:
        ax.axhline(hline, color="black", linestyle="--", linewidth=1.2, zorder=3)
        if hlabel:
            ax.annotate(hlabel, (len(cats) - 0.58, hline), textcoords="offset points",
                        xytext=(0, 4), ha="right", fontsize=8)
    _grid(ax)
    if legend and n > 1:
        _legend(ax, legend_ncol, -0.16 - (0.12 if rot else 0))
    _title(ax, title)
    return _fin(fig, name)


def barline(name, title, xs, bar_label, bar_vals, line_series, ylabel,
            figsize=(7.0, 3.4), ylim=None, pct=True, boxlabel=True,
            linefmt="%.1f%%", legend_ncol=4):
    """広域連合を棒、構成3町を折れ線で表す（前回計画14頁と同じ形式）。"""
    fig, ax = plt.subplots(figsize=figsize)
    xi = list(range(len(xs)))
    ax.bar(xi, bar_vals, width=0.62, label=bar_label, color=G_LIGHT,
           edgecolor="black", linewidth=0.6, zorder=1)
    for i, (lab, ys) in enumerate(line_series):
        st = LINES[i % len(LINES)]
        ax.plot(xi, ys, label=lab, linewidth=1.2, markersize=4.2, zorder=3,
                color=st["color"], linestyle=st["linestyle"], marker=st["marker"],
                markerfacecolor=st["markerfacecolor"], markeredgecolor="black",
                markeredgewidth=0.9)
        # ラベルの上下は系列ごとに固定し、棒の上端（広域連合）と近い年は退避させる
        for x, y, bv in zip(xi, ys, bar_vals):
            up = LINE_LABEL_UP[i % len(LINE_LABEL_UP)]
            dy = 9 if up else -11
            if up and abs(y - bv) < (ax.get_ylim()[1] - ax.get_ylim()[0]) * 0.055:
                dy = 19 if y >= bv else -13
            ax.annotate(linefmt % y, (x, y), textcoords="offset points",
                        xytext=(0, dy), ha="center", fontsize=6.4)
    if boxlabel:
        for x, y in zip(xi, bar_vals):
            _box(ax, x, y, "%.1f%%" % y, fs=6.4, dy=-8)
    ax.set_xticks(xi)
    ax.set_xticklabels(xs, fontsize=7.6)
    ax.set_ylabel(ylabel, fontsize=8.5)
    if ylim:
        ax.set_ylim(*ylim)
    if pct:
        ax.yaxis.set_major_formatter(PCT)
    ax.tick_params(axis="y", labelsize=8)
    _grid(ax)
    _legend(ax, legend_ncol, -0.15)
    _title(ax, title)
    return _fin(fig, name)


def stackbar(name, title, xs, segs, ylabel, figsize=(7.2, 3.6), total=True,
             labelfmt="%s", legend_ncol=5):
    """積上げ縦棒（前回計画14頁「人口の推移」と同じ形式）。"""
    fig, ax = plt.subplots(figsize=figsize)
    xi = list(range(len(xs)))
    bottom = [0] * len(xs)
    for i, (lab, vs) in enumerate(segs):
        ax.bar(xi, vs, bottom=bottom, width=0.66, label=lab,
               color=GRAYS[i % len(GRAYS)], edgecolor="black", linewidth=0.5)
        for x, (v, b) in enumerate(zip(vs, bottom)):
            ax.annotate("{:,}".format(v), (x, b + v / 2), ha="center", va="center",
                        fontsize=6.4, color="white" if i == 0 else "black")
        bottom = [a + b for a, b in zip(bottom, vs)]
    if total:
        for x, t in zip(xi, bottom):
            ax.annotate("{:,}".format(t), (x, t), textcoords="offset points",
                        xytext=(0, 3), ha="center", fontsize=6.8)
        ax.set_ylim(0, max(bottom) * 1.12)
    ax.set_xticks(xi)
    ax.set_xticklabels(xs, fontsize=7.6)
    ax.set_ylabel(ylabel, fontsize=8.5)
    ax.tick_params(axis="y", labelsize=8)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: "{:,.0f}".format(v)))
    _grid(ax)
    _legend(ax, legend_ncol, -0.14)
    _title(ax, title)
    return _fin(fig, name)


def hbars(name, title, cats, vals, xlabel, figsize=(6.6, None), labelfmt="%.2f",
          ref=None, reflabel=None, highlight=None):
    h = figsize[1] or max(2.0, 0.30 * len(cats) + 1.0)
    fig, ax = plt.subplots(figsize=(figsize[0], h))
    ys = range(len(cats))
    cols = [G_DARK if (highlight and i in highlight) else G_MID for i in range(len(cats))]
    ax.barh(list(ys), vals, color=cols, edgecolor="black", linewidth=0.6, height=0.66)
    ax.set_yticks(list(ys))
    ax.set_yticklabels(cats, fontsize=8)
    ax.invert_yaxis()
    ax.set_xlabel(xlabel, fontsize=8.5)
    ax.tick_params(axis="x", labelsize=8)
    lo, mx = min(vals + [0]), max(vals + [0])
    span = (mx - lo) or 1
    if ref is not None:
        ax.axvline(ref, color="black", linestyle="--", linewidth=1.0, zorder=3)
        if reflabel:
            ax.annotate(reflabel, (ref, 0.985), xycoords=("data", "axes fraction"),
                        fontsize=7.4, ha="center", va="top",
                        bbox=dict(boxstyle="square,pad=0.20", fc="white", ec="black", lw=0.5))
    for y, v in zip(ys, vals):
        off = span * 0.015 if v >= 0 else -span * 0.015
        ax.annotate(labelfmt % v, (v + off, y), va="center", fontsize=7.4,
                    ha="left" if v >= 0 else "right")
    ax.set_xlim(lo - span * 0.12 if lo < 0 else 0, mx + span * 0.18)
    _grid(ax, axis="x")
    _title(ax, title)
    return _fin(fig, name)


def stackh(name, title, cats, segs, xlabel, figsize=(6.8, None), labelfmt="%.1f",
           sep_after=None):
    """100%積上げ横棒（前回計画15頁と同じ形式）。"""
    h = figsize[1] or max(2.2, 0.46 * len(cats) + 1.1)
    fig, ax = plt.subplots(figsize=(figsize[0], h))
    ys = range(len(cats))
    left = [0.0] * len(cats)
    tot = [sum(x) for x in zip(*[s[1] for s in segs])]
    for i, (lab, vs) in enumerate(segs):
        ax.barh(list(ys), vs, left=left, label=lab, height=0.56,
                color=GRAYS[i % len(GRAYS)], edgecolor="black", linewidth=0.5)
        for y, (v, l) in enumerate(zip(vs, left)):
            if v > tot[y] * 0.06:
                ax.annotate(labelfmt % v, (l + v / 2, y), ha="center", va="center",
                            fontsize=7.2, color="white" if i == 0 else "black")
        left = [a + b for a, b in zip(left, vs)]
    ax.set_yticks(list(ys))
    ax.set_yticklabels(cats, fontsize=8)
    ax.invert_yaxis()
    ax.set_xlabel(xlabel, fontsize=8.5)
    ax.tick_params(axis="x", labelsize=8)
    if sep_after is not None:
        ax.axhline(sep_after + 0.5, color="#808080", linestyle=":", linewidth=0.9)
    _grid(ax, axis="x")
    _legend(ax, len(segs), -0.14 - 0.04 * (6 - len(cats)) if len(cats) < 6 else -0.16)
    _title(ax, title)
    return _fin(fig, name)


# ==================================================================== 作図の実行
# 数値は data_zuhyo.py を正本とし、図表データ管理台帳（xlsx）に数値があれば
# そちらを優先する（台帳で数値を直せば本文の図が追随する）。
# 図の体裁（種類・軸・凡例・ラベルの書式）は data_zuhyo.ZU の項目による。
import data_zuhyo as DZ

DRAW = {"line": line, "bars": bars, "hbars": hbars, "stackbar": stackbar,
        "barline": barline, "stackh": stackh}
# 作図関数が受け取らない項目（台帳の管理のためだけに持つもの）
META = ("no",)

for _d in DZ.load():
    _kw = {k: v for k, v in _d.items() if k not in META and k != "kind"}
    _name = _kw.pop("name")
    _kind = _d["kind"]
    if _kind == "line":
        _args = (_name, _kw.pop("title"), _kw.pop("xs"), _kw.pop("series"), _kw.pop("ylabel"))
    elif _kind == "bars":
        _args = (_name, _kw.pop("title"), _kw.pop("cats"), _kw.pop("series"), _kw.pop("ylabel"))
    elif _kind == "hbars":
        _args = (_name, _kw.pop("title"), _kw.pop("cats"), _kw.pop("vals"), _kw.pop("xlabel"))
    elif _kind == "stackbar":
        _args = (_name, _kw.pop("title"), _kw.pop("xs"), _kw.pop("segs"), _kw.pop("ylabel"))
    elif _kind == "stackh":
        _args = (_name, _kw.pop("title"), _kw.pop("cats"), _kw.pop("segs"), _kw.pop("xlabel"))
    else:  # barline
        _args = (_name, _kw.pop("title"), _kw.pop("xs"), _kw.pop("bar_label"),
                 _kw.pop("bar_vals"), _kw.pop("line_series"), _kw.pop("ylabel"))
    DRAW[_kind](*_args, **_kw)

print("saved figures:", len(_saved))
for n in _saved:
    print("  ", n)
