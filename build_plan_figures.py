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

OUT = "/home/user/repository/output/figures"
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
         labelfmt=None, legend=True, ylim=None, legend_ncol=3, pct=False):
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
                            xytext=(0, 2.5 if y >= 0 else -9), ha="center", fontsize=7)
    ax.set_xticks(list(xs))
    ax.set_xticklabels(cats, rotation=rot, ha="right" if rot else "center", fontsize=8)
    ax.set_ylabel(ylabel, fontsize=8.5)
    if ylim:
        ax.set_ylim(*ylim)
    if pct:
        ax.yaxis.set_major_formatter(PCT)
    ax.tick_params(axis="y", labelsize=8)
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


# ==================================================== 第2章第1節 高齢者の状況
# 前回計画14頁と同じ12年分に、第10期の更新分（令和6〜8年）を加える構成とする。
# 令和6〜8年の町別・年齢4区分別人口は3町からの提供を要するため、
# 現時点では前回計画と同じ令和5年までを作図し、更新欄は本文の表に設けている。
NEN = ["平成24年", "平成25年", "平成26年", "平成27年", "平成28年", "平成29年",
       "平成30年", "令和元年", "令和2年", "令和3年", "令和4年", "令和5年"]

# 図1 人口の推移（広域連合）　前回計画14頁と同じ積上げ縦棒
stackbar("fig01_population", "人口の推移（広域連合）", NEN,
         [("0〜39歳", [10590, 10483, 10450, 10458, 10324, 10309,
                     10144, 9924, 9572, 9525, 9421, 9248]),
          ("40〜64歳", [9903, 9832, 9740, 9706, 9629, 9589,
                      9540, 9479, 9433, 9450, 9443, 9448]),
          ("65〜74歳", [3744, 3862, 4032, 4119, 4183, 4166,
                      4210, 4215, 4269, 4233, 4165, 3986]),
          ("75歳以上", [4398, 4491, 4578, 4667, 4768, 4865,
                     4933, 5006, 5014, 5039, 5109, 5205])],
         "（人）", figsize=(7.4, 3.6), legend_ncol=4)

# 図2 高齢化率の推移　前回計画14頁と同じ（広域連合＝棒、3町＝折れ線）
barline("fig02_aging_rate", "高齢化率の推移", NEN,
        "広域連合",
        [28.4, 29.1, 29.9, 30.3, 31.0, 31.2, 31.7, 32.2, 32.8, 32.8, 33.0, 33.0],
        [("東川町", [28.7, 29.7, 31.1, 31.5, 32.1, 32.0,
                   32.2, 32.1, 32.8, 32.3, 31.8, 31.4]),
         ("美瑛町", [33.6, 34.4, 35.2, 35.9, 36.5, 36.7,
                   37.3, 37.9, 38.4, 38.6, 38.7, 38.8]),
         ("東神楽町", [22.5, 23.0, 23.4, 23.8, 24.5, 25.2,
                    25.8, 26.7, 27.4, 27.7, 28.4, 28.7])],
        "", figsize=(7.6, 4.2), ylim=(20.0, 40.5), legend_ncol=4)

# 図6-1 認定率の3系列（新規・本検証の中心的所見）
bars("fig06_1_ninteiritsu", "年齢階級別にみた要介護認定率（粗認定率）",
     ["全年齢", "75歳以上", "85歳以上"],
     [("平成30年3月末（85歳以上は令和2年3月末）", [20.8, 35.3, 63.7]),
      ("令和8年3月末", [21.8, 33.5, 61.6])],
     "認定率（％）", labelfmt="%.1f", legend_ncol=2, figsize=(5.6, 3.0))

bars("fig06_2_chosei", "性・年齢調整済み要介護2以上認定率と変化率（令和5年調査）",
     ["大雪地区広域連合", "北海道", "全国"],
     [("認定率（％）", [9.25, 8.59, 8.99]),
      ("変化率（％）", [1.76, -1.71, -0.37])],
     "％", labelfmt="%.2f", legend_ncol=2)

# ==================================================== 第2章第2節 給付の分析
bars("fig18_kyufu", "地域差指数による給付水準の要因分解（全国＝1.00）",
     ["調整済み\n給付月額", "調整済み\n認定率", "受給率", "受給者単価"],
     [("大雪地区広域連合", [1.08, 1.02, 1.08, 1.06])],
     "全国＝1.00", labelfmt="%.2f", legend=False, ylim=(0.9, 1.15))

bars("fig20_jukyuritsu", "受給率の内訳（令和7年・第1号被保険者に対する割合）",
     ["在宅サービス", "施設及び居住系サービス", "合計"],
     [("大雪地区広域連合", [11.2, 5.3, 16.5]),
      ("北海道", [10.3, 4.5, 14.8]),
      ("全国", [11.0, 4.3, 15.3])],
     "受給率（％）", labelfmt="%.1f")

hbars("fig21_service", "サービス種類別　第1号1人当たり給付月額の増減率（平成30→令和7年度）",
      ["訪問看護", "訪問介護", "通所リハビリテーション", "福祉用具貸与",
       "認知症対応型共同生活介護", "地域密着型通所介護"],
      [69.2, 52.8, 38.6, 36.4, -10.5, -27.0], "増減率（％）", labelfmt="%+.1f",
      ref=0, highlight={0, 5})

line("fig17_riyoritsu", "介護サービス利用率の推移（受給者÷認定者）",
     ["H28", "R6"],
     [("介護サービス利用率（％）", [82.2, 74.3])],
     "利用率（％）", figsize=(4.2, 2.6), labelfmt="%.1f", ylim=(60, 90), legend_ncol=1)

bars("fig17_2_mishiyo", "認定者と受給者の推移",
     ["平成28年度", "令和6年度"],
     [("認定者", [1842, 1962]), ("受給者", [1515, 1458]), ("未利用認定者", [327, 504])],
     "人数（人）", labelfmt="%d", figsize=(5.2, 2.9))

bars("fig14_riyokyodo", "受給者1人当たり利用日数・回数（令和7年度）",
     ["訪問介護\n（回／月）", "訪問看護\n（回／月）", "通所介護\n（日／月）",
      "通所リハ\n（日／月）", "地域密着型\n通所介護（日／月）"],
     [("大雪地区広域連合", [55.4, 8.0, 8.8, 5.0, 7.9]),
      ("北海道", [29.9, 6.7, 8.1, 4.8, 7.8]),
      ("全国", [29.7, 9.1, 10.7, 5.6, 9.2])],
     "回数・日数", labelfmt="%.1f")

bars("fig22_kenoiki", "調整済み第1号1人当たり給付月額（北海道との比較）",
     ["施設及び居住系サービス", "在宅サービス"],
     [("大雪地区広域連合", [13097, 9164]), ("北海道", [10504, 8918])],
     "給付月額（円）", labelfmt="%d", figsize=(5.0, 2.9), legend_ncol=2)

# ==================================================== 第2章第3節 供給体制
line("fig24_1_jigyosho", "増減のあったサービスの事業所数の推移",
     ["H24", "H29", "H30", "R元", "R3", "R4", "R5", "R6"],
     [("訪問看護", [2, 4, 4, 4, 5, 6, 7, 7]),
      ("居宅療養管理指導", [0, 3, 3, 3, 4, 6, 8, 10]),
      ("訪問介護", [8, 10, 10, 10, 10, 10, 10, 13]),
      ("通所介護", [4, 5, 2, 2, 2, 2, 2, 2]),
      ("認知症対応型共同生活介護", [6, 6, 6, 6, 6, 5, 5, 5]),
      ("認知症対応型通所介護", [2, 0, 0, 0, 0, 0, 0, 0])],
     "事業所数（箇所）", yint=True, legend_ncol=3, figsize=(6.6, 3.4))

bars("fig24_2_zero", "域内に事業所が存在しないサービス（令和6年度・人口10万対）",
     ["定期巡回・随時\n対応型訪問介護看護", "看護小規模\n多機能型居宅介護",
      "認知症対応型\n通所介護", "福祉用具貸与", "訪問入浴介護", "介護医療院"],
     [("大雪地区広域連合", [0, 0, 0, 0, 0, 0]),
      ("北海道", [2.6, 1.6, 2.8, 6.0, 1.1, 1.0]),
      ("全国", [1.2, 0.9, 2.3, 6.0, 1.3, 0.8])],
     "事業所数（人口10万対）", labelfmt="%.1f", figsize=(6.6, 3.2))

hbars("fig24_3_hikaku", "人口10万対の事業所数　大雪／全国（令和6年度）",
      ["小規模多機能型居宅介護", "介護老人保健施設", "特定施設入居者生活介護",
       "訪問看護", "介護老人福祉施設", "訪問介護", "居宅介護支援",
       "地域密着型通所介護", "居宅療養管理指導", "通所介護"],
      [4.00, 3.18, 2.20, 1.71, 1.57, 1.56, 1.19, 0.93, 0.75, 0.35],
      "全国＝1.00", labelfmt="%.2f", ref=1.0, reflabel="全国", highlight={8, 9})

line("fig25_1_jujisha", "サービス別の従事者数の推移（実数）",
     ["H29", "H30", "R元", "R3", "R4", "R5", "R6"],
     [("介護老人保健施設", [90, 158, 165, 162, 160, 113, 124]),
      ("介護老人福祉施設", [116, 104, 123, 119, 115, 119, 105]),
      ("地域密着型介護老人福祉施設", [24, 40, 59, 84, 77, 67, 50]),
      ("訪問看護", [15, 22, 23, 19, 21, 29, 37])],
     "従事者数（人）", yint=True, legend_ncol=2, figsize=(6.6, 3.2))

line("fig25_3_shokushu", "職種別の従事者数の推移（介護老人福祉施設）",
     ["H29", "H30", "R元", "R3", "R4", "R5", "R6"],
     [("介護職員", [80, 66, 78, 74, 77, 74, 70]),
      ("准看護師", [7, 7, 6, 6, 5, 5, 3]),
      ("看護師", [5, 7, 8, 7, 7, 6, 7]),
      ("生活相談員", [5, 5, 5, 6, 5, 5, 3]),
      ("機能訓練指導員", [2, 3, 2, 2, 2, 2, 1])],
     "従事者数（人）", yint=True, legend_ncol=3, figsize=(6.6, 3.2))

# ==================================================== 第2章 その他
line("fig13_nenrei", "高齢者の年齢構成（5歳階級別）の推移と将来推計",
     ["R2", "R5", "R8", "R12", "R17", "R22", "R32"],
     [("65〜74歳", [4183, 3882, 3574, 3387, 3196, 3038, 3261]),
      ("75〜84歳", [3237, 3466, 3558, 3612, 3403, 3151, 3444]),
      ("85歳以上", [1930, 2021, 2190, 2320, 2673, 2810, 2534])],
     "人口（人）", legend_ncol=3, figsize=(6.2, 3.0))

# R9以降は見える化の自然体推計を、人口の基礎の変更（総合戦略ベース）に合わせて
# 認定者数の比で置き直した値である。第6章第3節3の表と同じ。
line("fig07_kyufuhi", "保険給付費・地域支援事業費の推移と中長期見通し",
     ["R3", "R4", "R5", "R9", "R11", "R17", "R22"],
     [("保険給付費（億円）", [29.92, 29.07, 29.40, 31.03, 32.19, 34.50, 36.06])],
     "億円", labelfmt="%.1f", legend_ncol=1, figsize=(6.2, 2.9))

# 見える化A1（国勢調査及び社人研 令和5年推計）の5年刻みの値による。
# 図表集16シート及び第2章第8節の本文表と同じ系列である。
line("fig15_machibetsu", "町別総人口の推移と将来推計",
     ["R2", "R7", "R12", "R17", "R22", "R27", "R32"],
     [("東川町", [8314, 8213, 8059, 7849, 7607, 7339, 7088]),
      ("美瑛町", [9668, 8893, 8160, 7476, 6851, 6258, 5681]),
      ("東神楽町", [10127, 9996, 9750, 9453, 9124, 8735, 8289])],
     "人口（人）", legend_ncol=3, figsize=(6.2, 3.0))

# 見える化A9（15〜64歳人口÷65歳以上人口）。図表集17シートと同じ系列である。
line("fig16_ninaite", "高齢者1人当たり現役世代数（15〜64歳人口÷65歳以上人口）",
     ["R2", "R7", "R12", "R17", "R22", "R27", "R32"],
     [("大雪地区広域連合", [1.60, 1.56, 1.48, 1.37, 1.19, 1.10, 1.04])],
     "人", labelfmt="%.2f", legend_ncol=1, figsize=(6.2, 2.8))

# ==================================================== 第2章第4節 ニーズ調査
bars("fig08_needs", "健康とくらしの調査の主要指標（令和7年度・広域連合）",
     ["フレイル\n該当割合", "月1回以上の\n社会参加率\n（6区分合成）", "通いの場\n参加率\n（ニーズ調査）"],
     [("令和4年度", [18.5, None, 7.3]), ("令和7年度", [19.1, 36.0, 8.8])],
     "割合（％）", labelfmt="%.1f", legend_ncol=2, figsize=(5.4, 3.0))

bars("fig08_2_machibetsu", "月1回以上の社会参加率（町別・令和7年度）",
     ["東川町", "美瑛町", "東神楽町", "広域連合計"],
     [("社会参加率（6区分合成）", [38.2, 35.4, 34.8, 36.0])],
     "割合（％）", labelfmt="%.1f", legend=False, ylim=(0, 45), figsize=(5.0, 2.7))

# ==================================================== 第3章 第9期の評価
bars("fig19_kpi", "第9期計画の代表KPIの達成状況",
     ["① 要介護認定率", "② 重度要介護認定率", "③ フレイル該当割合", "④ 通いの場参加率"],
     [("基準値", [20.8, 6.7, 18.5, 7.3]),
      ("第9期目標", [20.8, 6.4, 18.0, 10.0]),
      ("実績", [21.8, 6.3, 19.1, 8.8])],
     "％", labelfmt="%.1f", figsize=(6.4, 3.0))

line("fig23_kayoinoba", "通いの場への参加率（介護予防・日常生活支援総合事業ベース）",
     ["H25", "H26", "H27", "H28", "H29", "H30", "R元", "R2"],
     [("大雪地区広域連合", [1.1, 1.4, 1.4, 1.9, 3.1, 3.0, 2.7, 1.6]),
      ("北海道", [1.9, 1.9, 2.6, 2.8, 3.8, 3.9, 4.8, 3.9]),
      ("全国", [2.7, 3.2, 3.9, 4.2, 4.9, 5.7, 6.7, 5.2])],
     "月1回以上の参加率（％）", legend_ncol=3, figsize=(6.2, 3.0))

bars("fig23_2_kasho", "週1回以上の通いの場の箇所数（65歳以上1万人当たり・令和2年度）",
     ["大雪地区広域連合", "北海道", "全国"],
     [("箇所数", [3.24, 10.48, 13.12])],
     "箇所", labelfmt="%.2f", legend=False, figsize=(4.6, 2.6))

bars("fig26_1_kofukin", "保険者機能強化推進交付金等の総合得点（令和5年調査）",
     ["保険者機能強化\n推進交付金", "介護保険保険者\n努力支援交付金", "合計"],
     [("大雪地区広域連合", [164.7, 167.3, 332.0]),
      ("北海道", [210.4, 203.3, 413.6]),
      ("全国", [205.6, 216.7, 422.4])],
     "得点（点）", labelfmt="%.1f", figsize=(5.8, 3.0))

bars("fig26_2_shihyogun", "指標群別の得点（令和5年調査）",
     ["推進\n取組・体制", "推進\n活動", "推進\n成果", "支援\n取組・体制", "支援\n活動", "支援\n成果"],
     [("大雪地区広域連合", [96.7, 8.0, 60.0, 82.3, 25.0, 60.0]),
      ("全国", [122.5, 34.5, 48.6, 123.1, 45.0, 48.6])],
     "得点（点）", labelfmt="%.1f", legend_ncol=2, figsize=(6.4, 3.0))

hbars("fig26_3_mokuhyo", "目標別の得点　大雪／全国（令和5年調査）",
      ["共通Ⅳ 自立した日常生活（成果）", "推進Ⅲ 介護人材確保・基盤整備",
       "支援Ⅰ 介護予防／日常生活支援", "推進Ⅰ 持続可能な地域のあるべき姿",
       "支援Ⅱ 認知症総合支援", "推進Ⅱ 公正・公平な給付体制",
       "支援Ⅲ 在宅医療・在宅介護連携"],
      [1.235, 1.000, 0.918, 0.717, 0.666, 0.390, 0.382],
      "全国＝1.000", labelfmt="%.3f", ref=1.0, reflabel="全国平均", highlight={5, 6})

# ==================================================== 新規（追加調査の結果を可視化）
# 図27 特定地域の判定（第1章第9節）
fig, ax = plt.subplots(figsize=(6.0, 3.2))
towns = ["美瑛町", "東川町", "東神楽町"]
dens = [3.05, 6.18, 22.6]
cols = ["#1A1A1A", "#8C8C8C", "#C8C8C8"]
ax.bar(range(3), dens, color=[G_DARK, G_MID, G_LIGHT], edgecolor="black",
       linewidth=0.7, width=0.55)
for i, v in enumerate(dens):
    ax.annotate("%.1f" % v, (i, v), textcoords="offset points", xytext=(0, 3),
                ha="center", fontsize=9, fontweight="bold")
ax.axhline(5.0, color="black", linestyle="--", linewidth=1.2, zorder=3)
ax.annotate("基準②　5人／km²未満", (2.42, 5.0), textcoords="offset points",
            xytext=(0, 4), ha="right", fontsize=8)
ax.set_xticks(range(3))
ax.set_xticklabels(["%s\n（%s km²）" % (t, a) for t, a in
                    zip(towns, ["676.78", "247.30", "68.50"])], fontsize=8.5)
ax.set_ylabel("75歳以上人口密度（人／km²）", fontsize=8.5)
ax.set_ylim(0, 26)
ax.tick_params(axis="y", labelsize=8)
_title(ax, "特定地域の基準（75歳以上人口密度）による3町の位置")
_grid(ax)
_fin(fig, "fig27_tokutei")

# 図28 美瑛町のスクールバス路線別乗車人員（第1章第7節）
hbars("fig28_biei_bus", "美瑛町のスクールバス路線別乗車人員（令和2年度）",
      ["俵真布線（美瑛〜朗根内〜俵真布）", "宇莫別線（美瑛〜下宇莫別〜上宇莫別）",
       "美田・五稜線（美瑛〜美田〜五稜）", "旭線（美瑛〜北瑛〜旭）",
       "二股線（美瑛〜ルベシベ〜二股）", "置杵牛線（美瑛〜置杵牛）",
       "水沢線（美瑛〜春日台〜千代田）", "美馬牛線（美瑛〜美馬牛）"],
      [8188, 6268, 4817, 4612, 6268, 2888, 2643, 2408],
      "乗車人員（人／年）", labelfmt="%,d".replace(",", ""), highlight={0, 3})

# 図29 必要保険料月額と基準額の対比（第6章第6節）
bars("fig29_1_hokenryo", "必要保険料月額と条例上の基準額の対比（第9期）",
     ["令和6年度", "令和7年度"],
     [("必要保険料月額（推計）", [6063, 6352]),
      ("条例上の基準額", [6400, 6400])],
     "月額（円）", labelfmt="%d", legend_ncol=2, ylim=(5500, 6700), figsize=(4.8, 2.9))

# 図30 上川管内21保険者の保険料分布（第6章第6節）
IN_NAMES = ["当麻町", "愛別町", "鷹栖町", "中川町", "大雪地区広域連合", "比布町", "旭川市",
            "富良野市", "上川町", "剣淵町", "下川町", "和寒町", "美深町", "中富良野町",
            "南富良野町", "名寄市", "上富良野町", "占冠村", "士別市", "幌加内町", "音威子府村"]
IN_VALS = [6800, 6706, 6700, 6550, 6400, 6300, 6190, 6000, 6000, 6000, 6000,
           5950, 5900, 5700, 5700, 5400, 5400, 5100, 5025, 5000, 3600]
hbars("fig29_2_kannai", "上川管内21保険者の第9期保険料基準額（月額）",
      IN_NAMES, IN_VALS, "月額基準額（円）", labelfmt="%d",
      ref=5829.6, reflabel="管内単純平均 5,830円", highlight={4},
      figsize=(6.4, 4.6))

# 図31 KPIのデータ源の確保状況（第4章第3節）
stackh("fig30_kpi_source", "代表KPI16項目のデータ源の確保状況",
       ["第10期 代表KPI"],
       [("確保（基準値算定済み）", [6]), ("確保（調査の受領待ち）", [2]),
        ("様式整備が必要", [2]), ("データ源が未確保", [4]),
        ("抽出可否の確認が必要", [2])],
       "項目数", figsize=(6.6, 1.9), labelfmt="%d")


# 図31 第9期の対計画比（第3章第3節）
bars("fig31_taikeikakuhi", "第9期の対計画比（令和6・7年度）",
     ["第1号\n被保険者数", "要介護\n認定者数", "要介護\n認定率", "総給付費",
      "施設\nサービス", "居住系\nサービス", "在宅\nサービス",
      "1人あたり\n給付費"],
     [("令和6年度", [99.3, 97.7, 98.4, 96.3, 104.3, 83.8, 93.7, 97.0]),
      ("令和7年度", [98.5, 96.0, 97.4, 97.8, 100.8, 88.3, 98.4, 99.3])],
     "対計画比（％）", figsize=(6.8, 3.2), labelfmt="%.1f",
     ylim=(75, 112), legend_ncol=2)

# 図31-2 給付費3区分の対計画比の推移（第3章第3節）
line("fig31_2_kubun", "給付費3区分の対計画比の推移",
     ["H30", "R元", "R2", "R3", "R4", "R5", "R6", "R7"],
     [("施設サービス", [96, 94, 95, 95, 90, 90, 104, 101]),
      ("居住系サービス", [99, 101, 105, 94, 85, 82, 84, 88]),
      ("在宅サービス", [95, 97, 95, 96, 96, 96, 94, 98])],
     "対計画比（％）", figsize=(6.8, 3.0), ylim=(75, 112), labelfmt="%d")

# 図32 計画との乖離が大きいサービス（第3章第3節）
bars("fig32_kairi", "計画との乖離が大きいサービスの対計画比",
     ["定期巡回・随時対応型訪問介護看護", "地域密着型通所介護",
      "短期入所生活介護", "訪問看護", "特定施設入居者生活介護",
      "特定福祉用具販売", "居宅療養管理指導", "住宅改修",
      "訪問入浴介護"],
     [("令和6年度", [0.0, 66.5, 67.7, 76.5, 79.0, 106.3, 122.7, 125.4, 142.9]),
      ("令和7年度", [19.5, 66.6, 65.0, 95.6, 84.8, 128.5, 154.9, 126.3,
                 166.9])],
     "対計画比（％）", figsize=(7.2, 4.0), labelfmt="%.0f", rot=30,
     ylim=(0, 185), legend_ncol=2)

# 図33 期別の保険料基準額（第6章第6節）
bars("fig33_hokenryo", "期別の保険料基準額（月額・準備基金取崩後）",
     ["第7期", "第8期", "第9期"],
     [("計画値", [6077, 6237, 6428]),
      ("実績値", [6021, 6334, 6069])],
     "基準額（円）", figsize=(5.6, 3.0), labelfmt="%d",
     ylim=(5800, 6600), legend_ncol=2)

# 図33-2 第9期から第10期への要因分解（第6章第6節）
line("fig33_2_yoin", "第9期から第10期への要因分解（月額基準額）",
     ["第9期の再現", "給付費の水準", "地域支援事業費",
      "基金取崩額", "補正後被保険者数\nの係数", "被保険者数"],
     [("累計の月額基準額", [6428, 6767, 6738, 6863, 6710, 6755])],
     "月額基準額（円）", figsize=(6.8, 3.0), ylim=(6300, 6900),
     labelfmt="%d")

print("saved figures:", len(_saved))
for n in _saved:
    print("  ", n)

# 図3 高齢者を含む世帯の構成割合（前回計画15頁と同じ100％積上げ横棒）
stackh("fig03_setai", "高齢者を含む世帯の構成割合",
       ["広域連合\n(平成22年度)", "広域連合\n(平成27年度)", "広域連合\n(令和2年度)",
        "東川町\n(令和2年度)", "美瑛町\n(令和2年度)", "東神楽町\n(令和2年度)"],
       [("高齢者夫婦世帯", [29.9, 31.9, 34.7, 36.8, 31.6, 37.1]),
        ("高齢者単身世帯", [22.0, 25.0, 29.1, 30.2, 30.3, 26.3]),
        ("その他の高齢者を含む世帯", [48.1, 43.1, 36.2, 33.0, 38.1, 36.6])],
       "（％）", figsize=(6.8, 3.4), sep_after=1)
