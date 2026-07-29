# -*- coding: utf-8 -*-
"""素案本文に差し込む図のPNG生成スクリプト.

図表集（第10期計画_図表集_白黒.xlsx）は編集可能な原本として維持し、
本スクリプトは計画本文へ差し込むための画像を生成する。

方針
  1 白黒印刷を前提とし、色ではなく濃淡・ハッチング・線種・マーカーで区別する
  2 和文はIPAゴシック。数値ラベルは読み取り優先で必要な箇所のみ付す
  3 出典・注記は本文側に記載し、図中には最小限とする
  4 出力先は output/figures/。ファイル名は figNN[-M]_<英名>.png

図番号は図表集のシート見出し（A1）と一致させる。
"""

import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import rcParams
from matplotlib.ticker import MaxNLocator

rcParams["font.family"] = "IPAGothic"
rcParams["axes.unicode_minus"] = False
rcParams["figure.dpi"] = 200
rcParams["savefig.dpi"] = 200
rcParams["savefig.bbox"] = "tight"
rcParams["savefig.pad_inches"] = 0.06
rcParams["axes.grid"] = True
rcParams["grid.color"] = "#CCCCCC"
rcParams["grid.linewidth"] = 0.5
rcParams["axes.edgecolor"] = "#333333"
rcParams["axes.linewidth"] = 0.8
rcParams["font.size"] = 9

OUT = "/home/user/repository/output/figures"
os.makedirs(OUT, exist_ok=True)

# 白黒印刷用の系列スタイル（順に適用）
GRAYS = ["#1A1A1A", "#6E6E6E", "#A8A8A8", "#D0D0D0", "#454545", "#8C8C8C"]
LSTY = ["-", "--", "-.", ":", (0, (5, 1, 1, 1)), (0, (3, 1, 1, 1, 1, 1))]
MARK = ["o", "s", "^", "D", "v", "P"]
HATCH = ["", "///", "...", "xxx", "\\\\\\", "+++"]

_saved = []


def _fin(fig, name, title=None, ax=None):
    if title:
        (ax or fig.gca()).set_title(title, fontsize=10, fontweight="bold", pad=8)
    p = os.path.join(OUT, name + ".png")
    fig.savefig(p, facecolor="white")
    plt.close(fig)
    _saved.append(name)
    return p


def line(name, title, xs, series, ylabel, figsize=(6.6, 3.2), ylim=None,
         labelfmt=None, labelidx=None, legend_ncol=3, yint=False):
    """series = [(ラベル, [値...]), ...]  値がNoneの点は描画しない。"""
    fig, ax = plt.subplots(figsize=figsize)
    for i, (lab, ys) in enumerate(series):
        ax.plot(xs, ys, label=lab, color=GRAYS[i % len(GRAYS)],
                linestyle=LSTY[i % len(LSTY)], marker=MARK[i % len(MARK)],
                markersize=3.4, linewidth=1.4, markerfacecolor="white",
                markeredgewidth=1.0)
        if labelfmt is not None and (labelidx is None or i in labelidx):
            for x, y in zip(xs, ys):
                if y is None:
                    continue
                ax.annotate(labelfmt % y, (x, y), textcoords="offset points",
                            xytext=(0, 5), ha="center", fontsize=7)
    ax.set_ylabel(ylabel, fontsize=8.5)
    if ylim:
        ax.set_ylim(*ylim)
    if yint:
        ax.yaxis.set_major_locator(MaxNLocator(integer=True))
    ax.tick_params(labelsize=8)
    ax.legend(fontsize=8, ncol=legend_ncol, frameon=False,
              loc="upper center", bbox_to_anchor=(0.5, -0.14))
    return _fin(fig, name, title, ax)


def bars(name, title, cats, series, ylabel, figsize=(6.6, 3.2), rot=0,
         labelfmt=None, legend=True, ylim=None, legend_ncol=3):
    """縦棒（系列を横に並べる）。"""
    fig, ax = plt.subplots(figsize=figsize)
    n = len(series)
    w = 0.8 / n
    xs = range(len(cats))
    for i, (lab, ys) in enumerate(series):
        pos = [x - 0.4 + w * (i + 0.5) for x in xs]
        vals = [float("nan") if y is None else y for y in ys]
        ax.bar(pos, vals, width=w * 0.92, label=lab,
               color=GRAYS[i % len(GRAYS)], edgecolor="#1A1A1A",
               linewidth=0.6, hatch=HATCH[i % len(HATCH)])
        if labelfmt:
            for x, y in zip(pos, ys):
                if y is None:
                    continue
                ax.annotate(labelfmt % y, (x, y), textcoords="offset points",
                            xytext=(0, 2 if y >= 0 else -9), ha="center", fontsize=6.8)
    ax.set_xticks(list(xs))
    ax.set_xticklabels(cats, rotation=rot,
                       ha="right" if rot else "center", fontsize=8)
    ax.set_ylabel(ylabel, fontsize=8.5)
    if ylim:
        ax.set_ylim(*ylim)
    ax.tick_params(axis="y", labelsize=8)
    if legend and n > 1:
        ax.legend(fontsize=8, ncol=legend_ncol, frameon=False,
                  loc="upper center", bbox_to_anchor=(0.5, -0.16 - (0.10 if rot else 0)))
    return _fin(fig, name, title, ax)


def hbars(name, title, cats, vals, xlabel, figsize=(6.6, None), labelfmt="%.2f",
          ref=None, reflabel=None, highlight=None):
    """横棒。highlight＝強調するインデックスの集合。"""
    h = figsize[1] or max(2.0, 0.30 * len(cats) + 0.9)
    fig, ax = plt.subplots(figsize=(figsize[0], h))
    ys = range(len(cats))
    cols = ["#1A1A1A" if (highlight and i in highlight) else "#9A9A9A"
            for i in range(len(cats))]
    ax.barh(list(ys), vals, color=cols, edgecolor="#1A1A1A", linewidth=0.6, height=0.68)
    ax.set_yticks(list(ys))
    ax.set_yticklabels(cats, fontsize=8)
    ax.invert_yaxis()
    ax.set_xlabel(xlabel, fontsize=8.5)
    ax.tick_params(axis="x", labelsize=8)
    lo, mx = min(vals + [0]), max(vals + [0])
    span = (mx - lo) or 1
    if ref is not None:
        ax.axvline(ref, color="#1A1A1A", linestyle="--", linewidth=1.0)
        if reflabel:
            ax.annotate(reflabel, (ref, 0.985), xycoords=("data", "axes fraction"),
                        fontsize=7.5, ha="center", va="top",
                        bbox=dict(boxstyle="round,pad=0.18", fc="white",
                                  ec="#1A1A1A", lw=0.5))
    for y, v in zip(ys, vals):
        off = span * 0.015 if v >= 0 else -span * 0.015
        ax.annotate(labelfmt % v, (v + off, y), va="center", fontsize=7.4,
                    ha="left" if v >= 0 else "right")
    ax.set_xlim(lo - span * 0.12 if lo < 0 else 0, mx + span * 0.16)
    ax.grid(axis="y", visible=False)
    return _fin(fig, name, title, ax)


def stackh(name, title, cats, segs, xlabel, figsize=(6.6, None), labelfmt="%.1f"):
    """100%積上げ横棒。segs = [(ラベル, [値...]), ...]"""
    h = figsize[1] or max(1.8, 0.42 * len(cats) + 1.0)
    fig, ax = plt.subplots(figsize=(figsize[0], h))
    ys = range(len(cats))
    left = [0.0] * len(cats)
    for i, (lab, vs) in enumerate(segs):
        ax.barh(list(ys), vs, left=left, label=lab, height=0.6,
                color=GRAYS[i % len(GRAYS)], edgecolor="#1A1A1A",
                linewidth=0.6, hatch=HATCH[i % len(HATCH)])
        for y, (v, l) in enumerate(zip(vs, left)):
            if v > max(sum(x) for x in zip(*[s[1] for s in segs])) * 0.05:
                ax.annotate(labelfmt % v, (l + v / 2, y), ha="center", va="center",
                            fontsize=7, color="white" if i % len(GRAYS) < 2 else "black")
        left = [a + b for a, b in zip(left, vs)]
    ax.set_yticks(list(ys))
    ax.set_yticklabels(cats, fontsize=8)
    ax.invert_yaxis()
    ax.set_xlabel(xlabel, fontsize=8.5)
    ax.tick_params(axis="x", labelsize=8)
    ax.legend(fontsize=8, ncol=len(segs), frameon=False,
              loc="upper center", bbox_to_anchor=(0.5, -0.22))
    ax.grid(axis="y", visible=False)
    return _fin(fig, name, title, ax)


# ==================================================== 第2章第1節 高齢者の状況
Y = ["H24", "H27", "R2", "R5", "R7", "R12", "R17", "R22", "R32"]

# 図1 人口の推移
line("fig01_population", "図1　総人口・高齢者人口の推移と将来推計（大雪地区広域連合）",
     ["R3", "R5", "R8", "R9", "R11", "R17", "R22", "R32"],
     [("総人口", [27900, 27505, 26876, 26649, 26195, 24778, 23582, 21058]),
      ("65歳以上", [9350, 9369, 9322, 9319, 9311, 9378, 9633, 9239]),
      ("75歳以上", [5300, 5487, 5748, None, None, None, None, None])],
     "人口（人）", legend_ncol=3)

# 図2 高齢化率
bars("fig02_aging_rate", "図2　高齢化率の比較（令和5年10月1日現在）",
     ["東川町", "美瑛町", "東神楽町", "広域連合計"],
     [("高齢化率", [31.4, 38.8, 28.7, 33.0])],
     "高齢化率（％）", labelfmt="%.1f", legend=False, ylim=(0, 45))

# 図6-1 認定率の3系列（新規・本検証の中心的所見）
bars("fig06_1_ninteiritsu", "図6-1　年齢階級別にみた要介護認定率（粗認定率）",
     ["全年齢", "75歳以上", "85歳以上"],
     [("平成30年3月末（85歳以上は令和2年3月末）", [20.8, 35.3, 63.7]),
      ("令和8年3月末", [21.8, 33.5, 61.6])],
     "認定率（％）", labelfmt="%.1f", legend_ncol=2, figsize=(5.6, 3.0))

bars("fig06_2_chosei", "図6-2　性・年齢調整済み要介護2以上認定率と変化率（令和5年調査）",
     ["大雪地区広域連合", "北海道", "全国"],
     [("認定率（％）", [9.25, 8.59, 8.99]),
      ("変化率（％）", [1.76, -1.71, -0.37])],
     "％", labelfmt="%.2f", legend_ncol=2)

# ==================================================== 第2章第2節 給付の分析
bars("fig18_kyufu", "図18　地域差指数による給付水準の要因分解（全国＝1.00）",
     ["調整済み\n給付月額", "調整済み\n認定率", "受給率", "受給者単価"],
     [("大雪地区広域連合", [1.08, 1.02, 1.08, 1.06])],
     "全国＝1.00", labelfmt="%.2f", legend=False, ylim=(0.9, 1.15))

bars("fig20_jukyuritsu", "図20　受給率の内訳（令和7年・第1号被保険者に対する割合）",
     ["在宅サービス", "施設及び居住系サービス", "合計"],
     [("大雪地区広域連合", [11.2, 5.3, 16.5]),
      ("北海道", [10.3, 4.5, 14.8]),
      ("全国", [11.0, 4.3, 15.3])],
     "受給率（％）", labelfmt="%.1f")

hbars("fig21_service", "図21　サービス種類別　第1号1人当たり給付月額の増減率（平成30→令和7年度）",
      ["訪問看護", "訪問介護", "通所リハビリテーション", "福祉用具貸与",
       "認知症対応型共同生活介護", "地域密着型通所介護"],
      [69.2, 52.8, 38.6, 36.4, -10.5, -27.0], "増減率（％）", labelfmt="%+.1f",
      ref=0, highlight={0, 5})

line("fig17_riyoritsu", "図17　介護サービス利用率の推移（受給者÷認定者）",
     ["H28", "R6"],
     [("介護サービス利用率（％）", [82.2, 74.3])],
     "利用率（％）", figsize=(4.2, 2.6), labelfmt="%.1f", ylim=(60, 90), legend_ncol=1)

bars("fig17_2_mishiyo", "図17-2　認定者と受給者の推移",
     ["平成28年度", "令和6年度"],
     [("認定者", [1842, 1962]), ("受給者", [1515, 1458]), ("未利用認定者", [327, 504])],
     "人数（人）", labelfmt="%d", figsize=(5.2, 2.9))

bars("fig14_riyokyodo", "図14　受給者1人当たり利用日数・回数（令和7年度）",
     ["訪問介護\n（回／月）", "訪問看護\n（回／月）", "通所介護\n（日／月）",
      "通所リハ\n（日／月）", "地域密着型\n通所介護（日／月）"],
     [("大雪地区広域連合", [55.4, 8.0, 8.8, 5.0, 7.9]),
      ("北海道", [29.9, 6.7, 8.1, 4.8, 7.8]),
      ("全国", [29.7, 9.1, 10.7, 5.6, 9.2])],
     "回数・日数", labelfmt="%.1f")

bars("fig22_kenoiki", "図22　調整済み第1号1人当たり給付月額（北海道との比較）",
     ["施設及び居住系サービス", "在宅サービス"],
     [("大雪地区広域連合", [13097, 9164]), ("北海道", [10504, 8918])],
     "給付月額（円）", labelfmt="%d", figsize=(5.0, 2.9), legend_ncol=2)

# ==================================================== 第2章第3節 供給体制
line("fig24_1_jigyosho", "図24-1　増減のあったサービスの事業所数の推移",
     ["H24", "H29", "H30", "R元", "R3", "R4", "R5", "R6"],
     [("訪問看護", [2, 4, 4, 4, 5, 6, 7, 7]),
      ("居宅療養管理指導", [0, 3, 3, 3, 4, 6, 8, 10]),
      ("訪問介護", [8, 10, 10, 10, 10, 10, 10, 13]),
      ("通所介護", [4, 5, 2, 2, 2, 2, 2, 2]),
      ("認知症対応型共同生活介護", [6, 6, 6, 6, 6, 5, 5, 5]),
      ("認知症対応型通所介護", [2, 0, 0, 0, 0, 0, 0, 0])],
     "事業所数（箇所）", yint=True, legend_ncol=3, figsize=(6.6, 3.4))

bars("fig24_2_zero", "図24-2　域内に事業所が存在しないサービス（令和6年度・人口10万対）",
     ["定期巡回・随時\n対応型訪問介護看護", "看護小規模\n多機能型居宅介護",
      "認知症対応型\n通所介護", "福祉用具貸与", "訪問入浴介護", "介護医療院"],
     [("大雪地区広域連合", [0, 0, 0, 0, 0, 0]),
      ("北海道", [2.6, 1.6, 2.8, 6.0, 1.1, 1.0]),
      ("全国", [1.2, 0.9, 2.3, 6.0, 1.3, 0.8])],
     "事業所数（人口10万対）", labelfmt="%.1f", figsize=(6.6, 3.2))

hbars("fig24_3_hikaku", "図24-3　人口10万対の事業所数　大雪／全国（令和6年度）",
      ["小規模多機能型居宅介護", "介護老人保健施設", "特定施設入居者生活介護",
       "訪問看護", "介護老人福祉施設", "訪問介護", "居宅介護支援",
       "地域密着型通所介護", "居宅療養管理指導", "通所介護"],
      [4.00, 3.18, 2.20, 1.71, 1.57, 1.56, 1.19, 0.93, 0.75, 0.35],
      "全国＝1.00", labelfmt="%.2f", ref=1.0, reflabel="全国", highlight={8, 9})

line("fig25_1_jujisha", "図25-1　サービス別の従事者数の推移（実数）",
     ["H29", "H30", "R元", "R3", "R4", "R5", "R6"],
     [("介護老人保健施設", [90, 158, 165, 162, 160, 113, 124]),
      ("介護老人福祉施設", [116, 104, 123, 119, 115, 119, 105]),
      ("地域密着型介護老人福祉施設", [24, 40, 59, 84, 77, 67, 50]),
      ("訪問看護", [15, 22, 23, 19, 21, 29, 37])],
     "従事者数（人）", yint=True, legend_ncol=2, figsize=(6.6, 3.2))

line("fig25_3_shokushu", "図25-3　職種別の従事者数の推移（介護老人福祉施設）",
     ["H29", "H30", "R元", "R3", "R4", "R5", "R6"],
     [("介護職員", [80, 66, 78, 74, 77, 74, 70]),
      ("准看護師", [7, 7, 6, 6, 5, 5, 3]),
      ("看護師", [5, 7, 8, 7, 7, 6, 7]),
      ("生活相談員", [5, 5, 5, 6, 5, 5, 3]),
      ("機能訓練指導員", [2, 3, 2, 2, 2, 2, 1])],
     "従事者数（人）", yint=True, legend_ncol=3, figsize=(6.6, 3.2))

# ==================================================== 第2章 その他
line("fig13_nenrei", "図13　高齢者の年齢構成（5歳階級別）の推移と将来推計",
     ["R2", "R5", "R8", "R12", "R17", "R22", "R32"],
     [("65〜74歳", [4183, 3882, 3574, 3387, 3196, 3038, 3261]),
      ("75〜84歳", [3237, 3466, 3558, 3612, 3403, 3151, 3444]),
      ("85歳以上", [1930, 2021, 2190, 2320, 2673, 2810, 2534])],
     "人口（人）", legend_ncol=3, figsize=(6.2, 3.0))

line("fig07_kyufuhi", "図7　保険給付費・地域支援事業費の推移と中長期見通し",
     ["R3", "R4", "R5", "R9", "R11", "R17", "R22"],
     [("保険給付費（億円）", [29.92, 29.07, 29.40, 30.57, 31.21, 33.42, 34.93])],
     "億円", labelfmt="%.1f", legend_ncol=1, figsize=(6.2, 2.9))

line("fig15_machibetsu", "図15　町別総人口の推移と将来推計",
     ["R2", "R5", "R8", "R12", "R17", "R22", "R32"],
     [("東川町", [8387, 8558, 8492, 8339, 8092, 7607, 7088]),
      ("美瑛町", [9676, 9471, 9143, 8697, 8069, 7495, 6480]),
      ("東神楽町", [9866, 9858, 9241, 9070, 8617, 8480, 7490])],
     "人口（人）", legend_ncol=3, figsize=(6.2, 3.0))

line("fig16_ninaite", "図16　高齢者1人当たり現役世代数（15〜64歳人口÷65歳以上人口）",
     ["R2", "R5", "R8", "R12", "R17", "R22", "R32"],
     [("大雪地区広域連合", [1.44, 1.38, 1.33, 1.28, 1.19, 1.10, 1.05])],
     "人", labelfmt="%.2f", legend_ncol=1, figsize=(6.2, 2.8))

# ==================================================== 第2章第4節 ニーズ調査
bars("fig08_needs", "図8　健康とくらしの調査の主要指標（令和7年度・広域連合）",
     ["フレイル\n該当割合", "月1回以上の\n社会参加率\n（6区分合成）", "通いの場\n参加率\n（ニーズ調査）"],
     [("令和4年度", [18.5, None, 7.3]), ("令和7年度", [19.1, 36.0, 8.8])],
     "割合（％）", labelfmt="%.1f", legend_ncol=2, figsize=(5.4, 3.0))

bars("fig08_2_machibetsu", "図8-2　月1回以上の社会参加率（町別・令和7年度）",
     ["東川町", "美瑛町", "東神楽町", "広域連合計"],
     [("社会参加率（6区分合成）", [38.2, 35.4, 34.8, 36.0])],
     "割合（％）", labelfmt="%.1f", legend=False, ylim=(0, 45), figsize=(5.0, 2.7))

# ==================================================== 第3章 第9期の評価
bars("fig19_kpi", "図19　第9期計画の代表KPIの達成状況",
     ["① 要介護認定率", "② 重度要介護認定率", "③ フレイル該当割合", "④ 通いの場参加率"],
     [("基準値", [20.8, 6.7, 18.5, 7.3]),
      ("第9期目標", [20.8, 6.4, 18.0, 10.0]),
      ("実績", [21.8, 6.3, 19.1, 8.8])],
     "％", labelfmt="%.1f", figsize=(6.4, 3.0))

line("fig23_kayoinoba", "図23　通いの場への参加率（介護予防・日常生活支援総合事業ベース）",
     ["H25", "H26", "H27", "H28", "H29", "H30", "R元", "R2"],
     [("大雪地区広域連合", [1.1, 1.4, 1.4, 1.9, 3.1, 3.0, 2.7, 1.6]),
      ("北海道", [1.9, 1.9, 2.6, 2.8, 3.8, 3.9, 4.8, 3.9]),
      ("全国", [2.7, 3.2, 3.9, 4.2, 4.9, 5.7, 6.7, 5.2])],
     "月1回以上の参加率（％）", legend_ncol=3, figsize=(6.2, 3.0))

bars("fig23_2_kasho", "図23-2　週1回以上の通いの場の箇所数（65歳以上1万人当たり・令和2年度）",
     ["大雪地区広域連合", "北海道", "全国"],
     [("箇所数", [3.24, 10.48, 13.12])],
     "箇所", labelfmt="%.2f", legend=False, figsize=(4.6, 2.6))

bars("fig26_1_kofukin", "図26-1　保険者機能強化推進交付金等の総合得点（令和5年調査）",
     ["保険者機能強化\n推進交付金", "介護保険保険者\n努力支援交付金", "合計"],
     [("大雪地区広域連合", [164.7, 167.3, 332.0]),
      ("北海道", [210.4, 203.3, 413.6]),
      ("全国", [205.6, 216.7, 422.4])],
     "得点（点）", labelfmt="%.1f", figsize=(5.8, 3.0))

bars("fig26_2_shihyogun", "図26-2　指標群別の得点（令和5年調査）",
     ["推進\n取組・体制", "推進\n活動", "推進\n成果", "支援\n取組・体制", "支援\n活動", "支援\n成果"],
     [("大雪地区広域連合", [96.7, 8.0, 60.0, 82.3, 25.0, 60.0]),
      ("全国", [122.5, 34.5, 48.6, 123.1, 45.0, 48.6])],
     "得点（点）", labelfmt="%.1f", legend_ncol=2, figsize=(6.4, 3.0))

hbars("fig26_3_mokuhyo", "図26-3　目標別の得点　大雪／全国（令和5年調査）",
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
b = ax.bar(range(3), dens, color=cols, edgecolor="#1A1A1A", linewidth=0.7, width=0.55)
for i, v in enumerate(dens):
    ax.annotate("%.1f" % v, (i, v), textcoords="offset points", xytext=(0, 3),
                ha="center", fontsize=9, fontweight="bold")
ax.axhline(5.0, color="#1A1A1A", linestyle="--", linewidth=1.2)
ax.annotate("基準②　5人／km²未満", (2.42, 5.0), textcoords="offset points",
            xytext=(0, 4), ha="right", fontsize=8)
ax.set_xticks(range(3))
ax.set_xticklabels(["%s\n（%s km²）" % (t, a) for t, a in
                    zip(towns, ["676.78", "247.30", "68.50"])], fontsize=8.5)
ax.set_ylabel("75歳以上人口密度（人／km²）", fontsize=8.5)
ax.set_ylim(0, 26)
ax.tick_params(axis="y", labelsize=8)
_fin(fig, "fig27_tokutei", "図27　特定地域の基準②（75歳以上人口密度）による3町の位置", ax)

# 図28 美瑛町のスクールバス路線別乗車人員（第1章第7節）
hbars("fig28_biei_bus", "図28　美瑛町のスクールバス路線別乗車人員（令和2年度）",
      ["俵真布線（美瑛〜朗根内〜俵真布）", "宇莫別線（美瑛〜下宇莫別〜上宇莫別）",
       "美田・五稜線（美瑛〜美田〜五稜）", "旭線（美瑛〜北瑛〜旭）",
       "二股線（美瑛〜ルベシベ〜二股）", "置杵牛線（美瑛〜置杵牛）",
       "水沢線（美瑛〜春日台〜千代田）", "美馬牛線（美瑛〜美馬牛）"],
      [8188, 6268, 4817, 4612, 6268, 2888, 2643, 2408],
      "乗車人員（人／年）", labelfmt="%,d".replace(",", ""), highlight={0, 3})

# 図29 必要保険料月額と基準額の対比（第6章第6節）
bars("fig29_1_hokenryo", "図29-1　必要保険料月額と条例上の基準額の対比（第9期）",
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
hbars("fig29_2_kannai", "図29-2　上川管内21保険者の第9期保険料基準額（月額）",
      IN_NAMES, IN_VALS, "月額基準額（円）", labelfmt="%d",
      ref=5829.6, reflabel="管内単純平均 5,830円", highlight={4},
      figsize=(6.4, 4.6))

# 図31 KPIのデータ源の確保状況（第4章第3節）
stackh("fig30_kpi_source", "図30　代表KPI16項目のデータ源の確保状況",
       ["第10期 代表KPI"],
       [("確保（基準値算定済み）", [6]), ("確保（調査の受領待ち）", [2]),
        ("様式整備が必要", [2]), ("データ源が未確保", [4]),
        ("抽出可否の確認が必要", [2])],
       "項目数", figsize=(6.6, 1.9), labelfmt="%d")

print("saved figures:", len(_saved))
for n in _saved:
    print("  ", n)
