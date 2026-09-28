"""小野町 協議会資料の図。

他町村の案件で用いている「策定委員会資料のデザインルール」の図の規則に合わせた。
**規則のみを取り入れたものであり、図に用いる数値はすべて小野町のものである。**

体裁の規則
  配色　　強調・最大 #C00000／主 #2E75B6／淡 #BDD7EE／参照・基準 #ED7D31
  　　　　減少・対象外 #A6A6A6／数値ラベル #1F3864（太字）
  書体　　IPAGothic（本文のBIZ UDPゴシックに近いゴシック体）
  形　　　順位・比較は横棒、推移は折れ線または縦棒
  規則　　①値ラベルを必ず付す　②軸ラベルに単位を（）で入れる
  　　　　③基準値は破線＋ラベル　④凡例は系列が2つ以上のときのみ
  　　　　⑤背景は白、枠線あり、目盛線は用いない（横棒のみ縦の目盛線を薄く）
  　　　　⑥縦軸を0から始めない場合は出典行にその旨を書く
  出力　　PNG・dpi200・幅はA4本文幅（17.0cm）に収まる比率

色で意味を決める。同じ色が図をまたいで別の意味を持たないようにする。
  小野町＝赤（弱点・注目）または青（主系列）、県＝オレンジ、全国＝グレー。
"""

import pathlib
import warnings

warnings.filterwarnings("ignore")

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt

import build_ono_tanka as T

FONT = "IPAGothic"
plt.rcParams["font.family"] = FONT
plt.rcParams["axes.unicode_minus"] = False

RED = "#C00000"      # 強調・最大・注目
BLUE = "#2E75B6"     # 主系列
PALE = "#BDD7EE"     # 淡・参考
ORANGE = "#ED7D31"   # 参照値・基準・県
GREY = "#A6A6A6"     # 減少・対象外・全国
NAVY = "#1F3864"     # 数値ラベル

OUT = pathlib.Path(__file__).parent / "output" / "ono_fig"
YS = ["令和3年度", "令和4年度", "令和5年度", "令和6年度", "令和7年度"]


def _save(fig, name):
    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / ("%s.png" % name)
    fig.savefig(p, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return p


def _lab(ax, x, y, t, color=NAVY, ha="left", va="center", sz=10):
    ax.text(x, y, t, color=color, ha=ha, va=va, fontsize=sz, fontweight="bold")


def _pm(v, unit="", dec=1):
    """増減を自治体の慣行にそろえて表す。負は ▲ で示す。"""
    f = "%%.%df" % dec
    return ("▲" + f % abs(v) if v < 0 else "＋" + f % v) + unit


def _xgrid(ax):
    ax.xaxis.grid(True, color="#E6E6E6", lw=0.8)
    ax.set_axisbelow(True)


# ================================================================ 図の定義

def f_nintei_suii():
    """認定者数の推移（年度末基準）。令和7年度の▲152人を示す。"""
    lab = ["平成30年\n3月末", "令和元年\n3月末", "令和2年\n3月末", "令和3年\n3月末",
           "令和4年\n3月末", "令和5年\n3月末", "令和6年\n3月末", "令和7年\n3月末",
           "令和8年\n3月末"]
    val = [691, 722, 754, 861, 872, 898, 919, 944, 792]
    col = [BLUE] * 8 + [RED]
    fig, ax = plt.subplots(figsize=(7.6, 3.4))
    ax.bar(range(len(val)), val, width=0.62, color=col)
    for i, v in enumerate(val):
        if i == 8:                       # 最終年は棒の中に白字で置く
            _lab(ax, i, v - 52, "%d" % v, "#FFFFFF", ha="center", sz=9)
        else:
            _lab(ax, i, v + 16, "%d" % v, NAVY, ha="center", sz=9)
    _lab(ax, 8, 880, "▲152人\n（▲16.1％）", RED, ha="center", sz=9)
    ax.set_xticks(range(len(lab)))
    ax.set_xticklabels(lab, fontsize=8)
    ax.set_ylim(0, 1090)
    ax.set_ylabel("認定者数（人）")
    return _save(fig, "f01_nintei_suii")


def f_nintei_bunkai():
    """令和7年3月末→令和8年3月末の減少の年齢階層別内訳。"""
    lab = ["85歳以上", "65〜74歳", "75〜84歳"]
    dec = [-117, -25, -10]
    share = [77.0, 16.4, 6.6]
    col = [RED, GREY, GREY]
    y = range(len(lab))
    fig, ax = plt.subplots(figsize=(7.0, 2.5))
    ax.barh(list(y), dec, height=0.55, color=col)
    for i, (v, s) in enumerate(zip(dec, share)):
        _lab(ax, v - 4, i, "%s人（寄与%.1f％）" % (_pm(v, "", 0), s), col[i],
             ha="right", sz=10)
    ax.set_yticks(list(y))
    ax.set_yticklabels(lab, fontsize=10)
    ax.invert_yaxis()
    ax.set_xlim(-175, 12)
    ax.set_xlabel("第1号被保険者の認定者数の増減（人）")
    _xgrid(ax)
    return _save(fig, "f02_nintei_bunkai")


def f_kyufu_suii(nen):
    """総給付費と標準給付費の推移（年報 様式3）。"""
    sou = [(nen[y]["介護諸費"] + nen[y]["予防諸費"]) / 1e6 for y in YS]
    std = [nen[y]["支払計"] / 1e6 for y in YS]
    x = range(len(YS))
    fig, ax = plt.subplots(figsize=(7.4, 3.4))
    ax.plot(list(x), std, "-o", color=BLUE, lw=2.4, ms=7, label="標準給付費")
    ax.plot(list(x), sou, "-s", color=PALE, lw=2.4, ms=7, label="総給付費")
    for i, v in enumerate(std):
        _lab(ax, i, v + 22, "%.0f" % v, RED if i == 4 else NAVY, ha="center",
             sz=9)
    for i, v in enumerate(sou):
        _lab(ax, i, v - 30, "%.0f" % v, NAVY, ha="center", sz=9)
    ax.set_xticks(list(x))
    ax.set_xticklabels([y.replace("年度", "\n年度") for y in YS], fontsize=9)
    ax.set_ylim(930, 1230)
    ax.set_ylabel("給付費（百万円）")
    ax.legend(loc="lower left", frameon=False, fontsize=9, ncol=2)
    return _save(fig, "f03_kyufu_suii")


def f_dai9_taihi(nen, mikomi):
    """第9期の見込みと実績（標準給付費）。"""
    lab = ["令和6年度", "令和7年度"]
    mi = [mikomi["令和6年度"] / 1e3, mikomi["令和7年度"] / 1e3]
    jt = [nen["令和6年度"]["支払計"] / 1e6, nen["令和7年度"]["支払計"] / 1e6]
    x = range(len(lab))
    w = 0.34
    fig, ax = plt.subplots(figsize=(6.4, 3.2))
    ax.bar([i - w / 2 for i in x], mi, width=w, color=PALE, label="第9期の見込み")
    ax.bar([i + w / 2 for i in x], jt, width=w, color=RED, label="実績")
    for i, v in enumerate(mi):
        _lab(ax, i - w / 2, v + 14, "%.0f" % v, NAVY, ha="center", sz=9)
    for i, (v, m) in enumerate(zip(jt, mi)):
        _lab(ax, i + w / 2, v + 14, "%.0f\n（%.1f％）" % (v, v / m * 100), RED,
             ha="center", sz=9)
    ax.set_xticks(list(x))
    ax.set_xticklabels(lab, fontsize=10)
    ax.set_ylim(0, 1420)
    ax.set_ylabel("標準給付費（百万円）")
    ax.legend(loc="lower right", frameon=False, fontsize=9, ncol=2)
    return _save(fig, "f04_dai9_taihi")


def f_service(act):
    """サービス別給付費の対前年増減率（令和6年度→令和7年度）。"""
    svc = ["短期入所生活介護", "訪問看護", "福祉用具貸与",
           "地域密着型介護老人福祉施設入所者生活介護",
           "居宅介護支援・介護予防支援", "小規模多機能型居宅介護",
           "通所介護", "介護老人福祉施設", "訪問介護",
           "認知症対応型共同生活介護", "介護老人保健施設"]

    def k(y, s):
        v = act[y]["給付費"].get(s, (0, 0, 0))
        return v[0] + v[1]

    rate = [(s, (k("令和7年度", s) / k("令和6年度", s) - 1) * 100)
            for s in svc if k("令和6年度", s)]
    rate.sort(key=lambda r: r[1], reverse=True)
    lab = [s.replace("地域密着型介護老人福祉施設入所者生活介護", "地域密着型\n介護老人福祉施設")
              .replace("居宅介護支援・介護予防支援", "居宅介護支援等")
              .replace("認知症対応型共同生活介護", "認知症対応型\n共同生活介護")
           for s, _v in rate]
    val = [v for _s, v in rate]
    col = [BLUE if v >= 0 else (RED if v <= -15 else GREY) for v in val]
    y = range(len(val))
    fig, ax = plt.subplots(figsize=(7.4, 4.6))
    ax.barh(list(y), val, height=0.6, color=col)
    for i, v in enumerate(val):
        _lab(ax, v + (0.6 if v >= 0 else -0.6), i, _pm(v, "％"), col[i],
             ha="left" if v >= 0 else "right", sz=9)
    ax.axvline(0, color="#808080", lw=0.9)
    ax.set_yticks(list(y))
    ax.set_yticklabels(lab, fontsize=9)
    ax.invert_yaxis()
    ax.set_xlim(-28, 16)
    ax.set_xlabel("給付費の対前年増減率（％）")
    _xgrid(ax)
    return _save(fig, "f05_service")


def f_chiiki(jisseki):
    """地域支援事業費の推移と内訳（積み上げ）。"""
    ys = ["令和3年度", "令和4年度", "令和5年度", "令和6年度"]
    a = [jisseki[y]["介護予防・生活支援サービス事業"] / 1e6 for y in ys]
    b = [jisseki[y]["一般介護予防事業"] / 1e6 for y in ys]
    c = [jisseki[y]["包括的支援事業・任意事業"] / 1e6 for y in ys]
    x = range(len(ys))
    fig, ax = plt.subplots(figsize=(6.8, 3.4))
    ax.bar(list(x), a, width=0.58, color=BLUE, label="介護予防・生活支援サービス事業")
    ax.bar(list(x), b, width=0.58, bottom=a, color=RED, label="一般介護予防事業")
    ax.bar(list(x), c, width=0.58, bottom=[i + j for i, j in zip(a, b)],
           color=PALE, label="包括的支援事業・任意事業")
    for i in x:
        tot = a[i] + b[i] + c[i]
        _lab(ax, i, tot + 1.4, "%.1f" % tot, NAVY, ha="center", sz=9)
    _lab(ax, 3, a[3] + b[3] / 2, "%.1f" % b[3], "#FFFFFF", ha="center", sz=8)
    ax.annotate("", xy=(3, a[3] + b[3]), xytext=(2.45, 52),
                arrowprops=dict(arrowstyle="->", color=RED, lw=1.5))
    _lab(ax, 2.40, 52, "一般介護予防事業費が\n5.3倍に", RED, ha="right", sz=9)
    ax.set_xticks(list(x))
    ax.set_xticklabels(ys, fontsize=9)
    ax.set_ylim(0, 78)
    ax.set_ylabel("地域支援事業費（百万円）")
    ax.legend(loc="upper left", frameon=False, fontsize=8, ncol=1)
    return _save(fig, "f06_chiiki")


def f_riyou():
    """受給率と利用強度の変化（平成26年度→令和7年度）。"""
    lab = ["在宅サービス受給率（％）", "在宅サービス利用率（％）",
           "訪問介護 1人あたり利用回数（回）", "通所介護 1人あたり利用日数（日）",
           "訪問看護 1人あたり利用回数（回）"]
    old = [12.2, 59.3, 10.9, 7.4, 3.1]
    new = [10.3, 44.1, 15.7, 9.3, 6.8]
    y = range(len(lab))
    h = 0.34
    fig, ax = plt.subplots(figsize=(7.2, 3.6))
    ax.barh([i + h / 2 for i in y], old, height=h, color=PALE, label="平成26年度")
    ax.barh([i - h / 2 for i in y], new, height=h,
            color=[RED if n < o else BLUE for n, o in zip(new, old)],
            label="令和7年度")
    for i, v in enumerate(old):
        _lab(ax, v + 0.6, i + h / 2, "%.1f" % v, NAVY, sz=9)
    for i, (v, o) in enumerate(zip(new, old)):
        _lab(ax, v + 0.6, i - h / 2, "%.1f" % v, RED if v < o else BLUE, sz=9)
    ax.set_yticks(list(y))
    ax.set_yticklabels(lab, fontsize=9)
    ax.invert_yaxis()
    ax.set_xlim(0, 72)
    ax.set_xlabel("値（指標ごとに単位が異なる）")
    _xgrid(ax)
    hd, lb = ax.get_legend_handles_labels()
    ax.legend(hd[::-1], lb[::-1], loc="lower right", frameon=False, fontsize=9,
              ncol=2)
    return _save(fig, "f07_riyou")


def f_hokenryo(cases, dai9):
    """保険料のケース（100円未満切上げ後）と第9期との対比。"""
    lab = [nm for nm, _v in cases]
    val = [v for _nm, v in cases]
    col = [RED if v == dai9 else (BLUE if v < dai9 else ORANGE) for v in val]
    y = range(len(val))
    fig, ax = plt.subplots(figsize=(7.4, 3.0))
    ax.barh(list(y), val, height=0.58, color=col)
    for i, v in enumerate(val):
        _lab(ax, v + 180, i, "%s円　%s" % (
            format(v, ","),
            "第9期と同額" if v == dai9 else _pm(v - dai9, "円", 0)), col[i],
             sz=10)
    ax.axvline(dai9, color=ORANGE, lw=1.6, ls="--")
    ax.set_yticks(list(y))
    ax.set_yticklabels(lab, fontsize=10)
    ax.set_ylim(len(val) - 0.4, -1.05)          # 反転と同時に上の余白を作る
    _lab(ax, dai9, -0.72, "第9期 %s円" % format(dai9, ","), ORANGE,
         ha="center", va="bottom", sz=9)
    ax.set_xlim(0, 10200)
    ax.set_xlabel("保険料基準額（月額・円。100円未満切上げ後）")
    _xgrid(ax)
    return _save(fig, "f08_hokenryo")


def f_zaisei(nen):
    """保険料収入と給付費の開き、繰越金の蓄積。"""
    lab = ["平成26年度", "平成30年度", "令和2年度", "令和5年度"]
    kuri = [32.4, 49.9, 155.9, 182.0]
    fig, ax = plt.subplots(figsize=(6.8, 3.2))
    ax.bar(range(len(kuri)), kuri, width=0.56, color=[PALE, PALE, BLUE, RED])
    for i, v in enumerate(kuri):
        _lab(ax, i, v + 5, "%.1f" % v, RED if i == 3 else NAVY, ha="center",
             sz=10)
    ax.set_xticks(range(len(lab)))
    ax.set_xticklabels(lab, fontsize=10)
    ax.set_ylim(0, 218)
    ax.set_ylabel("繰越金（百万円）")
    _lab(ax, 3.0, 200, "保険給付費の16.5％", RED, ha="center", sz=9)
    return _save(fig, "f09_zaisei")


def build_all(act, nen, mikomi, jisseki, cases, dai9):
    """協議会資料で用いる図をすべて作る。"""
    return {
        "認定推移": f_nintei_suii(),
        "認定分解": f_nintei_bunkai(),
        "給付費推移": f_kyufu_suii(nen),
        "第9期対比": f_dai9_taihi(nen, mikomi),
        "サービス別": f_service(act),
        "地域支援": f_chiiki(jisseki),
        "利用強度": f_riyou(),
        "保険料": f_hokenryo(cases, dai9),
        "財政": f_zaisei(nen),
    }


if __name__ == "__main__":
    act, _sub, nen = T.extract()
    MIKOMI = {"令和6年度": 1185894, "令和7年度": 1189712, "令和8年度": 1190880}
    CASES = [("本計画の算定（取崩なし）", 6100),
             ("見える化システムの算定", 6600),
             ("本計画＋基金12,000千円取崩", 6000),
             ("第1号負担割合24％の場合", 6400)]
    for k, v in build_all(act, nen, MIKOMI, T.CHIIKI_JISSEKI, CASES,
                          T.DAI9_KIJUN).items():
        print("  %-8s %s" % (k, v))


# ================================================================ 素案の図

def f_jinko(pop, years):
    """総人口と高齢化率の推移・推計。"""
    tot = [pop[y]["総人口"] for y in years]
    ko = [pop[y]["65+"] / pop[y]["総人口"] * 100 for y in years]
    x = range(len(years))
    fig, ax = plt.subplots(figsize=(7.6, 3.6))
    ax.bar(list(x), tot, width=0.6, color=PALE)
    for i, v in enumerate(tot):
        _lab(ax, i, v + 130, "%s" % format(v, ","), NAVY, ha="center", sz=8)
    ax.set_ylim(0, max(tot) * 1.22)
    ax.set_ylabel("総人口（人）")
    ax.set_xticks(list(x))
    ax.set_xticklabels([y.replace("年度", "") for y in years], fontsize=8)
    ax2 = ax.twinx()
    ax2.plot(list(x), ko, "-o", color=RED, lw=2.4, ms=6)
    for i, v in enumerate(ko):
        last = (i == len(ko) - 1)
        _lab(ax2, i + (-0.08 if last else 0), v - (1.4 if last else -1.1),
             "%.1f％" % v, RED, ha="right" if last else "center", sz=8)
    ax2.set_ylim(min(ko) - 6, max(ko) + 6)
    ax2.set_ylabel("高齢化率（％）", color=RED)
    ax2.tick_params(axis="y", colors=RED)
    _lab(ax2, len(years) - 1.2, min(ko) - 3.6, "棒＝総人口（左軸）／"
         "折れ線＝高齢化率（右軸）", NAVY, ha="right", sz=8)
    return _save(fig, "f21_jinko")


def f_nintei_mikomi(jisseki, mikomi):
    """認定者数の実績と第10期の見込み。"""
    lab = [y for y, _v in jisseki] + [y for y, _v in mikomi]
    val = [v for _y, v in jisseki] + [v for _y, v in mikomi]
    n = len(jisseki)
    col = [BLUE] * n + [ORANGE] * len(mikomi)
    x = range(len(val))
    fig, ax = plt.subplots(figsize=(7.2, 3.3))
    ax.bar(list(x), val, width=0.6, color=col)
    for i, v in enumerate(val):
        _lab(ax, i, v + 14, "%d" % v, col[i], ha="center", sz=9)
    ax.axvline(n - 0.5, color="#808080", lw=1.0, ls="--")
    _lab(ax, n - 0.45, max(val) * 1.12, "← 実績　｜　見込み →", NAVY,
         ha="center", sz=9)
    ax.set_xticks(list(x))
    ax.set_xticklabels([s.replace("年度", "\n年度") for s in lab], fontsize=8)
    ax.set_ylim(0, max(val) * 1.22)
    ax.set_ylabel("第1号被保険者の認定者数（人）")
    return _save(fig, "f22_nintei_mikomi")


def f_kubun(series, years, ylab="総給付費（百万円）", name="f23_kubun"):
    """区分別の積み上げ。series は (区分名, 各年度の値) の並び。"""
    cols = [BLUE, PALE, ORANGE, GREY]
    x = range(len(years))
    fig, ax = plt.subplots(figsize=(6.8, 3.4))
    bottom = [0.0] * len(years)
    for (lab, vals), c in zip(series, cols):
        ax.bar(list(x), vals, width=0.56, bottom=bottom, color=c, label=lab)
        bottom = [b + v for b, v in zip(bottom, vals)]
    for i in x:
        _lab(ax, i, bottom[i] + max(bottom) * 0.025,
             format(round(bottom[i]), ","), NAVY, ha="center", sz=9)
    ax.set_xticks(list(x))
    ax.set_xticklabels(years, fontsize=9)
    ax.set_ylim(0, max(bottom) * 1.24)
    ax.set_ylabel(ylab)
    ax.legend(loc="upper left", frameon=False, fontsize=9,
              ncol=len(series))
    return _save(fig, name)
