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


# 作図した図の数値を控える台帳。build_ono_zuhyo_daicho.py が読む。
# 各図の数値は関数ごとに一次資料から計算しているため、
# ここで書き出す以外に一覧する方法がない。
# 図の作り方には手を入れず、共通の出口（_save）で控えるだけとする。
LEDGER = {}


def _harvest(fig, name):
    """描かれた図から、系列名と数値を読み取って控える。"""
    try:
        ax = fig.axes[0]
    except (AttributeError, IndexError):
        return
    d = {"name": name, "title": ax.get_title() or "",
         "xlabel": ax.get_xlabel() or "", "ylabel": ax.get_ylabel() or "",
         "xs": [t.get_text().replace("\n", " ") for t in ax.get_xticklabels()],
         "ys": [t.get_text().replace("\n", " ") for t in ax.get_yticklabels()],
         "series": [], "kind": ""}
    kinds = set()
    for c in getattr(ax, "containers", []):
        patches = list(getattr(c, "patches", []) or [])
        if not patches:
            continue
        # 幅がそろっていれば縦棒、高さがそろっていれば横棒とみる
        ws = {round(p.get_width(), 6) for p in patches}
        hs = {round(p.get_height(), 6) for p in patches}
        if len(ws) == 1 and len(hs) > 1:
            vals = [p.get_height() for p in patches]
            kinds.add("縦棒")
        elif len(hs) == 1 and len(ws) >= 1:
            vals = [p.get_width() for p in patches]
            kinds.add("横棒")
        else:
            vals = [p.get_height() for p in patches]
            kinds.add("縦棒")
        lab = c.get_label() or ""
        if lab.startswith("_"):
            lab = ""
        d["series"].append((lab, [round(float(v), 4) for v in vals]))
    for ln in getattr(ax, "lines", []):
        yd = list(ln.get_ydata())
        if not yd or len(yd) < 2:
            continue
        lab = ln.get_label() or ""
        if lab.startswith("_"):
            lab = ""
        d["series"].append((lab, [round(float(v), 4) for v in yd]))
        kinds.add("折れ線")
    # 第2軸があれば拾う
    for ax2 in fig.axes[1:]:
        for ln in getattr(ax2, "lines", []):
            yd = list(ln.get_ydata())
            if not yd or len(yd) < 2:
                continue
            lab = (ln.get_label() or "")
            if lab.startswith("_"):
                lab = ""
            d["series"].append((lab + "（第2軸）" if lab else "（第2軸）",
                                [round(float(v), 4) for v in yd]))
            kinds.add("折れ線")
    d["kind"] = "・".join(sorted(kinds)) or "その他"
    LEDGER[name] = d
    _dump()


def _dump():
    """控えをJSONに残す。成果品ごとに別のプロセスで作図するため、
    図表データ管理台帳がまとめて読めるようにする。"""
    import json
    try:
        OUT.mkdir(parents=True, exist_ok=True)
        p = OUT / "_ledger.json"
        got = {}
        if p.exists():
            try:
                got = json.loads(p.read_text())
            except Exception:
                got = {}
        got.update(LEDGER)
        p.write_text(json.dumps(got, ensure_ascii=False, indent=1))
    except Exception:
        pass


def _save(fig, name):
    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / ("%s.png" % name)
    _harvest(fig, name)
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


# ================================================================ 総合事業ワークシートの図

def f_jinzai(years, juyo, kyokyu):
    """人材の視点。認定者数の推計（需要）と、生産年齢人口で縮めた供給可能量。"""
    x = range(len(years))
    fig, ax = plt.subplots(figsize=(7.4, 3.6))
    ax.bar(list(x), juyo, width=0.58, color=PALE, label="認定者数の推計（需要）")
    ax.plot(list(x), kyokyu, "-o", color=RED, lw=2.6, ms=7,
            label="生産年齢人口で縮めた供給可能量")
    for i, v in enumerate(juyo):
        _lab(ax, i, v + 26, "%d" % v, NAVY, ha="center", sz=9)
    for i, v in enumerate(kyokyu):
        if i == 0:                       # 基準年は需要と同じ値なので右にずらす
            _lab(ax, i + 0.34, v, "%d" % v, RED, ha="left", sz=9)
        else:
            _lab(ax, i, v - 64, "%d" % v, RED, ha="center", sz=9)
    gap = [k - j for j, k in zip(juyo, kyokyu)]
    for i, g in enumerate(gap):
        if g < 0:
            _lab(ax, i, 46, _pm(g, "人", 0), RED, ha="center", sz=8)
    ax.set_xticks(list(x))
    ax.set_xticklabels(years, fontsize=9)
    ax.set_ylim(0, max(juyo) * 1.26)
    ax.set_ylabel("要支援・要介護認定者数（人）")
    ax.legend(loc="upper right", frameon=False, fontsize=9)
    return _save(fig, "f31_jinzai")


def f_keido(years, vals, koku=None):
    """調整済み軽度認定率の推移。"""
    x = range(len(years))
    fig, ax = plt.subplots(figsize=(7.2, 3.2))
    ax.plot(list(x), vals, "-o", color=RED, lw=2.6, ms=7, label="小野町")
    for i, v in enumerate(vals):
        _lab(ax, i, v + 0.32, "%.1f％" % v, RED, ha="center", sz=9)
    if koku:
        ax.axhline(koku, color=GREY, lw=1.6, ls="--")
        _lab(ax, len(years) - 0.5, koku - 0.5, "参考 %.1f％" % koku, GREY,
             ha="right", sz=9)
    ax.set_xticks(list(x))
    ax.set_xticklabels([y.replace("年度", "") for y in years], fontsize=9)
    ax.set_ylim(min(vals) - 1.6, max(vals) + 1.4)
    ax.set_ylabel("調整済み軽度認定率（％）")
    return _save(fig, "f32_keido")


def f_chosa(years, series):
    """総合事業の実施状況調査への回答。空欄と、実人数＝延べ人数の年を見せる。

    series は (区分名, 実人数の並び, 延べ人数の並び)。None は調査票の空欄。
    """
    x = range(len(years))
    n = len(series)
    w = 0.8 / (n * 2)
    fig, ax = plt.subplots(figsize=(7.6, 3.6))
    cols = [(BLUE, PALE), (RED, ORANGE)]
    for k, (lab, jitsu, nobe) in enumerate(series):
        for j, (vals, nm, c) in enumerate(
                [(jitsu, "実人数", cols[k][0]), (nobe, "延べ人数", cols[k][1])]):
            off = (k * 2 + j - (n * 2 - 1) / 2) * w
            xs = [i + off for i, v in enumerate(vals) if v is not None]
            ys = [v for v in vals if v is not None]
            ax.bar(xs, ys, width=w * 0.92, color=c, label=f"{lab} {nm}")
            for xx, v in zip(xs, ys):
                _lab(ax, xx, v + 24, "%d" % v, c, ha="center", sz=8)
    for i in x:
        if all(v is None for _l, jt, nb in series for v in (jt[i], nb[i])):
            _lab(ax, i, 40, "調査票が空欄", GREY, ha="center", sz=8)
    ax.set_xticks(list(x))
    ax.set_xticklabels([y.replace("年度", "") for y in years], fontsize=9)
    ax.set_ylim(0, 1260)
    ax.set_ylabel("利用者数（人）")
    ax.legend(loc="upper left", frameon=False, fontsize=8, ncol=2)
    return _save(fig, "f33_chosa")


# ================================================================ 交付金・給付適正化の図

def f_kofukin_moku(rows):
    """交付金の目標別得点。rows は (目標名, 本町, 全国, 県)。"""
    lab = [r[0] for r in rows]
    y = range(len(rows))
    h = 0.26
    fig, ax = plt.subplots(figsize=(7.6, 4.6))
    ax.barh([i + h for i in y], [r[1] for r in rows], height=h,
            color=[RED if r[1] < r[3] else BLUE for r in rows], label="小野町")
    ax.barh(list(y), [r[3] for r in rows], height=h, color=ORANGE,
            label="福島県平均")
    ax.barh([i - h for i in y], [r[2] for r in rows], height=h, color=GREY,
            label="全国平均")
    for i, r in enumerate(rows):
        _lab(ax, r[1] + 0.8, i + h, "%g" % r[1],
             RED if r[1] < r[3] else BLUE, sz=9)
        _lab(ax, r[3] + 0.8, i, "%.1f" % r[3], ORANGE, sz=8)
        _lab(ax, r[2] + 0.8, i - h, "%.1f" % r[2], GREY, sz=8)
    ax.set_yticks(list(y))
    ax.set_yticklabels(lab, fontsize=9)
    ax.invert_yaxis()
    ax.set_xlim(0, max(max(r[1] for r in rows),
                       max(r[2] for r in rows)) * 1.16)
    ax.set_xlabel("得点（各目標100点満点）")
    _xgrid(ax)
    hd, lb = ax.get_legend_handles_labels()
    ax.legend(hd[::-1], lb[::-1], loc="lower right", frameon=False, fontsize=9,
              ncol=3, bbox_to_anchor=(1.0, 1.01))
    return _save(fig, "f41_kofukin_moku")


def f_kofukin_gun(rows):
    """指標群別の到達度。rows は (指標群, 本町, 全国)。"""
    lab = [r[0] for r in rows]
    y = range(len(rows))
    fig, ax = plt.subplots(figsize=(7.2, 2.7))
    ax.barh(list(y), [r[2] for r in rows], height=0.56, color=PALE,
            label="全国平均")
    ax.barh(list(y), [r[1] for r in rows], height=0.34,
            color=[RED if r[1] / r[2] < 0.5 else BLUE for r in rows],
            label="小野町")
    for i, r in enumerate(rows):
        _lab(ax, r[2] + 1.2, i, "全国 %.1f" % r[2], GREY, sz=8)
        _lab(ax, r[1] + 1.2, i - 0.02, "%g（%.0f％）" % (r[1], r[1] / r[2] * 100),
             RED if r[1] / r[2] < 0.5 else BLUE, sz=9)
    ax.set_yticks(list(y))
    ax.set_yticklabels(lab, fontsize=9)
    ax.invert_yaxis()
    ax.set_xlim(0, max(r[2] for r in rows) * 1.42)
    ax.set_xlabel("得点")
    _xgrid(ax)
    return _save(fig, "f42_kofukin_gun")


def f_torikaeshi(rows):
    """0点・低得点の項目と、全国平均まで取り返した場合の伸びしろ。"""
    lab = [r[0] for r in rows]
    y = range(len(rows))
    fig, ax = plt.subplots(figsize=(7.6, 4.4))
    ax.barh(list(y), [r[2] for r in rows], height=0.6, color=PALE,
            label="全国平均")
    ax.barh(list(y), [r[1] for r in rows], height=0.6, color=RED,
            label="小野町")
    for i, r in enumerate(rows):
        _lab(ax, r[2] + 0.35, i, "＋%.1f" % (r[2] - r[1]), NAVY, sz=9)
        if r[1]:
            _lab(ax, r[1] / 2, i, "%g" % r[1], "#FFFFFF", ha="center", sz=8)
    ax.set_yticks(list(y))
    ax.set_yticklabels(lab, fontsize=9)
    ax.invert_yaxis()
    ax.set_xlim(0, max(r[2] for r in rows) * 1.34)
    ax.set_xlabel("得点（赤＝小野町、淡い青＝全国平均。右の数字が伸びしろ）")
    _xgrid(ax)
    return _save(fig, "f43_torikaeshi")


def f_kofukin_tatsu(rows):
    """指標群別の到達度。rows は (指標群, 小野町の得点, 配点, 全国平均)。

    指標群によって配点が違うため、得点そのものではなく
    配点に対する到達度（％）で並べる。
    """
    lab = ["%s\n（配点%d点）" % (r[0], r[2]) for r in rows]
    y = range(len(rows))
    on = [r[1] / r[2] * 100 for r in rows]
    zn = [r[3] / r[2] * 100 for r in rows]
    fig, ax = plt.subplots(figsize=(7.4, 4.6))
    ax.barh(list(y), zn, height=0.62, color=PALE, label="全国平均")
    ax.barh(list(y), on, height=0.36,
            color=[RED if a < b else BLUE for a, b in zip(on, zn)],
            label="小野町")
    for i, (a, b, r) in enumerate(zip(on, zn, rows)):
        _lab(ax, max(a, b) + 1.6, i - 0.12, "%g点（%.0f％）" % (r[1], a),
             RED if a < b else BLUE, sz=9)
        _lab(ax, b + 1.6, i + 0.34, "全国 %.0f％" % b, GREY, sz=8)
    ax.set_yticks(list(y))
    ax.set_yticklabels(lab, fontsize=8.5)
    ax.invert_yaxis()
    ax.set_xlim(0, 128)
    ax.set_xlabel("配点に対する到達度（％）")
    _xgrid(ax)
    ax.legend(loc="lower right", frameon=False, fontsize=9, ncol=2)
    return _save(fig, "f44_kofukin_tatsu")


def f_kofukin_bunpu(ken, ono, ken_hei, zen_hei):
    """県内の保険者の得点の分布と小野町の位置。

    分布（階級ごとの団体数）だけを描き、個別の団体名も個別の得点も出さない。
    """
    lo, hi = 250, 650
    edges = list(range(lo, hi + 1, 50))
    cnt = [sum(1 for v in ken if e <= v < e + 50) for e in edges[:-1]]
    x = [e + 25 for e in edges[:-1]]
    col = [RED if e <= ono < e + 50 else BLUE for e in edges[:-1]]
    fig, ax = plt.subplots(figsize=(7.6, 3.4))
    ax.bar(x, cnt, width=44, color=col)
    for xi, c in zip(x, cnt):
        if c:
            _lab(ax, xi, c + 0.35, "%d" % c, NAVY, ha="center", sz=9)
    ax.axvline(zen_hei, color=GREY, ls="--", lw=1.4)
    ax.axvline(ken_hei, color=ORANGE, ls="--", lw=1.4)
    top = max(cnt) * 1.32
    _lab(ax, zen_hei - 6, top * 0.86, "全国平均 %.1f" % zen_hei, GREY,
         ha="right", sz=8)
    _lab(ax, ken_hei + 6, top * 0.98, "県平均 %.1f" % ken_hei, ORANGE, sz=8)
    _lab(ax, ono, top * 0.60, "小野町 %g点" % ono, RED, ha="center", sz=9.5)
    ax.set_ylim(0, top)
    ax.set_xticks(edges)
    ax.set_yticks(range(0, int(top) + 1, 5))
    ax.set_xlabel("令和8年度交付金の得点（800点満点）")
    ax.set_ylabel("県内の保険者数")
    ax.yaxis.grid(True, color="#E6E6E6", lw=0.8)
    ax.set_axisbelow(True)
    return _save(fig, "f45_kofukin_bunpu")


def f_nyutaiin(rows):
    """入退院支援と人生の最終段階における支援の算定者数割合の位置。

    交付金の活動指標群は、全保険者の中での順位（上位7割・5割・3割・1割）で
    評価される。4つの加算について、小野町がどこまで届いているかを示す。

    rows: [(指標名, 得点, 配点, 到達段階, 補足)]
    """
    names = [r[0] for r in rows]
    pts = [r[1] for r in rows]
    mans = [r[2] for r in rows]
    stage = [r[3] for r in rows]
    y = list(range(len(rows)))
    col = [RED if p == 0 else (BLUE if p >= m else ORANGE)
           for p, m in zip(pts, mans)]
    fig, ax = plt.subplots(figsize=(7.8, 3.2))
    # 配点までの薄い背景
    ax.barh(y, mans, height=0.52, color="#F0F0F0")
    ax.barh(y, pts, height=0.52, color=col)
    for i, (p, m, s) in enumerate(zip(pts, mans, stage)):
        _lab(ax, m + 0.25, i, "%d点／%d点　%s" % (p, m, s),
             col[i], ha="left", sz=9)
    # 段階の目盛り（2点きざみ＝上位7割・5割・3割・1割）
    ax.set_xticks([0, 2, 4, 6, 8])
    ax.set_xticklabels(["0\n（上位7割に\n届かない）", "2\n上位7割", "4\n上位5割",
                        "6\n上位3割", "8\n上位1割"], fontsize=8)
    ax.set_yticks(y)
    ax.set_yticklabels(names, fontsize=9.5)
    ax.invert_yaxis()
    ax.set_xlim(0, 15.5)
    ax.set_xlabel("得点（配点8点）と全保険者の中での位置")
    _xgrid(ax)
    return _save(fig, "f46_nyutaiin")


def f_ninchisho(rows):
    """認知症施策の体制と活動。

    体制（推進員・カフェ）と活動（初期集中支援チームの訪問実績）を
    同じ2時点で並べ、体制だけが伸びていることを示す。

    rows: [(指標名, 単位, 前の時点, 前の値, 後の時点, 後の値, 県内順位, 県内件数)]
    """
    n = len(rows)
    fig, axes = plt.subplots(1, n, figsize=(7.8, 3.0))
    if n == 1:
        axes = [axes]
    for ax, (nm, tani, t0, v0, t1, v1, rk, ken) in zip(axes, rows):
        col = RED if v1 == 0 else (BLUE if v1 > v0 else GREY)
        ax.bar([0, 1], [v0, v1], width=0.56, color=[GREY, col])
        top = max(v0, v1, 1) * 1.55
        for i, v in enumerate((v0, v1)):
            c = NAVY if (i == 0 or v1 != 0) else RED
            ax.text(i, v + top * 0.05, "%g" % v, color=c, ha="center",
                    va="bottom", fontsize=11, fontweight="bold")
        ax.set_xticks([0, 1])
        ax.set_xticklabels([t0, t1], fontsize=8.5)
        ax.set_ylim(0, top)
        ax.set_yticks([])
        ax.set_title("%s\n（%s）" % (nm, tani), fontsize=9.5, pad=6)
        ax.text(0.5, -0.30, "県内%d位／%d" % (rk, ken), transform=ax.transAxes,
                ha="center", va="top", fontsize=8.5, color=GREY)
        for sp in ("top", "right", "left"):
            ax.spines[sp].set_visible(False)
    fig.subplots_adjust(wspace=0.32, bottom=0.22)
    return _save(fig, "f47_ninchisho")


def f_kayoi(ken, chosa):
    """通いの場の実施状況と、調査の回答との食い違い。

    左　見える化システムの箇所数の推移（週1回以上／月1回以上）。
    　　令和3年度以降は同システムにデータがないため、網掛けで示す。
    右　令和7年度調査の参加頻度（n=1,835）。

    ken   {"年": [...], "週": [...], "月": [...]}
    chosa [(区分, ％)] 参加頻度の内訳。合計100％になるもの
    """
    fig, (ax1, ax2) = plt.subplots(
        1, 2, figsize=(7.9, 3.2), gridspec_kw={"width_ratios": [1.25, 1.0]})

    # ---- 左：見える化の箇所数
    ys, wk, mt = ken["年"], ken["週"], ken["月"]
    x = range(len(ys))
    ax1.bar([i - 0.19 for i in x], mt, width=0.36, color=BLUE,
            label="月1回以上")
    ax1.bar([i + 0.19 for i in x], wk, width=0.36, color=RED,
            label="週1回以上")
    top = max(max(mt), max(wk), 1) * 1.45
    for i, v in enumerate(mt):
        if v:
            _lab(ax1, i - 0.19, v + top * 0.03, "%g" % v, BLUE, ha="center", sz=8)
    for i, v in enumerate(wk):
        if v:
            _lab(ax1, i + 0.19, v + top * 0.03, "%g" % v, RED, ha="center", sz=8)
    # 令和3年度以降はデータがない
    ax1.axvspan(len(ys) - 0.5, len(ys) + 1.6, color="#F2F2F2")
    ax1.text(len(ys) + 0.55, top * 0.52, "令和3年度\n以降は\nデータなし",
             color=GREY, ha="center", va="center", fontsize=8.5)
    ax1.set_xlim(-0.7, len(ys) + 1.6)
    ax1.set_ylim(0, top)
    ax1.set_xticks(list(x))
    ax1.set_xticklabels([y.replace("年度", "") for y in ys], fontsize=7.5,
                        rotation=45, ha="right")
    ax1.set_ylabel("箇所数")
    ax1.set_title("見える化システム　通いの場の箇所数", fontsize=9.5, pad=6)
    ax1.legend(fontsize=8, frameon=False, loc="upper left")
    ax1.yaxis.grid(True, color="#E6E6E6", lw=0.8)
    ax1.set_axisbelow(True)

    # ---- 右：調査の参加頻度（横棒を上から並べる）
    cols = [RED, ORANGE, "#FFD966", PALE, GREY]
    names = [nm for nm, _v in chosa]
    vals = [v for _nm, v in chosa]
    y = list(range(len(chosa)))
    ax2.barh(y, vals, height=0.56, color=cols)
    for i, v in enumerate(vals):
        ax2.text(v + 1.6, i, "%.1f％" % v, color=NAVY, ha="left",
                 va="center", fontsize=9, fontweight="bold")
    ax2.set_yticks(y)
    ax2.set_yticklabels(names, fontsize=9)
    ax2.invert_yaxis()
    ax2.set_xlim(0, 100)
    ax2.set_xticks([0, 50, 100])
    ax2.set_xticklabels(["0", "50", "100％"], fontsize=8)
    ax2.set_title("令和7年度調査　通いの場への参加頻度\n（n=1,835）",
                  fontsize=9.5, pad=6)
    _xgrid(ax2)
    for sp in ("top", "right", "left"):
        ax2.spines[sp].set_visible(False)

    fig.subplots_adjust(wspace=0.46, bottom=0.26)
    return _save(fig, "f48_kayoi")
