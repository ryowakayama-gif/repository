# -*- coding: utf-8 -*-
"""
北塩原村　計画書の図（モノクロ印刷に耐える作り）

出力: output/図表/*.png

なぜ作るか
  正本（他メンバー版）には図が19点あるが、給付費・サービス別の構成・
  高齢化率には図がない。数値の推移と構成は図の方が早く読める。

色について（令和8年10月10日の指示）
  図は他メンバー版に合わせてカラーで作る。
  印刷の扱い（仕様書5はモノクロ・コピー・くるみ製本）は確認事項として
  村・他メンバーに残す（Ｍ-69・Ｓ-20）。
  ただしカラーのままでも、同じ図の中の系列はグレー値で25以上離す。
  計画書はコピーして配られることがあり、
  他メンバー版の図では緑 RGB(46,158,107) と 青 RGB(61,134,198) が
  いずれもグレー値119で、白黒コピーすると見分けられなくなっている。
  同じことが起きないようにするための歯止めである。

作りの決め
  ・色は他メンバー版と同じ3色（緑・橙・青）から使う。
  ・同じ図の中で系列が2つ以上あるときは、グレー値が25以上離れる組を選ぶ
    （緑と青は同じ濃さなので一緒に使わない。緑と橙なら52離れる）。
  ・系列が接する積み上げでは、色に加えて45度・135度の網かけを足す。
  ・枠線は下と左だけ。横の目盛線は細い実線（破線にしない）。
  ・棒は間隔を空けて細く。棒どうしが接する積み上げは白の細い境で分ける。
  ・値のラベルは全部には付けない。表が隣にある図は最初と最後の年度だけ。
  ・1系列の図に凡例は置かない（表題が何の図かを示す）。
  ・文字はデータの色を着ない（黒系で統一する）。
  ・二軸は作らない。単位の違うものは図を分ける。

データの出どころ
  給付費の推移・サービス別の構成
    build_kitashiobara_ishoku_okurijo.py の移植20の表から読む
    （申し送りと計画本文と図が、同じ数値から作られるようにするため）
  高齢化率
    正本（北塩原村_計画素案_正本_移植後.docx）第6章1（1）の表から読む
"""

import os
import re
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib import font_manager  # noqa: E402

import docx  # noqa: E402
from docx.table import Table  # noqa: E402
from docx.text.paragraph import Paragraph  # noqa: E402

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_kitashiobara_ishoku_okurijo import ISHOKU  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = f"{REPO_ROOT}/output/図表"
HONPON = f"{REPO_ROOT}/output/北塩原村_計画素案_正本_移植後.docx"

# ---------------------------------------------------------------------------
# モノクロの作り
# ---------------------------------------------------------------------------
GRAY_MIN = 25          # 256階調で、これ以上の差があれば見分けられるとみなす
# 他メンバー版が使っている3色。括弧内はグレー値
#   緑 119 ／ 橙 171 ／ 青 119
#   緑と青は同じ濃さになるため、同じ図の中で並べて使わない
COLORS = {"緑": "#2E9E6B", "橙": "#F0A030", "青": "#3D86C6"}
SERIES = ["緑", "橙"]          # 2系列までの既定（グレー値の差52）
HATCH = ["", "///", "xxx"]     # 接する面には網かけも足す
INK = "#1A1A1A"                # 文字
GRID = "#C8C8C8"               # 目盛線
AXIS = "#4D4D4D"               # 軸

BODY_W_CM = 16.6               # 正本の本文幅
DPI = 200

JP_FONT = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"


def setup_font():
    """日本語の字が出る状態にする（出ないと豆腐になる）。"""
    if not os.path.exists(JP_FONT):
        raise SystemExit(f"日本語フォントが見つかりません: {JP_FONT}")
    font_manager.fontManager.addfont(JP_FONT)
    name = font_manager.FontProperties(fname=JP_FONT).get_name()
    plt.rcParams["font.family"] = name
    plt.rcParams["axes.unicode_minus"] = False
    return name


def hex2rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def gray255(h):
    """色（#RRGGBB）をグレーにしたときの明るさ（0が黒・255が白）。"""
    r, g, b = hex2rgb(h)
    return round(0.299 * r + 0.587 * g + 0.114 * b)


def comma(ax, axis="y"):
    """目盛の数に桁区切りを入れる。"""
    from matplotlib.ticker import FuncFormatter
    f = FuncFormatter(lambda v, _p: f"{v:,.0f}")
    (ax.yaxis if axis == "y" else ax.xaxis).set_major_formatter(f)


def style_axes(ax, ylabel=None):
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for s in ("left", "bottom"):
        ax.spines[s].set_color(AXIS)
        ax.spines[s].set_linewidth(0.8)
    ax.tick_params(colors=INK, labelsize=10, length=3, width=0.8)
    ax.yaxis.grid(True, color=GRID, linewidth=0.6, linestyle="-")
    ax.set_axisbelow(True)
    if ylabel:
        # 単位は縦書きにすると読みにくいため、軸の上に横書きで置く
        ax.annotate(ylabel, xy=(0, 1), xycoords="axes fraction",
                    xytext=(-6, 10), textcoords="offset points",
                    ha="right", va="bottom", fontsize=9.5, color=INK)


def save(fig, name):
    os.makedirs(OUT_DIR, exist_ok=True)
    path = f"{OUT_DIR}/{name}.png"
    fig.savefig(path, dpi=DPI, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


# ---------------------------------------------------------------------------
# データの読み取り
# ---------------------------------------------------------------------------
def ishoku_tables(no):
    rec = next(r for r in ISHOKU if r["no"] == no)
    return [b[1] for b in rec["nakami"] if b[0] == "tbl"]


def to_num(s):
    """「36,002千円」「35.5％」「＋34.0％」から数を取り出す。"""
    s = str(s).replace(",", "").replace("＋", "").replace("△", "-")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group()) if m else None


def read_kyufu():
    """移植20（1）給付費の推移。単位は表と同じ千円のままにする。"""
    tbl = ishoku_tables("20")[0]
    years, kaigo, jido = [], [], []
    for row in tbl[1:]:
        years.append(row[0].replace("年度", ""))   # 「令和２」まで残す
        kaigo.append(to_num(row[1]))
        jido.append(to_num(row[2]))
    return years, kaigo, jido


def read_service():
    """移植20（3）サービス別の構成。単位は表と同じ千円。"""
    tbl = ishoku_tables("20")[2]
    names, vals, pcts = [], [], []
    for row in tbl[1:]:
        names.append(row[0])
        vals.append(to_num(row[1]))
        m = re.search(r"（([\d.]+)％）", row[1])
        pcts.append(float(m.group(1)) if m else None)
    return names, vals, pcts


def read_koureika():
    """正本 第6章1（1）高齢化率の表。"""
    doc = docx.Document(HONPON)
    for t in doc.tables:
        hdr = [c.text.strip() for c in t.rows[0].cells]
        if hdr[:2] == ["年度", "高齢化率"] and len(hdr) == 4:
            out = []
            for row in t.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                out.append((cells[0], to_num(cells[1])))
                if len(cells) > 3 and cells[2]:
                    out.append((cells[2], to_num(cells[3])))
            # 表は「年度｜高齢化率｜年度｜高齢化率」の2組並びであるため、
            # 行の順に読むと令和3・6・4・7…と入れ違う。年で並べ直す。
            out = [(a.replace("年度", ""), b) for a, b in out if b is not None]
            def nen(label):
                z = str.maketrans("０１２３４５６７８９", "0123456789")
                m = re.search(r"\d+", label.replace("令和", "").translate(z))
                return int(m.group()) if m else 0
            out.sort(key=lambda d: nen(d[0]))
            return out
    raise LookupError("高齢化率の表が見つかりません")


# ---------------------------------------------------------------------------
# 図
# ---------------------------------------------------------------------------
def fig_kyufu():
    """給付費の推移（小さな図を2つ並べる）。

    介護給付費等と障害児給付費は桁が15倍ほど違う。
    積み上げると障害児給付費が見えず、二軸にすると
    ありもしない相関を作ってしまうため、図を分ける。
    どちらも1系列なので凡例は置かない。
    """
    years, kaigo, jido = read_kyufu()
    labels = [y.replace("令和", "") for y in years]
    fig, axes = plt.subplots(
        1, 2, figsize=(BODY_W_CM * 0.94 / 2.54, 6.8 / 2.54))
    for ax, vals, title, marks in (
            (axes[0], kaigo, "介護給付費等", (0, len(kaigo) - 1)),
            (axes[1], jido, "障害児給付費",
             (jido.index(max(jido)), len(jido) - 1))):
        x = range(len(labels))
        ax.bar(x, vals, width=0.5, color=COLORS[SERIES[0]],
               edgecolor="none")
        style_axes(ax, "（千円）" if ax is axes[0] else None)
        comma(ax)
        ax.set_xticks(list(x))
        ax.set_xticklabels(labels, fontsize=9.5)
        ax.set_title(title, fontsize=10.5, color=INK, pad=14)
        ax.margins(x=0.14)
        # ラベルは要点だけ（左は最初と最後、右は山と最後）
        for i in marks:
            ax.annotate(f"{vals[i]:,.0f}", (i, vals[i]),
                        textcoords="offset points", xytext=(0, 5),
                        ha="center", fontsize=9.5, color=INK)
        ax.set_ylim(0, max(vals) * 1.22)
        ax.tick_params(axis="y", labelsize=9)
    axes[0].annotate("（各年度・令和２年度から令和７年度まで）",
                     xy=(0, 0), xycoords="axes fraction",
                     xytext=(0, -34), textcoords="offset points",
                     fontsize=9, color=INK)
    fig.subplots_adjust(wspace=0.35)
    return save(fig, "01_給付費の推移"), years, kaigo, jido


def fig_service():
    """サービス別の構成（横棒・1系列）。

    区分が8つで名前が長いため横棒にする。
    1系列なので凡例は置かず、濃さも1段で足りる。
    """
    names, vals, pcts = read_service()
    order = sorted(range(len(names)), key=lambda i: vals[i])
    fig, ax = plt.subplots(
        figsize=(BODY_W_CM * 0.86 / 2.54, 8.4 / 2.54))
    y = range(len(order))
    ax.barh(y, [vals[i] for i in order], height=0.56,
            color=COLORS[SERIES[0]], edgecolor="none")
    ax.set_yticks(list(y))
    ax.set_yticklabels([names[i] for i in order], fontsize=10)
    style_axes(ax)
    ax.yaxis.grid(False)
    ax.xaxis.grid(True, color=GRID, linewidth=0.6, linestyle="-")
    ax.set_xlabel("（千円）", fontsize=9.5, color=INK, labelpad=6)
    comma(ax, "x")
    for k, i in enumerate(order):
        lab = f"{vals[i]:,.0f}"
        if pcts[i] is not None:
            lab += f"（{pcts[i]}％）"
        ax.annotate(lab, (vals[i], k), textcoords="offset points",
                    xytext=(5, 0), va="center", fontsize=9.5, color=INK)
    ax.set_xlim(0, max(vals) * 1.34)
    ax.margins(y=0.04)
    return save(fig, "02_サービス別の給付費"), names, vals, pcts


def fig_koureika():
    """高齢化率の推移（折れ線・1系列）。"""
    data = read_koureika()
    xs = [d[0] for d in data]
    ys = [d[1] for d in data]
    fig, ax = plt.subplots(
        figsize=(BODY_W_CM * 0.9 / 2.54, 6.6 / 2.54))
    ax.plot(range(len(xs)), ys, color=COLORS[SERIES[0]], linewidth=2,
            marker="o", markersize=5, markerfacecolor=COLORS[SERIES[0]],
            markeredgecolor="white", markeredgewidth=1.5)
    style_axes(ax, "（％）")
    ax.set_xticks(range(len(xs)))
    ax.set_xticklabels([x.replace("令和", "") for x in xs],
                       fontsize=10)
    # 右上がりの線では、最初のラベルを下に、最後を上に置くと重ならない
    ax.annotate(f"{ys[0]}％", (0, ys[0]), textcoords="offset points",
                xytext=(0, -17), ha="center", fontsize=10, color=INK)
    ax.annotate(f"{ys[-1]}％", (len(xs) - 1, ys[-1]),
                textcoords="offset points", xytext=(0, 9), ha="center",
                fontsize=10, color=INK)
    ax.margins(x=0.1)
    lo, hi = min(ys), max(ys)
    ax.set_ylim(lo - (hi - lo) * 0.6, hi + (hi - lo) * 0.5)
    ax.annotate(f"（{xs[0]}年度から{xs[-1]}年度まで・各年４月１日現在）",
                xy=(0, 0), xycoords="axes fraction",
                xytext=(0, -34), textcoords="offset points",
                fontsize=9, color=INK)
    return save(fig, "03_高齢化率の推移"), xs, ys


# ---------------------------------------------------------------------------
# 正本に入れる図の仕様（build_kitashiobara_honpon.py が読む）
#   名前, 表題, 出所, 挿入位置（この段落の次に現れる表の直後に入れる）
# ---------------------------------------------------------------------------
GRAPHS = [
    dict(no="Z-1", file="01_給付費の推移.png",
         title="障がい福祉サービス等に係る給付費の推移",
         src="資料：村提供の障がいサービス給付実績",
         ichi="（１）給付費の推移"),
    dict(no="Z-2", file="02_サービス別の給付費.png",
         title="サービス別の給付費（令和７年度）",
         src="資料：村提供の障がいサービス給付実績",
         ichi="（３）サービス別の構成（令和７年度）"),
    dict(no="Z-3", file="03_高齢化率の推移.png",
         title="高齢化率の推移",
         src="資料：北塩原村（各年４月１日現在）",
         ichi="（１）高齢化率"),
]


# ---------------------------------------------------------------------------
# 自己点検
# ---------------------------------------------------------------------------
def verify(made):
    from PIL import Image
    import collections
    ng = []

    # ① 同じ図で並べて使う色が、グレーにしても25以上離れていること
    #    （白黒コピーされても系列を見分けられるようにするため）
    g = [gray255(COLORS[k]) for k in SERIES]
    for i in range(len(g)):
        for j in range(i + 1, len(g)):
            if abs(g[i] - g[j]) < GRAY_MIN:
                ng.append(f"並べて使う色の濃さが近すぎる: "
                          f"{SERIES[i]}{g[i]} と {SERIES[j]}{g[j]}")
    #    緑と青を同じ図で使っていないこと（グレー値が同じ）
    if abs(gray255(COLORS["緑"]) - gray255(COLORS["青"])) >= GRAY_MIN:
        ng.append("緑と青のグレー値が離れている（前提が変わっている）")
    if "緑" in SERIES and "青" in SERIES:
        ng.append("緑と青を同じ図で並べて使っている（白黒で潰れる）")

    # ② 図に使われている色が、決めた3色の色合いに収まっていること
    #    （縁のぼかしは白と混ざって薄くなるだけで、色合いは変わらない。
    #     濃さではなく色合い（色相）で見る）
    import colorsys

    def hue(rgb):
        r, g2, b = (v / 255 for v in rgb)
        return colorsys.rgb_to_hsv(r, g2, b)[0] * 360

    hues = [hue(hex2rgb(v)) for v in COLORS.values()]
    for path in made:
        im = Image.open(path).convert("RGB")
        small = im.resize((min(im.width, 300), min(im.height, 300)))
        try:
            px = list(small.get_flattened_data())
        except AttributeError:
            px = list(small.getdata())
        for rgb, n in collections.Counter(px).most_common(40):
            if n < 150 or max(rgb) - min(rgb) <= 14:
                continue      # 白地・黒文字・グレーの目盛線は見ない
            h = hue(rgb)
            if not any(min(abs(h - q), 360 - abs(h - q)) <= 12 for q in hues):
                ng.append(f"決めた色以外が使われている: "
                          f"{os.path.basename(path)} {rgb}（色相{h:.0f}度）")
                break

    # ③ 図の大きさが本文の幅に収まること
    for path in made:
        im = Image.open(path)
        cm = im.width / DPI * 2.54
        if cm > BODY_W_CM + 0.6:
            ng.append(f"本文の幅を超えている: {os.path.basename(path)} "
                      f"{cm:.1f}cm")

    # ④ 図の仕様と作った図が対応すること
    names = {os.path.basename(p) for p in made}
    for g2 in GRAPHS:
        if g2["file"] not in names:
            ng.append(f"仕様にある図が作られていない: {g2['file']}")

    # ⑤ 高齢化率の年が昇順であること（表が2組並びで入れ違いやすい）
    data = read_koureika()
    z = str.maketrans("０１２３４５６７８９", "0123456789")
    nen = [int(re.search(r"\d+", d[0].replace("令和", "").translate(z)).group())
           for d in data]
    if nen != sorted(nen):
        ng.append(f"高齢化率の年が昇順でない: {nen}")

    # ⑥ 読み取った数が原典と合うこと（移植20の表の行数）
    if len(ishoku_tables("20")[0]) - 1 != 6:
        ng.append("給付費の推移の年度数が6でない")
    if len(ishoku_tables("20")[2]) - 1 != 8:
        ng.append("サービス別の構成の区分数が8でない")

    if ng:
        print("自己点検 不合格:")
        for e in ng:
            print("   -", e)
        raise SystemExit(1)
    print(f"  自己点検: 並べて使う色 {SERIES}＝グレー値{g}"
          f"（差が{GRAY_MIN}以上）／"
          f"作った図{len(made)}点が決めた3色の範囲／"
          f"本文幅{BODY_W_CM}cmに収まる／高齢化率の年が昇順／原典の行数と一致")


def main():
    name = setup_font()
    p1, years, kaigo, jido = fig_kyufu()
    p2, svc, vals, pcts = fig_service()
    p3, xs, ys = fig_koureika()
    made = [p1, p2, p3]
    verify(made)
    print(f"作成: {OUT_DIR}")
    print(f"  フォント: {name}")
    for p in made:
        from PIL import Image
        im = Image.open(p)
        print(f"  {os.path.basename(p)}  {im.width}×{im.height}px "
              f"（{im.width / DPI * 2.54:.1f}cm）")
    print(f"  給付費 {years[0]}〜{years[-1]}年度／"
          f"サービス{len(svc)}区分／高齢化率{xs[0]}〜{xs[-1]}年度")


if __name__ == "__main__":
    main()
