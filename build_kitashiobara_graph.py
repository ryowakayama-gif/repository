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
# 箱の地色。同じ色みの薄い色を2つ作ると、白黒にしたとき見分けられない
# （例 #FDF3E4 と #FDF4E6 はグレー値が1しか違わない）。色ごとに1つに決める。
USUIRO = {"緑": "#F4FAF6", "橙": "#FDF3E4"}
KOIME = {"緑": "#EAF4EE"}
HAI = "#F7F7F7"

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


def save(fig, name, tight=True):
    """図を保存する。

    関係を示す図（体系・連携など）は、箱の位置を0〜1で決め打ちしている。
    bbox_inches="tight" を使うと端の文字の分だけ外へ広がり、
    本文の幅を超えてしまうため、図の大きさをそのまま書き出す。
    """
    os.makedirs(OUT_DIR, exist_ok=True)
    path = f"{OUT_DIR}/{name}.png"
    fig.savefig(path, dpi=DPI, facecolor="white",
                **({"bbox_inches": "tight"} if tight else {}))
    plt.close(fig)
    return path


# ---------------------------------------------------------------------------
# データの読み取り
# ---------------------------------------------------------------------------
def tbls_juten(no):
    from build_kitashiobara_juten12 import JUTEN_EDITS
    rec = next(r for r in JUTEN_EDITS if r["no"] == no)
    return [b[1] for b in rec["nakami"] if b[0] == "tbl"]


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
# 関係を示す図（体系・連携・機能・場面・流れ）の部品
#   箱と線だけの飾りにしないため、次を守る。
#   ・箱の中は正本の表の文言をそのまま使い、図のために言い換えない
#   ・裏づけのない対応関係は線で結ばない（結べないことは図の中に書く）
#   ・色は塗り分けの意味がある場合だけ使い、意味がなければ白地に線
#   ・文字は幅を測って折り返す。入らない字は切らずに箱を高くする
# ---------------------------------------------------------------------------
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch  # noqa: E402

PAD = 0.010          # 箱の内側の余白（横。図の幅に対する割合）
LINE = 1.38          # 行送り


class Sheet:
    """図を組む台。位置は 0〜1 で指定する。

    文字の大きさ（ポイント）と図の大きさ（センチ）が分かるため、
    1行に何文字入るかを計算できる。日本語は1文字がほぼ全角1つ分なので、
    文字数で折り返せば幅に収まる。
    """

    def __init__(self, h_cm, w_cm=None):
        self.w_cm = w_cm or BODY_W_CM
        self.h_cm = h_cm
        self.fig, self.ax = plt.subplots(
            figsize=(self.w_cm / 2.54, h_cm / 2.54))
        self.ax.set_xlim(0, 1)
        self.ax.set_ylim(0, 1)
        self.ax.axis("off")
        self.fig.subplots_adjust(left=0, right=1, top=1, bottom=0)
        self.low = 1.0       # 中身の下端（箱から拾う）
        self.notes_h = 0.0   # 注の高さ

    # --- 寸法 ---------------------------------------------------------
    def pt_w(self, frac):
        """幅（割合）をポイントに直す。"""
        return frac * self.w_cm / 2.54 * 72

    def chars(self, w, fs):
        """幅 w の中に、大きさ fs の字が1行に何文字入るか。"""
        # 0.93 としているのは、行末に送れない字（、。）」）を1字だけ
        # ぶら下げても箱の線に当たらないようにするため。
        return max(4, int(self.pt_w(w) * 0.93 / fs))

    def wrap(self, text, w, fs):
        """幅に合わせて折り返す（改行はそのまま生かす）。"""
        n = self.chars(w, fs)
        # 行頭に置けない字をぶら下げられるのは、線に当たらない範囲まで。
        yoyuu = max(1, int(self.pt_w(w) / fs) - n)
        out = []
        for para in str(text).split("\n"):
            # 半角の英数字が続くところは割らない。「（11）」が
            # 「（1」と「1）」に割れると数が読めなくなる。
            toks = re.findall(r"[0-9A-Za-z.,%]+|.", para)
            line = ""
            for tk in toks:
                nobetu = tk in "、。）」％ー" and len(line) < n + yoyuu
                if line and len(line) + len(tk) > n and not nobetu:
                    # 開き括弧を行末に残さない（次の行へ送る）
                    okuri = ""
                    while line and line[-1] in "（「【":
                        okuri = line[-1] + okuri
                        line = line[:-1]
                    out.append(line)
                    line = okuri
                line += tk
            out.append(line)
        return "\n".join(out)

    def h_of(self, text, w, fs, pad_lines=0.9):
        """折り返した文字の高さ（割合）。"""
        lines = self.wrap(text, w, fs).count("\n") + 1
        return (lines + pad_lines) * fs * LINE / 72 * 2.54 / self.h_cm

    # --- 部品 ---------------------------------------------------------
    def box(self, x, y, w, h, fc="white", ec=None, lw=0.9):
        self.low = min(self.low, y)
        self.ax.add_patch(FancyBboxPatch(
            (x, y), w, h, boxstyle="round,pad=0.003,rounding_size=0.010",
            linewidth=lw, edgecolor=(ec or AXIS), facecolor=fc, zorder=2))

    def text(self, x, y, s, fs=9.5, w=None, ha="left", va="top",
             color=None):
        """w を渡すと幅に合わせて折り返す。"""
        self.ax.text(x, y, (self.wrap(s, w, fs) if w else s),
                     ha=ha, va=va, fontsize=fs, color=(color or INK),
                     zorder=3, linespacing=LINE)

    def panel(self, x, y, w, body, fs=9.0, head=None, head_fs=9.8,
              fc="white", ec=None, top=None):
        """見出しと本文を入れた箱。高さは中身から決める。上端 top も指定可。

        戻り値は (下端, 高さ)。
        """
        inner = w - 2 * PAD
        h = 0.0
        if head:
            h += self.h_of(head, inner, head_fs, pad_lines=0.55)
        if body:
            h += self.h_of(body, inner, fs, pad_lines=0.7)
        if top is not None:
            y = top - h
        self.box(x, y, w, h, fc=fc, ec=ec)
        cur = y + h - 0.012
        if head:
            self.text(x + w / 2, cur, head, fs=head_fs, w=inner,
                      ha="center")
            cur -= self.h_of(head, inner, head_fs, pad_lines=0.0)
        if body:
            self.text(x + PAD, cur, body, fs=fs, w=inner)
        return y, h

    def arrow(self, x1, y1, x2, y2, style="-|>", lw=1.0, color=None):
        self.ax.add_patch(FancyArrowPatch(
            (x1, y1), (x2, y2), arrowstyle=style, mutation_scale=11,
            linewidth=lw, color=(color or AXIS), zorder=1,
            shrinkA=0, shrinkB=0))

    def notes(self, lines, fs=8.4):
        """図の下に並べる注。下から積む。"""
        y = 0.012
        for t in reversed(lines):
            h = self.h_of(t, 1 - 2 * PAD, fs, pad_lines=0.1)
            self.text(PAD, y, t, fs=fs, w=1 - 2 * PAD, va="bottom")
            y += h
        self.notes_h = y

    def save(self, name):
        return save(self.fig, name, tight=False)


def fit(tsukuru, h0, aki_cm=0.32):
    """高さを中身に合わせる。

    箱の位置も字の大きさも、割合ではなくセンチで決まる（h_of が
    高さで割っている）。つまり台を低くしても中身は縮まず、
    下の余白だけが減る。1度組んで中身の下端を測り、同じ手順で
    ちょうどの高さに組み直す。
    """
    shi = tsukuru(h0)
    iru = (1 - shi.low) * h0 + shi.notes_h * h0 + aki_cm
    plt.close(shi.fig)
    if iru >= h0:                      # 足りなければ広げて測り直す
        shi = tsukuru(iru + 2.0)
        iru = (1 - shi.low) * (iru + 2.0) + shi.notes_h * (iru + 2.0) + aki_cm
        plt.close(shi.fig)
    return tsukuru(round(iru, 2))


# ---------------------------------------------------------------------------
# 図4 計画の体系
# ---------------------------------------------------------------------------
RINEN = ("障がいのあるなしに関わらず、お互いの人格や個性を尊重し、"
         "多様な価値観を認め合い、誰もが自分らしく輝くむら")
MOKUHYO = ["①障がいへの理解を深め、交流を育むむら",
           "②共に支え合い、誰もが安心して暮らせるむら",
           "③みんなが輝き、自立した生活を送れるむら"]
SHISAKU = ["①啓発・広報", "②保健・医療", "③福祉", "④教育・育成",
           "⑤雇用・就業", "⑥生活環境", "⑦スポーツ・文化"]
SEIKA8 = ["①施設入所者の地域生活への移行",
          "②精神障害にも対応した地域包括ケアシステムの構築",
          "③福祉施設から一般就労への移行等",
          "④障がい児支援の提供体制の整備等",
          "⑤地域生活支援の充実",
          "⑥相談支援体制の充実・強化等",
          "⑦障害福祉人材の確保・定着、生産性の向上",
          "⑧サービスの質を向上させる体制の構築"]


def _taikei(h_cm):
    """計画の体系（理念→基本目標→基本施策／本計画→成年後見）。

    基本目標と基本施策の対応は第4次障がい者計画の体系図にしかなく、
    当方は未入手である。線で結ばず、図の中にその旨を書く。
    """
    sh = Sheet(h_cm)
    top = 0.975
    y, h = sh.panel(0.02, 0, 0.96, None, head="基本理念　" + RINEN,
                    head_fs=10.2, fc=KOIME["緑"], ec=COLORS["緑"], top=top)
    sh.arrow(0.5, y, 0.5, y - 0.030)
    top = y - 0.030

    sh.text(0.02, top, "基本目標（第４次北塩原村障がい者計画）", fs=8.8,
            va="top")
    top -= 0.035
    hs = [sh.h_of(m, 0.307 - 2 * PAD, 9.4, pad_lines=0.9) for m in MOKUHYO]
    hm = max(hs)
    for i, m in enumerate(MOKUHYO):
        x = 0.02 + i * 0.327
        sh.box(x, top - hm, 0.307, hm, fc=USUIRO["緑"], ec=COLORS["緑"])
        sh.text(x + 0.307 / 2, top - hm / 2 + 0.018, m, fs=9.4,
                w=0.307 - 2 * PAD, ha="center", va="center")
    # 目標ごとに矢印を下ろすと、どの目標がどの施策につながるかを
    # 示したことになってしまう。対応は未確認なので、束ねて2本だけ下ろす。
    for x in (0.2525, 0.7475):
        sh.arrow(x, top - hm, x, top - hm - 0.028)
    top = top - hm - 0.028

    body_l = "\n".join(SHISAKU)
    body_r = "\n".join(SEIKA8)
    yl, hl = sh.panel(0.02, 0, 0.465, body_l, fs=9.0,
                      head="第４次北塩原村障がい者計画の基本施策",
                      head_fs=9.4, ec=COLORS["緑"], top=top)
    yr, hr = sh.panel(0.515, 0, 0.465, body_r, fs=8.4,
                      head="本計画の成果目標（第４章２）",
                      head_fs=9.4, ec=COLORS["橙"], top=top)
    ylow = min(yl, yr)
    mid = (top + ylow) / 2
    sh.arrow(0.485, mid, 0.515, mid, style="<|-|>", lw=0.9)

    sh.arrow(0.5, ylow, 0.5, ylow - 0.028)
    sh.panel(0.02, 0, 0.96,
             "成年後見制度の利用の促進に関する施策についての基本的な計画"
             "（成年後見制度の利用促進に関する法律第14条第１項）　"
             "計画期間：令和９年度～令和13年度（第６章）",
             fs=9.0, fc=USUIRO["橙"], ec=COLORS["橙"], top=ylow - 0.028)

    sh.notes([
        "※ 緑は第４次北塩原村障がい者計画の内容、橙は本計画の内容です。",
        "※ 基本目標と基本施策の対応関係は、第４次北塩原村障がい者計画の"
        "体系図によります。当方は同図を未入手のため、本図では線で結んで"
        "いません。村にご確認ください。",
        "※ 本計画は第４次障がい者計画の基本理念・基本目標を踏まえ、"
        "成年後見の計画を内包して一体的に策定するものです"
        "（見込量は第５章）。",
    ])
    return sh


def fig_taikei():
    return fit(_taikei, 12.0).save("04_計画の体系")


# ---------------------------------------------------------------------------
# 図5 関係機関との連携
# ---------------------------------------------------------------------------
def _renkei(h_cm):
    """本人・家族を中心に、分野ごとの関係機関を置く（移植19の表から）。"""
    rows = ishoku_tables("19")[0][1:]
    bunya = {}
    for b, kikan, _naiyo in rows:
        bunya.setdefault(b, []).append(kikan)
    left = ["就労", "教育", "保健・医療"]
    right = ["保育・子育て", "相談支援", "権利擁護", "防災"]

    sh = Sheet(h_cm)
    w = 0.305
    xl, xr = 0.015, 0.68
    top = 0.975
    ys = {}
    for side, keys in ((xl, left), (xr, right)):
        t = top
        for b in keys:
            body = "\n".join("・" + k for k in bunya[b])
            y, h = sh.panel(side, 0, w, body, fs=8.0, head=b, head_fs=9.4,
                            top=t)
            ys[b] = y + h / 2
            t = y - 0.022

    cx, cw = 0.345, 0.31
    y1, h1 = sh.panel(cx, 0, cw, None, head="障がいのある方と家族",
                      head_fs=10.4, fc=KOIME["緑"], ec=COLORS["緑"],
                      top=0.70)
    y2, h2 = sh.panel(cx, 0, cw,
                      "相談の入口。圏域の相談支援事業所と連携して"
                      "サービス等利用計画を作る",
                      fs=8.4, head="村保健福祉課", head_fs=9.4,
                      ec=COLORS["緑"], top=y1 - 0.030)
    sh.arrow(0.5, y1, 0.5, y1 - 0.030, style="<|-|>")
    # 村保健福祉課が相談の入口であり、分野を問わずここが窓口になる。
    # 分野ごとに本人側・村側へ振り分けると、未確認の区別を示すことになる。
    chushin = y2 + h2 / 2
    for b, yy in ys.items():
        x1 = (xl + w) if b in left else xr
        x2 = cx if b in left else (cx + cw)
        sh.arrow(x1, yy, x2, chushin, style="<|-|>", lw=0.8)

    sh.notes([
        "※ 国の基本指針 別表第二 五は、村の障がい保健福祉部局と、"
        "医療機関、教育機関、公共職業安定所、障害者職業センター、"
        "障害者就業・生活支援センターその他の関係機関との連携方法等を"
        "定めることを記載事項としています（第５章６）。",
        "※ 本村は人口2,316人（令和８年４月１日現在）であり、"
        "村内で完結する支援は限られます。会津北部圏域（猪苗代町・磐梯町・"
        "湯川村・本村）及び会津障がい保健福祉圏域の関係機関との連携が、"
        "サービス提供体制の確保そのものになります。",
    ])
    return sh


def fig_renkei():
    return fit(_renkei, 10.4).save("05_関係機関との連携")


# ---------------------------------------------------------------------------
# 図6 地域生活支援拠点等の5つの機能
# ---------------------------------------------------------------------------
def _kyoten(h_cm):
    """5つの機能と、会津北部4町村で面的に確保していることを示す。"""
    rows = tbls_juten("W-3")[0][1:]
    sh = Sheet(h_cm)
    top = 0.975
    y, h = sh.panel(0.02, 0, 0.96,
                    "令和３年度に整備。地域生活支援センターいなわしろに"
                    "コーディネーター業務を委託",
                    fs=9.0,
                    head="会津北部地域生活支援拠点"
                         "（猪苗代町・磐梯町・湯川村・北塩原村）",
                    head_fs=9.8, fc=KOIME["緑"], ec=COLORS["緑"], top=top)
    top = y - 0.030
    w = 0.188
    hs = []
    for kinou, houhou, _j in rows:
        inner = w - 2 * PAD
        hs.append(sh.h_of(kinou, inner, 8.8, pad_lines=0.55)
                  + sh.h_of(houhou, inner, 7.6, pad_lines=0.7))
    hm = max(hs)
    for i, (kinou, houhou, _j) in enumerate(rows):
        x = 0.02 + i * (w + 0.005)
        sh.box(x, top - hm, w, hm)
        sh.text(x + w / 2, top - 0.012, kinou, fs=8.8, w=w - 2 * PAD,
                ha="center")
        sh.text(x + PAD,
                top - sh.h_of(kinou, w - 2 * PAD, 8.8, pad_lines=0.55),
                houhou, fs=7.6, w=w - 2 * PAD)
        sh.arrow(x + w / 2, y, x + w / 2, top)
    low = top - hm
    sh.arrow(0.5, low, 0.5, low - 0.028)
    sh.panel(0.02, 0, 0.96,
             "村単独では確保が難しいため、４町村の拠点に参画する形で"
             "５つの機能を確保する。年１回以上、運用状況の検証・検討の場を"
             "設ける",
             fs=9.0, head="本村の関わり", head_fs=9.4, ec=COLORS["橙"],
             top=low - 0.028)
    sh.notes([
        "※ 国の基本指針 第二 五は、地域生活支援拠点等について５つの機能の"
        "確保を求めています（第４章２（５））。本計画期間は機能ごとに実績を"
        "記録し、年１回の検証・検討に用います。",
        "※ 機能ごとの令和６・７年度の実績は村に確認中です（Ｍ-54）。",
    ])
    return sh


def fig_kyoten():
    return fit(_kyoten, 11.4).save("06_地域生活支援拠点の5機能")


# ---------------------------------------------------------------------------
# 図7 緊急時の4つの場面
# ---------------------------------------------------------------------------
def _bamen(h_cm):
    """場面ごとに働く仕組みを並べ、確かめられていないところを示す。"""
    rows = tbls_juten("W-8")[0][1:]
    sh = Sheet(h_cm)
    midashi = "本人の身に何か起きたとき、場面ごとに働く仕組みが違う"
    sh.text(0.5, 0.985, midashi, fs=10.0, ha="center", va="top")
    top = 0.985 - sh.h_of(midashi, 0.96, 10.0, pad_lines=0.55)
    # 箱の間を空けておく。線が近すぎると、縁のぼかしで橙と緑が混ざった
    # 中間色の画素ができ、白黒にしたとき別の系列のように見えてしまう。
    aida = 0.014
    w = (0.97 - 3 * aida) / 4
    # 色は場面の区別ではなく、「村の高齢者施策（対象かを確認していない）」か
    # 「本計画に位置づけがある」かを表す。白黒でも濃さで読み分けられる。
    tone = [USUIRO["橙"], USUIRO["橙"], USUIRO["緑"], USUIRO["緑"]]
    edge = [COLORS["橙"], COLORS["橙"], COLORS["緑"], COLORS["緑"]]
    inner = w - 2 * PAD
    hs = []
    for bamen, jitai, shikumi in rows:
        hs.append(sh.h_of(bamen, inner, 9.8, pad_lines=0.6)
                  + sh.h_of(jitai, inner, 8.0, pad_lines=0.5)
                  + sh.h_of(shikumi, inner, 7.6, pad_lines=0.7))
    hm = max(hs)
    for i, (bamen, jitai, shikumi) in enumerate(rows):
        x = 0.015 + i * (w + aida)
        sh.box(x, top - hm, w, hm, ec=edge[i])
        hh = sh.h_of(bamen, inner, 9.8, pad_lines=0.6)
        sh.box(x, top - hh, w, hh, fc=tone[i], ec=edge[i])
        sh.text(x + w / 2, top - 0.010, bamen, fs=9.8, w=inner, ha="center")
        cur = top - hh - 0.008
        sh.text(x + PAD, cur, jitai, fs=8.0, w=inner)
        cur -= sh.h_of(jitai, inner, 8.0, pad_lines=0.5)
        sh.text(x + PAD, cur, shikumi, fs=7.6, w=inner)
    low = top - hm
    sh.panel(0.015, 0, 0.955,
             "平時と急病時の仕組みは高齢者を対象に作られており、"
             "障がいのある方が対象に含まれるかを確認していない"
             "（Ｍ-50・Ｍ-52）。"
             "介護者の不在時と災害時は本計画に位置づけがある",
             fs=8.8, head="確かめること", head_fs=9.2,
             fc=USUIRO["橙"], ec=COLORS["橙"], top=low - 0.028)
    sh.notes([
        "※ 橙は村の高齢者施策に依っているもの、"
        "緑は本計画に位置づけがあるものです。",
        "※ 災害時の備えは第５章２（11）、個別避難計画は第３章３、"
        "地域生活支援拠点の緊急時の受入れは第４章２（５）によります。",
    ])
    return sh


def fig_bamen():
    return fit(_bamen, 11.8).save("07_緊急時の4場面")


# ---------------------------------------------------------------------------
# 図8 65歳到達時の判定の流れ
# ---------------------------------------------------------------------------
def _flow65(h_cm):
    """一律に介護保険へ移さないことを、判定の順序で示す。"""
    rows = tbls_juten("W-1")[0][1:]
    sh = Sheet(h_cm)
    top = 0.975
    y, h = sh.panel(0.33, 0, 0.34, None, head="65歳に到達", head_fs=10.4,
                    fc=KOIME["緑"], ec=COLORS["緑"], top=top)
    top = y - 0.026
    steps = [
        ("①　本人の意向と心身の状況を聴き取る",
         "介護保険優先原則（障害者総合支援法第７条）は"
         "一律に適用するものではない"),
        ("②　介護保険の認定の有無を確かめる",
         "認定がないまま優先原則を当てはめない"),
        ("③　介護保険に相当するサービスがあるかを見る",
         "共同生活援助・施設入所支援・就労系・同行援護・行動援護は"
         "相当するものがない。障がい福祉で続ける"),
        ("④　介護保険で必要な量を確保できるかを見る",
         "空きがない、送迎の範囲外、医療的ケア・強度行動障害に"
         "対応できない場合は障がい福祉を続ける"),
    ]
    for t, sub in steps:
        y, h = sh.panel(0.055, 0, 0.89, sub, fs=8.4, head=t, head_fs=9.4,
                        top=top)
        if t != steps[-1][0]:
            sh.arrow(0.5, y, 0.5, y - 0.026)
        top = y - 0.026
    outs = [("介護保険へ移る", COLORS["橙"]),
            ("障がい福祉を続ける", COLORS["緑"]),
            ("両方を併せて使う", AXIS)]
    hs = [sh.h_of(t, 0.29 - 2 * PAD, 9.6, pad_lines=0.7) for t, _c in outs]
    hm = max(hs)
    for i, (t, c) in enumerate(outs):
        x = 0.055 + i * 0.3
        # ④の判断からは3つのいずれにもなりうる。1本だけ下ろすと
        # 真ん中の結論に決まったように読めるため、3本とも下ろす。
        sh.arrow(x + 0.145, top + 0.026, x + 0.145, top)
        sh.box(x, top - hm, 0.29, hm, ec=c)
        sh.text(x + 0.145, top - hm / 2, t, fs=9.6, w=0.29 - 2 * PAD,
                ha="center", va="center")
    top = top - hm - 0.026
    body = "\n".join("・" + r[0] for r in rows[:3])
    body2 = "\n".join("・" + r[0] for r in rows[3:])
    y, h = sh.panel(0.055, 0, 0.89, None,
                    head="いずれの場合も、次の６点を一人ずつ記録する"
                         "（第５章５（２））",
                    head_fs=9.2, fc=HAI, top=top)
    hb = sh.h_of(body, 0.42, 8.4, pad_lines=0.6)
    sh.box(0.055, y - hb, 0.89, hb, fc=HAI)
    sh.text(0.075, y - 0.008, body, fs=8.4, w=0.42)
    sh.text(0.075 + 0.44, y - 0.008, body2, fs=8.4, w=0.42)
    sh.notes([
        "※ 橙は介護保険、緑は障がい福祉、灰は両方を併せて使う場合です。",
        "※ 移行先は介護保険の給付だけではありません。村の高齢者福祉施策・"
        "村単独事業・地域支援事業への接続も確かめます"
        "（対象となるかは事業ごとに確認。Ｍ-50）。",
        "※ 65歳到達前５年間の支給決定等の要件を満たす方は、"
        "介護保険移行後の利用者負担が償還されます（第７章３）。",
    ])
    return sh


def _life(h_cm):
    """ライフコースと制度の移行。

    18歳到達（第5章5（1））と65歳到達（同（2））が別々の項に分かれて
    いるため、児童期から高齢期までを一続きで示す図がない。
    色は制度の別を表す。緑＝障がい福祉、橙＝ほかの制度。
    """
    sh = Sheet(h_cm)
    midashi = "児童期・成人期・高齢期を一続きでみる"
    sh.text(0.5, 0.985, midashi, fs=10.0, ha="center", va="top")
    top = 0.985 - sh.h_of(midashi, 0.96, 10.0, pad_lines=0.55)

    # --- 上段 ほかの制度（こども施策・高齢者施策） -----------------
    w, aida = 0.300, 0.021
    hoka = [
        ("こども施策",
         "保育所・幼稚園・学校\n子育て短期支援事業\n一時預かり・児童クラブ\n"
         "こども誰でも通園制度\n（第５章２（５））"),
        ("―", ""),
        ("介護保険・高齢者福祉",
         "介護保険の給付\n老人クラブ・シルバー人材センター\n介護予防教室\n"
         "村単独の在宅福祉事業\n（第５章５（２））"),
    ]
    inner = w - 2 * PAD
    hs = [sh.h_of(h, inner, 9.4, pad_lines=0.55)
          + sh.h_of(b2, inner, 8.0, pad_lines=0.7)
          for h, b2 in hoka if b2]
    hm = max(hs)
    for i, (head, body) in enumerate(hoka):
        if not body:
            continue
        x = 0.015 + i * (w + aida)
        sh.box(x, top - hm, w, hm, fc=USUIRO["橙"], ec=COLORS["橙"])
        sh.text(x + w / 2, top - 0.010, head, fs=9.4, w=inner, ha="center")
        sh.text(x + PAD,
                top - sh.h_of(head, inner, 9.4, pad_lines=0.55),
                body, fs=8.0, w=inner)
    ue_shita = top - hm

    # --- 中段 障がい福祉 -------------------------------------------
    #     上段との間に節目のラベルを入れるため、ここだけ広く空ける
    fushime = 0.070
    naka_top = ue_shita - fushime
    shogai = [
        ("児童期（０歳〜18歳）",
         "児童発達支援\n放課後等デイサービス\n保育所等訪問支援\n"
         "障がい児相談支援\n（第５章１（５））"),
        ("成人期（18歳〜65歳）",
         "訪問系・日中活動系\n居住系・相談支援\n地域生活支援事業\n"
         "（第５章１・３）"),
        ("高齢期（65歳〜）",
         "介護保険に相当しないもの\n（共同生活援助・就労系・\n"
         "同行援護・行動援護等）は続ける\n（第５章５（２））"),
    ]
    hs = [sh.h_of(h, inner, 9.4, pad_lines=0.55)
          + sh.h_of(b2, inner, 8.0, pad_lines=0.7) for h, b2 in shogai]
    hm2 = max(hs)
    for i, (head, body) in enumerate(shogai):
        x = 0.015 + i * (w + aida)
        sh.box(x, naka_top - hm2, w, hm2, fc=USUIRO["緑"], ec=COLORS["緑"])
        sh.text(x + w / 2, naka_top - 0.010, head, fs=9.4, w=inner,
                ha="center")
        sh.text(x + PAD,
                naka_top - sh.h_of(head, inner, 9.4, pad_lines=0.55),
                body, fs=8.0, w=inner)
        if i:
            sh.arrow(x - aida - 0.001, naka_top - hm2 / 2,
                     x, naka_top - hm2 / 2)
    naka_shita = naka_top - hm2

    # --- 制度が変わる2つの節目 ---------------------------------------
    for i, (t, sub) in enumerate((
            ("18歳到達", "障がい児相談支援→計画相談支援。"
                        "介護給付は障害支援区分の認定。"
                        "特別児童扶養手当→障害基礎年金（第５章５（１））"),
            ("65歳到達", "一律に移さず、４段の判定で決める"
                        "（第５章５（２）・図５-４）"))):
        # 縦線を引くと、上段のこども施策・介護保険の箱に矢印が刺さり、
        # そこへ移るように読めてしまう。段の間に見出しを置くだけにする。
        x = 0.015 + (i + 1) * (w + aida) - aida / 2
        sh.text(x, (ue_shita + naka_top) / 2, t, fs=9.4, ha="center",
                va="center", color=COLORS["橙"])
    y, h = sh.panel(0.015, 0, 0.955,
                    "18歳到達　障がい児相談支援は計画相談支援に、"
                    "特別児童扶養手当は障害基礎年金に切り替わる。"
                    "介護給付を使う場合は障害支援区分の認定が新たに要る"
                    "（第５章５（１））\n"
                    "65歳到達　一律に介護保険へ移さない。"
                    "本人の意向、認定の有無、相当するサービスの有無、"
                    "必要な量を確保できるかの４段で判定する"
                    "（第５章５（２）・図５-４）",
                    fs=8.4, head="制度が変わる二つの節目", head_fs=9.4,
                    ec=COLORS["橙"], top=naka_shita - 0.030)

    # --- 下段 どの時期も続くもの -------------------------------------
    sh.panel(0.015, 0, 0.955,
             "相談　村保健福祉課と相談支援専門員（第５章２（２）・６）\n"
             "権利擁護　成年後見制度、わたしの思いノート（第６章）\n"
             "緊急時・災害時　緊急時の受入れ、個別避難計画、"
             "事業所の業務継続計画（第４章２（５）・第５章２（11））\n"
             "冬季の地域生活　除雪、通院・通所の確保（第５章２（14））",
             fs=8.4, head="どの時期も続くもの", head_fs=9.4,
             fc=HAI, top=y - 0.026)

    sh.notes([
        "※ 橙はほかの制度（こども施策・介護保険・高齢者福祉）、"
        "緑は障がい福祉サービス等です。",
        "※ 一般の子育て施策と障がい福祉は入れ替わるものではなく、"
        "併せて使うものです。保育所等での受入れが進むことが、"
        "直ちに障がい児通所支援の減少を意味するものではありません"
        "（第５章２（５））。",
    ])
    return sh


def fig_life():
    return fit(_life, 18.0).save("09_ライフコースと制度の移行")


def _rendou(h_cm):
    """本計画と介護保険事業計画の関係。

    計画期間が完全に一致し、同じ方が両計画に現れる。
    数を揃えないと、村全体で二重に数えるか、どちらからも抜ける。
    """
    sh = Sheet(h_cm)
    midashi = "計画期間が同じで、同じ方が両方に現れる"
    sh.text(0.5, 0.985, midashi, fs=10.0, ha="center", va="top")
    top = 0.985 - sh.h_of(midashi, 0.96, 10.0, pad_lines=0.55)

    w = 0.465
    y1, h1 = sh.panel(0.015, 0, w,
                      "障害者総合支援法第88条第１項\n"
                      "児童福祉法第33条の20第１項\n"
                      "令和９年度～令和11年度",
                      fs=8.4, head="本計画\n（第８期障がい福祉計画・\n第４期障がい児福祉計画）",
                      head_fs=9.4, fc=USUIRO["緑"], ec=COLORS["緑"],
                      top=top)
    y2, h2 = sh.panel(0.52, 0, w,
                      "介護保険法第117条第１項\n"
                      "老人福祉法第20条の８第１項\n"
                      "令和９年度～令和11年度",
                      fs=8.4, head="第10期北塩原村高齢者福祉計画・"
                                   "介護保険事業計画",
                      head_fs=9.4, fc=USUIRO["橙"], ec=COLORS["橙"],
                      top=top)
    low = min(y1, y2)
    sh.arrow(0.48, low + min(h1, h2) / 2, 0.52, low + min(h1, h2) / 2,
             style="<|-|>", lw=1.0)

    y, h = sh.panel(0.015, 0, 0.97, None,
                    head="数を揃える四つの対象", head_fs=9.8,
                    fc=HAI, top=low - 0.030)
    rows = [
        ("65歳に到達する方",
         "障がい側で減らす数と介護保険側で増やす数を同じ年度に合わせる。"
         "年齢だけで自動的に動かさず、実際に移る方の分だけを動かす"),
        ("介護保険の適用除外",
         "指定障害者支援施設に生活介護及び施設入所支援の両方の支給決定を"
         "受けて入所している方は、介護保険の被保険者にならない"
         "（介護保険法施行法第11条第１項）。"
         "介護保険側の第１号被保険者数から除く"),
        ("第２号被保険者（40歳以上65歳未満）",
         "特定疾病により要介護認定を受けた方は65歳未満でも介護保険が"
         "優先する。両計画で同じ方を数えない"),
        ("共生型サービスの事業所",
         "同じ事業所が両方の指定を受けている場合、"
         "事業所数を足し合わせると地域の資源を過大に見せる"),
    ]
    inner = 0.97 - 2 * PAD
    cur = y - 0.012
    for i, (a2, b2) in enumerate(rows):
        hh = sh.h_of(f"{a2}　{b2}", inner, 8.2, pad_lines=0.25)
        sh.text(0.015 + PAD, cur, f"{a2}　{b2}", fs=8.2, w=inner)
        cur -= hh
    sh.box(0.015, cur - 0.004, 0.97, y - cur + 0.008, fc=HAI)
    sh.text(0.5, y - 0.012 + 0.0, "", fs=1)
    shita = cur - 0.004

    sh.panel(0.015, 0, 0.97,
             "成年後見制度の利用を支援する事業は、"
             "障害者総合支援法第77条の地域生活支援事業と、"
             "介護保険法第115条の45の地域支援事業の両方にあります。"
             "名前は同じでも別の事業であり、"
             "本計画が見込むのは前者です。"
             "両計画の数を足し合わせないよう、"
             "どちらの事業の数かを明らかにして示します"
             "（第５章３（２））。",
             fs=8.4, head="名前が同じで別のもの", head_fs=9.4,
             ec=COLORS["緑"], top=shita - 0.028)

    sh.notes([
        "※ 緑は本計画、橙は第10期北塩原村高齢者福祉計画・"
        "介護保険事業計画を表します。",
        "※ 両計画の整合は第７章３に、"
        "65歳到達時の判定は第５章５（２）によります。",
        "※ 高齢化率は、本計画が住民基本台帳の実績を、"
        "介護保険事業計画が国立社会保障・人口問題研究所の推計を"
        "用いているため、同じ年でも値が異なります。",
    ])
    return sh


def fig_rendou():
    return fit(_rendou, 16.0).save("10_両計画の関係")


def fig_flow65():
    return fit(_flow65, 13.2).save("08_65歳到達時の判定の流れ")


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
    dict(no="Z-4", file="04_計画の体系.png", kata="本文",
         title="計画の体系",
         src="資料：第４次北塩原村障がい者計画及び本計画により作成",
         ichi="第４次北塩原村障がい者計画では、７つの基本施策ごとに"
              "令和11年度の目標値を設定しています。"
              "本計画期間においては、障がい福祉サービス等の提供体制の確保と"
              "合わせて、次の取組みを進め、目標の達成を目指します。"),
    dict(no="Z-5", file="05_関係機関との連携.png", kata="表の後",
         title="関係機関との連携",
         src="資料：北塩原村（第５章６により作成）",
         ichi="本村は人口2,316人（令和８年４月１日現在）の"
              "小規模な自治体であり、村内で完結する支援は限られます。"),
    dict(no="Z-6", file="06_地域生活支援拠点の5機能.png", kata="表の後",
         title="地域生活支援拠点等の５つの機能",
         src="資料：国の基本指針 第二 五及び北塩原村の資料により作成",
         ichi="地域生活支援拠点等には、相談、緊急時の受入れ及び対応、"
              "体験の機会及び場、専門的人材の確保及び養成、"
              "地域の体制づくりの５つの機能が求められています。"),
    dict(no="Z-7", file="07_緊急時の4場面.png", kata="表の後",
         title="緊急時に働く仕組み（４つの場面）",
         src="資料：北塩原村（第５章２（11）により作成）",
         ichi="本村では、本人の身に何か起きたときの備えが"
              "場面ごとに分かれています。"),
    dict(no="Z-10", file="10_両計画の関係.png", kata="本文",
         title="本計画と介護保険事業計画の関係",
         src="資料：北塩原村（第７章３及び第10期北塩原村高齢者福祉計画・"
             "介護保険事業計画の素案により作成）",
         ichi="二つめは成年後見制度の利用を支援する事業です。"),
    dict(no="Z-9", file="09_ライフコースと制度の移行.png", kata="表の後",
         title="ライフコースと制度の移行",
         src="資料：北塩原村（第５章５及び第７章により作成）",
         ichi="障がい福祉サービスは、一定の年齢に達することにより、"
              "対象となる制度が変わります。"),
    dict(no="Z-8", file="08_65歳到達時の判定の流れ.png", kata="表の後",
         title="65歳到達時の判定の流れ",
         src="資料：障害者総合支援法第７条及び本計画により作成",
         ichi="65歳到達時に介護保険へ移行するかどうかは、"
              "サービスの種類により扱いが異なります。"),
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

    namae = list(COLORS)
    hues = [hue(hex2rgb(COLORS[k])) for k in namae]
    for path in made:
        im = Image.open(path).convert("RGB")
        small = im.resize((min(im.width, 300), min(im.height, 300)))
        try:
            px = list(small.get_flattened_data())
        except AttributeError:
            px = list(small.getdata())
        tsukatta = set()
        hoka = None
        for rgb, n in collections.Counter(px).most_common(40):
            if n < 150 or max(rgb) - min(rgb) <= 14:
                continue      # 白地・黒文字・グレーの目盛線は見ない
            h = hue(rgb)
            chikai = [k for k, q in zip(namae, hues)
                      if min(abs(h - q), 360 - abs(h - q)) <= 12]
            if chikai:
                tsukatta.add(chikai[0])
            elif hoka is None:
                hoka = (rgb, h)
        if hoka:
            ng.append(f"決めた色以外が使われている: "
                      f"{os.path.basename(path)} {hoka[0]}"
                      f"（色相{hoka[1]:.0f}度）")
        # ②b 同じ1枚の中で、グレーにすると潰れる組を使っていないこと
        #     （図の中で色に意味を持たせている以上、白黒でも読めること）
        nokori = sorted(tsukatta)
        for i in range(len(nokori)):
            for j in range(i + 1, len(nokori)):
                ga, gb = gray255(COLORS[nokori[i]]), gray255(COLORS[nokori[j]])
                if abs(ga - gb) < GRAY_MIN:
                    ng.append(f"同じ図で潰れる色を使っている: "
                              f"{os.path.basename(path)} "
                              f"{nokori[i]}{ga} と {nokori[j]}{gb}")

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
          f"作った図{len(made)}点が決めた3色の範囲／1枚ごとに潰れる組なし／"
          f"本文幅{BODY_W_CM}cmに収まる／高齢化率の年が昇順／原典の行数と一致")


def main():
    name = setup_font()
    p1, years, kaigo, jido = fig_kyufu()
    p2, svc, vals, pcts = fig_service()
    p3, xs, ys = fig_koureika()
    made = [p1, p2, p3,
            fig_taikei(), fig_renkei(), fig_kyoten(), fig_bamen(),
            fig_flow65(), fig_life(), fig_rendou()]
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
