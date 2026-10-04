# -*- coding: utf-8 -*-
"""概要版（8頁程度・3,500部）の下書き（令和8年10月4日）

08_作業順位 の順位8の続き。構成案（川崎町_概要版_構成案_R8.10.2.xlsx）に
沿って、**実際の本文と図を作る。**

  ① 1色刷り用の図（濃淡と模様で見分ける）　08_図表/概要版/*.png
  ② 概要版の下書き　01_第10期_最新版成果品/川崎町_計画素案_概要版_R8.10.4.docx

仕様書8　A4版／コート紙／1色刷り／8頁程度／3,500部

**数値は素案の本文と図（data_zuhyo.py）からのみ採る。**
概要版のために新しい数値を作らない。
素案を直したら本スクリプトを実行し直す。

⚠ 保険料は第3回策定委員会（令和9年1月）で決定するため、
  本下書きは算定Ｂ（基金50％取崩・6,143円）を仮に置き、
  確定前であることを明記している。
"""
import copy
import os
import re
import shutil
import sys

import docx
import matplotlib
matplotlib.use("Agg")
import matplotlib.font_manager as fm          # noqa: E402
import matplotlib.pyplot as plt               # noqa: E402
from docx.oxml.ns import qn                   # noqa: E402
from docx.shared import Emu, Pt               # noqa: E402

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from data_zuhyo import ZU                      # noqa: E402

SOAN = "01_第10期_最新版成果品/川崎町_計画書素案_v2.10_制度改正補完版.docx"
OUT = "01_第10期_最新版成果品/川崎町_計画素案_概要版_R8.10.4.docx"
FIGDIR = "08_図表/概要版"
EMU_IN = 914400

FONT = next(f for f in ("/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf",
                        "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf")
            if os.path.exists(f))
JP = fm.FontProperties(fname=FONT, size=11)
JP_S = fm.FontProperties(fname=FONT, size=9)
JP_M = fm.FontProperties(fname=FONT, size=10)
JP_T = fm.FontProperties(fname=FONT, size=13)
JP_XS = fm.FontProperties(fname=FONT, size=8)
plt.rcParams["axes.unicode_minus"] = False

# 1色刷り　濃淡と模様で見分ける
K0, K1, K2, K3 = "#1A1A1A", "#595959", "#A6A6A6", "#D9D9D9"
HATCH = ["", "///", "...", "xxx"]


def zu(no):
    return next(d for d in ZU if d["no"] == no)


def save(fig, name):
    os.makedirs(FIGDIR, exist_ok=True)
    p = os.path.join(FIGDIR, name)
    fig.savefig(p, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return p


def g_nenrei():
    d = zu("図2-1")
    fig, ax = plt.subplots(figsize=(6.4, 3.4), dpi=150)
    lab = [r[0].replace("（", "\n（") for r in d["rows"]]
    a = [r[1] for r in d["rows"]]
    b = [r[2] for r in d["rows"]]
    x = range(len(lab))
    w = 0.36
    ax.bar([i - w / 2 for i in x], a, w, color=K2, edgecolor=K0,
           linewidth=1.2, label="令和3年度末")
    ax.bar([i + w / 2 for i in x], b, w, color="white", edgecolor=K0,
           linewidth=1.2, hatch="///", label="令和7年度末")
    for i, (u, v) in enumerate(zip(a, b)):
        ax.text(i - w / 2, u + 25, f"{u:,}", ha="center", fontproperties=JP_S)
        ax.text(i + w / 2, v + 25, f"{v:,}", ha="center", fontproperties=JP_S)
    ax.set_xticks(list(x))
    ax.set_xticklabels(lab, fontproperties=JP_S)
    ax.set_ylabel("人", fontproperties=JP_S)
    ax.set_ylim(0, max(max(a), max(b)) * 1.25)
    ax.legend(prop=JP_S, frameon=False, ncol=2, loc="upper right")
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.set_title("65歳以上の方の年齢の内わけ", fontproperties=JP_T, color=K0)
    return save(fig, "g1_nenrei.png")


def g_koreika():
    d = zu("図2-2")
    fig, ax = plt.subplots(figsize=(6.4, 2.6), dpi=150)
    lab = [r[0] for r in d["rows"]]
    v = [r[1] for r in d["rows"]]
    bars = ax.barh(lab, v, color=[K0, K2, "white"], edgecolor=K0,
                   linewidth=1.2, hatch=["", "", "..."])
    for b, x in zip(bars, v):
        ax.text(x + 0.8, b.get_y() + b.get_height() / 2, f"{x}%",
                va="center", fontproperties=JP_M)
    ax.set_xlim(0, 50)
    ax.set_yticklabels(lab, fontproperties=JP_M)
    ax.grid(axis="x", linestyle=":", alpha=0.4)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.invert_yaxis()
    ax.set_title("高齢化率のくらべ方（令和7年3月31日）", fontproperties=JP_T,
                 color=K0)
    return save(fig, "g2_koreika.png")


def g_jukyusha():
    d = zu("図2-5")
    fig, ax = plt.subplots(figsize=(5.4, 3.4), dpi=150)
    v = [r[1] for r in d["rows"]]
    tot = sum(v)
    lab = [f"{r[0]}\n{r[1]:,.1f}人（{r[1] / tot * 100:.1f}%）"
           for r in d["rows"]]
    w, _ = ax.pie(v, colors=[K1, K3, "white"], startangle=90,
                  counterclock=False,
                  wedgeprops=dict(width=0.44, edgecolor=K0, linewidth=1.4),
                  hatch=["", "", "///"])
    ax.legend(w, lab, prop=JP_S, loc="center left",
              bbox_to_anchor=(0.96, 0.5), frameon=False)
    ax.text(0, 0.04, f"{tot:,.1f}人", ha="center", va="center",
            fontproperties=JP_T, color=K0)
    ax.set_title("介護サービスを使っている方（月あたり）",
                 fontproperties=JP_T, color=K0)
    return save(fig, "g3_jukyusha.png")


def g_kyufuhi():
    d = zu("図2-6")
    fig, ax = plt.subplots(figsize=(5.6, 3.2), dpi=150)
    lab = [r[0].replace("\n", "") for r in d["rows"]]
    v = [r[1] for r in d["rows"]]
    tot = sum(v)
    bars = ax.bar(range(len(v)), v, color=[K1, K3, "white"], edgecolor=K0,
                  linewidth=1.3, width=0.55, hatch=["", "", "///"])
    for b, x in zip(bars, v):
        ax.text(b.get_x() + b.get_width() / 2, x + 0.1,
                f"{x:.2f}億円\n({x / tot * 100:.1f}%)", ha="center",
                fontproperties=JP_S)
    ax.set_xticks(range(len(lab)))
    ax.set_xticklabels([s.replace("（", "\n（") for s in lab],
                       fontproperties=JP_S)
    ax.set_ylim(0, max(v) * 1.3)
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.set_title("介護にかかったお金（令和7年度・9.98億円）",
                 fontproperties=JP_T, color=K0)
    return save(fig, "g4_kyufuhi.png")


def g_taihi():
    d = zu("図3-2")
    fig, ax = plt.subplots(figsize=(6.4, 3.2), dpi=150)
    rows = d["rows"]
    lab = [r[0].replace("\n", "") for r in rows]
    kei = [r[1] / 1000 for r in rows]
    jis = [r[2] / 1000 for r in rows]
    x = [0, 1, 2, 3.4]
    w = 0.38
    ax.bar([i - w / 2 for i in x], kei, w, color=K3, edgecolor=K0,
           linewidth=1.2, label="計画")
    ax.bar([i + w / 2 for i in x], jis, w, color="white", edgecolor=K0,
           linewidth=1.2, hatch="///", label="実際")
    for i, (a, b) in zip(x, zip(kei, jis)):
        ax.text(i, max(a, b) + 30, f"{b / a * 100:.1f}%", ha="center",
                fontproperties=JP_S)
    ax.set_xticks(x)
    ax.set_xticklabels([s.replace("（", "\n（") for s in lab],
                       fontproperties=JP_S)
    ax.set_ylabel("百万円", fontproperties=JP_S)
    ax.set_ylim(0, max(max(kei), max(jis)) * 1.22)
    ax.legend(prop=JP_S, frameon=False, loc="upper left")
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.set_title("第9期の計画と実際（令和7年度）", fontproperties=JP_T,
                 color=K0)
    return save(fig, "g5_taihi.png")


def g_kpi():
    """認知症のKPIの3層構造（概要版は層と指標数だけを示す）"""
    import collections
    from matplotlib.patches import FancyArrow, FancyBboxPatch
    d = zu("図6-1")
    g = collections.OrderedDict()
    for layer, mok, name, joukyou in d["rows"]:
        g.setdefault(layer, []).append(name)
    fig, ax = plt.subplots(figsize=(6.0, 2.8), dpi=150)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 4.2)
    ax.axis("off")
    face = [K3, K2, K1]
    txtc = [K0, K0, "white"]
    for i, ((layer, items), fc, tc) in enumerate(
            zip(g.items(), face, txtc)):
        x = 0.4 + i * 3.25
        ax.add_patch(FancyBboxPatch(
            (x, 1.1), 2.9, 2.0,
            boxstyle="round,pad=0.04,rounding_size=0.12",
            facecolor=fc, edgecolor=K0, linewidth=1.4))
        ax.text(x + 1.45, 2.55, layer.split("（")[0], ha="center",
                fontproperties=JP_T, color=tc)
        ax.text(x + 1.45, 2.05, "（" + layer.split("（")[1], ha="center",
                fontproperties=JP_S, color=tc)
        ax.text(x + 1.45, 1.45, f"{len(items)}つの目安", ha="center",
                fontproperties=JP_M, color=tc)
        if i < 2:
            ax.add_patch(FancyArrow(x + 2.95, 2.1, 0.26, 0, width=0.14,
                                    head_width=0.4, head_length=0.2,
                                    length_includes_head=True,
                                    facecolor="white", edgecolor=K0,
                                    linewidth=1.2))
    ax.text(5.0, 0.5, "活動の量 → できたこと → 住民の変化、の順に"
                      "たしかめます", ha="center", fontproperties=JP_M,
            color=K0)
    ax.set_title("認知症の取組をどうたしかめるか", fontproperties=JP_T,
                 color=K0)
    return save(fig, "g6_kpi.png")


def g_hokenryo():
    d = zu("図9-1")
    fig, ax = plt.subplots(figsize=(6.4, 3.4), dpi=150)
    rows = d["rows"]
    lab = ["第8期\n(R3〜R5)", "第9期\n(R6〜R8)",
           "第10期\nＡ 基金を\n取りくずさない",
           "第10期\nＢ 基金の半分を\n取りくずす",
           "第10期\nＣ 基金を全額\n取りくずす"]
    v = [r[1] for r in rows]
    face = [K2, K2, "white", K3, "white"]
    hat = ["", "", "///", "", "..."]
    bars = ax.bar(range(len(v)), v, color=face, edgecolor=K0, linewidth=1.3,
                  width=0.56, hatch=hat)
    for i, (b, x) in enumerate(zip(bars, v)):
        s = f"{x:,}円"
        if i >= 2:
            sa = rows[i][2]
            s += "\n（第9期比 %s%s円）" % ("＋" if sa > 0 else "▲", f"{abs(sa):,}")
        ax.text(b.get_x() + b.get_width() / 2, x + 90, s, ha="center",
                fontproperties=JP_S)
    ax.set_xticks(range(len(lab)))
    ax.set_xticklabels(lab, fontproperties=JP_S)
    ax.set_ylabel("月額（円）", fontproperties=JP_S)
    ax.set_ylim(0, max(v) * 1.26)
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.set_title("介護保険料（月額）の見通し", fontproperties=JP_T, color=K0)
    return save(fig, "g7_hokenryo.png")


def g_hashira():
    """7つの柱と4つの重点（1色刷り）"""
    from matplotlib.patches import FancyBboxPatch
    HASHIRA = ["柱1\n健康づくりの推進",
               "柱2\n安心して暮らせる\nまちづくり",
               "柱3\n地域生活を支援する\n取り組み",
               "柱4\n地域支援事業の充実",
               "柱5\n認知症施策の推進",
               "柱6\n介護・介護予防\nサービスの充実",
               "柱7\n計画の進め方"]
    JUTEN = ["重点1 認知症とともに生きるまちづくり",
             "重点2 移動の手段の立て直し",
             "重点3 介護する家族への支援",
             "重点4 サービスを提供する体制の維持"]
    fig, ax = plt.subplots(figsize=(6.4, 3.6), dpi=150)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6.2)
    ax.axis("off")
    for i, h in enumerate(HASHIRA):
        c, r = i % 4, i // 4
        x = 0.25 + c * 2.45
        y = 4.0 - r * 1.35
        ax.add_patch(FancyBboxPatch(
            (x, y), 2.25, 1.1,
            boxstyle="round,pad=0.03,rounding_size=0.10",
            facecolor=K3 if i != 4 else K1, edgecolor=K0, linewidth=1.3))
        ax.text(x + 1.125, y + 0.55, h, ha="center", va="center",
                fontproperties=JP_XS, color=K0 if i != 4 else "white",
                linespacing=1.35)
    ax.text(5.0, 5.7, "7つの柱で進めます", ha="center",
            fontproperties=JP_T, color=K0)
    ax.text(5.0, 1.9, "とくに力を入れる4つのこと", ha="center",
            fontproperties=JP_M, color=K0)
    for i, j in enumerate(JUTEN):
        ax.text(0.3 + (i % 2) * 5.0, 1.35 - (i // 2) * 0.55,
                "■ " + j, ha="left", fontproperties=JP_XS, color=K0)
    return save(fig, "g8_hashira.png")


FIGS = [("g1", g_nenrei), ("g2", g_koreika), ("g3", g_jukyusha),
        ("g4", g_kyufuhi), ("g5", g_taihi), ("g6", g_kpi),
        ("g7", g_hokenryo), ("g8", g_hashira)]

# ══════════════════════════ 概要版の本文
#   （種別, 文）　種別は 表題／見出し／本文／囲み／図
P = [
    ("表題", "川崎町高齢者保健福祉計画\n第10期介護保険事業計画\n概要版"),
    ("本文", "令和9年度から令和11年度までの3年間の、"
             "高齢者福祉と介護保険の進め方を定めた計画です。"),
    ("本文", "この計画は、65歳以上の方の介護保険料と、"
             "町が行う高齢者福祉の3年間を定めるものです。"
             "計画の全文は町のホームページでご覧いただけます。"),
    ("見出し", "１　川崎町の高齢者のいま"),
    ("本文", "川崎町の高齢化率は令和8年6月末で42.0％です"
             "（総人口7,648人のうち65歳以上が3,209人）。"
             "全国平均29.1％・宮城県平均28.5％を大きく上回り、"
             "県内35市町村のなかで5番目に高い町です。"),
    ("図", "g1_nenrei.png"),
    ("本文", "65歳以上の方の数はほぼ横ばいですが、"
             "75歳以上の方が4年で15.4％増えています。"
             "介護が必要になる方が増えるのは、この年代からです。"),
    ("図", "g2_koreika.png"),
    ("見出し", "２　介護サービスの利用とお金"),
    ("図", "g3_jukyusha.png"),
    ("本文", "令和7年度は、ひと月あたり延べ450.0人の方が"
             "介護サービスを利用しました。"),
    ("図", "g4_kyufuhi.png"),
    ("本文", "介護にかかったお金は年間9.98億円です。"
             "そのうち約半分（49.4％）は特別養護老人ホームなどの"
             "施設サービスです。"
             "町外の施設を利用している方も多く、"
             "令和8年3月末で24人の方が住所地特例の対象となっています。"),
    ("見出し", "３　第9期（令和6〜8年度）をふりかえって"),
    ("図", "g5_taihi.png"),
    ("本文", "令和7年度は、計画の10億1,102万円に対し、"
             "実際には9億9,822万円でした（98.7％）。"
             "全体ではほぼ計画どおりですが、"
             "自宅で受けるサービスが計画を8.8％下回る一方、"
             "施設のサービスは3.3％上回りました。"
             "自宅から施設へという流れが進んでいます。"),
    ("囲み", "第10期の5つの課題　①サービスを提供する体制をどう保つか"
             "②自宅で暮らし続けるための支え　③認知症とともに生きるまちづくり"
             "④介護する家族の負担　⑤取組の成果をどう確かめるか"),
    ("見出し", "４　第10期で大切にすること"),
    ("本文", "基本理念は第9期から引き継ぎます。"
             "そのうえで、7つの柱に沿って施策を進め、"
             "4つのことを重点として取り組みます。"),
    ("囲み", "重点① サービスを提供する体制の維持　"
             "重点② 認知症施策の推進（認知症基本法への対応）　"
             "重点③ 介護する家族への支援　"
             "重点④ 取組と成果を結びつける仕組み"),
    ("図", "g8_hashira.png"),
    ("本文", "柱1は健康づくりと介護予防、柱2は住まい・防災・防犯、"
             "柱3は移動の手段・見守り・社会参加と介護する家族への支援、"
             "柱4は介護予防・日常生活支援総合事業と地域包括支援センター、"
             "柱5は認知症の取組、柱6は介護サービス、"
             "柱7は計画の進め方と介護人材の確保です。"),
    ("本文", "とくに、町内の介護事業所が小さくなっていることへの備えを"
             "新しい重点に加えました。"
             "ホームヘルパーの数が減っており、"
             "いま受けられているサービスを3年後も受けられるようにすることが、"
             "この計画のいちばんの課題です。"),
    ("見出し", "５　認知症とともに生きるまちへ"),
    ("本文", "令和6年に認知症基本法が施行されました。"
             "この法律は、認知症になってからも希望をもって自分らしく"
             "暮らすことができる社会をめざしています。"
             "川崎町は第10期から、認知症の取組をこの計画の中に位置づけます。"),
    ("図", "g6_kpi.png"),
    ("本文", "認知症の本人とご家族からお話をうかがったうえで"
             "計画をつくります（令和8年10月から11月に実施します）。"
             "認知症カフェ「喫茶みかん」（年47回）、"
             "認知症サポーターの養成、チームオレンジの活動、"
             "認知症初期集中支援チームによる訪問を続けます。"
             "第10期からは、認知症の本人が集まって語り合う"
             "「本人ミーティング」を新しく始めます。"),
    ("囲み", "もの忘れが気になったら　"
             "川崎町地域包括支援センター（保健福祉課内）にご相談ください。"
             "もの忘れ相談を毎月開いています。"),
    ("見出し", "６　介護保険料"),
    ("本文", "介護保険料は、3年間に必要な介護サービスのお金から、"
             "国・県・町が出す分と、40歳から64歳の方が出す分を差し引いて、"
             "65歳以上の方で分け合って決めます。"
             "これまでにためてきた基金（準備基金）を取りくずすと、"
             "保険料を抑えることができます。"),
    ("図", "g7_hokenryo.png"),
    ("本文", "第9期の月額6,500円に対して、3つの案があります。"
             "基金を取りくずさないと6,822円、"
             "半分を取りくずすと6,143円、"
             "全額を取りくずすと5,464円です。"
             "基金を使えばいまの保険料は下がりますが、"
             "次の第11期に使える分がなくなります。"),
    ("囲み", "⚠ 保険料の額は令和9年1月の策定委員会で決めます。"
             "上の金額は令和8年9月時点の試算であり、まだ確定していません。"
             "保険料は、所得に応じて13の段階に分かれます。"),
    ("見出し", "７　計画の進め方とご相談先"),
    ("本文", "計画は、毎年度その進み具合をたしかめ、"
             "結果を町のホームページで公表します。"
             "計画に定めた目安（KPI）がどこまで達成できたかを、"
             "策定委員会にご報告します。"),
    ("囲み", "ご相談は　川崎町保健福祉課／"
             "川崎町地域包括支援センター　までお願いします。"
             "介護予防の教室や通いの場のご案内もしています。"),
    ("本文", "この計画の案については、意見を募る手続"
             "（パブリックコメント）を行います。"
             "実施の時期は町の広報紙とホームページでお知らせします。"
             "みなさまのご意見をお寄せください。"),
    ("見出し", "８　この冊子に出てくることば"),
    ("本文", "要介護認定　介護が必要な状態かどうか、どの程度かを定める手続です。"
             "全国共通の基準により、要支援1・2と要介護1〜5の7つに分かれます。"),
    ("本文", "地域包括支援センター　高齢者の相談窓口です。"
             "保健師・社会福祉士などが、介護・福祉・権利を守ることについて"
             "相談に応じます。"),
    ("本文", "日常生活圏域　サービスを考えるうえでの区域です。"
             "川崎町は町全体を1つの区域とし、"
             "7つの地区ごとに状況を把握します。"),
    ("本文", "住所地特例　施設に入るために他の市町村へ住所を移した方について、"
             "移る前の市町村が引き続き保険者となる仕組みです。"),
    ("本文", "介護給付費準備基金　保険料の余りをためておき、"
             "次の3年間の保険料の上がり方をやわらげるために使うお金です。"),
    ("本文", "認知症サポーター　認知症について正しく理解し、"
             "認知症の人とご家族を見守る応援者です。"
             "養成講座を受けた方で、川崎町では累計642名を養成しています。"),
    ("本文", "チームオレンジ　認知症サポーターが、"
             "認知症の人やご家族の身近な支援につながるよう結成するチームです。"
             "川崎町は令和6年度に整備しました。"),
    ("本文", "通いの場　住民が主体となって運営する、"
             "体操や交流などの集まりです。"),
    ("本文", "［発行］川崎町保健福祉課"),
]


def strip_unused_media(path):
    """素案から引き継いだ、本文が参照しない画像を取り除く。

    素案を複製して中身を入れ替えているため、素案の図10点が
    使われないまま残る。成果品として配るものであり、取り除く。
    rels は正規表現ではなく XML として扱う（壊さないため）。
    """
    import zipfile
    from lxml import etree
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        data = {n: z.read(n) for n in names}
    doc = data["word/document.xml"].decode("utf-8")
    used = set(re.findall(r'r:embed="(rId\d+)"', doc))
    root = etree.fromstring(data["word/_rels/document.xml.rels"])
    drop = set()
    for rel in list(root):
        tgt = rel.get("Target") or ""
        rid = rel.get("Id") or ""
        if tgt.startswith("media/") and rid not in used:
            drop.add("word/" + tgt)
            root.remove(rel)
    if not drop:
        return 0
    data["word/_rels/document.xml.rels"] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            if n in drop:
                continue
            zout.writestr(n, data[n])
    shutil.move(tmp, path)
    return len(drop)


def set_el(el, text):
    from fix_soan_v111 import set_el as f
    return f(el, text)


def main():
    made = [fn() for _, fn in FIGS]
    print("1色刷りの図を%d点作りました" % len(made))
    for p in made:
        print("  ", p, f"{os.path.getsize(p):,}バイト")

    # ══════════════════════════ 概要版の docx（素案の書式を引き継ぐ）
    shutil.copy(SOAN, OUT)
    doc = docx.Document(OUT)
    body = doc.element.body

    tpl = {}
    for el in body.iterchildren():
        if el.tag != qn("w:p"):
            continue
        s = "".join(n.text or "" for n in el.iter(qn("w:t"))).strip()
        if "表題" not in tpl and s.startswith("第 1 章"):
            tpl["表題"] = copy.deepcopy(el)
        if "見出し" not in tpl and s.startswith("▌"):
            tpl["見出し"] = copy.deepcopy(el)
        if "本文" not in tpl and s.startswith("令和7年度から令和8年度への"):
            tpl["本文"] = copy.deepcopy(el)
        if "囲み" not in tpl and s.startswith("⚠ 宮城県自身の評価"):
            tpl["囲み"] = copy.deepcopy(el)
        if "図" not in tpl and el.findall(".//" + qn("w:drawing")):
            tpl["図"] = copy.deepcopy(el)
    for k in ("表題", "見出し", "本文", "囲み", "図"):
        if k not in tpl:
            raise SystemExit("雛形が見つからない：" + k)

    sectPr = body.find(qn("w:sectPr"))
    for el in list(body.iterchildren()):
        if el is not sectPr:
            body.remove(el)

    def add(el):
        if sectPr is not None:
            sectPr.addprevious(el)
        else:
            body.append(el)

    n_fig = 0
    for kind, v in P:
        if kind == "図":
            path = os.path.join(FIGDIR, v)
            doc.add_picture(path, width=Emu(int(5.6 * EMU_IN)))
            p = doc.paragraphs[-1]._p
            p.getparent().remove(p)
            pPr = tpl["図"].find(qn("w:pPr"))
            if pPr is not None:
                old = p.find(qn("w:pPr"))
                if old is not None:
                    p.remove(old)
                p.insert(0, copy.deepcopy(pPr))
            tr = tpl["図"].find(qn("w:r"))
            rPr = tr.find(qn("w:rPr")) if tr is not None else None
            if rPr is not None:
                for r in p.findall(qn("w:r")):
                    if r.find(qn("w:rPr")) is None:
                        r.insert(0, copy.deepcopy(rPr))
            add(p)
            n_fig += 1
        else:
            el = copy.deepcopy(tpl[kind])
            set_el(el, v)
            add(el)
    doc.save(OUT)
    strip_unused_media(OUT)

    # ══════════════════════════ 自己点検
    d2 = docx.Document(OUT)
    text = "\n".join(q.text for q in d2.paragraphs)
    ng = []
    import zipfile
    z = zipfile.ZipFile(OUT)
    n_media = sum(1 for n in z.namelist() if n.startswith("word/media"))
    n_img = len(re.findall(r"<w:drawing>",
                           z.read("word/document.xml").decode("utf-8")))
    if n_img != len(FIGS):
        ng.append(f"本文の図が{n_img}点（{len(FIGS)}点のはず）")
    if n_media != len(FIGS):
        ng.append(f"使っていない画像が残っている（{n_media}点）")
    # 概要版に受託者を主語とする語・確認事項の参照がないこと
    for w in ("受託者", "当社", "ビズアップ", "確認事項No",
              "【町確認】", "【委員会協議】"):
        if w in text:
            ng.append(f"概要版に「{w}」が残っている")
    # 数値が素案と合っているか（主なもの）
    for v in ("42.0", "9.98", "450.0", "6,500", "6,822", "6,143", "5,464",
              "98.7"):
        if v not in text:
            ng.append(f"主な数値が入っていない：{v}")
    n_char = sum(len(q.text) for q in d2.paragraphs)
    print("保存：", OUT)
    print(f"  段落 {len(d2.paragraphs)}／図 {n_img}"
          f"／画像 {n_media}／字数 {n_char:,}字")
    print("  ── 自己点検")
    if ng:
        for m in ng:
            print("   ×", m)
        sys.exit(1)
    print(f"   ○ 図{len(FIGS)}点・受託者を主語とする語なし・"
          "未確定箇所の表記なし・主な数値は素案と一致")
    print("  ⚠ 頁数は組版によります。Word で開いてご確認ください。")
    print("  ⚠ 保険料は令和9年1月の策定委員会で決まるため、"
          "本下書きは確定前である旨を明記しています。")


if __name__ == "__main__":
    main()
