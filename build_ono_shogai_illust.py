"""障がい計画：イラスト起こしの割付案（仕様書 成果品②「レイアウトデザイン（イラスト起こし数点）」）。

図表の点検（令和8年10月23日）で挙げた候補5点について、置く場所・大きさ・描く要素を決め、
ラフ（イラスト起こしの前の下絵）を付ける。ラフの文言は、他メンバー素案（変更履歴付き 第2弾）の
表から読む（書き写さない）。デザイナーへの発注・町との選定に用いる。

出力
  10_納品・印刷/小野町_障がい計画_イラスト起こしの割付案_YYYYMMDD.docx
  output/ono_fig/f9x_shogai_rough_*.png（ラフ）
"""

import pathlib
import re
import sys
import textwrap

import matplotlib.patches as mpatches

import build_ono_shogai_taso_hanei2 as B
import build_ono_soan_shogai as SO
import ono_figs as FG
import ono_style as S

plt = FG.plt
HERE = pathlib.Path(__file__).parent
OUT = HERE / "小野町_引継ぎ_整理済" / "10_納品・印刷"
ASOF = "20261023"
ASOF_JP = "令和8年10月23日"
W = B.W


def dodai_tables():
    """他メンバー素案（第2弾・承諾後の姿）の表を題名で読む。{題名: rows}"""
    root = B.accepted(B.OUT)
    kids = list(root.find(W + "body"))
    tx = lambda e: "".join(t.text or "" for t in e.iter(W + "t"))
    out = {}
    for i, e in enumerate(kids):
        m = re.match(r"【表\d+】(.+)", tx(e)) if e.tag == W + "p" else None
        if m and i + 1 < len(kids) and kids[i + 1].tag == W + "tbl":
            out[m.group(1).strip()] = [[tx(c) for c in r.findall(W + "tc")] for r in kids[i + 1].findall(W + "tr")]
    return out


def _box(ax, x, y, w, h, text, fc=FG.PALE, ec=FG.NAVY, sz=8, bold=False, wrap=14, color=FG.NAVY):
    ax.add_patch(mpatches.FancyBboxPatch((x - w / 2, y - h / 2), w, h, boxstyle="round,pad=0.01",
                                         fc=fc, ec=ec, lw=1))
    ax.text(x, y, "\n".join(textwrap.wrap(text, wrap)), ha="center", va="center", fontsize=sz,
            color=color, fontweight="bold" if bold else "normal")


def _line(ax, x1, y1, x2, y2):
    ax.plot([x1, x2], [y1, y2], color=FG.GREY, lw=1, zorder=0)


def _fig(w=7.4, h=4.4):
    fig, ax = plt.subplots(figsize=(w, h))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6)
    ax.axis("off")
    return fig, ax


def rough_taikei():
    bunya = [nm for ch, nm, _se in SO.BUNYA if ch != "第1章"]
    fig, ax = _fig(7.4, 4.8)
    _box(ax, 1.2, 3, 2.0, 1.2, "基本理念（協議会で設定）", fc=FG.BLUE, color="white", bold=True, wrap=7)
    _box(ax, 3.6, 3, 2.0, 1.2, "基本目標（協議会で設定）", fc=FG.PALE, wrap=7)
    ys = [5.6 - i * 0.66 for i in range(len(bunya))]
    for y, nm in zip(ys, bunya):
        _box(ax, 6.6, y, 3.2, 0.52, nm, sz=7, wrap=26)
        _line(ax, 4.6, 3, 5.0, y)
    _line(ax, 2.2, 3, 2.6, 3)
    ax.set_ylim(-0.5, 6)
    ax.text(6.6, -0.35, "第2部　分野別施策（9分野）", ha="center", fontsize=7.5, color=FG.NAVY)
    _box(ax, 9.25, 3, 1.3, 4.6, "第3部 成果目標（8分野）・見込量", fc="#E6E6E6", sz=7.5, wrap=5)
    fig.tight_layout()
    return FG._save(fig, "f90_shogai_rough_taikei")


def rough_kyoten(rows):
    fig, ax = _fig()
    _box(ax, 5, 3, 2.6, 1.2, "地域生活支援拠点等（田村地方）", fc=FG.BLUE, color="white", bold=True, wrap=10)
    pos = [(5, 5.3), (8.3, 3.9), (7.4, 0.9), (2.6, 0.9), (1.7, 3.9)]
    for (x, y), r in zip(pos, rows[1:]):
        _line(ax, 5, 3, x, y)
        _box(ax, x, y, 2.9, 1.2, "%s：%s" % (r[0], r[1]), sz=7, wrap=18)
    fig.tight_layout()
    return FG._save(fig, "f91_shogai_rough_kyoten")


def rough_renkei(rows):
    fig, ax = _fig(7.4, 4.8)
    _box(ax, 5, 3, 2.6, 1.0, "小野町（健康福祉課）・小野町自立支援協議会", fc=FG.BLUE, color="white", bold=True, wrap=12)
    bun = []
    for r in rows[1:]:
        if r[0]:
            bun.append([r[0], r[1]])
        else:
            bun[-1][1] += "、" + r[1]
    import math
    for k, (b, kikan) in enumerate(bun):
        a = math.pi / 2 - 2 * math.pi * k / len(bun)
        x, y = 5 + 3.9 * math.cos(a), 3 + 2.35 * math.sin(a)
        _line(ax, 5, 3, x, y)
        _box(ax, x, y, 2.4, 0.95, "【%s】%s" % (b, re.sub(r"（[^）]*）", "", kikan)[:30]), sz=6.5, wrap=16)
    fig.tight_layout()
    return FG._save(fig, "f92_shogai_rough_renkei")


def rough_life(rows):
    fig, ax = _fig(7.6, 4.0)
    ax.set_ylim(-0.6, 6.6)
    ax.annotate("", xy=(9.8, 3), xytext=(0.2, 3), arrowprops=dict(arrowstyle="->", color=FG.NAVY, lw=1.5))
    xs = [1.0, 2.6, 4.0, 5.4, 6.8, 8.6]
    for x, r in zip(xs, rows[1:]):
        ax.plot([x], [3], "o", color=FG.RED, ms=6)
        ax.text(x, 3.35, r[0], ha="center", fontsize=8, color=FG.RED, fontweight="bold")
        ax.text(x, 3.75, "\n".join(textwrap.wrap(r[1][:63], 9)), ha="center", va="bottom", fontsize=6.0, color=FG.NAVY)
        ax.text(x, 2.6, "\n".join(textwrap.wrap(r[2][:54], 9)), ha="center", va="top", fontsize=6.0, color=FG.BLUE)
    ax.text(0.2, 5.8, "制度上の内容", fontsize=7.5, color=FG.NAVY, fontweight="bold")
    ax.text(0.2, 0.15, "本計画での扱い", fontsize=7.5, color=FG.BLUE, fontweight="bold")
    fig.tight_layout()
    return FG._save(fig, "f93_shogai_rough_life")


def rough_ichi(rows):
    fig, ax = _fig()
    _box(ax, 5, 3, 3.2, 1.3, "おのまち障がい者計画・第4期障がい児福祉計画・第8期障がい福祉計画",
         fc=FG.BLUE, color="white", bold=True, wrap=13)
    _box(ax, 5, 5.3, 4.2, 0.8, rows[1][0] + "（上位）", fc="#E6E6E6", wrap=30)
    _box(ax, 5, 4.35, 4.2, 0.6, rows[2][0] + "（上位）", fc="#E6E6E6", wrap=30)
    _line(ax, 5, 4.05, 5, 3.65)
    rel = [re.sub(r"【[^】]*】", "", r[0]) for r in rows[3:]]
    pos = [(1.4, 1.0), (3.8, 0.7), (6.2, 0.7), (8.6, 1.0)]
    for (x, y), nm in zip(pos, rel):
        _line(ax, 5, 3, x, y)
        _box(ax, x, y, 2.3, 0.9, nm, sz=7, wrap=13)
    _box(ax, 1.3, 4.6, 2.2, 1.0, "国：障害者基本計画（第5次）・基本指針", fc=FG.PALE, sz=7, wrap=11)
    _box(ax, 8.7, 4.6, 2.2, 1.0, "福島県：障がい者計画・障がい福祉計画", fc=FG.PALE, sz=7, wrap=11)
    _line(ax, 2.4, 4.4, 3.4, 3.4)
    _line(ax, 7.6, 4.4, 6.6, 3.4)
    fig.tight_layout()
    return FG._save(fig, "f94_shogai_rough_ichi")


def main():
    ng = []
    T = dodai_tables()
    need = ["地域生活支援拠点等の5つの機能", "関係機関と連携の内容", "年齢到達に伴うサービスの移行", "町の関連計画との連携"]
    for k in need:
        if k not in T:
            ng.append("他メンバー素案に表「%s」がない" % k)
    if ng:
        print("\n".join(ng))
        return 1
    ill = [
        ("施策体系図", rough_taikei(), "土台 第2部第1章（計画の基本的な考え方）／当方 第2部第1章3", "1頁",
         "基本理念・基本目標・第2部の9分野・第3部の成果目標の対応", "当方 第2部の章と分野",
         "基本理念・基本目標は協議会で設定（11月下旬の第2回で協議）。設定後に描き起こす"),
        ("地域生活支援拠点等の5つの機能", rough_kyoten(T["地域生活支援拠点等の5つの機能"]),
         "土台 第3部第5章5（地域生活支援拠点等の5つの機能の表の前）", "半頁",
         "拠点等を中心に、5つの機能と本町での確保の方法", "土台 地域生活支援拠点等の5つの機能", "素材はそろう"),
        ("関係機関との連携", rough_renkei(T["関係機関と連携の内容"]),
         "土台 第3部第9章2（関係機関と連携の内容の表の前）", "半頁",
         "町と自立支援協議会を中心に、分野ごとの関係機関", "土台 関係機関と連携の内容", "素材はそろう"),
        ("ライフステージと制度の移行", rough_life(T["年齢到達に伴うサービスの移行"]),
         "土台 第3部第9章1（年齢到達に伴うサービスの移行の表の前）", "半頁",
         "6歳・18歳・20歳・40歳・65歳の制度の移行と、本計画での扱い", "土台 年齢到達に伴うサービスの移行", "素材はそろう"),
        ("計画の位置づけ", rough_ichi(T["町の関連計画との連携"]),
         "土台 第1部第1章2（現在の図1・図2を描き起こす）", "半頁",
         "上位計画（総合計画・地域福祉計画）、関連計画、国・県の計画との関係", "土台 町の関連計画との連携",
         "こども計画・健康増進計画の正式名称が町の確認待ち"),
    ]
    r = S.Report("障がい計画　イラスト起こしの割付案")
    r.cover(["おのまち障がい者計画・第4期障がい児福祉計画・第8期障がい福祉計画"],
            ["イラスト起こしの割付案", ASOF_JP, "ビズアップ公共コンサルティング株式会社"])
    r.toc_here()
    r.body_here()
    r.h1("1　割付の考え方")
    r.p("仕様書の成果品②は「原稿作成、レイアウトデザイン（イラスト起こし数点）、リライト作業を含む」としています。"
        "図にすると読みやすくなる5点を候補とし、置く場所・大きさ・描く要素を次のとおり割り付けます。"
        "**町と4点程度を選んだうえで、ラフをもとに描き起こします。**")
    r.tbl([["№", "イラスト", "置く場所", "大きさ", "描く要素", "状態"]]
          + [[str(i + 1), a, c, d, e, g] for i, (a, _p, c, d, e, _f, g) in enumerate(ill)],
          widths=[0.8, 3.0, 4.6, 1.2, 4.4, 3.2], fs=8, caption="イラスト起こしの割付")
    r.src("置く場所は他メンバー素案（変更履歴付き 第2弾）の章・節。配色は計画書の図と同じ（濃紺・青・淡青・灰）")
    r.h1("2　ラフ")
    for i, (a, png, c, d, e, f, g) in enumerate(ill):
        r.fig(png, "%s（ラフ）" % a, width=15.0, src="元にする表：%s。文言は表から読んだもので、描き起こしの際に短くする" % f)
    r.h1("3　進め方")
    r.tbl([["時期", "作業"],
           ["10月下旬", "町と候補5点から4点程度を選ぶ"],
           ["11月", "ラフを描き起こし、素案（第6版）に差し込む。施策体系図は第2回協議会の後に作る"],
           ["1月", "素案の協議・パブリックコメントでの見え方を確かめ、計画書の版下に反映する"]],
          widths=[3.0, 14.2], caption="進め方")
    p = OUT / f"小野町_障がい計画_イラスト起こしの割付案_{ASOF}.docx"
    r.build_toc()
    r.save(p)
    import docx
    t = "\n".join(x.text for x in docx.Document(str(p)).paragraphs)
    for w in ("北塩原", "大雪", "福島町", "訂正", "お詫び", "おわび"):
        if w in t:
            ng.append("「%s」が出ている" % w)
    print("■ イラスト起こしの割付案")
    print("   ", p.relative_to(HERE))
    print("\n■ 自己点検")
    for x in ng:
        print("   NG", x)
    print("   すべて適合" if not ng else "   **%d件**" % len(ng))
    return 1 if ng else 0


if __name__ == "__main__":
    sys.exit(main())
