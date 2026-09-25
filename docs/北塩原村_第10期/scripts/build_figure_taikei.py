# -*- coding: utf-8 -*-
"""施策体系図（基本理念→基本目標5→施策28）の作成。

   施策の一覧は soan_content から読むため、素案を直せば図も追随する。
   成果品はモノクロ印刷（仕様書5①）のため、区分は色相ではなく
   明度と枠線の太さで識別する。
   出力: output/figures/fig3-1_施策体系図.png（300dpi）
"""
import os, sys, textwrap
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
import soan_content as S

OUT = "/home/user/repository/output/figures/fig3-1_施策体系図.png"
os.makedirs(os.path.dirname(OUT), exist_ok=True)
plt.rcParams.update({"font.family": "IPAGothic", "figure.facecolor": "white"})

# 区分ごとの塗りと枠（モノクロ印刷で分離する3段階）
KUBUN = {
 "継続":    dict(fc="#FFFFFF", ec="#595959", lw=0.8),
 "継続・強化": dict(fc="#D9D9D9", ec="#404040", lw=1.0),
 "強化":    dict(fc="#D9D9D9", ec="#404040", lw=1.0),
 "改称":    dict(fc="#FFFFFF", ec="#1A1A1A", lw=1.2, ls=(0, (2.6, 1.6))),
 "新規":    dict(fc="#A6A6A6", ec="#1A1A1A", lw=1.6),
 "新設":    dict(fc="#A6A6A6", ec="#1A1A1A", lw=1.6),
}
NAVY = "#1F3864"


def collect():
    """第3章 3-2 から基本理念と基本目標、第4章から施策を取り出す"""
    ch = {c["no"]: c for c in S.CH}
    rinen = []
    for sec in ch["第3章"]["sections"]:
        if sec["no"] == "3-1":
            rinen = [b["v"] for b in sec["blocks"] if b["t"] == "h3"]
    assert len(rinen) == 2, rinen
    mokuhyo = []
    for sec in ch["第3章"]["sections"]:
        if sec["no"] == "3-2":
            for b in sec["blocks"]:
                if b["t"] == "table":
                    mokuhyo = [(r[0], r[1].replace("【新設】", "")) for r in b["rows"]]
    sisaku = {}
    for sec in ch["第4章"]["sections"]:
        if sec["no"].startswith("施策"):
            no = sec["no"].replace("施策", "")
            ti = sec["title"]
            sisaku.setdefault(no[0], []).append((no, ti.split("　【")[0], ti.split("【")[1].rstrip("】")))
        elif sec["no"] == "4-5":
            for b in sec["blocks"]:
                if b["t"] == "table":
                    for r in b["rows"]:
                        n, nm = r[0].split(" ", 1)
                        sisaku.setdefault(n[0], []).append((n, nm, "新設"))
    return rinen, mokuhyo, sisaku


def wrap(t, w):
    return "\n".join(textwrap.wrap(t, w)) or t


def main():
    rinen, mokuhyo, sisaku = collect()
    assert len(mokuhyo) == 5, mokuhyo
    assert sum(len(v) for v in sisaku.values()) == 28, {k: len(v) for k, v in sisaku.items()}

    PER = 4                       # 1行に並べる施策の数
    rows = [(no, nm, sisaku[no]) for no, nm in mokuhyo]
    # 各基本目標の高さ（施策の段数で決まる）
    heights = [max(1, -(-len(s) // PER)) for _, _, s in rows]
    H = 1.50 + sum(h * 0.86 + 0.20 for h in heights)
    fig = plt.figure(figsize=(6.3, H))
    ax = fig.add_axes([0, 0, 1, 1]); ax.set_xlim(0, 100); ax.set_ylim(0, H * 10)
    ax.axis("off")
    Y = H * 10

    def box(x, y, w, h, text, size, fc, ec, lw, bold=False, color="#262626", ls="-"):
        ax.add_patch(FancyBboxPatch((x, y - h), w, h,
                     boxstyle=f"round,pad=0,rounding_size={min(w, h) * 0.12}",
                     fc=fc, ec=ec, lw=lw, ls=ls, zorder=2))
        ax.text(x + w / 2, y - h / 2, text, ha="center", va="center", fontsize=size,
                fontweight="bold" if bold else "normal", color=color, zorder=3,
                linespacing=1.25)

    # ── 基本理念 ──
    y = Y - 1.0
    box(2, y, 96, 8.6, "【基本理念】\n" + "\n".join(wrap(r, 40) for r in rinen),
        8.0, "#1F3864", "#1F3864", 1.0, True, "white")
    y -= 8.6
    ax.annotate("", xy=(50, y - 2.2), xytext=(50, y),
                arrowprops=dict(arrowstyle="-|>", color=NAVY, lw=1.4))
    y -= 2.8

    # ── 基本目標と施策 ──
    LW, GAP = 20.0, 1.2           # 基本目標欄の幅
    for (no, nm, ss), nline in zip(rows, heights):
        blk = nline * 8.6
        box(2, y, LW, blk, f"基本目標{no}\n" + wrap(nm, 9), 7.6, "#DDEBF7", NAVY, 1.0, True, NAVY)
        x0 = 2 + LW + 2.0
        w = (98 - x0 - (PER - 1) * GAP) / PER
        for i, (sno, snm, kb) in enumerate(ss):
            r, c = divmod(i, PER)
            st = KUBUN[kb]
            box(x0 + c * (w + GAP), y - r * 8.6, w, 7.6,
                f"{sno}\n" + wrap(snm, 11) + f"\n〔{kb}〕", 5.9, **st)
        y -= blk + 2.0

    # ── 凡例 ──
    y -= 0.4
    ax.text(2, y - 1.6, "区分", fontsize=7, fontweight="bold", color=NAVY, va="center")
    for i, (kb, lab) in enumerate((("継続", "継続"), ("強化", "強化・継続強化"),
                                   ("改称", "改称"), ("新規", "新規・新設"))):
        x = 10 + i * 22.5
        ax.add_patch(FancyBboxPatch((x, y - 2.6), 3.4, 2.0,
                     boxstyle="round,pad=0,rounding_size=0.3", **KUBUN[kb], zorder=2))
        ax.text(x + 4.4, y - 1.6, lab, fontsize=6.8, va="center", color="#262626")
    fig.savefig(OUT, dpi=300, bbox_inches="tight", pad_inches=0.06, facecolor="white")
    plt.close(fig)
    n = sum(len(v) for v in sisaku.values())
    print(f"保存: {OUT}（基本目標{len(rows)}・施策{n}）")


if __name__ == "__main__":
    main()
