# -*- coding: utf-8 -*-
"""計画素案（output/06_…計画素案.docx）の頁数を章ごとに見積もり、第9期の実績と対比する。

   レンダリングができない環境のため、build_soan_docx.js が指定している実際の
   レイアウト値（用紙・余白・字送り・行送り・表の行高・図の寸法）から高さを積み上げる。
   目視確認の代わりにはならないが、章ごとの相対的な厚みと合計の桁は把握できる。

   A4縦 11906×16838 twip／余白 上下1418・左右1134 → 本文 9638×14002 twip
"""
import json, math, os, struct, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BODY_W = 11906 - 1134 * 2          # 9638 twip
BODY_H = 16838 - 1418 * 2          # 14002 twip
FIGDIR = "/home/user/repository/output/figures"

def chars_per_line(pt):            # 全角は1文字＝フォントサイズと同じ幅
    return (BODY_W / 20) / pt

def png_size(path):
    with open(path, "rb") as f:
        b = f.read(24)
    return struct.unpack(">II", b[16:24])

def para_h(text, pt=10.5, line=300, after=120):
    n = max(1, math.ceil(len(text) / chars_per_line(pt)))
    return n * line + after

def block_h(b):
    t = b["t"]
    if t == "p":
        return para_h(b["v"])
    if t == "h3":
        return para_h(b["v"], after=100)
    if t == "bullets":
        return sum(para_h(x, after=60) for x in b["v"])
    if t == "note":
        return para_h(b["v"], pt=9.5, line=280, after=160) + 100
    if t in ("table", "kpi"):
        # 表：8.5pt・行送り240・セル余白上下60。列幅に応じて折り返す
        widths = b.get("widths")
        cols = len(b["head"])
        w = widths or [100 / cols] * cols
        tot = sum(w)
        h = 0
        for row in [b["head"]] + [list(map(str, r)) for r in b["rows"]]:
            lines = 1
            for i, cell in enumerate(row):
                cw = (w[i] / tot) * (BODY_W - 90 * 2 * cols) / 20 / 8.5
                lines = max(lines, math.ceil(len(str(cell)) / max(1, cw)))
            h += lines * 240 + 120
        return h + 160
    return 0

def run():
    import soan_content as S
    from figures_map import FIGS
    figs = {}
    for sec, fn, cap, src in FIGS:
        figs.setdefault(sec, []).append(fn)

    rows, tot_h, tot_t, tot_f = [], 0, 0, 0
    for ch in S.CH:
        h = 300 + 30 * 20                       # 章見出し
        nt = nf = nsec = 0
        for sec in ch["sections"]:
            nsec += 1
            h += 320 + 24 * 20 + 180            # 節見出し
            for fn in figs.get(f'{ch["no"]}|{sec["no"]}', []):
                nf += 1
                w, hh = png_size(os.path.join(FIGDIR, fn))
                h += 160 + (6.3 * hh / w) * 1440 + 60 + 18 * 20 + 40 + 15 * 20 + 200
            for b in sec["blocks"]:
                if b["t"] in ("table", "kpi"):
                    nt += 1
                h += block_h(b)
        pages = h / BODY_H
        rows.append((ch["no"], ch["title"], nsec, nt, nf, pages))
        tot_h += h; tot_t += nt; tot_f += nf
    return rows, tot_h / BODY_H, tot_t, tot_f

# ── 第9期の実績（doc01 §5-1／計画本体123頁・本文104頁の頁範囲） ──
K9 = [
 ("第1章", "計画の目的と位置づけ",        1,   3),
 ("第2章", "北塩原村の現状・取組み状況",   5,  40),
 ("第3章", "計画の基本理念と基本目標",    41,  48),
 ("第4章", "施策の展開",                49,  80),
 ("第5章", "介護保険料の設定",           81,  92),
 ("第6章", "計画の推進",                93,  94),
 ("資料編", "基本指針のポイント・用語集ほか", 95, 104),
]

# 前付（build_soan_docx.js が生成するもの）
FRONT = [("表紙", 1), ("本書の見方", 1), ("目次", 2)]

if __name__ == "__main__":
    rows, total, nt, nf = run()
    print("■ 第10期 計画素案の分量（build_soan_docx.js のレイアウト値から積算）\n")
    print(f"{'章':6s} {'見出し':26s} {'節':>3s} {'表':>4s} {'図':>3s} {'推定頁':>7s}")
    print("─" * 60)
    est = {}
    for no, title, nsec, t, f, p in rows:
        est[no] = p
        print(f"{no:6s} {title[:24]:26s} {nsec:3d} {t:4d} {f:3d} {p:7.1f}")
    print("─" * 60)
    print(f"{'合計':6s} {'':26s} {sum(r[2] for r in rows):3d} {nt:4d} {nf:3d} {total:7.1f}")

    # 章ごとに改頁が入るため、章単位で切り上げたものが実際の丁合いに近い
    body = sum(math.ceil(r[5]) for r in rows)
    front = sum(n for _, n in FRONT)
    print(f"\n  章単位で切り上げ（章の変わり目で改頁）　本文 {body}頁")
    print(f"  前付（{'・'.join(n for n, _ in FRONT)}）　　　　　　 {front}頁")
    print(f"  計画書全体　　　　　　　　　　　　　　 {body + front}頁")

    print("\n■ 第9期の実績との対比（第9期＝本文104頁）\n")
    print(f"{'章':6s} {'第9期':>7s} {'第10期':>8s} {'差':>7s}  {'構成比 第9期→第10期':>0s}")
    print("─" * 62)
    k9tot = sum(e - s + 1 for _, _, s, e in K9)
    for (no, title, s, e), (no2, _, _, _, _, p) in zip(K9, rows):
        k = e - s + 1
        print(f"{no:6s} {k:7d} {p:8.1f} {p-k:+7.1f}   {k/k9tot:5.1%} → {p/total:5.1%}")
    print("─" * 62)
    print(f"{'合計':6s} {k9tot:7d} {total:8.1f} {total-k9tot:+7.1f}")
    print(f"\n仕様書5①「A4判・両面約100頁」に対し、推定 本文{body}頁＋前付{front}頁＝{body+front}頁")
    print(f"第9期は本文104頁（章の頁範囲の積算は{k9tot}頁。差1頁は第1章と第2章の間の白頁）、"
          f"計画本体は表紙・目次等を含め123頁")
    print(f"→ 仕様書の約100頁に対し {100-(body+front):+d}頁、第9期の計画本体123頁に対し {123-(body+front):+d}頁")

    # ── アンケート調査報告書（doc24 §1-1 の改訂案）との対比 ──
    print("\n\n■ アンケート調査報告書　第9期の実績との対比\n")
    Q_NEEDS, Q_HOME = 80, 26           # doc25 集計仕様書による設問数
    PLANS = [("第9期 実績",        None, None, 123, None),
             ("第10期 当初案",       40,  20, 118,  25),
             ("第10期 100頁案",      33,  14, 100,  22)]
    print(f"{'':16s} {'集計編':>8s} {'分析編':>8s} {'資料編':>8s} {'全体':>6s} {'頁/設問':>9s}")
    print("─" * 62)
    for nm, c3, c4, tot, siryo in PLANS:
        if c3 is None:
            # 第9期の構成は Ⅰ調査の概要／Ⅱニーズ調査結果／Ⅲ在宅調査結果（doc01 §5-5）。
            # 分析編・資料編にあたる部分はなく、概要を6頁とみて残りを集計編とみなす
            shukei, bunseki, si = tot - 6, 0, 0
        else:
            shukei, bunseki, si = c3 + c4, 25, siryo
        per = shukei / (Q_NEEDS + Q_HOME)
        print(f"{nm:16s} {shukei:8d} {bunseki:8d} {si:8d} {tot:6d} {per:8.2f}頁"
              f"  （1頁{1/per:.1f}設問）")
    print("─" * 62)
    print("※ 第9期の報告書は doc01 §5-5 に記録された構成に分析編・資料編がない。")
    print("※ 設問数は第10期の80＋26＝106設問（doc25）。第9期も同程度とみなしている。")
