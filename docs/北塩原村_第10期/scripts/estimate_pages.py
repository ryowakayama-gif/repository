# -*- coding: utf-8 -*-
"""計画素案（output/06_…計画素案.docx）の頁数を章ごとに見積もり、第9期の実績と対比する。

   レンダリングができない環境のため、build_soan_docx.js が指定している実際の
   レイアウト値（用紙・余白・字送り・行送り・表の行高・図の寸法）から高さを積み上げる。
   目視確認の代わりにはならないが、章ごとの相対的な厚みと合計の桁は把握できる。

   A4縦 11906×16838 twip／余白 上下1418・左右1134 → 本文 9638×14002 twip
"""
import sys
sys.dont_write_bytecode = True   # 古い .pyc で古い成果品ができるのを防ぐ
import json, math, os, struct, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BODY_W = 11906 - 1134 * 2          # 9638 twip
BODY_H = 16838 - 1418 * 2          # 14002 twip
FIGDIR = "/home/user/repository/output/figures"

def chars_per_line(pt):            # 全角は1文字＝フォントサイズと同じ幅
    return (BODY_W / 20) / pt


# 文字の幅（全角を1とした割合）。游ゴシックの実測ではなく、和文フォントの
# 一般的な作りによる概算である。半角の英数字・記号は全角の約半分の幅を占める。
# 表のセルは数字が多く、すべてを全角1文字と数えると行数を1.5〜2倍に見積もる。
HANKAKU = 0.5      # 0-9 A-Z a-z および半角の記号・空白
ZENKAKU = 1.0      # 仮名・漢字・全角の記号
NARROW = 0.55      # 半角の約物のうち幅の狭いもの（. , : ; ! | ' " ( ) [ ] / -）
_NARROW_CH = set(".,:;!|'\"()[]{}/-+<>=*")


def text_w(t):
    """文字列の幅を全角何文字ぶんかで返す。

    len() で数えると、数字ばかりの表のセルを実際の2倍近くに見積もる。
    紙面の推定はこの幅から行数を出すため、ここがずれると頁数もずれる。
    """
    w = 0.0
    for ch in str(t):
        o = ord(ch)
        if o < 0x80:
            w += NARROW if ch in _NARROW_CH else HANKAKU
        elif 0xFF61 <= o <= 0xFF9F:      # 半角カタカナ
            w += HANKAKU
        else:
            w += ZENKAKU
    return w

def png_size(path):
    with open(path, "rb") as f:
        b = f.read(24)
    return struct.unpack(">II", b[16:24])

# 図に使える高さの上限（インチ）。build_soan_docx.js／build_shiryo_docx.js の
# MAX_H_IN と同じ値でなければ頁数の積算が狂う。
MAX_FIG_H = 8.9
SH_MAX_FIG_H = 9.3

def fig_in(path, w_in=6.3, max_h=None):
    """図の寸法（インチ）。縦長の図は幅を縮めて1頁に収める。

    docx 側と同じ縮め方をしないと、推定と実際の紙面が食い違う。
    """
    w, h = png_size(path)
    h_in = w_in * h / w
    lim = MAX_FIG_H if max_h is None else max_h
    if h_in > lim:
        h_in = lim
        w_in = h_in * w / h
    return w_in, h_in

def para_h(text, pt=10.5, line=300, after=120):
    n = max(1, math.ceil(text_w(text) / chars_per_line(pt)))
    return n * line + after

def block_h(b):
    # 文字の大きさは build_soan_docx.js と揃える。ずれると頁数の積算が狂う
    t = b["t"]
    if t == "p":
        return para_h(b["v"])                                   # 10.5pt
    if t == "h3":
        return para_h(b["v"], pt=11, after=100)                 # 11pt
    if t == "bullets":
        return sum(para_h(x, pt=10, after=60) for x in b["v"])   # 10pt
    if t == "note":
        return para_h(b["v"], pt=8.5, line=280, after=160) + 100  # 8.5pt
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
                # 列幅から引く。セル余白（左右90twip）はその列のぶんだけ引く
                cw = ((w[i] / tot) * BODY_W - 90 * 2) / 20 / 8.5
                lines = max(lines, math.ceil(text_w(cell) / max(1.0, cw)))
            h += lines * 240 + 120
        return h + 160
    return 0

def run():
    import soan_content as S
    from figures_map import FIGS
    figs = {}
    for e in FIGS:                    # inline の図も同じ節に積む（位置は高さに効かない）
        figs.setdefault(e[0], []).append(e[1])

    rows, tot_h, tot_t, tot_f = [], 0, 0, 0
    for ch in S.CH:
        h = 300 + 30 * 20                       # 章見出し
        nt = nf = nsec = 0
        for sec in ch["sections"]:
            nsec += 1
            h += 320 + 25 * 20 + 180            # 節見出し 12.5pt
            for fn in figs.get(f'{ch["no"]}|{sec["no"]}', []):
                nf += 1
                _, h_in = fig_in(os.path.join(FIGDIR, fn))
                h += 160 + h_in * 1440 + 60 + 18 * 20 + 40 + 16 * 20 + 200
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


# ══════ 第2回策定委員会 資料（build_shiryo_docx.js のレイアウト値） ══════
# A4縦 11906×16838／余白 上下左右とも1134 → 本文 9638×14570 twip
SH_W, SH_H = 11906 - 1134 * 2, 16838 - 1134 * 2
SH_FIGDIR = "/home/user/repository/output/figures"


def shiryo_block_h(b):
    def para(text, pt=10.5, line=300, after=120):
        n = max(1, math.ceil(text_w(text) / ((SH_W / 20) / pt)))
        return n * line + after
    t = b["t"]
    if t == "p":
        return para(b["v"])
    if t == "h3":
        return para(b["v"], pt=11, after=100)
    if t == "key":
        return para(b["v"], line=300, after=160) + 200
    if t == "note":
        return para(b["v"], pt=8.5, line=280, after=160) + 100
    if t == "bullets":
        return sum(para(x, pt=10, after=60) for x in b["v"])
    if t == "fig":
        _, h_in = fig_in(os.path.join(SH_FIGDIR, b["file"]),
                         b.get("width", 6.3), SH_MAX_FIG_H)
        return 160 + h_in * 1440 + 60 + 18 * 20 + 40 + 200
    if t == "table":
        cols = len(b["head"])
        wid = b.get("widths") or [100 / cols] * cols
        tot = sum(wid)
        h = 0
        for row in [b["head"]] + [list(map(str, r)) for r in b["rows"]]:
            lines = 1
            for i, cell in enumerate(row):
                cw = ((wid[i] / tot) * SH_W - 90 * 2) / 20 / 8.5
                lines = max(lines, math.ceil(text_w(cell) / max(1.0, cw)))
            h += lines * 240 + 120
        return h + 160
    return 0


def shiryo():
    import shiryo_content as SH
    rows = []
    for c in SH.CH:
        h = 300 + 30 * 20
        nt = nf = 0
        for sec in c["sections"]:
            h += 320 + 24 * 20 + 180
            for b in sec["blocks"]:
                nt += b["t"] == "table"
                nf += b["t"] == "fig"
                h += shiryo_block_h(b)
        rows.append((c["no"], c["title"], len(c["sections"]), nt, nf, h / SH_H))
    return rows


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
    # ここまでは要素の高さを足しただけで、送りによる空白を見ていない下限値である。
    # 実際の紙面は estimate_layout.py が頁の境目まで追って出す。
    try:
        import estimate_layout as _EL
        n_soan = _EL.pages(_EL.soan_items(False), _EL.BODY_H)
        n_kei = _EL.pages(_EL.soan_items(True), _EL.BODY_H)
        print(f"\n  ※ 上は要素の高さを足しただけの**下限値**である。"
              f"見出しや図が次頁へ送られて生じる空白を見ていない。")
        print(f"  　 送りを追った推定（estimate_layout.py）"
              f"　素案のまま {n_soan}頁／計画書として {n_kei}頁（＋前付{front}頁）")
        print(f"  　 計画書は編集注記と5-12（据え置いた項目）を載せないため"
              f"{n_soan - n_kei}頁短くなる")
    except Exception as e:
        print(f"\n  ※ 送りを追った推定を出せない（{e}）")

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
             ("第10期 案A（不採用）",  33,  14, 100,  22),
             ("第10期 採用案（B）",   42,  21, 100,   6)]
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
    print("※ 採用案Bは調査票の再掲・単純集計表・自由記述全文を電子媒体に移し、"
          "浮いた頁を集計編に戻したもの（doc24§1-1）。")

    print("\n\n■ 第2回策定委員会 資料\n")
    sr = shiryo()
    print(f"{'資料':6s} {'見出し':34s} {'節':>3s} {'表':>4s} {'図':>3s} {'推定頁':>7s}")
    print("─" * 62)
    for no, ti, ns, nt, nf, pg in sr:
        print(f"{no:6s} {ti[:32]:34s} {ns:3d} {nt:4d} {nf:3d} {pg:7.1f}")
    print("─" * 62)
    body = sum(math.ceil(r[5]) for r in sr)
    print(f"{'合計':6s} {'':34s} {sum(r[2] for r in sr):3d} {sum(r[3] for r in sr):4d} "
          f"{sum(r[4] for r in sr):3d} {sum(r[5] for r in sr):7.1f}")
    nt_all = sum(r[3] for r in sr)
    print(f"\n  資料単位で切り上げ　{body}頁")
    print(f"  ＋ 資料1（次第・委員名簿・座席表）3頁 ＝ 全体 約{body + 3}頁")
    print(f"\n  ※ この積算は下限値である。資料単位でしか切り上げておらず、節の途中で表が"
          f"入りきらずに次頁へ送られる余りを見ていない。表{nt_all}点を抱えるため実際は"
          "これより増える。doc26 の手計算（資料2〜7で48頁）との差はこの余りにあたる。")
