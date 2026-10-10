# -*- coding: utf-8 -*-
"""紙面の点検（改頁のまたぎ・図の送り・見出しの取り残し）を機械で代替する.

【なぜ必要か】
  build_soan_docx.js が指定しているレイアウト値から要素を上から積み上げ、頁の
  境目に何が来るかを突き止める。どの頁を見るべきかを絞るための道具である。

  令和8年10月10日に libreoffice-writer と poppler-utils を入れ、docx を PDF に
  変換して紙面を見られるようになった（scripts/render_check.py）。実紙面と突き
  合わせた結果、次のことが分かっている。
    ・用紙・余白・本文の高さ・行送り（本文300・表240 twip）・図の寸法は一致する
    ・表の幅は本文幅（9638）ではなく TBLW（9360）である ← 直した
    ・表の1行の下駄は宣言値120ではなく罫線を含めて127 twip である ← 直した
    ・送りによる空白は、この推定よりも実際のほうが大きい（約3頁ぶん）
  したがって**この推定はなお下限であり、目視の代わりにはならない。** 字幅は
  全角1・半角0.5を目安とした概算であり、Word の実際の組版（禁則処理・
  プロポーショナルな字送り・行末の追い込み）とは食い違う。なお表の列幅は docx 側で
  TableLayoutType.FIXED を指定したため、宣言した widths どおりに組まれる。
  この点検が示すのは「ここを見てください」という箇所であり、「問題がない」ことの
  証明ではない。

【何を見るか】
  1 図が頁をまたぐ   Word は図を分割しないため、またぐ図はまるごと次頁へ送られる。
                     前の頁の下に大きな空白が残る。
  2 表が頁をまたぐ   見出し行は tableHeader により繰り返されるが、2行目以降が
                     1行だけ次頁に残ると読みにくい（孤立行）。
  3 見出しの取り残し  keepNext を入れたため Word は見出しを次の要素と同じ頁に送る。
                     送られた結果どれだけ空白が残るかを見る。
  4 空白の大きさ      送りによって生じる空白が頁の3割を超える箇所を挙げる。

【使い方】
  python3 scripts/estimate_layout.py            計画素案（素案のまま）
  python3 scripts/estimate_layout.py --keikaku  計画書として（編集注記と5-12を除く）
  python3 scripts/estimate_layout.py --shiryo   第2回策定委員会資料

【estimate_pages.py との違い】
  estimate_pages.py は要素の高さを足して本文の高さで割る。送りによる空白を
  見ないため、**下限値**である。本スクリプトは上から積み上げて頁の境目を見るため、
  送りの空白を含む。両者の差は送りの空白にあたる（点検49で確かめている）。
"""
import sys
sys.dont_write_bytecode = True
import math, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import estimate_pages as EP

BODY_H = EP.BODY_H                 # 14002 twip（素案）
SH_H = EP.SH_H                     # 14570 twip（委員会資料）
AKI = 0.30                         # 空白が頁のこの割合を超えたら挙げる

# 実紙面との差（令和8年10月10日の実測）。
#   素案　　推定135頁 → 実測137頁（LibreOffice・IPAゴシックに寄せた置換）
#   計画書　推定118頁 → 実測120頁
# どちらも＋2頁である。残りの差は送り（keepNext）による空白で、この推定は
# Word よりも空白を小さく見る。したがって**推定は下限であり**、分量が収まるか
# を判断するときはこの差を乗せる。節ごとの対比では推定と実測は±2頁で追えている。
#   ※ 実測は游ゴシックが無い環境での置換（IPAゴシック）による。字幅は
#     数字0.5em・漢字1.0emで推定の前提と一致するが、Word＋游ゴシックでは
#     禁則処理と欧文の字送りの違いでなお±1〜2頁動き得る。
#   ※ 突き合わせは scripts/render_check.py で再現できる。
JITSU_SA = 2


def flow(items, page_h, trace=None):
    """要素を上から積み上げ、頁の境目で起きることを拾う。

    items は (種別, 名, 高さ, 分割できるか, 次と離さないか, 最初のひとかたまり)
    の並び。返すのは (頁数, 気になる箇所の一覧)。

    **keepNext は「次の要素の先頭」としか結びつかない。** Word の keepNext は
    見出しを次の段落（表なら最初の行）と同じ頁に置くもので、表や段落の全体を
    引き連れるわけではない。全体の高さで判定すると、長い表の前の見出しが
    いつまでも次頁へ送られ、ありもしない空白を数えてしまう。
    """
    y, page, out = 0, 1, []
    i = 0
    while i < len(items):
        kind, name, h, splittable, keep = items[i][:5]
        if trace is not None:
            trace.append((kind, name, page if y < page_h else page + 1))
        first = items[i][5] if len(items[i]) > 5 else h
        if kind == "改頁":
            if y > 0:
                out.append(("改頁", page, name, page_h - y))
            page += 1
            y = 0
            i += 1
            continue
        # keepNext の連なり。最後の要素は「分割できない最初のひとかたまり」で見る
        grp, j = h, i
        while keep and j + 1 < len(items) and items[j + 1][0] != "改頁":
            j += 1
            nxt = items[j]
            if nxt[4]:                       # 次も keepNext なら全体を連れていく
                grp += nxt[2]
            else:                            # 連なりの終わり。先頭のひとかたまりだけ
                grp += nxt[5] if len(nxt) > 5 else nxt[2]
                break
        nokori = page_h - y
        if h > page_h:
            # 1頁に収まらない要素。挙げたうえで、送りは下の分岐と同じ数え方にする。
            #
            # 以前はここで page += max(1, int(h // page_h)) としていた。1.2頁の表が
            # 頁の途中から始まると、残りは 1.2 − 残り で2頁にまたがるのに1頁しか
            # 進めず、素案では5件で3頁ぶん少なく数えていた（PDF の実測で判明）。
            # 表は cantSplit でも行の境目では割れるため、分割できるものは
            # 「残りに入るぶんを入れ、続きを次頁へ」でよい。
            out.append(("頁超", page, name, h))
            if splittable:
                n = h - nokori
                page += 1 + int(n // page_h)
                y = n % page_h
            else:
                if y > 0:
                    page += 1
                page += max(1, int(h // page_h)) - 1
                y = h % page_h
            i += 1
            continue
        if h > nokori:
            # 入りきらない
            if splittable:
                out.append(("またぎ", page, name, nokori))
                y = h - nokori
                page += 1
            else:
                out.append(("送り", page, name, nokori))
                y = h
                page += 1
        elif keep and grp > nokori and grp <= page_h:
            out.append(("見出しの送り", page, name, nokori))
            y = h
            page += 1
        else:
            y += h
        i += 1
    return page, out


# 本文の高さのこの割合を超える図は、docx 側が頁の頭から置く（改頁を入れる）。
# build_soan_docx.js の ZENMEN_H_IN と同じ値でなければ推定が実際とずれる。
ZENMEN = 0.60


def fig_tall(h_in, page_h_in=9.72):
    return h_in >= page_h_in * ZENMEN


def first_unit(b, page_h=None):
    """分割できない最初のひとかたまりの高さ。

    表は「見出し行＋最初の1行」（cantSplit により行の途中では割れず、
    見出し行は tableHeader により次頁の先頭で繰り返される）。
    段落・箇条書きは1行。図は分けられないため全体。
    """
    t = b.get("t")
    if t in ("table", "kpi"):
        return (240 + EP.ROW_GETA) * 2 + 160   # 見出し行＋1行＋行の下駄＋表の前後
    if t in ("p", "note", "bullets", "key"):
        return 300 + 120
    # 分けられないもの（図など）は全体が最初のひとかたまりになる。
    # 呼び出し側が高さを渡していないときは、1行ぶんを下限として返す。
    return 300 + 120


def soan_items(keikaku=False):
    """素案の要素を上から並べる。

    keikaku=True のときは、計画書として納める形にする。
    編集注記（note）と5-12（現時点で据え置いた項目）は策定の過程を委員と
    共有するために置いたもので、計画書には載せない（素案5-12の注記による）。
    """
    import soan_content as S
    from figures_map import FIGS
    figs = {}
    for e in FIGS:
        figs.setdefault(e[0], []).append(e[1])
    it = []
    for ci, ch in enumerate(S.CH):
        if ci > 0:
            it.append(("改頁", f'{ch["no"]} の前', 0, False, False))
        it.append(("章見出し", f'{ch["no"]}　{ch["title"]}', 300 + 30 * 20, False, True))
        for sec in ch["sections"]:
            if keikaku and sec["no"] == "5-12":
                continue
            it.append(("節見出し", f'{sec["no"]}　{sec.get("title", "")}',
                       320 + 25 * 20 + 180, False, True))
            for fn in figs.get(f'{ch["no"]}|{sec["no"]}', []):
                _, h_in = EP.fig_in(os.path.join(EP.FIGDIR, fn))
                # 画像＋表題＋出典。keepNext により1つのかたまりとして動く
                h = 160 + h_in * 1440 + 60 + 18 * 20 + 40 + 16 * 20 + 200
                if fig_tall(h_in):
                    it.append(("改頁", f'{sec["no"]}　{fn} の前（1頁の図版）',
                               0, False, False, 0))
                it.append(("図", f'{sec["no"]}　{fn}', h, False, False, h))
            for b in sec["blocks"]:
                if keikaku and b["t"] == "note":
                    continue
                h = EP.block_h(b)
                if b["t"] in ("table", "kpi"):
                    nm = f'{sec["no"]}　表（{len(b["rows"])}行）{"／".join(map(str, b["head"]))[:34]}'
                    it.append(("表", nm, h + 160, True, False, first_unit(b)))
                elif b["t"] == "fig":
                    from figures_map import FIGS as _F
                    fn = next((e[1] for e in _F if e[0] == b["v"] or e[1].startswith(b["v"])),
                              None)
                    if fn:
                        _, h_in = EP.fig_in(os.path.join(EP.FIGDIR, fn))
                        h = 160 + h_in * 1440 + 60 + 18 * 20 + 40 + 16 * 20 + 200
                        if fig_tall(h_in):
                            it.append(("改頁", f'{sec["no"]}　{b["v"]} の前（1頁の図版）',
                                       0, False, False, 0))
                    it.append(("図", f'{sec["no"]}　{b["v"]}', h, False, False, h))
                elif b["t"] == "h3":
                    it.append(("小見出し", f'{sec["no"]}　{b["v"][:34]}', h, False, True, h))
                else:
                    it.append((b["t"], f'{sec["no"]}　{str(b.get("v"))[:34]}', h, True,
                               False, first_unit(b)))
    return it


def shiryo_items():
    import shiryo_content as SH
    it = []
    for ci, ch in enumerate(SH.CH):
        if ci > 0:
            it.append(("改頁", f'{ch["no"]} の前', 0, False, False))
        it.append(("資料見出し", f'{ch["no"]}　{ch["title"]}', 300 + 30 * 20, False, True))
        for sec in ch["sections"]:
            if not sec["no"].endswith("-0"):
                it.append(("節見出し", f'{sec["no"]}　{sec.get("title", "")}',
                           320 + 24 * 20 + 180, False, True))
            for b in sec["blocks"]:
                h = EP.shiryo_block_h(b)
                if b["t"] == "table":
                    nm = (f'{sec["no"]}　表（{len(b["rows"])}行）'
                          + "／".join(map(str, b["head"]))[:34])
                    it.append(("表", nm, h + 160, True, False, first_unit(b)))
                elif b["t"] == "fig":
                    _, hi = EP.fig_in(os.path.join(EP.SH_FIGDIR, b["file"]),
                                      b.get("width", 6.3), EP.SH_MAX_FIG_H)
                    if fig_tall(hi, 10.12):
                        it.append(("改頁", f'{sec["no"]}　{b["file"]} の前（1頁の図版）',
                                   0, False, False, 0))
                    it.append(("図", f'{sec["no"]}　{b["file"]}', h, False, False, h))
                elif b["t"] == "h3":
                    it.append(("小見出し", f'{sec["no"]}　{b["v"][:34]}', h, False, True, h))
                else:
                    it.append((b["t"], f'{sec["no"]}　{str(b.get("v"))[:34]}', h, True,
                               False, first_unit(b)))
    return it


def report(name, items, page_h):
    pages, out = flow(items, page_h)
    # 本文が頁をまたぐのは普通のことなので、空白の大きい箇所からは除く
    big = [o for o in out if o[0] != "改頁" and o[3] >= page_h * AKI
           and not (o[0] == "またぎ" and "表（" not in o[2])]
    print(f"■ {name}　推定 {pages}頁"
          f"（1頁 {page_h} twip ＝ 約{page_h / 300:.0f}行）\n")
    kinds = {}
    for k, pg, nm, v in out:
        kinds[k] = kinds.get(k, 0) + 1
    print("  種別ごとの件数")
    for k in ("またぎ", "送り", "見出しの送り", "頁超", "改頁"):
        if k in kinds:
            print(f"    {k:8s} {kinds[k]:4d}件")
    print(f"\n  空白が頁の{AKI:.0%}を超える箇所　{len(big)}件")
    if big:
        print(f"    {'頁':>4s} {'種別':10s} {'空白':>7s}  箇所")
        print("    " + "─" * 72)
        for k, pg, nm, v in big:
            print(f"    {pg:4d} {k:10s} {v / page_h:6.0%}  {nm[:58]}")
    cross = [o for o in out if o[0] == "またぎ" and "表（" in o[2]]
    print(f"\n  表が頁をまたぐ箇所　{len(cross)}件"
          f"（本文が頁をまたぐのは普通のことなので挙げていない）")
    if cross:
        print(f"    {'頁':>4s} {'残り':>7s}  箇所")
        print("    " + "─" * 72)
        for k, pg, nm, v in cross:
            print(f"    {pg:4d} {v / page_h:6.0%}  {nm[:62]}")
    over = [o for o in out if o[0] == "頁超"]
    print(f"\n  1頁に収まらない要素　{len(over)}件")
    for k, pg, nm, v in over:
        print(f"    {pg:4d} {v / page_h:6.1f}頁  {nm[:62]}")
    return pages, out


def pages(items, page_h):
    """頁数だけを返す（点検から呼ぶ）"""
    return flow(items, page_h)[0]


if __name__ == "__main__":
    if "--shiryo" in sys.argv:
        report("第2回策定委員会資料", shiryo_items(), SH_H)
    elif "--keikaku" in sys.argv:
        n1 = report("計画書（編集注記と5-12を除く）", soan_items(True), BODY_H)[0]
        n0 = pages(soan_items(False), BODY_H)
        print(f"\n  素案のまま {n0}頁 → 計画書として {n1}頁（差 {n0 - n1}頁）")
        print(f"  ＋前付（表紙・本書の見方・目次）4頁 ＝ 計画書全体 約{n1 + 4}頁")
        print(f"  仕様書5①「A4判・両面約100頁」に対し {100 - (n1 + 4):+d}頁")
    else:
        report("計画素案（本文）", soan_items(), BODY_H)
    print("\n※ この推定は目視確認の代わりにはならない。字幅は全角1・半角0.5を"
          "目安とした概算で、Word の禁則処理・プロポーショナルな字送り・"
          "行末の追い込みとは食い違う。")
    print("※ 表の行は cantSplit により行の途中では割れない。見出し行は tableHeader に"
          "より次頁の先頭で繰り返される。")
    print("※ 見出しと図は keepNext により次の要素と同じ頁へ送られる。"
          "「見出しの送り」「送り」はその結果生じる空白である。")
