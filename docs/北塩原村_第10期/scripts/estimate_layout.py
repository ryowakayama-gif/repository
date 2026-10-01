# -*- coding: utf-8 -*-
"""紙面の点検（改頁のまたぎ・図の送り・見出しの取り残し）を機械で代替する.

【なぜ必要か】
  この環境では LibreOffice が動かず docx を PDF に変換できない。紙面の目視確認が
  できないため、代わりに build_soan_docx.js が指定しているレイアウト値から
  要素を上から積み上げ、頁の境目に何が来るかを突き止める。

  **推定は目視の代わりにはならない。** 字幅は全角1文字＝フォントサイズとみなした
  概算であり、Word の実際の組版（禁則処理・和欧混植・表の自動列幅）とは食い違う。
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
  python3 scripts/estimate_layout.py            計画素案
  python3 scripts/estimate_layout.py --shiryo   第2回策定委員会資料
"""
import sys
sys.dont_write_bytecode = True
import math, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import estimate_pages as EP

BODY_H = EP.BODY_H                 # 14002 twip（素案）
SH_H = EP.SH_H                     # 14570 twip（委員会資料）
AKI = 0.30                         # 空白が頁のこの割合を超えたら挙げる


def flow(items, page_h):
    """要素を上から積み上げ、頁の境目で起きることを拾う。

    items は (種別, 名, 高さ, 分割できるか, 次と離さないか) の並び。
    返すのは (頁数, 気になる箇所の一覧)。
    """
    y, page, out = 0, 1, []
    i = 0
    while i < len(items):
        kind, name, h, splittable, keep = items[i]
        if kind == "改頁":
            if y > 0:
                out.append(("改頁", page, name, page_h - y))
            page += 1
            y = 0
            i += 1
            continue
        # keepNext の連なりをひとまとめにして高さを見る
        grp, j = h, i
        while keep and j + 1 < len(items) and items[j + 1][0] != "改頁":
            j += 1
            grp += items[j][2]
            if not items[j][4]:
                break
        nokori = page_h - y
        if h > page_h:
            # 1頁に収まらない要素。分割できるものは収まるところまで入れる
            if splittable:
                out.append(("頁超", page, name, h))
                y = (y + h) % page_h
                page += int((y + h) // page_h) if False else 0
                page += max(1, int(h // page_h))
            else:
                out.append(("頁超", page, name, h))
                page += 1
                y = 0
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


def soan_items():
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
            it.append(("節見出し", f'{sec["no"]}　{sec.get("title", "")}',
                       320 + 25 * 20 + 180, False, True))
            for fn in figs.get(f'{ch["no"]}|{sec["no"]}', []):
                _, h_in = EP.fig_in(os.path.join(EP.FIGDIR, fn))
                # 画像＋表題＋出典。keepNext により1つのかたまりとして動く
                h = 160 + h_in * 1440 + 60 + 18 * 20 + 40 + 16 * 20 + 200
                it.append(("図", f'{sec["no"]}　{fn}', h, False, False))
            for b in sec["blocks"]:
                h = EP.block_h(b)
                if b["t"] in ("table", "kpi"):
                    nm = f'{sec["no"]}　表（{len(b["rows"])}行）{"／".join(map(str, b["head"]))[:34]}'
                    it.append(("表", nm, h + 160, True, False))
                elif b["t"] == "fig":
                    from figures_map import FIGS as _F
                    fn = next((e[1] for e in _F if e[0] == b["v"] or e[1].startswith(b["v"])),
                              None)
                    if fn:
                        _, h_in = EP.fig_in(os.path.join(EP.FIGDIR, fn))
                        h = 160 + h_in * 1440 + 60 + 18 * 20 + 40 + 16 * 20 + 200
                    it.append(("図", f'{sec["no"]}　{b["v"]}', h, False, False))
                elif b["t"] == "h3":
                    it.append(("小見出し", f'{sec["no"]}　{b["v"][:34]}', h, False, True))
                else:
                    it.append((b["t"], f'{sec["no"]}　{str(b.get("v"))[:34]}', h, True, False))
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
                    it.append(("表", nm, h + 160, True, False))
                elif b["t"] == "fig":
                    it.append(("図", f'{sec["no"]}　{b["file"]}', h, False, False))
                elif b["t"] == "h3":
                    it.append(("小見出し", f'{sec["no"]}　{b["v"][:34]}', h, False, True))
                else:
                    it.append((b["t"], f'{sec["no"]}　{str(b.get("v"))[:34]}', h, True, False))
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


if __name__ == "__main__":
    if "--shiryo" in sys.argv:
        report("第2回策定委員会資料", shiryo_items(), SH_H)
    else:
        report("計画素案（本文）", soan_items(), BODY_H)
    print("\n※ この推定は目視確認の代わりにはならない。字幅は全角1文字＝フォントサイズ"
          "とみなした概算で、Word の禁則処理・和欧混植・表の自動列幅とは食い違う。")
    print("※ 表の行は cantSplit により行の途中では割れない。見出し行は tableHeader に"
          "より次頁の先頭で繰り返される。")
    print("※ 見出しと図は keepNext により次の要素と同じ頁へ送られる。"
          "「見出しの送り」「送り」はその結果生じる空白である。")
