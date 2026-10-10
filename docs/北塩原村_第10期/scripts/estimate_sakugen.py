# -*- coding: utf-8 -*-
"""計画書の分量を仕様書の目安に近づけるための削減案を試算する（WBS Ⅱ-113）

仕様書5①はA4判・両面約100頁を目安としている。現在の推定は本スクリプトが
毎回算出する（じか書きしない。素案を直すたびに動くため）。どこを削れば何頁減るかを、
estimate_layout.py と同じ積み上げで試算する。

【削ってよいものと、いけないもの】
  介護保険法第117条第2項が計画に定めることを求めている事項（サービスの種類ごとの
  量の見込み、地域支援事業の量、必要利用定員総数、給付適正化）は**削らない**。
  削る候補に挙げているのは、当方が算定の過程を示すために加えた
  「方法」「検証」「不確かさ」の記述である。第9期計画はこれらを載せておらず、
  第5章は12頁であった（第10期は36頁）。

  **どれを削るかは村の判断である。** 本スクリプトは、削った場合に何頁になるかを
  示すだけで、素案そのものは変えない。

  python3 scripts/estimate_sakugen.py
"""
import sys
sys.dont_write_bytecode = True
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import estimate_layout as EL
import soan_content as SC

# (番号, 名, [(節, 小見出し)…], 法定記載事項にあたるか, 理由と移し先)
AN = [
 ("A", "5-4 算定の方法の説明",
  [("5-4", "基準年度とデータの月数"), ("5-4", "利用率の分母"),
   ("5-4", "算定の対象とするサービス")],
  False,
  "どのデータをどう使ったかの説明。計画書の読み手に必要な事項ではなく、"
  "村が検算するための情報である。資料編または別冊の「算定の根拠」へ移す"),
 ("B", "5-4 推計の不確かさ（①単価の趨勢・②伸びの偏り・③推計の誤差）",
  [("5-4", "(10) 推計の不確かさ"), ("5-4", "① 単価の趨勢"),
   ("5-4", "② サービス別の伸びの偏り"), ("5-4", "③ 推計そのものの誤差"),
   ("5-4", "参考　受給者1人あたり利用日数・回数の実績")],
  False,
  "推計の確からしさを当方が検証した記録。第9期計画には対応する記載がない。"
  "別冊の「算定の根拠」へ移す"),
 ("C", "5-4 見込量の自己点検",
  [("5-4", "(6) 見込量の自己点検")],
  False,
  "当方が算定の整合を確かめた記録。別冊へ移す"),
 ("D", "5-4 施策の見込量への反映",
  [("5-4", "(11) 施策の見込量への反映")],
  False,
  "施策が見込量に効く筋道の説明。第4章の各施策と重なる。別冊へ移す"),
 ("E", "5-7 保険料の算定の過程（算定の構造・再現・感応度）",
  [("5-7", "算定の構造"), ("5-7", "算定の構造による第9期の再現"),
   ("5-7", "感応度")],
  False,
  "保険料がどう決まるかの式、第9期で検算した結果、前提が動いたときの振れ幅。"
  "委員会資料6に同じものがあり、計画書は結果（基準額とパターン）を示せば足りる。"
  "別冊へ移す"),
 ("F", "2-8 交付金の指標群ごとの明細",
  [("2-8", "指標群別にみると"), ("2-8", "目標別に全国平均と比べると"),
   ("2-8", "福島県の評価と重ねると"),
   ("2-8", "全国の多くの市町村が取得している項目のうち、未取得のもの")],
  False,
  "交付金の評価指標の明細と県内・全国との重ね合わせ。"
  "2-8の推移と5-9の活用方針があれば計画としては足りる。別冊へ移す"),
 ("G", "資5 用語の解説の絞り込み（本文に現れない8語を落とす）",
  [], False,
  "69語のうち、計画本文に一度も現れないのは8語である。"
  "ICF・運動器の機能向上・基本チェックリスト・経過的要介護・軽度認知障害（MCI）・"
  "支給限度基準額・フレイル・ロコモティブシンドローム。"
  "いずれも介護予防や制度の基本用語であり、本文で使っていないことのほうが"
  "問題という見方もできるため、落とすか本文で用いるかは村の判断による"),
 ("H", "10月6日に加えた図2点を資料編へ移す",
  [], False,
  "交付金の評価が求める分析の空白を埋めるために加えた2点"
  "（2-5 医療と介護の連携の加算、5-8 給付適正化の取組別得点）。"
  "評価には効くが本文の筋には必須でないため、資料編に移すことができる"),
 ("J", "1-8の3・1-8の5 令和8年9月の部会で示された事項と継続検討事項を資料編へ移す",
  [("1-8", "3　令和8年9月の社会保障審議会介護保険部会において示された事項"),
   ("1-8", "5　引き続き検討することとされた事項")],
  False,
  "令和8年10月8日に加えた2項。制度として定まっていない事項の一覧であり、"
  "本計画の算定には織り込んでいない（5-12に掲げている）。"
  "計画の期間中に結論が示され得るため委員の皆様には示す必要があるが、"
  "資料編に移しても本文の筋は通る。1-8の1・2（法律・政令により定まった事項）は"
  "本文に残す"),
 ("L", "5-7 保険料基準額の12のパターン（当方の算定ベース）を別冊へ移す",
  [("5-7", "保険料基準額のパターン")],
  False,
  "令和8年10月9日に5-7へ基金の取崩しの3案（国の算定ベース）を加えたため、"
  "節の中に前提の異なる2つの案の組が並んでいる。"
  "12のパターンは当方がワークシートの提供を受ける前に算定したもので、"
  "第1号被保険者負担割合と給付費の水準の効きは感応度の表が示している。"
  "決定の材料は3案であるため、12のパターンは別冊の「算定の根拠」へ移せる。"
  "「算定上の月額と条例上の基準額」の小見出しは3案の表が用いる四捨五入の説明であるため残す"),
 ("K", "2-5 介護経営DBの制度の概要を資料編へ移す",
  [("2-5", "供給の継続性を把握する手だて")],
  False,
  "令和8年10月8日に加えた項。村外を含む供給の継続性を把握する手だてとして"
  "介護サービス事業者経営情報データベースの制度を説明したもので、"
  "本計画は分析結果を用いていない（5-12）。制度の概要の表は資料編に移せる。"
  "供給の継続性を把握する旨は施策4-6に残る"),
]

# 本文に現れない用語（案G）。資5の表から落とす行を見出し語で指定する。
YOUGO_NAI = (
    "ICF（国際生活機能分類）", "運動器の機能向上", "基本チェックリスト", "経過的要介護",
    "軽度認知障害（MCI）", "支給限度基準額", "フレイル",
    "ロコモティブシンドローム（運動器症候群）",
)
# 資料編へ移せる図（案H）
FIG_UTSUSU = ("fig_連携の加算.png", "fig_給付適正化の得点.png")


def h3_range(sec_no, h3):
    """ある節の、ある小見出しから次の小見出しまでのブロックの位置を返す"""
    for ch in SC.CH:
        for sec in ch["sections"]:
            if sec["no"] != sec_no:
                continue
            bs = sec["blocks"]
            st = None
            for i, b in enumerate(bs):
                if b.get("t") == "h3" and b.get("v") == h3:
                    st = i
                elif st is not None and b.get("t") == "h3":
                    return (sec_no, st, i)
            if st is not None:
                return (sec_no, st, len(bs))
    return None


def items_without(targets, yougo=(), figs=()):
    """指定した範囲を落として要素を並べ直す。

    targets 小見出しの範囲（節, 小見出し）
    yougo   資5の表から落とす見出し語（案G）
    figs    本文から落とす図のファイル名（案H。資料編へ移す）
    """
    ranges = []
    for sec_no, h3 in targets:
        r = h3_range(sec_no, h3)
        if r:
            ranges.append(r)
    saved = {}
    for sec_no, st, en in ranges:
        saved.setdefault(sec_no, []).append((st, en))
    # soan_content を直接いじらず、落とす範囲を覚えておいて items を組み直す
    orig = {}
    for ch in SC.CH:
        for sec in ch["sections"]:
            if sec["no"] in saved:
                orig[sec["no"]] = sec["blocks"]
                drop = set()
                for st, en in saved[sec["no"]]:
                    drop |= set(range(st, en))
                sec["blocks"] = [b for i, b in enumerate(sec["blocks"])
                                 if i not in drop]
    # 案G　資5の表から、本文に現れない用語の行を落とす
    orig_rows = []
    if yougo:
        for ch in SC.CH:
            for sec in ch["sections"]:
                if sec["no"] != "資5":
                    continue
                for b in sec["blocks"]:
                    if b["t"] != "table":
                        continue
                    orig_rows.append((b, b["rows"]))
                    b["rows"] = [r for r in b["rows"] if str(r[0]) not in yougo]

    # 案H　本文に置いた図を落とす（資料編へ移す）
    import figures_map as _FM
    orig_figs = None
    if figs:
        orig_figs = list(_FM.FIGS)
        _FM.FIGS = [e for e in _FM.FIGS if e[1] not in figs]

    try:
        return EL.soan_items(True)
    finally:
        for ch in SC.CH:
            for sec in ch["sections"]:
                if sec["no"] in orig:
                    sec["blocks"] = orig[sec["no"]]
        for b, rows in orig_rows:
            b["rows"] = rows
        if orig_figs is not None:
            _FM.FIGS = orig_figs


def main():
    base = EL.pages(EL.soan_items(True), EL.BODY_H)
    MAE = 4          # 前付（表紙・本書の見方・目次）
    MOKUYASU = 100
    print("■ 計画書の分量を仕様書の目安に近づける削減案\n")
    print(f"  現在　本文 {base}頁 ＋ 前付 {MAE}頁 ＝ {base + MAE}頁"
          f"（目安 {MOKUYASU}頁に対し {base + MAE - MOKUYASU:+d}頁）\n")
    print(f"  {'案':3s} {'内容':44s} {'減る頁':>6s} {'法定':>4s}")
    print("  " + "─" * 70)
    rows = []
    for no, nm, tg, hotei, _riyu in AN:
        if no == "G":
            n = EL.pages(items_without([], yougo=YOUGO_NAI), EL.BODY_H)
        elif no == "H":
            n = EL.pages(items_without([], figs=FIG_UTSUSU), EL.BODY_H)
        elif not tg:
            rows.append((no, nm, None, hotei))
            print(f"  {no:3s} {nm[:42]:44s} {'（別途）':>6s} {'―':>4s}")
            continue
        else:
            n = EL.pages(items_without(tg), EL.BODY_H)
        rows.append((no, nm, base - n, hotei))
        print(f"  {no:3s} {nm[:42]:44s} {base - n:5d}頁 "
              f"{'★要' if hotei else '―':>4s}")
    # すべて行ったとき
    allt = [t for _n, _m, tg, _h, _r in AN for t in tg]
    n_all = EL.pages(items_without(allt, yougo=YOUGO_NAI, figs=FIG_UTSUSU), EL.BODY_H)
    print("  " + "─" * 70)
    print(f"  {'A〜H をすべて行った場合':50s} {base - n_all:5d}頁")
    print(f"\n  本文 {n_all}頁 ＋ 前付 {MAE}頁 ＝ {n_all + MAE}頁"
          f"（目安 {MOKUYASU}頁に対し {n_all + MAE - MOKUYASU:+d}頁）")

    # ── 組み合わせを総当たりし、目安に近いものを示す ──
    #    案ごとの減り方は足し算にならない。頁の境目のため、
    #    単独では1頁減るものが、他と重ねると0頁になることがある。
    import itertools
    keys = [a[0] for a in AN]
    kumi = {}
    for n in range(1, len(keys) + 1):
        for ks in itertools.combinations(keys, n):
            tg = [t for k in ks for t in dict((a[0], a[2]) for a in AN)[k]]
            pg = EL.pages(items_without(
                tg, yougo=YOUGO_NAI if "G" in ks else (),
                figs=FIG_UTSUSU if "H" in ks else ()), EL.BODY_H) + MAE
            kumi.setdefault(pg, []).append("".join(ks))
    mn = min(kumi)
    print("\n■ 目安に近い組み合わせ（総当たり）\n")
    print(f"  {'合計頁':>6s}  {'目安差':>5s}  最も少ない案で届く組み合わせ")
    for pg in sorted(kumi)[:5]:
        saitan = min(kumi[pg], key=len)
        print(f"  {pg:5d}頁  {pg - MOKUYASU:+5d}  {saitan}")
    print(f"\n  到達できる最小は {mn}頁（{min(kumi[mn], key=len)}）。"
          f"目安{MOKUYASU}頁に対し {mn - MOKUYASU:+d}頁。")
    print("  これ以上は、介護保険法第117条第2項が求める事項か前付に手を入れることになる。")
    print("\n■ 各案の理由と移し先\n")
    for no, nm, _tg, hotei, riyu in AN:
        print(f"  {no}　{nm}")
        print(f"      {riyu}")
        if hotei:
            print("      ★ 介護保険法第117条第2項の記載事項にあたる。削れない")
        print()
    print("  ※ どれを削るかは村の判断である。本スクリプトは素案を変えない。")
    print("  ※ 「別冊」は、計画書とは別に村へ渡す算定の根拠の綴りを指す。")
    print("     納品する電子媒体の 06_算定の根拠 に収めるCSVと対になる。")
    print("  ※ 介護保険法第117条第2項が求める事項（サービスの種類ごとの量の見込み、")
    print("     地域支援事業の量、必要利用定員総数、給付適正化）は削る候補にしていない。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
