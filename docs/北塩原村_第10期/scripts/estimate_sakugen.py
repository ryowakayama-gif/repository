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
 ("M", "資8 図表番号一覧を図表データ管理台帳へ移す（納品物と重複）",
  [("資8", None)],
  False,
  "図の番号・名称・掲載箇所・出典を並べた28行の表。同じ内容を図表データ管理台帳"
  "（納品する電子媒体 04_図表データ の 11_北塩原村第10期_図表データ管理台帳.xlsx。"
  "シート「04_資8の一覧との突合」で資8と突き合わせている）として別に納めており、重複している。"
  "計画の読み手が必要とするのは本文中の図そのものであって、番号の一覧ではない。"
  "第9期計画には対応する記載がない。1.3頁あり、削減案のなかでは効きが大きい。"
  "令和8年10月10日に紙面を実測して加えた案である"),
]

# 本文に現れない用語（案G）。資5の表から落とす行を見出し語で指定する。
YOUGO_NAI = (
    "ICF（国際生活機能分類）", "運動器の機能向上", "基本チェックリスト", "経過的要介護",
    "軽度認知障害（MCI）", "支給限度基準額", "フレイル",
    "ロコモティブシンドローム（運動器症候群）",
)
# 資料編へ移せる図（案H）
FIG_UTSUSU = ("fig_連携の加算.png", "fig_給付適正化の得点.png")


# ══════════════════════════════════════════════════════════════
# 実測（令和8年10月10日・scripts/render_check.py）
#   組み合わせ → (測ったときの推定本文頁, 実測本文頁)
#
# 推定は送りによる空白を小さく見るため下限である。実測との差は0〜3頁と
# ばらつき、一定ではない。**村にお示しする頁数は、組んで PDF に変換して
# 数えた値を用いる。** 測り直すには次を走らせる。
#   python3 scripts/estimate_sakugen.py --jissoku ABCDEFGHJKLM
# 素案が変わると推定も動く。測ったときの推定を一緒に持っておき、
# 現在の推定と食い違ったら「実測が古い」と分かるようにしている（点検52）。
JISSOKU = {
    "":              (118, 120),
    "ABCDE":         (112, 113),
    "ABCDEFH":       (107, 110),
    "ABCDEFHJ":      (105, 107),
    "ABCEFHJ":       (106, 107),
    "ABCDEFGHJKL":   (105, 106),
    "ABCDEFGHJKLM":  (104, 104),
}
MAE_PG = 4          # 前付（表紙・本書の見方・目次）
KYOYO = 110         # 村が許容する頁数（令和8年10月10日のご判断）

# 対応表に並べる組み合わせ（村が選ぶ順に意味のまとまりで並べる）
TAIOU = [
    ("",             "何も行わない"),
    ("ABCDE",        "第5章 5-4・5-7 の算定の方法と検証を別冊へ"),
    ("ABCDEFH",      "＋ 2-8 の交付金の明細と図2点を資料編へ"),
    ("ABCDEFHJ",     "＋ 1-8の3・5（部会で示された事項）を資料編へ"),
    ("ABCDEFGHJKL",  "＋ 資5の8語・2-5の介護経営DB・5-7の12パターン"),
    ("ABCDEFGHJKLM", "＋ 資8 図表番号一覧を図表データ管理台帳へ"),
]


def jissoku_furui():
    """実測が古くなっていないかを見る。古いものの一覧を返す。"""
    furui = []
    for k, (est0, _jis) in sorted(JISSOKU.items()):
        tg, yg, fg = an_targets(k)
        now = EL.pages(items_without(tg, yougo=yg, figs=fg), EL.BODY_H)
        if now != est0:
            furui.append((k or "なし", est0, now))
    return furui


def jissoku_of(keys):
    """その組み合わせの本文の頁数。実測があれば実測、なければ推定＋差。"""
    k = seiki(keys)
    tg, yg, fg = an_targets(k)
    est = EL.pages(items_without(tg, yougo=yg, figs=fg), EL.BODY_H)
    if k in JISSOKU and JISSOKU[k][0] == est:
        return JISSOKU[k][1], "実測"
    return est + EL.JITSU_SA, "推定＋差"


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


import contextlib


@contextlib.contextmanager
def nuki(targets, yougo=(), figs=()):
    """指定した範囲を落とした状態にして、抜けたら必ず元へ戻す。

    targets 小見出しの範囲（節, 小見出し）
    yougo   資5の表から落とす見出し語（案G）
    figs    本文から落とす図のファイル名（案H。資料編へ移す）

    推定（items_without）と、実紙面を測るための docx の組み立て（build_docx）の
    両方がこれを使う。**同じ落とし方でなければ、推定と実測を比べる意味がない。**
    """
    ranges = []
    marugoto = []          # 節まるごとを落とすもの（小見出しが None）
    for sec_no, h3 in targets:
        if h3 is None:
            marugoto.append(sec_no)
            continue
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
    # 節まるごとを落とす（案M）。章から節そのものを外す
    orig_secs = []
    if marugoto:
        for ch in SC.CH:
            if any(sec["no"] in marugoto for sec in ch["sections"]):
                orig_secs.append((ch, ch["sections"]))
                ch["sections"] = [sec for sec in ch["sections"]
                                  if sec["no"] not in marugoto]

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
    #   対応表（figures_map.FIGS）から外すだけでは、本文の {"t":"fig"} の
    #   ブロックが残り、docx を組むときに引く先が無くなって落ちる。
    #   推定の側は高さ0として素通ししていたため気づかなかった。本文からも外す。
    import figures_map as _FM
    orig_figs = None
    orig_fb = []
    if figs:
        orig_figs = list(_FM.FIGS)
        _FM.FIGS = [e for e in _FM.FIGS if e[1] not in figs]
        for ch in SC.CH:
            for sec in ch["sections"]:
                if any(b.get("t") == "fig" and b.get("v") in figs
                       for b in sec["blocks"]):
                    orig_fb.append((sec, sec["blocks"]))
                    sec["blocks"] = [b for b in sec["blocks"]
                                     if not (b.get("t") == "fig"
                                             and b.get("v") in figs)]

    try:
        yield
    finally:
        for ch in SC.CH:
            for sec in ch["sections"]:
                if sec["no"] in orig:
                    sec["blocks"] = orig[sec["no"]]
        for b, rows in orig_rows:
            b["rows"] = rows
        for sec, bs in orig_fb:
            sec["blocks"] = bs
        for ch, secs in orig_secs:
            ch["sections"] = secs
        if orig_figs is not None:
            _FM.FIGS = orig_figs


def items_without(targets, yougo=(), figs=()):
    """指定した範囲を落として要素を並べ直す（推定の側）。"""
    with nuki(targets, yougo, figs):
        return EL.soan_items(True)


def seiki(keys):
    """案の記号の並びを並べ替えて正規の形にする（JISSOKU の鍵と突き合わせるため）。"""
    return "".join(sorted(set(keys)))


def an_targets(keys):
    """案の記号の並び（例 "ABCEFHJ"）から、落とす範囲と案G・案Hの指定を返す。"""
    ks = [k for k in seiki(keys)]
    shiru = {a[0] for a in AN}
    for k in ks:
        if k not in shiru:
            raise SystemExit(f"案{k}は定義にない（あるのは {''.join(sorted(shiru))}）")
    tg = [t for a in AN if a[0] in ks for t in a[2]]
    return tg, (YOUGO_NAI if "G" in ks else ()), (FIG_UTSUSU if "H" in ks else ())


def build_docx(keys, out):
    """案を適用した計画書の docx を組む（実紙面を測るため）。

    推定は送りによる空白を小さく見るため下限にとどまる。どの組み合わせが
    村の許容に収まるかは、組んで変換して数えるのが確かである。
    納品物は上書きしない。out に書く。
    """
    import json, subprocess, tempfile
    tg, yg, fg = an_targets(keys)
    here = os.path.dirname(os.path.abspath(__file__))
    with nuki(tg, yg, fg):
        import figures_map as _FM
        figs_sec, figs_inline = {}, {}
        for _sec, _fn, _cap, _src, _inline in _FM.FIGS:
            rec = {"file": _fn, "caption": _cap, "source": _src}
            if _inline:
                figs_inline[_fn] = rec
            else:
                figs_sec.setdefault(_sec, []).append(rec)
        data = {"title": SC.TITLE, "title2": SC.TITLE2, "subtitle": SC.SUBTITLE,
                "draft": SC.DRAFT, "issuer": SC.ISSUER, "date": SC.DATE,
                "chapters": SC.CH, "figures": figs_sec,
                "figures_inline": figs_inline}
        fd, js = tempfile.mkstemp(suffix=".json", prefix="sakugen_")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
    r = subprocess.run(["node", os.path.join(here, "build_soan_docx.js"),
                        "--keikaku", out, "--in", js],
                       capture_output=True, text=True, timeout=900)
    os.unlink(js)
    if r.returncode != 0 or not os.path.exists(out):
        raise SystemExit((r.stdout + r.stderr)[-2000:])
    return out


def main():
    base = EL.pages(EL.soan_items(True), EL.BODY_H)
    MAE = 4 + EL.JITSU_SA   # 前付（表紙・本書の見方・目次）4頁 ＋ 実測との差
    #  令和8年10月10日に docx を PDF に変換して実測したところ、計画書は
    #  推定118頁に対し実測120頁であった（素案も推定135頁に対し実測137頁）。
    #  推定は送りの空白を小さく見るため下限である。村の許容に収まるかを
    #  判断する数であるから、ここでは実測との差（EL.JITSU_SA）を乗せる。
    MOKUYASU = 100
    print("■ 計画書の分量を仕様書の目安に近づける削減案\n")
    n_ima, doko_ima = jissoku_of("")
    print(f"  現在　本文 {n_ima}頁 ＋ 前付 {MAE_PG}頁 ＝ {n_ima + MAE_PG}頁（{doko_ima}）"
          f"（目安 {MOKUYASU}頁に対し {n_ima + MAE_PG - MOKUYASU:+d}頁、"
          f"村が許容する{KYOYO}頁に対し {n_ima + MAE_PG - KYOYO:+d}頁）\n")
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
    print(f"  {'すべて行った場合（推定ベース）':50s} {base - n_all:5d}頁")
    zenbu = seiki("".join(a[0] for a in AN))
    n_zen, doko_zen = jissoku_of(zenbu)
    print(f"\n  本文 {n_zen}頁 ＋ 前付 {MAE_PG}頁 ＝ {n_zen + MAE_PG}頁（{doko_zen}）"
          f"（目安 {MOKUYASU}頁に対し {n_zen + MAE_PG - MOKUYASU:+d}頁、"
          f"村が許容する{KYOYO}頁に対し {n_zen + MAE_PG - KYOYO:+d}頁）")

    # ── 実測による対応表（村が選ぶための表） ──
    furui = jissoku_furui()
    print("\n■ 実測による対応表　どの案を行うと何頁になるか\n")
    if furui:
        print("  ※ 実測が古い（素案が変わっている）：" +
              "・".join(f"{k}（測ったときの推定{a}頁→いまは{b}頁）"
                        for k, a, b in furui))
        print("    python3 scripts/estimate_sakugen.py --jissoku <案の記号> で測り直す\n")
    print(f"  {'行う案':<14}{'本文':>5}{'前付':>5}{'合計':>5}"
          f"{'目安100頁':>10}{'許容110頁':>10}  {'出どころ':<8}内容")
    print("  " + "─" * 104)
    for k, nm in TAIOU:
        n, doko = jissoku_of(k)
        g = n + MAE_PG
        print(f"  {k or '（なし）':<14}{n:5d}{MAE_PG:5d}{g:5d}"
              f"{g - MOKUYASU:+10d}{g - KYOYO:+10d}  {doko:<8}{nm}")
    print("  " + "─" * 104)
    n_min, _ = jissoku_of(TAIOU[-1][0])
    if n_min + MAE_PG <= KYOYO:
        print(f"  すべて行えば{n_min + MAE_PG}頁で、村が許容する{KYOYO}頁に"
              f"{KYOYO - n_min - MAE_PG}頁の余裕がある。")
    else:
        print(f"  すべて行っても{n_min + MAE_PG}頁で、村が許容する{KYOYO}頁に"
              f"{n_min + MAE_PG - KYOYO}頁足りない。")
    print("  ※ 本文の頁数は docx を PDF に変換して数えた値である（doc83）。"
          "推定は送りによる空白を小さく見るため0〜3頁少なく出る。")
    print("  ※ 游ゴシックが無い環境のため IPAゴシックで組んだ紙面である。"
          "Word＋游ゴシックではなお±1〜2頁動き得る。")
    print("  ※ 案Mを行う場合は、資9（関連計画との対照）を資8に繰り上げる。")

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
    print("\n■ 目安に近い組み合わせ（総当たり・推定ベース）\n")
    print("  推定に実測との差を一律に乗せたものである。上の対応表（実測）と"
          "1〜2頁食い違うことがある。\n")
    print(f"  {'合計頁':>6s}  {'目安差':>5s}  最も少ない案で届く組み合わせ")
    for pg in sorted(kumi)[:5]:
        saitan = min(kumi[pg], key=len)
        print(f"  {pg:5d}頁  {pg - MOKUYASU:+5d}  {saitan}")
    print(f"\n  到達できる最小は {mn}頁（{min(kumi[mn], key=len)}）。"
          f"目安{MOKUYASU}頁に対し {mn - MOKUYASU:+d}頁、"
          f"村が許容する110頁に対し {mn - 110:+d}頁。")
    if mn > 110:
        print(f"  **A〜L をすべて行っても村の許容に {mn - 110}頁足りない。** 次の手だてがある。")
        print("    ① 1頁に収まらない表（3-5・4-6・資8）を分けるか詰める")
        print("       estimate_layout.py が「1頁に収まらない要素」として挙げている。"
              "いずれも1.0〜1.1頁で、残りの0.1頁ぶんが次頁に1行だけ乗って空白を生む")
        print("    ② 送りで3割以上の空白が出る箇所（estimate_layout.py が挙げる13件）の"
              "図と見出しの並びを入れ替える")
        print("    ③ 村に111頁の可否を諮る（許容110頁は令和8年10月10日の口頭の判断）")
    else:
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


def jissoku_measure(kumi):
    """組み合わせごとに docx を組んで PDF に変換し、本文の頁数を数える。

    JISSOKU に貼る行をそのまま出す。素案を直したら測り直す。
    """
    import tempfile
    import render_check as RC
    import pymupdf
    out = tempfile.mkdtemp(prefix="sakugen_")
    print("■ 実測（組んで変換して数える）\n")
    for k in kumi:
        nm = k or "なし"
        tg, yg, fg = an_targets(k)
        est = EL.pages(items_without(tg, yougo=yg, figs=fg), EL.BODY_H)
        dx = os.path.join(out, nm + ".docx")
        build_docx(k, dx)
        pdf = RC.convert(dx, os.path.join(out, nm))
        n = pymupdf.open(pdf).page_count - 3        # 前付3頁（目次は空のまま出る）
        mae = JISSOKU.get(k)
        shirushi = "" if mae is None else ("　（前は %d）" % mae[1] if mae[1] != n else "　（前と同じ）")
        print(f'    "{k}": ({est}, {n}),'.ljust(34)
              + f"# 推定{est}頁 → 実測{n}頁{shirushi}")
    print("\n  上の行を estimate_sakugen.py の JISSOKU に貼る。")
    print(f"  組んだ docx と PDF： {out}")


if __name__ == "__main__":
    if "--docx" in sys.argv:
        i = sys.argv.index("--docx")
        keys = sys.argv[i + 1] if len(sys.argv) > i + 1 else ""
        j = sys.argv.index("--out") if "--out" in sys.argv else -1
        out = sys.argv[j + 1] if j >= 0 else os.path.join(
            __import__("tempfile").gettempdir(), f"06_計画書_{keys or 'なし'}.docx")
        print("組んだ:", build_docx(keys.replace("なし", ""), out))
    elif "--jissoku" in sys.argv:
        i = sys.argv.index("--jissoku")
        jissoku_measure([a.replace("なし", "") for a in sys.argv[i + 1:]]
                        or [k for k, _ in TAIOU])
    else:
        sys.exit(main())
