# -*- coding: utf-8 -*-
"""地域支援事業費の置き方の点検（令和8年10月7日）

08_作業順位 の順位23。

当初の見立ては「地域支援事業の事業別の費用が未受領で置けない」であった。
**調べた結果、置けていないのは2区分だけであり、それより重い問題が見つかった。**

  ⚠ 包括的支援事業（地域包括支援センターの運営）と任意事業の費用が、
    令和9〜11年度にかけて**直線で大きく伸びる値**で入力されている。
    伸びは令和6年度→令和8年度の平均年増をそのまま延ばしたものであり、
    **令和12年度でいったん下がる**（不連続になる）。
    地域支援事業費は第1号被保険者負担分（23％）に直に乗るため、
    **保険料に効く。**

  出力　05_試算・管理シート/川崎町_地域支援事業費の置き方_R8.10.7.xlsx
          00_この表について
          01_事業別に置けているもの・置けていないもの
          02_見える化システムの入力値の伸びの検証
          03_置き方の4案と保険料への影響
          04_受領をお願いしたい資料
          05_確認事項

⚠ **当方は見える化システムのデータを直接触らない。** 本表は、町から
  ご提供いただいた総括表（令和8年9月11日出力・訂正後）の値を読んで
  検証したものであり、入力値を変えるものではない。

⚠ **推測で数値を置き換えていない。** 4案はいずれも、総括表にある
  実績・実績見込みの値から機械的に作った**比較のための試算**である。
  どれを採るかは町のご判断（確認事項No.168）による。
"""
import importlib.util
import os
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)


def load(name, fname):
    spec = importlib.util.spec_from_file_location(name,
                                                  os.path.join(HERE, fname))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


GY = load("gyotei5", "build_iinkai_gyotei_R8.10.1.py")

SOKATSU = ("09_元資料/R8実績データ/R8.9.11受領版/"
           "【川崎町】第10期_将来推計総括表_R8.9.11出力_④訂正後.xlsx")
MIKOMI = "05_試算・管理シート/川崎町_第10期_計画見込量_R8.9.15.xlsx"
OUT = "05_試算・管理シート/川崎町_地域支援事業費の置き方_R8.10.7.xlsx"

# 保険料の算定に用いる値（⑦保険料算定シートから読む）
FUTAN = 0.23          # 第1号被保険者負担割合
SHUNO = 0.96          # 予定保険料収納率

NEN = ["R6", "R7", "R8", "R9", "R10", "R11", "R12"]
NEN_J = {"R6": "令和6年度", "R7": "令和7年度", "R8": "令和8年度",
         "R9": "令和9年度", "R10": "令和10年度", "R11": "令和11年度",
         "R12": "令和12年度"}

KAKUNIN = ["13", "39", "88", "141", "150", "33", "11"]


def yomu_sokatsu():
    """総括表の 3_地域支援事業費 から、区分別・事業別の額を読む。"""
    wb = openpyxl.load_workbook(SOKATSU, data_only=True)
    ws = wb["3_地域支援事業費"]
    out = {}
    for r in range(1, ws.max_row + 1):
        nm = ws.cell(r, 1).value
        if not isinstance(nm, str) or not nm.strip():
            continue
        nm = nm.strip()
        vals = [ws.cell(r, c).value for c in range(2, 9)]   # R6〜R12
        if all(v is None for v in vals):
            continue
        if not any(isinstance(v, (int, float)) for v in vals):
            continue
        out.setdefault(nm, dict(zip(NEN, vals)))
    return out


def yomu_hokenryo():
    """計画見込量 ⑦保険料算定 から、保険料の算定に要る値を読む。"""
    wb = openpyxl.load_workbook(MIKOMI, data_only=True)
    ws = wb["⑦保険料算定"]
    d = {}
    for r in range(1, ws.max_row + 1):
        k = str(ws.cell(r, 1).value or "").strip()
        if k == "地域支援事業費":
            d["chiiki"] = [ws.cell(r, c).value for c in (2, 3, 4)]
        if k == "保険料収納必要額":
            d["shuno_hitsuyo"] = ws.cell(r, 5).value
        if k.startswith("補正後被保険者数"):
            d["hosei"] = ws.cell(r, 2).value
        if k.startswith("予定保険料収納率"):
            d["ritsu"] = ws.cell(r, 2).value
        if k.startswith("保険料基準額"):
            d["gaku"] = ws.cell(r, 2).value
    for k in ("chiiki", "shuno_hitsuyo", "hosei", "ritsu", "gaku"):
        if k not in d:
            raise SystemExit("⑦保険料算定から読めない：" + k)
    return d


def main():
    for p in (SOKATSU, MIKOMI):
        if not os.path.exists(p):
            raise SystemExit("入力がない：" + p)
    S = yomu_sokatsu()
    H = yomu_hokenryo()

    HOUKATSU = "包括的支援事業(地域包括支援センターの運営)"
    NINI = "任意事業"
    for k in (HOUKATSU, NINI):
        if k not in S:
            raise SystemExit("総括表に行がない：" + k)

    # ── 包括＋任意の合計（この2つが伸びの元である）
    def goukei(nen):
        return S[HOUKATSU][nen] + S[NINI][nen]

    # ── 置き方の4案（令和9〜11年度の包括＋任意の額）
    r8 = goukei("R8")
    r7 = goukei("R7")
    zogen = (goukei("R8") - goukei("R6")) / 2      # 令和6→8の平均年増
    AN = [
        ("案1（現行・見える化システムの入力値）",
         [goukei("R9"), goukei("R10"), goukei("R11")],
         "令和6年度→令和8年度の平均年増（{:,.0f}円）をそのまま"
         "令和9〜11年度に延ばした値。総括表に現に入力されている。"
         .format(zogen),
         "⚠ 令和12年度は{:,.0f}円に下がり、不連続になる。"
         .format(goukei("R12"))),
        ("案2（令和8年度の実績見込み額で据え置き）",
         [r8, r8, r8],
         "令和8年度の実績見込み額{:,.0f}円を3か年とも置く。".format(r8),
         "直近の水準を保つ置き方。令和8年度は年度途中の見込みであり、"
         "決算により動く。"),
        ("案3（令和7年度の実績額で据え置き）",
         [r7, r7, r7],
         "直近の完結年度である令和7年度の実績額{:,.0f}円を"
         "3か年とも置く。".format(r7),
         "完結した年度の実績による置き方。"
         "令和8年度に増えた分（{:,.0f}円）を見込まないことになる。"
         .format(r8 - r7)),
        ("案4（平均年増の2分の1で延ばす）",
         [r8 + zogen / 2, r8 + zogen, r8 + zogen * 3 / 2],
         "令和6年度→令和8年度の平均年増の2分の1（{:,.0f}円）で"
         "延ばす。".format(zogen / 2),
         "案1と案2の中間。伸びの根拠が確かめられるまでの仮置き。"),
    ]

    # ── 保険料への影響
    #    地域支援事業費は第1号被保険者負担分（23％）に直に乗る。
    #    調整交付金は標準給付費と総合事業費によるため、
    #    包括的支援事業・任意事業を動かしても変わらない。
    genko = sum(AN[0][1])

    def hokenryo_sa(an_goukei):
        sa = an_goukei - genko
        return (sa * FUTAN / H["ritsu"] / H["hosei"] / 12, sa)

    # ── 事業別に置けているもの・置けていないもの
    JIGYO = []
    for nm in (HOUKATSU, NINI):
        JIGYO.append([nm.replace("(", "（").replace(")", "）"),
                      "区分の合計のみ",
                      f"{S[nm]['R7']:,.0f}", f"{S[nm]['R11']:,.0f}",
                      "⚠ 国の様式が区分の合計で求めるものであり、"
                      "事業別の内訳は様式にない。"
                      "町の決算の科目と突き合わせて確かめる必要がある。"])
    SHAKAI = ["在宅医療・介護連携推進事業", "生活支援体制整備事業",
              "認知症初期集中支援推進事業", "認知症地域支援・ケア向上事業",
              "認知症サポーター活動促進・地域づくり推進事業",
              "地域ケア会議推進事業"]
    for nm in SHAKAI:
        if nm not in S:
            raise SystemExit("総括表に行がない：" + nm)
        r7v, r11v = S[nm]["R7"], S[nm]["R11"]
        JIGYO.append([nm, "事業別に置けている",
                      f"{r7v:,.0f}", f"{r11v:,.0f}",
                      "置けている。" if r7v or r11v
                      else "⚠ 全年度0。事業を行っていないのか、"
                           "他の区分で経理しているのかを確かめる必要がある。"])

    wb = openpyxl.Workbook()

    # ────────────────────────── 00
    ws = GY.sheet(
        wb, "00_この表について", "地域支援事業費の置き方",
        "当初の見立ては「事業別の費用が未受領で置けない」でしたが、"
        "調べたところ、置けていないのは2区分"
        "（包括的支援事業の地域包括支援センターの運営・任意事業）だけで、"
        "これは国の様式が区分の合計で求めるためでした。\n"
        "⚠ それより重い問題が見つかりました。この2区分の費用が、"
        "令和9〜11年度にかけて直線で大きく伸びる値で入力されています。"
        "伸びは令和6年度→令和8年度の平均年増をそのまま延ばしたもので、"
        "令和12年度でいったん下がります（不連続）。"
        "地域支援事業費は第1号被保険者負担分（23％）に直に乗るため、"
        "保険料に効きます。\n"
        "⚠ 当方は見える化システムのデータを直接触りません。"
        "本表は町からご提供いただいた総括表（令和8年9月11日出力・訂正後）の"
        "値を読んで検証したものであり、入力値を変えるものではありません。"
        "4案はいずれも総括表にある実績・実績見込みから機械的に作った"
        "比較のための試算です。推測で数値を置き換えていません。",
        ["シート", "内容"], [34, 96], first=True)
    GY.put(ws, [
        ["01_事業別の状況",
         f"8つの事業・区分について、置けているかどうか。"],
        ["02_伸びの検証",
         "令和6〜8年度の実績と、令和9〜11年度の入力値、"
         "令和12年度の値を並べ、伸びの作られ方を確かめます。"],
        ["03_置き方の4案と保険料への影響",
         "4案それぞれの3か年の額と、保険料基準額（月額）への影響。"],
        ["04_受領をお願いしたい資料",
         "伸びの根拠を確かめるために要る資料。"],
        ["05_確認事項", f"{len(KAKUNIN)}件。"],
    ])

    # ────────────────────────── 01
    ws = GY.sheet(
        wb, "01_事業別の状況", "事業別に置けているもの・置けていないもの",
        "総括表（令和8年9月11日出力・訂正後）3_地域支援事業費 によります。\n"
        "⚠ 社会保障充実分の6事業のうち4事業が全年度0です。"
        "事業を行っていないのか、包括的支援事業（地域包括支援センターの運営）"
        "の中で経理しているのかを確かめる必要があります。"
        "認知症初期集中支援推進事業は現に行われており"
        "（チーム員会議12回／年・訪問12件／年）、"
        "費用が0であることと矛盾します。",
        ["事業・区分", "置けているか", "令和7年度（円）", "令和11年度（円）",
         "備考"], [42, 20, 20, 20, 56])
    GY.put(ws, JIGYO,
           [GY.NGF if "⚠" in x[4] else GY.OKF for x in JIGYO])

    # ────────────────────────── 02
    ws = GY.sheet(
        wb, "02_伸びの検証", "見える化システムの入力値の伸びの検証",
        "令和6〜8年度の実績・実績見込みと、令和9〜11年度の入力値、"
        "令和12年度の値を並べます。\n"
        "⚠ 令和9〜11年度は毎年同じ額ずつ増えており、"
        "その増え方は令和6年度→令和8年度の平均年増と一致します。"
        "つまり直線で延ばした値です。\n"
        "⚠ 令和12年度は別の延ばし方によるため、令和11年度から下がります。"
        "第10期（令和9〜11年度）だけが高く出る形になっています。",
        ["区分", "令和6年度", "令和7年度", "令和8年度", "令和9年度",
         "令和10年度", "令和11年度", "令和12年度"],
        [40, 16, 16, 16, 16, 16, 16, 16])
    rows, fills = [], []
    for nm in (HOUKATSU, NINI):
        rows.append([nm.replace("(", "（").replace(")", "）")]
                    + [f"{S[nm][n]:,.0f}" for n in NEN])
        fills.append(GY.NGF)
    rows.append(["　2区分の計"] + [f"{goukei(n):,.0f}" for n in NEN])
    fills.append(GY.WARN)
    rows.append(["　（前年からの増減）", "―"]
                + [f"{goukei(NEN[i]) - goukei(NEN[i - 1]):+,.0f}"
                   for i in range(1, len(NEN))])
    fills.append(None)
    GY.put(ws, rows, fills)
    r = 5 + len(rows) + 1
    ws.cell(r, 1).value = "確かめたこと"
    ws.cell(r, 1).font = openpyxl.styles.Font(name="游ゴシック", size=10,
                                              bold=True)
    KAKUMETA = [
        ["令和9〜11年度は直線である",
         "地域包括支援センターの運営は毎年 {:+,.0f}円、"
         "任意事業は毎年 {:+,.0f}円ずつ増えています。"
         "いずれも年による違いがありません。"
         .format(S[HOUKATSU]["R10"] - S[HOUKATSU]["R9"],
                 S[NINI]["R10"] - S[NINI]["R9"])],
        ["その増え方は令和6年度→令和8年度の平均年増と一致する",
         "地域包括支援センターの運営は（令和8年度 {:,.0f}円 － "
         "令和6年度 {:,.0f}円）÷2 ＝ {:+,.0f}円。"
         "上の毎年の増え方と同じです。"
         .format(S[HOUKATSU]["R8"], S[HOUKATSU]["R6"],
                 (S[HOUKATSU]["R8"] - S[HOUKATSU]["R6"]) / 2)],
        ["令和12年度で下がる",
         "2区分の計は令和11年度 {:,.0f}円 に対し、"
         "令和12年度 {:,.0f}円（{:+.1f}％）です。"
         "第10期だけが高く出る形になっています。"
         .format(goukei("R11"), goukei("R12"),
                 (goukei("R12") / goukei("R11") - 1) * 100)],
        ["令和8年度は年度途中の見込みである",
         "令和7年度 {:,.0f}円 から令和8年度 {:,.0f}円 へ"
         "{:+.1f}％ 増えています。この1年の増え方が"
         "そのまま第10期の3か年の伸びの土台になっています。"
         "令和8年度の決算により動きます。"
         .format(goukei("R7"), goukei("R8"),
                 (goukei("R8") / goukei("R7") - 1) * 100)],
        ["認定者数の伸びとは合わない",
         "認定者数は令和7年度560人から令和11年度576人へ"
         "2.9％の増加です。2区分の計は同じ期間に"
         "{:+.1f}％ 増えています。"
         .format((goukei("R11") / goukei("R7") - 1) * 100)],
    ]
    GY.put(ws, [[a, b, "", "", "", "", "", ""] for a, b in KAKUMETA],
           r0=r + 1)

    # ────────────────────────── 03
    ws = GY.sheet(
        wb, "03_置き方の4案と保険料への影響", "置き方の4案と保険料への影響",
        "地域支援事業費は第1号被保険者負担分（23％）に直に乗ります。"
        "調整交付金は標準給付費と総合事業費によるため、"
        "包括的支援事業・任意事業を動かしても変わりません。\n"
        f"保険料基準額（月額）の現行値は {H['gaku']:,.0f}円"
        "（算定Ａ・基金の取崩しなし）です。\n"
        "⚠ 4案は比較のための試算です。どれを採るかは町のご判断により、"
        "根拠となる資料（04シート）のご提供をお願いします。",
        ["案", "令和9年度（円）", "令和10年度（円）", "令和11年度（円）",
         "3か年計（円）", "案1との差（円）", "保険料月額への影響（円）",
         "内容"], [34, 18, 18, 18, 20, 18, 20, 56])
    rows, fills = [], []
    for name, vals, naiyo, bikou in AN:
        g = sum(vals)
        eikyo, sa = hokenryo_sa(g)
        rows.append([name] + [f"{v:,.0f}" for v in vals]
                    + [f"{g:,.0f}", f"{sa:+,.0f}", f"{eikyo:+,.0f}",
                       naiyo + bikou])
        fills.append(GY.NGF if name.startswith("案1") else GY.WARN)
    GY.put(ws, rows, fills)
    r = 5 + len(AN) + 1
    ws.cell(r, 1).value = "試算の式"
    ws.cell(r, 1).font = openpyxl.styles.Font(name="游ゴシック", size=10,
                                              bold=True)
    GY.put(ws, [[
        "保険料月額への影響 ＝ 3か年計の差 × 第1号被保険者負担割合"
        f"{FUTAN:.0%} ÷ 予定保険料収納率{H['ritsu']:.0%} ÷ "
        f"補正後被保険者数（3か年計）{H['hosei']:,.1f}人 ÷ 12",
        "", "", "", "", "", "", ""],
        ["⚠ 補正後被保険者数は令和7年度末の所得段階別の実績分布（13段階）"
         "によります。見える化システムの登録値（第10〜13段階が0人）の"
         "ままでは保険料が過大に出ます（確認事項No.33）。",
         "", "", "", "", "", "", ""]], r0=r + 1)

    # ────────────────────────── 04
    ws = GY.sheet(
        wb, "04_受領をお願いしたい資料", "伸びの根拠を確かめるために要る資料",
        "下記をいただければ、置き方を根拠のあるものに改められます。\n"
        "⚠ いただけない場合の当方の扱いは、各行の最後に書いています。",
        ["資料", "なぜ要るか", "いただけない場合の当方の扱い"],
        [40, 56, 56])
    SHIRYO = [
        ["令和6年度・令和7年度の地域支援事業の決算（科目別）",
         "総括表の令和6・7年度の額が決算と合っているかを確かめる。"
         "合っていれば、伸びの土台が確かなものになる。",
         "総括表の値をそのまま用い、出所を注記する。"],
        ["令和8年度の地域支援事業の予算と執行の見込み",
         "令和8年度は年度途中の見込みであり、"
         "この1年の増え方が第10期の3か年の伸びの土台になっている。"
         "予算額が分かれば、見込みの確からしさを確かめられる。",
         "案4（平均年増の2分の1）を協議用の仮置きとしてお示しする。"],
        ["地域包括支援センターの運営の委託契約の額と期間",
         "地域包括支援センターの運営費が毎年2,187,101円ずつ増える"
         "根拠があるか（人員の増・委託料の改定など）を確かめる。",
         "伸びの根拠を示せないため、素案 9-2 に"
         "「見える化システムの入力値による」と注記する。"],
        ["任意事業の事業別の内訳（成年後見・家族介護支援・住宅改修等）",
         "任意事業が令和7年度から令和11年度にかけて2倍になる"
         "根拠があるかを確かめる。",
         "同上。"],
        ["包括的支援事業（社会保障充実分）の6事業の経理の仕方",
         "認知症初期集中支援推進事業は現に行われているのに"
         "費用が全年度0である。どの区分で経理しているかを確かめる。",
         "素案 9-2 に「社会保障充実分として経理していない事業がある」"
         "旨を注記する。"],
    ]
    GY.put(ws, SHIRYO, [GY.WARN] * len(SHIRYO))

    # ────────────────────────── 05
    kak = yomu_kakunin(KAKUNIN)
    ws = GY.sheet(
        wb, "05_確認事項", "本表に関わる確認事項",
        "業務工程管理表 03_確認事項一覧 から読んでいます。"
        "内容はそちらが正本です。",
        ["No.", "確認事項", "状態", "決着しない場合の当方の扱い"],
        [8, 40, 12, 62])
    GY.put(ws, kak, [GY.OKF if x[2] == "完了" else None for x in kak])

    wb.save(OUT)

    # ══════════════════════════ 自己点検
    ng = []
    wb2 = openpyxl.load_workbook(OUT)
    if len(wb2.worksheets) != 6:
        ng.append(f"シートが{len(wb2.worksheets)}枚（6枚のはず）")
    # 案1が総括表の値そのものであること（読み違えていないこと）
    for i, n in enumerate(("R9", "R10", "R11")):
        if abs(AN[0][1][i] - goukei(n)) > 1:
            ng.append(f"案1の{NEN_J[n]}が総括表と一致しない")
    # 令和9〜11年度が直線であること（検証の前提）
    d1 = AN[0][1][1] - AN[0][1][0]
    d2 = AN[0][1][2] - AN[0][1][1]
    if abs(d1 - d2) > 2:
        ng.append(f"令和9〜11年度が直線でない（{d1:,.0f}／{d2:,.0f}）。"
                  "本表の見立ての前提が崩れている")
    # その増え方が令和6→8の平均年増と一致すること
    if abs(d1 - zogen) > 2:
        ng.append(f"増え方（{d1:,.0f}）が令和6→8の平均年増"
                  f"（{zogen:,.0f}）と一致しない")
    # 伸びが直線であることを確かめた記述があること
    zen = "\n".join(str(c.value) for w_ in wb2.worksheets
                    for row in w_.iter_rows() for c in row
                    if c.value is not None)
    for w in ("直線", "令和12年度", "第1号被保険者負担分"):
        if w not in zen:
            ng.append(f"核心の記述がない：{w}")
    for w in ("**", "__"):
        if w in zen:
            ng.append(f"強調記号が残っている：{w}")
    for w in ("大雪", "東川", "東神楽", "上川", "美瑛", "金ヶ崎", "川崎市"):
        if w in zen:
            ng.append(f"他団体の名称が入っている：{w}")
    # 見える化システムを当方が触らない旨
    if "直接触りません" not in zen:
        ng.append("見える化システムを当方が直接触らない旨の断りがない")
    # 保険料への影響が0でない案があること（試算が効いていること）
    eikyo = [hokenryo_sa(sum(v))[0] for _, v, _, _ in AN]
    if all(abs(e) < 1 for e in eikyo):
        ng.append("保険料への影響がすべて0（試算が効いていない）")

    print("地域支援事業費の置き方の点検")
    print("保存：", OUT)
    print(f"  シート{len(wb2.worksheets)}枚／事業・区分{len(JIGYO)}件"
          f"／置き方の案{len(AN)}件／要る資料{len(SHIRYO)}件"
          f"／確認事項{len(kak)}件")
    print("  ── 見つかったこと")
    print(f"   包括的支援事業（運営）＋任意事業　"
          f"令和7年度 {goukei('R7'):,.0f}円 → 令和11年度 "
          f"{goukei('R11'):,.0f}円（{goukei('R11') / goukei('R7') - 1:+.1%}）")
    print(f"   令和12年度は {goukei('R12'):,.0f}円 に下がる"
          f"（{goukei('R12') / goukei('R11') - 1:+.1%}・不連続）")
    for (name, vals, _, _), e in zip(AN, eikyo):
        print(f"   {name:36s} 3か年計 {sum(vals):>14,.0f}円"
              f"／保険料月額 {e:+,.0f}円")
    print("  ── 自己点検")
    if ng:
        for m in ng:
            print("   ×", m)
        sys.exit(1)
    print("   ○ 数値は総括表と計画見込量から読んでいる"
          "（推測で置き換えていない）")
    print("   ○ 見える化システムのデータを当方が直接触らない旨を明記している")
    print("  ⚠ どの案を採るかは町のご判断です（確認事項No.168）。")


def yomu_kakunin(nos):
    wb = openpyxl.load_workbook("川崎町_業務工程管理表.xlsx")
    ws = wb["03_確認事項一覧"]
    out, want = [], list(nos)
    for r in range(5, ws.max_row + 1):
        v = str(ws.cell(r, 1).value or "").strip()
        if v in want:
            out.append([v, str(ws.cell(r, 4).value or ""),
                        str(ws.cell(r, 8).value or ""),
                        str(ws.cell(r, 10).value or "")])
    out.sort(key=lambda x: want.index(x[0]))
    miss = set(want) - {x[0] for x in out}
    if miss:
        raise SystemExit("確認事項が見つからない：" + "／".join(sorted(miss)))
    return out


if __name__ == "__main__":
    main()
