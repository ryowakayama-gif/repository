# -*- coding: utf-8 -*-
"""認知症の人の数の**将来**の推計（令和8年10月19日）　— 算定の正本

  python3 07_ソーススクリプト/calc_ninchisho_shorai_R8.10.19.py

何をするものか
  令和12年度・令和17年度・令和22年度の認知症の人の数と軽度認知障害
  （ＭＣＩ）の人の数を推計する。現在（令和8年）の推計は
  `calc_ninchisho_R8.10.19.py` にあり、本スクリプトはそれを土台にする。

⚠ **これは暫定の算定である。**
  有病率は6つの年齢階級（65-69/70-74/75-79/80-84/85-89/90以上）で
  公表されているのに対し、町からご提供いただいた将来の人口は
  **3つの区分（前期65〜74歳／後期75〜84歳／後期85歳〜）**しかない。
  そのため、区分の中の内訳を当方で置いている。

  ⚠ **正道は按分しないことである。** 地域包括ケア「見える化」システムの
  「自然体推計の計算過程確認シート」は、令和12・17・22・27・32年度の
  第1号被保険者数を**性・年齢5歳階級別（65-69〜90歳以上）**で持つ
  （国のガイド「2.1 総人口と被保険者数」2)）。町に同シートを出力して
  いただければ、按分は一切要らなくなる。
  現に、町の見える化の画面（令和8年9月11日のキャプチャ）では、
  令和8年度の実績も令和9年度の推計値も、男女別・6階級で表示されている。

区分の中の内訳の置き方（4つを計算して並べる）
  Ａ　現在（令和8年6月末）の町の年齢構成で固定する
  Ｃ1　町の1歳刻みの人口をコーホートとして送る（死亡率を当てない）
  Ｃ2　同上（全国の死亡率を当てる）　← **本スクリプトの正本**
  簡便　第1号被保険者数×全国の有病率（12.3→14.2→14.9％）

  ⚠ **Ａは採らない。** 本町の75〜84歳の現在の内訳（75-79歳784人・
  80-84歳354人）は、団塊の世代とその直後の出生の落ち込みが隣り合った
  一過性の姿であり、年とともに反転する。Ａはこの歪みを将来まで凍結する。

  ⚠ **Ｃ1は採らない。** 死亡率を当てないと最高齢の階級に人が残りすぎ、
  90歳以上を多く見積もる。

出所
  ① 有病率・現在の人口　calc_ninchisho_R8.10.19.py（同じ原典を使う）
  ② 将来の第1号被保険者数（3区分）
     09_元資料/町提供資料/R8.10.19受領/
     【見える化】第10期_将来推計総括表_R8.10.19出力_パターン0911.xlsx
     シート「5_保険料推計」行111〜115
     ⚠ この3区分は令和8年9月11日出力の版と**1つも変わっていない。**
  ③ 全国の死亡率（令和6年・男女別・5歳階級）
     09_元資料/その他データ/
     日本の統計2026_表2-17_年齢男女別死亡数と死亡率_令和6年.xlsx
     （総務省統計局「日本の統計2026」。資料：厚生労働省「人口動態調査」）
  ④ 全国の認知症有病率の将来値（12.3％→14.2％(2030)→14.9％(2040)）
     09_元資料/認知症施策/R8.10.19取得版/
     001279920_認知症及びMCIの高齢者数と有病率の将来推計.pdf 1頁の表

⚠ **母集団の違いに注意。** 現在の推計は住民基本台帳の65歳以上人口
  （3,209人・令和8年6月末）を分母にしている。将来の3区分は
  **第1号被保険者数**（令和8年度3,257人）である。48人（1.5％）違う。
  本スクリプトは将来値を第1号被保険者数の基準で出し、
  **有病率（％）**を基準に依らない指標として併せて示す。

  終了コード　0＝検査1〜7がすべて適合／1＝1件でも不適合
"""
import importlib.util
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location(
    "calc_ninchisho", os.path.join(HERE, "calc_ninchisho_R8.10.19.py"))
C = importlib.util.module_from_spec(spec)
sys.modules["calc_ninchisho"] = C
spec.loader.exec_module(C)

SOKATSU = ("09_元資料/町提供資料/R8.10.19受領/"
           "【見える化】第10期_将来推計総括表_R8.10.19出力_パターン0911.xlsx")
SHIBOU = ("09_元資料/その他データ/"
          "日本の統計2026_表2-17_年齢男女別死亡数と死亡率_令和6年.xlsx")

# 町の人口表の基準日（令和8年6月末）からの年数。
#   年度の値は年度の中央（10月1日）に置く。
#   令和12年度＝2030年10月1日 … 2026年6月30日から 4.25年
NENSU = {"令和12年度": 4.25, "令和17年度": 9.25, "令和22年度": 14.25}

# 総括表 5_保険料推計 の列（行111〜115）
RETSU = {"令和12年度": 10, "令和17年度": 11, "令和22年度": 12}
GEN_RETSU = 6          # 合計の列（使わない。年度列の起点の確かめに用いる）

# 全国の認知症有病率の将来値（出所④・65歳以上全体・％）
ZENKOKU_SHORAI = {"令和12年度": 14.2, "令和17年度": 14.9,
                  "令和22年度": 14.9}
#   ⚠ 原典は 2030年14.2％・2040年14.9％。令和17年度（2035年度）の
#     公表値は無いため、簡便法の比較には2040年の値を当てる。
#     **簡便法は参考であり本計画では採らない。**

# 3区分と、それが含む年齢階級
KUBUN = [("前期（65〜74歳）", ["65〜69歳", "70〜74歳"], 112),
         ("後期（75〜84歳）", ["75〜79歳", "80〜84歳"], 114),
         ("後期（85歳〜）", ["85〜89歳", "90歳以上"], 115)]


def yomu_sokatsu():
    """将来の第1号被保険者数（3区分）を総括表から読む。"""
    import openpyxl
    if not os.path.exists(SOKATSU):
        raise SystemExit("総括表がない：" + SOKATSU)
    ws = openpyxl.load_workbook(SOKATSU, data_only=True)["5_保険料推計"]
    # 行110 の見出しで年度の列を確かめる（列の位置を決め打ちにしない）
    retsu = {}
    for c in range(5, 16):
        v = str(ws.cell(110, c).value or "").strip()
        if v in NENSU:
            retsu[v] = c
    nai = [k for k in NENSU if k not in retsu]
    if nai:
        raise SystemExit("総括表に年度の列が見つからない：" + "／".join(nai))
    out = {}
    for nendo, c in retsu.items():
        out[nendo] = {
            "第1号": int(ws.cell(111, c).value),
            "前期（65〜74歳）": int(ws.cell(112, c).value),
            "後期（75〜84歳）": int(ws.cell(114, c).value),
            "後期（85歳〜）": int(ws.cell(115, c).value),
        }
    # 令和8年度（現在）の第1号被保険者数も読む（母集団の違いを示すため）
    for c in range(5, 16):
        if str(ws.cell(110, c).value or "").strip() == "令和9年度":
            out["令和9年度"] = {
                "第1号": int(ws.cell(111, c).value),
                "前期（65〜74歳）": int(ws.cell(112, c).value),
                "後期（75〜84歳）": int(ws.cell(114, c).value),
                "後期（85歳〜）": int(ws.cell(115, c).value),
            }
    return out, retsu


def yomu_shibou():
    """全国の死亡率（令和6年・男女別）を単年齢に開く。戻り　{性: {年齢: 率}}"""
    import openpyxl
    if not os.path.exists(SHIBOU):
        raise SystemExit("死亡率の表がない：" + SHIBOU)
    ws = openpyxl.load_workbook(SHIBOU, data_only=True)["02-17"]
    # 行7〜27 が年齢階級。列6＝総数・列7＝男・列8＝女（人口10万につき）
    kaikyu = []
    for r in range(7, 28):
        na = str(ws.cell(r, 2).value or "").replace("　", " ").strip()
        if not na:
            continue
        na = na.replace("歳", "").replace(" ", "")
        if na.startswith("100"):
            lo, hi = 100, 109
        elif "～" in na:
            a, b = na.split("～")
            lo, hi = int(a), int(b)
        else:
            continue
        otoko = float(ws.cell(r, 7).value) / 100000.0
        onna = float(ws.cell(r, 8).value) / 100000.0
        kaikyu.append((lo, hi, otoko, onna))
    if len(kaikyu) != 21:
        raise SystemExit(f"年齢階級が21件でない（{len(kaikyu)}件）")
    # 階級の中央の年齢に率を置き、対数線形で単年齢へ内挿する
    #   （死亡率は年齢に対しておおむね指数で増えるため）
    ten = {"男": [], "女": []}
    for lo, hi, o, n in kaikyu:
        chu = (lo + hi) / 2.0
        ten["男"].append((chu, o))
        ten["女"].append((chu, n))
    out = {}
    for sei in ("男", "女"):
        t = ten[sei]
        d = {}
        for a in range(45, 115):
            if a <= t[0][0]:
                d[a] = t[0][1]
            elif a >= t[-1][0]:
                d[a] = t[-1][1]
            else:
                for i in range(len(t) - 1):
                    if t[i][0] <= a <= t[i + 1][0]:
                        x0, y0 = t[i]
                        x1, y1 = t[i + 1]
                        w = (a - x0) / (x1 - x0)
                        d[a] = math.exp(math.log(y0) * (1 - w)
                                        + math.log(y1) * w)
                        break
        out[sei] = d
    return out


def okuru(nen, shibou=None, bairitsu=1.0):
    """町の1歳刻みの人口を nen 年だけ送る。

      shibou  None なら死亡率を当てない（方法Ｃ1）
      bairitsu 死亡率にかける倍率（感度を見るため）

    戻り　{階級: (男, 女, 計)}　※6階級。合計は見える化に合わせる前の値。
    """
    # 1歳刻みの元の姿（日本人＋外国人）
    moto = {}
    for (a, nm, nf, gm, gf) in C.MACHI_1SAI:
        moto[a] = (nm + gm, nf + gf)

    waku = [(65, 69), (70, 74), (75, 79), (80, 84), (85, 89), (90, 999)]
    out = {na: [0.0, 0.0, 0.0] for na in C.KAIKYU}
    for a, (o, n) in moto.items():
        atarashii = a + nen          # nen 年後の年齢（小数）
        if atarashii < 65:
            continue
        seizon_o = seizon_n = 1.0
        if shibou is not None:
            # a 歳から atarashii 歳まで、各年の生存率を掛ける
            nokori = nen
            toshi = a
            while nokori > 0:
                haba = min(1.0, nokori)
                ko = int(min(max(toshi, 45), 110))
                seizon_o *= (1.0 - min(shibou["男"][ko] * bairitsu, 0.95)) ** haba
                seizon_n *= (1.0 - min(shibou["女"][ko] * bairitsu, 0.95)) ** haba
                toshi += haba
                nokori -= haba
        o2, n2 = o * seizon_o, n * seizon_n
        # 小数の年齢は、下の階級と上の階級に按分する
        shita = int(math.floor(atarashii))
        wa = atarashii - shita
        for toshi2, omomi in ((shita, 1.0 - wa), (shita + 1, wa)):
            if omomi <= 0 or toshi2 < 65:
                continue
            for (lo, hi), na in zip(waku, C.KAIKYU):
                if lo <= toshi2 <= hi:
                    out[na][0] += o2 * omomi
                    out[na][1] += n2 * omomi
                    out[na][2] += (o2 + n2) * omomi
                    break
    return {k: tuple(v) for k, v in out.items()}


def awaseru(okutta, mieruka):
    """6階級の姿を、見える化の3区分の合計に比例調整する。"""
    out = {}
    for na, kaikyu, _ in KUBUN:
        wa = sum(okutta[k][2] for k in kaikyu)
        mato = mieruka[na]
        hi = mato / wa if wa > 0 else 0.0
        for k in kaikyu:
            o, n, g = okutta[k]
            out[k] = (o * hi, n * hi, g * hi)
    return out


def suikei(jinko6):
    """6階級の人口から認知症・ＭＣＩの人数を積み上げる（性年齢階級別）。"""
    out = {}
    for shihyo in ("認知症", "ＭＣＩ"):
        r = C.YUBYO[shihyo]
        wa = 0.0
        meisai = []
        for i, na in enumerate(C.KAIKYU):
            o, n, g = jinko6[na]
            v = o * r["男性"][i] / 100.0 + n * r["女性"][i] / 100.0
            wa += v
            meisai.append((na, g, v))
        out[shihyo] = {"計": wa, "meisai": meisai}
    return out


def main():
    futeki = []

    def judge(ok, mi, setsu=""):
        print(("   ○ " if ok else "   × ") + mi + (("　" + setsu)
                                                   if setsu else ""))
        if not ok:
            futeki.append(mi)

    print("認知症の人の数の将来の推計（令和8年10月19日・暫定）")
    print("　川崎町 高齢者保健福祉計画・第10期介護保険事業計画")
    print()
    print("  ⚠ **暫定である。** 有病率は6階級で公表されているのに対し、")
    print("    町からご提供いただいた将来の人口は3区分しかない。")
    print("    見える化システムの「自然体推計の計算過程確認シート」を")
    print("    出力いただければ、按分は要らなくなる。")
    print()

    mieruka, retsu = yomu_sokatsu()
    shibou = yomu_shibou()
    genzai = C.machi_kaikyu()
    korei_ima = sum(v[2] for v in genzai.values())

    print("  ── 町からご提供いただいた将来の第1号被保険者数（3区分）")
    print(f"  {'年度':<12}{'第1号':>8}{'65〜74':>9}{'75〜84':>9}{'85〜':>9}"
          f"{'85歳以上の割合':>18}")
    for nendo in ["令和9年度"] + list(NENSU):
        m = mieruka[nendo]
        print(f"  {nendo:<12}{m['第1号']:>8,}"
              f"{m['前期（65〜74歳）']:>9,}{m['後期（75〜84歳）']:>9,}"
              f"{m['後期（85歳〜）']:>9,}"
              f"{m['後期（85歳〜）'] / m['第1号'] * 100:>16.1f}％")
    print(f"  （参考）令和8年6月末の住民基本台帳　65歳以上 {korei_ima:,}人"
          f"／65〜74歳 "
          f"{genzai['65〜69歳'][2] + genzai['70〜74歳'][2]:,}人"
          f"／75〜84歳 "
          f"{genzai['75〜79歳'][2] + genzai['80〜84歳'][2]:,}人"
          f"／85歳〜 "
          f"{genzai['85〜89歳'][2] + genzai['90歳以上'][2]:,}人")
    print("  ⚠ 将来は**第1号被保険者数**、現在は**住民基本台帳**であり"
          "母集団が違う（令和8年度で48人・1.5％）。")
    print()

    # ── 方法ごとの推計
    kekka = {}
    for nendo, nen in NENSU.items():
        m = {na: mieruka[nendo][na] for na, _, _ in KUBUN}
        # Ａ　現在の構成で固定
        a6 = awaseru(genzai, m)
        # Ｃ1　コーホートを送る（死亡率なし）
        c1 = awaseru(okuru(nen, None), m)
        # Ｃ2　コーホートを送る（死亡率あり）　← 正本
        c2raw = okuru(nen, shibou)
        c2 = awaseru(c2raw, m)
        kekka[nendo] = {"Ａ": (a6, suikei(a6)), "Ｃ1": (c1, suikei(c1)),
                        "Ｃ2": (c2, suikei(c2)), "c2raw": c2raw,
                        "第1号": mieruka[nendo]["第1号"]}

    print("  ── 認知症の人の数（方法別）")
    print(f"  {'年度':<12}{'Ａ 現在固定':>12}{'Ｃ1 送る(死亡率なし)':>22}"
          f"{'Ｃ2 送る(死亡率あり)':>22}{'簡便法':>10}")
    for nendo in NENSU:
        k = kekka[nendo]
        kan = k["第1号"] * ZENKOKU_SHORAI[nendo] / 100.0
        print(f"  {nendo:<12}"
              f"{k['Ａ'][1]['認知症']['計']:>11.0f}人"
              f"{k['Ｃ1'][1]['認知症']['計']:>21.0f}人"
              f"{k['Ｃ2'][1]['認知症']['計']:>21.0f}人"
              f"{kan:>9.0f}人")
    print()

    print("  ── ▶ 本スクリプトの正本（方法Ｃ2）")
    print(f"  {'年度':<12}{'65歳以上':>10}{'認知症':>9}{'割合':>9}"
          f"{'ＭＣＩ':>9}{'割合':>9}{'合計':>9}{'割合':>9}")
    hyo = [("令和8年度（現在）", korei_ima,
            C.santei(genzai, "認知症")["sei"],
            C.santei(genzai, "ＭＣＩ")["sei"])]
    for nendo in NENSU:
        k = kekka[nendo]
        hyo.append((nendo, k["第1号"], k["Ｃ2"][1]["認知症"]["計"],
                    k["Ｃ2"][1]["ＭＣＩ"]["計"]))
    for na, j, n, m2 in hyo:
        print(f"  {na:<12}{j:>9,}人{n:>8.0f}人{n / j * 100:>8.1f}％"
              f"{m2:>8.0f}人{m2 / j * 100:>8.1f}％"
              f"{n + m2:>8.0f}人{(n + m2) / j * 100:>8.1f}％")
    print("  （参考）全国の認知症有病率　令和4年12.3％／2030年14.2％／"
          "2040年14.9％")
    print()

    print("  ── 85歳以上の中身（窪みの通過）")
    print(f"  {'年度':<12}{'85〜89歳':>11}{'90歳以上':>11}"
          f"{'90歳以上の割合':>18}")
    o, n, g = genzai["85〜89歳"]
    o2, n2, g2 = genzai["90歳以上"]
    print(f"  {'令和8年度':<12}{g:>10,.0f}人{g2:>10,.0f}人"
          f"{g2 / (g + g2) * 100:>16.1f}％")
    for nendo in NENSU:
        j = kekka[nendo]["Ｃ2"][0]
        a, b = j["85〜89歳"][2], j["90歳以上"][2]
        print(f"  {nendo:<12}{a:>10,.0f}人{b:>10,.0f}人"
              f"{b / (a + b) * 100:>16.1f}％")
    print("  ⚠ **90歳以上の割合は令和17年度にいったん下がる。**")
    print("    終戦前後の出生の落ち込み（令和8年に79〜89歳・1歳あたり約70人。")
    print("    直前の団塊は1歳あたり約171人）が90歳以上の階級を通る時期に")
    print("    当たるためである。その後に団塊が85歳以上へ入る。")
    print()

    # ══════ 検査1　送らない（0年）と現在に戻るか
    print("  ── 検査1　送る年数を0にすると現在の姿に戻るか")
    rei = okuru(0.0, None)
    chigai = max(abs(rei[na][2] - genzai[na][2]) for na in C.KAIKYU)
    judge(chigai < 0.5, f"6階級が現在の人口と一致（最大の差 {chigai:.2f}人）")
    rei2 = suikei(rei)
    moto = C.santei(genzai, "認知症")["sei"]
    judge(abs(rei2["認知症"]["計"] - moto) < 0.5,
          f"認知症が現在の推計と一致（{rei2['認知症']['計']:.1f} 対 "
          f"{moto:.1f}人）")

    # ══════ 検査2　3区分の合計が見える化と一致するか
    print("  ── 検査2　3区分の合計が町のご提供の値と一致するか")
    ok = True
    for nendo in NENSU:
        j = kekka[nendo]["Ｃ2"][0]
        for na, kaikyu, _ in KUBUN:
            wa = sum(j[k][2] for k in kaikyu)
            if abs(wa - mieruka[nendo][na]) > 0.5:
                ok = False
                print(f"      × {nendo} {na} {wa:.1f} 対 "
                      f"{mieruka[nendo][na]}")
    judge(ok, "9つの区分すべてで一致（比例調整が効いている）")

    # ══════ 検査3　推計値が人口を超えないか
    print("  ── 検査3　推計値が人口を超えていないか")
    ok = True
    for nendo in NENSU:
        for shihyo in ("認知症", "ＭＣＩ"):
            for na, g3, v in kekka[nendo]["Ｃ2"][1][shihyo]["meisai"]:
                if v > g3 + 1e-6:
                    ok = False
    judge(ok, "いずれの階級でも推計値が人口以下")
    for nendo in NENSU:
        n = kekka[nendo]["Ｃ2"][1]["認知症"]["計"]
        m2 = kekka[nendo]["Ｃ2"][1]["ＭＣＩ"]["計"]
        judge(n + m2 < kekka[nendo]["第1号"],
              f"{nendo}　認知症＋ＭＣＩ {n + m2:.0f}人 ＜ 65歳以上 "
              f"{kekka[nendo]['第1号']:,}人")

    # ══════ 検査4　死亡率の感度
    print("  ── 検査4　死亡率の置き方にどれだけ振れるか")
    for nendo in NENSU:
        atai = []
        for b in (0.7, 1.0, 1.4):
            j = awaseru(okuru(NENSU[nendo], shibou, b),
                        {na: mieruka[nendo][na] for na, _, _ in KUBUN})
            atai.append(suikei(j)["認知症"]["計"])
        haba = (max(atai) - min(atai)) / atai[1] * 100
        judge(haba < 5.0,
              f"{nendo}　死亡率を0.7〜1.4倍に振っても "
              f"{min(atai):.0f}〜{max(atai):.0f}人（振れ {haba:.1f}％）")

    # ══════ 検査5　方法の違いによる振れ
    print("  ── 検査5　区分の中の内訳の置き方による振れ")
    for nendo in NENSU:
        k = kekka[nendo]
        atai = [k[m][1]["認知症"]["計"] for m in ("Ａ", "Ｃ1", "Ｃ2")]
        haba = (max(atai) - min(atai)) / k["Ｃ2"][1]["認知症"]["計"] * 100
        print(f"   ・{nendo}　Ａ {atai[0]:.0f}人／Ｃ1 {atai[1]:.0f}人／"
              f"Ｃ2 {atai[2]:.0f}人　振れ {haba:.1f}％")
    print("     ⚠ **死亡率の置き方（検査4）より、内訳の置き方の方が"
          "大きく効く。**")
    print("       だからこそ、6階級の将来人口を町からいただくのが正道である。")
    judge(True, "方法による振れを数えた")

    # ══════ 検査6　区分ごとにどれだけ効くか
    print("  ── 検査6　3区分のうちどこが効くか（令和22年度）")
    k = kekka["令和22年度"]
    for na, kaikyu, _ in KUBUN:
        atai = []
        for m in ("Ａ", "Ｃ1", "Ｃ2"):
            j = k[m][0]
            r = C.YUBYO["認知症"]
            v = 0.0
            for kk in kaikyu:
                i = C.KAIKYU.index(kk)
                v += (j[kk][0] * r["男性"][i] + j[kk][1] * r["女性"][i]) / 100
            atai.append(v)
        print(f"   ・{na:<16}Ａ {atai[0]:>5.0f}人／Ｃ1 {atai[1]:>5.0f}人／"
              f"Ｃ2 {atai[2]:>5.0f}人　幅 {max(atai) - min(atai):.0f}人")
    judge(True, "区分ごとの効きを数えた")

    # ══════ 検査7　簡便法が現実と逆の形を出さないか
    print("  ── 検査7　簡便法（第1号×全国の有病率）の当否")
    kan = {n: kekka[n]["第1号"] * ZENKOKU_SHORAI[n] / 100.0 for n in NENSU}
    print(f"   ・簡便法　令和12年度 {kan['令和12年度']:.0f}人／"
          f"令和17年度 {kan['令和17年度']:.0f}人／"
          f"令和22年度 {kan['令和22年度']:.0f}人")
    gyaku = kan["令和22年度"] < kan["令和12年度"]
    print(f"   ・Ｃ2　　 令和12年度 "
          f"{kekka['令和12年度']['Ｃ2'][1]['認知症']['計']:.0f}人／"
          f"令和17年度 {kekka['令和17年度']['Ｃ2'][1]['認知症']['計']:.0f}人／"
          f"令和22年度 {kekka['令和22年度']['Ｃ2'][1]['認知症']['計']:.0f}人")
    judge(gyaku, "簡便法では令和22年度が令和12年度より**少なく**なる",
          "⚠ 85歳以上は605人→893人と増えるのに減る。"
          "年齢構成を無視したことによる見かけの減であり、採れない")
    nobi = all(kekka[a]["Ｃ2"][1]["認知症"]["計"]
               < kekka[b]["Ｃ2"][1]["認知症"]["計"]
               for a, b in (("令和12年度", "令和17年度"),
                            ("令和17年度", "令和22年度")))
    judge(nobi, "Ｃ2は年を追って増える（85歳以上の増加と整合する）")

    print()
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        for m in futeki:
            print(f"      {m}")
        sys.exit(1)
    print("   ○ 検査1〜7　不適合0件")
    print()
    print("   ▶ 暫定の推計値（方法Ｃ2）")
    for nendo in NENSU:
        k = kekka[nendo]
        n = k["Ｃ2"][1]["認知症"]["計"]
        m2 = k["Ｃ2"][1]["ＭＣＩ"]["計"]
        print(f"      {nendo}　認知症 {n:.0f}人（{n / k['第1号'] * 100:.1f}％）"
              f"／ＭＣＩ {m2:.0f}人（{m2 / k['第1号'] * 100:.1f}％）")
    print()
    print("   ⚠ **計画書に載せる前に、見える化システムの")
    print("     「自然体推計の計算過程確認シート」を町にご出力いただきたい。**")
    print("     性・年齢5歳階級別の第1号被保険者数（令和12・17・22年度）が")
    print("     そこにあり、按分が要らなくなる。")


if __name__ == "__main__":
    main()
