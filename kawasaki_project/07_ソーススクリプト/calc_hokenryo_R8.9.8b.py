# -*- coding: utf-8 -*-
"""
保険料試算の再是正（令和8年9月8日・第2版）
― 令和6年度の年報（決算・所得段階別）の受領により、前提を実績に置き換える

令和8年9月8日に受領した『年報データ_2024_川崎町.xlsx』（令和6年度の介護保険事業状況報告・47シート）
により、次の3点が実データで確定した。同日午前の試算（calc_hokenryo_R8.9.8.py）を是正する。

■ 是正1　「保険給付費見込額」は総給付費ではなく標準給付費見込額であった
  令和6年度の様式4（介護保険特別会計経理状況）の歳出は
      介護サービス等諸費        982,753,997
    ＋ 介護予防サービス等諸費     27,005,382   （＝総給付費 1,009,759,379）
    ＋ 高額介護サービス等費       30,238,272
    ＋ 高額医療合算介護サービス等費 1,786,877
    ＋ 特定入所者介護サービス等費  50,836,225
    ＋ 審査支払手数料              632,700
    ＝ 1,093,253,453 円
  であり、町提供実績データ05シートの「保険給付費見込額 令和6年度 1,093,253千円」と完全に一致する。
  すなわち **05シートの値は既に標準給付費見込額(A)そのもの** である。
  同日午前の試算は、これを総給付費とみなして特定入所者等の比率を上乗せしていたため、
  Aを約8％過大に計算していた。

■ 是正2　所得段階別加入割合補正係数は1.0を上回る
  令和6・7年度の様式1（所得段階別第1号被保険者数）により、13段階の実人数が判明した。
  乗率を乗じた補正係数は 令和6年度 0.99749／令和7年度 1.01375（2か年平均 1.00559）。
  第9期の将来推計総括表の 0.9667 は、13段階化前の9段階の分布で推計したもので、
  第10〜13段階（令和6年度76人・令和7年度78人、乗率1.9〜2.4）が0のまま計算されていた。
  同日午前の試算で用いた 0.9685 も同じ理由で過小である。

■ 是正3　調整交付金の交付実績が判明した
  令和6年度の様式4の歳入「調整交付金 53,478,000円」は、標準給付費見込額に対し 4.891％。
  第9期計画の見込率（令和8年度3.92％）より高い。確認事項No.11の一部が確定した。
"""

# ------------------------------------------------------------ 令和6年度の決算（様式4）
R6 = dict(
    kaigo_shohi=982_753_997,          # 介護サービス等諸費
    yobo_shohi=27_005_382,            # 介護予防サービス等諸費
    kogaku=30_238_272,                # 高額介護サービス等費
    gassan=1_786_877,                 # 高額医療合算介護サービス等費
    tokutei=50_836_225,               # 特定入所者介護サービス等費
    tesuryo=632_700,                  # 審査支払手数料
    # 地域支援事業
    yobo_seikatsu=6_782_146,          # 介護予防・生活支援サービス事業費
    ippan_yobo=4_308_049,             # 一般介護予防事業費
    hokatsu_nini=30_792_643,          # 包括的支援事業・任意事業
    chiiki_sonota=4_200,
    # 歳入
    chosei_kofukin=53_478_000,        # 調整交付金
    kyoka_suishin=1_277_000,          # 保険者機能強化推進交付金
    kyoka_doryoku=2_461_000,          # 保険者努力支援交付金
    kikin_hoyu=146_700_000,           # 介護給付費準備基金保有額（年度末）
)

# 標準給付費見込額(A) の実績・見込（千円）＝ 町提供実績データ05シート（令和6年度の決算と一致）
A_JISSEKI = {"R6": 1_093_253, "R7": 1_085_203, "R8": 1_127_782}   # R8は当初予算・見込

# 地域支援事業費（千円）
B_JISSEKI = {"R6": 41_871, "R7": 42_040, "R8": 51_906}
# 令和6年度決算による総合事業費の割合（調整交付金相当額の算定基礎に入る部分）
SOGO_RATIO = (R6["yobo_seikatsu"] + R6["ippan_yobo"]) / (
    R6["yobo_seikatsu"] + R6["ippan_yobo"] + R6["hokatsu_nini"] + R6["chiiki_sonota"])

# 所得段階別第1号被保険者数（年報 様式1 所得段階別）と川崎町の乗率
JORITSU = [0.455, 0.685, 0.690, 0.900, 1.000, 1.200, 1.300, 1.500, 1.700,
           1.900, 2.100, 2.300, 2.400]
DANKAI = {
    "R6": [412, 308, 304, 383, 699, 375, 419, 218, 67, 25, 13, 7, 31],
    "R7": [383, 298, 305, 343, 679, 410, 427, 232, 85, 26, 17, 7, 28],
}

HIHOKENSHA_10KI = [3_214, 3_199, 3_184]     # 第1号被保険者数の推計（実績トレンドの外挿）
KIKIN_R8 = 152_500_000                      # 準備基金残高（令和8年度末見込）
SHUNORITSU = 0.96                           # 予定保険料収納率


def line(c="-", n=88):
    print(c * n)


def hosei_keisu(year):
    n = DANKAI[year]
    tot = sum(n)
    hosei = sum(a * b for a, b in zip(n, JORITSU))
    return hosei / tot, tot, hosei


def kijungaku(L, shunoritsu, hosei_ninzu):
    return L / shunoritsu / hosei_ninzu / 12


def suikei(A_sen, B_sen, futan, chosei_ritsu, hosei_ninzu, torikuzushi=0, kyoka=0):
    """千円単位の A・B から保険料基準額（月額・円）を求める。"""
    A = A_sen * 1000
    B = B_sen * 1000
    sogo = B * SOGO_RATIO
    D = (A + B) * futan
    E = (A + sogo) * 0.05
    I = A * chosei_ritsu
    L = D + E - I - torikuzushi - kyoka
    return dict(A=A, B=B, D=D, E=E, I=I, L=L,
                tsuki=kijungaku(L, SHUNORITSU, hosei_ninzu))


def main():
    print("■ 1. 令和6年度の決算（様式4）による標準給付費見込額の検算")
    line()
    A6 = (R6["kaigo_shohi"] + R6["yobo_shohi"] + R6["kogaku"]
          + R6["gassan"] + R6["tokutei"] + R6["tesuryo"])
    sokyufu = R6["kaigo_shohi"] + R6["yobo_shohi"]
    print(f"  総給付費（介護サービス等諸費＋介護予防サービス等諸費）　{sokyufu:>15,}円")
    print(f"  ＋ 高額介護サービス等費　　　　　　　　　　　　　　　　{R6['kogaku']:>15,}円")
    print(f"  ＋ 高額医療合算介護サービス等費　　　　　　　　　　　　{R6['gassan']:>15,}円")
    print(f"  ＋ 特定入所者介護サービス等費　　　　　　　　　　　　　{R6['tokutei']:>15,}円")
    print(f"  ＋ 審査支払手数料　　　　　　　　　　　　　　　　　　　{R6['tesuryo']:>15,}円")
    print(f"  ＝ 標準給付費見込額(A)　　　　　　　　　　　　　　　　 {A6:>15,}円")
    print(f"  町提供実績データ05シート「保険給付費見込額 令和6年度」 {A_JISSEKI['R6']*1000:>15,}円"
          f"　差 {A6 - A_JISSEKI['R6']*1000:+,}円")
    print()
    chiiki6 = (R6["yobo_seikatsu"] + R6["ippan_yobo"]
               + R6["hokatsu_nini"] + R6["chiiki_sonota"])
    print(f"  地域支援事業費（決算）　{chiiki6:>15,}円"
          f"（町提供データ {B_JISSEKI['R6']*1000:,}円・差 {chiiki6 - B_JISSEKI['R6']*1000:+,}円）")
    print(f"  　うち総合事業費（介護予防・生活支援＋一般介護予防）"
          f"　{R6['yobo_seikatsu'] + R6['ippan_yobo']:>12,}円　"
          f"（構成比 {SOGO_RATIO*100:.1f}％）")
    print(f"  　うち包括的支援事業・任意事業　　　　　　　　　　　"
          f"　{R6['hokatsu_nini']:>12,}円")
    print()
    print(f"  調整交付金（歳入）　{R6['chosei_kofukin']:>15,}円"
          f"　→ 標準給付費に対し {R6['chosei_kofukin']/A6*100:.3f}％")
    print(f"  保険者機能強化推進交付金 {R6['kyoka_suishin']:>10,}円"
          f" ＋ 保険者努力支援交付金 {R6['kyoka_doryoku']:>10,}円"
          f" ＝ {R6['kyoka_suishin']+R6['kyoka_doryoku']:,}円")
    print(f"  介護給付費準備基金保有額（令和6年度末）　{R6['kikin_hoyu']:>15,}円")
    print()

    print("■ 2. 所得段階別加入割合補正係数（年報 様式1 所得段階別による実績）")
    line()
    print(f"  {'段階':6s} {'乗率':>6s} {'令和6年度':>10s} {'令和7年度':>10s}")
    for i, r in enumerate(JORITSU):
        print(f"  第{i+1:<2d}段階 {r:>6.3f} {DANKAI['R6'][i]:>10,} {DANKAI['R7'][i]:>10,}")
    for y in ("R6", "R7"):
        k, tot, hosei = hosei_keisu(y)
        print(f"  令和{6 if y=='R6' else 7}年度　合計 {tot:,}人／補正後 {hosei:,.1f}人／"
              f"補正係数 {k:.5f}")
    k6, t6, h6 = hosei_keisu("R6")
    k7, t7, h7 = hosei_keisu("R7")
    k_avg = (h6 + h7) / (t6 + t7)
    print(f"  2か年平均の補正係数 {k_avg:.5f}")
    print(f"  （第9期 将来推計総括表の 0.9667 は13段階化前の9段階の分布によるもので過小。"
          f"当日午前に用いた 0.9685 も同様）")
    print()

    print("■ 3. 第10期 保険料基準額の再是正試算")
    line()
    hihokensha = sum(HIHOKENSHA_10KI)
    print(f"  第1号被保険者数（3か年計・実績トレンドの外挿）　{hihokensha:,}人")
    for label, k in (("令和7年度の係数", k7), ("2か年平均の係数", k_avg)):
        print(f"  {label} {k:.5f} → 補正後被保険者数 {hihokensha*k:,.0f}人")
    print()

    CASES = [("低位", 0.005, 0.000), ("中位", 0.015, 0.015), ("高位", 0.025, 0.025)]
    chiiki_nen = sum(B_JISSEKI.values()) / 3

    for keisu_label, keisu in (("2か年平均 1.00559", k_avg), ("令和7年度 1.01375", k7)):
        hosei_ninzu = hihokensha * keisu
        print(f"  【補正係数 {keisu_label} ／ 補正後被保険者数 {hosei_ninzu:,.0f}人】")
        print(f"  {'ケース':8s} {'負担割合':>8s} {'調整交付金':>10s}"
              f" {'A取崩なし':>10s} {'B50%取崩':>10s} {'C全額取崩':>10s}")
        line("-", 68)
        for futan in (0.23, 0.24):
            for chosei in (0.0392, 0.0489):
                for name, shizen, kaitei in CASES:
                    if name != "中位":
                        continue
                    v = A_JISSEKI["R8"]
                    A3 = 0.0
                    for i in range(3):
                        v = v * (1 + shizen) * ((1 + kaitei) if i == 0 else 1)
                        A3 += v
                    out = []
                    for wari in (0.0, 0.5, 1.0):
                        r = suikei(A3, chiiki_nen * 3, futan, chosei, hosei_ninzu,
                                   KIKIN_R8 * wari)
                        out.append(r["tsuki"])
                    print(f"  {name:8s} {futan*100:>7.0f}% {chosei*100:>9.2f}%"
                          f" {out[0]:>10,.0f} {out[1]:>10,.0f} {out[2]:>10,.0f}")
        print()

    print("■ 4. 試算値の変遷（中位・基金50％取崩・負担割合23％）")
    line()
    hosei_ninzu = hihokensha * k_avg
    v = A_JISSEKI["R8"]
    A3 = 0.0
    for i in range(3):
        v = v * 1.015 * (1.015 if i == 0 else 1)
        A3 += v
    r = suikei(A3, chiiki_nen * 3, 0.23, 0.0489, hosei_ninzu, KIKIN_R8 * 0.5)
    print(f"  9月2日　　　（分母を9,740人と誤り）　　　　　　　　 7,023円")
    print(f"  9月8日午前（分母9,295人・Aを8％過大）　　　　　　 8,025円")
    print(f"  9月8日午後（本試算・実績ベース）　　　　　　　　　{r['tsuki']:>6,.0f}円")
    print()
    print("■ 5. 中位ケースの内訳（負担割合23％・調整交付金4.89％・50％取崩）")
    line()
    print(f"  標準給付費見込額(A) 3か年　　　{r['A']:>15,.0f}円")
    print(f"  地域支援事業費(B) 3か年　　　　{r['B']:>15,.0f}円")
    print(f"  第1号被保険者負担分相当額(D)　 {r['D']:>15,.0f}円")
    print(f"  ＋ 調整交付金相当額(E)　　　　 {r['E']:>15,.0f}円")
    print(f"  − 調整交付金見込額(I)　　　　　{r['I']:>15,.0f}円")
    print(f"  − 準備基金取崩額　　　　　　　 {KIKIN_R8*0.5:>15,.0f}円")
    print(f"  ＝ 保険料収納必要額(L)　　　　 {r['L']:>15,.0f}円")
    print(f"  ÷ 予定収納率 {SHUNORITSU} ÷ 補正後被保険者数 {hosei_ninzu:,.0f}人 ÷ 12")
    print(f"  ＝ 保険料基準額（月額）　　　　{r['tsuki']:>15,.0f}円")
    print()
    print("※ 本試算は、見える化システムの自然体推計に代わるものではない。")
    print("　 サービス種類別の見込量・給付費はシステムが年報・月報の実績から推計する。")
    print("　 確定値は第1回推計（令和8年9月30日〜10月5日提出）以降に置き換える。")


if __name__ == "__main__":
    main()
