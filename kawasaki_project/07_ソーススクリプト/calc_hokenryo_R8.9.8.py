# -*- coding: utf-8 -*-
"""
保険料算定の再検証（令和8年9月8日）
― 第9期の「将来推計総括表」（令和8年9月7日受領）により算式を確定し、第10期試算を是正する

令和8年9月7日に町から受領した『【川崎町】第９期　将来推計総括表.xlsx』
（地域包括ケア「見える化」システムの出力・出力日 令和5年10月5日・
　推計パターン名「第9期川崎町介護保険（R5.8.12）」）により、
システムが用いる算式と係数が判明した。これを用いて

  (1) 総括表の第9期基準額 7,644.21円／月 を再現できるか検証する
  (2) 令和8年9月2日にお示しした第10期の概算試算を是正する

■ 判明した算式（総括表 5_保険料推計 による）

  標準給付費見込額(A) ＝ 総給付費 ＋ 特定入所者介護サービス費等 ＋ 高額介護サービス費等
                        ＋ 高額医療合算介護サービス費等 ＋ 算定対象審査支払手数料
  地域支援事業費(B)
  第1号被保険者負担分相当額(D) ＝ (A ＋ B) × 第1号被保険者負担割合
  調整交付金相当額(E)          ＝ (A ＋ 介護予防・日常生活支援総合事業費) × 5％
      ※ 包括的支援事業費は調整交付金の対象外である点に注意（B 全体の5％ではない）
  調整交付金見込額(I)          ＝ A × 調整交付金見込交付割合(H)
  保険料収納必要額(L)          ＝ D ＋ E − I − 準備基金取崩額 − 保険者機能強化推進交付金等
  所得段階別加入割合補正後被保険者数(C) ＝ Σ(所得段階別被保険者数 × 基準額に対する割合)
  保険料基準額（月額）         ＝ L ÷ 予定保険料収納率 ÷ C ÷ 12

■ 令和8年9月2日の試算の誤り（本スクリプトで是正するもの）

  1. 補正後被保険者数(C)を「第1号被保険者数 × 1.0149」としていた。
     総括表では C ＝ 9,546.1人、第1号被保険者数 ＝ 9,875人であり、比は 0.9667。
     すなわち **補正後被保険者数は第1号被保険者数より少ない**。逆方向に誤っていた。
     （9,875人は「補正後被保険者数」ではなく、推計上の第1号被保険者数の3か年計であった）
  2. 保険者機能強化推進交付金等の交付見込額を10,800千円控除していたが、
     第9期の総括表では 0 である。
  3. 調整交付金相当額を「標準給付費×5％」としていたが、正しくは
     「(標準給付費＋総合事業費)×5％」である。
  4. 第1号被保険者負担割合を23％固定としていたが、総括表の係数表では
     令和12年度が24％であり、第10期（令和9〜11年度）も24％となる可能性が高い。
"""

# ------------------------------------------------------------ 第9期 総括表の値
K9 = dict(
    # 標準給付費見込額(A) の内訳（円・3か年計）
    sokyufu=3_189_830_000,          # 総給付費（財政影響額調整後）
    tokutei=162_299_291,            # 特定入所者介護サービス費等
    kogaku=82_949_253,              # 高額介護サービス費等
    gassan=7_410_739,               # 高額医療合算介護サービス費等
    tesuryo=2_094_300,              # 算定対象審査支払手数料（@60円×34,905件）
    # 地域支援事業費(B)
    sogo=25_560_000,                # 介護予防・日常生活支援総合事業費
    hokatsu_unei=39_000_000,        # 包括的支援事業（地域包括支援センターの運営）
    hokatsu_jujitsu=27_000_000,     # 包括的支援事業（社会保障充実分）
    # 係数
    futan_wariai=0.23,              # 第1号被保険者負担割合
    chosei_mikomi=146_178_000,      # 調整交付金見込額(I)
    kyoka_kofu=0,                   # 保険者機能強化推進交付金等の交付見込額
    torikuzushi=0,                  # 準備基金取崩額
    shunoritsu=0.96,                # 予定保険料収納率
    hosei_ninzu=9_546.1,            # 所得段階別加入割合補正後被保険者数(C)
    hihokensha=9_875,               # 第1号被保険者数（3か年計・推計値）
    kohyo_kijun=7_644.21,           # 総括表の保険料基準額（月額）
)

# 第9期 総括表の所得段階別被保険者数（3か年計）と標準乗率
DANKAI_K9 = [
    ("第1段階", 1_356, 0.445), ("第2段階", 890, 0.68), ("第3段階", 932, 0.69),
    ("第4段階", 1_217, 0.90), ("第5段階", 2_088, 1.00), ("第6段階", 1_363, 1.20),
    ("第7段階", 1_152, 1.30), ("第8段階", 565, 1.50), ("第9段階", 312, 1.70),
]
# 川崎町が第9期で実際に条例で定めた13段階の乗率（第9期計画書95頁）
JORITSU_13 = [0.455, 0.685, 0.69, 0.90, 1.00, 1.20, 1.30, 1.50, 1.70, 1.90, 2.10, 2.30, 2.40]


def line(c="-", n=82):
    print(c * n)


def hyojun_kyufu(d):
    return d["sokyufu"] + d["tokutei"] + d["kogaku"] + d["gassan"] + d["tesuryo"]


def chiiki_shien(d):
    return d["sogo"] + d["hokatsu_unei"] + d["hokatsu_jujitsu"]


def shunyu_hitsuyogaku(A, B, sogo, futan, I, kyoka=0, torikuzushi=0):
    D = (A + B) * futan
    E = (A + sogo) * 0.05
    return D + E - I - kyoka - torikuzushi, D, E


def kijungaku(L, shunoritsu, hosei_ninzu):
    return L / shunoritsu / hosei_ninzu / 12


# ============================================================ 1. 第9期の再現
def check_k9():
    print("■ 1. 第9期 将来推計総括表（令和5年10月5日出力）の再現")
    line()
    A = hyojun_kyufu(K9)
    B = chiiki_shien(K9)
    print(f"  標準給付費見込額(A)　　　　　　{A:>15,.0f}円　（総括表 3,444,583,583円）")
    print(f"  地域支援事業費(B)　　　　　　　{B:>15,.0f}円　（総括表 　 91,560,000円）")
    L, D, E = shunyu_hitsuyogaku(A, B, K9["sogo"], K9["futan_wariai"],
                                 K9["chosei_mikomi"], K9["kyoka_kofu"], K9["torikuzushi"])
    print(f"  第1号被保険者負担分相当額(D)　 {D:>15,.0f}円　（総括表 　813,313,000円）")
    print(f"  調整交付金相当額(E)　　　　　　{E:>15,.0f}円　（総括表 　173,507,000円）")
    print(f"  − 調整交付金見込額(I)　　　　　{K9['chosei_mikomi']:>15,.0f}円")
    print(f"  ＝ 保険料収納必要額(L)　　　　 {L:>15,.0f}円　（総括表 　840,642,000円）")
    m = kijungaku(L, K9["shunoritsu"], K9["hosei_ninzu"])
    print(f"  保険料基準額（月額）　　　　　 {m:>15,.2f}円　（総括表 　　　7,644.21円）")

    # 補正後被保険者数の検算
    c_std = sum(n * r for _, n, r in DANKAI_K9)
    print(f"  補正後被保険者数(C)の検算　　　{c_std:>15,.1f}人　（総括表 　　　9,546.1人）")
    hosei_hi = c_std / K9["hihokensha"]
    print(f"  　→ 補正係数 ＝ C ÷ 第1号被保険者数 ＝ {hosei_hi:.4f}"
          f"　（第1号被保険者数 {K9['hihokensha']:,}人）")

    # 13段階の乗率を当てはめた場合の補正係数
    wariai = [n / K9["hihokensha"] for _, n, _ in DANKAI_K9] + [0.0] * 4
    c13 = sum(w * r for w, r in zip(wariai, JORITSU_13))
    print(f"  　→ 川崎町の13段階の乗率を当てはめた場合の補正係数 {c13:.4f}")
    print()
    return hosei_hi, c13


# ============================================================ 2. 第10期の是正試算
# 令和8年9月2日の試算の前提（calc_hokenryo_R8.9.2.py）
KYUFU_R8 = 1_106_907_000            # 令和8年度の保険給付費見込額（町提供実績データ 05シート）
CHIIKI_JISSEKI = {"R6": 41_871_000, "R7": 42_040_000, "R8": 51_906_000}
HIHOKENSHA_10KI = [3_214, 3_199, 3_184]     # 第1号被保険者数の推計（実績トレンドの外挿）
KIKIN_R8 = 152_500_000                      # 準備基金残高（令和8年度末見込）

# 標準給付費見込額に対する各項目の比（第9期総括表による）
R_TOKUTEI = K9["tokutei"] / K9["sokyufu"]
R_KOGAKU = K9["kogaku"] / K9["sokyufu"]
R_GASSAN = K9["gassan"] / K9["sokyufu"]
R_TESURYO = K9["tesuryo"] / K9["sokyufu"]

CASES = [
    ("低位", 0.005, 0.000),
    ("中位", 0.015, 0.015),
    ("高位", 0.025, 0.025),
]


def suikei_10ki(shizen, kaitei, chiiki_nen, sogo_nen, futan, chosei_ritsu,
                shunoritsu, hosei_ninzu, torikuzushi=0):
    v = KYUFU_R8
    sokyufu = 0.0
    for i in range(3):
        v = v * (1 + shizen) * ((1 + kaitei) if i == 0 else 1)
        sokyufu += v
    A = sokyufu * (1 + R_TOKUTEI + R_KOGAKU + R_GASSAN + R_TESURYO)
    B = chiiki_nen * 3
    I = A * chosei_ritsu
    L, D, E = shunyu_hitsuyogaku(A, B, sogo_nen * 3, futan, I, 0, torikuzushi)
    return dict(A=A, B=B, D=D, E=E, I=I, L=L,
                tsuki=kijungaku(L, shunoritsu, hosei_ninzu))


def main():
    hosei_hi, c13 = check_k9()

    print("■ 2. 第10期の是正試算（令和8年9月8日版）")
    line()
    hihokensha = sum(HIHOKENSHA_10KI)
    hosei_ninzu = hihokensha * c13
    chiiki_nen = sum(CHIIKI_JISSEKI.values()) / 3
    # 総合事業費は第9期計画値（8,520千円/年）が続くと仮定する
    sogo_nen = K9["sogo"] / 3
    print(f"  第1号被保険者数（3か年計・実績トレンドの外挿）　{hihokensha:>10,}人")
    print(f"  補正係数（川崎町13段階の乗率・第9期の分布）　　 {c13:>10.4f}")
    print(f"  所得段階別加入割合補正後被保険者数(C)　　　　　 {hosei_ninzu:>10,.1f}人")
    print(f"  　（9月2日の試算では 9,740人としていた。差 {hosei_ninzu - 9_740:+,.0f}人）")
    print(f"  地域支援事業費（直近3か年平均）　　　　　　　　 {chiiki_nen:>10,.0f}円/年")
    print(f"  うち介護予防・日常生活支援総合事業費（第9期計画値）{sogo_nen:>8,.0f}円/年")
    print(f"  予定保険料収納率　　　　　　　　　　　　　　　　 {0.96:>10.2f}")
    print()

    print("  ケース別の保険料基準額（月額・円）")
    line()
    print(f"  {'ケース':34s} {'負担割合':>8s} {'調整交付金':>10s}"
          f" {'A取崩なし':>10s} {'B50%取崩':>10s} {'C全額取崩':>10s}")
    line()
    rows = {}
    for futan in (0.23, 0.24):
        for chosei in (0.0392, 0.05):
            for name, shizen, kaitei in CASES:
                out = []
                for wari in (0.0, 0.5, 1.0):
                    r = suikei_10ki(shizen, kaitei, chiiki_nen, sogo_nen, futan,
                                    chosei, 0.96, hosei_ninzu, KIKIN_R8 * wari)
                    out.append(r["tsuki"])
                rows[(futan, chosei, name)] = out
                label = f"{name}"
                print(f"  {label:34s} {futan*100:>7.0f}% {chosei*100:>9.2f}%"
                      f" {out[0]:>10,.0f} {out[1]:>10,.0f} {out[2]:>10,.0f}")
        print()

    print("■ 3. 9月2日の試算との比較（中位・50％取崩）")
    line()
    print(f"  9月2日の試算　　　　　　　　　　　　　　　　　　　 7,023円")
    for futan in (0.23, 0.24):
        for chosei in (0.0392, 0.05):
            v = rows[(futan, chosei, "中位")][1]
            print(f"  是正後（負担割合{futan*100:.0f}％・調整交付金{chosei*100:.2f}％）"
                  f"　　{v:>8,.0f}円　（{v - 7_023:+,.0f}円）")
    print()

    print("■ 4. 中位ケースの内訳（負担割合24％・調整交付金3.92％・50％取崩）")
    line()
    r = suikei_10ki(0.015, 0.015, chiiki_nen, sogo_nen, 0.24, 0.0392, 0.96,
                    hosei_ninzu, KIKIN_R8 * 0.5)
    print(f"  標準給付費見込額(A)　　　　　　{r['A']:>15,.0f}円"
          f"　（第9期総括表 3,444,584千円）")
    print(f"  地域支援事業費(B)　　　　　　　{r['B']:>15,.0f}円"
          f"　（第9期総括表 　 91,560千円）")
    print(f"  第1号被保険者負担分相当額(D)　 {r['D']:>15,.0f}円")
    print(f"  ＋ 調整交付金相当額(E)　　　　 {r['E']:>15,.0f}円")
    print(f"  − 調整交付金見込額(I)　　　　　{r['I']:>15,.0f}円")
    print(f"  − 準備基金取崩額　　　　　　　 {KIKIN_R8*0.5:>15,.0f}円")
    print(f"  ＝ 保険料収納必要額(L)　　　　 {r['L']:>15,.0f}円")
    print(f"  保険料基準額（月額）　　　　　 {r['tsuki']:>15,.0f}円")
    print()

    # ------------------------------------------------ 5. 確認事項No.14 の検証
    print("■ 5. 第9期計画書95頁の算出表との突合（確認事項No.14）")
    line()
    keikaku_nashi = 804_662_755      # 計画書「取崩なしの保険料収納必要額」
    keikaku_go = 712_662_755         # 計画書「取崩後の保険料収納必要額」
    keikaku_torikuzushi = 78_000_000  # 計画書「準備基金取崩額」
    sa = keikaku_nashi - keikaku_go
    print(f"  計画書の取崩なし収納必要額　　 {keikaku_nashi:>15,}円")
    print(f"  計画書の取崩後収納必要額　　　 {keikaku_go:>15,}円")
    print(f"  　→ 差引　　　　　　　　　　　 {sa:>15,}円")
    print(f"  計画書の準備基金取崩額　　　　 {keikaku_torikuzushi:>15,}円"
          f"　（差引との差 {sa - keikaku_torikuzushi:+,}円）")
    print()
    for label, ninzu in (("総括表の補正後被保険者数", K9["hosei_ninzu"]),
                         ("第1号被保険者数（誤った分母）", K9["hihokensha"])):
        m_nashi = kijungaku(keikaku_nashi, 0.96, ninzu)
        m_go = kijungaku(keikaku_go, 0.96, ninzu)
        print(f"  {label}（{ninzu:,.1f}人）を分母にした場合")
        print(f"    取崩なし {m_nashi:>8,.0f}円／月　　取崩後 {m_go:>8,.0f}円／月")
    print()
    ninzu_kohyo = keikaku_go / 0.96 / 78_000
    print(f"  公表値 年額78,000円 を取崩後の収納必要額から再現する補正後被保険者数")
    print(f"    ＝ {keikaku_go:,}円 ÷ 0.96 ÷ 78,000円 ＝ {ninzu_kohyo:,.1f}人")
    print(f"    （総括表〈令和5年10月時点〉の 9,546.1人 との差 "
          f"{ninzu_kohyo - K9['hosei_ninzu']:+,.1f}人・"
          f"{(ninzu_kohyo / K9['hosei_ninzu'] - 1) * 100:+.2f}％）")
    print()
    print("  → 計画書の取崩後収納必要額 712,662,755円 は、補正後被保険者数 約9,517人 を")
    print("　　 分母とすれば月額6,480円となり、公表値6,500円とおおむね一致する。")
    print("　　 差引92,000千円が実際の取崩額であり、計画書記載の78,000千円とは14,000千円ずれる。")
    print("　　 令和8年9月2日の試算で「実効取崩額65,223千円」としたのは、")
    print("　　 分母に第1号被保険者数9,875人を用いた誤りによるものであり、撤回する。")
    print()

    print("※ 本試算は、見える化システムの自然体推計に代わるものではない。")
    print("　 サービス種類別の見込量・給付費は、システムが年報・月報の実績から自動で推計する。")
    print("　 確定値は、システムでの第1回推計（令和8年9月30日〜10月5日提出）以降に置き換える。")


if __name__ == "__main__":
    main()
