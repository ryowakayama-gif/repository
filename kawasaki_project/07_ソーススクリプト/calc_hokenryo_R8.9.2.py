# -*- coding: utf-8 -*-
"""
第10期 介護保険料基準額の概算試算（令和8年9月2日時点・仮試算）

国基本指針の8ステップにより、第9期計画書（95頁）の保険料算出表と同じ算式で試算する。
入力は次の確定資料のみを用い、未確定の事項は第9期計画の前提を踏襲する。

  ・町提供実績データ R8.9.1受領（05シート：保険給付費・地域支援事業費・基金残高・収納率）
  ・介護保険事業状況報告 令和7年度（認定者・受給者・給付費）
  ・第9期計画書 95頁 保険料算出表

■ 第9期計画の算出表の内部不整合について
  計画書は「準備基金取崩額 78,000,000円」「取崩後の保険料収納必要額 712,662,755円」と
  記載するが、804,662,755 − 78,000,000 = 726,662,755 であり14,000千円が合わない
  （確認事項No.14）。また712,662,755では基準額が年75,175円となり、公表値78,000円と
  一致しない。公表値78,000円（月額6,500円）を再現する取崩額は 65,222,755円であり、
  これは基金残高131,700千円の49.5％にあたる。本試算では、この65,223千円を第9期の
  実効取崩額とみなし、同じ算式で第10期を試算する（第9期を正確に再現できる）。
"""

# ------------------------------------------------------------------ 確定データ
# 標準給付費（＝保険給付費。補完給付費を含む）千円
JISSEKI = {"R2": 1_044_362, "R3": 1_042_856, "R4": 1_040_193, "R5": 1_061_269,
           "R6": 1_093_253, "R7": 1_085_203, "R8": 1_106_907}
# 地域支援事業費 千円
CHIIKI = {"R2": 46_340, "R3": 32_079, "R4": 35_236, "R5": 38_914,
          "R6": 41_871, "R7": 42_040, "R8": 51_906}
# 第1号被保険者数 人
HIHOKEN = {"R2": 3228, "R3": 3255, "R4": 3285, "R5": 3270,
           "R6": 3261, "R7": 3240, "R8": 3229}
KIKIN_R8 = 152_500          # 準備基金残高 R8年度末見込（千円・05シート）
KYOKA_KOFU = 10_800         # 保険者機能強化推進交付金等（3年計・千円）

# 第9期計画書 95頁 保険料算出表（円）
K9 = dict(hyojun=3_294_553_410, chiiki=91_560_000, futan=778_806_084,
          chosei_soto=166_005_671, chosei_mikomi=129_349_000,
          kyoka=10_800_000, hitsuyo_torikuzushi_nashi=804_662_755,
          kikin=131_700_000, hosei_hihokensha=9_875, shunoritsu=0.96,
          kijun_nen=78_000, kijun_tsuki_santei=6_508)


def kijungaku(hitsuyo_sen, hosei_ninzu, shunoritsu):
    """保険料収納必要額（千円）から基準額（年額円・月額円）を求める。"""
    nen = hitsuyo_sen * 1000 / shunoritsu / hosei_ninzu
    return nen, nen / 12


def line(c="-", n=78):
    print(c * n)


# ============================================================ 1. 第9期の再現
def check_k9():
    print("■ 1. 第9期計画（95頁）の算出表の再現")
    line()
    nashi = K9["hitsuyo_torikuzushi_nashi"]
    n_a, m_a = kijungaku(nashi / 1000, K9["hosei_hihokensha"], K9["shunoritsu"])
    print(f"  取崩なしの収納必要額 {nashi:,}円")
    print(f"    → 基準額 年額{n_a:,.0f}円／月額{m_a:,.0f}円")
    # 公表値を再現する取崩額
    hitsuyo_kohyo = K9["kijun_nen"] * K9["hosei_hihokensha"] * K9["shunoritsu"]
    torikuzushi = nashi - hitsuyo_kohyo
    print(f"  公表値 年額{K9['kijun_nen']:,}円 を再現する収納必要額 {hitsuyo_kohyo:,.0f}円")
    print(f"    → 実効取崩額 {torikuzushi:,.0f}円"
          f"（基金{K9['kikin']:,}円の{torikuzushi / K9['kikin'] * 100:.1f}％）")
    n_b, m_b = kijungaku(hitsuyo_kohyo / 1000, K9["hosei_hihokensha"], K9["shunoritsu"])
    print(f"    → 基準額 年額{n_b:,.0f}円／月額{m_b:,.0f}円　"
          f"（計画の公表値 年額{K9['kijun_nen']:,}円・月額6,500円と一致）")
    print(f"  ※計画書記載の取崩額78,000,000円では 804,662,755−78,000,000＝726,662,755円となり、")
    print(f"    計画書記載の712,662,755円と14,000,000円が合わない（確認事項No.14）。")
    print()
    return torikuzushi


# ============================================================ 2. 第10期の前提
def hosei_hihokensha_10ki():
    """補正後被保険者数（3年計）。第9期計画と同じ補正の考え方が続くと仮定する。"""
    k9_jitsu = HIHOKEN["R6"] + HIHOKEN["R7"] + HIHOKEN["R8"]        # 9,730人
    k10 = [3214, 3199, 3184]                                        # ▲15人／年
    ratio = K9["hosei_hihokensha"] / k9_jitsu
    return sum(k10) * ratio, k9_jitsu, sum(k10), ratio


def suikei(shizen, kaitei, chiiki_nen, chosei_ritsu, shunoritsu, hosei_ninzu):
    """標準給付費・地域支援事業費から取崩なしの収納必要額（千円）を求める。"""
    base = JISSEKI["R8"]
    kyufu = []
    v = base
    for i in range(3):
        v = v * (1 + shizen) * ((1 + kaitei) if i == 0 else 1)
        kyufu.append(v)
    hyojun = sum(kyufu)
    chiiki = chiiki_nen * 3
    futan = (hyojun + chiiki) * 0.23
    soto = hyojun * 0.05
    mikomi = hyojun * chosei_ritsu
    hitsuyo = futan + soto - mikomi - KYOKA_KOFU
    return dict(kyufu=kyufu, hyojun=hyojun, chiiki=chiiki, futan=futan,
                soto=soto, mikomi=mikomi, hitsuyo=hitsuyo)


def main():
    check_k9()

    hosei, k9_jitsu, k10_ninzu, ratio = hosei_hihokensha_10ki()
    print("■ 2. 第10期の前提")
    line()
    print(f"  第1号被保険者数（実績 R6〜R8）　{k9_jitsu:,}人")
    print(f"  第1号被保険者数（推計 R9〜R11）　{k10_ninzu:,}人（3,214／3,199／3,184・▲15人/年）")
    print(f"  補正後被保険者数（第9期の補正比 {ratio:.4f} を適用）　{hosei:,.0f}人")
    print(f"  準備基金残高（R8年度末見込）　{KIKIN_R8:,}千円")
    print(f"  予定収納率　96.0％（第9期と同水準。実績は総合96.4〜97.0％）")
    chiiki_heikin = (CHIIKI["R6"] + CHIIKI["R7"] + CHIIKI["R8"]) / 3
    print(f"  地域支援事業費　直近3か年平均 {chiiki_heikin:,.0f}千円/年"
          f"（第9期計画は30,520千円/年）")
    print(f"  調整交付金見込率　3.7％（第9期計画のR8見込率3.67％）")
    print()

    CASES = [
        ("低位", 0.005, 0.000, chiiki_heikin, "自然増+0.5%/年・報酬改定なし"),
        ("中位", 0.015, 0.015, chiiki_heikin, "自然増+1.5%/年・報酬改定+1.5%（R9）"),
        ("高位", 0.025, 0.025, chiiki_heikin, "自然増+2.5%/年・報酬改定+2.5%（R9）"),
        ("中位・地域支援事業費が第9期計画水準", 0.015, 0.015, 30_520,
         "地域支援事業費を30,520千円/年に据置"),
    ]

    print("■ 3. 第10期 保険料基準額の概算試算（月額・円）")
    line()
    print(f"  {'ケース':38s} {'A 取崩なし':>10s} {'B 50%取崩':>10s} {'C 全額取崩':>10s}")
    line()
    results = {}
    for name, shizen, kaitei, chiiki_nen, memo in CASES:
        r = suikei(shizen, kaitei, chiiki_nen, 0.037, 0.96, hosei)
        row = []
        for wari in (0.0, 0.5, 1.0):
            h = r["hitsuyo"] - KIKIN_R8 * wari
            _, m = kijungaku(h, hosei, 0.96)
            row.append(m)
        results[name] = (r, row)
        print(f"  {name:38s} {row[0]:>10,.0f} {row[1]:>10,.0f} {row[2]:>10,.0f}")
    line()
    print(f"  {'（参考）第9期 基準額':38s} {7072:>10,} {6500:>10,} {'―':>10s}")
    print("   ※第9期は基金131,700千円の49.5％（65,223千円）を取り崩して6,500円としている。")
    print()

    # 中位ケースの内訳
    r, row = results["中位"]
    print("■ 4. 中位ケースの内訳（千円）")
    line()
    print(f"  標準給付費見込額　R9 {r['kyufu'][0]:>10,.0f}／R10 {r['kyufu'][1]:>10,.0f}"
          f"／R11 {r['kyufu'][2]:>10,.0f}")
    print(f"  　　　　　　3年計　{r['hyojun']:>12,.0f}"
          f"（第9期計画 3,294,553千円・{r['hyojun'] / 3_294_553 * 100 - 100:+.1f}％）")
    print(f"  地域支援事業費 3年計　{r['chiiki']:>12,.0f}"
          f"（第9期計画 91,560千円・{r['chiiki'] / 91_560 * 100 - 100:+.1f}％）")
    print(f"  第1号被保険者負担相当額（23％）　{r['futan']:>12,.0f}")
    print(f"  ＋ 調整交付金相当額（5％）　　　 {r['soto']:>12,.0f}")
    print(f"  − 調整交付金見込額（3.7％）　　 {r['mikomi']:>12,.0f}")
    print(f"  − 保険者機能強化推進交付金等　　{KYOKA_KOFU:>12,.0f}")
    print(f"  ＝ 取崩なしの保険料収納必要額　 {r['hitsuyo']:>12,.0f}"
          f"（第9期 804,663千円・{r['hitsuyo'] / 804_663 * 100 - 100:+.1f}％）")
    print()

    # 感度
    print("■ 5. 感度（中位ケース・50％取崩を基準）")
    line()
    base = results["中位"][1][1]
    r0 = results["中位"][0]
    def sens(label, hitsuyo=None, ninzu=None, shuno=0.96):
        h = r0["hitsuyo"] - KIKIN_R8 * 0.5 if hitsuyo is None else hitsuyo
        n = ninzu if ninzu else hosei
        _, m = kijungaku(h, n, shuno)
        print(f"  {label:44s} {m:>8,.0f}円　（{m - base:+,.0f}円）")
    print(f"  {'基準（中位・50％取崩・収納率96％）':44s} {base:>8,.0f}円")
    sens("予定収納率を96.7％（直近実績）とした場合", shuno=0.967)
    sens("補正後被保険者数が3％少ない場合", ninzu=hosei * 0.97)
    sens("調整交付金見込率が5.0％だった場合",
         hitsuyo=(r0["futan"] + r0["soto"] - r0["hyojun"] * 0.05 - KYOKA_KOFU) - KIKIN_R8 * 0.5)
    hitsuyo_b = r0["hitsuyo"] - KIKIN_R8 * 0.5
    # 入所・居住系1人増（年3,539千円×3年）
    zoka = 3_539 * 3 * 0.23 * (1 + 0.05 - 0.037)
    sens("入所・居住系の利用者が1人増えた場合", hitsuyo=hitsuyo_b + zoka)
    sens("同 10人増えた場合", hitsuyo=hitsuyo_b + zoka * 10)
    sens("同 125人増えた場合（施設申込33.8％の外挿）", hitsuyo=hitsuyo_b + zoka * 125)
    print()
    print("※ 本試算は令和8年9月2日時点で入手している資料のみによる仮試算である。")
    print("　 国保連データによるサービス見込量、所得段階別被保険者数、調整交付金の交付実績、")
    print("　 令和9年度介護報酬改定の内容が確定した段階で、第3回策定委員会に確定版を示す。")


if __name__ == "__main__":
    main()
