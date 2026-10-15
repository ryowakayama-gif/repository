# -*- coding: utf-8 -*-
"""中長期（令和12・17・22年度）の、サービスの種類ごとの量と地域支援事業の量。

【なぜ必要か】
  社会福祉法等の一部を改正する法律（令和8年法律第51号。要綱 第2の14(1)イ）は、
  市町村介護保険事業計画に

    「介護給付等対象サービスの種類ごとの量及び地域支援事業の量に関する
      中長期的な推計」

  を定めることを**義務**として加えた。施行は令和9年4月1日で、第10期計画
  そのものに掛かる。

  素案5-11は中長期の見通しを置いているが、**総量**（在宅サービス受給者数・
  サービス諸費）であって種類ごとではなく、地域支援事業は**費用**であって
  量ではない。法が求めているのは種類ごとの量と地域支援事業の量である。

【算定】
  サービスの種類ごとの量
    estimate_services.py が既に令和12・17・22年度まで算定し、
    data/第10期_サービス種類別見込量.csv に書き出している。ここではそれを読む。
    じか書きしない。

  地域支援事業の量
    素案5-5(1)と同じ延ばし方による。
      介護予防・生活支援サービス事業に属するもの … 要支援・要介護認定者数の伸び
      一般介護予防事業・包括的支援事業・任意事業 … 第1号被保険者数の伸び
    基準は令和6年度の実績で、5-5(1)の令和9〜11年度と同じ値になる。
    **同じ値を2か所に置くことになるため、点検79が5-5(1)と突き合わせる。**
    ずれたら不適合になる。

【延ばす基礎】（estimate_certified.py の案B。素案5-2・5-3の採用値）
  基準は令和8年度＝100（第1号被保険者数1,012人・認定者数214人）。素案5-5と同じ。
  第1号被保険者数   令和11年 987 → 令和12年 971 → 令和17年 891 → 令和22年 814
  認定者数          令和11年 220 → 令和12年 220 → 令和17年 212 → 令和22年 193

【留意】
  報酬改定・制度改正・供給の制約を織り込まない参考値である。
  実績の提供を受けていない事業は「―」のままとし、受けた時点で同じ方法により算定する。

【使い方】
  python3 scripts/estimate_longterm_ryo.py
"""
import csv
import io
import os
import sys
sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths as P

# 掲げる年度（令和11年度＝計画の最終年度を起点に置く）
YEARS = ["令和11年度", "令和12年度", "令和17年度", "令和22年度"]

# 延ばす基礎（estimate_certified.py 案B）
INSURED = {"令和9年度": 1004, "令和10年度": 995, "令和11年度": 987,
           "令和12年度": 971, "令和17年度": 891, "令和22年度": 814}
NINTEI = {"令和9年度": 214, "令和10年度": 217, "令和11年度": 220,
          "令和12年度": 220, "令和17年度": 212, "令和22年度": 193}
# 基準は令和8年度＝100。素案5-5(1)の「量の見込みの立て方」と同じ置き方である
# （第1号被保険者数 5-2 の 1,012人／認定者数 5-3 の 214人）。
# 基準を令和5年度末に取ると5-5と1〜2人ずれる。点検79が突き合わせる。
KIJUN = {"insured": 1012, "nintei": 214}

# ══════════════════════════════════════════════════════════════
# 地域支援事業の量
#   (事業, 単位, 令和6年度の実績, 延ばす基礎, 小数点以下の桁)
#   実績の提供を受けていないものは jisseki を None とする（「―」で出す）
#   延ばす基礎 "ninsho"＝認定者数／"insured"＝第1号被保険者数
# ══════════════════════════════════════════════════════════════
CHIIKI = [
 ("訪問型サービス（従前相当）", "人／月", None, "ninsho", 1),
 ("訪問型サービスA〜D", "人／月", 0, "ninsho", 0),
 ("通所型サービス（従前相当）", "人／月", None, "ninsho", 1),
 ("通所型サービスA〜C", "人／月", 0, "ninsho", 0),
 ("介護予防ケアマネジメント", "件／月", None, "ninsho", 1),
 ("一般介護予防事業（通いの場）", "か所", 23, "insured", 0),
 ("　同　参加者数", "人", 281, "insured", 0),
 ("　同　参加率", "%", 27.8, "sueoki", 1),
 ("　うち週1回以上開催", "か所", 8, "insured", 0),
 ("　同　参加者数", "人", 123, "insured", 0),
 ("総合相談支援", "件／年", None, "insured", 0),
 ("地域ケア会議", "回／年", 6, "insured", 0),
 ("認知症初期集中支援推進事業（対象者数）", "人／年", None, "insured", 0),
 ("認知症地域支援・ケア向上事業（推進員数）", "人", 1, "insured", 0),
 ("認知症カフェ", "か所", None, "insured", 0),
 ("認知症サポーター養成講座", "回／年", None, "insured", 0),
 ("　同　養成者数（累計）", "人", 237, "ruikei", 0),
 ("チームオレンジ", "か所", 0, "ruikei", 0),
 ("家族介護者支援（リフレッシュ事業）", "人／年", None, "insured", 0),
]
# 箇所数・推進員数は村の判断により置く数であり、人口に比例させない
SUEOKI = {"一般介護予防事業（通いの場）", "　うち週1回以上開催",
          "地域ケア会議", "認知症地域支援・ケア向上事業（推進員数）"}


def _nobasu(jisseki, kiso, year, keta, nm):
    """令和6年度の実績を、延ばす基礎により year まで延ばす"""
    if jisseki is None:
        return "―"
    if kiso == "ruikei":
        # 累計・立ち上げの数は延ばす基礎を持たない（5-5(1)も「―」としている）
        return "―"
    if kiso == "sueoki" or nm in SUEOKI:
        v = jisseki
    elif kiso == "ninsho":
        v = jisseki * NINTEI[year] / KIJUN["nintei"]
    else:
        v = jisseki * INSURED[year] / KIJUN["insured"]
    return f"{v:,.{keta}f}"


def chiiki_rows(years=YEARS):
    """地域支援事業の量。(事業, 単位, 令和6年度, *years) の行を返す"""
    out = []
    for nm, tani, jis, kiso, keta in CHIIKI:
        r = [nm, tani, "―" if jis is None else f"{jis:,.{keta}f}"]
        r += [_nobasu(jis, kiso, y, keta, nm) for y in years]
        out.append(r)
    return out


def service_rows(years=YEARS):
    """サービスの種類ごとの量。算定の正本（CSV）から読む。じか書きしない"""
    f = P.data("第10期_サービス種類別見込量.csv")
    out = []
    with io.open(f, encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            row = [r["区分"], r["サービス種類"]]
            for y in years:
                row.append(r[f"{y}(人/月)"])
            out.append(row)
    return out


def main():
    print("■ 延ばす基礎\n")
    print(f"  {'基礎':<16}{'令和5年度末':>12}" + "".join(f"{y:>12}" for y in YEARS))
    print(f"  {'第1号被保険者数':<16}{KIJUN['insured']:>10}人"
          + "".join(f"{INSURED[y]:>10}人" for y in YEARS))
    print(f"  {'認定者数':<16}{KIJUN['nintei']:>10}人"
          + "".join(f"{NINTEI[y]:>10}人" for y in YEARS))

    sv = service_rows()
    print(f"\n■ サービスの種類ごとの量（人／月）　{len(sv)}種類\n")
    print(f"  {'サービス':<24}" + "".join(f"{y:>11}" for y in YEARS))
    for r in sv:
        print(f"  {r[1]:<24}" + "".join(f"{x:>11}" for x in r[2:]))

    ch = chiiki_rows()
    print(f"\n■ 地域支援事業の量　{len(ch)}事業\n")
    print(f"  {'事業':<26}{'単位':<7}{'令和6年度':>10}"
          + "".join(f"{y:>11}" for y in YEARS))
    for r in ch:
        print(f"  {r[0]:<26}{r[1]:<7}{r[2]:>10}" + "".join(f"{x:>11}" for x in r[3:]))
    print("\n  「―」は村からの実績の提供を受けていないもの。"
          "提供を受けた時点で同じ方法により算定する。")
    print("  箇所数・会議の回数・推進員数は村の判断により置く数であるため"
          "人口に比例させず据え置いている。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
