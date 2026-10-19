# -*- coding: utf-8 -*-
"""派生変数の算出（集計仕様書 doc25 の DERIVED に対応）

リスク判定の基準は国の手引きに明記がなく、村・県・全国と比較できる形にする
必要がある。ここでは第9期の一般的な運用に合わせた案を実装し、確定前の
ものには JUDGE_PROVISIONAL に印を付けている（村への確認事項 doc25 参照）。
"""

# 令和8年10月19日に「介護予防・日常生活圏域ニーズ調査 実施の手引き」（第10期用）を
# 受領し、R01〜R06・R10 の該当基準を原典で確定した（同手引き Q34〜Q37・参考資料2）。
# 原典に該当基準の記載がないのは R07・R08・R09 の3つで、これらは老研式活動能力指標
# （手段的自立・知的能動性・社会的役割）の区分に準じた当方の定義である。
# 当方の定義によるものは、見える化システムの全国・県の値と比べない。
JUDGE_PROVISIONAL = {"R07", "R08", "R09"}

# ⚠ 令和8年9月9日の校正版（doc34 §4）により、次の対応が必要になる。
#   ・R09（社会的役割の低下）は問4(13)〜(16)の4項目すべてを要する。
#     (16) は削除済み、(15) は削除の指示があるため、そのまま印刷されると算出できない。
#     算出できない場合、R99 は10領域から9領域に減り、第9期調査との比較も成立しない。
#     L01（生活機能4層）は R99 に依存するため定義の見直しを要する。
#   ・S01・S02 は「そのような人はいない」の選択肢番号を参照している。
#     ニーズ問7(5)(8)(10) の順序が入れ替わるため、番号を確認して修正する。
#   ・derive_zaitaku の疾病数 D01 は ZA_問7 → ZA_問8 に変わり、
#     新設の「15.なし」を 0 として扱えるようになる（改善）。
#   ・利用サービス数 U01 は ZA_問9_A〜L → ZA_問10_A〜I に変わる。


def _i(r, k):
    v = r.get(k, "")
    if v in ("", None):
        return None
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def _multi_on(r, prefix, nums):
    """複数回答で指定した選択肢のいずれかが選ばれているか"""
    return any(str(r.get(f"{prefix}_c{n}", "")) == "1" for n in nums)


def _multi_count(r, prefix, total, exclude=()):
    return sum(1 for n in range(1, total + 1)
               if n not in exclude and str(r.get(f"{prefix}_c{n}", "")) == "1")


def derive_needs(r):
    """ニーズ調査の派生変数を計算して r に書き戻す"""
    d = {}

    # R01 運動器の機能低下：問2の5設問のうち3問以上が該当（手引き Q34）
    #     (1)(2)(3)は「3.できない」、(4)は「1.何度もある」「2.1度ある」、
    #     (5)は「1.とても不安である」「2.やや不安である」が該当の選択肢である。
    #     3設問のうち2項目以上としていたのを、原典に合わせて改めた。
    gai = [(1, (3,)), (2, (3,)), (3, (3,)), (4, (1, 2)), (5, (1, 2))]
    v = [(_i(r, f"N_問2_{k}"), ok) for k, ok in gai]
    d["R01"] = (int(sum(1 for x, ok in v if x in ok) >= 3)
                if any(x is not None for x, _ in v) else None)

    # R02 転倒リスク：問2(4)で「1.何度もある」または「2.1度ある」
    x = _i(r, "N_問2_4")
    d["R02"] = int(x in (1, 2)) if x is not None else None

    # R03 閉じこもり傾向：問2(6)で「1.ほとんど外出しない」または「2.週1回」
    x = _i(r, "N_問2_6")
    d["R03"] = int(x in (1, 2)) if x is not None else None

    # B01 BMI／R04 低栄養の傾向：BMI 18.5以下（手引き Q36）
    #     手引きは2段構えである。BMI 18.5以下で「低栄養が疑われる高齢者」、
    #     これに問3(7)「6ヶ月間で2〜3kg以上の体重減少」が加わると「低栄養状態にある高齢者」。
    #     R04は前者（国の基準）とし、後者は R04S に分けて持つ。
    #     「BMI 18.5以下 または 体重減少」としていたのを、原典に合わせて改めた。
    h, w = _i(r, "N_問3_1_身長"), _i(r, "N_問3_1_体重")
    bmi = round(w / (h / 100) ** 2, 1) if h and w and h > 0 else None
    d["B01"] = bmi
    x = _i(r, "N_問3_7")
    d["R04"] = int(bmi <= 18.5) if bmi is not None else None
    d["R04S"] = (int(bmi <= 18.5 and x == 1)
                 if (bmi is not None and x is not None) else None)

    # R05 口腔機能の低下：問3(2)(3)(4)のうち2項目以上が「1.はい」
    v = [_i(r, f"N_問3_{k}") for k in (2, 3, 4)]
    d["R05"] = int(sum(1 for y in v if y == 1) >= 2) if any(y is not None for y in v) else None

    # R06 認知機能の低下：問4(1)「物忘れが多いと感じますか」で「1.はい」（手引き Q37）
    #     (2)(3)も認知機能を問う設問ではあるが、手引きが判定に用いるのは(1)だけである。
    #     (1)(2)(3)のいずれかとしていたのを、原典に合わせて改めた。
    a = _i(r, "N_問4_1")
    d["R06"] = int(a == 1) if a is not None else None

    # R07 IADLの低下：問4(4)〜(8)のうち「3.できない」が1項目以上
    #     【手引きに該当基準の記載がない。老研式活動能力指標の手段的自立に準じた当方の定義】
    v = [_i(r, f"N_問4_{k}") for k in (4, 5, 6, 7, 8)]
    d["R07"] = int(any(y == 3 for y in v)) if any(y is not None for y in v) else None

    # R08 知的能動性の低下：問4(10)(11)(12)がすべて「2.いいえ」
    #     【手引きに該当基準の記載がない。老研式活動能力指標に準じた当方の定義】
    v = [_i(r, f"N_問4_{k}") for k in (10, 11, 12)]
    d["R08"] = int(all(y == 2 for y in v)) if all(y is not None for y in v) else None

    # R09 社会的役割の低下：問4(13)〜(16)がすべて「2.いいえ」
    #     【手引きに該当基準の記載がない。老研式活動能力指標に準じた当方の定義】
    v = [_i(r, f"N_問4_{k}") for k in (13, 14, 15, 16)]
    d["R09"] = int(all(y == 2 for y in v)) if all(y is not None for y in v) else None

    # R10 うつ傾向：問8(3)(4)のいずれかが「1.はい」
    a, b = _i(r, "N_問8_3"), _i(r, "N_問8_4")
    d["R10"] = int(a == 1 or b == 1) if (a is not None or b is not None) else None

    # R99 リスク該当領域数（0〜10）
    rs = [d[f"R{k:02d}"] for k in range(1, 11)]
    d["R99"] = sum(x for x in rs if x) if any(x is not None for x in rs) else None
    d["R99_3"] = None if d["R99"] is None else ("0領域" if d["R99"] == 0
                                                else "1〜2領域" if d["R99"] <= 2 else "3領域以上")

    # D01 疾病数：問8(7)の選択数（「1.ない」は0）
    if str(r.get("N_問8_7_c1", "")) == "1":
        d["D01"] = 0
    else:
        n = _multi_count(r, "N_問8_7", 19, exclude=(1,))
        d["D01"] = n if n or any(str(r.get(f"N_問8_7_c{k}", "")) == "1" for k in range(1, 20)) else None
    d["D01_3"] = None if d["D01"] is None else ("0疾病" if d["D01"] == 0
                                                else "1〜2疾病" if d["D01"] <= 2 else "3疾病以上")

    # P01 社会参加あり：問5(1)①〜⑧のいずれかが「月1〜3回」以上（値1〜4）
    v = [_i(r, f"N_問5_1_{k}") for k in range(1, 9)]
    d["P01"] = int(any(y is not None and y <= 4 for y in v)) if any(y is not None for y in v) else None

    # S01 支え合いの受領／S02 提供：「8.そのような人はいない」以外を選んでいるか
    d["S01"] = int(_multi_on(r, "N_問7_1", range(1, 8)) or _multi_on(r, "N_問7_3", range(1, 8)))
    d["S02"] = int(_multi_on(r, "N_問7_2", range(1, 8)) or _multi_on(r, "N_問7_4", range(1, 8)))
    d["S99"] = int(not d["S01"] and not d["S02"])

    # L01 生活機能4層
    cert = r.get("認定状況", "認定なし")
    if cert and cert != "認定なし":
        d["L01"] = "第4層　要支援・要介護層"
    elif d["R99"] is None:
        d["L01"] = None
    elif d["R99"] == 0:
        d["L01"] = "第1層　元気・活躍層" if d["P01"] else "第2層　元気・未参加層"
    else:
        d["L01"] = "第3層　リスク層"

    r.update({k: ("" if v is None else v) for k, v in d.items()})
    return r


def derive_zaitaku(r):
    """在宅介護実態調査の派生変数"""
    d = {}
    # 利用しているサービスの数（問9のA〜Lで「利用していない」以外）
    n = 0
    have = False
    for c in "ABCDEFGHIJKL":
        x = _i(r, f"ZA_問9_{c}")
        if x is not None:
            have = True
            if x >= 2:
                n += 1
    d["U01"] = n if have else None
    # 疾病数（問7。「ない」に相当する選択肢がない設計のため単純な選択数）
    d["D01"] = _multi_count(r, "ZA_問7", 20)
    r.update({k: ("" if v is None else v) for k, v in d.items()})
    return r
