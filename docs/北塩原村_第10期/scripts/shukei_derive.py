# -*- coding: utf-8 -*-
"""派生変数の算出（集計仕様書 doc25 の DERIVED に対応）

リスク判定の基準は国の手引きに明記がなく、村・県・全国と比較できる形にする
必要がある。ここでは第9期の一般的な運用に合わせた案を実装し、確定前の
ものには JUDGE_PROVISIONAL に印を付けている（村への確認事項 doc25 参照）。
"""

JUDGE_PROVISIONAL = {"R01", "R02", "R03", "R05", "R06", "R07", "R08", "R09", "R10"}


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

    # R01 運動器の機能低下：問2(1)(2)(3)のうち「3.できない」が2項目以上
    v = [_i(r, f"N_問2_{k}") for k in (1, 2, 3)]
    d["R01"] = int(sum(1 for x in v if x == 3) >= 2) if any(x is not None for x in v) else None

    # R02 転倒リスク：問2(4)で「1.何度もある」または「2.1度ある」
    x = _i(r, "N_問2_4")
    d["R02"] = int(x in (1, 2)) if x is not None else None

    # R03 閉じこもり傾向：問2(6)で「1.ほとんど外出しない」または「2.週1回」
    x = _i(r, "N_問2_6")
    d["R03"] = int(x in (1, 2)) if x is not None else None

    # B01 BMI／R04 低栄養の傾向：BMI 18.5以下（手引きに明記）または問3(7)「1.はい」
    h, w = _i(r, "N_問3_1_身長"), _i(r, "N_問3_1_体重")
    bmi = round(w / (h / 100) ** 2, 1) if h and w and h > 0 else None
    d["B01"] = bmi
    x = _i(r, "N_問3_7")
    if bmi is None and x is None:
        d["R04"] = None
    else:
        d["R04"] = int((bmi is not None and bmi <= 18.5) or x == 1)

    # R05 口腔機能の低下：問3(2)(3)(4)のうち2項目以上が「1.はい」
    v = [_i(r, f"N_問3_{k}") for k in (2, 3, 4)]
    d["R05"] = int(sum(1 for y in v if y == 1) >= 2) if any(y is not None for y in v) else None

    # R06 認知機能の低下：問4(1)「1.はい」／(2)「2.いいえ」／(3)「1.はい」のいずれか
    a, b, c = _i(r, "N_問4_1"), _i(r, "N_問4_2"), _i(r, "N_問4_3")
    d["R06"] = (int(a == 1 or b == 2 or c == 1)
                if any(x is not None for x in (a, b, c)) else None)

    # R07 IADLの低下：問4(4)〜(8)のうち「3.できない」が1項目以上
    v = [_i(r, f"N_問4_{k}") for k in (4, 5, 6, 7, 8)]
    d["R07"] = int(any(y == 3 for y in v)) if any(y is not None for y in v) else None

    # R08 知的能動性の低下：問4(10)(11)(12)がすべて「2.いいえ」
    v = [_i(r, f"N_問4_{k}") for k in (10, 11, 12)]
    d["R08"] = int(all(y == 2 for y in v)) if all(y is not None for y in v) else None

    # R09 社会的役割の低下：問4(13)〜(16)がすべて「2.いいえ」
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
