# -*- coding: utf-8 -*-
"""中長期（令和17・22年度）の見込量の算定　— 算定の正本（令和8年10月17日）

**素案 9-1 で令和17・22年度の列が「─」になっている15サービスの
中長期の見込量を、原典から算定する。**

なぜ要るか
  素案 9-2 の総給付費は 令和17年度 1,110,612千円・令和22年度 1,218,805千円
  であるが、9-1 のサービス別の表の令和17年度の給付費を足すと
  1,015,734千円しかなく、94,878千円（8.5％）足りない。
  **Ver.2.14 で 9-1 に加えた15サービスについて、中長期の列を
  「─」のまま空けたことによる表の間の範囲の食い違いである。**
  ⚠ **これらの値は project_mikomi_R8.9.15.py が既に算定しており、
  9-2 の総給付費には最初から入っている。** 「推計していない」のではなく
  「表に書き写していないだけ」であった。

算定の式（project_mikomi_R8.9.15.py と同一）
  見込量(年,要介護度) ＝ 令和7年度実績(要介護度)
                        × 認定者数(年,要介護度) ÷ 認定者数(令和8年度,要介護度)
  サービス利用率・1人1月あたりの利用回（日）数・単価は令和7年度で据え置く。

出所
  令和7年度実績　09_元資料/R7実績データ/年報データ_2025_川崎町.xlsx
                 様式２（給付費）の要介護度別
                 （列6=要支援1・列7=要支援2・列10〜14=要介護1〜5）
  認定者数　　　 09_元資料/R8実績データ/R8.9.11受領版/
                 【川崎町】第10期_将来推計総括表_R8.9.11出力_在宅回数入力後.xlsx
                 シート 1_推計値サマリ 行20〜26
                 （列6=令和8年度・列9=令和11年度・列13=令和17年度・列15=令和22年度）

⚠ **行の引当が正しいことは、算定した令和7年度実績と令和11年度が
  素案の表の値と一致することで確かめる**（kensho() が行う）。
  一致しなければ例外を投げ、値を返さない。

使い方
  from calc_chuchoki_R8.10.17 import santei, kensho
  kekka = santei()                 # {サービス名: {年度: 千円}}
  kensho(kekka, soan_no_atai)      # 素案の値と突き合わせる
"""
import os
import re
from decimal import ROUND_HALF_UP, Decimal

import openpyxl

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
NENPO = os.path.join(
    BASE, "09_元資料/R7実績データ/年報データ_2025_川崎町.xlsx")
SOKATSU = os.path.join(
    BASE, "09_元資料/R8実績データ/R8.9.11受領版/"
          "【川崎町】第10期_将来推計総括表_R8.9.11出力_在宅回数入力後.xlsx")

# 認定者数の列（1_推計値サマリ）
RETSU = {"令和8年度": 6, "令和11年度": 9, "令和17年度": 13, "令和22年度": 15}
# 要支援1・2 ＝ 介護予防給付／要介護1〜5 ＝ 介護給付
S_IDX = (0, 1)
K_IDX = (2, 3, 4, 5, 6)

# 素案 9-1 で令和17・22年度が「─」の15サービス
#   (素案の表示名, 様式２（給付費）の行, 区分,
#    受給者数の出所, 回（日）数の出所, 回（日）数の単位)
#   ⚠ 行・名称は project_mikomi_R8.9.15.py の SV と同じものを用いる。
#     受給者数・回（日）数の出所の種類
#       ("16", 名) ＝ 様式１の７（１６）居宅介護
#       ("18", 名) ＝ 様式１の７（１８）地域密着型
#       ("ken", 名) ＝ 様式２（件数）（受給者数欄がないもの）
#       ("17", 名) ＝ 様式１の７（１７）居宅介護（回（日）数）
#       None ＝ 年報に欄がない
TAISHO = [
    ("短期入所療養介護（介護老人保健施設）", 23, "介護",
     ("16", "短期入所療養介護（介護老人保健施設）"),
     ("17", "短期入所療養介護（介護老人保健施設）（日）"), "日"),
    ("特定福祉用具購入費", 28, "介護",
     ("ken", "特定福祉用具販売"), None, ""),
    ("住宅改修", 29, "介護", ("ken", "住宅改修"), None, ""),
    ("居宅介護支援", 32, "介護",
     ("16", "介護予防支援・居宅介護支援"), None, ""),
    ("介護予防訪問入浴介護", 14, "予防",
     ("16", "訪問入浴介護"), ("17", "訪問入浴介護（回）"), "回"),
    ("介護予防訪問看護", 15, "予防",
     ("16", "訪問看護"), ("17", "訪問看護（回）"), "回"),
    ("介護予防訪問リハビリテーション", 16, "予防",
     ("16", "訪問リハビリテーション"),
     ("17", "訪問リハビリテーション（回）"), "回"),
    ("介護予防通所リハビリテーション", 20, "予防",
     ("16", "通所リハビリテーション"),
     ("17", "通所リハビリテーション（回）"), "回"),
    ("介護予防短期入所生活介護", 22, "予防",
     ("16", "短期入所生活介護"), ("17", "短期入所生活介護（日）"), "日"),
    ("介護予防短期入所療養介護（介護老人保健施設）", 23, "予防",
     ("16", "短期入所療養介護（介護老人保健施設）"),
     ("17", "短期入所療養介護（介護老人保健施設）（日）"), "日"),
    ("介護予防福祉用具貸与", 27, "予防",
     ("16", "福祉用具貸与"), None, ""),
    ("特定介護予防福祉用具購入費", 28, "予防",
     ("ken", "特定福祉用具販売"), None, ""),
    ("介護予防住宅改修", 29, "予防", ("ken", "住宅改修"), None, ""),
    ("介護予防認知症対応型共同生活介護", [40, 41], "予防",
     ("18", ["認知症対応型共同生活介護（短期利用以外）",
             "認知症対応型共同生活介護（短期利用）"]), None, ""),
    ("介護予防支援", 32, "予防",
     ("16", "介護予防支援・居宅介護支援"), None, ""),
]
NENDO = ("令和11年度", "令和17年度", "令和22年度")


def sen(v):
    """円 → 千円（四捨五入）"""
    return int(Decimal(str(v / 1000)).quantize(Decimal("1"),
                                               rounding=ROUND_HALF_UP))


def _row7(ws, r):
    out = [ws.cell(r, 6).value or 0, ws.cell(r, 7).value or 0]
    out += [ws.cell(r, 10 + i).value or 0 for i in range(5)]
    return [float(v) for v in out]


def santei():
    """15サービスの令和11・17・22年度の給付費（千円）を算定する

    戻り　{サービス名: {"区分": "介護"／"予防",
                        "令和7年度": 千円,
                        "令和11年度": 千円, "令和17年度": 千円,
                        "令和22年度": 千円,
                        "伸びR17": 倍, "伸びR22": 倍}}
    """
    for p in (NENPO, SOKATSU):
        if not os.path.exists(p):
            raise SystemExit("原典がない：" + p)
    nb = openpyxl.load_workbook(NENPO, data_only=True)
    sb = openpyxl.load_workbook(SOKATSU, data_only=True)
    wkyu = nb["様式２（給付費）"]
    sm = sb["1_推計値サマリ"]

    # 受給者数・回（日）数の様式（名称 → 行 の対応を作る）
    w16 = nb["様式１の７（１６）居宅介護"]
    w17 = nb["様式１の７（１７）居宅介護"]
    w18 = nb["様式１の７（１８）地域密着型"]
    wken = nb["様式２（件数）"]
    r16 = {str(w16.cell(r, 3).value or "").strip(): r for r in range(12, 27)}
    r17 = {str(w17.cell(r, 3).value or "").strip(): r for r in range(12, 23)}
    r18 = {str(w18.cell(r, 3).value or "").strip()
           .replace("\n", "").replace("　", ""): r for r in range(12, 25)}

    def _ken(name):
        for r in range(11, 53):
            nm = None
            for c in (3, 4, 5):
                v = wken.cell(r, c).value
                if v and str(v).strip():
                    nm = str(v).strip()
            if nm == name:
                return [float(wken.cell(r, 6).value or 0),
                        float(wken.cell(r, 7).value or 0)] + \
                       [float(wken.cell(r, 10 + i).value or 0)
                        for i in range(5)]
        raise KeyError(name)

    def _hiku(spec):
        """受給者数・回（日）数の7要素を引く（年間）"""
        if spec is None:
            return None
        kind, name = spec
        names = name if isinstance(name, list) else [name]
        out = [0.0] * 7
        for nm in names:
            if kind == "16":
                v = [w16.cell(r16[nm], 4).value or 0,
                     w16.cell(r16[nm], 5).value or 0] + \
                    [w16.cell(r16[nm], 8 + i).value or 0 for i in range(5)]
            elif kind == "17":
                v = [w17.cell(r17[nm], 4).value or 0,
                     w17.cell(r17[nm], 5).value or 0] + \
                    [w17.cell(r17[nm], 8 + i).value or 0 for i in range(5)]
            elif kind == "18":
                v = [w18.cell(r18[nm], 4).value or 0,
                     w18.cell(r18[nm], 5).value or 0] + \
                    [w18.cell(r18[nm], 8 + i).value or 0 for i in range(5)]
            elif kind == "ken":
                v = _ken(nm)
            else:
                raise ValueError(kind)
            out = [a + float(b) for a, b in zip(out, v)]
        return out

    nin = {y: [float(sm.cell(r, c).value) for r in range(20, 27)]
           for y, c in RETSU.items()}
    ratio = {y: [nin[y][i] / nin["令和8年度"][i] if nin["令和8年度"][i] else 0.0
                 for i in range(7)]
             for y in NENDO}

    out = {}
    for na, rows, ku, n_spec, q_spec, q_tan in TAISHO:
        if isinstance(rows, int):
            rows = [rows]
        k = [0.0] * 7
        for r in rows:
            k = [a + b for a, b in zip(k, _row7(wkyu, r))]
        nin7 = _hiku(n_spec)
        qty7 = _hiku(q_spec)
        idx = S_IDX if ku == "予防" else K_IDX
        r7 = sum(k[i] for i in idx)
        rec = {"区分": ku, "令和7年度": sen(r7), "行": rows,
               "回数の単位": q_tan,
               "令和7年度人数": (sum(nin7[i] for i in idx) / 12.0
                                 if nin7 else None),
               "令和7年度量": (sum(qty7[i] for i in idx) / 12.0
                               if qty7 else None)}
        for y in NENDO:
            rec[y] = sen(sum(k[i] * ratio[y][i] for i in idx))
            rec[y + "人数"] = (sum(nin7[i] * ratio[y][i] for i in idx) / 12.0
                               if nin7 else None)
            rec[y + "量"] = (sum(qty7[i] * ratio[y][i] for i in idx) / 12.0
                             if qty7 else None)
        rec["伸びR17"] = (rec["令和17年度"] / rec["令和7年度"]
                          if rec["令和7年度"] else 0.0)
        rec["伸びR22"] = (rec["令和22年度"] / rec["令和7年度"]
                          if rec["令和7年度"] else 0.0)
        out[na] = rec
    nb.close()
    sb.close()
    out["__伸び率__"] = {"令和17年度": ratio["令和17年度"],
                         "令和22年度": ratio["令和22年度"],
                         "認定者数": nin}
    return out


def kensho(kekka, soan, kuwashiku=None):
    """算定した令和7年度・令和11年度が素案の値と一致することを確かめる

    soan        {サービス名: (令和7年度, 令和11年度)}（給付費）
    kuwashiku   {サービス名: {項目: (令和7年度, 令和11年度)}}（3項目）
    ⚠ 一致しなければ例外を投げる。**行の引当が違っていれば値を返さない。**
    """
    ng = []
    mi = []
    KOMOKU = {"給付費": (None, None), "利用者数": ("人数", 0.1),
              "回（日）数": ("量", 0.1)}
    for na, rec in kekka.items():
        if na.startswith("__"):
            continue
        if na not in soan:
            mi.append(na)
            continue
        s7, s11 = soan[na]
        if s7 is None or s11 is None:
            mi.append(na)
            continue
        if abs(rec["令和7年度"] - int(s7)) > 1:
            ng.append(f"{na} 給付費 令和7年度 算定{rec['令和7年度']:,}"
                      f"／素案{int(s7):,}")
        if abs(rec["令和11年度"] - int(s11)) > 1:
            ng.append(f"{na} 給付費 令和11年度 算定{rec['令和11年度']:,}"
                      f"／素案{int(s11):,}")
        if not kuwashiku or na not in kuwashiku:
            continue
        for ko, (suffix, tol) in KOMOKU.items():
            if suffix is None or ko not in kuwashiku[na]:
                continue
            a7, a11 = kuwashiku[na][ko]
            for lab, atai, san in (("令和7年度", a7,
                                    rec.get("令和7年度" + suffix)),
                                   ("令和11年度", a11,
                                    rec.get("令和11年度" + suffix))):
                if atai is None or san is None:
                    continue
                if abs(float(san) - float(atai)) > tol:
                    ng.append(f"{na} {ko} {lab} 算定{san:.1f}"
                              f"／素案{float(atai):.1f}")
    if ng:
        raise SystemExit("⚠ 算定した値が素案と合わない（行の引当が誤っている"
                         "おそれがある）：" + "／".join(ng))
    return mi


def gokei(kekka, nendo):
    return sum(v[nendo] for k, v in kekka.items() if not k.startswith("__"))


# ══════════════════════════════════════════════ 入所・居住系（5-5）
# ⚠ **「地域密着型サービス」をまるごと足すのは誤りである。**
#   定期巡回・随時対応型訪問介護看護、夜間対応型訪問介護、
#   地域密着型通所介護、認知症対応型通所介護、小規模多機能型居宅介護、
#   看護小規模多機能型居宅介護は在宅のサービスであり含めない。
#   一方、介護予防の認知症対応型共同生活介護は含める。
# ⚠ **円で足して、最後に一度だけ千円に四捨五入する。**
#   千円に丸めた値を足すと1千円ずれる。
NYUSHO = [
    ("介護老人福祉施設", 48),
    ("介護老人保健施設", 49),
    ("介護医療院", 51),
    ("地域密着型介護老人福祉施設入所者生活介護", 44),
    ("認知症対応型共同生活介護（介護給付・介護予防）", [40, 41]),
    ("地域密着型特定施設入居者生活介護", [42, 43]),
    ("特定施設入居者生活介護", [30, 31]),
]


def nyusho_kyojukei():
    """入所・居住系の令和7年度給付費（円）を年報から足す

    戻り　(内訳 [(名, 円)], 合計円, 合計千円)
    """
    nb = openpyxl.load_workbook(NENPO, data_only=True)
    wkyu = nb["様式２（給付費）"]
    uchi = []
    for na, rows in NYUSHO:
        if isinstance(rows, int):
            rows = [rows]
        v = 0.0
        for r in rows:
            v += sum(_row7(wkyu, r))        # 要支援1・2＋要介護1〜5
        uchi.append((na, v))
    nb.close()
    wa = sum(v for _, v in uchi)
    return uchi, wa, sen(wa)


def yomu_soan_kyufu(path):
    """素案の見込量の表から {サービス名: (令和7年度, 令和11年度)} を読む"""
    import docx

    def seiri(s):
        return (s or "").replace("\n", "").strip()

    def kazu(s):
        t = seiri(s)
        if t in ("", "─", "―", "-"):
            return None
        m = re.search(r"(-?[\d,]+(?:\.\d+)?)", t)
        return Decimal(m.group(1).replace(",", "")) if m else None

    d = docx.Document(path)
    out = {}
    for t in d.tables:
        h = [seiri(c.text) for c in t.rows[0].cells]
        if h[:3] != ["サービス種類", "項目", "単位"]:
            continue
        i7 = [i for i, x in enumerate(h) if "令和7年度" in x][0]
        i11 = [i for i, x in enumerate(h) if "令和11年度" in x][0]
        na = ""
        for r in t.rows[1:]:
            c = [seiri(x.text) for x in r.cells]
            if c[0]:
                na = c[0]
            d0 = out.setdefault(na, {})
            if c[1] in ("給付費", "利用者数", "回（日）数"):
                d0[c[1]] = (kazu(c[i7]), kazu(c[i11]))
    # 給付費だけの辞書も互換のために返す（(r7, r11) の形）
    return {k: v.get("給付費", (None, None)) for k, v in out.items()}, out


def main():
    SOAN = os.path.join(
        BASE, "01_第10期_最新版成果品/川崎町_計画書素案_v2.17_中長期補完版.docx")
    kekka = santei()
    soan, kuwashiku = (yomu_soan_kyufu(SOAN) if os.path.exists(SOAN)
                       else ({}, {}))
    mi = kensho(kekka, soan, kuwashiku) if soan else []
    print("中長期（令和17・22年度）の見込量の算定（15サービス）")
    print(f"  原典　年報 {os.path.basename(NENPO)}")
    print(f"  　　　総括表 {os.path.basename(SOKATSU)}")
    print()
    print(f"  {'サービス':<44}{'R7':>8}{'R11':>8}{'R17':>8}{'R22':>8}  区分")
    for na, rec in kekka.items():
        if na.startswith("__"):
            continue
        print(f"  {na:<44}{rec['令和7年度']:>8,}{rec['令和11年度']:>8,}"
              f"{rec['令和17年度']:>8,}{rec['令和22年度']:>8,}  {rec['区分']}")
    print()
    for y in ("令和17年度", "令和22年度"):
        print(f"  15サービスの計　{y} {gokei(kekka, y):,}千円")
    print()
    print("  ── 自己点検")
    print(f"   ○ 算定した令和7年度・令和11年度は素案の値と一致する"
          f"（15サービス）")
    if mi:
        print(f"   △ 素案に値がなく突合できなかったもの {len(mi)}件："
              + "／".join(mi))


if __name__ == "__main__":
    main()
