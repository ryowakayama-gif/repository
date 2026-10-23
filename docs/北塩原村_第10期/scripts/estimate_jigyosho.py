# -*- coding: utf-8 -*-
"""村内の介護サービス事業所の利用者数・従事者数（令和3〜7年）

【出どころ】
  村より受領した事業所別の従事者数・利用者数・回転率・定員充足率の5か年推移
  （介護サービス情報公表システムの項目による。全国47都道府県分）から、
  本村の5事業所の値だけを取り出した表（data/原典_事業所別5か年/、doc116）。
  他市町村の値は取り込まない。

【数を書き写さない】
  素案（2-5・5-4(5)・5-4(12)・資6）はこのモジュールの値を引く。

【欠けている年の扱い】
  通所介護（社協）の令和7年の従事者数と、通所介護の令和7年・令和6年の定員の一部は
  掲載がない。従事者数は前年の値で補い（hokan に記録する）、定員は他の年と同じ値で補う
  （定員は平成30年度以降変わっていない。2-5）。

  python3 scripts/estimate_jigyosho.py
"""
import csv
import os
import sys

sys.dont_write_bytecode = True
import paths as P

CSV = os.path.join(P.data("原典_事業所別5か年"), "北塩原村_事業所別5か年推移.csv")
Y = ("令和3年", "令和4年", "令和5年", "令和6年", "令和7年")
SV = ("訪問介護", "通所介護", "認知症対応型共同生活介護")


def _rows():
    with open(CSV, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def _v(x):
    return None if x in ("", None) else float(x)


def load():
    """{(サービス, 事業所番号): {"名": 事業所名, 項目: {年: 値}}}"""
    d = {}
    for r in _rows():
        k = (r["サービス"], r["事業所番号"])
        e = d.setdefault(k, {"名": r["事業所名"]})
        e[r["項目"]] = {y: _v(r[y]) for y in Y}
    return d


def _hokan(s, hokan, label):
    """欠けた年を前の年（なければ後の年）の値で補う"""
    out = dict(s)
    for i, y in enumerate(Y):
        if out[y] is None:
            mae = [out[z] for z in Y[:i] if out[z] is not None]
            ato = [s[z] for z in Y[i + 1:] if s[z] is not None]
            v = mae[-1] if mae else (ato[0] if ato else None)
            if v is not None:
                out[y] = v
                hokan.append(f"{label} {y}")
    return out


def summary():
    """サービス別の集計。

    返り値 {サービス: {"事業所数", "定員", "利用者": {年}, "延べ": {年},
                        "介護職員": {年}, "看護職員等": {年}}}, 補った箇所の一覧
    """
    d = load()
    hokan = []
    out = {s: {"事業所数": 0, "定員": 0,
               "利用者": dict.fromkeys(Y, 0.0), "延べ": dict.fromkeys(Y, 0.0),
               "介護職員": dict.fromkeys(Y, 0.0), "看護職員等": dict.fromkeys(Y, 0.0)}
           for s in SV}
    for (sv, no), e in d.items():
        o = out[sv]
        o["事業所数"] += 1
        tei = e.get("定員（人）")
        if tei:
            vals = [v for v in tei.values() if v is not None]
            o["定員"] += vals[-1] if vals else 0
        for key, it in (("利用者", "1月あたり利用者数（実数・人）"),
                        ("延べ", "1月あたり延べ利用者数（人）"),
                        ("介護職員", "従事者数_介護職員_常勤換算数（人）"),
                        ("看護職員等", "従事者数_看護職員等_常勤換算数（人）")):
            if it not in e:
                continue
            s = _hokan(e[it], hokan, f"{e['名'][:14]}の{key}")
            for y in Y:
                o[key][y] += s[y] or 0
    return out, hokan


def kaigo_kei():
    """5事業所の介護職員の常勤換算数の計 {年: 人}"""
    s, _h = summary()
    return {y: sum(s[sv]["介護職員"][y] for sv in SV) for y in Y}


def main():
    s, hokan = summary()
    for sv in SV:
        o = s[sv]
        print(f"■ {sv}（{o['事業所数']}事業所・定員{o['定員']:.0f}）")
        for key in ("利用者", "延べ", "介護職員", "看護職員等"):
            print("   %-6s" % key + "".join(f"{o[key][y]:8.1f}" for y in Y))
    k = kaigo_kei()
    print("■ 介護職員の常勤換算数の計", {y: round(v, 1) for y, v in k.items()})
    print("■ 補った箇所", hokan)


if __name__ == "__main__":
    main()
