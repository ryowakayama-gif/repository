# -*- coding: utf-8 -*-
"""認定者数・第1号被保険者数を年報と突き合わせる検査（令和8年10月16日）

**他メンバー作成の第10期計画素案（令和8年10月16日受領）と当方の正本を
突き合わせたところ、当方の素案 2-3 の実績値が年報と食い違っていた。**

  当方の素案 2-3 の第1号被保険者数　令和6年度3,270人／令和7年度3,285人
  年報の値　　　　　　　　　　　　　令和6年度末3,261人／令和7年度末3,240人
  ⚠ **3,270は令和5年度末、3,285は令和4年度末の値であり、年度が2年ずれていた。**

  当方の素案 2-3 の要介護度別認定者数（令和7年度）
    要支援1 33／要支援2 85／要介護1 70／要介護2 116／要介護3 131／
    要介護4 84／要介護5 49　計568人
  年報（令和7年度）様式1の5 (12) の総数
    要支援1 36／要支援2 90／要介護1 61／要介護2 103／要介護3 130／
    要介護4 96／要介護5 50　計566人
  ⚠ **当方の値は年報とも月報とも別系統（総括表）のものであった。**
  ⚠ 他メンバー作成の素案の値は年報と完全に一致していた。

  ⚠ **保険料と見込量には影響しない。** 保険料は calc_hokenryo の HIHOKEN
  （年報と同じ値）を用いており、見込量は令和8年度以降の見える化システムの
  推計値の比で伸ばしているためである。**記載の誤りである。**

本検査は、素案に載せる実績値が年報と一致するかを機械で確かめる。
**同じ誤りを二度と出さないための検査である。**

  python3 07_ソーススクリプト/check_nintei_R8.10.16.py [素案のパス]

検査
  1　第1号被保険者数（令和6・7年度）が年報の当年度末現在と一致するか
  2　要介護度別認定者数（令和6・7年度）が年報 様式1の5 の総数と一致するか
  3　第1号被保険者の認定者数（令和6・7年度）が年報と一致するか
  4　認定率が 認定者数÷第1号被保険者数 で再現するか
  5　年齢階級別の第1号被保険者数（2-1 の表）が年報と一致するか
"""
import os
import re
import sys

import docx
import openpyxl
from docx.oxml.ns import qn

SOAN = "01_第10期_最新版成果品/川崎町_計画書素案_v2.31_協議用.docx"
# 年報（受領した原本。リポジトリの中にある）
NENPO = {
    "令和6年度": "09_元資料/R8実績データ/R8.9.8受領版/年報データ_2024_川崎町.xlsx",
    "令和7年度": "09_元資料/R7実績データ/年報データ_2025_川崎町.xlsx",
}
KAIGO = ["要支援1", "要支援2", "要介護1", "要介護2", "要介護3",
         "要介護4", "要介護5"]


def yomu_nenpo(path):
    """年報から 第1号被保険者数・要介護度別認定者数・第1号の認定者数を読む。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    out = {}
    # ── 様式1 (2) 第1号被保険者数
    #   ⚠ (1)第1号被保険者のいる世帯数 の「計」が先に出てくるため、
    #     **(2)第1号被保険者数 の見出しの後ろ**で「計」を探す。
    ws = wb["様式１"]
    hajime = None
    for r in range(1, ws.max_row + 1):
        if "(2)" in str(ws.cell(r, 2).value or "") and \
                "第１号被保険者数" in str(ws.cell(r, 2).value or ""):
            hajime = r
    if hajime is None:
        raise SystemExit("様式1 に (2)第1号被保険者数 が見つからない：" + path)
    nen = {}
    for r in range(hajime, min(hajime + 12, ws.max_row + 1)):
        na = str(ws.cell(r, 3).value or "").strip()
        if na == "計":
            out["hihokensha_zen"] = ws.cell(r, 4).value   # 前年度末現在
            out["hihokensha"] = ws.cell(r, 7).value       # 当年度末現在
        if na in ("65歳以上75歳未満", "75歳以上85歳未満", "85歳以上"):
            nen[na] = ws.cell(r, 7).value
        if na == "(再掲)住所地特例被保険者":
            out["jusho_tokurei"] = ws.cell(r, 7).value
    out["nenrei"] = nen
    # 世帯数（(1)の「計」の当年度末現在）
    for r in range(1, hajime):
        if str(ws.cell(r, 3).value or "").strip() == "計":
            out["setai"] = ws.cell(r, 7).value
    # ── 様式1の5 (12) 認定者数
    #   ⚠ 「男」「女」「計」の3つの区分があり、**最後の「計」の区分**を見る。
    #     要支援1・2 は5・6列、要介護1〜5 は9〜13列（8列は経過的要介護）。
    ws = wb["様式１の５ 総数"]
    RETSU = (5, 6, 9, 10, 11, 12, 13)
    kei = None
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(r, 3).value or "").strip() == "計":
            kei = r
    if kei is None:
        raise SystemExit("様式1の5 に「計」の区分が見つからない：" + path)
    for r in range(kei, min(kei + 14, ws.max_row + 1)):
        na = str(ws.cell(r, 4).value or "").strip()
        if na == "第１号被保険者":
            out["nintei_1go"] = [ws.cell(r, c).value for c in RETSU]
        if na == "総　　数":
            out["nintei"] = [ws.cell(r, c).value for c in RETSU]
    for k in ("nintei", "nintei_1go"):
        if k not in out:
            raise SystemExit(f"様式1の5 から {k} を読めない：" + path)
    return out


def yomu_soan(path):
    """素案から 2-3 の2つの表と 2-1 の年齢階級別の表を読む。"""
    d = docx.Document(path)
    cur, out = "", {"nintei": {}, "ritsu": [], "nenrei": {}}
    for el in d.element.body.iterchildren():
        if el.tag == qn("w:p"):
            t = "".join(n.text or "" for n in el.iter(qn("w:t")))
            m = re.match(r"^(\d+-\d+)[　 ]", t.strip())
            if m:
                cur = m.group(1)
        elif el.tag == qn("w:tbl"):
            tb = docx.table.Table(el, d)
            hd = [c.text.replace("\n", "").strip() for c in tb.rows[0].cells]
            # 2-3 の要介護度別の表
            if cur == "2-3" and hd[:2] == ["年度", "要支援1"]:
                for r in tb.rows[1:]:
                    cs = [c.text.replace("\n", "").strip() for c in r.cells]
                    y = cs[0].split("（")[0]
                    vs = []
                    for x in cs[1:8]:
                        s = x.replace(",", "").replace("人", "")
                        vs.append(float(s) if re.fullmatch(
                            r"-?\d+(?:\.\d+)?", s) else None)
                    out["nintei"][y] = vs
            # 2-3 の認定率の表
            if cur == "2-3" and "第1号被保険者数" in " | ".join(hd):
                for r in tb.rows[1:]:
                    cs = [c.text.replace("\n", "").strip() for c in r.cells]
                    def n(x):
                        s = x.replace(",", "").replace("人", "").replace("%", "")
                        return float(s) if re.fullmatch(
                            r"-?\d+(?:\.\d+)?", s) else None
                    out["ritsu"].append((cs[0], n(cs[1]), n(cs[2]), n(cs[3])))
            # 2-1 の年齢階級別の表
            if cur == "2-1" and "令和7年度末" in " | ".join(hd):
                for r in tb.rows[1:]:
                    cs = [c.text.replace("\n", "").strip() for c in r.cells]
                    na = cs[0].split("（")[0].strip()
                    s = cs[2].replace(",", "").replace("人", "")
                    if re.fullmatch(r"-?\d+", s):
                        out["nenrei"][na] = float(s)
    return out


# 2-1 の表の区分と年報の年齢区分の読み替え
YOMIKAE_NENREI = {
    "65〜75歳未満": "65歳以上75歳未満",
    "75〜85歳未満": "75歳以上85歳未満",
    "85歳以上": "85歳以上",
}


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else SOAN
    if not os.path.exists(path):
        raise SystemExit("素案がない：" + path)
    for p in NENPO.values():
        if not os.path.exists(p):
            raise SystemExit("年報がない：" + p)
    nen = {k: yomu_nenpo(v) for k, v in NENPO.items()}
    sv = yomu_soan(path)
    ng, ok = [], []

    # ── 検査1・2・3
    for y in ("令和6年度", "令和7年度"):
        n = nen[y]
        # 検査2　要介護度別
        s = sv["nintei"].get(y)
        if s is None:
            ng.append(f"検査2　素案に{y}の要介護度別認定者数の行がない")
        else:
            for i, k in enumerate(KAIGO):
                a, b = n["nintei"][i], s[i]
                if a is None or b is None:
                    continue
                if abs(float(a) - b) > 0.5:
                    ng.append(f"検査2　{y}の{k}　年報{a:g}人と"
                              f"素案{b:g}人が食い違う")
            if all(x is not None for x in s):
                if abs(sum(float(x) for x in n["nintei"]) - sum(s)) > 0.5:
                    ng.append(f"検査2　{y}の認定者数の計　"
                              f"年報{sum(float(x) for x in n['nintei']):g}人と"
                              f"素案{sum(s):g}人が食い違う")
                else:
                    ok.append(f"{y}の要介護度別認定者数が年報と一致")
        # 検査1・3　第1号被保険者数と認定者数
        row = [r for r in sv["ritsu"] if r[0].startswith(y)]
        if not row:
            ng.append(f"検査1　素案に{y}の第1号被保険者数の行がない")
            continue
        _y, hi, nin, ritsu = row[0]
        if hi is not None and abs(float(n["hihokensha"]) - hi) > 0.5:
            ng.append(f"検査1　{y}の第1号被保険者数　"
                      f"年報{n['hihokensha']:,}人と素案{hi:,.0f}人が食い違う"
                      f"（年報の当年度末現在）")
        else:
            ok.append(f"{y}の第1号被保険者数が年報と一致")
        n1 = sum(float(x) for x in n["nintei_1go"])
        if nin is not None and abs(n1 - nin) > 0.5:
            ng.append(f"検査3　{y}の第1号被保険者の認定者数　"
                      f"年報{n1:g}人と素案{nin:g}人が食い違う")
        # 検査4　認定率
        if hi and nin and ritsu is not None:
            keisan = round(nin / hi * 100, 1)
            if abs(keisan - ritsu) > 0.15:
                ng.append(f"検査4　{y}の認定率　"
                          f"{nin:g}÷{hi:,.0f}＝{keisan}％だが"
                          f"素案は{ritsu}％")

    # ── 検査5　年齢階級別（2-1 の表・令和7年度末）
    n7 = nen["令和7年度"]["nenrei"]
    for k, nk in YOMIKAE_NENREI.items():
        a = sv["nenrei"].get(k)
        b = n7.get(nk)
        if a is None or b is None:
            continue
        if abs(float(b) - a) > 0.5:
            ng.append(f"検査5　2-1 の{k}（令和7年度末）　"
                      f"年報{b:g}人と素案{a:g}人が食い違う")
        else:
            ok.append(f"2-1 の{k}が年報と一致")

    print("認定者数・第1号被保険者数を年報と突き合わせる検査")
    print("  素案：", path.split("/")[-1])
    for y, p in NENPO.items():
        print(f"  年報（{y}）：{p.split('/')[-1]}")
    for y in ("令和6年度", "令和7年度"):
        n = nen[y]
        print(f"  ── 年報（{y}）")
        print(f"   第1号被保険者数　当年度末現在 {n['hihokensha']:,}人"
              f"（前年度末 {n['hihokensha_zen']:,}人）")
        print("   要介護度別認定者数（総数）　"
              + "／".join(f"{k} {v:g}" for k, v in zip(KAIGO, n["nintei"]))
              + f"　計{sum(float(x) for x in n['nintei']):g}人")
        print(f"   うち第1号被保険者 "
              f"{sum(float(x) for x in n['nintei_1go']):g}人")
    print("  ── 自己点検")
    if ng:
        for m in ng:
            print("   ×", m)
        print(f"   不適合 {len(ng)}件")
        print("  ⚠ 素案に載せる実績値は年報の値としてください。"
              "年報は受領した原本であり、リポジトリの中にあります。")
        sys.exit(1)
    for m in ok[:12]:
        print("   ○", m)
    print("   ○ 認定率は 認定者数÷第1号被保険者数 で再現する")


if __name__ == "__main__":
    main()
