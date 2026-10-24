# -*- coding: utf-8 -*-
"""業務工程管理表の更新（事業所分析の協議用素案への反映：素案Ver.2.28・資料v14・問答集v11）
（令和8年10月23日・同日5度目）

  python3 07_ソーススクリプト/update_wbs_R8.10.23e.py

⚠ 列8（状態）は動かさない。
"""
import copy
import sys

import openpyxl
from openpyxl.styles import Alignment, Font

WBS = "川崎町_業務工程管理表.xlsx"
FONT = Font(name="游ゴシック", size=9)
HIZUKE = "【令和8年10月23日 更新⑤（事業所分析の素案への反映）】"
OKIKAE = [("川崎町_計画書素案_v2.27_協議用.docx", "川崎町_計画書素案_v2.28_協議用.docx"),
          ("川崎町_計画書案_公表用_v2.27.docx", "川崎町_計画書案_公表用_v2.28.docx"),
          ("川崎町_第2回策定委員会資料_R8.11_v13.docx", "川崎町_第2回策定委員会資料_R8.11_v14.docx"),
          ("同_想定問答集_R8.11_v10.docx", "同_想定問答集_R8.11_v11.docx")]


def iro(c):
    c.font = FONT
    c.alignment = Alignment(wrap_text=True, vertical="top")


def main():
    wb = openpyxl.load_workbook(WBS)
    ws = wb["04_成果品管理"]
    n = 0
    for x in range(5, ws.max_row + 1):
        no = str(ws.cell(x, 1).value or "").strip()
        na = str(ws.cell(x, 3).value or "")
        nb = na
        for a, b in OKIKAE:
            nb = nb.replace(a, b)
        if nb != na and HIZUKE not in str(ws.cell(x, 8).value or ""):
            ws.cell(x, 3).value = nb
            moji = (f"{HIZUKE}協議用素案を Ver.2.28（117頁）に上げ、公表用も組み直した（114頁）。"
                    f"2-7に見える化システムの事業所分析による5年間の推移（稼働率・短期入所・"
                    f"通所の回転率・介護職員）を加え、別表・7-1・7-2・7-6・9-1に反映した。"
                    f"地域密着型の空き床の扱いを［要協議］として置いた。　") if no == "4" else (
                    f"{HIZUKE}資料を v14、想定問答集を v11（全68問）に上げた。資料３に5年間の"
                    f"推移の表、問答集に QB-68（地域密着型の空き床）を加えた。頁数を117頁に。　")
            ws.cell(x, 8).value = moji + str(ws.cell(x, 8).value or "")
            iro(ws.cell(x, 8))
            n += 1
    ws3 = wb["03_確認事項一覧"]
    k = 0
    for y in range(5, ws3.max_row + 1):
        if (str(ws3.cell(y, 1).value or "").strip() == "159"
                and HIZUKE not in str(ws3.cell(y, 10).value or "")):
            ws3.cell(y, 10).value = (
                f"{HIZUKE}延べ利用日数による稼働率と短期入所の利用者数を協議用素案 Ver.2.28 の"
                f"2-7・別表・9-1に書きました。地域密着型介護老人福祉施設（2施設計）の稼働率は"
                f"令和7年で78.7％、増やした定員のうち約6人分が埋まっていません。空きが職員不足"
                f"による受入れの制限によるものかのご確認をお願いします［要確認］。見込量の扱いは"
                f"［要協議］として素案9-1に置いています。　"
                + str(ws3.cell(y, 10).value or ""))
            iro(ws3.cell(y, 10))
            k += 1
    ws2 = wb["08_作業順位"]
    if not any(str(ws2.cell(x, 1).value or "").strip() == "44"
               for x in range(5, ws2.max_row + 1)):
        moto = ws2.max_row
        r = moto + 1
        for c in range(1, 11):
            ws2.cell(r, c)._style = copy.copy(ws2.cell(moto, c)._style)
        ws2.row_dimensions[r].height = ws2.row_dimensions[moto].height
        for c, v in enumerate([
                "44", "見える化システムの事業所分析の協議用素案への反映", "R8.10.23",
                "ご指示による。出所を確定し、手引きの定義で計算し直した後に反映。",
                "No.159・No.50", "協議用素案 Ver.2.28／公表用 v2.28／資料 v14／問答集 v11",
                "半日", "完了",
                "2-7に5年間の推移の表。⚠ 事業所ごとの職員数は掲げず、施設・居住系の合計"
                "（指定の重複を除く）と定員100人あたりで示した。⚠ 地域密着型の稼働率78.7％"
                "（増床分が埋まっていない）を9-1に［要協議］。資料・問答集を同じ版で直した。"], 1):
            ws2.cell(r, c).value = v
    wb.save(WBS)
    print("業務工程管理表の更新（事業所分析の反映）")
    ok = n in (0, 2) and k in (0, 1)
    print(("   ○ " if ok else "   × ") + f"成果品 {n}件・No.159 {k}件・作業順位44")
    if not ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
