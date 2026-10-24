# -*- coding: utf-8 -*-
"""業務工程管理表の更新（資料v13・問答集v9：第1回の書式・修正履歴を除く）
（令和8年10月22日・同日5度目）

  python3 07_ソーススクリプト/update_wbs_R8.10.22e.py

⚠ 列8（状態）には書き込まない。
"""
import copy
import sys

import openpyxl
from openpyxl.styles import Alignment, Font

WBS = "川崎町_業務工程管理表.xlsx"
SHIRYO_OLD = "川崎町_第2回策定委員会資料_R8.11_v12.docx"
SHIRYO = "川崎町_第2回策定委員会資料_R8.11_v13.docx"
FONT = Font(name="游ゴシック", size=9)
HIZUKE = "【令和8年10月22日 更新⑤（資料v13・第1回の書式）】"


def main():
    wb = openpyxl.load_workbook(WBS)
    ws = wb["04_成果品管理"]
    n = 0
    for x in range(5, ws.max_row + 1):
        na = str(ws.cell(x, 3).value or "")
        if str(ws.cell(x, 1).value or "").strip() == "6" and SHIRYO_OLD in na:
            ws.cell(x, 3).value = (na.replace(SHIRYO_OLD, SHIRYO)
                                   .replace("同_想定問答集_R8.11_v8.docx",
                                            "同_想定問答集_R8.11_v9.docx"))
            ws.cell(x, 8).value = (
                f"{HIZUKE}ご指示により、資料を第1回の書式に揃え（表紙の帯と"
                f"町のロゴ・表の1行おきの網掛け・注記は※）、表の本文を左揃えに"
                f"した。計画素案は策定委員会に未提出のため、修正履歴（当社資料の"
                f"訂正11件・素案と資料の版）を除いた。⚠ v12 の表紙は帯の網掛けが"
                f"失われ、題名が白地に白で見えなくなっていた。問答集を v9 に"
                f"揃えた（QB-45 の置き換え・頁数の統一）。　"
                + str(ws.cell(x, 8).value or ""))
            for c in range(1, 10):
                ws.cell(x, c).font = FONT
                ws.cell(x, c).alignment = Alignment(wrap_text=True,
                                                    vertical="top")
            n += 1

    ws2 = wb["08_作業順位"]
    if not any(str(ws2.cell(x, 1).value or "").strip() == "40"
               for x in range(5, ws2.max_row + 1)):
        moto = ws2.max_row
        r = moto + 1
        for c in range(1, 11):
            ws2.cell(r, c)._style = copy.copy(ws2.cell(moto, c)._style)
        ws2.row_dimensions[r].height = ws2.row_dimensions[moto].height
        atai = [
            "40", "第2回策定委員会資料の書式の整理（第1回の踏襲）", "R8.10.22",
            "ご指示による。計画素案は策定委員会に未提出のため修正履歴は不要。"
            "表の中央揃えを左揃えに。書式は第1回の資料を踏襲する。",
            "成果品6", "第2回策定委員会資料 v13／想定問答集 v9", "半日", "完了",
            "⚠ **v12 の表紙は帯の網掛けが clear になっており、白抜きの題名が"
            "見えなかった。**第1回と同じ solid に戻し、町のロゴを入れ、2章から"
            "改頁。表は本文を左揃え・1行おきの網掛け。⚠ は第1回にならい ※ に"
            "（0件）。当社資料の訂正11件の節・版の記述を除いた。古い値の残り"
            "（検算6,100項目→6,400項目、圧縮後102頁→106頁）を直した。",
        ]
        for c, v in enumerate(atai, 1):
            ws2.cell(r, c).value = v
    wb.save(WBS)

    wb2 = openpyxl.load_workbook(WBS, data_only=True)
    s = "\n".join(str(v) for sh in wb2 for row in sh.iter_rows(values_only=True)
                  for v in row if v is not None)
    jun = [str(wb2["08_作業順位"].cell(x, 1).value or "").strip()
           for x in range(5, wb2["08_作業順位"].max_row + 1)]
    ok = (n in (0, 1) and jun.count("40") == 1 and SHIRYO in s
          and "想定問答集_R8.11_v9" in s
          and not any(na in s for na in ("蔵王グリーンホーム", "鶴寿会",
                                          "北塩原", "金ヶ崎", "川崎市")))
    print("業務工程管理表の更新（資料v13・問答集v9）")
    print(("   ○ " if ok else "   × ") + f"成果品6 {n}件・作業順位40・名称の混入なし")
    if not ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
