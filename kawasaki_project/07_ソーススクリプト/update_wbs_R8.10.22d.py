# -*- coding: utf-8 -*-
"""業務工程管理表の更新（軽費老人ホームの廃止の確認結果と、素案Ver.2.25・資料v12・問答集v8）
（令和8年10月22日・同日4度目）

  python3 07_ソーススクリプト/update_wbs_R8.10.22d.py

⚠ 列8（状態）には書き込まない。⚠ 施設名・法人名は書かない。
"""
import copy
import sys

import openpyxl
from openpyxl.styles import Alignment, Font

WBS = "川崎町_業務工程管理表.xlsx"
SOAN_OLD = "川崎町_計画書素案_v2.24_交付金整理版.docx"
SOAN = "川崎町_計画書素案_v2.25_軽費老人ホーム確認版.docx"
SHIRYO_OLD = "川崎町_第2回策定委員会資料_R8.11_v11.docx"
SHIRYO = "川崎町_第2回策定委員会資料_R8.11_v12.docx"
FONT = Font(name="游ゴシック", size=9)
HIZUKE = "【令和8年10月22日 更新④（軽費老人ホームの廃止・素案Ver.2.25）】"


def iro(ws, r, n):
    for c in range(1, n + 1):
        ws.cell(r, c).font = FONT
        ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")


TSUIKI_154 = (
    f"{HIZUKE}廃止日は令和6年3月31日でした（法人の公表。県の入所状況調の"
    f"掲載の推移と整合）。令和5年1〜9月の退所32人の退所先は他の施設28・"
    f"入院3・在宅1（県の入所状況調）。⚠ 以前の更新で「住まい70人分が"
    f"失われた」と書きましたが、休止前の入所者は令和5年4月1日時点で13人で"
    f"あり、なくなったのは定員70人分の住まいの選択肢です。素案 Ver.2.25 で"
    f"改めました。⚠ 次の3点をお願いします。①宮城県 長寿社会政策課に、"
    f"受理された廃止届（様式第24号）の「廃止の理由」「現に入所している者に"
    f"対する措置」の記載内容をご確認ください。②入所されていた方のうち本町の"
    f"住民の数（匿名の人数のみ。数えた定義を添えて）。③運営法人への照会の"
    f"要否（既定値＝照会しない）。詳細は 川崎町_軽費老人ホーム廃止の"
    f"確認結果_R8.10.22.md。　")


def main():
    wb = openpyxl.load_workbook(WBS)
    ws = wb["03_確認事項一覧"]
    n = 0
    for x in range(5, ws.max_row + 1):
        if (str(ws.cell(x, 1).value or "").strip() == "154"
                and HIZUKE not in str(ws.cell(x, 10).value or "")):
            ws.cell(x, 10).value = TSUIKI_154 + str(ws.cell(x, 10).value or "")
            iro(ws, x, 10)
            n += 1

    ws = wb["04_成果品管理"]
    n_seika = 0
    for x in range(5, ws.max_row + 1):
        no = str(ws.cell(x, 1).value or "").strip()
        na = str(ws.cell(x, 3).value or "")
        if no == "4" and SOAN_OLD in na:
            ws.cell(x, 3).value = na.replace(SOAN_OLD, SOAN)
            ws.cell(x, 8).value = (
                f"{HIZUKE}素案を Ver.2.25 に上げた（全118頁）。2-7・別表 事項八に"
                f"町内の軽費老人ホームの廃止日・入所者の推移・退所先を書き、"
                f"「住まい70人分が失われた」を定員と入所者の別が分かる形に改めた。　"
                + str(ws.cell(x, 8).value or ""))
            iro(ws, x, 9)
            n_seika += 1
        elif no == "6" and SHIRYO_OLD in na:
            ws.cell(x, 3).value = (na.replace(SHIRYO_OLD, SHIRYO)
                                   .replace("同_想定問答集_R8.11_v7.docx",
                                            "同_想定問答集_R8.11_v8.docx"))
            ws.cell(x, 8).value = (
                f"{HIZUKE}資料を v12（全12版）、想定問答集を v8（全67問）に上げた。"
                f"表6・QB-62 を素案 Ver.2.25 と同じ内容に改めた。　"
                + str(ws.cell(x, 8).value or ""))
            iro(ws, x, 9)
            n_seika += 1

    ws2 = wb["08_作業順位"]
    if not any(str(ws2.cell(x, 1).value or "").strip() == "39"
               for x in range(5, ws2.max_row + 1)):
        moto = ws2.max_row
        r = moto + 1
        for c in range(1, 11):
            ws2.cell(r, c)._style = copy.copy(ws2.cell(moto, c)._style)
        ws2.row_dimensions[r].height = ws2.row_dimensions[moto].height
        atai = [
            "39", "町内の軽費老人ホームの廃止の確認結果の反映", "R8.10.22",
            "ご指示による。廃止前後を公表資料で調べた結果の提供を受け、"
            "手元の県の入所状況調と突き合わせた。",
            "No.154 のうち軽費老人ホーム",
            "計画素案 Ver.2.25／資料 v12／想定問答集 v8／"
            "軽費老人ホーム廃止の確認結果（md）",
            "半日", "完了",
            "廃止日 令和6年3月31日。⚠ **提供資料で未確認とされた廃止前の"
            "入所者数と退所先の一部が、県の入所状況調で埋まった**"
            "（13人→7人→0人、退所32人のうち他の施設28）。⚠ 素案の"
            "「住まい70人分が失われた」を改めた。本町の住民の数・移られた先の"
            "種別と市町村・廃止の理由は、県への廃止届の確認を町にお願いする。",
        ]
        for c, v in enumerate(atai, 1):
            ws2.cell(r, c).value = v
    wb.save(WBS)

    futeki = []

    def judge(ok, mi):
        print(("   ○ " if ok else "   × ") + mi)
        if not ok:
            futeki.append(mi)

    print("業務工程管理表の更新（軽費老人ホームの廃止・素案Ver.2.25）")
    judge(n in (0, 1), f"No.154 に書いた（{n}件）")
    judge(n_seika in (0, 2), f"成果品4・6 を更新した（{n_seika}件）")
    wb2 = openpyxl.load_workbook(WBS, data_only=True)
    ws3 = wb2["03_確認事項一覧"]
    nos = sorted(int(s) for s in (str(ws3.cell(x, 1).value or "").strip()
                                  for x in range(5, ws3.max_row + 1))
                 if s.isdigit())
    judge(nos == list(range(1, max(nos) + 1)), f"欠番0（{len(nos)}件）")
    jun = [str(wb2["08_作業順位"].cell(x, 1).value or "").strip()
           for x in range(5, wb2["08_作業順位"].max_row + 1)]
    judge(jun.count("39") == 1, "作業順位39 が1行だけある")
    s = "\n".join(str(v) for sh in wb2 for row in sh.iter_rows(values_only=True)
                  for v in row if v is not None)
    for na in ("蔵王グリーンホーム", "鶴寿会", "0224-84-4853", "釜房みどりの園",
               "アルパイン", "ながさか", "リリーハイム", "川崎市", "北塩原",
               "金ヶ崎"):
        judge(na not in s, f"「{na}」が入っていない")
    judge(SOAN in s and SHIRYO in s and "想定問答集_R8.11_v8" in s,
          "成果品管理に素案 Ver.2.25・資料 v12・問答集 v8 が載っている")
    if futeki:
        sys.exit(1)
    print("   ○ 自己点検 不適合0件")


if __name__ == "__main__":
    main()
