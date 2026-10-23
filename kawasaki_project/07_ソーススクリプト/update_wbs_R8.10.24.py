# -*- coding: utf-8 -*-
"""業務工程管理表の更新（Red Team レビュー（令和8年10月24日）への対応）

  python3 07_ソーススクリプト/update_wbs_R8.10.24.py

  00_管理方法　　　基準日を令和8年10月24日に（件数は台帳から数える）
  03_確認事項一覧　No.85・117・148・163 に追記
  04_成果品管理　　稿を 素案 Ver.2.32／公表用・概要版 v2.32／資料 v17／問答集 v14 に
  08_作業順位　　　49（完了）を加える
再実行しても二重に書かない。
"""
import collections
import copy
import sys

import openpyxl
from openpyxl.styles import Alignment, Font

WBS = "川崎町_業務工程管理表.xlsx"
FONT = Font(name="游ゴシック", size=9)
HIZUKE = "【令和8年10月24日 更新（Red Team レビューへの対応）】"
TAIOU = "川崎町_RedTeamレビュー対応表_R8.10.24.md"
ZIP = "11_提出一式/川崎町_RedTeamレビュー対応_成果物とエビデンス_R8.10.24.zip"
OKIKAE = [("川崎町_計画書素案_v2.31_協議用.docx", "川崎町_計画書素案_v2.32_協議用.docx"),
          ("川崎町_計画書案_公表用_v2.31.docx", "川崎町_計画書案_公表用_v2.32.docx"),
          ("川崎町_計画書案_概要版_v2.31.docx", "川崎町_計画書案_概要版_v2.32.docx"),
          ("川崎町_第2回策定委員会資料_R8.11_v16.docx", "川崎町_第2回策定委員会資料_R8.11_v17.docx"),
          ("同_想定問答集_R8.11_v13.docx", "同_想定問答集_R8.11_v14.docx")]
TSUIKI = {
    "85": (f"{HIZUKE}Red Team レビューを受け、資料を v17・想定問答集を v14 に上げた。選択２の保険料（5,932円・"
           "6,354円。旧の5,959円・6,382円は算定Ｂ6,170円の時の値）、表27 の基準（算定Ａ）と算定Ｂ・Ｃの差、"
           "交付金の要件の言い切り、国の基金の水準の記述（根拠を確かめられず削除）、8事業所・14指定、"
           f"選択１の仮置きの注記。{TAIOU}。　"),
    "117": (f"{HIZUKE}生活支援コーディネーター数を原票で確かめた。町の実績データ（R6・R7見込とも26人）と国の"
            "令和8年度評価（該当状況調査26人）が一致。素案の25名は地域福祉計画の資料の生活支援サポーターの"
            "値であったため、26人に改めた（素案 Ver.2.32）。地域ケア会議への出席は令和6年度4人。　"),
    "148": (f"{HIZUKE}保険料を動かす要因の差額は算定Ａ（取崩しなし）を基準とした値であることを資料・問答集に"
            "明記。報酬改定1％は算定Ａ＋66円・Ｂ＋92円・Ｃ＋119円。入所・居住系±10人は約211円"
            "（基金の取崩しによらない。表101 から検算211.7円）。　"),
    "163": (f"{HIZUKE}素案 1-5 の基本指針案の別表11事項の件数を表から数え直した（対応済1・一部8・未対応2）。"
            "国の告示・県の第10期計画（中間案）を受けて照合し直す。　"),
}


def iro(c):
    c.font = FONT
    c.alignment = Alignment(wrap_text=True, vertical="top")


def mae_ni(c, s):
    if HIZUKE in str(c.value or ""):
        return 0
    c.value = s + str(c.value or "")
    iro(c)
    return 1


def main():
    wb = openpyxl.load_workbook(WBS)
    kekka = []
    ws3 = wb["03_確認事項一覧"]
    k3 = collections.Counter(str(ws3.cell(y, 8).value).strip() for y in range(5, ws3.max_row + 1)
                             if str(ws3.cell(y, 1).value or "").strip().isdigit())
    n3 = sum(k3.values())

    ws0 = wb["00_管理方法"]
    for r in ws0.iter_rows():
        for c in r:
            v = str(c.value or "")
            if v.startswith("令和8年10月23日　本表の記載は同日時点のもの。"):
                c.value = ("令和8年10月24日　本表の記載は同日時点のもの。Red Team レビュー（令和8年10月24日）を受け、"
                           "協議用素案 Ver.2.32（118頁）・公表用の計画書（案）v2.32（116頁）・概要版 v2.32（8頁）・"
                           "第2回策定委員会資料 v17（21頁）・想定問答集 v14（全69問）に是正した。指摘27件のうち"
                           "是正済16・一部是正3・既に対応1、残る7件は資料の未受領又は町のご判断による。"
                           f"点検ログ14件はすべて終了コード0。確認事項は{n3}件のうち{k3['完了']}件が完了。"
                           f"一式は {ZIP}。\n（前の基準日の記載）" + v)
                iro(c)
                kekka.append((True, "基準日を令和8年10月24日に"))

    k = 0
    for y in range(5, ws3.max_row + 1):
        no = str(ws3.cell(y, 1).value or "").strip()
        if no in TSUIKI:
            k += mae_ni(ws3.cell(y, 10), TSUIKI[no])
    kekka.append((k in (0, len(TSUIKI)), f"確認事項への追記 {k}件"))

    ws4 = wb["04_成果品管理"]
    n = 0
    for y in range(5, ws4.max_row + 1):
        na = str(ws4.cell(y, 3).value or "")
        nb = na
        for a, b in OKIKAE:
            nb = nb.replace(a, b)
        if nb != na:
            ws4.cell(y, 3).value = nb
            mae_ni(ws4.cell(y, 8), f"{HIZUKE}Red Team レビューの指摘を是正した版に上げた（{TAIOU}）。　")
            n += 1
    kekka.append((n in (0, 3), f"成果品の稿 {n}行"))

    ws8 = wb["08_作業順位"]
    aru = {str(ws8.cell(x, 1).value or "").strip() for x in range(5, ws8.max_row + 1)}
    if "49" not in aru:
        moto = ws8.max_row
        r = moto + 1
        for c in range(1, 11):
            ws8.cell(r, c)._style = copy.copy(ws8.cell(moto, c)._style)
        ws8.row_dimensions[r].height = ws8.row_dimensions[moto].height
        for c, v in enumerate([
                "49", "Red Team レビュー（令和8年10月24日）への対応と成果物・エビデンスの一式", "R8.10.24",
                "ご指示による。指摘27件を最新の版で確かめ、当てはまるものを是正した",
                "No.85・117・148・163・189", "素案 Ver.2.32／公表用・概要版 v2.32／資料 v17／問答集 v14／ZIP",
                "1日", "完了",
                "新たに見つけた誤り：9-4（実在しない節）の参照2か所、目次のご挨拶の頁、生活支援コーディネーター"
                "の人数。対応表・点検ログ・検算・差分・PDF を ZIP に同梱。"], 1):
            ws8.cell(r, c).value = v
    aru = {str(ws8.cell(x, 1).value or "").strip() for x in range(5, ws8.max_row + 1)}
    kekka.append(("49" in aru, "作業順位 49"))

    wb.save(WBS)
    print("業務工程管理表の更新（Red Team レビューへの対応）")
    ng = 0
    for ok, mi in kekka:
        print(("   ○ " if ok else "   × ") + mi)
        ng += not ok
    if ng:
        sys.exit(1)


if __name__ == "__main__":
    main()
