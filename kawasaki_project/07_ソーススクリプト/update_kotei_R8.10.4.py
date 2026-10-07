# -*- coding: utf-8 -*-
"""業務工程管理表の更新（令和8年10月4日・順位6の完了）

  03_確認事項一覧　No.164 を追加（全164件）
  07_ペンディング整理　束12（素案の個別の記述）に No.164 を加える
  08_作業順位　　　順位6 を完了に改める
  00・01・02・04　素案 Ver.2.9 の状況

点検スキル `.claude/skills/plan-draft-check/` の7「確認事項を起こす前に」に
よる。別の成果品で挙げた論点は、必ず本表へ登録する。
"""
import copy
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

WB = "川崎町_業務工程管理表.xlsx"
KIJUN = "令和8年10月4日"
DONE = PatternFill("solid", fgColor="E2EFDA")

NEW = [
    ("164", "(7)", "R8.11",
     "事業一覧表に加えた「内容」のうち、当方が書き起こした48件のご確認",
     "計画素案 Ver.2.9 で、事業一覧表（17表・138事業）に「内容」の欄を"
     "加えました。内容の出所は3つで、①第9期計画書の本文から取り出したもの"
     "57件、②統合体系表の新規・拡充の内容33件、③当方が書き起こしたもの"
     "48件です。③は介護保険法・国の指針が内容を定めているもの"
     "（施設サービス等）と、素案の本文に記述があるもの（町独自の事業）に"
     "限っていますが、実際の運用と食い違っていないかは町でなければ"
     "確かめられません。1件ずつご確認をお願いします。",
     "計画素案 第5章・第6章の事業一覧表\n計画書（案）",
     "発注者", "確認待ち", "R8.11",
     "対象の一覧は 05_試算・管理シート/川崎町_事業の内容の出所_R8.10.4.xlsx"
     "（橙色の行が③）。"
     "当方の既定値は、ご指摘のないものはそのまま計画書（案）に載せる"
     "こととします。"
     "⚠ 第9期計画書に記述のある57件も、第9期の文をそのまま写したもの"
     "ではなく第10期の記述として整えています。あわせてご一読ください。"),
]


def unmerge_below(ws, at_row):
    kept = []
    for m in list(ws.merged_cells.ranges):
        if m.min_row >= at_row:
            kept.append((m.min_row, m.max_row, m.min_col, m.max_col))
            ws.unmerge_cells(str(m))
    return kept


def remerge(ws, kept, n):
    for r1, r2, c1, c2 in kept:
        ws.merge_cells(start_row=r1 + n, end_row=r2 + n,
                       start_column=c1, end_column=c2)


def copy_style(ws, src_row, dst_row, ncol):
    for c in range(1, ncol + 1):
        ws.cell(dst_row, c)._style = copy.copy(ws.cell(src_row, c)._style)
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height


def insert(ws, at, tpl, recs, ncol, count_col, old_last):
    n = len(recs)
    kept = unmerge_below(ws, at)
    ws.insert_rows(at, n)
    remerge(ws, kept, n)
    for i, rec in enumerate(recs):
        r = at + i
        copy_style(ws, tpl, r, ncol)
        for c, v in enumerate(rec, start=1):
            ws.cell(r, c).value = v
    new_last = old_last + n
    for r in range(at + n, ws.max_row + 1):
        cell = ws.cell(r, 2)
        if isinstance(cell.value, str) and cell.value.startswith("=COUNTIF"):
            cell.value = cell.value.replace(
                f"{count_col}5:{count_col}{old_last}",
                f"{count_col}5:{count_col}{new_last}")
    return new_last


def main():
    wb = openpyxl.load_workbook(WB)
    ws = wb["03_確認事項一覧"]
    last = None
    for r in range(5, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").strip() == "163":
            last = r
            break
    if last is None:
        raise SystemExit("No.163 の行が見つからない")
    insert(ws, last + 1, last, NEW, 10, "H", last)

    # ══════════════════════════ 07_ペンディング整理
    ws = wb["07_ペンディング整理"]
    for r in range(5, ws.max_row + 1):
        if str(ws.cell(r, 2).value or "").startswith("12　素案の個別の記述"):
            ws.cell(r, 5).value = (ws.cell(r, 5).value or 0) + 1
            ws.cell(r, 7).value = str(ws.cell(r, 7).value or "") + "、164"
            break

    # ══════════════════════════ 08_作業順位
    ws = wb["08_作業順位"]
    for r in range(5, ws.max_row + 1):
        if ws.cell(r, 1).value == 6:
            ws.cell(r, 8).value = "完了"
            ws.cell(r, 9).value = (
                f"{KIJUN}　計画素案 Ver.2.9（事業内容版・全109頁）。"
                "17の事業一覧表すべてを5欄にし、138行すべてに内容を入れた。"
                "出所は、第9期計画書57件・統合体系表（新規拡充）33件・"
                "当方が書き起こしたもの48件。作文はしていない。"
                "③の48件は町のご確認をお願いする（確認事項No.164）。"
                "頁数は本文104頁→109頁（＋5頁）。"
                "圧縮案を当てた v2.9c は本文102頁であり、"
                "内容の欄を加えてもVer.2.8（104頁）より2頁少ない。")
            for c in range(1, 10):
                ws.cell(r, c).fill = DONE
                ws.cell(r, c).font = Font(name="游ゴシック", size=9)
                ws.cell(r, c).alignment = Alignment(wrap_text=True,
                                                    vertical="top")
            break

    # ══════════════════════════ 00・01・04
    ws = wb["00_管理方法"]
    ws.cell(14, 2).value = (
        f"{KIJUN}　本表の記載は同日時点のもの。"
        "計画素案 Ver.2.9（事業内容版・全109頁・図10点・"
        "事業一覧表17表に「内容」の欄）、"
        "圧縮案 v2.9c（本文102頁）と別冊 詳細分析（9頁）、"
        "第2回策定委員会資料 v6、想定問答集 v3（全39問）、"
        "策定委員会 逆算工程表（6シート）、第9期 計画値実績対比表（8シート）、"
        "調査報告書3文書（納品前処理済み）、公表版への整理、概要版 構成案、"
        "事業の内容の出所（2シート）、資料編の3表（3シート）、"
        "図表データ管理台帳（17シート・図10点）まで作成済み。"
        "点検は、素案の自己点検32件・保険料の検算39項目・書体の点検13件・"
        "図表の自己点検10件・委員会資料の自己点検27件がいずれも適合。"
        "確認事項は164件のうち19件が決着し、145件が未決着。"
        "作業順位10件のうち、順位1は町対応のため取り下げ、他の9件は実施済み。")

    ws = wb["01_業務内容別の進捗"]
    ws.cell(11, 5).value = (
        "計画書素案_v2.9_事業内容版（全109頁・図10点・事業一覧表に内容の欄）\n"
        "同_v2.9c_圧縮案（本文102頁）／計画書_別冊_詳細分析（9頁）\n"
        "事業の内容の出所_R8.10.4（2シート・138事業）\n"
        "図表データ管理台帳_R8.9.30（17シート・自己点検10件）\n"
        "第9期_計画値実績対比表_R8.10.1（8シート）\n"
        "策定委員会_逆算工程表_R8.10.1（6シート）\n"
        "公表版への整理_R8.10.2／概要版_構成案_R8.10.2\n"
        "資料編_3表_R8.10.2／施策・事業統合体系表_R8.9.25b（143事業）")
    ws.cell(11, 7).value = 0.99
    ws.cell(11, 9).value = (
        "素案はVer.2.9まで作成（全109頁・660段落・図10点）。"
        "Ver.2.9で事業一覧表17表に「内容」の欄を加え、"
        "138事業すべてに内容を入れた。これにより、名称だけの一覧から"
        "計画として読める形になった。"
        "残るのは第3章の町確認欄（ペンディング束3）と、"
        "当方が書き起こした48件のご確認（確認事項No.164）、"
        "成果品の頁数の割付け（確認事項No.83）。")
    ws.cell(12, 7).value = 0.5
    ws.cell(12, 9).value = (
        "中間報告の提出期日が指定されていない（確認事項No.17）。"
        "計画書（案）は素案Ver.2.9を母体とする。"
        "頁数は圧縮案 v2.9c（本文102頁・計115頁）をお示しした。"
        "事業一覧表の内容の書き下ろしが終わったため、"
        "残るのは概要版（8頁程度・3,500部）の作成と、"
        "町確認欄の確定を待った差し替えである。")

    ws = wb["04_成果品管理"]
    ws.cell(8, 3).value = ("川崎町_計画書素案_v2.9_事業内容版.docx"
                           "／同_v2.9c_圧縮案_R8.10.4.docx")
    ws.cell(8, 6).value = 0.85
    ws.cell(8, 8).value = (
        f"【{KIJUN} 更新】素案Ver.2.9（本文109頁・図10点）。"
        "事業一覧表17表に「内容」の欄を加え、138事業すべてを埋めた。"
        "圧縮案 v2.9c は本文102頁・資料編13頁を含めて115頁。"
        "**仕様書8の「100頁程度」を本文のみと読むなら、"
        "圧縮案の本文102頁で収まる。**"
        "この読み方を含めてご判断いただきたい（確認事項No.83）。"
        "当方が書き起こした48件の内容は確認事項No.164。")

    wb.save(WB)
    print("保存：", WB)
    wb2 = openpyxl.load_workbook(WB)
    ws2 = wb2["03_確認事項一覧"]
    nos = [int(str(ws2.cell(r, 1).value).strip())
           for r in range(5, ws2.max_row + 1)
           if str(ws2.cell(r, 1).value or "").strip().isdigit()]
    miss = sorted(set(range(1, max(nos) + 1)) - set(nos))
    print("  確認事項", len(nos), "件／欠番", miss if miss else "なし")
    # 07シートの束の件数の合計が未決着の件数と一致するか
    ws3 = wb2["07_ペンディング整理"]
    tot = 0
    for r in range(5, 5 + 12):
        v = ws3.cell(r, 5).value
        if isinstance(v, int):
            tot += v
    nokori = sum(1 for r in range(5, ws2.max_row + 1)
                 if str(ws2.cell(r, 1).value or "").strip().isdigit()
                 and str(ws2.cell(r, 8).value or "").strip() != "完了")
    print(f"  ペンディングの束の合計 {tot}件／未決着 {nokori}件")
    if tot != nokori:
        print("   × 束の合計と未決着の件数が合わない")
        sys.exit(1)
    print("   ○ 束の合計と未決着の件数が一致")


if __name__ == "__main__":
    main()
