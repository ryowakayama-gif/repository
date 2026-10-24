# -*- coding: utf-8 -*-
"""継続101事業の本文記述のための作業台帳（令和8年10月2日）

08_作業順位 の順位6。素案に残る作業のうち最も大きいのが
「継続101事業の本文の記述」である。

**事業名と所管は統合体系表にあるが、事業の内容そのものは書いていない。**
内容を当方で作文すると、実際に行われていない事業を書くおそれがある。
継続事業は第9期から続くものであるから、**第9期計画書（令和6年3月）の
記述を素材とする**のが確かである。

本スクリプトは、統合体系表の継続101事業について

  ① 第9期計画書に対応する記述があるか
  ② あるならその頁と記述の抜粋
  ③ ないなら、何を確認すれば書けるか

を機械で突き合わせ、作業台帳にする。
**記述そのものは書かない。素材を揃えるところまでを機械で行う。**

  出力　05_試算・管理シート/川崎町_継続101事業_本文記述_作業台帳_R8.10.2.xlsx

⚠ 第9期計画書の記述をそのまま写すのではない。
  第10期の内容として書き改める必要があり、
  その際に町のご確認（確認事項No.8 第9期の施策・事業実績）を要する。
"""
import re
import sys

import openpyxl
import pypdf
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

TAIKEI = "05_試算・管理シート/川崎町_第10期_施策事業統合体系表_R8.9.25b.xlsx"
PDF9 = "09_元資料/川崎町_第9期計画_04324.pdf"
OUT = "05_試算・管理シート/川崎町_継続101事業_本文記述_作業台帳_R8.10.2.xlsx"

TITLE = PatternFill("solid", fgColor="1F4E78")
HEAD = PatternFill("solid", fgColor="5B9BD5")
LEAD = PatternFill("solid", fgColor="F2F2F2")
OKF = PatternFill("solid", fgColor="E2EFDA")
NGF = PatternFill("solid", fgColor="FCE4D6")


def norm(s):
    """突き合わせ用に、番号・括弧・記号・空白を取り除く。"""
    s = re.sub(r"^（[0-9０-９]+）", "", str(s or ""))
    s = re.sub(r"【.*?】", "", s)
    return re.sub(r"[\s　（）()・,，、。]", "", s)


def main():
    wb = openpyxl.load_workbook(TAIKEI, data_only=True)
    ws = wb["01_統合体系表"]
    rows = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 4).value
        if not name:
            continue
        rows.append(dict(
            sho=str(ws.cell(r, 1).value or ""), setsu=str(ws.cell(r, 2).value or ""),
            ko=str(ws.cell(r, 3).value or ""), jigyo=str(name),
            kubun=str(ws.cell(r, 5).value or ""),
            shokan=str(ws.cell(r, 7).value or ""),
            kpi=str(ws.cell(r, 8).value or ""),
            bikou=str(ws.cell(r, 9).value or "")))
    keizoku = [x for x in rows if x["kubun"] == "継続"]

    # 第9期計画書の本文を頁ごとに読む
    rd = pypdf.PdfReader(PDF9)
    pages = [(i + 1, (p.extract_text() or "")) for i, p in enumerate(rd.pages)]
    npages = [(i, re.sub(r"[\s　（）()・,，、。]", "", t)) for i, t in pages]

    out = []
    n_hit = 0
    for x in keizoku:
        key = norm(x["jigyo"])[:10]
        hit = None
        for i, t in npages:
            if key and key in t:
                hit = i
                break
        if hit:
            n_hit += 1
            raw = pages[hit - 1][1]
            # 抜粋　事業名のあたりから200字
            k2 = norm(x["jigyo"])[:8]
            pos = re.sub(r"[\s　]", "", raw).find(k2)
            flat = re.sub(r"[\s　]+", " ", raw)
            j = flat.find(x["jigyo"].replace("（", "").replace("）", "")[:6])
            nuki = flat[max(0, j):max(0, j) + 220] if j >= 0 else flat[:220]
            out.append((x["sho"], x["setsu"], x["ko"], x["jigyo"], x["shokan"],
                        "あり", f"第9期計画 {hit}頁", nuki,
                        "第9期の記述を素材に、第10期の内容として書き改める。"
                        "実績の数値は確認事項No.8のご記入による。"))
        else:
            out.append((x["sho"], x["setsu"], x["ko"], x["jigyo"], x["shokan"],
                        "なし", "―", "",
                        "第9期計画書に同じ名称の記述がない。"
                        "名称が変わったか、第9期の本文に書かれていない事業。"
                        "町に事業の概要をご確認いただく（確認事項No.8）。"))

    nb = openpyxl.Workbook()
    ws2 = nb.active
    ws2.title = "01_継続101事業"
    ws2.cell(1, 1).value = "継続101事業の本文記述のための作業台帳"
    ws2.cell(1, 1).font = Font(name="游ゴシック", size=14, bold=True,
                               color="FFFFFF")
    ws2.cell(1, 1).fill = TITLE
    ws2.cell(2, 1).value = (
        "統合体系表の継続101事業について、第9期計画書（令和6年3月）に"
        "対応する記述があるかを機械で突き合わせたものです。"
        "⚠ 事業の内容を当方で作文すると、実際に行われていない事業を"
        "書くおそれがあります。継続事業は第9期から続くものですから、"
        "第9期計画書の記述を素材とし、第10期の内容として書き改めます。"
        "「なし」の事業は、町に事業の概要をご確認いただく必要があります"
        "（確認事項No.8）。")
    ws2.cell(2, 1).font = Font(name="游ゴシック", size=9)
    ws2.cell(2, 1).fill = LEAD
    ws2.cell(2, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws2.merge_cells(start_row=2, start_column=1, end_row=2, end_column=9)
    ws2.row_dimensions[2].height = 62
    head = ["章", "節", "項", "事業", "所管", "第9期計画の記述",
            "該当箇所", "第9期計画の記述（抜粋）", "本文を書くときの扱い"]
    for c, (h, w) in enumerate(zip(head, [26, 30, 26, 36, 20, 12, 15, 76, 50]),
                               start=1):
        cell = ws2.cell(4, c)
        cell.value = h
        cell.font = Font(name="游ゴシック", size=9, bold=True, color="FFFFFF")
        cell.fill = HEAD
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        ws2.column_dimensions[get_column_letter(c)].width = w
    for i, rec in enumerate(out):
        for c, v in enumerate(rec, start=1):
            cell = ws2.cell(5 + i, c)
            cell.value = v
            cell.font = Font(name="游ゴシック", size=9)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.fill = OKF if rec[5] == "あり" else NGF
    ws2.freeze_panes = "E5"

    ws3 = nb.create_sheet("02_まとめ")
    ws3.cell(1, 1).value = "まとめ"
    ws3.cell(1, 1).font = Font(name="游ゴシック", size=14, bold=True,
                               color="FFFFFF")
    ws3.cell(1, 1).fill = TITLE
    import collections
    by_sho = collections.Counter(x[0] for x in out)
    hit_sho = collections.Counter(x[0] for x in out if x[5] == "あり")
    rr = 3
    for c, (h, w) in enumerate(zip(["章", "継続事業", "記述あり", "記述なし"],
                                   [34, 10, 10, 10]), start=1):
        cell = ws3.cell(rr, c)
        cell.value = h
        cell.font = Font(name="游ゴシック", size=9, bold=True, color="FFFFFF")
        cell.fill = HEAD
        ws3.column_dimensions[get_column_letter(c)].width = w
    rr += 1
    for k in by_sho:
        for c, v in enumerate((k, by_sho[k], hit_sho[k],
                               by_sho[k] - hit_sho[k]), start=1):
            cell = ws3.cell(rr, c)
            cell.value = v
            cell.font = Font(name="游ゴシック", size=9)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        rr += 1
    for c, v in enumerate(("計", len(out), n_hit, len(out) - n_hit), start=1):
        cell = ws3.cell(rr, c)
        cell.value = v
        cell.font = Font(name="游ゴシック", size=9, bold=True)
        cell.fill = OKF
    nb.save(OUT)

    print("継続101事業の作業台帳")
    print("  保存：", OUT)
    print(f"  継続事業 {len(keizoku)}件"
          f"／第9期計画に記述あり {n_hit}件・なし {len(out) - n_hit}件")
    print("  ── 自己点検")
    if len(keizoku) != 101:
        print(f"   × 継続事業が{len(keizoku)}件（統合体系表は101件のはず）")
        sys.exit(1)
    print("   ○ 継続事業は101件（統合体系表と一致）")
    print("  ⚠ 記述そのものは書いていません。素材を揃えるところまでです。")


if __name__ == "__main__":
    main()
