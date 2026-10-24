# -*- coding: utf-8 -*-
"""公表版に向けた整理（令和8年10月2日）

08_作業順位 の順位7。確認事項No.162（未確定箇所の表記の扱い）の
ご判断を待たずに、**何がどこにいくつあるかを数え、
除去と置換の手順を用意する。**

  既定では **数えるだけ**（--apply を付けたときだけ書き換える）。

  python3 07_ソーススクリプト/build_kohyo_seiri_R8.10.2.py
      → 05_試算・管理シート/川崎町_公表版への整理_R8.10.2.xlsx（一覧）
  python3 07_ソーススクリプト/build_kohyo_seiri_R8.10.2.py --apply
      → 01_第10期_最新版成果品/川崎町_計画書_公表版_準備_R8.10.2.docx

取り除くもの（協議の段階では有用であり、公表版でのみ落とす）

  ① 本文・表に残る「確認事項No.〜」の参照
  ② 【町確認】【委員会協議】等の、［　］以外の括弧による未確定箇所
  ③ 受託者を主語とする語（策定支援の表示・当方の作業環境の断り・工程表の担当）

**計画は川崎町が定めるものである。本文にも見出しにも受託者を主語とする語を
置かない。** 一方、協議用の素案では、町が業務工程管理表を引けるよう
確認事項No.の参照を残す。

⚠ 図の中の文字は機械では拾えない。公表の前に目で確かめる。
"""
import collections
import copy
import os
import re
import shutil
import sys

import docx
import openpyxl
from docx.oxml.ns import qn
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SRC = "01_第10期_最新版成果品/川崎町_計画書素案_v2.8_第9期対比版.docx"
DST = "01_第10期_最新版成果品/川崎町_計画書_公表版_準備_R8.10.2.docx"
XL = "05_試算・管理シート/川崎町_公表版への整理_R8.10.2.xlsx"

# ① 確認事項の参照　「（確認事項No.12）」「確認事項No.12・No.13」ほか
RE_KAKUNIN = re.compile(r"（?確認事項No\.[0-9０-９]+(?:[・，,]\s*No\.[0-9０-９]+)*）?")
# ② ［　］以外の括弧による未確定箇所
RE_KAKKO = re.compile(r"[\[【〔][^\]】〕]{0,24}[\]】〕]")
# ③ 受託者を主語とする語（文脈を見て決めるため、置換は手当てを個別に書く）
JUTAKU = [
    ("（策定支援：ビズアップ公共コンサルティング株式会社）", "",
     "表紙の策定支援の表示。計画は町が定めるものであるため落とす"),
    ("当社の作業環境のネットワーク制限により厚生労働省のウェブサイトを"
     "参照できないため、", "",
     "当方の作業経過であり計画の記載事項ではない"),
    ("（受託者＋町担当課）", "（町担当課）",
     "工程表の担当欄。公表版では町の作業として示す"),
]

TITLE = PatternFill("solid", fgColor="1F4E78")
HEAD = PatternFill("solid", fgColor="5B9BD5")
LEAD = PatternFill("solid", fgColor="F2F2F2")
NGF = PatternFill("solid", fgColor="FFC7CE")
WARN = PatternFill("solid", fgColor="FCE4D6")


def walk(doc):
    """本文・表（入れ子を含む）・ヘッダー・フッターの段落を、場所とともに返す。"""
    for i, p in enumerate(doc.paragraphs, 1):
        yield ("本文", f"段落{i}", p._p)
    for ti, t in enumerate(doc.tables, 1):
        for ri, r in enumerate(t.rows, 1):
            for ci, c in enumerate(r.cells, 1):
                for p in c.paragraphs:
                    yield ("表", f"表{ti}[{ri},{ci}]", p._p)
    for si, s in enumerate(doc.sections, 1):
        for nm, part in (("ヘッダー", s.header), ("フッター", s.footer)):
            for p in part.paragraphs:
                yield (nm, f"{nm}{si}", p._p)


def text_of(el):
    return "".join(n.text or "" for n in el.iter(qn("w:t")))


def set_el(el, text):
    rs = el.findall(qn("w:r"))
    if not rs:
        return False
    for r in rs[1:]:
        el.remove(r)
    ts = rs[0].findall(qn("w:t"))
    if not ts:
        return False
    for t in ts[1:]:
        rs[0].remove(t)
    ts[0].text = text
    ts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return True


def scan(doc):
    rows = []
    for kind, where, el in walk(doc):
        t = text_of(el)
        if not t.strip():
            continue
        for m in RE_KAKUNIN.finditer(t):
            rows.append(("① 確認事項の参照", kind, where, m.group(0),
                         "取り除く", t.strip()[:70]))
        for m in RE_KAKKO.finditer(t):
            g = m.group(0)
            # ［　］以外の括弧には2種類ある。
            #   未確定箇所　　…… 決着しだい内容に置き換えるもの
            #   見出しのラベル…… 体系の表などで区分を示すもの（残す）
            if g in ("【町確認】", "【委員会協議】", "【町確認・最新値】"):
                rows.append(("②-1 未確定箇所", kind, where, g,
                             "確定した内容に置き換える", t.strip()[:70]))
            else:
                rows.append(("②-2 見出しのラベル", kind, where, g,
                             "公表版でも残す（未確定箇所ではない）",
                             t.strip()[:70]))
        for a, b, note in JUTAKU:
            if a in t:
                rows.append(("③ 受託者を主語とする語", kind, where, a[:30],
                             note, t.strip()[:70]))
    return rows


def apply(doc):
    n = collections.Counter()
    for kind, where, el in walk(doc):
        t = text_of(el)
        if not t.strip():
            continue
        new = t
        for a, b, note in JUTAKU:
            if a in new:
                new = new.replace(a, b)
                n["③"] += 1
        c = len(RE_KAKUNIN.findall(new))
        if c:
            new = RE_KAKUNIN.sub("", new)
            n["①"] += c
        new = re.sub(r"[。、]\s*。", "。", new)
        new = re.sub(r"（\s*）", "", new)
        if new != t:
            if not set_el(el, new):
                n["書き換えられず"] += 1
    return n


def write_xl(rows, doc):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "01_取り除くもの"
    ws.cell(1, 1).value = "公表版に向けて取り除くもの"
    ws.cell(1, 1).font = Font(name="游ゴシック", size=14, bold=True,
                              color="FFFFFF")
    ws.cell(1, 1).fill = TITLE
    ws.cell(2, 1).value = (
        "計画素案 Ver.2.8 を走査した結果です。"
        "いずれも協議の段階では有用なものであり、"
        "当方の判断では消していません（確認事項No.162）。"
        "②の【町確認】【委員会協議】は、公表版では確定した内容に"
        "置き換える必要があります。"
        "⚠ 図の中の文字は機械では拾えません。公表の前に目で確かめます。")
    ws.cell(2, 1).font = Font(name="游ゴシック", size=9)
    ws.cell(2, 1).fill = LEAD
    ws.cell(2, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)
    ws.row_dimensions[2].height = 56
    head = ["区分", "場所", "位置", "当たった文字", "公表版での扱い",
            "前後の文"]
    for c, (h, w) in enumerate(zip(head, [22, 9, 16, 26, 40, 76]), start=1):
        cell = ws.cell(4, c)
        cell.value = h
        cell.font = Font(name="游ゴシック", size=9, bold=True, color="FFFFFF")
        cell.fill = HEAD
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = w
    for i, rec in enumerate(rows):
        for c, v in enumerate(rec, start=1):
            cell = ws.cell(5 + i, c)
            cell.value = v
            cell.font = Font(name="游ゴシック", size=9)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.fill = NGF if rec[0].startswith("③") else WARN
    ws.freeze_panes = "A5"

    # まとめ
    ws2 = wb.create_sheet("02_まとめ")
    cnt = collections.Counter(r[0] for r in rows)
    kakko = collections.Counter(r[3] for r in rows if r[0].startswith("②"))
    mikakutei = sum(v for k, v in kakko.items()
                    if k in ("【町確認】", "【委員会協議】",
                             "【町確認・最新値】"))
    ws2.cell(1, 1).value = "まとめ"
    ws2.cell(1, 1).font = Font(name="游ゴシック", size=14, bold=True,
                               color="FFFFFF")
    ws2.cell(1, 1).fill = TITLE
    r = 3
    for k, v in cnt.most_common():
        ws2.cell(r, 1).value = k
        ws2.cell(r, 2).value = v
        r += 1
    r += 1
    ws2.cell(r, 1).value = (
        f"②のうち、確定した内容に置き換えを要する未確定箇所は{mikakutei}件。"
        "残りは体系の表などの見出しのラベルであり、公表版でも残す。")
    ws2.cell(r, 1).font = Font(name="游ゴシック", size=9)
    r += 2
    ws2.cell(r, 1).value = "②の内訳（括弧の種類ごと）"
    ws2.cell(r, 1).font = Font(name="游ゴシック", size=10, bold=True)
    r += 1
    for k, v in kakko.most_common():
        ws2.cell(r, 1).value = k
        ws2.cell(r, 2).value = v
        r += 1
    for c, w in zip("AB", (40, 10)):
        ws2.column_dimensions[c].width = w
    wb.save(XL)
    return cnt, kakko, mikakutei


def main():
    do_apply = "--apply" in sys.argv
    doc = docx.Document(SRC)
    rows = scan(doc)
    cnt, kakko, mikakutei = write_xl(rows, doc)
    print("公表版への整理")
    print("  走査した素案：", os.path.basename(SRC))
    print("  保存：", XL)
    for k, v in cnt.most_common():
        print(f"   {k}　{v}件")
    print("   ②の内訳：" + "／".join(f"{k} {v}" for k, v in kakko.most_common()))
    print(f"   うち置き換えを要する未確定箇所 {mikakutei}件"
          "（残りは体系の表などの見出しのラベルであり、公表版でも残す）")
    if not do_apply:
        print("  ── 数えただけです。書き換えるには --apply を付けます。")
        print("  ⚠ 確認事項No.162 のご判断（現在の形のまま進めるか、"
              "［　］に統一するか）を待ちます。")
        return
    shutil.copy(SRC, DST)
    d2 = docx.Document(DST)
    n = apply(d2)
    d2.save(DST)
    print("  書き換えました：", DST)
    for k, v in n.items():
        print(f"   {k}　{v}件")
    # 自己点検　確認事項の参照が残っていないか
    d3 = docx.Document(DST)
    left = sum(1 for _, _, el in walk(d3)
               if RE_KAKUNIN.search(text_of(el)))
    print("  ── 自己点検")
    if left:
        print(f"   × 確認事項の参照が{left}件残っている")
        sys.exit(1)
    print("   ○ 確認事項の参照は残っていない")
    print("   ⚠ ②の未確定箇所は、確定した内容に置き換える必要があるため"
          "機械では消していません。")


if __name__ == "__main__":
    main()
