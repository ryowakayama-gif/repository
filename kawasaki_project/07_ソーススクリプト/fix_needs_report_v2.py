# -*- coding: utf-8 -*-
"""
川崎町 ニーズ調査結果報告書 R8.9.2版 → R8.9.2版rev2
redteamレビュー R9-1／R9-2／R9-3／R9-4／R9-5 の反映

  R9-1 第4章見出しの計画名から「保健」が脱落 → 正式名称に統一
  R9-2 KPI・第9期計画との接続が本報告書に無い → 取扱いを明記する注記を追加
  R9-3 BMI区分の判定で丸めの有無が未記載 → 丸め前の値で判定した旨と感度を追記
  R9-4 分割した2表の列構成が揃っていない（結合セルが重複列に見える） → 2列表に整理
  R9-5 BMI区分と低栄養リスクが1表に同居 → 2表に分割

数値は変更しない。
"""
import copy
import re
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

BASE = Path(__file__).resolve().parent.parent
SRC = BASE / "01_第10期_最新版成果品/川崎町_ニーズ調査結果報告書_R8.9.2版.docx"
OUT = SRC

T_ACT, T_INT, T_BMI = 50, 51, 95     # 活動別／参加意向／BMI・低栄養
GRID_ACT = ["4000", "3600"]
GRID_INT = ["5800", "1800"]

NOTE_KPI = (
    "※ 第９期計画で設定した目標値（KPI）の達成状況、及び前回計画で整理した課題との対応関係については、"
    "第１回策定委員会資料において別途整理している。本報告書は、今回の調査結果から読み取れる課題の整理に限定している。"
)
ADD_BMI = (
    "なお、区分の判定は、算出したBMIを丸める前の値により行っている"
    "（小数第1位に丸めてから判定した場合、境界値付近の該当が変わり「やせ」27人・「標準」351人・「肥満」155人となる）。"
)

log = []


def tables_of(doc):
    return [Table(c, doc) for c in doc.element.body.iterchildren() if c.tag == qn("w:tbl")]


def node_text(el):
    return "".join(n.text or "" for n in el.iter(qn("w:t")))


def node_sub(el, pattern, repl):
    ts = list(el.iter(qn("w:t")))
    if not ts:
        return False
    whole = "".join(t.text or "" for t in ts)
    new = re.sub(pattern, repl, whole)
    if new == whole:
        return False
    ts[0].text = new
    ts[0].set(qn("xml:space"), "preserve")
    for t in ts[1:]:
        t.text = ""
    return True


def find_par(doc, needle):
    for p in doc.paragraphs:
        if needle in p.text:
            return p
    raise KeyError(needle)


def clone_par(tpl_par, text):
    p = copy.deepcopy(tpl_par._p)
    par = Paragraph(p, tpl_par._parent)
    for r in par.runs[1:]:
        r._element.getparent().remove(r._element)
    par.runs[0].text = text
    return p


def blank_par(tpl_par):
    p = copy.deepcopy(tpl_par._p)
    for r in Paragraph(p, tpl_par._parent).runs:
        r._element.getparent().remove(r._element)
    return p


# ---------------------------------------------------------------- R9-1
def fix_plan_name(doc):
    p = find_par(doc, "４．高齢者福祉計画・介護保険事業計画策定に向けた課題整理")
    node_sub(p._p, "高齢者福祉計画", "高齢者保健福祉計画")
    log.append("[R9-1] 第4章見出しを「４．高齢者保健福祉計画・介護保険事業計画策定に向けた課題整理」に統一")


# ---------------------------------------------------------------- R9-2
def add_kpi_note(doc, note_tpl):
    anchor = find_par(doc, "母数が小さいからこそ地域包括支援センター等による個別のフォローアップ体制")
    anchor._p.addnext(clone_par(note_tpl, NOTE_KPI))
    log.append("[R9-2] 第4章末に、KPI・第９期計画との接続の取扱いを示す注記を追加")


# ---------------------------------------------------------------- R9-3
def add_bmi_note(doc):
    p = find_par(doc, "【BMI区分】身長・体重の両方に有効回答がある方について")
    node_sub(p._p, r"「不明」とした。$", "「不明」とした。" + ADD_BMI)
    log.append("[R9-3] BMI区分の判定方法に、丸め前の値で判定した旨と丸めた場合の感度を追記")


# ---------------------------------------------------------------- R9-4
def flatten_span(doc, ti, grid):
    """結合セルを解消して素直な2列表にする（見た目の列幅は維持）。"""
    tbl = tables_of(doc)[ti]._tbl
    g = tbl.find(qn("w:tblGrid"))
    for col in g.findall(qn("w:gridCol")):
        g.remove(col)
    for w in grid:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), w)
        g.append(col)
    for tr in tbl.findall(qn("w:tr")):
        for tc, w in zip(tr.findall(qn("w:tc")), grid):
            tcpr = tc.find(qn("w:tcPr"))
            gs = tcpr.find(qn("w:gridSpan")) if tcpr is not None else None
            if gs is not None:
                tcpr.remove(gs)
            tcw = tcpr.find(qn("w:tcW")) if tcpr is not None else None
            if tcw is not None:
                tcw.set(qn("w:w"), w)
                tcw.set(qn("w:type"), "dxa")
    log.append(f"[R9-4] 表{ti}：結合セルを解消し {len(grid)}列表に整理（列幅 {'/'.join(grid)}）")


# ---------------------------------------------------------------- R9-5
def split_bmi_table(doc, blank_tpl):
    tbl = tables_of(doc)[T_BMI]._tbl
    rows = tbl.findall(qn("w:tr"))
    at = next(i for i, tr in enumerate(rows) if node_text(tr).startswith("低栄養リスク"))
    new_tbl = copy.deepcopy(tbl)
    for tr in rows[at:]:
        tbl.remove(tr)
    for tr in new_tbl.findall(qn("w:tr"))[:at]:
        new_tbl.remove(tr)
    tbl.addnext(new_tbl)
    tbl.addnext(blank_par(blank_tpl))
    log.append(f"[R9-5] 表{T_BMI}を「BMI区分」と「低栄養リスク」の2表に分割（第{at}行で分割）")


# ---------------------------------------------------------------- 実行
def main():
    doc = Document(str(SRC))
    note_tpl = find_par(doc, "※ 転倒予防・移動手段の確保は")

    fix_plan_name(doc)
    add_bmi_note(doc)
    add_kpi_note(doc, note_tpl)
    flatten_span(doc, T_ACT, GRID_ACT)
    flatten_span(doc, T_INT, GRID_INT)
    split_bmi_table(doc, note_tpl)      # 表番号がずれるため最後

    doc.save(str(OUT))
    print("\n".join(log))
    print(f"\n出力: {OUT}")


if __name__ == "__main__":
    main()
