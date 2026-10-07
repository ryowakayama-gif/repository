# -*- coding: utf-8 -*-
"""
川崎町 介護実態調査結果報告書 R8.9.2版 — redteamレビュー R10-1／R10-2 の反映

  R10-1 第13章総括の外挿値が実態推計として読める／「最も届いていない」「最も大きな供給源」が断定的
        → 感度試算であることを明記し、追加確認が必要な層としての記述に改める
  R10-2 第10章③の換算値に「感度試算」の語がない → 補う

数値は変更しない。
"""
import re
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

BASE = Path(__file__).resolve().parent.parent
DOC = BASE / "01_第10期_最新版成果品/川崎町_介護実態調査結果報告書_R8.9.2版.docx"

log = []


def find_par(doc, needle):
    for p in doc.paragraphs:
        if needle in p.text:
            return p
    raise KeyError(needle)


def set_text(par, text):
    ts = list(par._p.iter(qn("w:t")))
    ts[0].text = text
    ts[0].set(qn("xml:space"), "preserve")
    for t in ts[1:]:
        t.text = ""


def sub_text(par, old, new):
    ts = list(par._p.iter(qn("w:t")))
    whole = "".join(t.text or "" for t in ts)
    out = re.sub(re.escape(old), lambda m: new, whole)
    if out == whole:
        return False
    ts[0].text = out
    ts[0].set(qn("xml:space"), "preserve")
    for t in ts[1:]:
        t.text = ""
    return True


def main():
    doc = Document(str(DOC))

    # ---------------- R10-1 第13章 総括
    p = find_par(doc, "在宅認定者約369人に換算すると約134人と推計され")
    set_text(
        p,
        "家族・親族による介護がなく、かつ介護保険サービスも利用していない方は50件"
        "（両設問に有効回答があった138件の36.2％）確認された。在宅認定者約369人に単純外挿した感度試算では"
        "約134人規模となるが、実人数の確定値ではない。在宅生活継続支援や施設入所相談につながる可能性のある層として、"
        "入所申込みの状況、サービス未利用の理由、家族の支援状況を追加で確認する必要がある"
        "（９－９・９－11参照）。",
    )
    log.append("[R10-1] 第13章総括の「約134人と推計」「最も届いていない層」「最も大きな供給源」を感度試算・追加確認の記述に是正")

    # ---------------- R10-2 第10章③
    p = find_par(doc, "４章のとおり在宅認定者約369人に換算すると申込済み約125人")
    sub_text(
        p,
        "（４章のとおり在宅認定者約369人に換算すると申込済み約125人・検討中約53人）",
        "（４章の感度試算では、在宅認定者約369人に単純外挿した場合、申込済み約125人・検討中約53人規模となる）",
    )
    log.append("[R10-2] 第10章③の換算値に「感度試算」「単純外挿」の語を補った")

    doc.save(str(DOC))
    print("\n".join(log))


if __name__ == "__main__":
    main()
