# -*- coding: utf-8 -*-
"""docx の保存前に OOXML の不適合を直す。

`/mnt/skills/public/docx/scripts/office/validate.py` に適合させるための手当て。
生の要素を足すと順序が崩れるため、保存の直前に本モジュールの `fix(doc)` を通す。

直すもの
- `w:zoom` の必須属性 `w:percent` の欠落（既定のテンプレートが欠いている）。
- `w:tblPr` 及び `w:tcPr` の子要素の順序。
  OOXML は順序を定めており、`w:tblCellMar` は `w:tblLook` より前、
  `w:shd` は `w:tcPr` では `w:vAlign` より前でなければならない
  （CLAUDE.md §4）。

使い方
    import docx_fix
    docx_fix.fix(doc)
    doc.save(OUT)
"""

from docx.oxml.ns import qn

# CT_TblPrBase の順序
TBLPR = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
         "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
         "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout",
         "tblCellMar", "tblLook", "tblCaption", "tblDescription"]

# CT_TcPrBase の順序
TCPR = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders",
        "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign",
        "hideMark"]

# CT_TrPrBase の順序
TRPR = ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter",
        "cantSplit", "trHeight", "tblHeader", "tblCellSpacing", "jc",
        "hidden"]


def _sort(el, order):
    """el の子要素を order の並びに直す。order にないものは末尾に残す。"""
    if el is None:
        return
    idx = {qn("w:" + n): i for i, n in enumerate(order)}
    kids = list(el)
    if all(idx.get(k.tag, len(order)) <= idx.get(n.tag, len(order))
           for k, n in zip(kids, kids[1:])):
        return                      # 既に順序どおり
    for k in sorted(kids, key=lambda c: idx.get(c.tag, len(order))):
        el.append(k)


def fix(doc):
    """保存前の docx を直す。直した箇所の数を返す。"""
    n = 0
    z = doc.settings.element.find(qn("w:zoom"))
    if z is not None and z.get(qn("w:percent")) is None:
        z.set(qn("w:percent"), "100")
        n += 1
    body = doc.element.body
    for tag, order in (("w:tblPr", TBLPR), ("w:tcPr", TCPR),
                       ("w:trPr", TRPR)):
        for el in body.iter(qn(tag)):
            before = [c.tag for c in el]
            _sort(el, order)
            if [c.tag for c in el] != before:
                n += 1
    return n
