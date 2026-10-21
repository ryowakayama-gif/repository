# -*- coding: utf-8 -*-
"""第1回策定委員会資料の様式の部品（fmt.py の後に連結する）。

　COVER　表紙（委員会の回・計画名3行・資料名の帯・目次・日付と会場）
　SEC　　節の見出し（緑の帯・白文字）。2つ目以降は改ページしてから置く
　SQ　　 小見出し（■）
　BL　　 箇条（・で始め、ぶら下げ）
　BOX　　囲み。kind='o' はご意見をいただきたい点、kind='g' は読み取り・考え方

　表は fmt.py の TBL をそのまま用いる（見出し行の塗りは STYLE の TBLFILL）。
"""
G = '2C6E52'
C_TXT = _rgb('1B3A2F')
_BOXK = {'o': ('FBF0E6', 'EBC9AB', 'B85C22'), 'g': ('EAF2ED', 'CBD8D1', G)}

_f0 = _f


def _f(run, f=FONT, sz=BODY_PT, b=False, color=None):
    """文字色の既定を第1回資料の #1B3A2F にする。"""
    _f0(run, f, sz, b, color or C_TXT)


def _shd(p, fill):
    pPr = p._p.get_or_add_pPr()
    sh = OxmlElement('w:shd')
    sh.set(qn('w:val'), 'clear'); sh.set(qn('w:color'), 'auto'); sh.set(qn('w:fill'), fill)
    pPr.append(sh)


def _pbdr(p, color, sides=('top', 'left', 'bottom', 'right'), sz='4', space='4'):
    pPr = p._p.get_or_add_pPr()
    b = OxmlElement('w:pBdr')
    for s in sides:
        e = OxmlElement('w:' + s)
        e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), sz)
        e.set(qn('w:space'), space); e.set(qn('w:color'), color)
        b.append(e)
    pPr.append(b)


def _runs(p, text, sz, color=None, bold=False):
    for seg in re.split(r'(\*\*.*?\*\*)', text):
        if not seg:
            continue
        bb = seg.startswith('**') and seg.endswith('**')
        _f(p.add_run(seg[2:-2] if bb else seg), FONT, sz, bold or bb, color)


_NSEC = [0]


def SEC(t):
    """節の見出し。「１　　題 ― 結論」。"""
    if _NSEC[0]:
        doc.add_page_break()
    _NSEC[0] += 1
    p = doc.add_paragraph(); pf = p.paragraph_format
    pf.space_before = Pt(0); pf.space_after = Pt(8)
    pf.left_indent = Cm(0.2); pf.right_indent = Cm(0.2)
    pf.keep_with_next = True
    _shd(p, G)
    _f(p.add_run(t), GFONT, 17, True, _rgb('FFFFFF'))
    return _register(p, t, 1)


def SQ(t):
    """小見出し（■）。"""
    p = doc.add_paragraph(); pf = p.paragraph_format
    pf.space_before = Pt(8); pf.space_after = Pt(4)
    pf.keep_with_next = True
    _f(p.add_run('■　' + t), GFONT, 13.5, True, _rgb(G))
    return _register(p, t, 2)


def BL(t, sz=None):
    """箇条。「・」で始め、2行目以降を字下げする。"""
    sz = BODY_PT if sz is None else sz
    p = doc.add_paragraph(); pf = p.paragraph_format
    pf.space_before = Pt(0); pf.space_after = Pt(2)
    pf.left_indent = Cm(0.45); pf.first_line_indent = Cm(-0.45)
    _ls(pf)
    _runs(p, '・' + t, sz)
    return p


def BOX(title, lines, kind='g'):
    """囲み。題と本文の行を、塗りと枠のある段落の連なりとして置く。"""
    fill, line, tc = _BOXK[kind]
    ps = []
    p = doc.add_paragraph(); ps.append(p)
    _runs(p, title, BODY_PT - 0.5, _rgb(tc), True)
    for s in lines:
        q = doc.add_paragraph(); ps.append(q)
        _runs(q, s, BODY_PT - 1)
    for i, q in enumerate(ps):
        pf = q.paragraph_format
        pf.space_before = Pt(4 if i == 0 else 0)
        pf.space_after = Pt(6 if i == len(ps) - 1 else 2)
        pf.left_indent = Cm(0.25); pf.right_indent = Cm(0.25)
        pf.keep_with_next = i < len(ps) - 1
        pf.keep_together = True
        _shd(q, fill)
        _pbdr(q, line)
    return ps


def COVER(no, sub, toc, date_place, info):
    """表紙。no＝「第2回」、sub＝資料名の下に置く行、toc＝[(番号, 題)]、info＝[(項目, 内容)]。"""
    C = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph(); p.alignment = C
    p.paragraph_format.space_before = Pt(36); p.paragraph_format.space_after = Pt(20)
    _pbdr(p, '000000', ('top', 'bottom'), '24', '1')
    _f(p.add_run(no + '　策定委員会'), GFONT, 16, False, _rgb(G))
    for t, sz in (('金ケ崎町高齢者福祉計画', 13), ('第10期金ケ崎町介護保険事業計画', 24),
                  ('金ケ崎町認知症施策推進計画', 24)):
        q = doc.add_paragraph(); q.alignment = C
        q.paragraph_format.space_after = Pt(6)
        _f(q.add_run(t), GFONT, sz, sz > 13)
    q = doc.add_paragraph(); q.alignment = C
    q.paragraph_format.space_before = Pt(18); q.paragraph_format.space_after = Pt(6)
    _shd(q, G)
    _f(q.add_run(' '.join(no + '策定委員会資料')), GFONT, 20, True, _rgb('FFFFFF'))
    for s in sub:
        q = doc.add_paragraph(); q.alignment = C
        q.paragraph_format.space_after = Pt(2)
        _runs(q, s, 12.5, _rgb(G), True)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    for k, v in info:
        q = doc.add_paragraph()
        q.paragraph_format.space_after = Pt(0)
        q.paragraph_format.left_indent = Cm(0.25); q.paragraph_format.right_indent = Cm(0.25)
        _shd(q, 'EAF2ED'); _pbdr(q, 'CBD8D1')
        _f(q.add_run(k + '　'), FONT, 11, True, _rgb(G))
        _runs(q, v, 11)
    q = doc.add_paragraph(); q.alignment = C
    q.paragraph_format.space_before = Pt(18); q.paragraph_format.space_after = Pt(6)
    _f(q.add_run('目　　次'), GFONT, 13, True, _rgb(G))
    for n, t in toc:
        q = doc.add_paragraph()
        q.paragraph_format.space_after = Pt(1)
        q.paragraph_format.left_indent = Cm(2.0)
        _f(q.add_run('%s　' % n), FONT, 11, True, _rgb(G))
        _f(q.add_run(t), FONT, 11)
    q = doc.add_paragraph(); q.alignment = C
    q.paragraph_format.space_before = Pt(24)
    _runs(q, date_place, 12)
    _NSEC[0] = 1          # 表紙の次の節から改ページする
