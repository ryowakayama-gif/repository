# -*- coding: utf-8 -*-
"""WBSと進捗状況（Word）の作成。"""
import os, sys
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from wbs_data import WBS, STALLED, RECOVERY, SPEC, TODAY, DEADLINE

HERE = os.path.dirname(os.path.abspath(__file__))
NAVY, BLUE, RED, GREY = RGBColor(0x1F, 0x38, 0x64), RGBColor(0x2E, 0x75, 0xB6), \
                        RGBColor(0xC0, 0, 0), RGBColor(0x59, 0x59, 0x59)
ORANGE, GREEN = RGBColor(0xC5, 0x5A, 0x11), RGBColor(0x37, 0x56, 0x23)
JP, JPG = 'ＭＳ 明朝', 'ＭＳ ゴシック'

doc = Document()
st = doc.styles['Normal']
st.font.name = JP
st.font.size = Pt(10.5)
st.element.rPr.rFonts.set(qn('w:eastAsia'), JP)
for s in doc.sections:
    s.orientation = WD_ORIENT.LANDSCAPE
    s.page_width, s.page_height = Cm(29.7), Cm(21.0)
    s.top_margin = s.bottom_margin = Cm(1.5)
    s.left_margin = s.right_margin = Cm(1.5)


def setfont(run, name=JP, size=None, bold=None, color=None):
    run.font.name = name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color:
        run.font.color.rgb = color


def P(text='', size=10.5, bold=False, color=None, align=None, space_after=6, font=JP):
    p = doc.add_paragraph()
    if align:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    if text:
        setfont(p.add_run(text), font, size, bold, color)
    return p


def H1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(7)
    setfont(p.add_run(text), JPG, 13.5, True, NAVY)
    pbd = OxmlElement('w:pBdr')
    b = OxmlElement('w:bottom')
    b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), '12')
    b.set(qn('w:space'), '2'); b.set(qn('w:color'), '1F3864')
    pbd.append(b)
    p._p.get_or_add_pPr().append(pbd)
    return p


def H2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    setfont(p.add_run(text), JPG, 11.5, True, BLUE)
    return p


def bullet(text, size=10):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.12
    setfont(p.add_run(text), JP, size)
    return p


def shade(cell, hexcolor):
    e = OxmlElement('w:shd')
    e.set(qn('w:val'), 'clear')
    e.set(qn('w:fill'), hexcolor)
    cell._tc.get_or_add_tcPr().append(e)


STATE_BG = {'完了': 'E2EFDA', '進行中': 'FFF2CC', '未着手': 'F2F2F2', '町回答待ち': 'FCE4E4'}
STATE_FG = {'完了': GREEN, '進行中': ORANGE, '未着手': GREY, '町回答待ち': RED}


def table(headers, rows, widths=None, size=8.5, note=None, center=None, state_col=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    if widths:
        t.autofit = False
        for i, w in enumerate(widths):
            for row in t.rows:
                row.cells[i].width = Cm(w)
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = ''
        pr = c.paragraphs[0]
        pr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pr.paragraph_format.space_after = Pt(2)
        setfont(pr.add_run(str(h)), JPG, size, True, RGBColor(0xFF, 0xFF, 0xFF))
        shade(c, '2E75B6')
    for r in rows:
        cells = t.add_row().cells
        head = str(r[0]).startswith('【')
        for i, v in enumerate(r):
            cells[i].text = ''
            pr = cells[i].paragraphs[0]
            pr.paragraph_format.space_after = Pt(2)
            txt = str(v)
            pr.alignment = (WD_ALIGN_PARAGRAPH.CENTER
                            if (center and i in center) else WD_ALIGN_PARAGRAPH.LEFT)
            col = None
            bold = head or txt.startswith('★')
            if txt.startswith('★'):
                col = RED
            if state_col is not None and i == state_col:
                shade(cells[i], STATE_BG.get(txt, 'FFFFFF'))
                col = STATE_FG.get(txt)
                bold = (txt == '町回答待ち')
            setfont(pr.add_run(txt.lstrip('★')), JP, size, bold, col)
            if head:
                shade(cells[i], 'DDEBF7')
            if widths:
                cells[i].width = Cm(widths[i])
        if widths:
            for i, w in enumerate(widths):
                cells[i].width = Cm(w)
    if note:
        P(note, size=8, color=GREY, space_after=8)
    else:
        P('', size=4, space_after=4)
    return t


def pagebreak():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


# ============================== 表紙相当 ==============================
P('令和8年度　上下水道経営指標評価書作成等業務委託（女川町）', size=10.5,
  align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=2, font=JPG)
P('ビズアップ公共コンサルティング株式会社', size=10.5,
  align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=10)
H1('業務分解構成（WBS）と進捗状況')
P(f'作成日：{TODAY}　／　履行期限：{DEADLINE}　／　中間報告：令和8年9月末（提出済）',
  size=10, color=GREY, space_after=10)

# ============================== 1 全体 ==============================
H1('1　全体の状況')
n_total = sum(1 for x in WBS if x[1] == 1)
n_done = sum(1 for x in WBS if x[1] == 1 and x[7] == '完了')
n_prog = sum(1 for x in WBS if x[1] == 1 and x[7] == '進行中')
n_none = sum(1 for x in WBS if x[1] == 1 and x[7] == '未着手')
n_wait = sum(1 for x in WBS if x[1] == 1 and x[7] == '町回答待ち')
avg = sum(x[8] for x in WBS if x[1] == 1) / n_total

P('中間報告までの作業は完了しているが、業務全体でみると残りの比重が大きい。'
  '特に、業務名にある下水道の経営指標評価が手つかずであること、'
  '評価書の本体となる指標分析が要約にとどまっていることの2点が、履行期限に対する主な懸念である。')
table(['状態', '件数', '内容'],
 [['完了', f'{n_done}件', '中間報告書、料金体系の比較、前提条件切替計算表、パッケージ整備など'],
  ['進行中', f'{n_prog}件', '総括原価の各費目、3パターンの試算、影響分析'],
  ['未着手', f'{n_none}件', '下水道の指標評価、経営指標評価書の取りまとめ、説明会資料、年次資金収支'],
  ['★町回答待ち', f'★{n_wait}件', '★決算確定値、固定資産台帳残高、繰出基準の整理、使用者別調定明細'],
  ['【合計】', f'{n_total}件', f'当方作業の進捗率（単純平均）　約{avg:.0f}％']],
 widths=[3.4, 2.4, 20.8], size=9.5, center={1}, state_col=0)

H2('1-1　期限までの残り')
table(['時期', '節目', '状態'],
 [['令和8年9月末', '中間報告', '完了'],
  ['令和8年10月', '未回答データの再依頼・下水道データの依頼・仕様書原本の受領', '未着手'],
  ['令和8年11月', '総括原価の確定、経営指標評価書の草案', '未着手'],
  ['令和8年12月', '3パターンの数値確定、年次資金収支', '未着手'],
  ['令和9年1月', '改定案の決定、住民・事業者説明', '未着手'],
  ['令和9年2月末', '最終報告書の提出', '未着手'],
  ['★令和9年3月31日', '★履行期限', '★―']],
 widths=[4.0, 16.0, 6.6], size=9.5, center={2})
pagebreak()

# ============================== 2 WBS ==============================
H1('2　業務分解構成（WBS）')
P('担当欄の「町」は女川町、「当方」はビズアップ公共コンサルティング、「協議」は両者。'
  '進捗率は当方の作業量に対する割合。', size=9, color=GREY, space_after=6)
rows = []
for num, lvl, name, deliv, owner, start, end, state, pct, note in WBS:
    if lvl == 0:
        rows.append([f'【{num}】', f'【{name}】', '', '', '', '', ''])
        continue
    rows.append([num, name, deliv, owner, f'{start}〜{end}', state,
                 f'{pct}％' + (('　' + note) if note else '')])
table(['WBS', '作業項目', '成果物', '担当', '期間', '状態', '進捗率・備考'], rows,
 widths=[1.5, 6.4, 3.4, 1.5, 2.6, 2.2, 9.0], size=8, center={0, 3, 4, 5}, state_col=5,
 note='※ 備考の★印は工程上の要注意事項。')
pagebreak()

# ============================== 3 停滞事項 ==============================
H1('3　進捗が滞っている事項')
P('本業務の進行を妨げている事項を、放置期間の長いものから整理した。'
  'S1・S2は当方の着手判断で動かせるもの、S3・S4・S7は女川町からの回答がないと動かないもの、'
  'S5・S6・S8は当方の作業で解消できるものである。')
rows = []
for no, kind, what, since, dur, impact, how, due in STALLED:
    rows.append([no, kind, what, f'{since}\n（{dur}）', impact, how, due])
table(['No', '区分', '滞っている事項', '起算日', '影響', '解消の手立て', '期限'], rows,
 widths=[1.2, 2.2, 5.6, 2.2, 6.8, 6.8, 1.8], size=7.5, center={0, 1, 3, 6}, state_col=1)
pagebreak()

# ============================== 4 リカバリ ==============================
H1('4　優先して着手する作業')
P('10月上旬に片づける5件（順位1〜5）は、いずれも所要半日〜1日でありながら、'
  '着手しないと後工程が1か月単位で遅れる。まずこの5件を先に処理することを提案する。')
table(['順', '作業', '所要', '期限', '効果'],
 [[p, t, c, d, e] for p, t, c, d, e in RECOVERY],
 widths=[1.4, 8.6, 2.0, 2.6, 12.0], size=9, center={0, 2, 3})

H2('4-1　女川町への依頼をまとめた場合の内容')
P('順位1・2・5は1回の連絡にまとめられる。次の6点を、様式と期限を添えて依頼する。')
for t in ['令和6年度・令和7年度決算の確定値（損益計算書・貸借対照表）',
          '令和8年4月1日時点の固定資産台帳残高（取水・浄水・配水の区分別）',
          '使用者別の調定明細（使用者番号は匿名可。業種・口径・用途・月別水量・調定額・減免の別）',
          '繰出基準に基づく一般会計繰入金の整理（基準内・基準外の区分と算定根拠）',
          '建設改良費調査票の未記入5項目（借入予定利率、償却年数の区分、江島海底送水管の資産種別、'
          '構造物及び設備の差4.75億円の内容、量水器更新費の計上方針）',
          '下水道の段階量調定集計表（5年分）、現行使用料条例、総括原価の予測、収入の予測',
          '業務仕様書の写し']: bullet(t)
P('あわせて、南三陸町上下水道事業所へ20mm以上の口径別料金を照会する（当方で実施、所要1日）。',
  space_after=8)
pagebreak()

# ============================== 5 仕様書対応 ==============================
H1('5　業務仕様書の要求事項と成果物の対応')
P('仕様書原本は未入手のため、令和8年6月2日作成の仕様書整合確認の記載に基づいている。'
  '原本の受領後に本表を更新する。', color=RED, size=10, space_after=6)
table(['要求事項', '求められる内容', '現在の成果物', '達成度', '残作業'],
 [[a, b, c, d, e] for a, b, c, d, e in SPEC],
 widths=[3.2, 6.4, 5.4, 1.8, 9.8], size=8.5, center={3},
 note='※ 達成度は当方の判断による目安。')

H2('5-1　特に注意を要する点')
for t in ['業務名は「上下水道経営指標評価書作成等業務委託」であり、'
          '料金改定の検討はその一部である。現時点の成果物は水道の料金改定に偏っている',
          '評価書の本体となる財務・経営指標分析は、修正前後決算の対比と'
          '経営比較分析表の指標を体系的にそろえる必要がある。'
          '現在は中間報告書第3章に主要4指標の要約があるにとどまる',
          '資産維持率3％について、仕様書がこの水準を要求しているかは原本で確認する。'
          '現在は日本水道協会の算定要領の標準値として採用している',
          'パターン2の「現状赤字額」がどの定義を指すかも、原本と町の確認による']: bullet(t)

P('', space_after=14)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
setfont(p.add_run('以　上'), JP, 10.5)

cp = doc.core_properties
cp.title = '女川町 上下水道経営指標評価書作成等業務委託　WBSと進捗状況'
cp.subject = '令和8年度上下水道経営指標評価書作成等業務委託'
cp.author = '若山諒太'
cp.last_modified_by = '若山諒太'
cp.comments = 'ビズアップ公共コンサルティング株式会社'

OUT = os.path.join(HERE, '..', '女川町上下水道経営指標評価_WBSと進捗状況.docx')
doc.save(OUT)
print('作成:', os.path.abspath(OUT))
