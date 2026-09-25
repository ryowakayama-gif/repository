# -*- coding: utf-8 -*-
"""成果品（Word文書）の体裁・図・表・文言を機械的に点検する。

　使い方
　　python3 check_docx.py out/計画素案_第10期.docx [ほかの.docx ...] [--figs figs.py figs_soan.py]

　点検する内容
　　図　　キャプションの形式と連番／出典の有無／画像数との一致／keepNextの有無
　　表　　列数の一致／列幅の合計／セル幅とグリッドの一致／見出し行の塗り
　　記号　【　】【案】《要確定》《要回答》の実数と、0-1の集計表との一致
　　参照　《要回答N》《要確定N》の番号が、附の一覧に存在するか／「図N」の参照先の有無
　　文言　表記ゆれ（金ヶ崎／半角％／半角カナ／全角数字の年度）
　　図の素　系列（label付き）に条件付きの色指定がないか／規定パレット外の色がないか

　終了コード　0＝指摘なし　1＝指摘あり
"""
import re
import sys

import docx
from docx.shared import Cm
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

BODY_W = 17.0                       # 本文幅（cm）A4縦・左右余白2.0cm
PALETTE = {'#C00000', '#2E75B6', '#BDD7EE', '#ED7D31', '#A6A6A6', '#1F3864',
           '#F2F2F2', '#D9D9D9', '#E6E6E6', '#E8E8E8', '#9DC3E6', '#1F4E79',
           '#BFBFBF', '#808080', 'white'}
NG_WORDS = [('金ヶ崎', '金ケ崎'), ('%', '％'), ('ｱ', '全角カタカナ'),
            ('ｶ', '全角カタカナ'), ('ｻ', '全角カタカナ')]


def _text(doc):
    """本文と表をつないだ全文。"""
    out = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for r in t.rows:
            out += [c.text for c in r.cells]
    return '\n'.join(out)


def _is_src(p):
    return any(r.font.size is not None and r.font.size.pt <= 8.5 for r in p.runs)


def check_figs(doc, name, ng):
    caps = []
    paras = doc.paragraphs
    for i, p in enumerate(paras):
        m = re.match(r'^【図(\d+)】(.+)$', p.text.strip())
        if not m:
            continue
        caps.append((i, int(m.group(1)), m.group(2)))
        if not p.paragraph_format.keep_with_next:
            ng.append('%s 図%s　キャプションにkeepNextがない（図と別ページに分かれる）'
                      % (name, m.group(1)))
        # 直後＝画像、その次＝出典
        nxt = paras[i + 1] if i + 1 < len(paras) else None
        if nxt is None or not nxt._p.findall('.//' + qn('a:blip')):
            ng.append('%s 図%s　キャプションの直後に画像がない' % (name, m.group(1)))
        src = paras[i + 2] if i + 2 < len(paras) else None
        if src is None or not _is_src(src) or not src.text.strip():
            ng.append('%s 図%s　出典行がない' % (name, m.group(1)))
    nums = [c[1] for c in caps]
    if nums != list(range(1, len(nums) + 1)):
        ng.append('%s 図番号が連番でない　%s' % (name, nums))
    n_img = len(doc.inline_shapes)
    if n_img != len(caps):
        ng.append('%s 画像%d点に対しキャプション%d点' % (name, n_img, len(caps)))
    return caps


def _body(doc):
    """本文と表を、文書に現れる順で返す。"""
    for ch in doc.element.body.iterchildren():
        if ch.tag == qn('w:p'):
            yield Paragraph(ch, doc)
        elif ch.tag == qn('w:tbl'):
            yield Table(ch, doc)


def check_fig_table(doc, name, ng):
    """図は、表・その出典行・他の図の直後に置く（本文段落の直後に置かない）。

    　数値は表で示し、図はその表を読みやすくするために添える、という規則による。
    """
    prev = None          # 'tbl' / 'src' / 'fig' / 'body' / 'head'
    for el in _body(doc):
        if isinstance(el, Table):
            prev = 'tbl'
            continue
        t = el.text.strip()
        if el.style.name.startswith('Heading'):
            prev = 'head'
            continue
        m = re.match(r'^【図(\d+)】(.+)$', t)
        if m:
            if prev not in ('tbl', 'src', 'fig'):
                ng.append('%s 図%s　表の直後に置かれていない　「%s」'
                          % (name, m.group(1), m.group(2)[:30]))
            prev = 'figcap'
            continue
        if prev == 'figcap':
            prev = 'figimg'
            continue
        if prev == 'figimg':
            prev = 'fig'          # 図の出典行
            continue
        if not t:
            continue
        prev = 'src' if _is_src(el) else 'body'


def check_tables(doc, name, ng):
    for ti, t in enumerate(doc.tables, 1):
        ncol = {len(r.cells) for r in t.rows}
        if len(ncol) > 1:
            ng.append('%s 表%d　行ごとに列数が違う　%s' % (name, ti, sorted(ncol)))
            continue
        w = [c.width.cm for c in t.rows[0].cells if c.width is not None]
        if len(w) == len(t.rows[0].cells):
            if sum(w) > BODY_W + 0.05:
                ng.append('%s 表%d　列幅の合計が%.2fcmで本文幅%.1fcmを超える'
                          % (name, ti, sum(w), BODY_W))
            grid = t._tbl.find(qn('w:tblGrid'))
            gc = [int(g.get(qn('w:w'))) for g in grid.findall(qn('w:gridCol'))] \
                if grid is not None else []
            want = [int(Cm(x).twips) for x in w]
            if gc and gc != want:
                ng.append('%s 表%d　セル幅とグリッドが一致しない（Wordと組版結果がずれる）'
                          % (name, ti))
        else:
            ng.append('%s 表%d　列幅が指定されていない' % (name, ti))
        sh = t.rows[0].cells[0]._tc.find('.//' + qn('w:shd'))
        if sh is None or (sh.get(qn('w:fill')) or '').upper() != '1F3864':
            ng.append('%s 表%d　見出し行の塗りが規定色（1F3864）でない' % (name, ti))


def check_marks(doc, name, ng):
    t = _text(doc)
    cnt = {'【　】': t.count('【　】'), '【案】': t.count('【案】')}
    kaku = sorted({int(x) for x in re.findall(r'《要確定(\d+)》', t)})
    kai = sorted({int(x) for x in re.findall(r'《要回答(\d+)》', t)})
    # 0-1 の集計表と突き合わせる（表から直接読む）
    for tb in doc.tables:
        for r in tb.rows:
            c0 = r.cells[0].text.strip().replace('*', '')
            if c0 in ('【　】', '【案】'):
                m = re.search(r'(\d+)\s*か所', r.cells[-1].text)
                if m and int(m.group(1)) != cnt[c0]:
                    ng.append('%s 0-1の集計「%s %s か所」に対し、文書中の実数は%d か所'
                              % (name, c0, m.group(1), cnt[c0]))
            if c0 in ('《要確定》', '《要回答》'):
                m = re.search(r'(\d+)\s*件', r.cells[-1].text)
                nn = kaku if c0 == '《要確定》' else kai
                if m and nn and max(nn) > int(m.group(1)):
                    ng.append('%s 0-1の集計「%s %s 件」に対し、本文は%d番を参照している'
                              % (name, c0, m.group(1), max(nn)))
    for pat, nums, label in ((r'町のご判断をいただく事項（(\d+)件）', kaku, '要確定'),
                             (r'確認事項（(\d+)件）', kai, '要回答')):
        m = re.search(pat, t)
        if not m or not nums:
            continue
        # 一覧の見出しの件数と、本文で参照されている最大番号を照らす
        if max(nums) > int(m.group(1)):
            ng.append('%s 《%s》の一覧は%s件だが、本文は%s番を参照している'
                      % (name, label, m.group(1), max(nums)))
        missing = [n for n in range(1, max(nums) + 1) if n not in nums]
        if missing and label == '要確定' and len(nums) >= 3:
            ng.append('%s 《%s%s》が本文から参照されていない'
                      % (name, label, '・'.join(str(x) for x in missing)))


def _num(x):
    x = x.replace('*', '').replace(',', '').replace('人', '').replace('件', '')
    x = x.replace('千円', '').replace('円', '').replace('点', '').replace('か所', '')
    x = x.replace('▲', '-').replace('△', '-').strip()
    try:
        return float(x)
    except ValueError:
        return None


def check_sums(doc, name, ng):
    """合計行のある表の検算。列ごとに、合計以外の行の和と突き合わせる。"""
    for ti, t in enumerate(doc.tables, 1):
        rows = [[c.text.strip() for c in r.cells] for r in t.rows]
        if len(rows) < 3:
            continue
        tot_i = [i for i, r in enumerate(rows)
                 if re.fullmatch(r'\**\s*(合?計)\s*\**', r[0].replace('　', ''))]
        if not tot_i:
            continue
        ti_row = tot_i[-1]
        body = rows[1:ti_row]      # 合計行より上の行だけを足す
        for ci in range(1, len(rows[0])):
            tv = _num(rows[ti_row][ci])
            if tv is None or tv == 0:
                continue
            col = [_num(r[ci]) for r in body]
            if any(v is None for v in col) or len(col) < 2:
                continue
            if any('％' in r[ci] or '%' in r[ci] or '倍' in r[ci] for r in body):
                continue
            ssum = sum(col)
            if abs(ssum - tv) > max(1.0, abs(tv) * 0.005):
                ng.append('%s 表%d 第%d列　合計%s に対し、各行の和は%s'
                          % (name, ti, ci + 1, ('%g' % tv), ('%g' % ssum)))


def check_words(doc, name, ng):
    t = _text(doc)
    for w, right in NG_WORDS:
        if w in t:
            n = t.count(w)
            ex = [l.strip()[:34] for l in t.split('\n') if w in l][:2]
            ng.append('%s 表記ゆれ「%s」が%d件（正：%s）　例：%s'
                      % (name, w, n, right, ' ／ '.join(ex)))


def check_refs(doc, name, caps, ng):
    t = _text(doc)
    have = {c[1] for c in caps}
    for n in sorted({int(x) for x in re.findall(r'（?図(\d+)）?', t)}):
        if have and n not in have:
            ng.append('%s 本文が「図%d」を参照しているが、その図がない' % (name, n))


def check_src(paths, ng):
    """図の生成スクリプトを静的に点検する。"""
    for p in paths:
        s = open(p, encoding='utf-8').read()
        for m in re.finditer(r'\.(?:bar|barh|plot)\((?:[^()]|\([^()]*\))*\)', s):
            call = m.group(0)
            if 'label=' in call and re.search(r'color=\[[^\]]*\bif\b', call):
                ln = s[:m.start()].count('\n') + 1
                ng.append('%s:%d　凡例のある系列を条件で塗り分けている'
                          '（1つの図の中で色は1つの意味しか持たせない）' % (p, ln))
        used = set(re.findall(r"'#[0-9A-Fa-f]{6}'", s))
        for c in used:
            if c.strip("'").upper() not in {x.upper() for x in PALETTE}:
                ng.append('%s　規定パレット外の色 %s' % (p, c))


def main(argv):
    docs, figs, mode = [], [], 'doc'
    for a in argv:
        if a == '--figs':
            mode = 'figs'
        elif mode == 'doc':
            docs.append(a)
        else:
            figs.append(a)
    ng = []
    for path in docs:
        d = docx.Document(path)
        name = path.split('/')[-1].replace('.docx', '')
        caps = check_figs(d, name, ng)
        check_tables(d, name, ng)
        check_fig_table(d, name, ng)
        check_marks(d, name, ng)
        check_sums(d, name, ng)
        check_words(d, name, ng)
        check_refs(d, name, caps, ng)
        print('点検　%s　段落%d　表%d　図%d'
              % (name, len(d.paragraphs), len(d.tables), len(d.inline_shapes)))
    if figs:
        check_src(figs, ng)
        print('点検　図の生成スクリプト　%s' % ' '.join(figs))
    print()
    if not ng:
        print('指摘なし')
        return 0
    print('指摘　%d件' % len(ng))
    for x in ng:
        print('  ・' + x)
    return 1


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
