# -*- coding: utf-8 -*-
"""xlsx を OOXML のまま編集し、数式を評価してキャッシュを更新するための共通処理。

openpyxl で開き直すとチャートシート・グラフ・図が失われるため、XMLを直接編集する。
LibreOffice がこの環境で動作しない（無修正の原本でも読み込めない）ため、再計算は
この中の簡易評価器で行う。fix_xlsx_kyogikai.py で使った処理を切り出したもの。
"""
import zipfile, re, math

# ── 列記号 ───────────────────────────────────────────────────────────
def col_n(c):
    v = 0
    for ch in c:
        v = v * 26 + (ord(ch) - 64)
    return v

def col_s(v):
    out = ''
    while v:
        v, r = divmod(v - 1, 26)
        out = chr(65 + r) + out
    return out

def colseq(a, b):
    return [col_s(i) for i in range(col_n(a), col_n(b) + 1)]

def split_ref(r):
    m = re.fullmatch(r'(\$?)([A-Z]{1,3})(\$?)(\d+)', r)
    return m.group(1), m.group(2), m.group(3), int(m.group(4))

# ── 数式の字句 ───────────────────────────────────────────────────────
CELL = r'\$?[A-Z]{1,3}\$?\d{1,7}'
PFX  = r"(?:'([^']*)'|([^\s!,()+\-*/:&%'\"=<>]+))!"
#  範囲と単一参照は1回で置換する（範囲を先に変換すると、生成した
#  RNG('sheet','E17','E23') の中身を単一参照の置換が二重に書き換えてしまう）
RE_REF    = re.compile(f'(?:{PFX})?({CELL})(?::({CELL}))?')
RE_PCT    = re.compile(r'(\d+(?:\.\d+)?)%')
RE_QUOTED = re.compile(r"'[^']*'")

class Unresolved(Exception):
    pass

def translate(text, src, dst):
    """共有数式の親(src)から子(dst)へ相対参照をずらす。"""
    _, sc, _, sr = split_ref(src)
    _, dc, _, dr = split_ref(dst)
    dcol, drow = col_n(dc) - col_n(sc), dr - sr
    masks = []
    t = RE_QUOTED.sub(lambda m: (masks.append(m.group(0)), f'\x01{len(masks)-1}\x01')[1], text)
    def shift(m):
        a, c, b, r = split_ref(m.group(0))
        return f'{a}{c if a else col_s(col_n(c) + dcol)}{b}{r if b else r + drow}'
    t = re.compile(CELL).sub(shift, t)
    for i, v in enumerate(masks):
        t = t.replace(f'\x01{i}\x01', v)
    return t

def _flat(a):
    for v in a:
        if isinstance(v, list):
            yield from _flat(v)
        else:
            yield v
def _round(x, n=0):
    m = 10.0 ** n
    return math.floor(abs(x) * m + 0.5) / m * (1 if x >= 0 else -1)
def _roundup(x, n=0):
    m = 10.0 ** n
    return math.ceil(abs(x) * m) / m * (1 if x >= 0 else -1)
def _rounddown(x, n=0):
    m = 10.0 ** n
    return math.floor(abs(x) * m) / m * (1 if x >= 0 else -1)
def _sum(*a):     return sum(_flat(a))
def _avg(*a):
    v = list(_flat(a))
    if not v:
        raise Unresolved('AVERAGEの対象が空')
    return sum(v) / len(v)
def _max(*a):     return max(_flat(a))
def _min(*a):     return min(_flat(a))
def _if(c, a, b): return a if c else b


class Workbook:
    """xlsx を展開して保持し、セル単位の編集と再計算を行う。"""

    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        self.parts = {n: self.z.read(n) for n in self.z.namelist()}
        self.infos = self.z.infolist()
        self.log = []
        wb = self.parts['xl/workbook.xml'].decode()
        rl = self.parts['xl/_rels/workbook.xml.rels'].decode()
        rid2tgt = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]*)"', rl))
        self.SHEET = {m.group(1): 'xl/' + rid2tgt[m.group(2)].lstrip('/')
                      for m in re.finditer(r'<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"', wb)}
        self._ss = self.parts['xl/sharedStrings.xml'].decode()
        self._si = [''.join(re.findall(r'<t[^>]*>([^<]*)</t>', s.split('<rPh')[0]))
                    for s in re.findall(r'<si>(.*?)</si>', self._ss, re.S)]
        self._added = 0

    # ── パート入出力 ──────────────────────────────────────────────
    def get(self, sheet):      return self.parts[self.SHEET[sheet]].decode('utf-8')
    def put(self, sheet, x):   self.parts[self.SHEET[sheet]] = x.encode('utf-8')
    @staticmethod
    def q(name):               return f"'{name}'" if re.search(r"[ ()%\-]", name) else name

    # ── 共有文字列 ───────────────────────────────────────────────
    def sstr(self, text):
        """共有文字列に追加（既存なら再利用）し、インデックスを返す。"""
        if text in self._si:
            return self._si.index(text)
        self._si.append(text)
        self._added += 1
        return len(self._si) - 1

    def _flush_ss(self):
        if not self._added:
            return
        esc = lambda t: t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        extra = ''.join(f'<si><t xml:space="preserve">{esc(t)}</t></si>'
                        for t in self._si[len(self._si) - self._added:])
        cnt = int(re.search(r'count="(\d+)"', self._ss).group(1))
        ss = re.sub(r'count="\d+" uniqueCount="\d+"',
                    f'count="{cnt + self._added}" uniqueCount="{len(self._si)}"',
                    self._ss, count=1).replace('</sst>', extra + '</sst>')
        self.parts['xl/sharedStrings.xml'] = ss.encode('utf-8')

    # ── セル編集 ─────────────────────────────────────────────────
    @staticmethod
    def cellm(x, ref):
        return re.search(r'<c r="%s"([^>]*?)(?:/>|>(.*?)</c>)' % ref, x, re.S)

    def cached(self, x, ref):
        m = self.cellm(x, ref)
        if not m:
            return None
        v = re.search(r'<v>([^<]*)</v>', m.group(2) or '')
        try:
            return float(v.group(1)) if v else None
        except ValueError:
            return None

    def set_cell(self, x, ref, inner, why, t=None):
        m = self.cellm(x, ref)
        if not m:
            raise SystemExit(f'!! セル {ref} が見つかりません')
        attrs = re.sub(r'\s+t="[^"]*"', '', m.group(1))
        if t:
            attrs += f' t="{t}"'
        self.log.append((ref, m.group(2) or '', inner, why))
        return x[:m.start()] + f'<c r="{ref}"{attrs}>{inner}</c>' + x[m.end():]

    def set_str(self, x, ref, text, why):
        return self.set_cell(x, ref, f'<v>{self.sstr(text)}</v>', why, t='s')

    def clear(self, x, ref, why):
        return self.set_cell(x, ref, '', why)

    # ── 保存 ─────────────────────────────────────────────────────
    def save(self, dst):
        self._flush_ss()
        wb = self.parts['xl/workbook.xml'].decode('utf-8')
        if 'fullCalcOnLoad' not in wb:
            wb = re.sub(r'<calcPr([^>]*?)/>', r'<calcPr\1 fullCalcOnLoad="1"/>', wb, count=1)
            self.parts['xl/workbook.xml'] = wb.encode('utf-8')
        self.parts['[Content_Types].xml'] = re.sub(
            r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', '',
            self.parts['[Content_Types].xml'].decode('utf-8')).encode('utf-8')
        self.parts['xl/_rels/workbook.xml.rels'] = re.sub(
            r'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', '',
            self.parts['xl/_rels/workbook.xml.rels'].decode('utf-8')).encode('utf-8')
        self.parts.pop('xl/calcChain.xml', None)
        zo = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
        done = set()
        for it in self.infos:
            if it.filename in self.parts:
                zo.writestr(it, self.parts[it.filename]); done.add(it.filename)
        for n, b in self.parts.items():
            if n not in done:
                zo.writestr(n, b)
        zo.close(); self.z.close()

    # ── 再計算 ───────────────────────────────────────────────────
    def recalc(self, sheets=None, verbose=True):
        bk = Book(self)
        total, fixed_err = 0, []
        for sh in (sheets or list(self.SHEET)):
            refs = [r for (s, r) in bk.cell if s == sh]
            if not refs:
                continue
            x = self.get(sh)
            chg = unres = 0
            for ref in sorted(refs, key=lambda r: (int(re.search(r'\d+', r).group(0)),
                                                   col_n(re.match(r'[A-Z]+', r).group(0)))):
                rec = bk.cell[(sh, ref)]
                if not rec['f'] or rec['t'] not in (None, 'n', 'e'):
                    continue
                try:
                    got = bk.eval(sh, ref)
                except Unresolved:
                    unres += 1
                    continue
                old = None
                if rec['v'] not in (None, '') and rec['t'] != 'e':
                    try:
                        old = float(rec['v'])
                    except ValueError:
                        old = None
                if old is None or abs(got - old) > max(1e-9, abs(old) * 1e-12):
                    x = _update_v(x, ref, _numfmt(got), drop_type=(rec['t'] == 'e'))
                    chg += 1
                    if rec['t'] == 'e':
                        fixed_err.append(f'{sh}!{ref}')
            if chg:
                self.put(sh, x)
                total += chg
                if verbose:
                    print(f'   {sh}: 更新{chg}セル（評価不能{unres}セル＝外部ブック参照・#REF!）')
        return total, fixed_err, bk


def _numfmt(v):
    if abs(v - round(v)) < 1e-12 and abs(v) < 1e15:
        return str(int(round(v)))
    return repr(float(v))

def _update_v(x, ref, val, drop_type=False):
    m = Workbook.cellm(x, ref)
    inner = m.group(2) or ''
    new = (re.sub(r'<v>.*?</v>', f'<v>{val}</v>', inner, count=1, flags=re.S)
           if '<v>' in inner else inner + f'<v>{val}</v>')
    attrs = re.sub(r'\s+t="e"', '', m.group(1)) if drop_type else m.group(1)
    return x[:m.start()] + f'<c r="{ref}"{attrs}>{new}</c>' + x[m.end():]


class Book:
    """ブック内の全セルを読み、数式を評価する。解決できないものは例外にする。"""

    def __init__(self, wbk):
        self.wbk = wbk
        self.cell = {}
        self.memo = {}
        for name, pn in wbk.SHEET.items():
            x = wbk.parts[pn].decode('utf-8')
            shared, cells = {}, []
            for m in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*?)(?:/>|>(.*?)</c>)', x, re.S):
                ref, attrs, inner = m.group(1), m.group(2), m.group(3) or ''
                t = re.search(r'\st="([^"]*)"', attrs)
                fm = re.search(r'<f([^>]*?)(?:/>|>(.*?)</f>)', inner, re.S)
                vm = re.search(r'<v>(.*?)</v>', inner, re.S)
                rec = {'f': None, 'v': vm.group(1) if vm else None,
                       't': t.group(1) if t else None, 'si': None}
                if fm:
                    fa, ft = fm.group(1), fm.group(2)
                    si = re.search(r'si="(\d+)"', fa)
                    if 't="shared"' in fa and si:
                        rec['si'] = si.group(1)
                        if ft:
                            shared[si.group(1)] = (ref, ft)
                    if ft:
                        rec['f'] = ft
                cells.append((ref, rec))
            for ref, rec in cells:
                if rec['f'] is None and rec['si'] is not None and rec['si'] in shared:
                    mref, mtext = shared[rec['si']]
                    rec['f'] = translate(mtext, mref, ref)
                self.cell[(name, ref)] = rec

    def eval(self, sheet, ref, stack=()):
        key = (sheet, ref)
        if key in self.memo:
            if isinstance(self.memo[key], Exception):
                raise self.memo[key]
            return self.memo[key]
        if key in stack:
            raise Unresolved('循環参照')
        rec = self.cell.get(key)
        if rec is None or not rec['f']:
            return self._V(sheet, ref, stack)
        try:
            val = self._run(rec['f'], sheet, stack + (key,))
        except Unresolved as e:
            self.memo[key] = e
            raise
        except Exception as e:
            err = Unresolved(f'{sheet}!{ref}: {type(e).__name__}')
            self.memo[key] = err
            raise err
        self.memo[key] = val
        return val

    def _run(self, formula, sheet, stack):
        masks = []
        t = RE_QUOTED.sub(lambda m: (masks.append(m.group(0)[1:-1]),
                                     f'\x01{len(masks)-1}\x01')[1], formula)
        if any(e in t for e in ('#REF', '#DIV', '#VALUE', '#N/A', '#NAME')):
            raise Unresolved('エラー値を含む')
        t = RE_PCT.sub(lambda m: f'({m.group(1)}/100.0)', t)
        def shname(m):
            if m.group(1) is not None:
                return m.group(1)
            g = m.group(2) or ''
            mm = re.fullmatch(r'\x01(\d+)\x01', g)
            return masks[int(mm.group(1))] if mm else g
        def ref(m):
            sh = shname(m) if (m.group(1) is not None or m.group(2)) else sheet
            a = m.group(3).replace('$', '')
            if m.group(4):
                return f'RNG({sh!r},{a!r},{m.group(4).replace("$", "")!r})'
            return f'V({sh!r},{a!r})'
        t = RE_REF.sub(ref, t).replace('^', '**')
        if '\x01' in t:
            raise Unresolved('未解決のシート名')
        env = {'V': lambda s, r: self._V(s, r, stack),
               'RNG': lambda s, a, b: self._RNG(s, a, b, stack),
               'ROUND': _round, 'SUM': _sum, 'AVERAGE': _avg, 'ROUNDUP': _roundup,
               'ROUNDDOWN': _rounddown, 'IF': _if, 'ABS': abs, 'MAX': _max, 'MIN': _min}
        return eval(t, {'__builtins__': {}}, env)

    def _V(self, sh, ref, stack):
        if sh not in self.wbk.SHEET:
            raise Unresolved(f'未知のシート {sh}')
        rec = self.cell.get((sh, ref))
        if rec is None:
            return 0.0
        #  エラー型(t="e")でも数式があれば再評価する（古いエラーが残っている場合がある）
        if rec['t'] not in (None, 'n') and not (rec['t'] == 'e' and rec['f']):
            raise Unresolved(f'{sh}!{ref} は数値以外')
        if rec['f']:
            return self.eval(sh, ref, stack)
        if rec['v'] in (None, ''):
            return 0.0
        try:
            return float(rec['v'])
        except ValueError:
            raise Unresolved(f'{sh}!{ref} の値が数値でない')

    def _RNG(self, sh, a, b, stack):
        _, c1, _, r1 = split_ref(a)
        _, c2, _, r2 = split_ref(b)
        out = []
        for ci in range(min(col_n(c1), col_n(c2)), max(col_n(c1), col_n(c2)) + 1):
            for ri in range(min(r1, r2), max(r1, r2) + 1):
                ref = f'{col_s(ci)}{ri}'
                rec = self.cell.get((sh, ref))
                if rec is None or (rec['f'] is None and rec['v'] in (None, '')):
                    continue
                #  Excelと同様、文字列・論理値は無視し、エラー値は伝播させる
                if rec['t'] in (None, 'n', 'e'):
                    out.append(self._V(sh, ref, stack))
        return out
