# -*- coding: utf-8 -*-
"""中札内村下水道 経営戦略シミュレーション：協議会案（令和9年度適用）への組み替え。

 C-1 使用料改定は令和9年度から適用（Excelは既に令和9年度起点のため数値変更なし。注記を追加）
 C-7 他会計繰入金ブロックの不具合
      ① 99・100行の年度見出しが1年ずれ（かつ令和17年度が重複）→ 25・26行から複写
      ② 102行「うち基準内繰入金」が直接入力 → 本体23行への参照へ
      ③ 令和7年度の他会計補助金88,514千円（33行）の根拠を注記
 C-9 一律％改定 → 協議会案（水質使用料＋一般改定100円案）の2本立てへ組み替え
 E-6 「簡水現況予測」の給水原価が #DIV/0!（19行 有収水量の参照先が空セル）

openpyxlで開き直すとチャートシート7枚・グラフ7点・図14点が失われるため、xlsx内のXMLを直接編集する。
LibreOfficeがこの環境で動作しないため、再計算は本スクリプト内の簡易評価器で行う。
"""
import zipfile, re, sys, collections

NSM = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
SRC = sys.argv[1] if len(sys.argv) > 1 else 'work/sim.xlsx'
DST = sys.argv[2] if len(sys.argv) > 2 else 'work/sim_kyogikai.xlsx'

z = zipfile.ZipFile(SRC)
parts = {n: z.read(n) for n in z.namelist()}
infos = z.infolist()
log = []

# ── シート名 → パート名 ───────────────────────────────────────────────────
_wb = parts['xl/workbook.xml'].decode()
_rl = parts['xl/_rels/workbook.xml.rels'].decode()
_rid2tgt = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]*)"', _rl))
SHEET = {}
for m in re.finditer(r'<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"', _wb):
    tgt = _rid2tgt[m.group(2)]
    SHEET[m.group(1)] = 'xl/' + tgt.lstrip('/')
del _wb, _rl

def get(sheet):      return parts[SHEET[sheet]].decode('utf-8')
def put(sheet, x):   parts[SHEET[sheet]] = x.encode('utf-8')
def q(name):         return f"'{name}'" if re.search(r"[ ()%\-]", name) else name

# ── 共有文字列 ───────────────────────────────────────────────────────────
_ss = parts['xl/sharedStrings.xml'].decode()
_si_texts = [''.join(re.findall(r'<t[^>]*>([^<]*)</t>', s.split('<rPh')[0]))
             for s in re.findall(r'<si>(.*?)</si>', _ss, re.S)]
_new_si, _added = {}, 0

def sstr(text):
    """共有文字列に追加（既存なら再利用）し、インデックスを返す。"""
    global _added
    if text in _new_si:
        return _new_si[text]
    if text in _si_texts:
        _new_si[text] = _si_texts.index(text)
        return _new_si[text]
    idx = len(_si_texts)
    _si_texts.append(text)
    _new_si[text] = idx
    _added += 1
    return idx

def flush_sharedstrings():
    global _ss
    if not _added:
        return
    esc = lambda t: t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    extra = ''.join(f'<si><t xml:space="preserve">{esc(t)}</t></si>'
                    for t in _si_texts[len(_si_texts) - _added:])
    cnt = int(re.search(r'count="(\d+)"', _ss).group(1))
    _ss = re.sub(r'count="\d+" uniqueCount="\d+"',
                 f'count="{cnt + _added}" uniqueCount="{len(_si_texts)}"', _ss, count=1)
    _ss = _ss.replace('</sst>', extra + '</sst>')
    parts['xl/sharedStrings.xml'] = _ss.encode('utf-8')

# ── セル編集ヘルパ ───────────────────────────────────────────────────────
def cellm(x, ref):
    return re.search(r'<c r="%s"([^>]*?)(?:/>|>(.*?)</c>)' % ref, x, re.S)

def cell_inner(x, ref):
    m = cellm(x, ref)
    return m.group(2) if m and m.group(2) else ''

def cached(x, ref):
    v = re.search(r'<v>([^<]*)</v>', cell_inner(x, ref))
    if not v:
        return None
    try:
        return float(v.group(1))
    except ValueError:
        return None

def set_cell(x, ref, inner, why, t=None):
    """セルの中身を差し替える（t属性も必要なら書き換える）。"""
    m = cellm(x, ref)
    if not m:
        raise SystemExit(f'!! セル {ref} が見つかりません')
    attrs = m.group(1)
    attrs = re.sub(r'\s+t="[^"]*"', '', attrs)
    if t:
        attrs += f' t="{t}"'
    log.append((ref, m.group(2) or '', inner, why))
    return x[:m.start()] + f'<c r="{ref}"{attrs}>{inner}</c>' + x[m.end():]

def set_str(x, ref, text, why):
    return set_cell(x, ref, f'<v>{sstr(text)}</v>', why, t='s')

def copy_cell_value(x, src_ref, dst_ref, why):
    """src の値と t属性を dst へ複写し、dst の書式(s)は保持する。"""
    sm, dm = cellm(x, src_ref), cellm(x, dst_ref)
    if not sm or not dm:
        raise SystemExit(f'!! 複写元/先が見つかりません {src_ref}→{dst_ref}')
    st = re.search(r'\st="([^"]*)"', sm.group(1))
    return set_cell(x, dst_ref, sm.group(2) or '', why, t=st.group(1) if st else None)

CO = [chr(c) for c in range(ord('A'), ord('Z') + 1)]
def colseq(a, b):
    """列記号 a..b の並び（2文字列まで）。"""
    def n(c):
        v = 0
        for ch in c:
            v = v * 26 + (ord(ch) - 64)
        return v
    def s(v):
        out = ''
        while v:
            v, r = divmod(v - 1, 26)
            out = chr(65 + r) + out
        return out
    return [s(i) for i in range(n(a), n(b) + 1)]

PLAN   = colseq('P', 'Y')      # 現況予測系：令和7〜16年度
PLAN_S = colseq('F', 'O')      # 予測系：令和7〜16年度
YSUM   = colseq('D', 'M')      # まとめ：令和7〜16年度

print('■ 1. シート名の変更（パターン2 → 協議会案）')
RENAME = [('予測 (2)P2-15% (2)', '予測 (2)協議会案'),
          ('下水道現況予測 P2-15% (2)', '下水道現況予測 協議会案')]
TOUCH = ['xl/workbook.xml', 'xl/worksheets/sheet16.xml',
         'xl/worksheets/sheet21.xml', 'docProps/app.xml']
for pn in TOUCH:
    x = parts[pn].decode('utf-8')
    before = x
    for old, new in RENAME:
        x = x.replace(old, new)
    if x != before:
        parts[pn] = x.encode('utf-8')
for old, new in RENAME:
    SHEET[new] = SHEET.pop(old)
    print(f'   {old} → {new}')
# 改名後に取り残しがないか確認
for pn, b in parts.items():
    if pn.endswith(('.xml', '.rels')):
        s = b.decode('utf-8', 'replace')
        for old, _ in RENAME:
            assert old not in s, f'!! {pn} に旧シート名が残存'

print('\n■ 2. C-9「予測 (2)協議会案」を水質使用料＋一般改定の構造へ')
S15 = '予測 (2)協議会案'
x = get(S15)

# 改定の適用は令和9年度（H列）から。令和7・8年度（F・G列）は現行単価のまま。
COLS15 = colseq('F', 'P')          # 令和7〜17年度
REV15  = colseq('H', 'P')          # 令和9年度以降

# 一般家庭等：現行単価＋一般改定（基本＋超過）の増収
for c in REV15:
    x = set_cell(x, f'{c}23', f'<f>ROUND({c}19*170.5/1000,0)+{c}48</f><v>0</v>',
                 '一般家庭等 使用料収入＝現行単価＋一般改定（基本＋100円/月・超過＋10円/㎥）')
# 工場排水：現行単価＋水質使用料
for c in REV15:
    x = set_cell(x, f'{c}24', f'<f>ROUND({c}20*170.5/1000,0)+{c}50</f><v>0</v>',
                 '工場排水 使用料収入＝現行単価＋水質使用料')
# 平均単価の表示（まとめシートが参照）
x = set_cell(x, 'I26', '<f>ROUND(L25/(L19+L20)*1000,2)</f><v>0</v>',
             '改定後の平均単価（令和13年度の使用料収入÷総有収水量）')

# ── 協議会案の内訳ブロック（44〜51行）を追加 ─────────────────────────────
def c_str(ref, text, s=None):
    sa = f' s="{s}"' if s else ''
    return f'<c r="{ref}"{sa} t="s"><v>{sstr(text)}</v></c>'
def c_num(ref, val, s='8'):
    return f'<c r="{ref}" s="{s}"><v>{val}</v></c>'
def c_f(ref, f, s='8'):
    return f'<c r="{ref}" s="{s}"><f>{f}</f><v>0</v></c>'

NEW15 = []
NEW15.append((44, [c_str('E44', '○協議会案の内訳（令和９年度改定／第３回下水道使用料改定検討協議会）')]))
NEW15.append((45, [c_str('E45', '年間賦課件数（件）')] +
                  [c_f(f'{c}45', f'ROUND({c}14*8,0)') for c in COLS15]))
NEW15.append((46, [c_str('E46', '基本使用料の増収（＋100円／月）（千円）')] +
                  [c_num(f'{c}46', 0) for c in ('F', 'G')] +
                  [c_f(f'{c}46', f'ROUND({c}45*100/1000,0)') for c in REV15]))
NEW15.append((47, [c_str('E47', '超過使用料の増収（＋10円／㎥）（千円）')] +
                  [c_num(f'{c}47', 0) for c in ('F', 'G')] +
                  [c_f(f'{c}47', f'ROUND({c}19*10/1000,0)') for c in REV15]))
NEW15.append((48, [c_str('E48', '一般改定（100円案）増収計（千円）')] +
                  [c_f(f'{c}48', f'{c}46+{c}47', s='30') for c in COLS15]))
NEW15.append((49, [c_str('E49', '水質使用料の単価（円／㎥）')] +
                  [c_num(f'{c}49', 0) for c in ('F', 'G')] +
                  [c_num(f'{c}49', 70) for c in REV15]))
NEW15.append((50, [c_str('E50', '水質使用料の収入（千円）')] +
                  [c_f(f'{c}50', f'ROUND({c}20*{c}49/1000,0)') for c in COLS15]))
NEW15.append((51, [c_str('E51', '改定による増収計（千円）')] +
                  [c_f(f'{c}51', f'{c}48+{c}50', s='30') for c in COLS15]))
NEW15.append((52, [c_str('E52', '※年間賦課件数は協議会の令和13年度予測21,000件を同年度の水洗化人口で除した'
                                '8.0件／人／年で各年度に按分。水質使用料は薬品費12,000千円相当のうち'
                                '不足充足に必要な額（協議会資料p.23の約8,245千円）を工場排水量で除した70円／㎥。')]))

add = ''.join(f'<row r="{r}" spans="5:16" x14ac:dyDescent="0.2">{"".join(cs)}</row>'
              for r, cs in NEW15)
assert '</sheetData>' in x
x = x.replace('</sheetData>', add + '</sheetData>', 1)
x = re.sub(r'<dimension ref="A1:R42"/>', '<dimension ref="A1:R52"/>', x, count=1)
put(S15, x)
for r, cs in NEW15:
    log.append((f'{r}行', '(新規)', f'{len(cs)}セル', '協議会案の内訳ブロックを追加'))
print(f'   23・24行を組み替え（{len(REV15)}列×2）、I26を平均単価へ、44〜52行を新設')

print('\n■ 3. C-7 他会計繰入金ブロックの不具合')

# ── ① 年度見出しのずれ（1年先・令和17年度の重複）を本体から複写 ─────────────
LABELS = [('下水道現況予測',          99, 100),
          ('下水道現況予測 P1-10%',    99, 100),
          ('下水道現況予測 協議会案',   99, 100),
          ('下水道現況予測 P3-20%',    99, 100),
          ('下水道現況予測 P4-30%',    99, 100),
          ('簡水現況予測',            100, 101)]
for sheet, rlab, ryear in LABELS:
    x = get(sheet)
    cols = [c for c in colseq('P', 'AR')
            if cellm(x, f'{c}{rlab}') and cellm(x, f'{c}25')]
    before = cached(x, f'P{ryear}')
    for c in cols:
        x = copy_cell_value(x, f'{c}25', f'{c}{rlab}', f'{sheet} 年度見出しを本体25行から複写')
        x = copy_cell_value(x, f'{c}26', f'{c}{ryear}', f'{sheet} 西暦を本体26行から複写')
    put(sheet, x)
    print(f'   ① {sheet}：{rlab}・{ryear}行の{len(cols)}列を修正'
          f'（先頭 {before:.0f}年 → {cached(get(sheet), f"P{ryear}"):.0f}年）')

# ── ② 基準内繰入金の直接入力を本体23行への参照へ（本体シートのみ） ───────────
x = get('下水道現況予測')
chk = [(c, cached(x, f'{c}102'), cached(x, f'{c}23')) for c in PLAN]
bad = [(c, a, b) for c, a, b in chk if a is None or b is None or abs(a - b) > 0.5]
assert not bad, f'!! 102行と23行の値が一致しません: {bad}'
for c, v, _ in chk:
    x = set_cell(x, f'{c}102', f'<f>{c}23</f><v>{v:.0f}</v>',
                 'うち基準内繰入金 直接入力→本体23行（＝収益的収支の基準内繰入額）への参照')
# ── ③ 令和7年度アンカー（33行88,514千円）の根拠を注記 ───────────────────
anchor  = cached(x, 'P33')
in_amt  = cached(x, 'P102')
out_amt = cached(x, 'P103')
note = (f'※令和７年度の他会計補助金{anchor:,.0f}千円は令和６年度決算額を据置いたもの。'
        f'内訳は基準内{in_amt:,.0f}千円（＝P49−P35＋P51）、基準外{out_amt:,.0f}千円で、'
        f'令和８年度以降の基準外繰入金は年５％逓減としている（103行）。'
        f'令和７年度のみ循環参照を避けるため33行に直接入力している。')
assert not re.search(r'<row r="108"', x), '!! 108行が既に存在します'
x = x.replace('</sheetData>',
              f'<row r="108" spans="3:48" ht="30" customHeight="1" x14ac:dyDescent="0.55">'
              f'<c r="C108" s="144" t="s"><v>{sstr(note)}</v></c></row></sheetData>', 1)
put('下水道現況予測', x)
log.append(('C108', '(新規)', note[:40] + '…', 'C-7③ アンカー値の根拠を注記'))
print(f'   ② 下水道現況予測：102行の{len(chk)}列を「＝23行」参照へ')
print(f'   ③ 下水道現況予測：108行に根拠注記を追加（{anchor:,.0f}＝基準内{in_amt:,.0f}＋基準外{out_amt:,.0f}）')

print('\n■ 4. E-6「簡水現況予測」の給水原価 #DIV/0! を解消')
x = get('簡水現況予測')
y = get('予測')
assert cached(x, 'P19') in (0, None), '!! 19行が既に修正済みの可能性'
tot = [cached(y, f'{c}36') for c in PLAN_S]
assert all(t and t > 0 for t in tot), f'!! 予測シートの簡水有収水量が取得できません: {tot}'
for i, c in enumerate(PLAN):
    x = set_cell(x, f'{c}19', f'<f>予測!{PLAN_S[i]}36</f><v>{tot[i]:.0f}</v>',
                 '有収水量の参照先を空セル(行8)→「予測」シートの簡水有収水量へ')
# 共有数式の親を Z19 へ付け替える（P19 を実式にしたため）
m = cellm(x, 'Z19')
assert m and 't="shared"' in (m.group(2) or ''), '!! Z19 が共有数式ではありません'
x = set_cell(x, 'Z19', '<f t="shared" ref="Z19:AR19" si="5">Z8</f><v>0</v>',
             '共有数式の親をP19→Z19へ付け替え（計画期間外は従前どおり）')
put('簡水現況予測', x)
print(f'   19行 P〜Y の{len(PLAN)}列を「＝予測!F36〜O36」へ（令和7年度 {tot[0]:,.0f}㎥）')

print('\n■ 4b. 指標名の誤記を訂正（汚染処理費→汚水処理費）')
#  共有文字列を直接書き換えるため、参照している全シート（各6箇所）に一度に効く
FIXWORD = [('汚染処理費（千円）', '汚水処理費（千円）'),
           ('汚染処理原価（円）', '汚水処理原価（円）')]
for old, new in FIXWORD:
    assert old in _ss, f'!! 共有文字列に {old} がありません'
    n = _ss.count(f'<t>{old}</t>')
    _ss = _ss.replace(f'<t>{old}</t>', f'<t>{new}</t>')
    i = _si_texts.index(old)
    _si_texts[i] = new
    log.append(('共有文字列', old, new, '総務省の指標名は「汚水処理費」「汚水処理原価」'))
    print(f'   {old} → {new}（{n}件の定義を書換え＝各シート6箇所に反映）')

print('\n■ 5. 「シミュレーションまとめ」を協議会案に合わせて再構成')
S21 = 'シミュレーションまとめ'
x = get(S21)
KG = q('下水道現況予測 協議会案')
KY = q('予測 (2)協議会案')

# パターン2 → 採用案（協議会案）
x = set_str(x, 'B14', '採用案：協議会案（一般改定100円案＋水質使用料）', 'パターン2→協議会案に改称')
x = set_cell(x, 'J14', '<f>F14/$F$2-1</f><v>0</v>',
             '改定率 直接入力0.15→平均単価から算出（協議会の全体改定率120.9%に対応）')

# 空の32〜38行ブロックを協議会案の増収内訳へ
x = set_str(x, 'B32', '協議会案の内訳（令和９年度改定による増収額）', '空のサンプル枠（パターン4の重複表記）を内訳表へ')
BREAK = [(34, '一般改定：基本使用料（＋100円／月）（千円）', 46),
         (35, '一般改定：超過使用料（＋10円／㎥）（千円）',  47),
         (36, '水質使用料（工場排水・70円／㎥）（千円）',    50),
         (37, '増収計（千円）',                            51)]
for r, label, srow in BREAK:
    x = set_str(x, f'C{r}', label, f'{r}行の見出しを協議会案の内訳へ')
    for i, c in enumerate(YSUM):
        ref = f'{c}{r}'
        assert cellm(x, ref), f'!! まとめ {ref} が見つかりません'
        x = set_cell(x, ref, f'<f>{KY}!{PLAN_S[i]}{srow}</f><v>0</v>',
                     f'{label} を予測シートから参照')
x = set_str(x, 'C38', '改定後の使用料収入（千円）', '38行の見出しを改定後使用料収入へ')
for i, c in enumerate(YSUM):
    if cellm(x, f'{c}38'):
        x = set_cell(x, f'{c}38', f'<f>{KG}!{PLAN[i]}28</f><v>0</v>',
                     '改定後の使用料収入を現況予測シートから参照')

# 注記（48〜50行を新設）
NOTES = [
    '※１　使用料改定は令和９（２０２７）年度から適用する。使用料算定期間は令和９〜令和１３年度'
    '（第２回・第３回下水道使用料改定検討協議会）。令和７・令和８年度は現行単価のまま。',
    '※２　改定の構造は協議会案による２本立て。①一般改定（基本使用料＋100円／月・超過使用料＋10円／㎥）、'
    '②水質使用料（高濃度汚水を排出する工場の排水へ70円／㎥を新設）。'
    '令和１３年度の不足額12,735千円を①約4,490千円と②約8,245千円で充足する整理に対応している。',
    '※３　経費回収率は総務省の経営比較分析表の定義（使用料収入÷汚水処理費。資本費を含む）で算定している。'
    '協議会資料の経費回収率は分母を「維持管理費−不明水処理費」（資本費０）とした別定義であり、'
    '同一年度でも数値が異なる。',
]
assert not re.search(r'<row r="48"', x), '!! 48行が既に存在します'
addn = ''.join(f'<row r="{48 + i}" spans="2:14" ht="28" customHeight="1" x14ac:dyDescent="0.55">'
               f'<c r="B{48 + i}" t="s"><v>{sstr(t)}</v></c></row>' for i, t in enumerate(NOTES))
x = x.replace('</sheetData>', addn + '</sheetData>', 1)
x = re.sub(r'<dimension ref="B2:N46"/>', '<dimension ref="B2:N50"/>', x, count=1)
put(S21, x)
for i, t in enumerate(NOTES):
    log.append((f'B{48+i}', '(新規)', t[:36] + '…', 'C-1/C-9/C-8の注記を追加'))
print('   14行を採用案（協議会案）へ、32〜38行を増収内訳へ、48〜50行に注記3件')

# ════════════════════════════════════════════════════════════════════════
#  再計算（LibreOfficeがこの環境で動作しないため簡易評価器で行う）
#  解決できない参照（外部ブック・#REF!）を含むセルはキャッシュを温存する。
# ════════════════════════════════════════════════════════════════════════
import math

class Unresolved(Exception):
    pass

CELL = r'\$?[A-Z]{1,3}\$?\d{1,7}'
PFX  = r"(?:'([^']*)'|([^\s!,()+\-*/:&%'\"=<>]+))!"
#  範囲と単一参照は1回で置換する（範囲を先に変換すると、生成した
#  RNG('sheet','E17','E23') の中身を単一参照の置換が二重に書き換えてしまう）
RE_REF    = re.compile(f'(?:{PFX})?({CELL})(?::({CELL}))?')
RE_PCT    = re.compile(r'(\d+(?:\.\d+)?)%')
RE_QUOTED = re.compile(r"'[^']*'")

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

def split_ref(r):
    m = re.fullmatch(r'(\$?)([A-Z]{1,3})(\$?)(\d+)', r)
    return m.group(1), m.group(2), m.group(3), int(m.group(4))

def translate(text, src, dst):
    """共有数式の親(src)から子(dst)へ相対参照をずらす。"""
    _, sc, _, sr = split_ref(src)
    _, dc, _, dr = split_ref(dst)
    dcol, drow = col_n(dc) - col_n(sc), dr - sr
    masks = []
    def mask(m):
        masks.append(m.group(0))
        return f'\x01{len(masks)-1}\x01'
    t = RE_QUOTED.sub(mask, text)
    def shift(m):
        a, c, b, r = split_ref(m.group(0))
        return f'{a}{c if a else col_s(col_n(c) + dcol)}{b}{r if b else r + drow}'
    t = re.compile(CELL).sub(shift, t)
    for i, v in enumerate(masks):
        t = t.replace(f'\x01{i}\x01', v)
    return t

class Book:
    def __init__(self):
        self.cell = {}          # (sheet, ref) -> {'f':str|None, 'v':str|None, 't':str|None}
        self.memo = {}
        for name, pn in SHEET.items():
            x = parts[pn].decode('utf-8')
            shared = {}
            cells = []
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

    def num(self, sheet, ref):
        rec = self.cell.get((sheet, ref))
        if rec is None:
            return 0.0
        #  エラー型(t="e")でも数式があれば再評価する（古いエラーが残っている場合がある）
        if rec['t'] not in (None, 'n') and not (rec['t'] == 'e' and rec['f']):
            raise Unresolved(f'{sheet}!{ref} は数値以外')
        if rec['f']:
            return self.eval(sheet, ref)
        if rec['v'] in (None, ''):
            return 0.0
        try:
            return float(rec['v'])
        except ValueError:
            raise Unresolved(f'{sheet}!{ref} の値が数値でない: {rec["v"]}')

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
            return self.num(sheet, ref)
        try:
            val = self._run(rec['f'], sheet, stack + (key,))
        except Unresolved as e:
            self.memo[key] = e
            raise
        except (ZeroDivisionError, ValueError, OverflowError, TypeError, SyntaxError,
                NameError, AttributeError, KeyError, IndexError) as e:
            err = Unresolved(f'{sheet}!{ref}: {type(e).__name__}')
            self.memo[key] = err
            raise err
        self.memo[key] = val
        return val

    def _run(self, formula, sheet, stack):
        masks = []
        def mask(m):
            masks.append(m.group(0)[1:-1])
            return f'\x01{len(masks)-1}\x01'
        t = RE_QUOTED.sub(mask, formula)
        if '#REF' in t or '#DIV' in t or '#VALUE' in t or '#N/A' in t:
            raise Unresolved('エラー値を含む')
        t = RE_PCT.sub(lambda m: f'({m.group(1)}/100.0)', t)
        def shname(m):
            if m.group(1) is not None:
                return m.group(1)
            g = m.group(2)
            mm = re.fullmatch(r'\x01(\d+)\x01', g or '')
            return masks[int(mm.group(1))] if mm else g
        def ref(m):
            sh = shname(m) if (m.group(1) is not None or m.group(2)) else sheet
            a = m.group(3).replace('$', '')
            if m.group(4):
                return f'RNG({sh!r},{a!r},{m.group(4).replace("$", "")!r})'
            return f'V({sh!r},{a!r})'
        t = RE_REF.sub(ref, t)
        t = t.replace('^', '**')
        if '\x01' in t:
            raise Unresolved('未解決のシート名')
        env = {'V': lambda s, r: self._V(s, r, stack),
               'RNG': lambda s, a, b: self._RNG(s, a, b, stack),
               'ROUND': _round, 'SUM': _sum, 'AVERAGE': _avg, 'ROUNDUP': _roundup,
               'ROUNDDOWN': _rounddown, 'IF': _if, 'ABS': abs, 'MAX': _max, 'MIN': _min}
        return eval(t, {'__builtins__': {}}, env)

    def _V(self, sh, ref, stack):
        if sh not in SHEET:
            raise Unresolved(f'未知のシート {sh}')
        key = (sh, ref)
        rec = self.cell.get(key)
        if rec is None:
            return 0.0
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
    v = [x for x in _flat(a)]
    if not v:
        raise Unresolved('AVERAGEの対象が空')
    return sum(v) / len(v)
def _max(*a):     return max(_flat(a))
def _min(*a):     return min(_flat(a))
def _if(c, a, b): return a if c else b

print('\n■ 6. 評価器の自己検証（無修正シートのキャッシュを再現できるか）')
bk = Book()
#  (a) 本修正の影響を受けないシートは完全一致でなければならない
for probe in ('予測', '予測 (2)P1-10%', '予測 (2)P3-20%', '予測 (2)P4-30%',
              '下水過去', '簡水過去', '使用料収入'):
    ok = ng = skip = 0
    worst = (0.0, None)
    for (sh, ref), rec in bk.cell.items():
        if sh != probe or not rec['f'] or rec['t'] not in (None, 'n') or rec['v'] in (None, ''):
            continue
        try:
            got, want = bk.eval(sh, ref), float(rec['v'])
        except (Unresolved, ValueError):
            skip += 1
            continue
        d = abs(got - want) / max(1.0, abs(want))
        if d < 1e-9:
            ok += 1
        else:
            ng += 1
            if d > worst[0]:
                worst = (d, f'{ref} 計算{got:,.2f} / キャッシュ{want:,.2f}')
    print(f'   {probe}: 一致{ok} 不一致{ng} 評価不能{skip}'
          + (f'  最大乖離→{worst[1]}' if ng else ''))
    assert ng == 0, f'!! 評価器が {probe} のキャッシュを再現できません: {worst[1]}'

#  (b) 既に納品済の修正（建設改良費93,100千円化・有収水量の付替え）の波及で
#      キャッシュが古いままになっているシートは、その波及行に限られることを確認する
CASCADE = {29, 33, 68, 74, 78, 81, 87, 92, 93, 96, 114, 115, 125}
for probe in ('【下水道】建設改良費の年度別事業費 (2)', '下水道現況予測 P1-10%',
              '下水道現況予測 P3-20%', '下水道現況予測 P4-30%'):
    stale = []
    for (sh, ref), rec in bk.cell.items():
        if sh != probe or not rec['f'] or rec['t'] not in (None, 'n') or rec['v'] in (None, ''):
            continue
        try:
            got, want = bk.eval(sh, ref), float(rec['v'])
        except (Unresolved, ValueError):
            continue
        if abs(got - want) / max(1.0, abs(want)) > 1e-9:
            stale.append(ref)
    rows = {int(re.search(r'\d+', r).group(0)) for r in stale}
    print(f'   {probe}: 旧キャッシュ{len(stale)}セル（行 {sorted(rows)}）')
    assert rows <= CASCADE, f'!! 想定外の行にずれ: {sorted(rows - CASCADE)}'

print('\n■ 7. 全シートのキャッシュを再計算')
def numfmt(v):
    if abs(v - round(v)) < 1e-12 and abs(v) < 1e15:
        return str(int(round(v)))
    return repr(float(v))

def update_v(x, ref, val, drop_type=False):
    m = cellm(x, ref)
    inner = m.group(2) or ''
    new = (re.sub(r'<v>.*?</v>', f'<v>{val}</v>', inner, count=1, flags=re.S)
           if '<v>' in inner else inner + f'<v>{val}</v>')
    attrs = re.sub(r'\s+t="e"', '', m.group(1)) if drop_type else m.group(1)
    return x[:m.start()] + f'<c r="{ref}"{attrs}>{new}</c>' + x[m.end():]

bk.memo.clear()
total_chg = 0
fixed_err = []
for sh in SHEET:
    refs = [r for (s, r) in bk.cell if s == sh]
    if not refs:
        continue
    x = get(sh)
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
            x = update_v(x, ref, numfmt(got), drop_type=(rec['t'] == 'e'))
            chg += 1
            if rec['t'] == 'e':
                fixed_err.append(f'{sh}!{ref}')
    if chg:
        put(sh, x)
        total_chg += chg
        print(f'   {sh}: 更新{chg}セル（評価不能{unres}セル＝外部ブック参照・#REF!のためキャッシュ温存）')
print(f'   計 {total_chg} セルのキャッシュを更新')
if fixed_err:
    print(f'   うちエラー値(#DIV/0!等)が解消したセル {len(fixed_err)}件: '
          + ', '.join(fixed_err[:6]) + (' …' if len(fixed_err) > 6 else ''))

print('\n■ 8. 協議会資料との突合せ（令和13年度）')
KYO = '予測 (2)協議会案'
GEN = '下水道現況予測 協議会案'
chk = [
    ('一般改定：基本使用料の増収', bk.eval(KYO, 'L46'), 2100, '千円'),
    ('一般改定：超過使用料の増収', bk.eval(KYO, 'L47'), 2390, '千円'),
    ('一般改定（100円案）計',      bk.eval(KYO, 'L48'), 4490, '千円'),
    ('水質使用料の収入',          bk.eval(KYO, 'L50'), 8245, '千円'),
    ('改定による増収計',          bk.eval(KYO, 'L51'), 12735, '千円'),
    ('改定後の使用料収入',        bk.eval(GEN, 'V28'), 73543, '千円'),
    ('改定後の平均単価',          bk.eval(KYO, 'I26'), 170.5 * 1.209, '円/㎥'),
    ('経常収支比率（改定後）',     bk.eval(GEN, 'V61') * 100, 105.02, '％'),
]
for name, got, want, unit in chk:
    d = got - want
    print(f'   {name:24s} 本モデル {got:>12,.2f} {unit}  協議会 {want:>10,.2f}  差 {d:+,.2f}')
base = bk.eval('下水道現況予測', 'V28')
print(f'\n   （参考）現行料金継続の使用料収入 {base:,.0f}千円・'
      f'経常収支比率 {bk.eval("下水道現況予測","V61")*100:.2f}％・'
      f'経費回収率 {bk.eval("下水道現況予測","V62")*100:.2f}％（総務省定義）')
print(f'   （参考）改定後の経費回収率 {bk.eval(GEN,"V62")*100:.2f}％（総務省定義）／'
      f'{bk.eval(GEN,"V28")/84619*100:.2f}％（協議会定義：控除後汚水処理費84,619千円）')
print(f'   （参考）全体改定率 {bk.eval(KYO,"I26")/170.5*100:.1f}％（協議会 120.9％）')

print('\n■ 9. 保存')
flush_sharedstrings()
wb = parts['xl/workbook.xml'].decode('utf-8')
if 'fullCalcOnLoad' not in wb:
    wb = re.sub(r'<calcPr([^>]*?)/>', r'<calcPr\1 fullCalcOnLoad="1"/>', wb, count=1)
    parts['xl/workbook.xml'] = wb.encode('utf-8')
parts['[Content_Types].xml'] = re.sub(
    r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', '',
    parts['[Content_Types].xml'].decode('utf-8')).encode('utf-8')
parts['xl/_rels/workbook.xml.rels'] = re.sub(
    r'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', '',
    parts['xl/_rels/workbook.xml.rels'].decode('utf-8')).encode('utf-8')
parts.pop('xl/calcChain.xml', None)

zo = zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED)
done = set()
for it in infos:
    if it.filename in parts:
        zo.writestr(it, parts[it.filename]); done.add(it.filename)
for n, b in parts.items():
    if n not in done:
        zo.writestr(n, b)
zo.close(); z.close()

print(f'\n=== 編集 {len(log)} 件／共有文字列 +{_added} 件 → {DST} ===')
strip = lambda t: re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', t)).strip()
seen = collections.Counter()
for ref, b, a, why in log:
    seen[why] += 1
for why, n in seen.items():
    print(f'  [{n:>3d}件] {why}')
