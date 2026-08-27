# -*- coding: utf-8 -*-
"""中札内村下水道 経営戦略シミュレーションの確定分修正。
openpyxlで開き直すとチャートシート7枚・グラフ・図が失われるため、xlsx内のXMLを直接編集する。"""
import zipfile, json, re

SRC, DST = 'work/sim.xlsx', 'work/sim_fixed.xlsx'
z = zipfile.ZipFile(SRC)
NAMES = json.load(open('sheetmap.json'))
parts = {n: z.read(n) for n in z.namelist()}
log = []

def sx(s):    return 'xl/' + NAMES[s]
def get(s):   return parts[sx(s)].decode('utf-8')
def put(s,x): parts[sx(s)] = x.encode('utf-8')
def cellm(x, ref): return re.search(r'<c r="%s"([^>]*)>(.*?)</c>' % ref, x, re.S)

def cached(x, ref):
    m = cellm(x, ref)
    if not m: return None
    v = re.search(r'<v>([^<]*)</v>', m.group(2))
    return float(v.group(1)) if v else None

def set_cell(x, ref, inner, why):
    m = cellm(x, ref)
    if not m: raise SystemExit(f'!! セル {ref} が見つかりません')
    log.append((ref, m.group(2), inner, why))
    return x[:m.start()] + f'<c r="{ref}"{m.group(1)}>{inner}</c>' + x[m.end():]

def set_value_only(x, ref, val, why):
    """共有数式を壊さないよう <v> だけ差し替える。"""
    m = cellm(x, ref)
    if not m: raise SystemExit(f'!! セル {ref} が見つかりません')
    inner = m.group(2)
    new = re.sub(r'<v>[^<]*</v>', f'<v>{val}</v>', inner) if '<v>' in inner else inner + f'<v>{val}</v>'
    log.append((ref, inner, new, why))
    return x[:m.start()] + f'<c r="{ref}"{m.group(1)}>{new}</c>' + x[m.end():]

COLS = list('PQRSTUVWXY')      # 下水道現況予測 の令和7〜16年度
SRCC = list('FGHIJKLMNO')      # 予測系シート の令和7〜16年度

# ── 1. 論点E：シミュレーションパターン1の使用料単価 ────────────────────────────
x = get('シミュレーションまとめ')
x = set_cell(x, 'F8', '<v>187.55</v>',
             'パターン1使用料単価 180.8→187.55（170.5×110%）。改定額10.3→17.05は式で連動')
put('シミュレーションまとめ', x)

# ── 2. 確認事項4：令和7年度の建設改良費を公表版の93,100千円へ ──────────────────
s = '【下水道】建設改良費の年度別事業費 (2)'
x = get(s)
x = set_cell(x, 'E6',  '<v>12500</v>', '❶管渠更新 令和7年度 0→12,500')
x = set_cell(x, 'E12', '<v>80600</v>', '❷処理場設備等更新 令和7年度 27,424→80,600')
x = set_cell(x, 'E17', '<f>ROUND((E14-E6)*50%,0)</f><v>40300</v>',
             '国庫補助金 令和7年度 直接入力7,712→他年度と同じ算式に復元（40,300）')
x = set_cell(x, 'E21', '<v>46550</v>', '企業債 令和7年度 13,900→46,550（事業費×50%）')
put(s, x)
#  E14 事業費合計 =E6+E8+E10+E12 → 93,100 ／ E23 自己財源 =E14-E17-E19-E21 → 6,250（式のまま自動）

# ── 3. 確認事項5：汚水処理原価の分母を総有収水量（工場排水込み）へ ────────────────
#     基準シート「予測」には合計行(21行)があるが、パターン別シートには無いため
#     各シート内で 19行(一般家庭等)＋20行(工場排水) を足す式にする。
PAIRS = [('下水道現況予測', '予測', True),
         ('下水道現況予測 P1-10%',     '予測 (2)P1-10%',     False),
         ('下水道現況予測 P2-15% (2)', '予測 (2)P2-15% (2)', False),
         ('下水道現況予測 P3-20%',     '予測 (2)P3-20%',     False),
         ('下水道現況予測 P4-30%',     '予測 (2)P4-30%',     False)]

for dst, src, has_total in PAIRS:
    sxml, x = get(src), get(dst)
    q = "'" + src + "'" if re.search(r"[ ()%]", src) else src
    cost = [cached(x, f'{c}16') for c in COLS]
    for i, c in enumerate(COLS):
        sc = SRCC[i]
        if has_total:
            tot, f = cached(sxml, f'{sc}21'), f'{q}!{sc}21'
        else:
            tot = cached(sxml, f'{sc}19') + cached(sxml, f'{sc}20')
            f   = f'{q}!{sc}19+{q}!{sc}20'
        x = set_cell(x, f'{c}19', f'<f>{f}</f><v>{tot:.0f}</v>', f'{dst} 有収水量(C)→総有収水量')
        if cost[i] is not None:
            x = set_value_only(x, f'{c}20', f'{cost[i]*1000/tot:.10f}', f'{dst} 汚水処理原価 再計算')
    put(dst, x)

# ── 4. 開いたときに全再計算させ、古い計算チェーンは除去 ──────────────────────────
parts['xl/workbook.xml'] = get.__globals__['parts']['xl/workbook.xml'].decode('utf-8').replace(
    '<calcPr calcId="191029"/>', '<calcPr calcId="191029" fullCalcOnLoad="1"/>').encode('utf-8')
parts['[Content_Types].xml'] = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', '',
    parts['[Content_Types].xml'].decode('utf-8')).encode('utf-8')
parts['xl/_rels/workbook.xml.rels'] = re.sub(r'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', '',
    parts['xl/_rels/workbook.xml.rels'].decode('utf-8')).encode('utf-8')
parts.pop('xl/calcChain.xml', None)

zo = zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED)
for item in z.infolist():
    if item.filename in parts: zo.writestr(item, parts[item.filename])
zo.close(); z.close()

strip = lambda t: re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', t)).strip()
print(f'=== 適用 {len(log)} セル ===')
for ref, b, a, why in log:
    if ref[-2:] == '20' and ref[0] in 'PQRSTUVWXY': continue
    print(f'  {ref:5s} {strip(b)[:40]:42s} → {strip(a)[:40]:42s} {why}')
print(f'\n（汚水処理原価行 {sum(1 for l in log if l[0][-2:]=="20" and l[0][0] in "PQRSTUVWXY")} セルの再計算は表示省略）')
