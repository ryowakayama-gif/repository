#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""成果物のチェック（v2）。Word文書の記載値を独立に再計算して突合する。

外部レビューの指摘を反映したv2では、料金を条例別表と同じ税抜単価で設計しているため、
チェックも税抜基準で行う。あわせて、レビューで問題とされた表現が残っていないかを
確認する否定チェックを追加した。
"""
import io, json, math, os, re, sys, zipfile
import xml.etree.ElementTree as ET
import openpyxl
sys.path.insert(0, '/tmp/claude-0/-home-user-repository/670c168c-8281-57ba-9df0-b54358bb5879/scratchpad')
os.chdir('/root/.claude/uploads/670c168c-8281-57ba-9df0-b54358bb5879')
from compare import GYO, KOU, CUR_G, CUR_K, SPECIAL, BASIC, NORMAL

D = '/tmp/claude-0/-home-user-repository/670c168c-8281-57ba-9df0-b54358bb5879/scratchpad/'
DOCX = D + 'doc/下水道使用料改定説明資料.docx'
XLSX = D + '階上町下水道使用料改定_試算エビデンス.xlsx'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
TOTAL = CUR_G + CUR_K

results = []
def check(name, ok, detail=''):
    results.append((name, bool(ok), detail))

# ---- 料金計算（税抜ベース） ----
CURRENT = (1008, 37, 174, 200)
def net2m(V, b, r1, r2, r3):
    n = b * 2
    if V > 10:  n += r1 * min(V - 10, 10)
    if V > 20:  n += r2 * min(V - 20, 80)
    if V > 100: n += r3 * (V - 100)
    return n
def bill(V, *p): return math.floor(net2m(V, *p) * 1.1)
def mnet(v, b, r1, r2, r3):
    n = b
    if v > 5:  n += r1 * min(v - 5, 5)
    if v > 10: n += r2 * min(v - 10, 40)
    if v > 50: n += r3 * (v - 50)
    return n
def revenue(p, ratio):
    t = 0
    for recs in (GYO, KOU):
        for k, v, c, cur in recs:
            t += bill(v, *p) * c if k == NORMAL else (bill(10, *p) * c if k == BASIC else round(cur * ratio))
    return t
def solve_mid(target, r1, r3):
    best = None
    for r2 in range(120, 301):
        t = revenue((1008, r1, r2, r3), 1 + target)
        g = abs(t / TOTAL - 1 - target)
        if best is None or g < best[0]: best = (g, r2, t)
    return best[1], best[2]

# ---- Word文書 ----
z = zipfile.ZipFile(DOCX)
root = ET.fromstring(z.read('word/document.xml'))
doc = ''.join(t.text or '' for t in root.iter(W + 't'))
n_tbl = len(list(root.iter(W + 'tbl')))
n_img = len([n for n in z.namelist() if n.startswith('word/media/') and n.lower().endswith('.png')])
def has(*v): return all(str(x) in doc for x in v)

check('Word文書が開ける', True)
check('図が4点埋め込まれている', n_img == 4, '検出 %d 点' % n_img)
check('表が13点ある', n_tbl >= 13, '検出 %d 点（注記ボックスを含む）' % n_tbl)
check('図1〜図4の見出しがある', has('図1', '図2', '図3', '図4'))
check('表1〜表13の見出しがある', all('表%d' % i in doc for i in range(1, 14)))
check('章立て1〜10＋参考がある', all(h in doc for h in
      ['本資料の趣旨', '現行使用料体系の構造', '単価差の縮小余地', '単価差をどこまで縮めるか',
       '中心案と各家庭への影響', '平年度増収額と参考経費回収率', '段階的な縮小の考え方',
       '条例改正に向けた整理', '令和8年度予算と経営戦略の関係', '留意事項', '数値の出典']))

# ---- レビュー指摘への対応（否定チェック） ----
check('【指摘1】「既に達成」の記載がない', '既に達成' not in doc)
check('【指摘1】経営戦略と直接比較できない旨を明記', has('直接比較はできない'))
check('【指摘2】「調定額ベースの平均増収率」と明記', has('調定額ベースの平均増収率'))
check('【指摘2】6〜10㎥の改定率+97.3%を明記', has('+97.3'))
check('【指摘2】第2段階の6〜10㎥改定率+194.6%を明記', has('+194.6'))
check('【指摘3】「完全是正」の語を使っていない', '完全是正' not in doc)
check('【指摘3】「是正率」の語を使っていない', '是正率' not in doc)
check('【指摘3】「単価差縮小率」を使用', has('単価差縮小率'))
check('【指摘3】174円まで引上げが逆転体系である旨を明記', has('逆転体系'))
check('【指摘4】条例別表が税抜である旨を明記', has('別表は税抜'))
check('【指摘4】改正対象が2条例である旨を明記', has('2条例'))
check('【指摘5】公衆浴場汚水に言及', has('公衆浴場'))
check('【指摘6】「推奨案」ではなく「中心案」', '推奨案' not in doc and has('中心案'))
check('【指摘6】中心案を中心に置いた理由を記載', has('中心に置いた理由'))
check('【指摘6】負担のばらつきを数値で提示', has('5.4ポイント', '10.0ポイント', '23.5ポイント'))
check('【指摘7】「平年度増収額」と明記', has('平年度増収額'))
check('【指摘7】初年度影響額が別途である旨を明記', has('初年度影響額'))
check('【指摘8】料金表が暫定である旨を明記', has('料金表は暫定'))
check('【指摘8】データ不足が単価設計に影響する旨を明記', has('料金単価そのものにも影響'))
check('【指摘9】従量課金対象水量の定義を明記', has('従量課金対象水量'))
check('【指摘9】全汚水量ではない旨を明記', has('全汚水量ではない'))
check('【指摘10】「本来この区分が負担すべき」を削除', '本来この区分が負担すべき' not in doc)
check('【指摘11】参考経費回収率・静態試算と明記', has('参考経費回収率', '静態試算'))
check('【指摘12】経営戦略は負担配分まで決定していない旨を明記', has('負担配分までを事前に決定しているわけではない'))
check('【指摘12】R8予算との差異を分解して提示', has('41,906', '52,273', '10,367'))

# ---- 第9章 R8予算差異ブリッジ ----
check('ブリッジ 改定未実施 4,736千円', has('4,736') and 44463 - 40410 + 6828 * 0.1 == 4735.8)
check('ブリッジ 漁集水洗化率分 299千円', has('299') and round(7810 - 6828 * 1.1) == 299)
check('ブリッジ 消費税相当 4,294千円', has('4,294') and round(47238 - 47238 / 1.1) == 4294)
check('ブリッジ 残差 1,038千円', has('1,038') and round(47238 / 1.1 - 41906) == 1038)
check('ブリッジ 各段階の残高', has('47,538', '47,238', '42,944'))
check('ブリッジ 合計が整合', round(52273 - 4736 - 299 - 4294 - 1038) == 41906)
check('R6検証 公共 税込換算 37,017千円', has('37,017') and round(33652 * 1.1) == 37017)
check('R6検証 漁集 税込換算 7,952千円', has('7,952') and round(7229 * 1.1) == 7952)
check('R6検証 逆算水量 212,865㎥', has('212,865') and round(33652 * 1.1 * 1000 / 173.9) == 212865)
check('R6検証 漁集逆算水量 46,475㎥', has('46,475') and round(7229 * 1.1 * 1000 / 171.1) == 46475)
check('予算の税込換算 46,097千円', has('46,097') and round(41906 * 1.1) == 46097)
check('現況継続との差 1,141千円（2.4%）', has('1,141', '2.4') and round(47238 - 41906 * 1.1) == 1141)
check('R7調定実績 税抜換算 40,526千円', has('40,526') and round(44578708 / 1.1 / 1000) == 40526)
check('R8予算のR6決算比 +2.5%', has('+2.5') and abs(41906 / 40881 - 1 - 0.025) < 0.001)
check('R8予算のR7実績比 +3.4%', has('+3.4') and abs(41906 / 40526 - 1 - 0.034) < 0.001)
check('中心案適用後 46,067千円・増収4,161千円',
      has('46,067', '4,161') and round(41906 * 1.0993) == 46067 and round(41906 * 0.0993) == 4161)
check('経営戦略を税込ベースと明記', has('税込ベース'))
check('単価173.9円がR1〜R4実績平均である旨を明記',
      has('173.9') and abs((172.1 + 174.4 + 175.1 + 174.1) / 4 - 173.9) < 0.05)
check('ケース①「5年毎に10%」を明記', has('5年毎に10%'))
check('二重計上の注意を明記', has('二重に見込む'))

# ---- 調定データ範囲の補正 ----
check('年間全件の調定額 44,578,708円を明記', has('44,578,708'))
check('年間全件の件数 9,220件を明記', has('9,220'))
check('年間全件ベースの平年度増収額 4,427千円／4,024千円',
      has('4,427', '4,024') and round(44578708 * 0.0993) == 4426666)
check('特殊算定の実額 71件64,144円を明記', has('64,144', '903'))
check('補正が改定率・単価に影響しない旨を明記', has('改定率および税抜単価には影響しない'))
check('予算額の前提（合算・税抜・当初予算）を明記', has('合算・税抜・当初予算'))
check('公共単独なら+24%となる旨を明記', has('33,782', '+24'))
check('目標管理を税抜へ統一する旨を明記', has('42,944', '47,521') and round(52273 / 1.1) == 47521)

# ---- 数値の再計算突合 ----
cnt = sum(c for _, _, c, _ in GYO) + sum(c for _, _, c, _ in KOU)
check('調定件数 9,070件', has('9,070') and cnt == 9070, '再計算 %d' % cnt)
check('調定額 42,987,132円', has('42,987,132') and TOTAL == 42987132)
for lab, val in [('基本使用料', 1008), ('6〜10㎥', 37), ('11〜50㎥', 174), ('51㎥〜', 200)]:
    check('表1 現行 %s 税抜%s円' % (lab, format(val, ',')), has(format(val, ',') if val >= 1000 else str(val)))
check('表1 税込換算 1,108.8 / 40.7 / 191.4 / 220.0', has('1,108.8', '40.7', '191.4', '220.0'))
check('単価の開き 4.7分の1 / 4.70倍', has('4.7') and abs(174 / 37 - 4.70) < 0.005)

lvl = revenue((1008, 174, 174, 200), 1.0)
check('表3 6〜10㎥を174円に揃えた調定額 51,480,990円', has('51,480,990') and lvl == 51480990,
      '再計算 %s' % format(lvl, ','))
check('表3 増収余地 8,493,858円', has('8,493,858') and lvl - TOTAL == 8493858,
      '再計算 %s' % format(lvl - TOTAL, ','))
check('表3 現行比 約+19.8%', has('19.8') and abs((lvl / TOTAL - 1) * 100 - 19.8) < 0.06,
      '再計算 %.2f%%' % ((lvl / TOTAL - 1) * 100))

# 従量課金対象水量（公共）
a1 = a23 = a4 = 0
for k, v, c, _ in KOU:
    if k in (SPECIAL, BASIC): continue
    a1 += max(0, min(v, 20) - 10) * c; a23 += max(0, min(v, 100) - 20) * c; a4 += max(0, v - 100) * c
tv = a1 + a23 + a4
check('表2 6〜10㎥の従量課金対象水量 46,394㎥', has('46,394') and a1 == 46394)
check('表2 同 水量シェア 37.9%', has('37.9') and abs(a1 / tv * 100 - 37.9) < 0.06, '再計算 %.2f%%' % (a1 / tv * 100))
for lab, val, shown, pct in [('6〜10㎥', 37 * a1, '1,888,236', 5.3), ('11〜50㎥', 174 * a23, '12,699,581', 35.7),
                             ('51㎥〜', 200 * a4, '2,156,440', 6.1)]:
    check('表2 %s の収入 %s円' % (lab, shown), has(shown) and abs(round(val * 1.1) - int(shown.replace(',', ''))) <= 2,
          '再計算 %s' % format(round(val * 1.1), ','))
    check('表2 %s の収入シェア %.1f%%' % (lab, pct), abs(val * 1.1 / CUR_K * 100 - pct) < 0.06,
          '再計算 %.2f%%' % (val * 1.1 / CUR_K * 100))

LEVELS = [('縮小なし', 37, 214, 0.0, 13.6, 19.6), ('弱い縮小', 55, 202, 7.5, 12.6, 14.8),
          ('中心案', 73, 190, 15.1, 11.6, 10.1), ('強い縮小', 100, 172, 26.4, 10.1, 2.9),
          ('参考174円', 174, 122, 57.4, 5.6, -17.1)]
for lab, r1, r2e, m10, m20, m50 in LEVELS:
    r2, _ = solve_mid(0.10, r1, 220)
    check('表4 %s の11〜50㎥ %d円' % (lab, r2e), has(str(r2e)) and r2 == r2e, '再計算 %d' % r2)
    p = (1008, r1, r2, 220)
    for v, exp in ((10, m10), (20, m20), (50, m50)):
        got = (mnet(v, *p) / mnet(v, *CURRENT) - 1) * 100
        check('表4 %s 月%d㎥ %+.1f%%' % (lab, v, exp), abs(got - exp) < 0.06, '再計算 %+.2f%%' % got)
    red = (r1 - 37) / (174 - 37) * 100
    check('表4 %s の単価差縮小率' % lab, True, '再計算 %.1f%%' % red)

R6 = dict(kou=(33652, 67464), gyo=(7229, 19614))
PLANS = [('参考 約5%', (1008, 55, 182, 210), 1.05, 2135226, 52.3, 38.7),
         ('中心案 約10%', (1008, 73, 190, 220), 1.10, 4269216, 54.8, 40.6),
         ('参考 約15%', (1008, 82, 205, 230), 1.15, 6493678, 57.4, 42.7)]
for lab, p, ratio, inc, rke, rge in PLANS:
    g = sum(bill(v, *p) * c if k == NORMAL else (bill(10, *p) * c if k == BASIC else round(cur * ratio))
            for k, v, c, cur in GYO)
    kk = sum(bill(v, *p) * c if k == NORMAL else (bill(10, *p) * c if k == BASIC else round(cur * ratio))
             for k, v, c, cur in KOU)
    check('表8 %s の平年度増収額 %s円' % (lab, format(inc, ',')), has(format(inc, ',')) and g + kk - TOTAL == inc,
          '再計算 %s' % format(g + kk - TOTAL, ','))
    rk = R6['kou'][0] * (kk / CUR_K) / R6['kou'][1] * 100
    rg = R6['gyo'][0] * (g / CUR_G) / R6['gyo'][1] * 100
    check('表8 %s の参考経費回収率 公共%.1f%%' % (lab, rke), abs(rk - rke) < 0.06, '再計算 %.2f%%' % rk)
    check('表8 %s の参考経費回収率 漁集%.1f%%' % (lab, rge), abs(rg - rge) < 0.06, '再計算 %.2f%%' % rg)

CEN = (1008, 73, 190, 220); ST2 = (1008, 109, 207, 240)
HH = [(5, 2217, 2217, 0, 0.0), (8, 2461, 2699, 238, 9.7), (10, 2624, 3020, 396, 15.1),
      (15, 4538, 5110, 572, 12.6), (20, 6452, 7200, 748, 11.6), (25, 8366, 9290, 924, 11.0),
      (30, 10280, 11380, 1100, 10.7), (40, 14108, 15560, 1452, 10.3),
      (50, 17936, 19740, 1804, 10.1), (100, 39936, 43940, 4004, 10.0)]
for v, c2, r2b, diff, pct in HH:
    gc, gr = bill(2 * v, *CURRENT), bill(2 * v, *CEN)
    check('表6 月%d㎥ %s→%s円' % (v, format(c2, ','), format(r2b, ',')), gc == c2 and gr == r2b,
          '再計算 %s→%s' % (format(gc, ','), format(gr, ',')))
    got = (mnet(v, *CEN) / mnet(v, *CURRENT) - 1) * 100
    check('表6 月%d㎥ 差額%+d円・増減率%+.1f%%' % (v, diff, pct),
          gr - gc == diff and abs(got - pct) < 0.06, '再計算 %+d円 %+.2f%%' % (gr - gc, got))

r2s, tot2 = solve_mid(0.20, 109, 240)
check('表9 第2段階の11〜50㎥ 207円', has('207') and r2s == 207, '再計算 %d' % r2s)
check('表9 第2段階の平年度増収 8,632,400円', tot2 - TOTAL == 8632400, '再計算 %s' % format(tot2 - TOTAL, ','))
for lab, r1, r2v, exp in [('現行', 37, 174, 4.70), ('第1段階', 73, 190, 2.60), ('第2段階', 109, 207, 1.90)]:
    check('図4 %s の単価の開き %.2f倍' % (lab, exp), has('%.2f' % exp) and abs(r2v / r1 - exp) < 0.005,
          '再計算 %.3f' % (r2v / r1))
for v, s2 in [(5, 2217), (10, 3416), (20, 7970), (30, 12524), (50, 21632), (100, 48032)]:
    got = bill(2 * v, *ST2)
    check('表10 第2段階 月%d㎥ %s円' % (v, format(s2, ',')), has(format(s2, ',')) and got == s2,
          '再計算 %s' % format(got, ','))
check('表11 中心案の税込換算 80.3 / 209.0 / 242.0', has('80.3', '209.0', '242.0') and
      abs(73 * 1.1 - 80.3) < 0.05 and abs(190 * 1.1 - 209.0) < 0.05)

# ---- Excel ----
wb = openpyxl.load_workbook(XLSX)
check('Excelに29シートある', len(wb.sheetnames) == 29, '検出 %d' % len(wb.sheetnames))
for sh in ['20_単価差の縮小余地', '21_単価差縮小の水準別比較', '22_料金案と平年度増収',
           '23_世帯への影響', '24_段階的な縮小', '25_条例改正に向けた整理',
           '26_R8予算差異ブリッジ', '27_経営戦略の税込検証', '28_調定実績の集計範囲']:
    check('Excelにシート「%s」がある' % sh, sh in wb.sheetnames)
def cells(sh): return [list(r) for r in wb[sh].iter_rows(values_only=True)]
c20 = cells('20_単価差の縮小余地')
check('Excel 20 の増収余地が8,493,858円', any(r[1] == 8493858 for r in c20 if len(r) > 1))
c22 = cells('22_料金案と平年度増収')
check('Excel 22 の平年度増収額が説明資料 表8 と一致',
      sorted(r[16] for r in c22 if isinstance(r[16], int) and r[16] > 0) == [2135226, 4269216, 6493678, 8632400])
def nums(sh): return {c for r in wb[sh].iter_rows(values_only=True) for c in r
                      if isinstance(c, (int, float))}
n26 = nums('26_R8予算差異ブリッジ')
check('Excel 26 にブリッジ各段階がある', {52273, 47538, 47238, 42944, 41906} <= n26)
check('Excel 26 に増減額がある', {-4736, -299, -4294, -1038} <= n26)
check('Excel 26 に税抜読み替え値がある', {46097, 47521} <= n26)
n27 = nums('27_経営戦略の税込検証')
check('Excel 27 にR6検証値がある', {37017, 37065, 212865, 213135, 33652} <= n27)
n28 = nums('28_調定実績の集計範囲')
check('Excel 28 に年間全件の調定額がある', {44578708, 9220, 40526099} <= n28)
check('Excel 28 に公表値との対比がある', {42987132, 42472607, 9070} <= n28)
check('Excel 28 の年間全件増収額が再計算と一致',
      4426666 in n28 and round(44578708 * 0.0993) == 4426666)
c23 = cells('23_世帯への影響')
m = {r[0]: r for r in c23 if isinstance(r[0], int)}
check('Excel 23 の2か月請求額が説明資料 表6 と一致',
      all(m[v][3] == c2 and m[v][5] == r2b for v, c2, r2b, _, _ in HH))
c25 = cells('25_条例改正に向けた整理')
check('Excel 25 に公衆浴場汚水の行がある', any('公衆浴場' in str(r[0]) for r in c25 if r[0]))

ng = [r for r in results if not r[1]]
print('チェック項目 %d 件 / 合格 %d 件 / 不合格 %d 件' % (len(results), len(results) - len(ng), len(ng)))
for n, ok, d in results:
    if not ok: print('  [NG] %s  %s' % (n, d))
io.open(D + 'check_results.json', 'w', encoding='utf-8').write(
    json.dumps([{'name': n, 'ok': o, 'detail': d} for n, o, d in results], ensure_ascii=False, indent=1))
