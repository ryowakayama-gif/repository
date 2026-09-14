#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""成果物のチェック。Word文書の記載値を独立に再計算して突合する。"""
import io, json, math, os, re, sys, zipfile
import xml.etree.ElementTree as ET
import openpyxl
sys.path.insert(0, '/tmp/claude-0/-home-user-repository/670c168c-8281-57ba-9df0-b54358bb5879/scratchpad')
os.chdir('/root/.claude/uploads/670c168c-8281-57ba-9df0-b54358bb5879')
from compare import monthly, charge2m, CURRENT, GYO, KOU, CUR_G, CUR_K, SPECIAL, BASIC, NORMAL

D = '/tmp/claude-0/-home-user-repository/670c168c-8281-57ba-9df0-b54358bb5879/scratchpad/'
DOCX = D + 'doc/下水道使用料改定説明資料.docx'
XLSX = D + '階上町下水道使用料改定_試算エビデンス.xlsx'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

results = []
def check(name, ok, detail=''):
    results.append((name, bool(ok), detail))

# ---- Word文書を読む ----
z = zipfile.ZipFile(DOCX)
root = ET.fromstring(z.read('word/document.xml'))
doc_text = ''.join(t.text or '' for t in root.iter(W + 't'))
n_tables = len(list(root.iter(W + 'tbl')))
n_images = len([n for n in z.namelist() if n.startswith('word/media/') and n.lower().endswith('.png')])

def has(*vals):
    return all(str(v) in doc_text for v in vals)

# ---- 1 文書の構造 ----
check('Word文書が開ける（XMLパース成功）', True)
check('図が4点埋め込まれている', n_images == 4, '検出 %d 点' % n_images)
check('表が10点ある', n_tables >= 10, '検出 %d 点（注記ボックス3点を含む）' % n_tables)
check('図1〜図4の見出しがある', has('図1', '図2', '図3', '図4'))
check('表1〜表10の見出しがある', all('表%d' % i in doc_text for i in range(1, 11)))
check('章立て1〜8＋参考がある', all(h in doc_text for h in
      ['本資料の趣旨', '現行使用料体系の構造', '6〜10㎥区分の是正余地', '是正水準の選択肢',
       '推奨案と各家庭への影響', '増収額と経費回収率', '段階的な是正の考え方', '留意事項', '数値の出典']))

# ---- 2 調定データの基礎値 ----
cnt = sum(c for _, _, c, _ in GYO) + sum(c for _, _, c, _ in KOU)
tot = CUR_G + CUR_K
check('調定件数 9,070件', has('9,070') and cnt == 9070, '再計算 %d' % cnt)
check('調定額 42,987,132円', has('42,987,132') and tot == 42987132, '再計算 %s' % format(tot, ','))

# ---- 3 係数と是正余地 ----
def coef(recs):
    c = dict(n=0, a1=0, a23=0, a4=0, sp=0)
    for kind, v, k, amt in recs:
        if kind == SPECIAL: c['sp'] += amt; continue
        c['n'] += k
        if kind == BASIC: continue
        c['a1'] += max(0, min(v, 20) - 10) * k
        c['a23'] += max(0, min(v, 100) - 20) * k
        c['a4'] += max(0, v - 100) * k
    return c
cg, ck = coef(GYO), coef(KOU)
A1 = cg['a1'] + ck['a1']; A23 = cg['a23'] + ck['a23']; A4 = cg['a4'] + ck['a4']
N = cg['n'] + ck['n']; SP = cg['sp'] + ck['sp']
gap = (191.4 - 40.7) * A1
check('6〜10㎥の従量水量 56,362㎥', has('56,362') and A1 == 56362, '再計算 %s' % format(A1, ','))
check('単価差 150.7円', has('150.7') and abs((191.4 - 40.7) - 150.7) < 1e-9)
check('是正余地 8,493,753円', has('8,493,753') and round(gap) == 8493753, '再計算 %s' % format(round(gap), ','))
check('是正余地は現行比+19.8%', has('19.8') and abs(gap / tot * 100 - 19.8) < 0.05,
      '再計算 %.2f%%' % (gap / tot * 100))
check('段差 4.7分の1 / 4.70倍', has('4.7') and abs(191.4 / 40.7 - 4.70) < 0.005)

# ---- 4 公共の区分別収入（表2） ----
kv = dict(基本=1108.8 * 2 * ck['n'], 低=40.7 * ck['a1'], 中=191.4 * ck['a23'], 大=220.0 * ck['a4'], 特=ck['sp'])
for label, val, shown, pct in [('基本使用料', kv['基本'], '15,864,710', 44.5),
                               ('6〜10㎥', kv['低'], '1,888,236', 5.3),
                               ('11〜50㎥', kv['中'], '12,699,581', 35.7),
                               ('51㎥〜', kv['大'], '2,156,440', 6.1),
                               ('特殊算定', kv['特'], '3,007,563', 8.4)]:
    check('表2 %s の収入 %s円' % (label, shown), has(shown) and abs(round(val) - int(shown.replace(',', ''))) <= 1,
          '再計算 %s' % format(round(val), ','))
    check('表2 %s の収入シェア %.1f%%' % (label, pct), abs(val / CUR_K * 100 - pct) < 0.06,
          '再計算 %.2f%%' % (val / CUR_K * 100))
vol_tot = ck['a1'] + ck['a23'] + ck['a4']
for label, v, shown, pct in [('6〜10㎥', ck['a1'], '46,394', 37.9), ('11〜50㎥', ck['a23'], '66,351', 54.1),
                             ('51㎥〜', ck['a4'], '9,802', 8.0)]:
    check('表2 %s の水量 %s㎥' % (label, shown), has(shown) and v == int(shown.replace(',', '')))
    check('表2 %s の水量シェア %.1f%%' % (label, pct), abs(v / vol_tot * 100 - pct) < 0.06,
          '再計算 %.2f%%' % (v / vol_tot * 100))

# ---- 5 是正水準（表4・図3） ----
def solve_mid(t, base, r1, r4):
    return round(191.4 + (tot * t - SP * t - (base - 1108.8) * 2 * N - (r1 - 40.7) * A1 - (r4 - 220.0) * A4) / A23, 1)
def plan(base, r1, r2, r4): return dict(base=base, tiers=[(10, r1), (50, r2), (None, r4)])
LEVELS = [('是正なし', 40.7, 0.0, 235.9), ('弱い是正', 60, 12.8, 223.0),
          ('中程度の是正', 80, 26.1, 209.5), ('強い是正', 110, 46.0, 189.4),
          ('完全是正', 191.4, 100.0, 134.7)]
for label, r1, rate, r2 in LEVELS:
    calc = solve_mid(0.10, 1108.8, r1, 242.0)
    check('表4 %s の11〜50㎥ %s円' % (label, r2), has(str(r2)) and abs(calc - r2) < 0.05, '再計算 %.1f' % calc)
    cr = (r1 - 40.7) / (191.4 - 40.7) * 100
    check('表4 %s の是正率 %.1f%%' % (label, rate), abs(cr - rate) < 0.06, '再計算 %.1f%%' % cr)
    p = plan(1108.8, r1, calc, 242.0)
    for v, shown in [(10, None), (20, None), (50, None)]:
        pass
for label, r1, r2, m10, m20, m50 in [('是正なし', 40.7, 235.9, 0.0, 13.8, 19.8),
                                     ('弱い是正', 60, 223.0, 7.4, 12.8, 15.2),
                                     ('中程度', 80, 209.5, 15.0, 11.7, 10.3),
                                     ('強い是正', 110, 189.4, 26.4, 10.1, 3.0),
                                     ('完全是正', 191.4, 134.7, 57.4, 5.8, -16.9)]:
    p = plan(1108.8, r1, solve_mid(0.10, 1108.8, r1, 242.0), 242.0)
    for v, exp in ((10, m10), (20, m20), (50, m50)):
        got = (monthly(v, **p) / monthly(v, **CURRENT) - 1) * 100
        check('表4 %s 月%d㎥ %+.1f%%' % (label, v, exp), abs(got - exp) < 0.06, '再計算 %+.2f%%' % got)

# ---- 6 推奨案（表5・表8） ----
def revenue(recs, p, ratio):
    t = 0
    for kind, v, c, cur in recs:
        t += charge2m(v, **p) * c if kind == NORMAL else (charge2m(10, **p) * c if kind == BASIC else round(cur * ratio))
    return t
R6 = dict(kou=(33652, 67464), gyo=(7229, 19614))
for up, r1, r2x, r4x, inc, rk_e, rg_e in [(0.05, 60, 200.7, 231.0, 2148290, 52.4, 38.7),
                                          (0.10, 80, 209.5, 242.0, 4295086, 54.8, 40.6),
                                          (0.15, 90, 225.1, 253.0, 6447884, 57.3, 42.6)]:
    r2 = solve_mid(up, 1108.8, r1, r4x)
    check('表5 推奨案+%.0f%% の11〜50㎥ %s円' % (up * 100, r2x), has(str(r2x)) and abs(r2 - r2x) < 0.05, '再計算 %.1f' % r2)
    p = plan(1108.8, r1, r2, r4x)
    g, k = revenue(GYO, p, 1 + up), revenue(KOU, p, 1 + up)
    check('表8 推奨案+%.0f%% の増収額 %s円' % (up * 100, format(inc, ',')),
          has(format(inc, ',')) and g + k - tot == inc, '再計算 %s' % format(g + k - tot, ','))
    rk = R6['kou'][0] * (k / CUR_K) / R6['kou'][1] * 100
    rg = R6['gyo'][0] * (g / CUR_G) / R6['gyo'][1] * 100
    check('表8 推奨案+%.0f%% の経費回収率 公共%.1f%%' % (up * 100, rk_e), abs(rk - rk_e) < 0.06, '再計算 %.2f%%' % rk)
    check('表8 推奨案+%.0f%% の経費回収率 漁集%.1f%%' % (up * 100, rg_e), abs(rg - rg_e) < 0.06, '再計算 %.2f%%' % rg)

# ---- 7 世帯への影響（表6） ----
rec = plan(1108.8, 80, solve_mid(0.10, 1108.8, 80, 242.0), 242.0)
HH = [(5, 2217, 2217, 0, 0.0), (8, 2461, 2697, 236, 9.6), (10, 2624, 3017, 393, 15.0),
      (15, 4538, 5112, 574, 12.6), (20, 6452, 7207, 755, 11.7), (25, 8366, 9302, 936, 11.2),
      (30, 10280, 11397, 1117, 10.9), (40, 14108, 15587, 1479, 10.5),
      (50, 17936, 19777, 1841, 10.3), (100, 39936, 43977, 4041, 10.1)]
for v, c2, r2m, diff, pct in HH:
    gc, gr = charge2m(2 * v, **CURRENT), charge2m(2 * v, **rec)
    check('表6 月%d㎥ 現行%s円→推奨%s円' % (v, format(c2, ','), format(r2m, ',')),
          gc == c2 and gr == r2m, '再計算 %s→%s' % (format(gc, ','), format(gr, ',')))
    check('表6 月%d㎥ 差額%+d円・増減率%+.1f%%' % (v, diff, pct),
          gr - gc == diff and abs((monthly(v, **rec) / monthly(v, **CURRENT) - 1) * 100 - pct) < 0.06,
          '再計算 %+d円 %+.2f%%' % (gr - gc, (monthly(v, **rec) / monthly(v, **CURRENT) - 1) * 100))

# ---- 8 段階的是正（表9・表10・図4） ----
st2 = plan(1108.8, 120, solve_mid(0.20, 1108.8, 120, 264.0), 264.0)
check('表9 第2段階の11〜50㎥ 227.2円', has('227.2') and abs(solve_mid(0.20, 1108.8, 120, 264.0) - 227.2) < 0.05,
      '再計算 %.1f' % solve_mid(0.20, 1108.8, 120, 264.0))
check('表9 第2段階の是正率 52.6%', has('52.6') and abs((120 - 40.7) / 150.7 * 100 - 52.6) < 0.06,
      '再計算 %.2f%%' % ((120 - 40.7) / 150.7 * 100))
for label, p, exp in [('現行', 4.70, None), ('第1段階', 2.62, None), ('第2段階', 1.89, None)]:
    pass
for label, r1, r2v, exp in [('現行', 40.7, 191.4, 4.70), ('第1段階', 80, 209.5, 2.62), ('第2段階', 120, 227.2, 1.89)]:
    check('図4 %s の段差 %.2f倍' % (label, exp), has('%.2f' % exp) and abs(r2v / r1 - exp) < 0.005,
          '再計算 %.3f' % (r2v / r1))
for v, s2 in [(5, 2217), (10, 3417), (20, 7961), (30, 12505), (50, 21593), (100, 47993)]:
    got = charge2m(2 * v, **st2)
    check('表10 第2段階 月%d㎥ %s円' % (v, format(s2, ',')), has(format(s2, ',')) and got == s2,
          '再計算 %s' % format(got, ','))

# ---- 9 本文の主要数値 ----
check('経費回収率 公共49.9%・漁集36.9%', has('49.9', '36.9'))
check('本文の単価カーブ記載 221.8 / 131.2 / 220', has('221.8', '131.2', '220'))
check('図1の注記値 221.8 / 131.2 / 179.4 円/㎥ が正しい',
      all(abs(monthly(v, **CURRENT) / v - e) < 0.05 for v, e in ((5, 221.8), (10, 131.2), (50, 179.4))),
      '図中の注記。本文には221.8・131.2を引用')
check('中央値 月10.5㎥', has('10.5'))
check('繰入金が減り始める倍率 2.00倍・2.71倍', has('2.00', '2.71'))
mins = min(((v, monthly(v, **CURRENT) / v) for v in range(5, 201)), key=lambda x: x[1])
check('単価の最小は月10㎥（131.2円/㎥）', mins[0] == 10 and abs(mins[1] - 131.2) < 0.05,
      '再計算 月%d㎥ %.2f円/㎥' % mins)

# ---- 10 Excelエビデンスとの整合 ----
wb = openpyxl.load_workbook(XLSX)
check('Excelに25シートある', len(wb.sheetnames) == 25, '検出 %d' % len(wb.sheetnames))
for sh in ['20_是正余地の算定', '21_是正水準の比較', '22_推奨案の料金表と増収', '23_世帯への影響', '24_段階的是正']:
    check('Excelにシート「%s」がある' % sh, sh in wb.sheetnames)
def cells(sheet):
    return [[c for c in r] for r in wb[sheet].iter_rows(values_only=True)]
c20 = cells('20_是正余地の算定')
check('Excel 20 の増収額が8,493,753円', any(r[1] == 8493753 for r in c20 if len(r) > 1),
      '説明資料 表3 と一致')
c21 = cells('21_是正水準の比較')
data21 = [r for r in c21 if r[0] and sum(1 for c in r if c is not None) > 3 and r[0] != '是正水準']
check('Excel 21 に是正水準5行がある', len(data21) == 5, '検出 %d 行' % len(data21))
c22 = cells('22_推奨案の料金表と増収')
check('Excel 22 の増収額が説明資料 表8 と一致',
      sorted(r[9] for r in c22 if isinstance(r[9], int) and r[9] > 0) == [2148290, 4295086, 6447884])
c23 = cells('23_世帯への影響')
m = {r[0]: r for r in c23 if isinstance(r[0], int)}
check('Excel 23 の2か月請求額が説明資料 表6 と一致',
      all(m[v][3] == c2 and m[v][5] == r2m for v, c2, r2m, _, _ in HH))
c24 = cells('24_段階的是正')
data24 = [r for r in c24 if r[0] and sum(1 for c in r if c is not None) > 3 and r[0] != '段階']
check('Excel 24 に現行・第1段階・第2段階の3行がある', len(data24) == 3, '検出 %d 行' % len(data24))

# ---- 出力 ----
ng = [r for r in results if not r[1]]
print('チェック項目 %d 件 / 合格 %d 件 / 不合格 %d 件' % (len(results), len(results) - len(ng), len(ng)))
for n, ok, d in results:
    if not ok:
        print('  [NG] %s  %s' % (n, d))
io.open(D + 'check_results.json', 'w', encoding='utf-8').write(
    json.dumps([{'name': n, 'ok': o, 'detail': d} for n, o, d in results], ensure_ascii=False, indent=1))
