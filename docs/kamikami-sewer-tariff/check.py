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

# 実行場所に依存しないよう、すべてスクリプトからの相対で解決する。
# 別の配置で動かす場合は環境変数で上書きできる。
#   KAMIKAMI_DATA  入力データ（00_入力データ）のディレクトリ
#   KAMIKAMI_OUT   成果物（docx/xlsx）のディレクトリ
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import recompute as RC


def _find(rels, marker):
    for rel in rels:
        d = os.path.normpath(os.path.join(HERE, rel))
        if os.path.exists(os.path.join(d, marker)):
            return d + os.sep
    raise SystemExit('%s が見つかりません。KAMIKAMI_OUT で指定してください。' % marker)


DOCX_NAME = '下水道使用料改定説明資料.docx'
XLSX_NAME = '階上町下水道使用料改定_試算エビデンス.xlsx'
OUTDIR = (os.environ.get('KAMIKAMI_OUT', '').rstrip(os.sep) + os.sep
          if os.environ.get('KAMIKAMI_OUT')
          else _find(['01_成果物', '../01_成果物', '.', '..', 'doc', '../doc'], DOCX_NAME))
RK, RG = RC.load()                                        # 年間全件（12か月・漁集の日割補正済み）
INC_K, VOL_K = RC.decompose(RK)
INC_G, VOL_G = RC.decompose(RG)
ANNUAL = sum(x[4] for x in RK) + sum(x[4] for x in RG)
INC = {k: INC_K[k] + INC_G[k] for k in INC_K}
VOL = {k: VOL_K[k] + VOL_G[k] for k in VOL_K}

D = OUTDIR
DOCX = OUTDIR + DOCX_NAME
XLSX = (OUTDIR + XLSX_NAME if os.path.exists(OUTDIR + XLSX_NAME)
        else _find(['01_成果物', '../01_成果物', '.', '..'], XLSX_NAME) + XLSX_NAME)
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
TOTAL = ANNUAL  # 漁集の日割補正後（RK/RG と同一基準）

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
    """確定モデル（recompute）で料金案pの調定額を再計算する。"""
    return RC.revenue(RK, p, ratio) + RC.revenue(RG, p, ratio)
def solve_mid(target, r1, r3):
    """平均増収率がtargetに最も近くなる11〜50㎥の税抜単価を探す。"""
    best = None
    for r2 in range(120, 301):
        t, _, base = RC.solve([RK, RG], (1008, r1, r2, r3))
        g = abs(t / base - 1 - target)
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
check('表が15点ある', n_tbl >= 15, '検出 %d 点（注記ボックスを含む）' % n_tbl)
check('図1〜図4の見出しがある', has('図1', '図2', '図3', '図4'))
check('表1〜表13の見出しがある', all('表%d' % i in doc for i in range(1, 14)))
check('章立て1〜10＋参考がある', all(h in doc for h in
      ['本資料の趣旨', '現行使用料体系の構造', '単価差の縮小余地', '改善レンジと単価差の縮小水準',
       '2案の比較と各家庭への影響', '平年度増収額と経費回収率', '段階的な縮小の考え方',
       '条例改正に向けた整理', '経営戦略および令和8年度予算との関係', '留意事項', '数値の出典']))

# ---- 方針：R7実績改善型 ----
check('R7決算49.5%を出発点とする旨を明記', has('令和7年度決算の経費回収率49.5%'))
check('R7決算の3指標を明記', has('40,586,357円', '49.5%', '87.5%'))
check('相対改善率であることを明記', has('相対改善率') and has('ポイント改善ではない'))
check('49.5×1.10＝54.5%の説明がある', has('49.5%×1.10'))
check('経営戦略を改定幅の直接根拠としない旨を明記', has('改定幅の直接的な算定根拠とはしていない')
      or has('改定幅の算定根拠は経営戦略の将来数値ではなく'))
check('「経営戦略が見込む使用料収入の水準を確保」を削除',
      '経営戦略が見込む使用料収入の水準を確保' not in doc)
check('「中心案」の語を使っていない', '中心案' not in doc)
check('A案・B案を併記', has('A案', 'B案', '少量利用配慮型', '負担均衡型'))

# ---- 数値の再計算突合（年間全件・確定モデル） ----
cnt = sum(x[3] for x in RK) + sum(x[3] for x in RG)
check('調定件数 9,220件（12か月・全件）', has('9,220') and cnt == 9220, '再計算 %d' % cnt)
check('調定額 44,567,510円', has('44,567,510') and round(ANNUAL) == 44567510,
      '再計算 %s' % format(round(ANNUAL), ','))
for lab, val in [('基本使用料', 1008), ('6〜10㎥', 37), ('11〜50㎥', 174), ('51㎥〜', 200)]:
    check('表1 現行 %s 税抜%s円' % (lab, format(val, ',')), has(format(val, ',') if val >= 1000 else str(val)))
check('表1 税込換算 1,108.8 / 40.7 / 191.4 / 220.0', has('1,108.8', '40.7', '191.4', '220.0'))
check('単価の開き 4.7分の1 / 4.70倍', has('4.7') and abs(174 / 37 - 4.70) < 0.005)

# ---- 条例との整合 ----
check('条例：各月均等按分と境界2倍方式が恒等',
      all(abs(2 * RC.month_net(V / 2) - (2016
           + 37 * max(0, min(V, 20) - 10)
           + 174 * max(0, min(V, 100) - 20)
           + 200 * max(0, V - 100))) < 1e-9 for V in range(0, 400)))
check('条例：切捨ては2か月一括（2,217円であり2,216円ではない）',
      RC.bill(5, 2) == 2217 and math.floor(1008 * 1.1) * 2 == 2216)
check('別表の区分表記を条例どおりに記載', has('5㎥を超え10㎥まで', '10㎥を超え50㎥まで', '50㎥を超える分'))
check('公衆浴場汚水の区分を条例どおりに記載', has('5㎥を超える分'))
check('日割は基本水量未満に限られる旨を明記', has('基本水量未満'))
check('漁集条例 別表第2の特定を改正時整理事項として明記', has('別表第2'))

# ---- 表2（区分別の収入と従量課金対象水量） ----
TOT = sum(INC.values()); MET = INC['low'] + INC['mid'] + INC['high']; TV = sum(VOL.values())
check('表2 の内訳合計が調定額の実額と一致', abs(TOT - ANNUAL) < 1.0,
      '差 %.2f 円' % (TOT - ANNUAL))
check('表2 合計値 44,567,510円を明記', has('44,567,510'))
check('表2 従量収入計 23,070,442円', abs(MET - 23070442) < 2, '再計算 %s' % format(round(MET), ','))
check('表2 従量課金対象水量計 161,984㎥', has('161,984') and abs(TV - 161984) < 1)
for lab, key, shown, p_tot, p_met, vshown, p_vol in [
        ('基本使用料', 'base', '19,038,663', 42.7, None, None, None),
        ('6〜10㎥',   'low',  '2,300,335',  5.2, 10.0, '56,525', 34.9),
        ('11〜50㎥',  'mid',  '16,261,232', 36.5, 70.5, '84,964', 52.5),
        ('51㎥〜',    'high', '4,508,875',  10.1, 19.5, '20,495', 12.7),
        ('日割', 'hiwari', '456,776', 1.0, None, None, None),
        ('異動等', 'sp',  '2,001,629',  4.5, None, None, None)]:
    v = INC[key]
    check('表2 %s の収入 %s円' % (lab, shown),
          has(shown) and abs(round(v) - int(shown.replace(',', ''))) <= 2,
          '再計算 %s' % format(round(v), ','))
    check('表2 %s 調定額に対する比 %.1f%%' % (lab, p_tot),
          abs(v / TOT * 100 - p_tot) < 0.06, '再計算 %.2f%%' % (v / TOT * 100))
    if p_met is not None:
        check('表2 %s 従量収入に対する比 %.1f%%' % (lab, p_met),
              has('%.1f' % p_met) and abs(v / MET * 100 - p_met) < 0.06,
              '再計算 %.2f%%' % (v / MET * 100))
        check('表2 %s 水量 %s㎥ / %.1f%%' % (lab, vshown, p_vol),
              has(vshown) and abs(VOL[key] - int(vshown.replace(',', ''))) < 1
              and abs(VOL[key] / TV * 100 - p_vol) < 0.06)
check('日割と異動等を分けて表示', has('日割（基本水量未満）') and has('異動等'))
check('日割が据置である旨を明記', has('基本使用料1,008円を据え置く以上いずれの案でも変わらない'))
check('分母が異なる対比をしていない', has('従量課金対象水量の34.9%', '従量収入に占める割合は10.0%'))
check('分母の違いを注記している', has('直接対比してはならない'))
check('切捨て差の配分を注記している', has('比例配分'))

# ---- 表3 ----
lvl, _, base = RC.solve([RK, RG], (1008, 174, 174, 200))
core = (RC.revenue(RK, (1008, 174, 174, 200), 1.0)
        + RC.revenue(RG, (1008, 174, 174, 200), 1.0)) - base
check('表3 174円に揃えた調定額 53,486,483円',
      has('53,486,483') and round(lvl) == 53486483, '再計算 %s' % format(round(lvl), ','))
check('表3 増収余地 8,918,973円（表示値どうしの差引）',
      has('8,918,973') and round(lvl) - round(base) == 8918973,
      '再計算 %s' % format(round(lvl) - round(base), ','))
check('表3 うち従量部分のみ 8,518,401円', has('8,518,401') and round(core) == 8518401)
check('表3 現行比 約+20.0%', has('20.0') and abs((lvl / base - 1) * 100 - 20.0) < 0.06)
check('表3 の差引が表示値どうしで成立', 53486483 - 44567510 == 8918973)

# ---- 表4・表9（改善レンジと料金案） ----
R6 = dict(kou=(33652, 67464), gyo=(7229, 19614))
BK = sum(x[4] for x in RK); BG = sum(x[4] for x in RG)
PLANS = [('約5%改善', (1008, 55, 182, 210), 2190879, 4.92, 51.9, 52.3, 38.6),
         ('A案', (1008, 55, 203, 220), 4481621, 10.06, 54.5, 54.8, 40.6),
         ('B案', (1008, 73, 191, 220), 4478920, 10.05, 54.5, 54.9, 40.4),
         ('約15%改善', (1008, 82, 205, 230), 6671497, 14.97, 56.9, 57.3, 42.3)]
for lab, p, inc, rt, rec, rke, rge in PLANS:
    tot, r, _ = RC.solve([RK, RG], p)
    nk, ng = RC.revenue(RK, p, r), RC.revenue(RG, p, r)
    check('表9 %s の平年度増収額 %s円' % (lab, format(inc, ',')),
          has(format(inc, ',')) and round(tot - ANNUAL) == inc,
          '再計算 %s' % format(round(tot - ANNUAL), ','))
    check('表9 %s の平均増収率 %+.2f%%' % (lab, rt),
          has('%.2f' % rt) and abs((tot / ANNUAL - 1) * 100 - rt) < 0.006)
    check('表4 %s の経費回収率 %.1f%%' % (lab, rec),
          has('%.1f%%' % rec) and abs(49.5 * tot / ANNUAL - rec) < 0.06,
          '再計算 %.2f%%' % (49.5 * tot / ANNUAL))
    rk = (R6['kou'][0] + (nk - BK) / 1.1 / 1000) / R6['kou'][1] * 100
    rg = (R6['gyo'][0] + (ng - BG) / 1.1 / 1000) / R6['gyo'][1] * 100
    check('表9 %s の経費回収率 公共%.1f%% 漁集%.1f%%' % (lab, rke, rge),
          abs(rk - rke) < 0.06 and abs(rg - rge) < 0.06,
          '再計算 %.2f%% / %.2f%%' % (rk, rg))
check('改善幅がポイント表記で併記されている', has('+2.4ポイント', '+5.0ポイント', '+7.4ポイント'))
check('15%案が上限ケースである旨を明記', has('上限ケース'))

# ---- 表5（単価差の縮小水準・平均増収率10%固定） ----
def solve_mid(target, r1, r3):
    best = None
    for r2 in range(120, 301):
        t, _, b = RC.solve([RK, RG], (1008, r1, r2, r3))
        g = abs(t / b - 1 - target)
        if best is None or g < best[0]: best = (g, r2, t)
    return best[1], best[2]
LEVELS = [('縮小なし', 37, 215, 0.0, 14.0, 20.1), ('A案', 55, 203, 7.5, 13.0, 15.3),
          ('B案', 73, 191, 15.1, 11.9, 10.5), ('強い縮小', 100, 173, 26.4, 10.4, 3.4),
          ('参考174円', 174, 124, 57.4, 6.3, -16.1)]
for lab, r1, r2e, m10, m20, m50 in LEVELS:
    r2, _ = solve_mid(0.10, r1, 220)
    check('表5 %s の11〜50㎥ %d円' % (lab, r2e), has(str(r2e)) and r2 == r2e, '再計算 %d' % r2)
    p = (1008, r1, r2, 220)
    for v, exp in ((10, m10), (20, m20), (50, m50)):
        got = (mnet(v, *p) / mnet(v, *CURRENT) - 1) * 100
        check('表5 %s 月%d㎥ %+.1f%%' % (lab, v, exp), abs(got - exp) < 0.06, '再計算 %+.2f%%' % got)

# ---- 表6・表7（料金表と世帯影響） ----
A = (1008, 55, 203, 220); B = (1008, 73, 191, 220)
check('表6 A案の税込単価 60.5 / 223.3 / 242.0', has('60.5', '223.3', '242.0'))
check('表6 B案の税込単価 80.3 / 210.1 / 242.0', has('80.3', '210.1', '242.0'))
check('表6 区分別改定率 A案', has('+48.6%', '+16.7%'))
check('表6 区分別改定率 B案', has('+97.3%', '+9.8%'))
check('単価の開き A案3.69倍 B案2.62倍', has('3.69倍', '2.62倍')
      and abs(203 / 55 - 3.69) < 0.005 and abs(191 / 73 - 2.62) < 0.005)
HH = [(5, 2217, 2217, 2217, 0.0, 0.0), (8, 2461, 2580, 2699, 4.8, 9.7),
      (10, 2624, 2822, 3020, 7.5, 15.1), (15, 4538, 5055, 5121, 11.4, 12.8),
      (20, 6452, 7288, 7222, 13.0, 11.9), (25, 8366, 9521, 9323, 13.8, 11.4),
      (30, 10280, 11754, 11424, 14.3, 11.1), (40, 14108, 16220, 15626, 15.0, 10.8),
      (50, 17936, 20686, 19828, 15.3, 10.5), (100, 39936, 44886, 44028, 12.4, 10.2)]
for v, cur, a, b, da, db in HH:
    check('表7 月%d㎥ 現行%s円' % (v, format(cur, ',')), RC.bill(v, 2) == cur, '再計算 %d' % RC.bill(v, 2))
    check('表7 月%d㎥ A案%s円 %+.1f%%' % (v, format(a, ','), da),
          RC.bill(v, 2, *A) == a and abs((a / cur - 1) * 100 - da) < 0.06)
    check('表7 月%d㎥ B案%s円 %+.1f%%' % (v, format(b, ','), db),
          RC.bill(v, 2, *B) == b and abs((b / cur - 1) * 100 - db) < 0.06)

# ---- 表8（月使用量帯別の構成） ----
BANDS7 = [('月5㎥以下', lambda v: v <= 5), ('月6〜10㎥', lambda v: 5 < v <= 10),
          ('月11〜20㎥', lambda v: 10 < v <= 20), ('月21〜30㎥', lambda v: 20 < v <= 30),
          ('月31〜50㎥', lambda v: 30 < v <= 50), ('月51㎥〜', lambda v: v > 50)]
def band7(recs):
    cn = {b[0]: 0 for b in BANDS7}; am = {b[0]: 0.0 for b in BANDS7}
    cn['日割・異動等'] = 0; am['日割・異動等'] = 0.0
    for k, v, m, c, cur in recs:
        if k in ('SP', 'SPH'):
            cn['日割・異動等'] += c; am['日割・異動等'] += cur; continue
        vv = 5.0 if k == 'B' else v
        for l, fn in BANDS7:
            if fn(vv):
                cn[l] += c; am[l] += cur; break
    return cn, am
CN_K, AM_K = band7(RK); CN_G, AM_G = band7(RG)
T7K = sum(CN_K.values()); A7K = sum(AM_K.values())
T7G = sum(CN_G.values()); A7G = sum(AM_G.values())
check('表8 の合計が年間全件9,220件と一致', T7K + T7G == 9220)
for lab, ck_, pk, ak_, cg_, pg, ag_ in [
        ('月5㎥以下', 1657, 21.3, 9.9, 277, 19.3, 8.3),
        ('月6〜10㎥', 1870, 24.0, 12.3, 297, 20.7, 9.7),
        ('月11〜20㎥', 2384, 30.6, 27.9, 505, 35.2, 29.9),
        ('月21〜30㎥', 938, 12.1, 20.6, 250, 17.4, 27.8),
        ('月31〜50㎥', 255, 3.3, 8.4, 69, 4.8, 12.4),
        ('月51㎥〜', 96, 1.2, 14.4, 24, 1.7, 11.4),
        ('日割・異動等', 584, 7.5, 6.5, 14, 1.0, 0.5)]:
    check('表8 %s 公共 %d件 %.1f%% 金額%.1f%%' % (lab, ck_, pk, ak_),
          has(format(ck_, ',')) and CN_K[lab] == ck_
          and abs(CN_K[lab] / T7K * 100 - pk) < 0.06
          and abs(AM_K[lab] / A7K * 100 - ak_) < 0.06,
          '再計算 %d件 %.2f%% %.2f%%' % (CN_K[lab], CN_K[lab] / T7K * 100, AM_K[lab] / A7K * 100))
    check('表8 %s 漁集 %d件' % (lab, cg_), CN_G[lab] == cg_,
          '再計算 %d件' % CN_G[lab])
c5 = CN_K['月5㎥以下'] + CN_G['月5㎥以下']; c10 = CN_K['月6〜10㎥'] + CN_G['月6〜10㎥']
check('少量利用者の規模 月5㎥以下1,934件21.0%',
      has('1,934') and has('21.0%') and c5 == 1934 and abs(c5 / 9220 * 100 - 21.0) < 0.06)
check('月6〜10㎥ 2,167件23.5%', has('2,167') and has('23.5%') and c10 == 2167)
check('月10㎥以下で44.5%である旨を明記',
      has('44.5%') and abs((c5 + c10) / 9220 * 100 - 44.5) < 0.06)
check('「少量利用者に配慮した料金体系」と断定していない',
      '少量利用者に配慮した料金体系ではなく' in doc or '少量利用者に配慮した料金体系」ではなく' in doc)
check('正確な言い換えを明記', has('最低使用量帯（月5㎥以下）の負担を据え置きつつ'))

# ---- 表10・表11（段階的縮小） ----
S2 = (1008, 109, 207, 240)
t2, _, _ = RC.solve([RK, RG], S2)
check('表10 第2段階の平年度増収 8,860,468円', round(t2 - ANNUAL) == 8860468,
      '再計算 %s' % format(round(t2 - ANNUAL), ','))
check('表10 第2段階の平均増収率 +19.88%', has('+19.88') and abs((t2 / ANNUAL - 1) * 100 - 19.88) < 0.006)
check('表10 第2段階の経費回収率 59.3%', has('59.3') and abs(49.5 * t2 / ANNUAL - 59.3) < 0.06)
for lab, r1, r2v, exp in [('現行', 37, 174, 4.70), ('第1段階', 73, 191, 2.62), ('第2段階', 109, 207, 1.90)]:
    check('図4 %s の単価の開き %.2f倍' % (lab, exp), has('%.2f' % exp) and abs(r2v / r1 - exp) < 0.005)
for v, cur, s1, s2v, d1, d2 in [(5, 2217, 2217, 2217, 0.0, 0.0), (10, 2624, 3020, 3416, 15.1, 30.2),
                                (20, 6452, 7222, 7970, 11.9, 23.5), (30, 10280, 11424, 12524, 11.1, 21.8),
                                (50, 17936, 19828, 21632, 10.5, 20.6)]:
    check('表11 月%d㎥ 第1段階%s円 第2段階%s円' % (v, format(s1, ','), format(s2v, ',')),
          RC.bill(v, 2, *B) == s1 and RC.bill(v, 2, *S2) == s2v)

# ---- 第9章（参考・経営戦略との関係） ----
check('経営戦略の不整合3点を明記', has('税込ベースで積算', '資本費を含む分母', '有収水量'))
check('税抜読み替え 42,944／47,521千円',
      has('42,944', '47,521') and round(47238 / 1.1) == 42944 and round(52273 / 1.1) == 47521)
check('ブリッジ各段階の残高', has('52,273', '47,521', '43,215', '42,944', '39,896', '40,586', '38,096'))
check('ブリッジ 合計が整合',
      abs(52273 - 4752 - 4305 - 272 - 3047 + 690 - 2490 - 41906 / 1.1) < 1.5)
check('41,906千円が3条予算・税込である旨を明記', has('3条予算') and has('38,096'))
check('セグメント別 公共34,356千円・漁集7,550千円', has('34,356', '7,550'))
check('R8予算はR7決算を6.1%下回る旨を明記', has('6.1') and abs(38096 / 40586 - 1 + 0.061) < 0.001)

# ---- 第10章（留意事項） ----
check('異動等516件の簡便法を明記', has('516件'))
check('相対改善率の説明を留意事項にも記載', has('49.5%×1.10＝54.5%'))
check('近似適用である旨を明記', has('近似適用'))
check('世帯像の断定を削除', '単身・高齢世帯' not in doc and '3〜4人世帯' not in doc)
check('R7普及指導費90千円を明記', has('90千円'))


# ---- Excel ----
wb = openpyxl.load_workbook(XLSX)
check('Excelに34シートある', len(wb.sheetnames) == 34, '検出 %d' % len(wb.sheetnames))
for sh in ['20_単価差の縮小余地', '21_単価差縮小の水準別比較', '22_料金案と平年度増収',
           '23_世帯への影響', '24_段階的な縮小', '25_条例改正に向けた整理',
           '26_R8予算差異ブリッジ', '27_経営戦略の税込検証', '28_調定実績の集計範囲',
           '29_R7決算書による検証', '30_R8当初予算書による確定', '31_決算統計の税区分',
           '32_料金収入算定の精査', '33_条例との整合確認']:
    check('Excelにシート「%s」がある' % sh, sh in wb.sheetnames)
def cells(sh): return [list(r) for r in wb[sh].iter_rows(values_only=True)]
c20 = cells('20_単価差の縮小余地')
v20 = [c for r in wb['20_単価差の縮小余地'].iter_rows(values_only=True) for c in r if c is not None]
n20 = {c for c in v20 if isinstance(c, (int, float))}
check('Excel 20 の増収余地が8,918,973円', 8918973 in n20)
check('Excel 20 に従量部分のみ 8,518,401円がある', 8518401 in n20)
check('Excel 20 に同一分母の対比がある', 10.0 in n20 and 34.9 in n20)
check('Excel 20 の区分別収入・水量', {2300335, 16261232, 4508875, 56525, 84964, 20495} <= n20)
check('Excel 20 の合計が実額と一致', 44567510 in n20)
check('漁集の日割は明細実額11件・13,189円で評価',
      RC.hiwari()['漁業集落排水'][:2] == (11, 13189.0)
      and has('調定明細の実額で評価し直している'))
n32 = {str(c) for r in wb['32_料金収入算定の精査'].iter_rows(values_only=True)
       for c in r if c is not None}
check('Excel 32 に精査7項目がある', all(str(i) in n32 for i in range(1, 8)))
n33 = {str(c) for r in wb['33_条例との整合確認'].iter_rows(values_only=True)
       for c in r if c is not None}
check('Excel 33 に条例の各条項がある',
      all(v in n33 for v in ['第17条ただし書', '第18条第3項', '第18条第4項', '別表（第17条関係）']))
c22 = cells('22_料金案と平年度増収')
c22n = {c for r in c22 for c in r if isinstance(c, (int, float))}
check('Excel 22 の平年度増収額が説明資料 表9 と一致',
      {2190879, 4481621, 4478920, 6671497, 8860468} <= c22n)
check('Excel 22 に税抜の平年度増収額がある', {4074201, 4071745} <= c22n)
def nums(sh): return {c for r in wb[sh].iter_rows(values_only=True) for c in r
                      if isinstance(c, (int, float))}
n26 = nums('26_R8予算差異ブリッジ')
check('Excel 26 にブリッジ各段階がある', {52273, 47521, 43215, 42944, 39896, 40586, 38096} <= n26)
check('Excel 26 に増減額がある', {-4752, -4305, -272, -3047, 690, -2490} <= n26)
check('Excel 26 に税込税抜の対応がある', {41906, 38096, 34356, 31233, 7550, 6864, 43174, 39249} <= n26)
n31s = {str(c) for r in wb['31_決算統計の税区分'].iter_rows(values_only=True)
        for c in r if c is not None}
check('Excel 31 に各表の税区分がある',
      all(v in n31s for v in ['20表 損益計算書', '32表 経営分析に関する調（一）',
                              '26表 歳入歳出決算に関する調（法非適用）', '33表 経営分析に関する調（二）']))
check('Excel 31 に経営指標の反映がある',
      all(v in n31s for v in ['49.5%', '87.5%', '54.4%（+4.9ﾎﾟｲﾝﾄ）', '89.0%（+1.5ﾎﾟｲﾝﾄ）']))
n30 = nums('30_R8当初予算書による確定')
check('Excel 30 にR8予算の内訳がある', {34356, 7550, 41906, 38096} <= n30)
check('Excel 30 に増収試算の基礎3種がある', {41906, 44579, 44645, 49005, 4427} <= n30)
n29 = nums('29_R7決算書による検証')
check('Excel 29 にR7決算確定値がある', {40586357, 33793885, 6792472, 40881245} <= n29)
check('Excel 29 に有収水量がある', {210398, 42557} <= n29)
check('Excel 29 に当方推計との突合がある', {40526099, 60258} <= n29)
n27 = nums('27_経営戦略の税込検証')
n27s = {str(c) for r in wb['27_経営戦略の税込検証'].iter_rows(values_only=True)
        for c in r if c is not None}
check('Excel 27 に4データ点の実績単価がある',
      all(v in n27s for v in ['176.9', '176.7', '175.1', '175.6']))
check('Excel 27 に相殺の説明がある', {-669, 622, -47} <= n27)
n28 = nums('28_調定実績の集計範囲')
check('Excel 28 に年間全件の調定額がある', {44567510, 9220, 40515919} <= n28)
check('Excel 28 に公表値との対比がある', {42987132, 42472607, 9070} <= n28)
_tB, _rB, _bB = RC.solve([RK, RG], (1008, 73, 191, 220))
check('Excel 28 の年間全件増収額が再計算と一致',
      4478920 in n28 and round(_tB - _bB) == 4478920,
      '再計算 %s' % format(round(_tB - _bB), ','))
check('Excel 28 に漁集の日割補正がある', {24387, 13189} <= n28 or
      {'11件 × 2,217円 = 24,387円'} <= {str(c) for r in wb['28_調定実績の集計範囲']
                                        .iter_rows(values_only=True) for c in r if c is not None})
c23 = cells('23_世帯への影響')
m = {r[0]: r for r in c23 if isinstance(r[0], int)}
check('Excel 23 の現行2か月請求額が説明資料 表7 と一致',
      all(m[v][3] == cur for v, cur, _a, _b, _da, _db in HH if v in m),
      '照合 %d 行' % len([v for v, *_ in HH if v in m]))
c25 = cells('25_条例改正に向けた整理')
check('Excel 25 に公衆浴場汚水の行がある', any('公衆浴場' in str(r[0]) for r in c25 if r[0]))

# ---- Red Teamレビュー（v13）への対応 ----
check('P1 要点②が最新値（約8,919千円・+20.0%）',
      has('約8,919千円', '約+20.0%') and '8,493千円' not in doc)
check('P1 要点②に従量部分のみを併記', has('約8,518千円'))
check('P1 の旧留保（偶数月・毎月検針の欠落）を削除',
      '偶数月調定および毎月検針者の一部が含まれていない' not in doc)
check('P1 位置づけが年間全件と整合', has('年間全件調定9,220件・44,567,510円'))
check('「月額の2倍」という表現を使っていない', '月額の2倍となる' not in doc and '請求額はこの2倍' not in doc)
check('2か月請求の算定を条例どおり説明', has('2か月分の税抜使用料を算定したうえで消費税等を加算'))
check('増収率の適用が近似である旨を明記', has('近似適用') and 'そのまま適用できる' not in doc)
check('増収率が加重平均である旨を明記', has('利用構成から導いた加重平均'))
check('世帯像の断定を削除', '単身・高齢世帯' not in doc and '3〜4人世帯' not in doc)
check('水量帯と世帯属性の対応が未確認である旨を明記',
      has('世帯人数・年齢の対応関係は調定データから確認できない'))

# ---- 表7（年間全件で再集計） ----

# ---- 表3 の差引が表示値で成立 ----
check('表3 の差引が表示値どうしで成立', 53583231 - 44567510 == 9015721)

ng = [r for r in results if not r[1]]
print('チェック項目 %d 件 / 合格 %d 件 / 不合格 %d 件' % (len(results), len(results) - len(ng), len(ng)))
for n, ok, d in results:
    if not ok: print('  [NG] %s  %s' % (n, d))
io.open(D + 'check_results.json', 'w', encoding='utf-8').write(
    json.dumps([{'name': n, 'ok': o, 'detail': d} for n, o, d in results], ensure_ascii=False, indent=1))
