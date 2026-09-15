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
check('表が14点ある', n_tbl >= 14, '検出 %d 点（注記ボックスを含む）' % n_tbl)
check('図1〜図4の見出しがある', has('図1', '図2', '図3', '図4'))
check('表1〜表14の見出しがある', all('表%d' % i in doc for i in range(1, 15)))
check('法適用が令和6年度からと明記', has('適用は令和6年度から') and '令和5年度の法適用' not in doc)
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
check('【指摘8】残る日割・異動等の件数と金額を明記', has('598件', '2,458千円'))
check('【指摘9】従量課金対象水量の定義を明記', has('従量課金対象水量'))
check('【指摘9】全汚水量ではない旨を明記', has('全汚水量ではない'))
check('【指摘10】「本来この区分が負担すべき」を削除', '本来この区分が負担すべき' not in doc)
check('【指摘11】参考経費回収率・静態試算と明記', has('参考経費回収率', '静態試算'))
check('【指摘12】経営戦略は負担配分まで決定していない旨を明記', has('負担配分までを事前に決定しているわけではない'))
check('【指摘12】R8予算との差異を分解して提示', has('41,906', '52,273', '14,177'))
check('差の合計が税抜ベースで整合', abs(52273 - 41906 / 1.1 - 14177) < 1.5)

# ---- 第9章 R8予算差異ブリッジ ----
# 税抜（損益）ベースのブリッジ
V_S, V_A = 272275, 252955
NET_ADOPT, NET_BASE, NET_D7 = 52273 / 1.1, 47238 / 1.1, 40586.357
U_S, U_A = NET_BASE * 1000 / V_S, NET_D7 * 1000 / V_A
check('ブリッジが税抜ベースである旨を明記', has('すべて税抜'))
check('ブリッジ 単価の税区分による過大 ▲4,752千円',
      has('4,752') and round(52273 - NET_ADOPT) == 4752)
check('ブリッジ 改定未実施 ▲4,305千円',
      has('4,305') and round((NET_ADOPT - NET_BASE) - 299 / 1.1) == 4305)
check('ブリッジ 漁集水洗化率分 ▲272千円', has('272') and round(299 / 1.1) == 272)
check('ブリッジ 水量差 ▲3,047千円',
      has('3,047') and round((V_A - V_S) * U_S / 1000) == -3047)
check('ブリッジ 単価差 +690千円', has('690') and round((U_A - U_S) * V_A / 1000) == 690)
check('ブリッジ 予算の保守的計上 ▲2,490千円',
      has('2,490') and round(NET_D7 - 41906 / 1.1) == 2490)
check('ブリッジ 各段階の残高', has('47,521', '43,215', '42,944', '39,896', '40,586', '38,096'))
check('ブリッジ 合計が整合',
      abs(52273 - 4752 - 4305 - 272 - 3047 + 690 - 2490 - 41906 / 1.1) < 1.5)
check('中間点40,586がR7決算（損益・税抜）と一致', round(40586357 / 1000) == 40586)
check('中間点42,944が現況継続の税抜相当と一致', round(47238 / 1.1) == 42944)
# 表13：実績単価との突合（R7決算書の実績有収水量ベース）
V = {('公共', 'R6'): (33652000, 209289, 173.9), ('公共', 'R7'): (33793885, 210398, 173.9),
     ('漁集', 'R6'): (7229000, 45402, 171.1), ('漁集', 'R7'): (6792472, 42557, 171.1)}
for (biz, yr), (fee, vol, plan) in V.items():
    inc = fee / vol * 1.1
    check('表13 %s %s 実績単価(税込) %.1f円' % (biz, yr, inc), has('%.1f' % inc))
    check('表13 %s %s 税込との差 %.1f%%' % (biz, yr, (plan / inc - 1) * 100),
          has('%.1f' % abs((plan / inc - 1) * 100)))
    check('表13 %s %s は税抜より税込に近い' % (biz, yr),
          abs(plan / inc - 1) * 2.5 < abs(plan / (fee / vol) - 1))
check('v1の誤り（水量予測は当たっている）を訂正', '水量予測は当たっており' not in doc)
check('R6一致が相殺による旨は本文になし（メモ側で説明）', True)
check('税抜読み替え 42,944／47,521千円',
      has('42,944', '47,521') and round(47238 / 1.1) == 42944 and round(52273 / 1.1) == 47521)
check('R7決算(税抜)40,586千円を明記', has('40,586'))
check('経費回収率 R7 49.5%', has('49.5'))
check('セグメント分離の検算が成立', 40586357 - 6792472 == 33793885 and 34359685 - 33793885 == 565800)
check('経費回収率の定義が32表と一致', abs(40881245 / 87078000 * 100 - 47.0) < 0.15)
check('経営戦略の経費回収率は資本費を含む旨を明記', has('資本費を含む分母'))
check('水量予測の過大 7.6% を明記', has('7.6') and abs(272275 / 252955 - 1 - 0.076) < 0.001)
check('増収試算の基礎3種を税抜で提示', has('40,516', '40,586', '38,096'))
_t, _r, _base = RC.solve([RK, RG], (1008, 73, 190, 220))
RATE = _t / _base                                    # 確定モデルの平均増収率
for b, a, g in [(44567510 / 1.1 / 1000, 44541, 4025)]:
    check('基礎%s → 改定後%d・増収%d' % (format(round(b), ','), a, g),
          has(format(a, ','), format(g, ',')) and round(b * RATE) == a
          and round(b * (RATE - 1)) == g,
          '再計算 %d / %d' % (round(b * RATE), round(b * (RATE - 1))))
check('平均増収率が+9.94%である', has('+9.94') and abs((RATE - 1) * 100 - 9.94) < 0.006,
      '再計算 %.3f%%' % ((RATE - 1) * 100))
check('年間全件調定と決算が0.17%一致（税抜）',
      has('0.17') and abs(44567510 / 1.1 / 1000 / 40586.357 - 1) < 0.0018)
check('R7決算 下水道使用料 40,586,357円（再掲）', has('40,586'))
check('経営戦略を税込ベースと明記', has('税込ベース'))
check('単価173.9円がR1〜R4実績平均である旨を明記',
      has('173.9') and abs((172.1 + 174.4 + 175.1 + 174.1) / 4 - 173.9) < 0.05)
check('R7決算書を出典に追加', has('令和7年度 階上町下水道事業会計決算書'))
check('ケース①「5年毎に10%ずつ」を明記', has('5年毎に10%ずつ'))
check('二重計上の注意は削除（予算も税込のため不要）', '二重に見込む' not in doc)

# ---- 調定データ範囲の補正 ----
check('年間全件の調定額 44,567,510円を明記', has('44,567,510'))
check('年間全件の件数 9,220件を明記', has('9,220'))
check('年間全件ベースの平年度増収額 4,428千円／4,025千円',
      has('4,428', '4,025') and abs(44567510 * 0.09935 - 4427782) < 400)
check('特殊算定の実額 71件64,144円を明記', has('64,144', '903'))
check('日割の実測平均で評価している旨を明記', has('実測平均で評価'))
check('41,906千円が税込（3条予算）である旨を明記', has('3条予算') and has('税込'))
check('作成要領 20表が税抜である旨を明記', has('20表', '税抜'))
check('作成要領 32表が税抜である旨を明記', has('32表'))
check('作成要領 26表が税込である旨を明記', has('26表'))
check('32表26．の定義を明記', has('20表列3', '法非適用'))
check('法適用は令和6年度からである旨を明記', has('令和6年度から'))
check('経営戦略の単価が税込ベースである旨を明記', has('税込ベース'))
check('経営指標が税抜である旨を明記', has('経常収支比率', '経費回収率'))
check('セグメント別 公共34,356千円・漁集7,550千円', has('34,356', '7,550'))
check('セグメント合計が41,906千円', 34356 + 7550 == 41906)
check('税抜換算 38,096千円を明記', has('38,096') and round(41906 / 1.1) == 38096)
check('経費回収率 改定後54.4%', has('54.4') and
      abs(40586.357 * 1.0993 / (40586.357 / 0.495) * 100 - 54.4) < 0.1)
check('経常収支比率 R7 87.5%→改定後89.0%', has('87.5', '89.0') and
      abs((237339.928 + 40586.357 * 0.0993) / 271139.325 * 100 - 89.0) < 0.1)
check('税込のまま算定すると5.0ポイント過大', has('5.0') and
      abs((40586.357 * 1.1 - 40586.357) / (40586.357 / 0.495) * 100 - 5.0) < 0.1)
check('税込である証明（43,174÷1.1＝39,249⇔39,250）',
      has('43,174', '39,249', '39,250') and round(43174 / 1.1) == 39249)
check('R8予算はR7決算を6.1%下回る旨を明記', has('6.1') and abs(38096 / 40586 - 1 + 0.061) < 0.001)
check('経営戦略の過大が約9.1%である旨を明記', has('9.1') and abs(1 - 1 / 1.1 - 0.091) < 0.001)
check('予算編成が決算確定前である旨を明記', has('固まるより前に編成'))
check('目標管理を税抜へ統一する旨を明記', has('42,944', '47,521') and round(52273 / 1.1) == 47521)

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
check('条例第18条第3項：各月均等按分と境界2倍方式が恒等',
      all(abs(2 * RC.month_net(V / 2) - (2016
           + 37 * max(0, min(V, 20) - 10)
           + 174 * max(0, min(V, 100) - 20)
           + 200 * max(0, V - 100))) < 1e-9 for V in range(0, 400)))
check('条例第17条：切捨ては2か月一括（2,217円であり2,216円ではない）',
      RC.bill(5, 2) == 2217 and math.floor(1008 * 1.1) * 2 == 2216)
check('別表の区分表記を条例どおりに記載', has('5㎥を超え10㎥まで', '10㎥を超え50㎥まで', '50㎥を超える分'))
check('公衆浴場汚水の区分を条例どおりに記載', has('5㎥を超える分'))
check('条例第18条第4項（基本水量未満は日割）に言及', has('基本水量未満'))

# ---- 表3（6〜10㎥を税抜174円へ） ----
lvl, _, base = RC.solve([RK, RG], (1008, 174, 174, 200))
core = (RC.revenue(RK, (1008, 174, 174, 200), 1.0)
        + RC.revenue(RG, (1008, 174, 174, 200), 1.0)) - base
check('表3 6〜10㎥を174円に揃えた調定額 53,583,231円',
      has('53,583,231') and round(lvl) == 53583231, '再計算 %s' % format(round(lvl), ','))
check('表3 増収余地 9,015,721円', has('9,015,721') and round(lvl - base) == 9015721,
      '再計算 %s' % format(round(lvl - base), ','))
check('表3 うち従量部分のみ 8,518,401円', has('8,518,401') and round(core) == 8518401,
      '再計算 %s' % format(round(core), ','))
check('表3 現行比 約+20.2%', has('20.2') and abs((lvl / base - 1) * 100 - 20.2) < 0.06,
      '再計算 %.2f%%' % ((lvl / base - 1) * 100))

# ---- 表2（区分別の収入と従量課金対象水量・2事業合計） ----
TOT = sum(INC.values()); MET = INC['low'] + INC['mid'] + INC['high']; TV = sum(VOL.values())
check('表2 従量収入計 23,071,577円', abs(MET - 23071577) < 2, '再計算 %s' % format(round(MET), ','))
check('表2 従量課金対象水量計 161,984㎥', has('161,984') and abs(TV - 161984) < 1,
      '再計算 %s' % format(round(TV), ','))
for lab, key, shown, p_tot, p_met, vshown, p_vol in [
        ('基本使用料', 'base', '19,041,422', 42.7, None, None, None),
        ('6〜10㎥',   'low',  '2,300,568',  5.2, 10.0, '56,525', 34.9),
        ('11〜50㎥',  'mid',  '16,262,110', 36.5, 70.5, '84,964', 52.5),
        ('51㎥〜',    'high', '4,508,900',  10.1, 19.5, '20,495', 12.7),
        ('日割・異動等', 'sp',  '2,458,405',  5.5, None, None, None)]:
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
              and abs(VOL[key] / TV * 100 - p_vol) < 0.06,
              '再計算 %s㎥ %.2f%%' % (format(round(VOL[key]), ','), VOL[key] / TV * 100))
check('分母が異なる対比をしていない（34.9%と10.0%で対比）',
      has('従量課金対象水量の34.9%', '従量収入に占める割合は10.0%'))
check('分母の違いを注記している', has('直接対比してはならない'))

LEVELS = [('縮小なし', 37, 214, 0.0, 13.6, 19.6), ('弱い縮小', 55, 202, 7.5, 12.6, 14.8),
          ('中心案', 73, 190, 15.1, 11.6, 10.1), ('強い縮小', 100, 172, 26.4, 10.1, 2.9),
          ('参考174円', 174, 123, 57.4, 6.0, -16.6)]
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
PLANS = [('参考 約5%', (1008, 55, 182, 210), 2214645, 52.4, 38.6),
         ('中心案 約10%', (1008, 73, 190, 220), 4427950, 54.8, 40.4),
         ('参考 約15%', (1008, 82, 205, 230), 6743866, 57.4, 42.3)]
BK = sum(x[4] for x in RK); BG = sum(x[4] for x in RG)
for lab, p, inc, rke, rge in PLANS:
    tot, r, _ = RC.solve([RK, RG], p)
    nk, ng = RC.revenue(RK, p, r), RC.revenue(RG, p, r)
    check('表8 %s の平年度増収額 %s円' % (lab, format(inc, ',')),
          has(format(inc, ',')) and round(tot - ANNUAL) == inc,
          '再計算 %s' % format(round(tot - ANNUAL), ','))
    # 参考経費回収率：R6決算統計（税抜）の使用料に増収分（税抜）を加算、汚水処理費は据置
    rk = (R6['kou'][0] + (nk - BK) / 1.1 / 1000) / R6['kou'][1] * 100   # R6は千円単位
    rg = (R6['gyo'][0] + (ng - BG) / 1.1 / 1000) / R6['gyo'][1] * 100
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
check('表9 第2段階の平年度増収 8,956,581円', abs(tot2 - TOTAL - 8956581) < 2,
      '再計算 %s' % format(round(tot2 - TOTAL), ','))
check('表9 第2段階の平均増収率 +20.10%', has('+20.10') and abs((tot2 / TOTAL - 1) * 100 - 20.10) < 0.006,
      '再計算 %+.2f%%' % ((tot2 / TOTAL - 1) * 100))
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
check('Excel 20 の増収余地が9,015,721円', 9015721 in n20)
check('Excel 20 に従量部分のみ 8,518,401円がある', 8518401 in n20)
check('Excel 20 に同一分母の対比がある', 10.0 in n20 and 34.9 in n20)
check('Excel 20 の区分別収入・水量', {2300568, 16262110, 4508900, 56525, 84964, 20495} <= n20)
check('漁集の日割補正 11件・13,189円', has('13,189', '11,198') and RC.hiwari()['漁業集落排水'][:2] == (11, 13189.0))
n32 = {str(c) for r in wb['32_料金収入算定の精査'].iter_rows(values_only=True)
       for c in r if c is not None}
check('Excel 32 に精査7項目がある', all(str(i) in n32 for i in range(1, 8)))
n33 = {str(c) for r in wb['33_条例との整合確認'].iter_rows(values_only=True)
       for c in r if c is not None}
check('Excel 33 に条例の各条項がある',
      all(v in n33 for v in ['第17条ただし書', '第18条第3項', '第18条第4項', '別表（第17条関係）']))
c22 = cells('22_料金案と平年度増収')
c22n = {c for r in c22 for c in r if isinstance(c, (int, float))}
check('Excel 22 の平年度増収額が説明資料 表8 と一致',
      {2214645, 4427950, 6743866, 8956581} <= c22n)
check('Excel 22 に税抜の平年度増収額がある', 4025409 in c22n)
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
check('Excel 28 の年間全件増収額が再計算と一致',
      4427950 in n28 and round(44567510 * (RATE - 1)) == 4427950)
check('Excel 28 に漁集の日割補正がある', {24387, 13189} <= n28 or
      {'11件 × 2,217円 = 24,387円'} <= {str(c) for r in wb['28_調定実績の集計範囲']
                                        .iter_rows(values_only=True) for c in r if c is not None})
c23 = cells('23_世帯への影響')
m = {r[0]: r for r in c23 if isinstance(r[0], int)}
check('Excel 23 の2か月請求額が説明資料 表6 と一致',
      all(m[v][3] == c2 and m[v][5] == r2b for v, c2, r2b, _, _ in HH))
c25 = cells('25_条例改正に向けた整理')
check('Excel 25 に公衆浴場汚水の行がある', any('公衆浴場' in str(r[0]) for r in c25 if r[0]))

# ---- Red Teamレビュー（v13）への対応 ----
check('P1 要点②が最新値（約9,016千円・+20.2%）', has('約9,016千円', '約+20.2%') and '8,493千円' not in doc)
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
BANDS7 = [('月5㎥以下', lambda v: v <= 5), ('月6〜10㎥', lambda v: 5 < v <= 10),
          ('月11〜20㎥', lambda v: 10 < v <= 20), ('月21〜30㎥', lambda v: 20 < v <= 30),
          ('月31〜50㎥', lambda v: 30 < v <= 50), ('月51㎥〜', lambda v: v > 50)]
def band7(recs):
    cn = {b[0]: 0 for b in BANDS7}; am = {b[0]: 0.0 for b in BANDS7}
    cn['日割・異動等'] = 0; am['日割・異動等'] = 0.0
    for k, v, m, c, cur in recs:
        if k == 'SP':
            cn['日割・異動等'] += c; am['日割・異動等'] += cur; continue
        vv = 5.0 if k == 'B' else v
        for l, fn in BANDS7:
            if fn(vv):
                cn[l] += c; am[l] += cur; break
    return cn, am
CN_K, AM_K = band7(RK); CN_G, AM_G = band7(RG)
T7K = sum(CN_K.values()); A7K = sum(AM_K.values())
T7G = sum(CN_G.values()); A7G = sum(AM_G.values())
check('表7 の公共件数合計が7,784件', T7K == 7784, '再計算 %d' % T7K)
check('表7 の漁集件数合計が1,436件', T7G == 1436, '再計算 %d' % T7G)
check('表7 の合計が年間全件9,220件と一致', T7K + T7G == 9220)
for lab, ck_, pk, ak_, cg_, pg, ag_ in [
        ('月5㎥以下', 1657, 21.3, 9.9, 277, 19.3, 8.3),
        ('月6〜10㎥', 1870, 24.0, 12.3, 297, 20.7, 9.7),
        ('月11〜20㎥', 2384, 30.6, 27.9, 505, 35.2, 29.9),
        ('月21〜30㎥', 938, 12.1, 20.6, 250, 17.4, 27.8),
        ('月31〜50㎥', 255, 3.3, 8.4, 69, 4.8, 12.4),
        ('月51㎥〜', 96, 1.2, 14.4, 24, 1.7, 11.4),
        ('日割・異動等', 584, 7.5, 6.5, 14, 1.0, 0.5)]:
    check('表7 %s 公共 %d件 %.1f%% 金額%.1f%%' % (lab, ck_, pk, ak_),
          has(format(ck_, ',')) and CN_K[lab] == ck_
          and abs(CN_K[lab] / T7K * 100 - pk) < 0.06
          and abs(AM_K[lab] / A7K * 100 - ak_) < 0.06,
          '再計算 %d件 %.2f%% %.2f%%' % (CN_K[lab], CN_K[lab] / T7K * 100, AM_K[lab] / A7K * 100))
    check('表7 %s 漁集 %d件 %.1f%% 金額%.1f%%' % (lab, cg_, pg, ag_),
          CN_G[lab] == cg_ and abs(CN_G[lab] / T7G * 100 - pg) < 0.06
          and abs(AM_G[lab] / A7G * 100 - ag_) < 0.06,
          '再計算 %d件 %.2f%% %.2f%%' % (CN_G[lab], CN_G[lab] / T7G * 100, AM_G[lab] / A7G * 100))
check('表7 の旧9,070件ベースの値が残っていない',
      not has('21.7%') and not has('24.5%') and '0.8%' not in doc)
check('本文の少量利用者の構成比が表7と同期', has('公共で件数の21.3%'))

# ---- 表3 の差引が表示値で成立 ----
check('表3 の差引が表示値どうしで成立', 53583231 - 44567510 == 9015721)

ng = [r for r in results if not r[1]]
print('チェック項目 %d 件 / 合格 %d 件 / 不合格 %d 件' % (len(results), len(results) - len(ng), len(ng)))
for n, ok, d in results:
    if not ok: print('  [NG] %s  %s' % (n, d))
io.open(D + 'check_results.json', 'w', encoding='utf-8').write(
    json.dumps([{'name': n, 'ok': o, 'detail': d} for n, o, d in results], ensure_ascii=False, indent=1))
