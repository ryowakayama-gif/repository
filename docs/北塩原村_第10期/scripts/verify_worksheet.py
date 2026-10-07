# -*- coding: utf-8 -*-
"""受領した国の推計ワークシート・県の整備量見込みシートと、doc66 に書いた数の照合。

   doc66（国の推計ワークシートの受領と内容確認）は、素案第5章の組み替えの土台になる。
   土台の数が受領データと食い違っていれば、組み替えた素案も食い違う。
   このため doc66 に引いた数を、受領データと素案から機械で引き直して照合する。

   【並びの位置まで見る。】「その数がどこかにあるか」だけを見ていたため、
   表の3つの値のうち2つが入れ替わっている誤り（＋2.2／2.1／2.3pt と書いたが
   正しくは ＋2.1／2.1／2.2pt）を通してしまった。行ごと組み立てて突き合わせる。

   python3 scripts/verify_worksheet.py
"""
import io, os, sys
sys.dont_write_bytecode = True
B = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(B, 'scripts'))
import parse_worksheet as W
import soan_content as SC

print('■ 国の推計ワークシート・県シートと doc66 の照合')

doc = io.open(os.path.join(B, '66_国の推計ワークシートの受領と内容確認.md'),
              encoding='utf-8').read()
j = W.jinko(); k = W.keisu(); n = W.nenji(); h = W.hokenryo()
riyou, teiin = W.seibi()

def f(x): return '{:,.0f}'.format(x)
ng, N = [], [0]
def c(nm, ok, mi=''):
    print(('  ○ ' if ok else '  ✕ ') + nm + (('  … ' + mi) if mi else ''))
    N[0] += 1
    if not ok:
        ng.append(nm)

# 4-4 の内訳
c('前期 444／425／405', all(('%d' % x) in doc for x in j['前期(65～74歳)']),
  '／'.join('%d' % x for x in j['前期(65～74歳)']))
c('後期 560／575／589', '560 | 575 | 589' in doc.replace('　',''),
  '／'.join('%d' % x for x in j['後期(75歳～)']))
c('75〜84歳 357／374／391', all(('| %d |' % x) in doc for x in j['後期(75歳～84歳)']))
c('85歳以上 203／201／198', all(('| %d |' % x) in doc for x in j['後期(85歳～)']))
kz = [j['後期(75歳～)'][i] / j['第1号被保険者数'][i] * 100 for i in range(3)]
c('後期の割合 国 55.8／57.5／59.3%', all(('%.1f%%' % x) in doc for x in kz),
  '／'.join('%.1f' % x for x in kz))
tj = W.TOUHOU_JINKO
sz = [tj['後期'][i] / (tj['前期'][i] + tj['後期'][i]) * 100 for i in range(3)]
c('後期の割合 素案 53.6／55.4／57.0%', all(('%.1f%%' % x) in doc for x in sz),
  '／'.join('%.1f' % x for x in sz))
# 並びの位置まで見る。「どこかにあるか」では取りこぼす（現にそれで誤りを通した）
_gyo = [l for l in doc.splitlines() if l.startswith('| 差 |')]
c('差の行が1本', len(_gyo) == 1, repr(_gyo[:1]))
c('差 ＋2.1／2.1／2.2pt（並びの位置まで）',
  len(_gyo) == 1 and _gyo[0] == '| 差 | ' + ' | '.join(
      '＋%.1fpt' % (kz[i] - sz[i]) for i in range(3)) + ' |',
  '正しくは ' + '／'.join('%.1f' % (kz[i] - sz[i]) for i in range(3)))
_han = [l for l in doc.splitlines() if l.startswith('| 後期高齢者の割合　国 |')]
c('後期の割合 国の行が並びどおり',
  len(_han) == 1 and _han[0] == '| 後期高齢者の割合　国 | ' + ' | '.join(
      '%.1f%%' % x for x in kz) + ' |', repr(_han[:1]))
_hs = [l for l in doc.splitlines() if l.startswith('| 同　素案 |')]
c('後期の割合 素案の行が並びどおり',
  len(_hs) == 1 and _hs[0] == '| 同　素案 | ' + ' | '.join(
      '%.1f%%' % x for x in sz) + ' |', repr(_hs[:1]))
c('開きの幅 2.1〜2.2ポイント',
  ('%.1f〜%.1fポイント開く' % (min(kz[i]-sz[i] for i in range(3)),
                              max(kz[i]-sz[i] for i in range(3)))) in doc)
# 第11期の後期（令和12年度）602人
import openpyxl
wb = openpyxl.load_workbook(W.WS, data_only=True)
s5 = wb['5_保険料推計']
r113 = [c_.value for c_ in list(s5.iter_rows(min_row=113, max_row=113))[0]]
c('令和12年度の後期 602人', ('602' in doc) and r113[9] == 602, '実データ %s' % r113[9])

# 4-3 の見込み
c('県シート 認知症GH R9 27人', riyou['認知症対応型共同生活介護／認知症高齢者 グループホーム'][3] == 27)
c('県シート R8 26人', riyou['認知症対応型共同生活介護／認知症高齢者 グループホーム'][2] == 26)
c('県シート R7 24人', riyou['認知症対応型共同生活介護／認知症高齢者 グループホーム'][1] == 24)
c('到達率 96.3%（26/27）', '96.3%' in doc and abs(26/27*100 - 96.3) < 0.05)
# 素案 5-4(5) の表の値
hyo = None
for ch in SC.CH:
    for sec in ch['sections']:
        if sec['no'] != '5-4':
            continue
        for b in sec['blocks']:
            if b['t'] == 'table' and b['head'] and b['head'][0] == 'サービス' \
               and '村内定員' in b['head']:
                hyo = b
c('素案5-4(5)の表がある', hyo is not None)
if hyo:
    gh = [r for r in hyo['rows'] if r[0].startswith('認知症')][0]
    c('素案 R9 21.6／R10 21.4／R11 21.2',
      gh[5:8] == ['21.6', '21.4', '21.2'], '／'.join(gh[5:8]))
    tou = [r for r in hyo['rows'] if r[0].startswith('通所介護')][0]
    c('本文の最大 83.6% が表にある', '83.6%' in str(hyo['rows']))
    c('R5〜R7 20.2／21.0／23.8', gh[2:5] == ['20.2', '21.0', '23.8'])
# 差 5.4人・年約17,000千円
c('差 5.4人', '▲5.4' in doc and abs(27 - 21.6 - 5.4) < 0.01)
c('1人あたり年3,116千円', '3,116千円' in doc and abs(67304/21.6 - 3115.9) < 1.0,
  '%.1f' % (67304/21.6))
c('5.4人＝年約17,000千円', '17,000千円' in doc and abs(67304/21.6*5.4 - 16826) < 20,
  '%.0f' % (67304/21.6*5.4))
# サービス諸費の差 69,419千円
tg = dict((a, (b, cc)) for a, b, cc, _ in W.totsugou())
sa = tg['サービス諸費'][1] - tg['サービス諸費'][0]
c('サービス諸費の差 69,419千円', f(sa) in doc, f(sa))
c('差の7割を説明', '7割' in doc and 0.6 < (67304/21.6*5.4*3)/sa < 0.8,
  '%.0f%%' % ((67304/21.6*5.4*3)/sa*100))
# 他の施設
for nm, ours, key in (('介護老人保健施設', 15.8, '介護老人保健施設'),
                      ('介護医療院', 1.2, '介護医療院'),
                      ('特別養護老人ホーム', 14.3, '特別養護老人ホーム')):
    c('県シート %s R9 %d人' % (nm, riyou[key][3]), ('| %d |' % riyou[key][3]) in doc)

# 係数
c('負担割合 23%', k['第1号被保険者負担割合'][0] == 0.23)
c('見込交付割合 5.36／5.03／4.67%',
  [round(x*100, 2) for x in k['調整交付金見込交付割合']] == [5.36, 5.03, 4.67])
c('収納率 99.50%', abs(k['予定保険料収納率'][0] - 0.995) < 1e-9)
print()
if ng:
    print('不適合 %d件' % len(ng))
    for x in ng:
        print('  -', x)
    sys.exit(1)
print('適合：doc66 に引いた %d件のすべてが受領データ・素案と一致する' % N[0])
sys.exit(0)
