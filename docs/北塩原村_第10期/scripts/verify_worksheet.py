# -*- coding: utf-8 -*-
"""受領した国の推計ワークシート・県の整備量見込みシートと、doc66 に書いた数の照合。

   doc66（国の推計ワークシートの受領と内容確認）は、素案第5章の組み替えの土台になる。
   土台の数が受領データと食い違っていれば、組み替えた素案も食い違う。
   このため doc66 に引いた数を、受領データと素案から機械で引き直して照合する。

   【並びの位置まで見る。】「その数がどこかにあるか」だけを見ていたため、
   表の3つの値のうち2つが入れ替わっている誤り（＋2.2／2.1／2.3pt と書いたが
   正しくは ＋2.1／2.1／2.2pt）を通してしまった。行ごと組み立てて突き合わせる。

   doc67（被保険者数と認定者数）・doc68（保険料と中長期）・
   doc69（施設・居住系の見込量）に引いた数も同じ考えで照合する。

   python3 scripts/verify_worksheet.py
"""
import io, os, sys
sys.dont_write_bytecode = True
B = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(B, 'scripts'))
import parse_worksheet as W
import soan_content as SC
import shihyo_dict as SD

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
    # 令和8年10月8日より、認知症対応型共同生活介護は県の整備量見込みシートを正本とする。
    # 表は令和6〜11年度の6か年で、県シートの値と一致すること。
    gh = [r for r in hyo['rows'] if r[0].startswith('認知症')][0]
    ken = W.seibi_riyou('認知症対応型共同生活介護')
    SY = ('令和6年度', '令和7年度', '令和8年度',
          '令和9年度', '令和10年度', '令和11年度')
    c('素案5-4(5) 認知症GHが県シートと一致（6か年）',
      [str(x) for x in gh[2:8]] == ['%d' % ken[y] for y in SY],
      '素案%s ／ 県%s' % ('・'.join(str(x) for x in gh[2:8]),
                          '・'.join('%d' % ken[y] for y in SY)))
    teiin = W.seibi_teiin('認知症対応型共同生活介護')['R9']
    c('素案5-4(5) 定員が県シートと一致', gh[1] == '%d人' % teiin, gh[1])
    tatsu = [r for r in hyo['rows'] if str(r[0]).strip().startswith('村内定員に対する')][0]
    c('素案5-4(5) 到達率が県シート÷定員と一致',
      [str(x) for x in tatsu[2:8]] == ['%.1f%%' % (ken[y] / teiin * 100) for y in SY],
      '／'.join(str(x) for x in tatsu[2:8]))
    c('本文の最大 83.6% が表にある', '83.6%' in str(hyo['rows']))
# 差 5.4人・年約17,000千円。
# 21.6 は当方が3か年平均で算定していた値である（10/8に県シートの27へ置き換えた）。
# doc67 は置き換える前の対比を記録したものであるため、素案からは引かない。
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
print('■ doc67（素案第5章の組み替え）の照合')
doc = io.open(os.path.join(B, '67_素案第5章の組み替え_被保険者数と認定者数.md'),
              encoding='utf-8').read()
YM = ('令和9年度','令和10年度','令和11年度')
KH, KZ, KO = W.ws('第1号被保険者数'), W.ws('前期(65～74歳)'), W.ws('後期(75歳～)')
K84, K85, KN = W.ws('後期(75歳～84歳)'), W.ws('後期(85歳～)'), W.ws('第1号_計')

# §4-1 の表が行ごと合うこと
for y in YM + ('令和12年度',):
    want = '| %s | %d | %d | %d | %d | %s | %.1f%% |' % (
        y, KZ[y], KO[y], K84[y], K85[y], '{:,}'.format(int(KH[y])), KO[y]/KH[y]*100)
    c('§4-1 %s の行' % y, want in doc, want if want not in doc else '')

# §4-2 の表が行ごと合うこと
LV = ('要支援1','要支援2','要介護1','要介護2','要介護3','要介護4','要介護5')
for y, lab in (('令和8年度','令和8年度（実績）'),) + tuple((y, y) for y in YM):
    want = '| %s | %s | %d | %.1f%% |' % (
        lab, ' | '.join('%d' % W.ws('第1号_'+lv)[y] for lv in LV), KN[y], KN[y]/KH[y]*100)
    c('§4-2 %s の行' % y, want in doc, want if want not in doc else '')

# §4-3 の幅
tot = int(sum(KN[y] for y in YM))
c('§4-3 3か年計 629', '| 629 |' in doc and tot == 629, str(tot))
c('§4-3 差 ＋6／＋6／＋10／＋22',
  '| ②と①の差 | ＋6 | ＋6 | ＋10 | ＋22 | ― |' in doc
  and [214-KN[YM[0]], 217-KN[YM[1]], 220-KN[YM[2]]] == [6,6,10] and 651-tot == 22)
c('§4-3 ＋3.5%', '＋3.5%' in doc and abs((651-tot)/tot*100 - 3.5) < 0.05,
  '%.2f' % ((651-tot)/tot*100))
c('§4-3 約200円', '＋約200円' in doc and 190 <= (651-tot)/tot*100*59 <= 215,
  '%.0f円' % ((651-tot)/tot*100*59))

# §3-1 の取り違え
c('§3-1 総数635', '635人' in doc and int(sum(W.ws('総数_計')[y] for y in YM)) == 635)
c('§3-1 第1号629・▲3.4%', '**629人**' in doc and '**▲3.4%**' in doc
  and abs((tot-651)/651*100 + 3.4) < 0.05, '%.2f' % ((tot-651)/651*100))

# §3-2 当方の認定率を国の内訳に当てた値
for y, v in zip(YM, (220, 224, 228)):
    calc = KZ[y]*0.0586 + KO[y]*0.3470
    c('§3-2 %s %d人' % (y, v), ('**%d**' % v) in doc and round(calc) == v, '%.1f' % calc)
# 認定率
for y, v in zip(YM, (20.7, 21.1, 21.1)):
    c('§3-2 国の認定率 %s %.1f%%' % (y, v), abs(KN[y]/KH[y]*100 - v) < 0.05)
c('§3-2 当方21.3／21.7／22.1%',
  all(abs(v/KH[y]*100 - w) < 0.05 for y, v, w in
      zip(YM, (214,217,220), (21.3,21.7,22.1))))

# §4-1 の1人あたり給付費（参考シート）
import openpyxl
wb = openpyxl.load_workbook(W.WS, data_only=True)
sk = wb['(参考)保険料の推計に要する係数']
g = {}
for r in sk.iter_rows():
    v = [x.value for x in r]
    for x in v[:4]:
        if x and '1人あたり給付費' in str(x):
            g[str(x).strip()] = v[3]
c('§4-1 85歳以上 月80,362円', '80,362円' in doc
  and g.get('85歳以上後期高齢者の1人あたり給付費') == 80362)
c('§4-1 75〜84歳 17,647円', '17,647円' in doc
  and g.get('85歳未満後期高齢者の1人あたり給付費') == 17647)
c('§4-1 前期 4,296円', '4,296円' in doc
  and g.get('前期高齢者の1人あたり給付費') == 4296)
c('§4-1 令和17年度の85歳以上228人', '令和17年度（228人）' in doc and K85['令和17年度'] == 228)

# 指標辞書が受領ファイルから引けていること
c('指標辞書 令和9年度1,004人', SD.S['第1号被保険者数_令和9年度'][0] == int(KH['令和9年度']))
c('指標辞書 認定者数 令和9年度208人', SD.S['認定者数_令和9年度'][0] == int(KN['令和9年度']))


print()
print('■ doc68（保険料と中長期）の照合')
doc = io.open(os.path.join(B, '68_素案第5章の組み替え_保険料と中長期.md'), encoding='utf-8').read()
m = W.hokenryo_meisai(); k = W.hokenryo_ki(); kk = W.keisu_nendo()
# §1 の内訳表を行ごと
for nm, lab in (('総給付費','| 総給付費 |'), ('在宅サービス','| 　在宅サービス |'),
                ('居住系サービス','| 　居住系サービス |'), ('施設サービス','| 　施設サービス |'),
                ('その他給付費','| その他給付費 |'), ('地域支援事業費','| 地域支援事業費 |')):
    v, p = m[nm]
    want = '%s %s円 | %.0f%% |' % (lab, '{:,.2f}'.format(v), p*100)
    c('§1 %s' % nm, want in doc, want if want not in doc else '')
c('§1 市町村特別給付費等 ▲115.17円 ▲2%',
  '| 市町村特別給付費等 | ▲115.17円 | ▲2% |' in doc and abs(m['市町村特別給付費等'][0]+115.17)<0.01)
c('§1 収納必要額 6,610.55円', '**6,610.55円**' in doc and abs(m['保険料収納必要額（月額）'][0]-6610.55)<0.01)
c('§1 基準額 5,513.72円', '**5,513.72円**' in doc and abs(m['基準保険料額（月額）'][0]-5513.72)<0.01)
c('§1 取崩 1,096.83円', '| 準備基金取崩額 | 1,096.83円 | 17% |' in doc)
c('§1 ▲1,247.04円 ▲18.45%', '1,247.04円（▲18.45%）' in doc
  and abs((6760.76-k['保険料基準額（月額）']['第10期'])-1247.04)<0.01
  and abs((k['保険料基準額（月額）']['第10期']-6760.76)/6760.76*100+18.45)<0.01)
c('§1 施設・居住系61%', '61%' in doc and round((m['居住系サービス'][1]+m['施設サービス'][1])*100)==61)
# §2 の期ごとの表
YM = ('令和11年度','令和12年度','令和17年度','令和22年度','令和27年度','令和32年度')
for ki, y in zip(W.KI, YM):
    base = k['保険料基準額（月額）']['第10期']
    v = k['保険料基準額（月額）'][ki]
    hi = '基準' if ki=='第10期' else '＋%.1f%%' % ((v-base)/base*100)
    cj = '%.2f%%' % (kk['調整交付金見込交付割合'][y]*100)
    ok = ('{:,.2f}'.format(v) in doc) and (hi in doc) and (cj in doc)
    c('§2 %s %s円 %s %s' % (ki, '{:,.2f}'.format(v), hi, cj), ok)
c('§2 F 0.9753→0.7790', '0.9753' in doc and '0.7790' in doc
  and abs(kk['後期高齢者加入割合補正係数']['令和17年度']-0.9753)<1e-9
  and abs(kk['後期高齢者加入割合補正係数']['令和22年度']-0.779)<1e-9)
g = W.kyufu_ki()
c('§2 総給付費 第12期6,945.82円・第14期6,282.81円',
  '6,945.82円' in doc and '6,282.81円' in doc
  and abs(g['総給付費']['第12期']-6945.82)<0.01 and abs(g['総給付費']['第14期']-6282.81)<0.01)
# §3 復元
zen={'令和17年度':(321,368,228),'令和22年度':(262,297,286)}
kuni={'令和17年度':(0.4076,0.3951,0.1973),'令和22年度':(0.3946,0.3898,0.2157)}
tanka=(4296,17647,80362)
for y,want in (('令和17年度',0.8604),('令和22年度',0.7459)):
    a=zen[y]; t=sum(a); mura=[x/t for x in a]
    f=sum(kn*u for kn,u in zip(kuni[y],tanka))/sum(mm*u for mm,u in zip(mura,tanka))
    c('§3 復元 %s %.4f' % (y, want), ('%.4f'%want) in doc and abs(f-want)<0.0001, '%.4f'%f)
c('§3 単価 4,296／17,647／80,362円',
  all(x in doc for x in ('4,296円','17,647円','80,362円')))
# §5 残高
c('§5 残高42,128,183円', '42,128,183円' in doc
  and int(k['準備基金の残高（前年度末の見込額）']['第10期'])==42128183)
c('§5 G 1.0197／1.0199／1.0193',
  '1.0197／1.0199／1.0193' in doc
  and [round(kk['所得段階別加入割合補正係数'][y],4) for y in ('令和9年度','令和10年度','令和11年度')]==[1.0197,1.0199,1.0193])
c('§5 収納率99.50%', '99.50%' in doc and abs(W.keisu()['予定保険料収納率'][0]-0.995)<1e-9)

print()
print('■ doc69（施設・居住系の見込量を県シートに揃えた）の照合')
doc69 = io.open(os.path.join(B, '69_施設居住系の見込量を県シートに揃えた.md'),
                encoding='utf-8').read()
SB69 = {m: W.seibi_riyou(m) for m in
        ("認知症対応型共同生活介護", "介護老人福祉施設", "介護老人保健施設",
         "介護医療院", "特定施設入居者生活介護")}
# §1 実績の対比（当方の算出と県シート）
for m69, a6, a7 in (("認知症対応型共同生活介護", '21.0', '23.8'),
                    ("介護老人福祉施設", '14.1', '12.8'),
                    ("介護老人保健施設", '15.9', '16.6'),
                    ("介護医療院", '1.2', '1.0'),
                    ("特定施設入居者生活介護", '2.6', '1.9')):
    want = '| %s | %s | %d | %s | %d |' % (m69, a6, SB69[m69]['令和6年度'],
                                           a7, SB69[m69]['令和7年度'])
    c('doc69 §1 %s の実績の対比' % m69, want in doc69, '' if want in doc69 else want)
# §2 見込みの対比。差の向きは「当方 − 県」（doc66 §4-3 と同じ）
for m69, ours in (("認知症対応型共同生活介護", 21.6), ("介護老人保健施設", 15.8),
                  ("介護医療院", 1.2), ("特定施設入居者生活介護", 2.3),
                  ("介護老人福祉施設", 14.3)):
    d69 = ours - SB69[m69]['令和9年度']
    lab = ('▲%.1f' if d69 < 0 else '＋%.1f') % abs(d69)
    if m69 == "認知症対応型共同生活介護":
        want = '| %s | %s | **%d** | **%s** |' % (m69, ours, SB69[m69]['令和9年度'], lab)
    else:
        want = '| %s | %s | %d | %s |' % (m69, ours, SB69[m69]['令和9年度'], lab)
    c('doc69 §2 %s の見込みの対比' % m69, want in doc69, '' if want in doc69 else want)
T69 = W.seibi_teiin('認知症対応型共同生活介護')['R9']
for y69, v69 in (('令和6年度', 77.8), ('令和7年度', 88.9), ('令和8年度', 96.3)):
    c('doc69 §3 到達率 %s %.1f%%' % (y69, v69),
      ('%.1f%%' % v69) in doc69
      and abs(SB69['認知症対応型共同生活介護'][y69] / T69 * 100 - v69) < 0.05)
c('doc69 §3 計画期間は満床（100%）',
  '100%' in doc69 and all(SB69['認知症対応型共同生活介護'][y] == T69
                          for y in ('令和9年度', '令和10年度', '令和11年度')))
c('doc69 §6 サービス諸費の差 858,882／789,463千円・8.8%',
  all(x in doc69 for x in ('858,882千円', '789,463千円', '8.8%')))

print()
if ng:
    print('不適合 %d件' % len(ng))
    for x in ng:
        print('  -', x)
    sys.exit(1)
print('適合：doc66〜doc69 に引いた %d件のすべてが受領データ・素案・辞書と一致する' % N[0])
sys.exit(0)
