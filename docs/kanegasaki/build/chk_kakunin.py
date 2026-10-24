# -*- coding: utf-8 -*-
"""金ケ崎町　確認事項が3つの成果品で食い違っていないかを突き合わせる。

  ① 計画素案の附《要回答》（ch_soan.py）  … 正本。件数は見出し「確認事項（N件）」から読む
  ② 確認事項照会_第10期（ch_shokai.py）   … 第2＝《要確定》／第3＝Ｃ・Ｄ／第4＝Ａ・Ｂ
  ③ 確認事項の区分（ch_shiryo.py）        … 第2 Ａ／第3 Ｂ／第4 Ｃ／第5 Ｄ

  使い方　python3 -I chk_kakunin.py          （スクラップパッドで）
          python3 -I chk_kakunin.py <ディレクトリ>
"""
import io
import os
import re
import sys

B = sys.argv[1] if len(sys.argv) > 1 else '.'


def load(n):
    return io.open(os.path.join(B, n), encoding='utf-8').read()


soan, shokai, shiryo = load('ch_soan.py'), load('ch_shokai.py'), load('ch_shiryo.py')


def seg(txt, a, b=None):
    s = txt.index(a)
    return txt[s:txt.index(b, s)] if b else txt[s:]


# 表のます目。番号だけのます目は「**1・6**」「**27・28・29**\n**30・31**」の形で
# 書かれており、表により1列目にも2列目にも来る。番号だけのます目を拾う。
CELL = re.compile(r"'([^']*)'")
ONLY = re.compile(r'^(?:\*\*[0-9・]+\*\*|\\n)+$')


def nos_in(txt):
    out = []
    for c in CELL.findall(txt):
        if ONLY.match(c):
            out += [int(x) for x in re.findall(r'[0-9]+', c)]
    return out


# ── ① 素案の附 ──
NH = re.search(r"H2\('《要回答》　確認事項（(\d+)件）'\)", soan)
NN = int(NH.group(1))
FU = [int(m) for m in re.findall(
    r"\['\*\*(\d+)\*\*', ", seg(soan, NH.group(0), '], widths='))]

# ── ② 照会文書 ──
KN = re.compile(r"'\*\*([Ａ-Ｄ])\*\*\\n\*\*(\d+)\*\*'")
SH3 = [(k, int(n)) for k, n in KN.findall(seg(shokai, "H1('第３　", "H1('第４　"))]
SH4 = [(k, int(n)) for k, n in KN.findall(seg(shokai, "H1('第４　"))]

# ── ③ 区分文書 ──
HAN = {'Ａ': ("H1('第２　Ａ", "H1('第３　Ｂ"), 'Ｂ': ("H1('第３　Ｂ", "H1('第４　Ｃ"),
       'Ｃ': ("H1('第４　Ｃ", "H1('第５　Ｄ"), 'Ｄ': ("H1('第５　Ｄ", "H2('ご依頼のまとめ方")}
SR = {k: nos_in(seg(shiryo, a, b)) for k, (a, b) in HAN.items()}

bad = []


def chk(na, ok, msg):
    print(('  ○ ' if ok else '  × ') + na + '　… ' + msg)
    if not ok:
        bad.append(na)


ZEN = list(range(1, NN + 1))
kaketa = lambda ns: sorted(set(ZEN) - set(ns)) or 'なし'
kasanaru = lambda ns: sorted({n for n in ns if ns.count(n) > 1}) or 'なし'

print('■ 金ケ崎町　確認事項の突合せ\n')

chk('①素案の附が1〜%dで欠番・重複なし' % NN, sorted(FU) == ZEN,
    '%d件／欠番 %s／重複 %s' % (len(FU), kaketa(FU), kasanaru(FU)))

chk('②素案の附が番号の順に並んでいる', FU == sorted(FU),
    '順に並んでいる' if FU == sorted(FU) else '並びが崩れている：' + '・'.join(
        '%dの次に%d' % (a, b) for a, b in zip(FU, FU[1:]) if b < a))

SA = [n for _k, n in SH3 + SH4]
chk('③照会文書の第3＋第4が%d件で欠番・重複なし' % NN, sorted(SA) == ZEN,
    '第3 %d件＋第4 %d件＝%d件／欠番 %s／重複 %s'
    % (len(SH3), len(SH4), len(SA), kaketa(SA), kasanaru(SA)))

RA = sum(SR.values(), [])
chk('④区分文書のＡ〜Ｄが%d件で欠番・重複なし' % NN, sorted(RA) == ZEN,
    '・'.join('%s%d' % (k, len(v)) for k, v in SR.items())
    + '＝%d件／欠番 %s／重複 %s' % (len(RA), kaketa(RA), kasanaru(RA)))

K = {n: k for k, n in SH3 + SH4}
S = {n: k for k, ns in SR.items() for n in ns}
chigau = sorted(n for n in set(K) & set(S) if K[n] != S[n])
chk('⑤照会文書と区分文書で区分が一致する', not chigau, '食い違い なし' if not chigau
    else '／'.join('№%d 照会%s・区分%s' % (n, K[n], S[n]) for n in chigau))

m3 = sorted(n for k, n in SH3 if k not in 'ＣＤ')
m4 = sorted(n for k, n in SH4 if k not in 'ＡＢ')
chk('⑥第3はＣ・Ｄだけ、第4はＡ・Ｂだけ', not (m3 or m4),
    '混在なし' if not (m3 or m4) else '第3に %s／第4に %s' % (m3, m4))

e = []
t3 = seg(shokai, "H1('第３　", "H1('第４　")
hs = list(re.finditer(r"H2\('([^']*?)（(\d+)件）'\)", t3))
for a, b in zip(hs, hs[1:] + [None]):
    part = t3[a.end():(b.start() if b else len(t3))]
    j = len(KN.findall(part))
    if j != int(a.group(2)):
        e.append('%s：宣言%s件・実際%d件' % (a.group(1), a.group(2), j))
chk('⑦照会文書の分野ごとの件数が行数と合う', not e,
    '%d分野すべて一致' % len(hs) if not e else '／'.join(e))

e2 = []
for k, (a, b) in HAN.items():
    m = re.search(r'（(\d+)件）', seg(shiryo, a, b).split('\n')[0])
    if m and int(m.group(1)) != len(SR[k]):
        e2.append('%s：見出し%s件・実際%d件' % (k, m.group(1), len(SR[k])))
tB = seg(shiryo, *HAN['Ｂ'])
hs2 = list(re.finditer(r"H2\('([^']*?)（(\d+)件）'\)", tB))
for a, b in zip(hs2, hs2[1:] + [None]):
    part = tB[a.end():(b.start() if b else len(tB))]
    j = len(nos_in(part))
    if j != int(a.group(2)):
        e2.append('Ｂの%s：宣言%s件・実際%d件' % (a.group(1), a.group(2), j))
chk('⑧区分文書の見出し・小見出しの件数が行数と合う', not e2,
    # 件数はじか書きしない。区分を動かすと成功の知らせだけ古くなるためである
    '・'.join('%s%d' % (k, len(v)) for k, v in SR.items())
    + ' と小見出し%d つが一致' % len(hs2) if not e2 else '／'.join(e2))

# 件数はじか書きしない。《要確定》を増やすと点検だけ古くなるためである
YH = re.search(r"H2\('《要確定》　町のご判断をいただく事項（(\d+)件）'\)", soan)
YN = int(YH.group(1)) if YH else 0
YK = [int(m) for m in re.findall(
    r"\['\*\*(\d+)\*\*', ", seg(soan, YH.group(0) if YH else '', "H2('《要回答》"))]
chk('⑨《要確定》が1〜%d で欠番・重複なく順に並ぶ' % YN, YN and YK == list(range(1, YN + 1)),
    '%d件／%s' % (len(YK), '順に並んでいる' if YK == sorted(YK) else '並びが崩れている'))

# ── 素案0-1の記号の件数 ──
HON = soan[:soan.index("H1('附　未確定事項の一覧')")]
sen = dict(re.findall(r"\['\*\*(【　】|【案】)\*\*'[^\]]*?\*\*(\d+)か所\*\*'\]", soan))
e3 = []
for k, v in sen.items():
    j = len(re.findall(re.escape(k), HON))
    if j != int(v):
        e3.append('%s：0-1は%sか所・本文は%dか所' % (k, v, j))
chk('⑩素案0-1の記号の件数が本文と合う', not e3 and len(sen) == 2,
    '【　】%sか所・【案】%sか所 が本文と一致' % (sen.get('【　】'), sen.get('【案】'))
    if not e3 and len(sen) == 2 else '／'.join(e3) or '0-1の表を読めない')

# ── Ａ（照会を取り下げた）の番号が本文に《要回答》として残っていないこと ──
#    Ａは「解決済み。照会を取り下げます」であり、本文が町の回答を待ち続けては
#    辻褄が合わない。Ｂ（当方で一次資料を確認する）は確認が済むまで残ってよい。
nokori = {n: len(re.findall('《要回答%d》' % n, HON)) for n in SR['Ａ']}
nokori = {n: c for n, c in nokori.items() if c}
chk('⑪Ａ（照会を取り下げた）が本文に《要回答》として残っていない', not nokori,
    'Ａ%s件とも残っていない' % len(SR['Ａ']) if not nokori
    else '／'.join('№%d が%dか所' % (n, c) for n, c in sorted(nokori.items())))

# ── 《要確定》と照会文書 第2の対応 ──
#    素案の附《要確定》と照会文書の第2「ご判断をいただきたい事項」は同じ番号で
#    同じ事項を扱う。片方だけに増やすと、町は判断を求められていない事項を
#    素案で「町のご判断」と読むことになる（令和8年10月20日、№14〜17の欠落で起きた）。
DH = re.search(r"H1\('第２　ご判断をいただきたい事項（(\d+)件）'\)", shokai)
DN = int(DH.group(1)) if DH else 0
DK = [int(m) for m in re.findall(
    r"\['\*\*(\d+)\*\*', ", seg(shokai, DH.group(0) if DH else '', "H1('第３　"))] if DH else []
chk('⑫照会文書 第2が素案の《要確定》と同じ番号の組である', DH and DK == YK and DN == YN,
    '第2 %d件（見出し%d件）・《要確定》%d件 が一致' % (len(DK), DN, YN) if DH and DK == YK and DN == YN
    else '第2 %d件（見出し%d件）／《要確定》%d件／第2にない№ %s／附にない№ %s' % (
        len(DK), DN, YN, sorted(set(YK) - set(DK)) or 'なし', sorted(set(DK) - set(YK)) or 'なし'))

print('\n%s：確認事項%d件（Ａ%d・Ｂ%d・Ｃ%d・Ｄ%d）／《要確定》%d件'
      % ('適合' if not bad else '不適合 %d件' % len(bad), NN,
         len(SR['Ａ']), len(SR['Ｂ']), len(SR['Ｃ']), len(SR['Ｄ']), YN))
sys.exit(1 if bad else 0)
