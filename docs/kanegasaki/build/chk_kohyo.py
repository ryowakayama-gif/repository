# -*- coding: utf-8 -*-
"""公表用素案の数値が、協議用の計画素案に現れるものだけであることを照合する。

　公表用は協議用から内容を採って書き直したものである。書き直しの際に
　数値を取り違えると、2本の成果品で数が食い違う。これを機械で防ぐ。

　使い方　python3 -I chk_kohyo.py
"""
import io
import re
import sys

from docx import Document


def alltext_docx(path):
    d = Document(path)
    t = [p.text for p in d.paragraphs]
    for tb in d.tables:
        for r in tb.rows:
            for c in r.cells:
                t.append(c.text)
    return '\n'.join(t)


SOAN = alltext_docx('out/計画素案_第10期.docx')
KOHYO = alltext_docx('out/計画素案_第10期_公表用.docx')

# 数値の形　1,234 / 12.3 / 43 ／ 単位つきのものも含める
NUM = re.compile(r'[0-9][0-9,]*(?:\.[0-9]+)?')

# 照合の対象外
#   年・月・日・期・章・節・号・項・条・ページ・記号の番号は体系の番号であり数値ではない
SKIP_CTX = re.compile(
    r'(令和|平成|昭和|第[0-9]+期|第[0-9]+章|[0-9]+-[0-9]|第[0-9]+条|第[0-9]+項|第[0-9]+号|'
    r'ＳＴＥＰ|STEP|年度|月|日|位／|／[0-9])'
)

def numbers(s):
    out = set()
    for m in NUM.finditer(s):
        v = m.group(0)
        if len(v.replace(',', '').replace('.', '')) <= 1:
            continue          # 1桁は番号の可能性が高い
        head = s[max(0, m.start() - 6):m.start()]
        tail = s[m.end():m.end() + 4]
        if SKIP_CTX.search(head + v + tail):
            continue
        out.add(v)
    return out


soan = numbers(SOAN)
kohyo = numbers(KOHYO)
only = sorted(kohyo - soan, key=lambda x: (-len(x), x))

print('協議用の数値 %d 種　公表用の数値 %d 種' % (len(soan), len(kohyo)))
if not only:
    print('\n公表用だけに現れる数値　なし')
    sys.exit(0)

print('\n公表用だけに現れる数値　%d 件' % len(only))
for v in only:
    # 文脈を1つ出す
    i = KOHYO.find(v)
    ctx = KOHYO[max(0, i - 50):i + 30].replace('\n', ' ')
    print('  %-10s …%s…' % (v, ctx))
print('\n　協議用に無い数値は、取り違えか、公表用で新たに置いた数値である。')
print('　新たに置いたものであれば、協議用の側にも入れるか、ここに掲げて記録に残す。')
sys.exit(1)
