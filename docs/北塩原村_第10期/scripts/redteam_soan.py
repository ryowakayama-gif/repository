# -*- coding: utf-8 -*-
"""素案・委員会資料のレッドチーム走査。
大雪の協議用素案の点検で見つかった誤りの型を当村の成果品に当てる。
(a) 日本語でない文字の混入 (b) 受託者を主語とする語・内部の作業物の名
(c) 禁止表現 (d) 個人情報の形 (e) ヘッダー・フッター・ページ番号
"""
import sys, io, re, os
sys.dont_write_bytecode = True

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TARGETS = [('素案', os.path.join(BASE, '18_計画素案.md'))]
# 書き下ろした説明文はどこにでも混入する。確認事項の本文も走査に含める
SRC_TARGETS = [('確認事項', os.path.join(BASE, 'scripts', 'wbs_kakunin.py'))]
BUILDER = os.path.join(BASE, 'scripts', 'build_soan_docx.js')

# (a) 日本語・記号として通す文字。これ以外が出たら混入とみなす
OK = ('ぁ-んァ-ヴ一-龥々ーヶ〆' 'ａ-ｚＡ-Ｚ０-９' '0-9A-Za-z'
      '、。・（）「」『』【】〔〕［］〈〉《》〜～％% 　\n\t'
      r'\-＋+±×÷=＝<>＜＞!！?？:：;；,，.．/／\\|｜'
      r'#＃$＄&＆*＊@＠^＾_＿`｀~‾"\'’‘”“()\[\]{}'
      '§¶†‡°′″№㎡㎞㎏ℓ①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳²³'
      '⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ'
      '─━│┃├┤┬┴┼╴▲▼△▽◆◇■□●○◎★☆※→←↑↓⇒⇔═╔╗╚╝╠╣╦╩╬'
      '⚙'
      # この文書が意図して用いる記号。掲げておかないと走査が雑音で埋まる
      '―'   # 該当がないことを示すダッシュ（U+2015）
      'α'   # 第1号被保険者負担割合の記号
      '－'   # 全角の減算記号（＋＝×÷と揃える）
      'ⅰ')  # 国の評価指標の目標Ⅲ（ⅰ）の表記
BAD_CHAR = re.compile('[^' + OK + ']')

# (b) 受託者を主語とする語・内部の作業物の名
JUCHU = ['受託者', '本素案', '協議用', '修正指示書', '確認事項No', '固定値',
         '実物から', '再実行', '章節ごとに', '判定しています', 'スクリプト',
         '弊社', '当社', 'ビズアップ', '再出力', '依頼中', '納品',
         '.py', '.js', '.json', 'runpy', '推計の案']

# (c) 禁止表現
KINSHI = ['に由来する', '1件も', '一件も', '有意差がないため', '全国トップ級',
          'と整合する']

# (d) 個人情報の形
PII = [(r'0\d{1,4}-\d{1,4}-\d{3,4}', '電話番号の形'),
       (r'[\w.+-]+@[\w.-]+\.\w{2,}', 'メールアドレスの形'),
       (r'\d{4}年\d{1,2}月\d{1,2}日生', '生年月日の形'),
       (r'被保険者番号[：:\s]*\d', '被保険者番号')]


def scan(name, path):
    s = io.open(path, encoding='utf-8').read()
    lines = s.split('\n')
    out = []

    # (a)
    hits = []
    for i, ln in enumerate(lines, 1):
        for m in BAD_CHAR.finditer(ln):
            hits.append((i, m.group(), repr(m.group()), ln.strip()[:60]))
    out.append(('a 日本語でない文字', hits))

    # (b)
    hits = [(i, w, '', ln.strip()[:70])
            for i, ln in enumerate(lines, 1) for w in JUCHU if w in ln]
    out.append(('b 受託者を主語とする語・内部の作業物', hits))

    # (c)
    hits = [(i, w, '', ln.strip()[:70])
            for i, ln in enumerate(lines, 1) for w in KINSHI if w in ln]
    out.append(('c 禁止表現', hits))

    # (d)
    hits = []
    for i, ln in enumerate(lines, 1):
        for pat, lab in PII:
            for m in re.finditer(pat, ln):
                hits.append((i, lab, m.group(), ln.strip()[:60]))
    out.append(('d 個人情報の形', hits))

    # (f) 半角・全角のゆれ。少数側を混在とみなして掲げる
    PAIRS = [('%', '％'), ('〜', '～'), (',', '，'), ('.', '．'), (':', '：')]
    hits = []
    for a, b in PAIRS:
        na, nb = s.count(a), s.count(b)
        if na and nb:
            minor = a if na < nb else b
            for i, ln in enumerate(lines, 1):
                if minor in ln:
                    hits.append((i, '%r が %d件 / %r が %d件' % (a, na, b, nb),
                                 repr(minor), ln.strip()[:60]))
    out.append(('f 半角・全角のゆれ', hits))
    return out


def scan_builder():
    """(e) 節の区切りを落としてヘッダー・フッター・ページ番号が失われていないか。"""
    s = io.open(BUILDER, encoding='utf-8').read()
    need = [('ヘッダー', 'headers:'), ('フッター', 'footers:'),
            ('ページ番号', 'PageNumber.CURRENT'),
            ('節の区切り', 'sections: [{'), ('表頭の繰り返し', 'tableHeader: true')]
    return [(0, lab, '', key) for lab, key in need if key not in s]


def selftest():
    """走査が現に検出するかを確かめる。検出しない走査は点検の意味がない。"""
    import tempfile
    cases = [('日本語でない文字', 'Привет　高齢者', 'a'),
             ('受託者を主語とする語', '受託者が判定しています。', 'b'),
             ('禁止表現', '1件も見当たりません。', 'c'),
             ('電話番号の形', '連絡先は0241-23-4567です。', 'd')]
    bad = []
    for lab, text, ax in cases:
        fd, path = tempfile.mkstemp(suffix='.md')
        os.close(fd)
        io.open(path, 'w', encoding='utf-8').write(text)
        got = {k[0]: len(v) for k, v in scan('自己試験', path)}
        os.unlink(path)
        if not got.get(ax):
            bad.append(lab)
    return bad


def run():
    ng = 0
    bad = selftest()
    print('■ 走査の自己試験', '適合（4件すべて検出）' if not bad
          else '不適合：検出しなかった走査 ' + '／'.join(bad))
    ng += len(bad)
    for name, path in TARGETS:
        print('■', name, path)
        for label, hits in scan(name, path):
            mark = '適合' if not hits else '要確認 %d件' % len(hits)
            print('  %-38s %s' % (label, mark))
            for h in hits[:40]:
                print('      %5d  %s  %s  | %s' % h)
            if hits:
                ng += len(hits)
    # 書き下ろした説明文（確認事項）も a と c と d で走査する
    for name, path in SRC_TARGETS:
        print('■', name, path)
        for label, hits in scan(name, path):
            if label.startswith('b') or label.startswith('f'):
                continue          # 作業物の名と記号のゆれはスクリプトでは当然に出る
            mark = '適合' if not hits else '要確認 %d件' % len(hits)
            print('  %-38s %s' % (label, mark))
            for h in hits[:20]:
                print('      %5d  %s  %s  | %s' % h)
            if hits:
                ng += len(hits)

    hits = scan_builder()
    print('  %-38s %s' % ('e ヘッダー・フッター・ページ番号',
                          '適合' if not hits else '不適合 %d件' % len(hits)))
    for h in hits:
        print('      %5d  %s  %s  | %s' % h)
    ng += len(hits)
    return ng


if __name__ == '__main__':
    n = run()
    print('\n要確認 合計 %d件' % n)
