# -*- coding: utf-8 -*-
"""目次のページ番号を組版結果から作る。

　手順
　　1　mkf_soan.py でビルドする（ページ番号が無い状態では目次に「－」が出る）
　　2　LibreOffice で PDF にする
　　3　各見出しが何ページ目に現れるかを PDF から読む
　　4　out/_toc_pages.json に「ブックマーク名 → ページ番号」を書く
　　5　もう一度ビルドすると、目次にページ番号が入る

　ページ番号は本文の先頭を1とする（表紙・目次には番号を出さないため）。
　Word で開いた時点で PAGEREF フィールドが更新されるので、ここで入れる値は
　更新前に表示される値である。
"""
import json
import os
import re
import subprocess
import sys

DOCX = sys.argv[1] if len(sys.argv) > 1 else 'out/計画素案_第10期.docx'
BUILD = sys.argv[2] if len(sys.argv) > 2 else 'mkf_soan.py'
PAGEMAP = 'out/_toc_pages.json'


def build():
    subprocess.run([sys.executable, BUILD], check=True)


def to_pdf(path, outdir='out'):
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf',
                    '--outdir', outdir, path],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return os.path.join(outdir, os.path.splitext(os.path.basename(path))[0] + '.pdf')


def headings_of(build_py):
    """ビルドして HEADINGS を取り出す（level, text, bookmark名）。"""
    g = {'__name__': '__mkf__'}
    src = open(build_py, encoding='utf-8').read()
    # doc.save を無効にして実行する
    src = re.sub(r'^doc\.save\(.*\)$', 'pass', src, flags=re.M)
    exec(compile(src, build_py, 'exec'), g)
    return g['HEADINGS']


def norm(s):
    return re.sub(r'\s+', '', s)


def pages_of(pdf, heads):
    """各見出しのページ番号を、PDFのフッタ「- n -」から読み取る。

    　表紙・目次にはフッタの番号を出していないため、番号のあるページが本文である。
    　目次のページにも見出しの文字列が並ぶので、本文の範囲だけを探す。
    """
    import fitz
    d = fitz.open(pdf)
    txt, shown = [], []
    for p in d:
        raw = p.get_text()
        txt.append(norm(raw))
        m = [x for x in re.findall(r'(?m)^\s*-\s*(\d+)\s*-\s*$', raw)]
        shown.append(int(m[-1]) if m else None)
    base = next((i for i, v in enumerate(shown) if v is not None), 0)
    out = {}
    at = base
    for lv, text, name in heads:
        key = norm(text)
        found = None
        for i in range(at, len(txt)):
            if key in txt[i]:
                found = i
                break
        if found is None:                 # 行が折り返されている場合は前方一致で探す
            short = key[:12]
            for i in range(at, len(txt)):
                if short in txt[i]:
                    found = i
                    break
        if found is None:
            found = at
        out[name] = shown[found] if shown[found] is not None else found - base + 1
        at = found
    return out, len(d), base


if __name__ == '__main__':
    heads = headings_of(BUILD)
    if os.path.exists(PAGEMAP):
        os.remove(PAGEMAP)
    build()
    pdf = to_pdf(DOCX)
    pages, npage, base = pages_of(pdf, heads)
    with open(PAGEMAP, 'w', encoding='utf-8') as f:
        json.dump(pages, f, ensure_ascii=False, indent=1)
    build()
    pdf = to_pdf(DOCX)
    import fitz
    print('見出し %d　PDF %dページ（前付 %dページ・本文 %dページ）'
          % (len(heads), npage, base, npage - base))
    print('最終の見出しのページ', max(pages.values()))
