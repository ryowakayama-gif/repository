# -*- coding: utf-8 -*-
"""レンダリング結果から各ブックマークの実ページを求め、plan.py の目次キャッシュ値を更新する。"""
import re, subprocess, os, sys
from pypdf import PdfReader

def norm(s):
    return re.sub(r'\s+', '', s)

def render(docx, pdf):
    env = dict(os.environ, HOME=os.environ.get('SOFFICE_HOME', os.path.expanduser('~')))
    if os.path.exists(pdf):
        os.remove(pdf)
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', '.', docx],
                   check=True, capture_output=True, env=env)

def toc_pages(pdf):
    """目次ページ(1ページ目)に描画された番号を、LibreOffice が解決した実ページ番号として読む。"""
    r = PdfReader(pdf)
    out = []
    for line in (r.pages[0].extract_text() or '').split('\n'):
        m = re.match(r'^(.*?)\.{3,}\s*(\d+)\s*$', line.strip())
        if m:
            out.append((norm(m.group(1)), int(m.group(2))))
    return out

if __name__ == '__main__':
    render('restructured.docx', 'restructured.pdf')
    pages = dict(toc_pages('restructured.pdf'))
    src = open('plan.py', encoding='utf-8').read()
    changed = 0
    def repl(m):
        global changed
        label, name, old = m.group(1), m.group(2), int(m.group(3))
        new = pages.get(norm(label.replace('　', '')), pages.get(norm(label)))
        if new is None:
            print('  !! ページ未検出:', label); return m.group(0)
        if new != old:
            changed += 1
        return f'("{label}", "{name}", {new})'
    src = re.sub(r'\("([^"]+)", "(\w+)", (\d+)\)', repl, src)
    open('plan.py', 'w', encoding='utf-8').write(src)
    print(f'目次キャッシュ値を更新: {changed}件')
