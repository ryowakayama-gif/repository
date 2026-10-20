# -*- coding: utf-8 -*-
"""医療情報ネット（ナビイ）を2つのキーワードで検索し、件数と施設の一覧を出す。

    python3 -I navii.py <診療科目> <所在地> [保存するHTMLのファイル名]

　画面はスクリプトで動くが、中身は3つのリクエストで取れる。
　　1　S2300/initialize　　セッションを張る
　　2　S2300/iryoSearch　　keywordType（3＝診療科目／4＝所在地）とkeywordを送る。検索IDが返る
　　　　2つ目以降のキーワードは、同じ検索IDを付けて送る
　　3　S2400/initialize?id=<検索ID>　　結果の一覧
"""
import http.cookiejar
import json
import re
import sys
import urllib.parse
import urllib.request

B = 'https://www.iryou.teikyouseido.mhlw.go.jp/znk-web/juminkanja'
UA = ('Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) '
      'Chrome/120.0.0.0 Safari/537.36')

cj = http.cookiejar.CookieJar()
op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
op.addheaders = [('User-Agent', UA), ('Referer', B + '/S2300/initialize')]


def get(url):
    with op.open(url, timeout=90) as r:
        return r.read().decode('utf-8', errors='replace')


def search(kamoku, basho):
    get(B + '/S2300/initialize')
    q = urllib.parse.quote
    first = ('%s/S2300/iryoSearch?XCHARSET=utf-8&XPARAM=keyword&iyakuKbn=1&lang='
             '&keywordType=3&keyword=%s' % (B, q(kamoku)))
    sid = json.loads(get(first))['result']['id']
    if basho:
        get('%s/S2300/iryoSearch?XCHARSET=utf-8&XPARAM=keyword&keywordType=4&keyword=%s&id=%s'
            % (B, q(basho), sid))
    return get('%s/S2400/initialize?id=%s' % (B, sid))


def parse(html):
    m = re.search(r'が<span class="num">([0-9,]+)</span>件', html)
    if m:
        cnt = m.group(1)
    else:
        m0 = re.search(r'([0-9,]+)\s*件', html)          # 0件のときは文言が変わる
        cnt = m0.group(1) if m0 else '?'
    items = []
    for blk in re.split(r'<div class="item">', html)[1:]:
        nm = re.search(r'<h2 class="name">\s*<a[^>]*>(.*?)</a>', blk, re.S)
        ad = re.search(r'alt="住所".*?<dd>\s*<p>(.*?)<a', blk, re.S)
        km = re.search(r'alt="診療科目".*?<dd>(.*?)</dd>', blk, re.S)
        strip = lambda t: re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', t)).strip()
        items.append((strip(nm.group(1)) if nm else '',
                      strip(ad.group(1)) if ad else '',
                      strip(km.group(1)) if km else ''))
    return cnt, items


if __name__ == '__main__':
    kamoku, basho = sys.argv[1], (sys.argv[2] if len(sys.argv) > 2 else '')
    html = search(kamoku, basho)
    if len(sys.argv) > 3:
        open(sys.argv[3], 'w', encoding='utf-8').write(html)
    cnt, items = parse(html)
    print('【%s × %s】　%s件' % (kamoku, basho or '（全国）', cnt))
    for n, a, k in items:
        print('  ・%s' % n)
        if a:
            print('      %s' % a)
        if k:
            print('      診療科目　%s' % k[:110])
