# -*- coding: utf-8 -*-
"""体裁のプロファイル　―　第1回策定委員会資料の様式（町が作成・令和8年10月14日）

　第2回以降の策定委員会資料・次第・想定問答集に用いる。fmt.py の前に連結する。
　　{ cat style_iinkai1.py; cat fmt.py; echo; cat fmt_iinkai1.py; echo; cat ch_XXX.py; } > mkf_XXX.py

　第1回資料（受領した docx）から実測した値
　　用紙・余白　　A4縦　上・左・右2.0cm　下1.8cm
　　書体　　　　　BIZ UDPゴシック。文字色 #1B3A2F
　　節の見出し　　塗り #2C6E52・白文字18pt太字（「１　　題 ― 結論」の形）
　　小見出し　　　「■　」＋14pt太字 #2C6E52
　　本文　　　　　12pt（箇条は「・」で始め、ぶら下げ）
　　表　　　　　　見出し行 塗り #2C6E52・白文字太字／本文11pt／罫線は黒の細線
　　囲み　　　　　ご意見をいただきたい点＝塗り #FBF0E6・枠 #EBC9AB・題 #B85C22
　　　　　　　　　読み取り・考え方　　　＝塗り #EAF2ED・枠 #CBD8D1・題 #2C6E52
　　ヘッダ　　　　なし。フッタにページ番号
"""
STYLE = {
    'FONT': 'BIZ UDPゴシック',
    'GFONT': 'BIZ UDPゴシック',
    'BODY_PT': 11.5,
    'TBL_PT': 10,
    'SRC_PT': 8.5,
    'H1': (18, '2C6E52', None),
    'H2': (14, '2C6E52', None),
    'H3': (12, '2C6E52', None),
    'MARGIN': (2.0, 1.8, 2.0, 2.0),
    'HDRFT': (1.25, 1.0),
    'LINE': 1.08,
    'AFTER': 3,
    'TEXTW': 17.0,
    'HEADER': '',
    'PAGENO': '- {} -',
    'PAGENO_PT': 10,
    'TBLFILL': '2C6E52',
    'BORDER': (('top', '4', '000000'), ('left', '4', '000000'),
               ('bottom', '4', '000000'), ('right', '4', '000000'),
               ('insideH', '4', '000000'), ('insideV', '4', '000000')),
    'FIGCAP': (11, '2C6E52'),
    'TOC_NOTE': '',
}
