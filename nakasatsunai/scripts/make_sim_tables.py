# -*- coding: utf-8 -*-
"""協議会案を反映したシミュレーション表を、本文貼込用のPNGとして生成する。

 ・sim_pattern.png … 改定パターン一覧（本文 p.27）
 ・sim_kensho.png  … シミュレーションパターン検証（本文 p.28）

既存のEMF（Excel貼込）と同じ構成・配色に合わせる。数値は
シミュレーションExcelから抽出した sim_numbers.json を唯一の出所とする。
"""
import json, os, sys
from PIL import Image, ImageDraw, ImageFont

FONT_PATH = '/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf'
OUT = sys.argv[2] if len(sys.argv) > 2 else 'img'
os.makedirs(OUT, exist_ok=True)
NUM = json.load(open(sys.argv[1] if len(sys.argv) > 1 else 'sim_numbers.json', encoding='utf-8'))

DPI    = 220
HEADBG = '#7E9BB2'      # ヘッダー帯（本文中の既存表と同色）
TITLEBG = '#E3EAEF'     # 区分名の帯
ALTBG  = '#F2F5F7'
LINE   = '#9AA9B4'
GRID   = '#C9D3DA'
TXT    = '#333333'
HEADTX = '#FFFFFF'
EMPH   = '#1F4E66'      # 採用案の強調色

F = lambda s: ImageFont.truetype(FONT_PATH, int(s / 72 * DPI))
YEARS = [f'令和{n}年度' for n in range(7, 17)]


def draw_text(d, box, text, font, fill, align='center'):
    x0, y0, x1, y1 = box
    l, t, r, b = d.textbbox((0, 0), text, font=font)
    w, h = r - l, b - t
    if align == 'center':
        x = x0 + (x1 - x0 - w) / 2 - l
    elif align == 'left':
        x = x0 + 0.05 * DPI - l
    else:
        x = x1 - 0.05 * DPI - w - l
    d.text((x, y0 + (y1 - y0 - h) / 2 - t), text, font=font, fill=fill)


def table(path, w_in, rows, colw, rowh, fonts, note=None):
    """rows: [(種別, [セル…]), …]  種別は 'head'|'title'|'data'|'alt'|'emph'"""
    nh = sum(rowh)
    pad = 0.04 * DPI
    note_h = (0.17 * DPI * (1 + note.count('\n'))) if note else 0
    W = int(w_in * DPI)
    H = int(nh + pad * 2 + note_h)
    img = Image.new('RGB', (W, H), 'white')
    d = ImageDraw.Draw(img)
    xs = [pad]
    for c in colw:
        xs.append(xs[-1] + c * (W - 2 * pad) / sum(colw))
    y = pad
    for (kind, cells), h in zip(rows, rowh):
        bg = {'head': HEADBG, 'title': TITLEBG, 'alt': ALTBG, 'text': 'white',
              'data': 'white', 'emph': '#EAF2F6'}[kind]
        d.rectangle([xs[0], y, xs[-1], y + h], fill=bg)
        fg = HEADTX if kind == 'head' else (EMPH if kind == 'emph' else TXT)
        fnt = fonts['head'] if kind in ('head', 'title') else fonts['body']
        for i, cell in enumerate(cells):
            if cell is None:
                continue
            al = 'left' if i == 0 else 'right' if kind in ('data', 'alt', 'emph') else 'center'
            draw_text(d, (xs[i], y, xs[i + 1], y + h), str(cell), fnt, fg, al)
            if i:
                d.line([xs[i], y, xs[i], y + h], fill=GRID, width=1)
        d.line([xs[0], y + h, xs[-1], y + h], fill=LINE if kind == 'head' else GRID, width=1)
        y += h
    d.rectangle([xs[0], pad, xs[-1], y], outline=LINE, width=2)
    if note:
        draw_text(d, (xs[0], y, xs[-1], H - 1), note, fonts['note'], '#555555', 'left')
    img.save(path, dpi=(DPI, DPI))
    print(f'  {path}: {img.width}x{img.height}px → {img.width/DPI:.2f}×{img.height/DPI:.2f}in')
    return img


# ═══ 表① 改定パターン一覧 ═══════════════════════════════════════════
K  = NUM['採用案：協議会案']
P1 = NUM['（参考）一律10％']
P3 = NUM['（参考）一律20％']
P4 = NUM['（参考）一律30％']
fonts1 = {'head': F(8.5), 'body': F(8.5), 'note': F(7)}
rows1 = [
    ('head',  ['区　分', '採用案：協議会案', '（参考）一律10％', '（参考）一律20％', '（参考）一律30％']),
    ('text',  ['改定の方法', '一般改定＋水質使用料', '単価に一律10％', '単価に一律20％', '単価に一律30％']),
    ('alt',   ['改定後の平均単価（円／m³）', f'{K["単価"]:.2f}', f'{P1["単価"]:.2f}',
               f'{P3["単価"]:.2f}', f'{P4["単価"]:.2f}']),
    ('data',  ['平均改定額（円／m³）', f'＋{K["改定額"]:.1f}', f'＋{P1["改定額"]:.1f}',
               f'＋{P3["改定額"]:.1f}', f'＋{P4["改定額"]:.1f}']),
    ('emph',  ['引上げ率（現行比）', f'＋{K["改定率"]*100:.1f}％', f'＋{P1["改定率"]*100:.1f}％',
               f'＋{P3["改定率"]*100:.1f}％', f'＋{P4["改定率"]*100:.1f}％']),
]
rowh1 = [0.26 * DPI] + [0.24 * DPI] * 4
table(f'{OUT}/sim_pattern.png', 6.13, rows1, [1.55, 1.15, 1.0, 1.0, 1.0], rowh1, fonts1,
      note='※採用案の引上げ率は平均値。協議会資料では全体改定率120.9％（改定後÷現行）と表記している。')

# ═══ 表② シミュレーションパターン検証 ═══════════════════════════════
fonts2 = {'head': F(6.6), 'body': F(6.6), 'note': F(6)}
pct = lambda v: f'{v*100:.1f}％'
sen = lambda v: f'{v:,.0f}'
SCN = [('現状予測（現行料金継続）', NUM['現状予測（現行料金継続）'], False),
       ('採用案：協議会案（一般改定100円案＋水質使用料）', K, True),
       ('（参考）シミュレーションパターン１（一律10％）', P1, False),
       ('（参考）シミュレーションパターン３（一律20％）', P3, False),
       ('（参考）シミュレーションパターン４（一律30％）', P4, False)]
rows2, rowh2 = [], []
RH, HH = 0.165 * DPI, 0.175 * DPI
for name, rec, is_main in SCN:
    cap = (f'{name}　［単価 {rec["単価"]:.2f}円／改定額 ＋{rec["改定額"]:.1f}円'
           f'／引上げ率 ＋{rec["改定率"]*100:.1f}％］' if rec['改定額'] > 0
           else f'{name}　［単価 {rec["単価"]:.2f}円／改定なし］')
    rows2.append(('title', [cap] + [None] * 10));                      rowh2.append(HH)
    rows2.append(('head',  ['項　目'] + YEARS));                       rowh2.append(HH)
    rows2.append(('data',  ['使用料収入（千円）'] + [sen(v) for v in rec['使用料収入']])); rowh2.append(RH)
    if is_main:
        rows2.append(('emph', ['うち改定による増収（千円）']
                      + [sen(v) for v in NUM['内訳']['増収計']]));     rowh2.append(RH)
    rows2.append(('alt',   ['経常収支比率'] + [pct(v) for v in rec['経常収支比率']])); rowh2.append(RH)
    rows2.append(('data',  ['経費回収率'] + [pct(v) for v in rec['経費回収率']]));     rowh2.append(RH)
table(f'{OUT}/sim_kensho.png', 6.54, rows2, [1.72] + [1.0] * 10, rowh2, fonts2,
      note='※経費回収率は総務省の経営比較分析表の定義（使用料収入÷汚水処理費。資本費を含む）。\n'
           '※採用案の増収は、一般改定（基本＋100円／月・超過＋10円／m³）と水質使用料（工場排水70円／m³）の合計。')
print('完了')
