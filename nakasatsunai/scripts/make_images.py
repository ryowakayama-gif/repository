# -*- coding: utf-8 -*-
"""提供資料から、Word本文に差し込む図表画像を生成する。"""
import io, os, sys
import pymupdf
from PIL import Image, ImageDraw, ImageFont
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager

MAT = 'mat'
OUT = 'img'
os.makedirs(OUT, exist_ok=True)

FONT_PATH = '/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf'
font_manager.fontManager.addfont(FONT_PATH)
JP = font_manager.FontProperties(fname=FONT_PATH).get_name()
plt.rcParams['font.family'] = JP
plt.rcParams['axes.unicode_minus'] = False

DPI = 220
BAR = '#7190A8'          # 本文中の既存グラフに合わせた青灰
GRID = '#D7DCDE'
TXT = '#333333'
HEADBG = '#7E9BB2'       # 表のヘッダー帯
ROWBG = '#FFFFFF'
ALTBG = '#F2F5F7'
LINE = '#9AA9B4'


def pdf_to_png(src, out, width_in, dpi=DPI):
    d = pymupdf.open(src)
    pg = d[0]
    scale = (width_in * dpi) / (pg.rect.width / 72 * dpi) * dpi / 72
    pix = pg.get_pixmap(matrix=pymupdf.Matrix(scale, scale))
    im = Image.open(io.BytesIO(pix.tobytes('png'))).convert('RGB')
    im.save(out, dpi=(dpi, dpi))
    return im.size


# ── 1. 下水道処理区域図 ───────────────────────────────────────────────
w, h = pdf_to_png(f'{MAT}/区域.pdf', f'{OUT}/gesui_kuiki.png', 6.4)
print(f'下水区域図: {w}x{h}px  → {w/DPI:.2f}x{h/DPI:.2f}in')

# ── 2. 簡易水道給水区域図（xlsx内のJPEGを取り出し済みのものを使う）─────────
src = 'new/kansui_map_full.png'
im = Image.open(src).convert('RGB')
tw = int(6.4 * DPI)
im2 = im.resize((tw, int(im.height * tw / im.width)), Image.LANCZOS)
im2.save(f'{OUT}/kansui_kuiki.png', dpi=(DPI, DPI))
print(f'簡水区域図: {im2.width}x{im2.height}px  → {im2.width/DPI:.2f}x{im2.height/DPI:.2f}in')

# ── 3. 公共下水道事業 各年度整備延長グラフ ────────────────────────────
# 出典：提供資料「管渠延長.pdf」（年度別・管径別管渠延長）
DATA = [
    ('H4', 615.43), ('H5', 4343.51), ('H6', 2451.72), ('H7', 3066.37),
    ('H8', 2555.96), ('H9', 2287.57), ('H10', 4889.20), ('H11', 834.72),
    ('H12', 382.18), ('H13', 244.80), ('H14', 664.39), ('H15', 238.67),
    ('H16', 475.70), ('H17', 154.64), ('H18', 0.0), ('H19', 0.0),
    ('H20', 1013.30), ('H21', 0.0), ('H22', 0.0), ('H23', 0.0),
    ('H24', 0.0), ('H25', 94.33), ('H26', 417.24), ('H27', 0.0),
    ('H28', 306.93), ('H29', 0.0), ('H30', 0.0), ('R元', 113.00),
    ('R2', 0.0), ('R3', 0.0), ('R4', 0.0), ('R5', 121.00),
]
total = sum(v for _, v in DATA)
assert abs(total - 25270.66) < 0.01, total

W_IN, H_IN = 6.53, 4.26
fig, ax = plt.subplots(figsize=(W_IN, H_IN), dpi=DPI)
labels = [d[0] for d in DATA]
vals = [d[1] for d in DATA]
ax.bar(range(len(vals)), vals, color=BAR, width=0.72, zorder=3)
ax.set_xticks(range(len(labels)))
ax.set_xticklabels(labels, rotation=90, fontsize=6.4, color=TXT)
ax.set_ylim(0, 5500)
ax.set_yticks(range(0, 5501, 1000))
ax.set_yticklabels([f'{v:,}' for v in range(0, 5501, 1000)], fontsize=6.8, color=TXT)
ax.yaxis.grid(True, color=GRID, linewidth=0.6, zorder=0)
ax.set_axisbelow(True)
for s in ('top', 'right'):
    ax.spines[s].set_visible(False)
ax.spines['left'].set_color(LINE); ax.spines['bottom'].set_color(LINE)
ax.spines['left'].set_linewidth(0.6); ax.spines['bottom'].set_linewidth(0.6)
ax.tick_params(length=0, pad=2)
ax.text(0.995, 0.965, f'総延長　{total:,.2f} ｍ', transform=ax.transAxes,
        ha='right', va='top', fontsize=8.2, color=TXT)
ax.text(0.995, 0.905, '※　法定耐用年数（50年）に達する管渠は令和24年度以降',
        transform=ax.transAxes, ha='right', va='top', fontsize=6.8, color='#777777')
fig.subplots_adjust(left=0.085, right=0.99, top=0.97, bottom=0.11)
fig.savefig(f'{OUT}/kankyo_chart.png', dpi=DPI, facecolor='white')
plt.close(fig)
print(f'管渠グラフ: {W_IN}x{H_IN}in  合計 {total:,.2f}m')

# ── 4. 処理場の概要（表）────────────────────────────────────────────
# 出典：提供資料「施設の概要.pdf」（中札内村下水道事業の概要）
W_IN, H_IN = 6.53, 2.36
PW, PH = int(W_IN * DPI), int(H_IN * DPI)
img = Image.new('RGB', (PW, PH), 'white')
dr = ImageDraw.Draw(img)
F = lambda s: ImageFont.truetype(FONT_PATH, s)
f_head, f_cell, f_small, f_note = F(17), F(17), F(15), F(14)

HEADERS = ['施設名', '供用開始', '水処理方式', '計画処理人口', '処理能力', '整備率']
ROW = ['中札内浄化センター', '平成9年3月\n（平成8年度）', 'オキシデーション\nディッチ法',
       '2,661 人', '1,520 m³/日\n（760 m³/日×2池）', '95.7 ％']
WIDTHS = [0.215, 0.135, 0.185, 0.135, 0.205, 0.125]
x0, y0 = 6, 8
tw = PW - 2 * x0
hh, rh = 34, 60
xs = [x0]
for w_ in WIDTHS:
    xs.append(xs[-1] + int(tw * w_))
xs[-1] = x0 + tw

def center(d, box, text, font, fill):
    x1, y1, x2, y2 = box
    lines = text.split('\n')
    lh = font.size + 5
    ty = (y1 + y2) / 2 - lh * len(lines) / 2 + 1
    for ln in lines:
        w_ = d.textlength(ln, font=font)
        d.text(((x1 + x2) / 2 - w_ / 2, ty), ln, font=font, fill=fill)
        ty += lh

# ヘッダー
dr.rectangle([x0, y0, x0 + tw, y0 + hh], fill=HEADBG)
for i, hname in enumerate(HEADERS):
    center(dr, (xs[i], y0, xs[i + 1], y0 + hh), hname, f_head, 'white')
# データ行
yy = y0 + hh
dr.rectangle([x0, yy, x0 + tw, yy + rh], fill=ROWBG, outline=LINE)
for i, v in enumerate(ROW):
    center(dr, (xs[i], yy, xs[i + 1], yy + rh), v, f_cell if '\n' not in v else f_small, TXT)
for x in xs:
    dr.line([(x, y0), (x, yy + rh)], fill=LINE, width=1)
dr.line([(x0, yy + rh), (x0 + tw, yy + rh)], fill=LINE, width=1)

# 面積・区域の補足表
yy2 = yy + rh + 20
SUB = [('計画面積', '164.0 ha'), ('処理区域面積', '157.0 ha'),
       ('行政区域内人口', '3,869 人'), ('水洗化人口', '2,599 人'), ('水洗化率', '97.7 ％')]
sw = tw // len(SUB)
dr.rectangle([x0, yy2, x0 + sw * len(SUB), yy2 + 30], fill=ALTBG, outline=LINE)
dr.rectangle([x0, yy2 + 30, x0 + sw * len(SUB), yy2 + 30 + 34], fill=ROWBG, outline=LINE)
for i, (k, v) in enumerate(SUB):
    bx = x0 + sw * i
    center(dr, (bx, yy2, bx + sw, yy2 + 30), k, f_small, TXT)
    center(dr, (bx, yy2 + 30, bx + sw, yy2 + 64), v, f_cell, TXT)
    dr.line([(bx, yy2), (bx, yy2 + 64)], fill=LINE, width=1)
dr.line([(x0 + sw * len(SUB), yy2), (x0 + sw * len(SUB), yy2 + 64)], fill=LINE, width=1)

note = '※　出典：中札内村下水道事業の概要（事業計画・目標年次 令和7年度）。敷地面積・放流水質及び主要施設の一覧は資料受領後に追記。'
dr.text((x0 + 2, yy2 + 76), note, font=f_note, fill='#777777')
content_h = yy2 + 76 + 22
img = img.crop((0, 0, PW, content_h))
img.save(f'{OUT}/shorijo_table.png', dpi=(DPI, DPI))
ratio = img.height / img.width
print(f'処理場概要表: {img.width}x{img.height}px → {W_IN:.2f}x{W_IN*ratio:.3f}in (縦横比{ratio:.4f})')
print('\n生成物:', sorted(os.listdir(OUT)))
