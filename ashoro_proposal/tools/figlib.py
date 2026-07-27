# -*- coding: utf-8 -*-
"""提案書に差し込む図を描くための最小限の作図ライブラリ（Pillow）。

・キャンバス幅 1800px＝本文幅15.5cm（約295dpi）を前提とする。
・色は提案書の配色（濃紺1F4E78／スレート44546A）に合わせ、
  データを表す印だけは検証済みのカテゴリ配色（青2a78d6／橙eb6834）を使う。
"""
from PIL import Image, ImageDraw, ImageFont

W = 1800
FONT_PATH = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"

NAVY = (31, 78, 120)
SLATE = (68, 84, 106)
LIGHT = (232, 238, 245)
PALE = (244, 247, 250)
INK = (38, 38, 38)
MUTED = (89, 89, 89)
BORDER = (191, 191, 191)
WHITE = (255, 255, 255)
# 検証済みカテゴリ配色（validate_palette.js で全項目PASS）
S_BLUE = (42, 120, 214)     # 70歳未満
S_ORANGE = (235, 104, 52)   # 70歳以上

# 本文幅15.5cmに1800pxを割り当てるため、1px≒0.086mm。
# 図中の文字を印刷時8pt以上にするには em で32px以上が必要なので、
# 指定サイズを一律この係数で拡大する。
SCALE = 1.36

_cache = {}


def font(size):
    px = max(8, int(round(size * SCALE)))
    if px not in _cache:
        _cache[px] = ImageFont.truetype(FONT_PATH, px)
    return _cache[px]


def canvas(height, bg=WHITE):
    im = Image.new("RGB", (W, height), bg)
    return im, ImageDraw.Draw(im)


# ---------------------------------------------------------------- テキスト

_NO_HEAD = "、。）」』】〕・％%）,.…ー"
_NO_TAIL = "（「『【〔（"


def wrap(text, f, width):
    """日本語向けの折り返し（簡易な禁則処理つき）。改行は \n で明示できる。"""
    lines = []
    for para in text.split("\n"):
        cur = ""
        for ch in para:
            if f.getlength(cur + ch) <= width or not cur:
                cur += ch
            else:
                if ch in _NO_HEAD:          # 行頭にできない文字は前行に送る
                    cur += ch
                    lines.append(cur)
                    cur = ""
                elif cur[-1] in _NO_TAIL:   # 行末にできない文字は次行へ送る
                    lines.append(cur[:-1])
                    cur = cur[-1] + ch
                else:
                    lines.append(cur)
                    cur = ch
        lines.append(cur)
    return [l for l in lines]


def text(d, xy, s, size=30, fill=INK, anchor="la", bold=False):
    d.text(xy, s, font=font(size), fill=fill, anchor=anchor,
           stroke_width=1 if bold else 0, stroke_fill=fill)


def block(d, x, y, w, s, size=26, fill=INK, lh=1.45, align="left", bold=False):
    """折り返しつきのテキスト。描画後の下端 y を返す。"""
    f = font(size)
    lines = wrap(s, f, w)
    step = size * SCALE * lh
    for i, line in enumerate(lines):
        yy = y + i * step
        if align == "center":
            d.text((x + w / 2, yy), line, font=f, fill=fill, anchor="ma",
                   stroke_width=1 if bold else 0, stroke_fill=fill)
        else:
            d.text((x, yy), line, font=f, fill=fill, anchor="la",
                   stroke_width=1 if bold else 0, stroke_fill=fill)
    return y + len(lines) * step


def box(d, x, y, w, h, fill=WHITE, outline=BORDER, width=3, radius=12, dash=False):
    if dash:
        d.rounded_rectangle([x, y, x + w, y + h], radius=radius, fill=fill)
        _dashed_round(d, x, y, w, h, radius, outline, width)
    else:
        d.rounded_rectangle([x, y, x + w, y + h], radius=radius,
                            fill=fill, outline=outline, width=width)


def _dashed_round(d, x, y, w, h, r, color, width):
    seg, gap = 16, 12
    for x0 in range(int(x + r), int(x + w - r), seg + gap):
        d.line([x0, y, min(x0 + seg, x + w - r), y], fill=color, width=width)
        d.line([x0, y + h, min(x0 + seg, x + w - r), y + h], fill=color, width=width)
    for y0 in range(int(y + r), int(y + h - r), seg + gap):
        d.line([x, y0, x, min(y0 + seg, y + h - r)], fill=color, width=width)
        d.line([x + w, y0, x + w, min(y0 + seg, y + h - r)], fill=color, width=width)


def labelled_box(d, x, y, w, h, title, body, fill=WHITE, outline=NAVY,
                 title_color=None, title_size=28, body_size=24, radius=12,
                 title_bg=None, dash=False):
    """見出し付きのボックス。"""
    box(d, x, y, w, h, fill=fill, outline=outline, radius=radius, dash=dash)
    pad = 18
    ty = y + pad
    if title:
        if title_bg:
            nlines = len(wrap(title, font(title_size), w - pad * 2))
            barh = title_size * SCALE * 1.45 * nlines + pad * 1.1
            d.rounded_rectangle([x, y, x + w, y + barh], radius=radius, fill=title_bg)
            d.rectangle([x, y + radius, x + w, y + barh], fill=title_bg)
            block(d, x + pad, y + pad * 0.55, w - pad * 2, title,
                  size=title_size, fill=WHITE, align="center", bold=True)
            ty = y + barh + pad * 0.7
        else:
            ty = block(d, x + pad, ty, w - pad * 2, title, size=title_size,
                       fill=title_color or outline, bold=True) + 8
    if body:
        block(d, x + pad, ty, w - pad * 2, body, size=body_size, fill=INK)


def arrow(d, x1, y1, x2, y2, color=SLATE, width=7, head=20):
    d.line([x1, y1, x2, y2], fill=color, width=width)
    if x2 == x1:  # 縦
        s = 1 if y2 > y1 else -1
        d.polygon([(x2, y2 + s * head), (x2 - head * 0.7, y2 - s * head * 0.2),
                   (x2 + head * 0.7, y2 - s * head * 0.2)], fill=color)
    else:         # 横
        s = 1 if x2 > x1 else -1
        d.polygon([(x2 + s * head, y2), (x2 - s * head * 0.2, y2 - head * 0.7),
                   (x2 - s * head * 0.2, y2 + head * 0.7)], fill=color)


def band(d, x, y, w, h, s, fill=LIGHT, color=NAVY, size=24, radius=10):
    d.rounded_rectangle([x, y, x + w, y + h], radius=radius, fill=fill)
    f = font(size)
    lines = wrap(s, f, w - 40)
    step = size * SCALE * 1.4
    y0 = y + (h - len(lines) * step) / 2
    for i, line in enumerate(lines):
        d.text((x + w / 2, y0 + i * step), line, font=f, fill=color, anchor="ma")


def trim_bottom(im, margin=18):
    """下端の余白を切り詰める（図ごとの高さ指定を厳密にしなくてよくする）。"""
    px = im.load()
    w, h = im.size
    last = 0
    for y in range(h - 1, -1, -1):
        row_has_ink = False
        for x in range(0, w, 3):
            if px[x, y] != WHITE:
                row_has_ink = True
                break
        if row_has_ink:
            last = y
            break
    return im.crop((0, 0, w, min(h, last + 1 + margin)))


def save(im, path):
    im = trim_bottom(im)
    im.save(path, "PNG", optimize=True)
    return path
