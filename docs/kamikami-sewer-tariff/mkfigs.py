#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""説明資料に貼る図（PNG）を生成する。HTML+SVGで描いてブラウザで撮影する。"""
import io, json, math, os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from compare import monthly, CURRENT

TEAL, RUST = '#12939E', '#CF5B1A'          # 検証済みの2色（validate_palette.js）
RAMP = {'弱': '#5FB4BD', '中': '#12939E', '強': '#0A5A66'}
GRAY = '#8A9AA0'
INK, MUTED, RULE = '#16242A', '#6B8189', '#D3DEDE'

# 税抜単価（条例別表と同じ基準）。税込は ×1.1 の1円未満切捨てで算定
CUR_NET = (1008, 37, 174, 200)
PLANS = {
    '縮小なし':  (1008, 37, 214, 220),
    '弱い縮小':  (1008, 55, 202, 220),
    '中心案':    (1008, 73, 190, 220),
    '強い縮小':  (1008, 100, 172, 220),
}

def mnet(v, base, r1, r2, r3):
    """1か月あたりの税抜使用料。"""
    n = base
    if v > 5:  n += r1 * min(v - 5, 5)
    if v > 10: n += r2 * min(v - 10, 40)
    if v > 50: n += r3 * (v - 50)
    return n

def svg_open(w, h):
    return ['<svg viewBox="0 0 %d %d" width="%d" height="%d" xmlns="http://www.w3.org/2000/svg">' % (w, h, w, h),
            '<rect width="%d" height="%d" fill="#ffffff"/>' % (w, h)]

def txt(x, y, s, size=13, fill=INK, anchor='start', weight='400'):
    return ('<text x="%.1f" y="%.1f" font-size="%s" fill="%s" text-anchor="%s" font-weight="%s" '
            'font-family="IPAPGothic, Noto Sans JP, sans-serif">%s</text>' % (x, y, size, fill, anchor, weight, s))

# ---------- 図1 現行の単価カーブ ----------
def fig1():
    W, H, L, R, T, B = 880, 380, 78, 30, 34, 56
    out = svg_open(W, H)
    X = lambda v: L + (v - 5) / 45 * (W - L - R)
    Y = lambda u: T + (230 - u) / 110 * (H - T - B)
    for u in (130, 150, 170, 190, 210, 230):
        out.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="%s"/>' % (L, Y(u), W - R, Y(u), RULE))
        out.append(txt(L - 10, Y(u) + 5, str(u), 12, MUTED, 'end'))
    for v in range(5, 51, 5):
        out.append(txt(X(v), H - B + 22, str(v), 12, MUTED, 'middle'))
    out.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="%s" stroke-width="1.5" stroke-dasharray="6 5"/>'
               % (L, Y(220), W - R, Y(220), RUST))
    out.append(txt(W - R, Y(220) - 9, 'どれだけ使っても 220円/㎥ が上限', 12.5, RUST, 'end', '700'))
    pts = [(v, monthly(v, **CURRENT) / v) for v in range(5, 51)]
    out.append('<path d="M%s" fill="none" stroke="%s" stroke-width="2.5" stroke-linejoin="round"/>'
               % (' L'.join('%.1f %.1f' % (X(v), Y(u)) for v, u in pts), TEAL))
    for v, lab, dx, dy, anc in ((5, '月5㎥ 221.8円/㎥', 10, 5, 'start'),
                                (10, '月10㎥ 131.2円/㎥', 14, 4, 'start'),
                                (50, '月50㎥ 179.4円/㎥', -10, 5, 'end')):
        u = monthly(v, **CURRENT) / v
        out.append('<circle cx="%.1f" cy="%.1f" r="6" fill="%s" stroke="#fff" stroke-width="2.5"/>' % (X(v), Y(u), TEAL))
        out.append(txt(X(v) + dx, Y(u) + dy, lab, 13, INK, anc, '700'))
    out.append(txt(X(10) + 14, Y(131.23) + 20, '全区分で最も安い（中央値の世帯）', 12, MUTED))
    out.append(txt(L - 66, T - 12, '円/㎥', 12, MUTED))
    out.append(txt(W - R, H - 14, '月使用量（㎥）', 12, MUTED, 'end'))
    out.append('</svg>')
    return '\n'.join(out)

# ---------- 図2 収入シェアと水量シェアの乖離 ----------
def fig2():
    """区分別の従量収入シェアと従量課金対象水量シェア（同一分母で対比）"""
    rows = [('6〜10㎥', 10.0, 34.9), ('11〜50㎥', 70.5, 52.5), ('51㎥〜（大口）', 19.5, 12.7)]
    W, H, L, R, T = 880, 300, 190, 130, 78
    rowh, barh = 60, 19
    out = svg_open(W, H)
    out.append(txt(L - 22, 24, '2事業合計・令和7年度（12か月・全件）', 12.5, MUTED, 'start'))
    out.append(txt(L, 50, '従量収入シェア', 13, TEAL, 'start', '700'))
    out.append('<rect x="%d" y="40" width="14" height="14" fill="%s"/>' % (L - 22, TEAL))
    out.append(txt(L + 132, 50, '従量課金対象水量シェア', 13, RUST, 'start', '700'))
    out.append('<rect x="%d" y="40" width="14" height="14" fill="%s"/>' % (L + 110, RUST))

    scale = (W - L - R) / 75.0
    for i, (lab, rev, vol) in enumerate(rows):
        y = T + i * rowh
        out.append(txt(L - 14, y + 15, lab, 13, INK, 'end'))
        out.append('<rect x="%d" y="%.1f" width="%.1f" height="%d" fill="%s" rx="3"/>' % (L, y, rev * scale, barh, TEAL))
        out.append(txt(L + rev * scale + 8, y + 15, '%.1f%%' % rev, 12.5, INK, 'start', '700'))
        out.append('<rect x="%d" y="%.1f" width="%.1f" height="%d" fill="%s" rx="3"/>'
                   % (L, y + barh + 3, vol * scale, barh, RUST))
        out.append(txt(L + vol * scale + 8, y + barh + 18, '%.1f%%' % vol, 12.5, INK, 'start', '700'))
    out.append(txt(24, H - 14, '※いずれも従量部分のみを分母とする。従量課金対象水量＝基本水量（1か月5㎥）を超える水量',
                   11.5, MUTED))
    y = T
    out.append('<rect x="%d" y="%.1f" width="%d" height="%d" fill="none" stroke="%s" stroke-width="2" rx="4"/>'
               % (L - 176, y - 8, W - R - L + 250, barh * 2 + 20, RUST))
    out.append(txt(W - R + 66, y + 22, '対象水量の34.9%に対し', 12.5, RUST, 'end', '700'))
    out.append(txt(W - R + 66, y + 39, '従量収入は10.0%', 12.5, RUST, 'end', '700'))
    out.append('</svg>')
    return '\n'.join(out)

# ---------- 図3 単価差縮小の水準別の世帯影響 ----------
def fig3():
    W, H, L, R, T, B = 880, 380, 70, 150, 34, 56
    vs = list(range(5, 51))
    X = lambda v: L + (v - 5) / 45 * (W - L - R)
    Y = lambda d: T + (30 - d) / 32 * (H - T - B)
    out = svg_open(W, H)
    for d in (0, 6, 12, 18, 24, 30):
        out.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="%s"/>' % (L, Y(d), W - R, Y(d), RULE))
        out.append(txt(L - 10, Y(d) + 5, '%+d%%' % d, 12, MUTED, 'end'))
    for v in range(5, 51, 5):
        out.append(txt(X(v), H - B + 22, str(v), 12, MUTED, 'middle'))
    styles = [('縮小なし', GRAY, '7 5', 2), ('弱い縮小', RAMP['弱'], '2 4', 2.5),
              ('中心案', RAMP['中'], None, 3.5), ('強い縮小', RAMP['強'], None, 2.5)]
    for lab, col, dash, wdt in styles:
        p = PLANS[lab]
        pts = [(v, mnet(v, *p) / mnet(v, *CUR_NET) - 1) for v in vs]
        d = ' L'.join('%.1f %.1f' % (X(v), Y(x * 100)) for v, x in pts)
        out.append('<path d="M%s" fill="none" stroke="%s" stroke-width="%s" stroke-linejoin="round"%s/>'
                   % (d, col, wdt, ' stroke-dasharray="%s"' % dash if dash else ''))
        ly = Y(pts[-1][1] * 100)
        out.append('<circle cx="%.1f" cy="%.1f" r="5" fill="%s" stroke="#fff" stroke-width="2"/>' % (X(50), ly, col))
        out.append(txt(X(50) + 10, ly + 5, lab, 12.5, col, 'start', '700'))
    out.append(txt(L - 58, T - 12, '増減率', 12, MUTED))
    out.append(txt(W - R, H - 14, '月使用量（㎥）', 12, MUTED, 'end'))
    out.append(txt(L + 6, T + 14, 'いずれも平均増収率 約10%・基本使用料は据置', 12, MUTED))
    out.append('</svg>')
    return '\n'.join(out)

# ---------- 図4 段階的な縮小で単価の開きが縮まる ----------
def fig4():
    groups = [('現行', 37, 174, '4.70倍'), ('第1段階（令和8年度）', 73, 190, '2.60倍'),
              ('第2段階（令和13年度）', 109, 207, '1.90倍')]
    W, H, L, R, T, B = 880, 360, 70, 40, 62, 74
    gw = (W - L - R) / 3
    Y = lambda u: T + (220 - u) / 220 * (H - T - B)
    out = svg_open(W, H)
    out.append('<rect x="%d" y="18" width="14" height="14" fill="%s"/>' % (L, TEAL))
    out.append(txt(L + 22, 30, '6〜10㎥の単価（税抜）', 13, TEAL, 'start', '700'))
    out.append('<rect x="%d" y="18" width="14" height="14" fill="%s"/>' % (L + 220, RUST))
    out.append(txt(L + 242, 30, '11〜50㎥の単価（税抜）', 13, RUST, 'start', '700'))
    for u in (0, 55, 110, 165, 220):
        out.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="%s"/>' % (L, Y(u), W - R, Y(u), RULE))
        out.append(txt(L - 10, Y(u) + 5, str(u), 12, MUTED, 'end'))
    for i, (lab, a_, b_, jump) in enumerate(groups):
        cx = L + gw * i + gw / 2
        bw = 62
        out.append('<rect x="%.1f" y="%.1f" width="%d" height="%.1f" fill="%s" rx="3"/>'
                   % (cx - bw - 3, Y(a_), bw, Y(0) - Y(a_), TEAL))
        out.append(txt(cx - bw / 2 - 3, Y(a_) - 9, '%g円' % a_, 12.5, INK, 'middle', '700'))
        out.append('<rect x="%.1f" y="%.1f" width="%d" height="%.1f" fill="%s" rx="3"/>'
                   % (cx + 3, Y(b_), bw, Y(0) - Y(b_), RUST))
        out.append(txt(cx + bw / 2 + 3, Y(b_) - 9, '%g円' % b_, 12.5, INK, 'middle', '700'))
        out.append(txt(cx, H - B + 24, lab, 12.5, INK, 'middle', '700'))
        out.append(txt(cx, H - B + 43, '単価の開き %s' % jump, 12.5, RUST if i == 0 else MUTED,
                       'middle', '700' if i == 0 else '400'))
    out.append(txt(L - 58, T - 14, '円/㎥', 12, MUTED))
    out.append('</svg>')
    return '\n'.join(out)

FIGS = {'fig1': fig1(), 'fig2': fig2(), 'fig3': fig3(), 'fig4': fig4()}
html = ['<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;700&display=swap">',
        '<style>body{margin:0;background:#fff;font-family:"IPAPGothic","Noto Sans JP",sans-serif}'
        '.f{display:inline-block;background:#fff}</style>']
for k, v in FIGS.items():
    html.append('<div class="f" id="%s">%s</div>' % (k, v))
io.open(os.environ.get('KAMIKAMI_FIGS') or os.path.join(HERE, 'figs.html'), 'w', encoding='utf-8').write('\n'.join(html))
print('figs.html written')
