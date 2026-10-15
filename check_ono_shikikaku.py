"""色覚への配慮の点検 ― 図の配色が、色だけに頼っていないかを確かめる。

なぜ要るか
  計画書は紙で配られ、白黒で複写されることがある。また、日本人男性の
  およそ20人に1人がＰ型・Ｄ型の色覚である。
  **色の違いだけで系列を見分けさせる図は、一定の割合の読み手に届かない。**

やること
  1. 配色の6色について、**すべての組み合わせ**の色の差を
     ①通常の色覚 ②Ｐ型（1型） ③Ｄ型（2型） ④Ｔ型（3型） ⑤白黒
     の5通りで数える。
  2. **Ｐ型・Ｄ型で近づく組み合わせ**と、**白黒にすると近づく組み合わせ**を挙げる。
     どちらも近い組み合わせは、**色だけで見分けさせてはいけない。**
  3. 実際に作った図（PNG）を読み、**使われている色の組み合わせ**を調べて、
     危ない組み合わせを現に使っている図を挙げる。

色覚の模擬は Viénot, Brettel & Mollon (1999) の線形変換による。
色の差は CIE L*a*b* の ΔE*ab（1976）で測る。
  ΔE 2.3 … 見分けの限界（JND）
  ΔE 10  … 離れていると言える目安（本書ではこれを下限とする）

出力
  画面のみ（成果品ではない）。
"""

import colorsys
import itertools
import pathlib
import sys
import warnings

warnings.filterwarnings("ignore")

ROOT = pathlib.Path(__file__).parent
sys.path.insert(0, str(ROOT))

DE_MIN = 10.0        # これを下回ったら「近い」とみなす
FIGDIR = ROOT / "output" / "ono_fig"


# ---------------------------------------------------------------- 色の変換
def hex2rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def _lin(c):
    return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4


def _unlin(c):
    c = min(max(c, 0.0), 1.0)
    return c * 12.92 if c <= 0.0031308 else 1.055 * c ** (1 / 2.4) - 0.055


def rgb2xyz(rgb):
    r, g, b = (_lin(c) for c in rgb)
    return (0.4124 * r + 0.3576 * g + 0.1805 * b,
            0.2126 * r + 0.7152 * g + 0.0722 * b,
            0.0193 * r + 0.1192 * g + 0.9505 * b)


def xyz2lab(xyz):
    wx, wy, wz = 0.95047, 1.0, 1.08883

    def f(t):
        return t ** (1 / 3) if t > 216 / 24389 else (841 / 108) * t + 4 / 29
    fx, fy, fz = f(xyz[0] / wx), f(xyz[1] / wy), f(xyz[2] / wz)
    return (116 * fy - 16, 500 * (fx - fy), 200 * (fy - fz))


def lab(h):
    return xyz2lab(rgb2xyz(hex2rgb(h)))


def de(a, b):
    la, lb = lab(a), lab(b)
    return sum((x - y) ** 2 for x, y in zip(la, lb)) ** 0.5


# ---------------------------------------------------------------- 色覚の模擬
# Viénot, Brettel & Mollon (1999) の線形RGB上の変換行列。
SIM = {
    "Ｐ型": ((0.11238, 0.88762, 0.0),
             (0.11238, 0.88762, 0.0),
             (0.00401, -0.00401, 1.0)),
    "Ｄ型": ((0.29275, 0.70725, 0.0),
             (0.29275, 0.70725, 0.0),
             (-0.02234, 0.02234, 1.0)),
    "Ｔ型": ((1.0, 0.14461, -0.14461),
             (0.0, 0.85659, 0.14341),
             (0.0, 0.85659, 0.14341)),
}


def simulate(h, kind):
    r, g, b = (_lin(c) for c in hex2rgb(h))
    m = SIM[kind]
    out = tuple(m[i][0] * r + m[i][1] * g + m[i][2] * b for i in range(3))
    return "#%02X%02X%02X" % tuple(round(_unlin(c) * 255) for c in out)


def grey(h):
    """白黒複写。CIE の明度（Y）で落とす。"""
    y = rgb2xyz(hex2rgb(h))[1]
    v = round(_unlin(y) * 255)
    return "#%02X%02X%02X" % (v, v, v)


# ---------------------------------------------------------------- 配色
def palette():
    import ono_figs as F
    return [("強調・最大", F.RED), ("主系列", F.BLUE), ("淡・参考", F.PALE),
            ("参照・基準", F.ORANGE), ("減少・対象外", F.GREY),
            ("数値ラベル", F.NAVY)]


def main():
    P = palette()
    print("■ 色覚への配慮の点検")
    print()
    print("── 1　配色（%d色）" % len(P))
    for nm, h in P:
        L = lab(h)[0]
        print("   %-12s %s  明度L*=%5.1f  Ｐ型%s  Ｄ型%s  白黒%s"
              % (nm, h, L, simulate(h, "Ｐ型"), simulate(h, "Ｄ型"), grey(h)))

    print()
    print("── 2　組み合わせごとの色の差（ΔE*ab）")
    print("   %-26s %6s %6s %6s %6s %6s" %
          ("組み合わせ", "通常", "Ｐ型", "Ｄ型", "Ｔ型", "白黒"))
    abunai = []
    for (n1, h1), (n2, h2) in itertools.combinations(P, 2):
        row = {"通常": de(h1, h2)}
        for k in SIM:
            row[k] = de(simulate(h1, k), simulate(h2, k))
        row["白黒"] = de(grey(h1), grey(h2))
        warn = [k for k, v in row.items() if v < DE_MIN]
        mark = "  **近い：%s**" % "・".join(warn) if warn else ""
        print("   %-26s %6.1f %6.1f %6.1f %6.1f %6.1f%s"
              % ("%s↔%s" % (n1, n2), row["通常"], row["Ｐ型"], row["Ｄ型"],
                 row["Ｔ型"], row["白黒"], mark))
        if warn:
            abunai.append(((n1, h1), (n2, h2), warn, row))

    print()
    print("── 3　判定")
    if not abunai:
        print("   **どの組み合わせも、Ｐ型・Ｄ型・Ｔ型・白黒のいずれでも"
              "ΔE %.0f 以上離れています。**" % DE_MIN)
        print("   色だけで見分けさせても届きます。")
    else:
        print("   **次の組み合わせは、色だけで見分けさせてはいけません。**")
        for (n1, h1), (n2, h2), warn, row in abunai:
            print("   ・%s（%s）↔ %s（%s）　%s で ΔE %.1f"
                  % (n1, h1, n2, h2, "・".join(warn),
                     min(row[k] for k in warn)))
            print("     → 凡例・数値ラベル・並び順のいずれかで区別できること")

    # ---- 4 実際の図
    print()
    print("── 4　現に作っている図")
    pngs = sorted(FIGDIR.glob("*.png")) if FIGDIR.exists() else []
    if not pngs:
        print("   図の出力先（%s）に PNG がありません。"
              % FIGDIR.relative_to(ROOT))
        print("   **素案の生成を一度通してから、もう一度この点検を行います。**")
    else:
        try:
            from PIL import Image
        except ImportError:
            print("   Pillow がないため、図の色は調べられません。")
            pngs = []
        for p in pngs:
            im = Image.open(p).convert("RGB")
            use = set()
            cnt = {}
            for c, rgb in im.getcolors(maxcolors=1 << 24) or []:
                cnt[rgb] = cnt.get(rgb, 0) + c
            for nm, h in P:
                want = tuple(round(x * 255) for x in hex2rgb(h))
                n = sum(v for rgb, v in cnt.items()
                        if max(abs(a - b) for a, b in zip(rgb, want)) <= 6)
                if n >= 200:            # 線1本ぶんより多く使われている
                    use.add(nm)
            hit = [x for x in abunai
                   if x[0][0] in use and x[1][0] in use]
            print("   %-34s 使用色 %s%s"
                  % (p.name, "・".join(sorted(use)) or "―",
                     "　**要確認：%s**"
                     % "／".join("%s↔%s" % (a[0], b[0]) for a, b, _w, _r in hit)
                     if hit else ""))

    print()
    print("── 5　この点検でできないこと")
    print("   ・**凡例や数値ラベルがあるかどうかは、図を見ないと分かりません。**")
    print("     本点検は色の差だけを数えます。")
    print("   ・印刷機の調子や紙質による見え方の違いは測れません。")
    print("   ・Ｐ型・Ｄ型の模擬は、最も強い場合（2色覚）を仮定しています。")
    print("     実際には程度の差があり、見分けられる方もいらっしゃいます。")
    return abunai


if __name__ == "__main__":
    main()
