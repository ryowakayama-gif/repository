# -*- coding: utf-8 -*-
"""docx を PDF に変換し、紙面の実測値をレイアウトの前提と突き合わせる。

【なぜ必要か】
  estimate_pages.py／estimate_layout.py は build_*_docx.js が指定した
  レイアウト値から頁数を積み上げる。積み上げの前提（用紙・余白・行送り・
  表の幅・行の下駄）が実際の組版と合っていなければ、推定した頁数も、
  「削減案で110頁に収まるか」という判断も当てにならない。
  この道具は PDF から実測値を読み、前提と食い違う点を挙げる。

【環境】
  令和8年10月10日まで、この環境には libreoffice-writer と poppler-utils が
  入っておらず変換ができなかった。--setup で入れ方を表示する。

  **フォントの置換に注意。** 游ゴシックはこの環境に無い。既定の置換では
  和文が WenQuanYi Zen Hei、欧文が DejaVu Sans に分かれ、DejaVu Sans の
  数字は全角比0.64と広いため紙面が実際より約3頁ぶん膨らむ。--setup が書く
  fontconfig の別名で IPAゴシック（欧文も半角固定・全角比0.5）に寄せる。
  それでも Word＋游ゴシックそのものではないため、**この変換は Word の紙面の
  代わりにはならない。** 分かるのは、積み上げの前提が組版と合っているか、
  表が頁をまたぐ箇所がどこか、といったことである。

【使い方】
  python3 scripts/render_check.py --setup              環境の確認と整え方
  python3 scripts/render_check.py output/06_….docx     変換して実測値を突き合わせる
  python3 scripts/render_check.py output/06_….docx --images 1-8
                                                       頁を画像にして目視する
"""
import sys
sys.dont_write_bytecode = True
import argparse, glob, os, re, shutil, subprocess, tempfile, collections, statistics

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

FC = os.path.expanduser("~/.config/fontconfig/fonts.conf")
FC_BODY = """<?xml version="1.0"?>
<!DOCTYPE fontconfig SYSTEM "fonts.dtd">
<!-- 游ゴシックが無い環境で、和文も欧文も IPAゴシックに寄せる。
     既定の置換では欧文が DejaVu Sans になり、数字が全角比0.64と広いため
     紙面が実際より膨らむ。IPAゴシックは欧文が半角固定（0.5）である。 -->
<fontconfig>
  <match target="pattern">
    <test name="family"><string>游ゴシック</string></test>
    <edit name="family" mode="assign" binding="strong"><string>IPAGothic</string></edit>
  </match>
  <match target="pattern">
    <test name="family"><string>Yu Gothic</string></test>
    <edit name="family" mode="assign" binding="strong"><string>IPAGothic</string></edit>
  </match>
</fontconfig>
"""


def setup(apply_fc=True):
    """何が入っていて何が足りないかを示す。フォントの別名は書く。"""
    ok = True
    for cmd, pkg in (("soffice", "libreoffice-writer"),
                     ("pdftoppm", "poppler-utils"),
                     ("pdfinfo", "poppler-utils")):
        w = shutil.which(cmd)
        print(f"  {cmd:<10}{'あり ' + w if w else 'なし → apt-get install -y ' + pkg}")
        ok = ok and bool(w)
    if not ok:
        print("\n  足りないものは次で入る（apt-get update を先に走らせる）")
        print("    apt-get update && apt-get install -y libreoffice-writer poppler-utils")
    # 日本語フォント
    try:
        fl = subprocess.run(["fc-list", ":lang=ja", "family"], capture_output=True,
                            text=True, timeout=60).stdout
        fam = sorted({l.split(",")[0].strip() for l in fl.splitlines() if l.strip()})
        print(f"  和文フォント  {'／'.join(fam) or 'なし'}")
    except Exception as e:
        print(f"  和文フォント  調べられなかった（{e}）")
    if apply_fc:
        os.makedirs(os.path.dirname(FC), exist_ok=True)
        if not os.path.exists(FC) or "IPAGothic" not in open(FC, encoding="utf-8").read():
            with open(FC, "w", encoding="utf-8") as f:
                f.write(FC_BODY)
            print(f"  フォントの別名  {FC} に書いた（游ゴシック → IPAゴシック）")
        else:
            print(f"  フォントの別名  {FC} にすでにある")
    try:
        m = subprocess.run(["fc-match", "游ゴシック"], capture_output=True,
                           text=True, timeout=60).stdout.strip()
        print(f"  游ゴシックの行き先  {m}")
    except Exception:
        pass
    return ok


def _soffice(args, outdir):
    """AF_UNIX が塞がれた環境でも通るように、docx スキルの補助を使う。

    スキルが無い環境でも動くよう、見つからなければ素の soffice を叩く。
    どちらもユーザープロファイルを渡さないと「User installation could not be
    completed」で何も変換せずに終わるため、必ず渡す。
    """
    cand = glob.glob(os.path.expanduser(
        "~/.claude/skills/**/docx/scripts/office/soffice.py"), recursive=True)
    env = dict(os.environ, SAL_USE_VCLPLUGIN="svp")
    with tempfile.TemporaryDirectory(prefix="lo_profile_") as prof:
        if cand:
            cmd = [sys.executable, cand[0]] + args + ["--outdir", outdir]
        else:
            cmd = ["soffice", f"-env:UserInstallation=file://{prof}"] + args + \
                  ["--outdir", outdir]
        return subprocess.run(cmd, capture_output=True, text=True, timeout=1800, env=env)


def convert(docx, outdir):
    os.makedirs(outdir, exist_ok=True)
    src = os.path.join(outdir, os.path.basename(docx))
    if os.path.abspath(src) != os.path.abspath(docx):
        shutil.copy2(docx, src)
    pdf = os.path.splitext(src)[0] + ".pdf"
    if os.path.exists(pdf):
        os.remove(pdf)
    # soffice は cwd 基準で開くため、絶対パスで渡す
    r = _soffice(["--headless", "--convert-to", "pdf", os.path.abspath(src)], outdir)
    if not os.path.exists(pdf):
        print(r.stdout[-2000:]); print(r.stderr[-2000:])
        raise SystemExit("変換できなかった")
    return pdf


# ══════ 実測 ══════
def _hex(c):
    return None if c is None else "%02X%02X%02X" % tuple(int(round(x * 255)) for x in c)


def measure(pdf, mae=3):
    """PDF から実測値を拾う。mae は前付の頁数（本文の頁数を出すために除く）。"""
    import pymupdf
    d = pymupdf.open(pdf)
    TH = {"2E75B6", "FFFFFF", "DDEBF7"}
    m = {"頁数": d.page_count, "本文の頁数": d.page_count - mae,
         "紙の幅": d[0].rect.width, "紙の高さ": d[0].rect.height}
    pitch, rowh, tblw, fonts = [], [], collections.Counter(), collections.Counter()
    wid = collections.defaultdict(list)
    for pg in d:
        if pg.number < mae:
            continue
        # 表の行高と幅
        by = collections.defaultdict(list)
        for dr in pg.get_drawings():
            if _hex(dr.get("fill")) not in TH:
                continue
            r = dr["rect"]
            if r.width < 15 or r.height < 4 or r.height > 500:
                continue
            by[round(r.y0 * 2) / 2].append(r)
        for y0, g in by.items():
            rowh.append(max(r.y1 for r in g) - y0)
            tblw[round(max(r.x1 for r in g) - min(r.x0 for r in g))] += 1
        # 本文の行送りと字幅
        prev = None
        for b in pg.get_text("rawdict")["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                for s in l["spans"]:
                    fonts[s["font"]] += 1
                    if round(s["size"], 1) != 10.5:
                        continue
                    for c in s["chars"]:
                        ch, adv = c["c"], c["bbox"][2] - c["bbox"][0]
                        if ch.isdigit():
                            wid["数字"].append(adv / 10.5)
                        elif ord(ch) >= 0x4E00:
                            wid["漢字"].append(adv / 10.5)
                if round(max(x["size"] for x in l["spans"]), 1) == 10.5:
                    if prev is not None and 8 < l["bbox"][1] - prev < 30:
                        pitch.append(l["bbox"][1] - prev)
                    prev = l["bbox"][1]
                else:
                    prev = None
    m["本文の行送り"] = statistics.median(pitch) if pitch else None
    m["表の1行の行高"] = statistics.median([h for h in rowh if h < 24]) if rowh else None
    m["表の幅"] = tblw.most_common(1)[0][0] if tblw else None
    m["表の行の数"] = len(rowh)
    m["字幅"] = {k: statistics.median(v) for k, v in wid.items() if v}
    m["フォント"] = fonts.most_common(6)
    return m


def report(pdf, mae=3, shiryo=False):
    """shiryo=True のときは委員会資料のレイアウト値（余白1134・本文14570）で見る。"""
    import estimate_pages as EP
    yohaku = 1134 if shiryo else 1418
    body_h = EP.SH_H if shiryo else EP.BODY_H
    m = measure(pdf, mae)
    print(f"■ {os.path.basename(pdf)}　{m['頁数']}頁"
          f"（前付{mae}頁を除く本文 {m['本文の頁数']}頁）")
    print(f"  紙 {m['紙の幅']:.1f}×{m['紙の高さ']:.1f}pt"
          f"（A4縦なら 595.3×841.9）")
    print(f"  フォント {m['フォント']}")
    sub = [f for f, _ in m["フォント"] if "DejaVu" in f or "WenQuanYi" in f]
    if sub:
        print(f"  ※ 游ゴシックが置換されている（{'／'.join(sorted(set(sub)))}）。"
              "--setup で IPAゴシックに寄せると紙面が近づく")
    print()
    print("  積み上げの前提と実測の対比")
    hits = []
    def cmp(nm, jitsu, mae_, tani, yurusi):
        if jitsu is None:
            print(f"    {nm:<16}{'測れなかった':>12}"); return
        au = abs(jitsu - mae_) <= yurusi
        print(f"    {nm:<16}実測 {jitsu:8.2f}{tani}   前提 {mae_:8.2f}{tani}"
              f"   {'合う' if au else '★食い違う'}")
        if not au:
            hits.append(nm)
    cmp("本文の行送り", m["本文の行送り"], 300 / 20, "pt", 0.3)
    # 許容を0.4ptとする。実測は計画素案18.35pt・委員会資料18.65ptで、同じ
    # 表の指定（セル余白60/60・行送り240 exact・罫線sz4）でも組版の丸めで
    # 0.3pt 動く。ROW_GETA は行数の多い計画素案に合わせてある。余白の取り違えの
    # ような本当の食い違い（1pt以上）はこの許容でも挙がる。
    cmp("表の1行の行高", m["表の1行の行高"], (240 + EP.ROW_GETA) / 20, "pt", 0.4)
    cmp("表の幅", m["表の幅"], EP.TBLW / 20, "pt", 1.0)
    cmp("本文の高さ", m["紙の高さ"] - yohaku / 20 * 2, body_h / 20, "pt", 1.0)
    for k, v in sorted(m["字幅"].items()):
        cmp(f"字幅（{k}）", v, 1.0 if k == "漢字" else 0.5, "em", 0.06)
    print()
    if hits:
        print(f"  ★ {len(hits)}件が食い違っている：{'／'.join(hits)}")
        print("     字幅だけが食い違うのはフォントの置換による。それ以外は"
              "estimate_pages.py の前提を直す必要がある。")
    else:
        print("  前提はすべて実測と合っている")
    return m


def aki(pdf, mae=3, yohaku=1418, n=15):
    """頁の下に残った空白の大きい順に挙げる。

    推定（estimate_layout.py）の「送りによる空白」は keepNext の効き方を
    強く見ており、実紙面では空白が出ないことが多い。分量を詰めるときは、
    推定ではなく実紙面のこの一覧を見る。
    """
    import pymupdf
    d = pymupdf.open(pdf)
    top = yohaku / 20
    bot = d[0].rect.height - yohaku / 20
    out = []
    for pg in d:
        if pg.number < mae:
            continue
        low = top
        for b in pg.get_text("dict")["blocks"]:
            if b["bbox"][3] <= bot + 1:
                low = max(low, b["bbox"][3])
        for dr in pg.get_drawings():
            r = dr["rect"]
            if r.y1 <= bot + 1 and r.height > 1:
                low = max(low, r.y1)
        out.append((bot - low, pg.number + 1))
    out.sort(reverse=True)
    h = bot - top
    print(f"\n  頁の下に残った空白（本文の高さ {h:.0f}pt に対する割合）")
    print(f"    {'頁':>5}{'空白':>9}{'割合':>7}   先頭の行")
    for v, pno in out[:n]:
        t = [l.strip() for l in d[pno - 1].get_text("text").split("\n") if l.strip()]
        atama = (t[1] if len(t) > 1 else "")[:40]
        print(f"    {pno:5d}{v:8.0f}pt{v / h * 100:6.0f}%   {atama}")
    print(f"    空白の合計 {sum(v for v, _ in out):.0f}pt ＝ {sum(v for v, _ in out) / h:.1f}頁ぶん")
    return out


def images(pdf, rng, outdir):
    a, _, b = rng.partition("-")
    b = b or a
    subprocess.run(["pdftoppm", "-jpeg", "-r", "100", "-f", a, "-l", b, pdf,
                    os.path.join(outdir, "p")], check=True, timeout=900)
    got = sorted(glob.glob(os.path.join(outdir, "p-*.jpg")))
    print(f"  画像 {len(got)}枚 → {outdir}")
    for g in got:
        print(f"    {g}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx", nargs="?")
    ap.add_argument("--setup", action="store_true", help="環境の確認とフォントの別名")
    ap.add_argument("--out", default=None, help="置き場所（既定は一時の場所）")
    ap.add_argument("--images", default=None, help="画像にする頁（例 1-8）")
    ap.add_argument("--mae", type=int, default=3, help="前付の頁数")
    ap.add_argument("--shiryo", action="store_true",
                    help="委員会資料として見る（余白1134・本文14570 twip）")
    ap.add_argument("--aki", action="store_true",
                    help="頁の下に残った空白の大きい順に挙げる（分量を詰めるとき）")
    a = ap.parse_args()
    if a.setup or not a.docx:
        print("■ 環境")
        setup()
        if not a.docx:
            return
        print()
    outdir = a.out or tempfile.mkdtemp(prefix="render_check_")
    pdf = convert(a.docx, outdir)
    report(pdf, a.mae, a.shiryo)
    if a.aki:
        aki(pdf, a.mae, 1134 if a.shiryo else 1418)
    if a.images:
        images(pdf, a.images, outdir)
    print(f"\n  PDF {pdf}")


if __name__ == "__main__":
    main()
