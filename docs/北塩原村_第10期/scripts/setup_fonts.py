# -*- coding: utf-8 -*-
"""紙面の確認に使う BIZ UDPゴシック をこの環境に入れる。

【なぜ要るか】
  計画書の本文の書体は BIZ UDPゴシック である（build_soan_docx.js の注記）。
  村の端末（Windows 10 以降）には標準で入っているが、この作業環境には無い。
  無いまま docx を PDF に変換すると別の書体に置き換わり、字幅も行の高さも
  変わるため、頁数の実測（jissoku_*）も見切れの確認も当てにならない。
  図（matplotlib）も同じ書体で描くために要る。

【どこから取るか】
  BIZ UDPゴシック は SIL Open Font License の書体である。
  この環境から github.com の配布物には届かないため、npm の
  @fontsource/biz-udpgothic（版を固定）から取る。fontsource は書体を
  字の範囲ごとに分けて配っているので、和文の本体（japanese）に、
  記号（※・①〜⑦・▲・→ など）を持つ分割片を足して1つにまとめる。
  ローマ数字（Ⅰ〜Ⅳ）・α などは fontsource の配布に無く、置き換わる。
  数文字であり頁数には効かない。

  docx が書体を「BIZ UDPゴシック」（和名）で指定するため、
  名前の表に和名を足してから ~/.fonts に置く。

  python3 scripts/setup_fonts.py
"""
import glob
import os
import subprocess
import sys
import tarfile
import tempfile

sys.dont_write_bytecode = True
PKG = "@fontsource/biz-udpgothic@5.3.0"
DEST = os.path.expanduser("~/.fonts")
KIGO = set("α―※ⅠⅡⅢⅣⅰ→①②③④⑤⑥⑦⑧⑨⑩▲△▽○●■□◆◇★☆〇～・")


def main():
    from fontTools.ttLib import TTFont
    from fontTools.merge import Merger
    os.makedirs(DEST, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="fnt_") as d:
        subprocess.run(["npm", "pack", PKG, "--silent"], cwd=d, check=True,
                       capture_output=True)
        tgz = glob.glob(os.path.join(d, "*.tgz"))[0]
        with tarfile.open(tgz) as t:
            t.extractall(d, filter="data")
        files = os.path.join(d, "package", "files")
        for w, sub in (("400", "Regular"), ("700", "Bold")):
            base = os.path.join(files, "biz-udpgothic-japanese-%s-normal.woff" % w)
            want = KIGO - {chr(c) for c in TTFont(base).getBestCmap()}
            picks = []
            for fn in sorted(glob.glob(os.path.join(files, "biz-udpgothic-*-%s-normal.woff" % w))):
                if fn == base:
                    continue
                got = {chr(c) for c in TTFont(fn).getBestCmap()} & want
                if got:
                    picks.append(fn)
                    want -= got
            tmp = []
            for i, fn in enumerate([base] + picks):
                f = TTFont(fn)
                f.flavor = None
                pth = os.path.join(d, "m%d.ttf" % i)
                f.save(pth)
                tmp.append(pth)
            m = Merger().merge(tmp)
            for nid, v in ((1, "BIZ UDPゴシック"), (2, sub), (16, "BIZ UDPゴシック")):
                m["name"].setName(v, nid, 3, 1, 0x411)
            out = os.path.join(DEST, "BIZUDPGothic-%s.ttf" % sub)
            m.save(out)
            print("  %s（%d字。fontsource に無い字：%s）" % (out, len(m.getBestCmap()),
                                                           "".join(sorted(want)) or "なし"))
    subprocess.run(["fc-cache", "-f"], check=False)
    # matplotlib は書体の一覧を控えに持つため、消して作り直させる
    import shutil
    shutil.rmtree(os.path.expanduser("~/.cache/matplotlib"), ignore_errors=True)
    print("  fc-match：", subprocess.run(["fc-match", "BIZ UDPゴシック"], capture_output=True,
                                       text=True).stdout.strip())


if __name__ == "__main__":
    main()
