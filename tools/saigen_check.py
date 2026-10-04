# -*- coding: utf-8 -*-
"""リポジトリの複製だけから成果品を作り直せることを確かめる.

作業領域（セッション固有の場所）は環境が作り直された時点で消える。
そこにしか本文・数値がない成果品は、同じ内容を作り直せない。

本点検は、別の場所に複製したリポジトリで処理を走らせた結果と、
収録してある成果品とを中身で比べる。

比べ方
  docx・xlsx・odt  zip の中身で比べる（作成日時の `docProps/core.xml`・
                   `app.xml` は作り直すたびに変わるため比べない）
  png              画素で比べる（作り直すと付帯情報だけが変わる）
  そのほか         バイトで比べる

使い方
  git clone -q --branch <ブランチ> --single-branch <このリポジトリ> <複製>
  cd <複製> && for f in build_*.py; do python3 "$f"; done
  python3 tools/saigen_check.py <複製>/output

  第2引数を与えると、収録してある側を差し替えられる（既定は本リポジトリ）。
"""

import hashlib
import os
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import repo_paths as RP                                       # noqa: E402

# 作成日時だけが変わるもの。中身の比較から除く
ZIP_SKIP = {"docProps/core.xml", "docProps/app.xml",
            "meta.xml", "settings.xml"}
ZIP_EXT = (".docx", ".xlsx", ".odt", ".pptx", ".ods")

# 中身が実行の時点により変わる成果品（作り直せないことを表すものではない）
JIKOKU = {
    "第10期計画_日次の状況と翌日の作業順位.xlsx":
        "基準日のコミットを履歴から読むため",
    "第10期計画_中間報告の根拠対照表.xlsx":
        "収録ファイルのハッシュ値を掲げているため",
}


def _img_bytes(b):
    """画像の中身を画素で表す。作り直すと付帯情報だけが変わるため。"""
    import io
    from PIL import Image
    with Image.open(io.BytesIO(b)) as im:
        return im.convert("RGBA").tobytes()


def _zip_digest(path):
    h = hashlib.md5()
    with zipfile.ZipFile(path) as z:
        for n in sorted(z.namelist()):
            if n in ZIP_SKIP:
                continue
            b = z.read(n)
            if n.lower().endswith((".png", ".jpg", ".jpeg")):
                try:
                    b = _img_bytes(b)
                except Exception:                        # noqa: BLE001
                    pass
            h.update(n.encode("utf-8"))
            h.update(b)
    return h.hexdigest()


def _png_digest(path):
    with open(path, "rb") as f:
        return hashlib.md5(_img_bytes(f.read())).hexdigest()


def _digest(path):
    if path.endswith(ZIP_EXT):
        return _zip_digest(path)
    if path.endswith(".png"):
        return _png_digest(path)
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 2
    new = os.path.abspath(argv[1])
    cur = os.path.abspath(argv[2]) if len(argv) > 2 else RP.OUTPUT

    def _rels(d):
        out = set()
        for root, _, fs in os.walk(d):
            if os.path.basename(root) == "zip":
                continue
            for f in fs:
                out.add(os.path.relpath(os.path.join(root, f), d))
        return out

    a, b = _rels(cur), _rels(new)
    same, diff, err, jik = 0, [], [], []
    for rel in sorted(a & b):
        try:
            if _digest(os.path.join(cur, rel)) == _digest(
                    os.path.join(new, rel)):
                same += 1
            elif os.path.basename(rel) in JIKOKU:
                jik.append(rel)
            else:
                diff.append(rel)
        except Exception as e:                           # noqa: BLE001
            err.append((rel, str(e)))

    only_cur = sorted(a - b)      # 複製では作られなかったもの
    only_new = sorted(b - a)      # 複製にだけできたもの

    print("■ リポジトリの複製からの作り直しの点検")
    print("  収録 %d件／複製 %d件" % (len(a), len(b)))
    print("  中身が同じ %d件／違う %d件／比べられない %d件"
          % (same, len(diff), len(err)))
    for rel in diff:
        print("  【違う】%s" % rel)
    for rel, e in err:
        print("  【比べられない】%s　%s" % (rel, e))
    if jik:
        print("")
        print("［実行の時点により中身が変わるもの］%d件" % len(jik))
        for rel in jik:
            print("  %s　%s" % (rel, JIKOKU[os.path.basename(rel)]))
    if only_cur:
        print("")
        print("［複製では作られなかったもの］%d件" % len(only_cur))
        for rel in only_cur:
            print("  %s" % rel)
        print("  ※ 作業領域の受領資料を読むものはここに現れる。"
              "実体が追跡下にあるため内容は失われない（tools/jizoku_check.py）。")
    if only_new:
        print("")
        print("［複製にだけできたもの］%d件" % len(only_new))
        for rel in only_new:
            print("  %s" % rel)
    return 1 if (diff or err or only_new) else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
