# -*- coding: utf-8 -*-
"""再出力しただけで中身の変わっていない成果品を元に戻す。

docx・xlsx は zip であり、再出力するたびに作成日時などが変わるため、
中身が同じでも git の差分に出る。本スクリプトは zip の中身を比べ、
**書式まで含めて**同じものだけを `git checkout --` で戻す。

  python3 tools/restore_unchanged.py          # 戻す
  python3 tools/restore_unchanged.py --dry    # 戻さずに一覧だけ出す

比べないもの（再出力のたびに必ず変わる）
  docProps/core.xml（作成日時・更新日時）
  docProps/app.xml（作成アプリケーション）

**本文の文字だけを比べると、太字・改ページ・行の分割などの体裁の変更が
「変わっていない」と判定されて消える。** zip の中身で比べること。
"""

import io
import os
import subprocess
import sys
import zipfile

SKIP = ("docProps/core.xml", "docProps/app.xml")
IMG = (".png", ".jpg", ".jpeg")


def img_bytes(b):
    """画像の中身を画素で表す。

    **PNG は作り直すと付帯情報だけが変わるため、バイトで比べると
    全件が「変わった」と出て本当の違いが埋もれる**（CLAUDE.md §4）。
    """
    from PIL import Image
    with Image.open(io.BytesIO(b)) as im:
        return im.convert("RGBA").tobytes()


def _body(n, b):
    if n.lower().endswith(IMG):
        try:
            return img_bytes(b)
        except Exception:                                # noqa: BLE001
            return b
    return b


def members(b):
    z = zipfile.ZipFile(io.BytesIO(b))
    return {n: _body(n, z.read(n)) for n in z.namelist() if n not in SKIP}


def main(dry=False):
    out = subprocess.run(["git", "diff", "--name-only", "-z", "--", "output"],
                         capture_output=True).stdout.decode("utf-8")
    mod = [m for m in out.split("\0") if m]
    same, diff = [], []
    for f in mod:
        if not os.path.exists(f):
            continue
        old = subprocess.run(["git", "show", "HEAD:" + f],
                             capture_output=True).stdout
        new = open(f, "rb").read()
        try:
            if f.endswith((".docx", ".xlsx", ".odt", ".zip")):
                eq = members(old) == members(new)
            elif f.lower().endswith(IMG):
                eq = img_bytes(old) == img_bytes(new)
            else:
                eq = old == new
        except zipfile.BadZipFile:
            eq = old == new
        (same if eq else diff).append(f)
    print("対象 %d ／中身が同じ %d ／変わった %d"
          % (len(mod), len(same), len(diff)))
    for f in diff:
        print("  *", f)
    if same and not dry:
        for i in range(0, len(same), 40):
            subprocess.run(["git", "checkout", "--"] + same[i:i + 40],
                           check=True)
    return 0


if __name__ == "__main__":
    sys.exit(main("--dry" in sys.argv))
