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


def members(b):
    z = zipfile.ZipFile(io.BytesIO(b))
    return {n: z.read(n) for n in z.namelist() if n not in SKIP}


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
            eq = (members(old) == members(new)
                  if f.endswith((".docx", ".xlsx", ".odt", ".zip"))
                  else old == new)
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
