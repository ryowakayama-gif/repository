# -*- coding: utf-8 -*-
"""成果品が作業領域の消失によって失われないかを点検する.

作業領域（セッション固有の場所）は、環境が作り直された時点で消える。
そこにしか本文・数値を置いていない成果品は、同じ内容を作り直せない。

本点検は、成果品ごとに次の3つを確かめる。

  ① 実体が output/ にあり、git の追跡下にあること
  ② 作成に用いた処理が git の追跡下にあること
  ③ 作業領域を読む処理によるものは、①により内容が担保されていること

いずれかを欠くものを不適合とし、1件でもあれば終了コード1で終わる。

併せて、次の2つを数える。

  ・出力先を絶対パスで固定している処理（`RP.ROOT` によらないもの）
  ・作業領域を読む処理

使い方
  python3 tools/jizoku_check.py
"""

import os
import re
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import repo_paths as RP                                       # noqa: E402
import data_dispatch as DD                                    # noqa: E402

ROOT = RP.ROOT

# 作業領域（セッション固有の場所）を指す書き方
RE_SESSION = re.compile(r"/root/\.claude/uploads|/tmp/claude-|scratchpad")
# 出力先を絶対パスで固定している書き方（`RP.ROOT` によらないもの）
RE_ABS_OUT = re.compile(r'"/home/user/repository/output')


def _tracked():
    """git の追跡下にあるファイルの集合を返す。"""
    out = subprocess.run(["git", "-C", ROOT, "ls-files", "-z"],
                         capture_output=True, check=True).stdout
    return set(out.decode("utf-8").split("\0")) - {""}


def _scripts():
    return sorted(f for f in os.listdir(ROOT)
                  if f.startswith("build_") and f.endswith(".py"))


def main():
    tracked = _tracked()
    scripts = _scripts()
    src = {}
    for f in scripts:
        with open(os.path.join(ROOT, f), encoding="utf-8") as fh:
            src[f] = fh.read()

    # 成果品を書き出している処理を作成元とみる。
    # 名が現れるだけのもの（一覧に掲げているだけのもの）は作成元ではない。
    # 引用符・空白・改行を落として `/output/<名>` 又は
    # `RP.OUTPUT,<名>` の形で探す。
    flat = {f: re.sub(r"[\"'\s]", "", s) for f, s in src.items()}
    maker = {}
    for name in DD.DISPATCH:
        # 源も名も引用符・空白・改行を落として比べる
        # （名そのものに空白を含むものがある）
        fn = re.sub(r"[\"'\s]", "", name)
        pat = re.compile(r"join\([A-Za-z_.]+," + re.escape(fn) + r"\)")
        maker[name] = [f for f in scripts
                       if ("output/" + fn) in flat[f] or pat.search(flat[f])]
        if not maker[name] and fn.endswith(".pdf"):
            # PDF は同名の xlsx 又は docx から変換したものである
            for x in (fn[:-4] + ".xlsx", fn[:-4] + ".docx"):
                px = re.compile(r"join\([A-Za-z_.]+," + re.escape(x) + r"\)")
                maker[name] = [f for f in scripts
                               if ("output/" + x) in flat[f]
                               or px.search(flat[f])]
                if maker[name]:
                    break

    ng, warn, ok = [], [], 0
    for name in sorted(DD.DISPATCH):
        if DD.DISPATCH[name][0] == DD.GAI:
            continue
        rel = "output/" + name
        has = rel in tracked
        mk = maker[name]
        # 点検のための走査（`SESSION_PATH`）は作業領域を読むことに当たらない
        sess = [f for f in mk
                if RE_SESSION.search(src[f]) and "SESSION_PATH" not in src[f]]
        untracked_mk = [f for f in mk if f not in tracked]
        if not has:
            ng.append((name, "実体が追跡下にない"))
            continue
        if not mk:
            # 作成に用いた処理を名から辿れないもの。
            # 実体は追跡下にあるため失われないが、作り直しの手掛かりがない。
            warn.append((name, "作成に用いた処理を名から辿れない"))
            ok += 1
            continue
        if untracked_mk:
            ng.append((name, "作成に用いた処理が追跡下にない（%s）"
                       % "、".join(untracked_mk)))
            continue
        if sess:
            # 作業領域を読むものは作り直せないが、実体が追跡下にあるため
            # 内容は失われない。
            warn.append((name, "作業領域を読むため作り直せない（実体は追跡下）"))
        ok += 1

    abs_out = [f for f in scripts if RE_ABS_OUT.search(src[f])]
    sess_all = sorted(f for f in scripts if RE_SESSION.search(src[f]))

    print("■ 成果品の持続性の点検")
    print("  適合 %d件／不適合 %d件" % (ok, len(ng)))
    for name, why in ng:
        print("")
        print("【不適合】%s　%s" % (name, why))
    if warn:
        print("")
        print("［注記］%d件" % len(warn))
        for name, why in warn:
            print("  %s　%s" % (name, why))

    print("")
    print("■ 出力先を絶対パスで固定している処理　%d件" % len(abs_out))
    for f in abs_out:
        print("  %s" % f)
    print("")
    print("■ 作業領域を読む処理　%d件" % len(sess_all))
    for f in sess_all:
        print("  %s%s" % (f, "（点検の走査のみ）"
                          if "SESSION_PATH" in src[f] else ""))
    return 1 if ng else 0


if __name__ == "__main__":
    sys.exit(main())
