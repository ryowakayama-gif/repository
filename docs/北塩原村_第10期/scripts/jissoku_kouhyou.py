# -*- coding: utf-8 -*-
"""公表用（素案・計画書）の紙面を実測する

【なぜ要るか】
  kouhyou_map.kouhyou_pages() の推定は、節の高さを足しただけである。
  公表用は協議用の節を束ねて並べ替えるため、束ねるときに足す小見出し・
  基本目標の見出し・章の変わり目の改頁を推定が見ていない。
  現に推定108頁に対し実測117頁で、9頁外れた（doc109）。
  素案を直したら、推定ではなくこれで測り直して
  kouhyou_build.JISSOKU_KOUHYOU・JISSOKU_KEIKAKU に貼る。

【使い方】
  python3 scripts/jissoku_kouhyou.py              # 素案と4通りの案を測る
  python3 scripts/jissoku_kouhyou.py ABCDE        # 案を指定して測る
"""
import os
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.dont_write_bytecode = True
import kouhyou_build as KB
import paths as P
import render_check as RC

MAE = 3          # 前付（目次は空のまま出る）。PDF の頁数から差し引く


def hakaru(an, keikaku, work):
    """docx を組んで PDF に変換し、本文の頁数を返す"""
    nm = ("keikaku" if keikaku else "soan") + "_" + (an or "nashi")
    js = os.path.join(work, nm + ".json")
    KB.to_json(out_path=js, an=an)
    dx = os.path.join(work, nm + ".docx")
    # build_soan_docx.js は --keikaku のあとに保存先を取る（--out は素案の保存先）
    cmd = ["node", P.base("scripts", "build_soan_docx.js"), "--in", js]
    cmd += ["--keikaku", dx] if keikaku else ["--out", dx]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL)
    import pymupdf
    return pymupdf.open(RC.convert(dx, os.path.join(work, nm))).page_count - MAE


def main():
    ans = [a.replace("なし", "") for a in sys.argv[1:] if not a.startswith("-")]
    if not ans:
        ans = list(KB.JISSOKU_KEIKAKU)
    work = tempfile.mkdtemp(prefix="kouhyou_")
    print("■ 公表用の実測（組んで変換して数える）\n")
    n = hakaru("", False, work)
    mae = KB.JISSOKU_KOUHYOU.get("")
    print("  JISSOKU_KOUHYOU = {")
    print('      "":%12s%d,   # 素案（注記を含む）＝%d頁%s'
          % ("", n, n + KB.MAE_PG,
             "" if mae is None else ("　（前は %d）" % mae if mae != n else "　（前と同じ）")))
    print("  }")
    print("  JISSOKU_KEIKAKU = {")
    for a in ans:
        k = hakaru(a, True, work)
        m = KB.JISSOKU_KEIKAKU.get(a)
        print('      "%s":%s%d,   # ＝%d頁%s'
              % (a, " " * max(1, 12 - len(a)), k, k + KB.MAE_PG,
                 "" if m is None else ("　（前は %d）" % m if m != k else "　（前と同じ）")))
    print("  }")
    print("\n  上を kouhyou_build.py の JISSOKU_KOUHYOU・JISSOKU_KEIKAKU に貼る。")
    print("  組んだ docx と PDF： %s" % work)
    return 0


if __name__ == "__main__":
    sys.exit(main())
