# -*- coding: utf-8 -*-
"""紙面の洗い出しの結果を記録し、悪くなっていないかを見る（WBS Ⅱ-75）

【なぜ必要か】
  令和8年10月10日（doc83）まで、この環境では docx を PDF に変換できなかった。
  そのため**紙面（改頁・表の分割・図の配置）は一度も確かめていなかった。**
  変換ができるようになったので、令和8年10月20日に初めて通して見た。

  洗い出しそのものは `render_check.py --toushi` が行う。このファイルは
  **そのときの結果を控えておき、次に測ったとき悪くなっていないかを見る**。
  素案は毎回変わるため、頁の番号は動く。見るのは件数である。

【游ゴシックがないことの意味】
  この環境には游ゴシックがなく、IPAゴシックに置き換わる。
  **したがって、ここで見ている改頁は村が Word で見る改頁ではない。**
  分かるのは、はみ出し・小さすぎる字・図と題の分離といった
  **フォントにあまり左右されないもの**である。
  孤立行（表が頁をまたいで1行だけ残る）はフォントで動くため、
  **件数が増えていないことだけを見て、頁を名ざして追わない。**
  納品のPDFは游ゴシックの入った環境で作り直す（build_nohin.py が申し送る）。

  python3 scripts/chk_shimen.py          控えを示す
  python3 scripts/chk_shimen.py --measure 測り直して控えと比べる
"""
import os
import re
import subprocess
import sys
sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths as P

SOAN = P.out("06_北塩原村第10期_計画素案.docx")

# 令和8年10月21日に測った結果（本文152頁。素案＝5-12と編集注記を含むもの）
#   測り直して増えていたら、増えた分を見る。
#   10月20日に本文145頁から150頁に伸びた。村単独事業の突合せで4-6(2)に3事業を
#   加えて147頁になり、さらに老人福祉法の原文に当たって4-6の対応表と
#   1-6の県の意見聴取を加えて150頁になった（doc106）。
#   10月21日に151頁になった。例規集の第8類第4章 介護保険を突合せ、
#   施策1-2に平成29年訓令第6号・第7号、施策4-6に令和4年訓令第2号、
#   4-6(2)に家族介護慰労金・住宅改修費の理由書作成を加えたことによる（doc110）。
#   さらに同日、他3町のブランチと突き合わせて交付金の試算を直し、
#   素案2-8に「条件の近い団体との比較」を加えて152頁になった（doc111）。
#   協議用に分量の制約はないため、伸びること自体は問題としない。
#   孤立行は145頁の時点で2件あり、本文が伸びて改頁の位置が動くたびに
#   2件→1件→0件と減った。**本文を直して消したものではない。**
#   5項目すべてが0になったが、本文が伸び縮みすれば孤立行は戻る。
#   戻ったときの扱いは下の KETTEI のとおりである。
HIKAE = {
    "はみ出し": 0,
    "孤立行": 0,
    "見出しの取り残し": 0,
    "図の離れ": 0,
    "小さすぎる字": 0,
}
HIKAE_HI = "令和8年10月21日"
HIKAE_PG = 152

# 孤立行について、見て決めたこと
#   本文145頁の時点では2件あった。
#     2-7(7)「本村の認知症の状況」の表の最後の1行
#     資7「主な参照資料」の表の最後の1行
#   いずれも、前の節が伸びて改頁の位置が動いたことで消えた。
#   頁を名ざして追わない（上の【游ゴシックがないことの意味】のとおり）。
#
# 孤立行が戻った場合も直さない。表頭が次の頁にも繰り返されるため
# （build_soan_docx.js の tableHeader: true）、読む人は見出しつきで1行を読む。
# 表をまるごと次の頁へ送れば孤立は消えるが、前の頁に空白が1頁近く残る。
# 本計画は仕様書の目安100頁を上回っており、空白を増やす余地がない。
# 納品の前に、游ゴシックの入った環境で作り直したPDFで改めて見る。
KETTEI = (
 "5項目すべてが0件である。孤立行は本文が伸びて改頁の位置が動いたことで消えたもので、"
 "本文を直して消したのではない。本文が伸び縮みすれば戻る。"
 "戻った場合も直さない。表頭が繰り返されるため読めること、"
 "表をまるごと送ると空白が1頁近く残ること、"
 "游ゴシックの環境では改頁の位置が変わることによる。"
 "納品の前に、游ゴシックで作り直したPDFで改めて見る。")


def measure():
    """render_check.py --toushi を走らせて件数を読む"""
    r = subprocess.run([sys.executable,
                        os.path.join(P.SCRIPTS, "render_check.py"),
                        SOAN, "--toushi"],
                       capture_output=True, text=True, timeout=1800)
    out = r.stdout + r.stderr
    if "紙面の洗い出し" not in out:
        raise SystemExit("洗い出せなかった\n" + out[-1500:])
    n = {}
    for k in HIKAE:
        m = re.search(re.escape(k) + r"　(\d+)件", out)
        n[k] = int(m.group(1)) if m else None
    m = re.search(r"紙面の洗い出し（本文(\d+)頁）", out)
    return n, (int(m.group(1)) if m else None), out


def main():
    print("■ 紙面の洗い出しの控え（%s・本文%d頁）\n" % (HIKAE_HI, HIKAE_PG))
    for k, v in HIKAE.items():
        print("    %-14s %d件" % (k, v))
    print("\n  %s" % KETTEI)
    if "--measure" not in sys.argv:
        print("\n  （控えのみ。測り直すには --measure を付ける）")
        return 0
    print("\n■ 測り直し")
    ima, pg, out = measure()
    bad = []
    for k, v in HIKAE.items():
        now = ima.get(k)
        shirushi = "―" if now is None else ("同じ" if now == v
                                            else ("増えた" if now > v else "減った"))
        print("    %-14s 控え%d件 → いま%s件　%s"
              % (k, v, "―" if now is None else now, shirushi))
        if now is None:
            bad.append("%s を読み取れない" % k)
        elif now > v:
            bad.append("%s が %d件→%d件 に増えた" % (k, v, now))
    print("\n    本文 控え%d頁 → いま%s頁" % (HIKAE_PG, pg))
    if bad:
        print("\n  ✕ " + "／".join(bad))
        print("  増えた分を見て、直すか控えを改めるかを決める")
        return 1
    print("\n  ○ 控えより悪くなっていない")
    return 0


if __name__ == "__main__":
    sys.exit(main())
