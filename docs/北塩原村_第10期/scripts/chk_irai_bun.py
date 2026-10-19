# -*- coding: utf-8 -*-
"""照会票の優先度S・Aの件について、村にお見せする文を組み立てて見る（WBS Ⅴ-100）

【なぜ必要か】
  照会票の「当方の現在の扱い」は `irai_bundle.yasashiku()` が文の単位で
  170字まで出す。**背景から書き始めた件は、お願いが170字の外へ出て村に届かない。**
  見出し（尋ねたいことそのもの）だけで何をすればよいか読める件は直さなくてよいが、
  見出しが状況の説明で、お願いが本文の後ろにある件は直さなければならない。

  資6の6件は令和8年10月12日に直した（doc85）。残りは見ていなかった。

【この道具が見るもの】
  ア 本文がお願いの文で始まっているか（文末が依頼の形か）
  イ お願いの文が170字の中に入っているか（切れていないか）
  ウ 見出しだけで何をすればよいか読めるか（読めるなら本文は背景でよい）

  python3 scripts/chk_irai_bun.py          優先度S・Aの件を一覧で示す
  python3 scripts/chk_irai_bun.py --all    すべての件を示す
  python3 scripts/chk_irai_bun.py --ng     直す必要のある件だけ示す
"""
import os
import re
import sys
sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import irai_bundle as IB
import wbs_kakunin as KK

# お願いの形で終わる文。照会票の相手は村であるため、依頼の語で見分ける
ONEGAI = re.compile(
    r"(いただきたい|いただけますか|いただけると|ご教示|ご確認|ご回答|ご提供|ご指示|"
    r"お願いしたい|お願いします|お知らせください|確認されたい|伺いたい|"
    r"お尋ねしたい|提供を求める|ご判断|ご意向|お示しください)")

# 見出しが「言い切りの文」になっているもの。
#
#   照会票の表の列見出しは「ご確認・ご提供をお願いしたいこと」である。
#   したがって見出しが名詞句であれば、それだけで何をお願いしているかは通じる
#   （「契約締結日・契約書の取り交わし状況」＝これをご提供ください、と読める）。
#   通じないのは、見出しが状況を述べる文で終わっている場合だけである
#   （「…は見える化に系列が一切ない」では、村が何をすればよいか分からない）。
#
#   【注意】この並びに裸の $ を入れてはいけない。$ はどの文字列にも当たるため、
#   すべての見出しが当たり、走査が何も見つけなくなる。
#   令和8年10月19日、自己試験（お願いの文を抜いても落ちない）で見つけて外した。
IIKIRI = re.compile(
    r"(である|ている|ており|いない|ない|した|する|なる|食い違う|"
    r"分からない|できない|残る|要する|当たる)$")


def bun(t):
    """村にお見せする本文（yasashiku が出すもの）"""
    return IB.yasashiku(t)


def saisho(t):
    """本文の1文目"""
    return bun(t).split("。")[0] + "。"


def onegai_ichi(t):
    """お願いの文が本文の何文目にあるか。無ければ None"""
    for i, sn in enumerate(bun(t).split("。"), 1):
        if ONEGAI.search(sn):
            return i
    return None


def kireta(k):
    """お願いの文が 170字で切り落とされているか。

    もとの本文（切る前）にお願いがあり、お見せする本文に無ければ切れている。
    """
    nama = IB.yasashiku(k[2], limit=10 ** 6)
    return bool(ONEGAI.search(nama)) and not bool(ONEGAI.search(bun(k[2])))


def midashi_de_yomeru(k):
    """見出しだけで何をお願いしているかが読めるか。

       表の列見出しが「ご確認・ご提供をお願いしたいこと」であるため、
       名詞句ならば読める。言い切りの文で終わっているものだけが読めない。
    """
    m = IB.midashi(k[1]).rstrip("）)").split("（")[0].strip()
    return not IIKIRI.search(m)


def judge(k):
    """(しるし, 理由) を返す。

       ✕ 直す　　お願いの文はあるのに、170字の外へ出て村に届かない。
                 これは目で見なくても決まる（切れているかどうかの話である）。
       △ 見る　　本文にお願いの文がなく、見出しも尋ねたいことの形になっていない。
                 照会票は資料の提供をお願いする文書であり、見出しが
                 「令和6年度の決算」のように要るものの名であれば、
                 本文にお願いの文がなくても村には伝わる。
                 名でも形でもないものだけを挙げる。
       ○ そのまま
    """
    if kireta(k):
        return "✕", "お願いの文が170字の外に出ており、村に届かない"
    i = onegai_ichi(k[2])
    if i == 1:
        return "○", "1文目がお願いになっている"
    if i is not None:
        return "○", "お願いが%d文目にあり、170字の中に入っている" % i
    if midashi_de_yomeru(k):
        return "○", "本文にお願いの文はないが、見出しが尋ねたいことの形になっている"
    return "△", "本文にお願いの文がなく、見出しも尋ねたいことの形になっていない"


def taishou(zenbu=False):
    """優先度S・Aの未解決の件（--all ならすべての未解決の件）"""
    out = []
    for k in KK.K:
        if k[-1] == "解決済":
            continue
        if not zenbu and k[-2] not in ("S", "A"):
            continue
        out.append(k)
    return out


def main():
    zenbu = "--all" in sys.argv
    ng_dake = "--ng" in sys.argv
    ks = taishou(zenbu)
    res = [(k,) + judge(k) for k in ks]
    ng = [r for r in res if r[1] == "✕"]
    mi = [r for r in res if r[1] == "△"]
    print("■ 照会票に出る文の見直し（%s）" %
          ("未解決すべて" if zenbu else "優先度S・A"))
    print("  対象 %d件　直す %d件／目で見る %d件／そのまま %d件\n"
          % (len(res), len(ng), len(mi), len(res) - len(ng) - len(mi)))
    for k, mark, why in ((ng + mi) if ng_dake else res):
        print("%s [%s] %s" % (mark, k[-2], IB.midashi(k[1])[:62]))
        print("      %s" % why)
        print("      本文の1文目：%s" % saisho(k[2])[:88])
        print()
    return 1 if ng else 0


if __name__ == "__main__":
    sys.exit(main())
