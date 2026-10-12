# -*- coding: utf-8 -*-
"""いま進められる作業の洗い出し。

【なぜ必要か】
  未完了のWBSは、翌営業日の作業（READY）か、着手しない理由（HOLD_REASON）の
  どちらかを持つ（点検14）。だが理由は書いた時点のものであり、**事情が変わっても
  理由だけが残る**。現に「PDFは変換ができる環境を要する」（Ⅳ-91）は、
  令和8年10月12日に変換ができるようになった後も残っていた。

  この道具は、未完了のWBSを「何を待っているか」で仕分ける。待っているものが
  当方の側で用意できるもの（道具・手順・記述）であれば、それは**待ちではなく
  着手できる作業**である。

【仕分けの軸】
  村・県の回答／相手の計画／告示・政令／日程・内部手続／データの到着／
  当方の作業（＝進められる）

  語で仕分けるため取りこぼしはある。**一覧を見て人が判断するための下ごしらえ**で
  あって、判断そのものではない。

【使い方】
  python3 scripts/wbs_susumeru.py
"""
import sys
sys.dont_write_bytecode = True
import os, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 待っている相手の見分け（上から順に当てる）
MACHI = [
 ("告示・政令", r"告示|政令|省令|通知を待|制度として定まって"),
 ("相手の計画", r"障がい福祉計画|障害福祉計画|県の整備量|相手の計画"),
 ("委員会・議会・日程", r"策定委員会の開催|議会|日程|開催を待|委嘱|改選|パブリックコメントの時期"),
 ("村の内部手続", r"契約|決裁|所管課|内部手続"),
 ("データの到着", r"回収|到着|入力データ|受領を待|提供を待|取込"),
 ("村・県の回答", r"ご確認を待|確認を待|回答を待|判断を待|ご回答|village|束\d"),
 ("前の工程", r"が前提|同上|の結果が前提|確定を待"),
]


def machi_saki(riyu):
    for nm, pat in MACHI:
        if re.search(pat, riyu):
            return nm
    return "（仕分けできない）"


def run():
    import wbs_pending as PD
    import wbs_progress as PR
    import wbs_kakunin as KK

    # 影響度1（停止）に結びつく確認事項の見出し
    tome = [k for k, v in PD.IMPACT.items() if v[0] == 1]

    ready = {r[1] for r in PD.READY}
    hold = PD.HOLD_REASON
    group = {}
    for wid, riyu in sorted(hold.items()):
        group.setdefault(machi_saki(str(riyu)), []).append((wid, str(riyu)))

    print("■ 未完了のWBSが何を待っているか\n")
    print(f"  翌営業日の作業（READY）　{len(ready)}件　{'・'.join(sorted(ready))}")
    print(f"  着手しない理由（HOLD_REASON）　{len(hold)}件\n")
    for nm, _ in MACHI + [("（仕分けできない）", "")]:
        if nm not in group:
            continue
        g = group[nm]
        print(f"  {nm}　{len(g)}件")
        print(f"    {'・'.join(w for w, _ in g)}")
    print()
    print(f"■ 止まっているもの（影響度1）　{len(tome)}件\n")
    for t in tome:
        print(f"  ・{t[:66]}")
    print()
    print("■ 理由のなかに「当方で用意できる」と読めるものがある項目\n")
    print("   （待ちと書かれていても、道具・手順・記述は当方で進められることがある）")
    JIBUN = r"当方で|作れる|着手できる|組める|整えられる|進められる|できるようになった"
    atta = False
    for wid, riyu in sorted(hold.items()):
        if re.search(JIBUN, riyu):
            atta = True
            print(f"  {wid}　{riyu[:90]}")
    if not atta:
        print("  なし")
    print()
    print("■ 進捗が0.9を超えていて、あと一歩のもの\n")
    for nm, v in sorted(PR.P.items(), key=lambda x: -x[1][1]):
        if v[0] == "着手" and v[1] >= 0.90:
            print(f"  {v[1]:.2f}  {nm[:40]}")


if __name__ == "__main__":
    run()
