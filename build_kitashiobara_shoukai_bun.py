# -*- coding: utf-8 -*-
"""
北塩原村　村への照会文（5通）

出力: output/北塩原村_村への照会_5通.docx

何のために作るものか
  令和8年10月10日までの各点検・レビューで立てた村への照会は55件に
  なったが、村に出す文書そのものは1通も作っていなかった。
  照会を出さない限り、見込量も成果目標も委員会資料も確定しない。
  令和8年10月15日（木）の打合せで手渡せる形にする。

便の分け方（令和8年11月の委員会から逆算する）
  55件を1通にすると回答が遅れる。
  回答がいつ要るかで5通に分ける。
    第1便 委員会（令和8年11月）に出す資料に直に要るもの
    第2便 見込量の確定に要るもの（第2次算定の入力）
    第3便 高齢者施策との接続（対象となるかどうか）
    第4便 児童福祉との接続
    第5便 既存施策の実績（計画の確定には間に合わなくてよい）

  Ｍ-40（工程の確認）は文書にせず電話で確かめる。
  工程が分からないと、この5通の期限そのものが決められないため。

単一出所
  照会の中身は各点検・レビューの生成器に置いてある。
  本書はそれを読み込んで文書に組むだけで、文言を書き足さない。
  番号の抜け・重複・分類漏れは自己点検で落ちる。
"""

import os
import re
import sys

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_kitashiobara_kaigi import (  # noqa: E402
    h2, new_doc, note, para, table,
)
from build_kitashiobara_murashiryo import SHOUKAI as S_MURA  # noqa: E402
from build_kitashiobara_juten12 import SHOUKAI as S_JUTEN  # noqa: E402
from build_kitashiobara_hp_shisaku import SHOUKAI as S_HP  # noqa: E402
from build_kitashiobara_hp_3bunya import SHOUKAI as S_HP3  # noqa: E402
from build_kitashiobara_redteam_kyukyu import SHOUKAI as S_RT  # noqa: E402
from build_kitashiobara_mece_houshu import SHOUKAI as S_ME  # noqa: E402
from build_kitashiobara_mikomi_redteam import SHOUKAI as S_SA  # noqa: E402
from build_kitashiobara_zenkai_hyouka import SHOUKAI as S_ZEN  # noqa: E402
from build_kitashiobara_zuhyo import KAKUNIN as K_ZUHYO  # noqa: E402
from build_kitashiobara_zuhyo_bangou import KAKUNIN as K_BAN  # noqa: E402
from build_kitashiobara_anke_mikomi import KAKUNIN as K_ANKE  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_FILE = f"{REPO_ROOT}/output/北塩原村_村への照会_5通.docx"

UCHIAWASE = "令和８年10月15日（木）"
IINKAI = "令和８年11月"


# ===========================================================================
# 1. 照会を集めて形を揃える
#    古い3本は (番号, 件名, 内容, 使い道, 期限)
#    新しい5本は (番号, 宛先, 件名, 内容, 期限)
# ===========================================================================
def _zen(no):
    """M-40 と Ｍ-40 が混ざっているため全角に揃える。"""
    return no.replace("M-", "Ｍ-").replace("S-", "Ｓ-")


def atsumeru():
    out = {}
    for recs in (S_MURA, S_JUTEN, S_HP):
        for r in recs:
            out[_zen(r[0])] = dict(no=_zen(r[0]), atesaki="村", kenmei=r[1],
                                   naiyo=r[2], tsukaimichi=r[3], kigen=r[4])
    for recs in (S_HP3, S_RT, S_ME, S_SA, S_ZEN,
                 K_ZUHYO, K_BAN, K_ANKE):
        for r in recs:
            out[_zen(r[0])] = dict(no=_zen(r[0]), atesaki=r[1], kenmei=r[2],
                                   naiyo=r[3], tsukaimichi="", kigen=r[4])
    return out


# ===========================================================================
# 2. 便の分け方（番号を明示する。漏れたら自己点検で落ちる）
# ===========================================================================
BIN = [
    dict(
        no=1,
        midashi="委員会に出す資料に直に要るもの",
        kigen="令和８年10月24日（金）",
        riyu=f"{IINKAI}の委員会に、前回計画の評価と計画の素案を出します。"
             "評価の表のうち、いまの値が分からない欄が５つあります。"
             "ここが埋まらないと、基準値から目標までの進み具合を"
             "お示しできません。"
             "計画の名称と数値の食い違いも、資料に載せる前に"
             "確かめておく必要があります。",
        nos=["Ｍ-41", "Ｍ-42", "Ｍ-43", "Ｍ-44", "Ｍ-69", "Ｍ-70",
             "Ｍ-95", "Ｍ-96"],
    ),
    dict(
        no=2,
        midashi="サービスの見込量を確定するために要るもの",
        kigen="令和８年10月31日（金）",
        riyu="見込量は、過去の実績を延ばすのではなく、"
             "お一人ずつの状況を積み上げて定めます。"
             "とくに、重度訪問介護・同行援護・行動援護を"
             "令和９年度から令和11年度まで０人としている根拠は、"
             "いまは「利用の希望がない」ことだけです。"
             "対象となる状態の方がいらっしゃるのに"
             "制度を知らずに申し出ていない場合と、"
             "本当に必要がない場合とを分けられていません。",
        nos=["Ｍ-87", "Ｍ-88", "Ｍ-89", "Ｍ-90", "Ｍ-91",
             "Ｍ-92", "Ｍ-93", "Ｍ-94"],
    ),
    dict(
        no=3,
        midashi="高齢者福祉の施策との接続",
        kigen="令和８年11月14日（金）",
        riyu="障がいのある方の暮らしを支える仕組みは、"
             "障がい福祉の制度の中だけにあるわけではありません。"
             "村の高齢者福祉の事業に障がいのある方が含まれるのであれば、"
             "新しい事業を起こす前にそちらで対応できます。"
             "逆に含まれないのであれば、そこが本計画で埋めるべき"
             "すき間になります。",
        nos=["Ｍ-45", "Ｍ-46", "Ｍ-47", "Ｍ-48", "Ｍ-49", "Ｍ-50",
             "Ｍ-51", "Ｍ-52", "Ｍ-53", "Ｍ-54", "Ｍ-55", "Ｍ-56",
             "Ｍ-57", "Ｍ-58", "Ｍ-64", "Ｍ-65", "Ｍ-71", "Ｍ-72",
             "Ｍ-73", "Ｍ-81"],
    ),
    dict(
        no=4,
        midashi="児童福祉の施策との接続",
        kigen="令和８年11月14日（金）",
        riyu="本村には短期入所の事業所がありません。"
             "保護者の方が病気になられたときや出産のときに、"
             "障がいのある子どもをどこで預かるかは、"
             "障がい福祉の制度だけでは答えが出ません。"
             "村が令和８年度から始められた子育て短期支援事業で"
             "受け入れていただけるかどうかが、"
             "家族の支援の実際の分かれ目になります。",
        nos=["Ｍ-74", "Ｍ-75", "Ｍ-76", "Ｍ-77"],
    ),
    dict(
        no=5,
        midashi="既にある施策の実績",
        kigen="令和８年11月28日（金）",
        riyu="既にある施策については、"
             "「あること」ではなく「使われていること」を書きます。"
             "実績が分かれば、計画の評価と次の３年間の目標を"
             "数字で示せます。"
             "計画の確定にすぐ要るものではありませんので、"
             "他の便より後で差し支えありません。",
        nos=["Ｍ-59", "Ｍ-60", "Ｍ-61", "Ｍ-62", "Ｍ-63", "Ｍ-66",
             "Ｍ-67", "Ｍ-68", "Ｍ-78", "Ｍ-79", "Ｍ-80",
             "Ｍ-82", "Ｍ-83", "Ｍ-84", "Ｍ-85", "Ｍ-86"],
    ),
]

# 文書にせず電話で確かめるもの
DENWA = ["Ｍ-40"]

# 村以外に出すもの（この文書には入れない。他メンバーへの確認）
HOKA = ["Ｓ-19", "Ｓ-20", "Ｓ-21", "Ｓ-22", "Ｓ-23", "Ｓ-24",
        "Ｓ-25", "Ｓ-26"]


# ===========================================================================
# 3. 自己点検
# ===========================================================================
def verify(zenbu):
    ng = []
    mura = {k: v for k, v in zenbu.items() if v["atesaki"] == "村"}

    # ① 便に入れた番号が実在し、重複しないこと
    ireta = []
    for b in BIN:
        for no in b["nos"]:
            if no not in zenbu:
                ng.append(f"第{b['no']}便に実在しない番号: {no}")
            elif zenbu[no]["atesaki"] != "村":
                ng.append(f"第{b['no']}便に村以外あての番号: {no}")
            ireta.append(no)
    if len(ireta) != len(set(ireta)):
        dup = [x for x in set(ireta) if ireta.count(x) > 1]
        ng.append(f"便をまたいで重複している番号: {sorted(dup)}")

    # ② 村あての照会が、いずれかの便か電話に必ず入っていること
    nokori = set(mura) - set(ireta) - set(DENWA)
    if nokori:
        ng.append(f"どの便にも入っていない照会: {sorted(nokori)}")

    # ③ 村以外あてのものを村の照会文に入れていないこと
    for no in HOKA:
        if no in ireta:
            ng.append(f"村以外あてを村の照会文に入れている: {no}")

    # ④ 番号に抜けがないこと（Ｍ-40からＭ-96まで）
    ban = sorted(int(k.split("-")[1]) for k in mura)
    nuke = [n for n in range(ban[0], ban[-1] + 1) if n not in ban]
    if nuke:
        ng.append(f"村あての番号に抜けがある: {nuke}")

    # ⑤ 期限が便の順に遅くなっていること
    kigen = [b["kigen"] for b in BIN]
    if kigen != sorted(kigen, key=lambda s: (
            int(re.search(r"(\d+)月", s).group(1)),
            int(re.search(r"月(\d+)日", s).group(1)))):
        ng.append(f"便の期限が順に並んでいない: {kigen}")

    # ⑥ 中身が空でないこと
    for no, v in mura.items():
        for key in ("kenmei", "naiyo"):
            if not str(v[key]).strip():
                ng.append(f"{no}の{key}が空")

    # ⑦ 他団体の名前を出していないこと
    s = "\n".join(str(v) for v in zenbu.values()) + "\n".join(
        b["riyu"] for b in BIN)
    for g in ["小野町", "金ケ崎", "金ヶ崎", "阿蘇", "札幌"]:
        if g in s:
            ng.append(f"他団体の名前が入っている: {g}")

    if ng:
        print("自己点検 不合格:")
        for x in ng:
            print("   -", x)
        raise SystemExit(1)
    print(f"  自己点検: 村あて{len(mura)}件が5通{len(ireta)}件と"
          f"電話{len(DENWA)}件に過不足なく収まる／"
          f"番号Ｍ-{ban[0]}〜Ｍ-{ban[-1]}に抜けなし／"
          f"便をまたぐ重複なし／村以外あて{len(HOKA)}件を含めていない／"
          f"期限が便の順")
    return mura


# ===========================================================================
# 4. 文書
# ===========================================================================
def midashi_bin(doc, b, n):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(f"第{b['no']}便　{b['midashi']}（{n}件）")
    r.font.size = Pt(13)
    r.font.bold = True
    return p


def build(zenbu):
    doc = new_doc("北塩原村　第８期障がい福祉計画・第４期障がい児福祉計画",
                  f"村へのご照会（{UCHIAWASE}　ビズアップ公共コンサルティング"
                  "株式会社）")

    h2(doc, "このご照会について")
    para(doc,
         "計画の素案を作る中で、村にご確認いただきたい事項が"
         f"{len([v for v in zenbu.values() if v['atesaki'] == '村'])}件に"
         "なりました。"
         "一度にお送りするとご負担が大きく、"
         "お急ぎでないものに引きずられて急ぐものの回答が遅れます。"
         "このため、回答がいつ必要かによって５通に分けました。"
         "第１便から順にお願いできますと幸いです。")
    para(doc,
         f"{IINKAI}の委員会に前回計画の評価と計画の素案をお出しする"
         "予定とうかがっています。"
         "第１便は、その資料に直に載る数値です。"
         "第２便は、サービスの見込量をお一人ずつ積み上げて"
         "確定するために要るものです。"
         "第３便以降は、計画の確定には間に合えば足りるものです。")
    para(doc,
         "人数が少なく個人が特定されるおそれがある事項については、"
         "実数ではなく「該当あり」「該当なし」のみでも差し支えありません。"
         "その場合は計画にも実数を載せず、"
         "確かめた結果があることだけを書きます。")

    note(doc,
         "【お電話でお願いしたいこと】"
         "計画策定方針（令和８年５月25日）の工程について、"
         "第２回協議会の開催日、パブリックコメントの時期、"
         "第３回協議会（４町村広域）の開催日、庁議と議会への説明の時期が"
         "現在どうなっているかをお聞かせください。"
         "工程が決まらないと、このご照会の期限そのものを"
         "決められないためです（Ｍ-40）。")

    h2(doc, "５通の構成")
    rows = [["便", "内容", "件数", "回答の希望期限", "なぜその期限か"]]
    for b in BIN:
        rows.append([f"第{b['no']}便", b["midashi"], f"{len(b['nos'])}件",
                     b["kigen"], b["riyu"][:60] + "…"])
    table(doc, rows, widths=[1100, 4200, 900, 2400, 6800])

    for b in BIN:
        midashi_bin(doc, b, len(b["nos"]))
        para(doc, b["riyu"])
        rows = [["番号", "ご確認いただきたいこと", "内容",
                 "計画のどこに使うか"]]
        for no in b["nos"]:
            v = zenbu[no]
            rows.append([v["no"], v["kenmei"], v["naiyo"],
                         v["tsukaimichi"] or "―"])
        table(doc, rows, widths=[900, 3200, 8200, 3100])

    h2(doc, "回答をいただいた後に当方が行うこと")
    table(doc, [
        ["便", "回答後に進めること"],
        ["第１便",
         "前回計画の評価の表を埋め、委員会資料として確定します。"
         "計画の名称の食い違いは、第１章２の関連計画の記載と"
         "図（計画の位置づけ・他の計画との関係）の両方を直します。"],
        ["第２便",
         "サービスの見込量を第２次算定として組み直します。"
         "０人としているサービスについては、"
         "確かめた結果として０人とするか、見込みを立てるかを決めます。"],
        ["第３便",
         "高齢者福祉の事業のうち障がいのある方が対象となるものを"
         "計画に書き、対象とならないものは本計画で埋めるべき"
         "すき間として整理します。"],
        ["第４便",
         "一般の子育て施策で受けられる範囲と、"
         "障がい福祉で受ける範囲の切り分けを計画に書きます。"],
        ["第５便",
         "既にある施策の実績を計画に書き、"
         "次の３年間の目標の基準値とします。"],
    ], widths=[1200, 14200])

    note(doc,
         "【ご回答の様式について】"
         "様式は定めておりません。"
         "番号を添えてメール本文にご記入いただく形で差し支えありません。"
         "既存の資料で足りるものは、その資料をお送りいただくだけで結構です。"
         "分からない事項は「分からない」「記録していない」と"
         "お答えいただければ、それ自体を計画に書きます"
         "（記録がないことも、次の３年間で記録を始める根拠になります）。")

    doc.save(OUT_FILE)
    return doc


def main():
    zenbu = atsumeru()
    mura = verify(zenbu)
    doc = build(zenbu)
    print(f"作成: {OUT_FILE}")
    print(f"  段落{len(doc.paragraphs)}・表{len(doc.tables)}")
    print(f"  村あて{len(mura)}件＝5通{sum(len(b['nos']) for b in BIN)}件"
          f"＋電話{len(DENWA)}件")
    for b in BIN:
        print(f"  第{b['no']}便 {b['midashi']}　{len(b['nos'])}件　"
              f"{b['kigen']}")


if __name__ == "__main__":
    main()
