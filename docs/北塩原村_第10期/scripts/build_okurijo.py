# -*- coding: utf-8 -*-
"""障がい福祉計画担当への申し送りの中身を JSON に落とす。
   docx 側（build_okurijo_docx.js）がこれを読む。

【なぜ必要か】
  同時に策定している第8期障がい福祉計画・第4期障がい児福祉計画との突合が、
  互いに古い版を見て行われていた。相手の照会Ｋ-7（令和8年10月8日）は、
  当方の成果物の中に置かれただけで当方の照会票に行として立っていなかったため
  届かなかった。相手は令和8年10月12日、同じ受け渡しをやり直す申し送りを
  当方の様式に合わせて送ってきた。

  **当方から相手へ渡す事項も、同じ形で渡さなければ落ちる。**
  口頭や文書の往復では落ちることは、Ｋ-7で分かっている。

  中身は seigo_han が持つ。じか書きしない。
  WATASHI に1件足せば申し送りにも現れ、点検75が節の実在を確かめる。

【使い方】
  python3 scripts/build_okurijo.py           中身を書き出す
  node    scripts/build_okurijo_docx.js      docx を組む
"""
import io
import json
import os
import sys
sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths as P
import seigo_han as SH

ZEN = "〇１２３４５６７８９"      # 節の番号は全角で振る（相手の様式に合わせる）


def zen(i):
    return "".join(ZEN[int(c)] for c in str(i))


META = {
 "title": "第10期北塩原村高齢者福祉計画・",
 "title2": "第10期北塩原村介護保険事業計画策定業務",
 "subtitle": "障がい福祉計画担当への申し送り",
 "date": "令和8年10月13日",
 "issuer": "ビズアップ公共コンサルティング株式会社",
 "atesaki": "第8期障がい福祉計画・第4期障がい児福祉計画　ご担当者様",
 "footer": "ビズアップ公共コンサルティング株式会社",
}


def build():
    out = {"meta": META, "maegaki": SH.MAEGAKI, "sections": []}

    # ■1　版の記録
    out["sections"].append({
     "no": zen(1),
     "title": "版の記録",
     "lead": "どの版を見た突合なのかが後から分かるよう、ブランチ・コミット・取得日・"
             "内容のハッシュ（sha256 の先頭16桁）を記録しております。"
             "記録した版は、読むたびに現物と照合しております。",
     "head": ["区分", "コミット", "日", "内容のハッシュ", "資料"],
     "widths": [20, 11, 13, 20, 36],
     "rows": [["貴計画　" + a["nm"].split("（")[-1].rstrip("）")
               if "（" in a["nm"] else "貴計画　整合の照会",
               a["han"], a["han_hi"], a["naiyou"] or "―", a["shiryo"]]
              for a in SH.AITE if a["naiyou"]]
             + [["本計画　貴側が読まれた版", m["han"], m["han_hi"],
                 m["naiyou"] or "―", m["path"] or "―"] for m in SH.MIRARETA],
     "note": "当方の最新の版は、本書の作成時点で貴側がお読みになった e591fe8 より"
             "2コミット進んでおります（第1部）。"
             "本計画の素案の内容のハッシュは、"
             "git show <コミット>:docs/北塩原村_第10期/18_計画素案.md | sha256sum "
             "の先頭16桁で取り出せます。",
    })

    # ■2　貴側の申し送り§2・§3への当方の対応
    out["sections"].append({
     "no": zen(2),
     "title": "貴側のご指摘と、当方の確認事項へのご回答への対応",
     "lead": "貴側の申し送り§2（当方の整理の前提のうち当たらないもの）と"
             "§3（当方の確認事項3件へのご回答）について、当方が行ったことです。",
     "head": ["貴側のお示し", "当方の対応"],
     "widths": [42, 58],
     "rows": [[k["nm"] + "（" + k["yu"] + "）", k["touhou"]] for k in SH.KOTAE],
     "note": "doc81 §1-2 の「貴計画の第7章に連携の節は置かれていなかった」および"
             "§4 の「現在はどちらにも無い」は、当方が令和8年10月4日の版を見て"
             "書いたもので当たりません。訂正いたしました。"
             "あわせて当方の施策一覧の条文が「介護保険法第115条の45第2項第2号」と"
             "なっていた誤りを第3項第3号に正しました"
             "（第2項第2号は権利擁護事業でございました）。",
    })

    # ■3　貴側が立ててほしいと言われた5件の状態
    out["sections"].append({
     "no": zen(3),
     "title": "貴側の申し送り§4の5件の状態",
     "lead": "貴側が当方の照会票に立ててほしいとお示しくださった5件です。"
             "貴側がお読みになった e591fe8 より当方が2コミット進んでいたため、"
             "3件は既に反映済みでございました。",
     "head": ["優先度", "事項", "当方の状態", "反映した節"],
     "widths": [8, 36, 44, 12],
     "rows": [[o["yu"], o["nm"], o["sumi"] + "　" + o["memo"], "・".join(o["saki"])]
              for o in SH.OKURI],
     "note": "受給者証については、貴側の第7章3の表11件のうち通所支援受給者証（児童）と"
             "自立支援医療受給者証が本計画と接点がないため、本計画の見込量と"
             "被保険者資格に関わる5つに絞って資9に置きました。"
             "11行の表を新設しなかったのは、本計画の紙面が村が許容する110頁に対し"
             "109頁で余地がないことも理由でございます。",
    })

    # ■4〜　当方から渡す事項（部ごと）
    #   7列の表にすると1列が数文字になって読めない。事項ごとの縦組みにする。
    #   相手の照会票に貼るときは、見出しの札がそのまま列の名に対応する。
    for bu in sorted(SH.BU):
        midashi, lead = SH.BU[bu]
        naka = [w for w in SH.WATASHI if w["bu"] == bu]
        out["sections"].append({
         "no": zen(3 + bu),
         "kind": "jikou",
         "title": midashi,
         "lead": lead,
         "items": [{
            "midashi": f"［優先度 {w['yu']}・{w['kubun']}］　{w['nm']}",
            "rows": [["当方が確かめた事実", w["jijitsu"]],
                     ["お願いしたいこと", w["onegai"]],
                     ["ご回答がない場合の当方の扱い", w["nai"]],
                     ["なぜ必要か", w["naze"]],
                     ["由来／本計画で関わる節",
                      w["yurai"] + "　／　" + "・".join(w["saki"])]],
         } for w in naka],
         "note": "",
        })

    # 最後　受け渡しの約束
    out["sections"].append({
     "no": zen(4 + max(SH.BU)),
     "title": "受け渡しの約束と、当方が守る場所",
     "lead": "貴側の申し送り§5の4つの約束について、当方の側で守る場所を決めました。"
             "いずれも当方の生成器の中に置き、組むたびに確かめております。",
     "head": ["約束", "中身", "当方が守る場所"],
     "widths": [20, 54, 26],
     "rows": [[nm, naka, doko] for nm, naka, doko in SH.YAKUSOKU],
     "note": "当方は貴ブランチへ書き込みません。別業務であり、本書は申し送りであって"
             "貴側の成果物を直すものではありません。"
             "個人が特定される情報は求めておりません。"
             "村への照会が二重になる事項は貴側の便に相乗りさせていただき、"
             "ご回答は村のご了解のうえ村から両業務へ同じものをいただく形を"
             "希望いたします。",
    })
    return out


def soto_muki(out):
    """宛先に向けた文書として、言い方が内向きになっていないかを見る。

    seigo_han は当方の控えでもあるため「相手」という言い方が混じりやすい。
    本書は**その相手にお渡しするもの**であり、「相手」と書くと誰を指すのか
    分からなくなる。組むたびに確かめる。
    """
    warui = []
    for sec in out["sections"]:
        hako = []
        for r in sec.get("rows", []):
            hako += [str(x) for x in r]
        for it in sec.get("items", []):
            hako.append(it["midashi"])
            hako += [str(x) for r in it["rows"] for x in r]
        hako += [sec.get("lead", ""), sec.get("note", ""), sec.get("title", "")]
        for t in hako:
            for go in ("相手", "当方の別業務"):
                if go in t:
                    warui.append("%s：%s …%s…" % (sec["no"], go, t[:36]))
    return warui


def main():
    out = build()
    warui = soto_muki(out)
    if warui:
        print("★内向きの言い方が残っている（宛先に向けた文書である）")
        for w in warui:
            print("   ", w)
        return 1
    f = P.build("okurijo.json")
    io.open(f, "w", encoding="utf-8").write(json.dumps(out, ensure_ascii=False, indent=1))
    print("JSON:", f)
    for s in out["sections"]:
        n = len(s.get("items", s.get("rows", [])))
        print("  %s　%s　%d件" % (s["no"], s["title"], n))
    n = sum(len(s["items"]) for s in out["sections"] if s.get("kind") == "jikou")
    print("  当方から渡す事項 %d件" % n)
    assert n == len(SH.WATASHI), "渡す事項の数が seigo_han.WATASHI と合わない"
    return 0


if __name__ == "__main__":
    sys.exit(main())
