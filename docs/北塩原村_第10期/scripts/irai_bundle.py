# -*- coding: utf-8 -*-
"""照会票（束ごと）の中身
   村が右端の欄に書き入れて返すだけで答えが揃う様式にする。

   中身は wbs_kakunin.K と wbs_pending.BUNDLE／IMPACT から導く。
   じか書きしない。確認事項を1件足したら照会票にも現れる。

   ・1件の確認事項は1つの束にだけ入れる。関連WBSの重なりが最も多い束に割り当て、
     同数なら先の束に入れる。2つの束に同じことを書くと、村に二度尋ねることになる。
   ・影響度0（社内）は入れない。当方で確かめれば済むものである。
   ・束7（調査工程）は他メンバーの担当のため、この照会票では扱わない。
"""
import os as _os_p
import re
import sys as _sys_p
_sys_p.path.insert(0, _os_p.path.dirname(_os_p.path.abspath(__file__)))
import paths as _P                     # noqa: F401  （置き場所はここで決める）
import wbs_kakunin as KK
import wbs_pending as PD

MACHI = ("未依頼", "一部回答", "再提起")      # まだ答えをいただいていないもの
YUSEN = {"S": 0, "A": 1, "B": 2, "C": 3}

META = {
 "title": "第10期北塩原村高齢者福祉計画・",
 "title2": "第10期北塩原村介護保険事業計画策定業務",
 "subtitle": "照会票",
 "date": "令和8年10月7日",
 "issuer": "ビズアップ公共コンサルティング株式会社",
 "atesaki": "北塩原村役場　保健福祉課　福祉係　御中",
}

# 影響度ごとの部立て（0と5はこの票では扱わない）
PARTS = [
 (1, "第1部　【至急】ご回答がないと策定の作業が進まない事項",
     "下表の事項は、回答がないと保険料・見込量の算定を確定できず、以降の工程が止まります。"
     "お手数ですが、他の部より先にご回答をお願いいたします。"),
 (2, "第2部　【仮置き中】当方で仮定を置いて進めている事項",
     "いずれも当方の仮定で算定を進めています。右端の欄にご回答をいただいた時点で算定し直し、"
     "成果品に反映します。仮定の内容は「当方の現在の扱い」の欄に記しています。"),
 (3, "第3部　【記載待ち】文言や施策の内容が変わる事項",
     "記載そのものは済んでおり、数値は動きません。ご回答により文言または施策の内容が変わります。"),
 (4, "第4部　村の内部手続・日程の決定を待つ事項",
     "当方の作業は止まりません。決まり次第お知らせください。"),
 (None, "第5部　資料のご提供をお願いしたい事項",
     "既存の資料をそのままお送りいただければ結構です。作成をお願いするものではありません。"
     "お手元にないものは「なし」とご記入ください。"),
]

GOANNAI = [
 "本票は、第10期計画の策定にあたり村にご確認・ご提供をお願いしたい事項を、"
 "まとまりごとに1通にまとめたものです。",
 "ご回答は、各表の右端の「ご回答」の欄にご記入のうえ、本ファイルをご返送ください。"
 "別紙やメール本文でのご回答でも差し支えありません。",
 "すべてにご回答いただく必要はありません。第1部からお願いできれば、以降の工程は"
 "当方で進めることができます。",
 "数値は、出所（どの資料の何年度のものか）を併せてお示しいただけますと、"
 "成果品に出所として記載できます。",
 "個人が特定される情報は不要です。集計された数値のみでお願いいたします。",
]


# ══════════════════════════════════════════════════════════════
# 村にお見せする文に直す
# ══════════════════════════════════════════════════════════════
#   確認事項の「影響」「当方の手当」の欄は、当方の作業の覚え書きとして書いている。
#   そのまま照会票に載せると、点検の番号・スクリプトの名・他案件の名といった
#   村に関わりのないものが村の目に触れる。現に次のものが混入していた。
#     ・「verify_soan.py 点検46」「data_zuhyo.FUSHO」（当方の仕組みの名）
#     ・「他案件（大雪）の点検で指摘された」（他団体の案件の名）
#     ・「★10/5：当方では確かめられなかった。」（日付つきの内部の追記）
#   そこで、内部のものに触れている文を落とし、長さを抑えてからお見せする。
#   落とした中身は確認事項とWBSの備考に残るため、記録は失われない。

# 内部のものを指す言い方。語ではなく形で見る。
#   「点検」をそのまま落とすと、制度の用語である「ケアプラン点検」（給付適正化の
#   主要3事業）まで落ちる。「doc」をそのまま落とすと「docx」が落ちる。
#   現に、この2つと「受託者」で照会の見出し5件が空になり、村に空欄の質問が
#   届く状態になっていた。だから当方の記録を指す形だけを落とす。
NAIBU = (
 r"点検\d",                     # 当方の自己点検の番号（制度の「ケアプラン点検」ではない）
 r"doc\d",                      # 当方の記録の番号
 r"\.(py|js|json)\b",           # 当方の仕組みのファイル名
 r"data_zuhyo|FUSHO|verify_",   # 当方の仕組みの中の名
 r"スクリプト",
 r"当方では確かめられ",           # 日付つきの内部の追記に付く言い方
 r"他案件|大雪|金ヶ崎|浜田",       # 他団体の案件の名（成果品に持ち込まない）
 r"レッドチーム",
 r"弊社|当社",
)
# 「受託者」は落とさない。照会票は誰が何をするかを尋ねるものであり、
# 「受託者側で入力するか」「受託者に求められる関与の範囲」は照会そのものである。
# 委員会資料で同じ扱いをしているのと揃える。
_NAIBU_RE = re.compile("|".join(NAIBU))

TSUTAERU_MAX = 170        # この長さを超えたら文の単位で切る（途中で切らない）
_HIZUKE = re.compile(r"★\s*\d{1,2}/\d{1,2}")


def midashi(t):
    """照会の見出し（尋ねたいことそのもの）。

    ここでは文を落とさない。落とすと質問が消え、村に空欄が届く。
    印の★と、日付つきの内部の追記だけを取る。
    """
    t = str(t or "").strip().lstrip("★").strip()
    m = _HIZUKE.search(t)
    if m:
        t = t[:m.start()].strip()
    return t or "―"


def yasashiku(t, limit=TSUTAERU_MAX):
    """「当方の現在の扱い」を村にお見せする形に直す。

    内部の記録を指している文を落とし、長さを文の単位で抑える。
    落とした中身は確認事項とWBSの備考に残るため、記録は失われない。
    """
    t = str(t or "").strip()
    m = _HIZUKE.search(t)
    if m:
        t = t[:m.start()]                      # 日付つきの内部の追記から先を落とす
    keep, n = [], 0
    for sn in t.split("。"):
        sn = sn.strip().lstrip("★").strip()
        if not sn:
            continue
        if _NAIBU_RE.search(sn):
            continue
        if keep and n + len(sn) > limit:
            break
        keep.append(sn)
        n += len(sn)
    return ("。".join(keep) + "。") if keep else "―"

def _ws(k):
    return set(re.split(r"[,、／/ ]+", str(k[3]).strip())) - {""}


def assign(k):
    """関連WBSの重なりが最も多い束を返す。重なりがなければ None。"""
    best, bn = None, 0
    for nm, _, ws, _ in PD.BUNDLE:
        n = len(_ws(k) & set(ws))
        if n > bn:
            best, bn = nm, n
    return best


def bundles():
    """束の名 → (表題, ねらい) の並び。束7は他メンバーの担当のため外す。"""
    return [(b[0], b[1], b[3]) for b in PD.BUNDLE if b[0] != "束7"]


def items(bundle):
    """その束の確認事項を、部ごとに分けて返す。

    返り値： [(影響度, 部の表題, 前置き,
                [(見出し, 当方の扱い, 時期, 優先度, 状況, 元の見出し), …]), …]

    末尾の「元の見出し」は確認事項との突き合わせに使う鍵である。
    村にお見せする見出しではなく、確認事項に書かれているそのものを持つ。
    """
    mine = [k for k in KK.K if k[6] in MACHI and assign(k) == bundle]
    out = []
    for lv, title, lead in PARTS:
        rows = []
        for k in mine:
            imp = PD.IMPACT.get(k[1])
            klv = imp[0] if imp else None
            if klv == 0:
                continue                      # 社内。村に尋ねるものではない
            if klv != lv:
                continue
            rows.append((midashi(k[1]),
                         yasashiku(imp[2] if imp else k[2]),
                         yasashiku(imp[3] if imp else "―", 40).rstrip("。"),
                         k[5], k[6], k[1]))
        rows.sort(key=lambda r: (YUSEN.get(r[3], 9), r[0]))
        if rows:
            out.append((lv, title, lead, rows))
    return out


def count(bundle):
    return sum(len(r[3]) for r in items(bundle))


if __name__ == "__main__":
    for nm, ti, nerai in bundles():
        print("%s %s（%d件）" % (nm, ti, count(nm)))
        for lv, title, _, rows in items(nm):
            print("   %s：%d件" % (title.split("　")[1] if "　" in title else title, len(rows)))
