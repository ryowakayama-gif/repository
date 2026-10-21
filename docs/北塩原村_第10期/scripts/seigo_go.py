# -*- coding: utf-8 -*-
"""両計画の素案を決めた語で数え、片側のみの記載を洗い出す。

【なぜ必要か】
  北塩原村は第8期障がい福祉計画・第4期障がい児福祉計画を同時に策定している。
  両計画の接点にある事項は、**片方の計画にだけ書かれていても気づけない**。
  現に障がい側の照会Ｋ-7（令和8年10月8日）は「介護保険事業計画側に介護保険
  優先原則・適用除外・共生型サービス・新高額障害福祉サービス等給付費の記載が
  いずれも0件」という、語を数えて見つけたものだった。

  その照会は受け渡しの事故で当方に届かず、4日遅れて反映した。語を数えるのは
  **相手の照会を待たずに自分で見つける**ための手だてである。
  令和8年10月12日の申し送りで、相手と「語を数えて互いに確かめる」ことを
  約束した（seigo_han.YAKUSOKU）。相手は16語で数えている。

【数え方】
  当方：docs/北塩原村_第10期/18_計画素案.md（組んだ本文そのもの）
  相手：output/北塩原村_計画素案_正本_移植後.docx の本文（版は AITE_HAN）。
        相手は素案のマークダウンを持たず、正本は docx である。
        生成器（build_kitashiobara_honpon.py・gaiyou.py）を数えても
        正本に入っている語を拾えない。現に受給者証・新高額・介護保険優先は
        生成器に0件で正本にはある。**正本を数える**。

【使い方】
  python3 scripts/seigo_go.py          語ごとの件数と片側のみを表示する
  python3 scripts/seigo_go.py --katagawa   片側のみの語だけを挙げる
"""
import sys
sys.dont_write_bytecode = True
import io
import os
import re
import subprocess
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", ".."))
TOUHOU = os.path.join(ROOT, "docs", "北塩原村_第10期", "18_計画素案.md")

# 相手の正本（seigo_han.AITE の版で読む）。内容のハッシュは seigo_han が照合する
AITE_HAN = "9e960b6a"
AITE_SEIHON = "output/北塩原村_計画素案_正本_移植後.docx"

# ══════════════════════════════════════════════════════════════
# 数える語
#   go      数える語
#   naze    なぜこの語を数えるか（片側のみだと何が起きるか）
#   doko    本計画でこの語を置いた節。置かないと決めた語は ""
#   katate  片側のみでよい語は True（相手の制度に固有のもの）
# ══════════════════════════════════════════════════════════════
GO = [
 {"go": "介護保険優先", "doko": "1-4",
  "naze": "65歳到達時にどちらの制度が先かの案内が、両計画で割れる", "katate": False},
 {"go": "適用除外", "doko": "資9・5-1",
  "naze": "指定障害者支援施設の入所者を第1号被保険者に数えると保険料の分母が動く",
  "katate": False},
 {"go": "共生型", "doko": "施策3-1・資9・2-5",
  "naze": "同一の事業所が両計画に現れ、事業所数を足すと実態の倍になる",
  "katate": False},
 {"go": "受給者証", "doko": "1-4・資9",
  "naze": "被保険者資格と見込量を決めるのは受給者証の記載内容である", "katate": False},
 {"go": "新高額障害福祉サービス等給付費", "doko": "資9・施策4-1",
  "naze": "申請の窓口は介護保険側であり、案内がないと該当者が申請にたどり着けない",
  "katate": False},
 {"go": "障害支援区分", "doko": "資9",
  "naze": "障害支援区分と要介護度は別の尺度であり、読み替えられない", "katate": False},
 {"go": "特定疾病", "doko": "1-4・5-3",
  "naze": "40歳以上65歳未満で介護保険が優先する場合の範囲が決まる", "katate": False},
 {"go": "計画相談支援", "doko": "資9",
  "naze": "介護保険に移ると居宅介護支援に替わる。移行分の増減が両計画で食い違う",
  "katate": False},
 {"go": "相談支援専門員", "doko": "施策4-1",
  "naze": "介護支援専門員との引継ぎの相手であり、片側だけに書くと引継ぎが成り立たない",
  "katate": False},
 {"go": "成年後見制度利用支援事業", "doko": "施策4-7・資9",
  "naze": "同じ名前で別の事業である。注記がないと両計画の数が足し合わされる",
  "katate": False},
 {"go": "処遇改善加算", "doko": "施策4-6",
  "naze": "介護と障がいで事業所が重なるため、働きかけが片方からしか行われない",
  "katate": False},
 {"go": "障がい福祉計画", "doko": "1-4・資9",
  "naze": "相互参照そのもの。片側だけだと連携の節が一方通行になる", "katate": False},
 {"go": "就労選択支援", "doko": "",
  "naze": "訓練等給付であり介護保険に相当するものがない。本計画の対象外であるため"
          "当方は置かない。片側のみでよい語を1つ置き、点検が本当に働くかを示す",
  "katate": True},
 {"go": "地域生活支援事業", "doko": "施策4-7・資9",
  "naze": "成年後見制度利用支援事業の根拠。これがないと別の事業である説明が立たない",
  "katate": False},
 {"go": "保険料基準額", "doko": "5-7",
  "naze": "介護保険の保険料は障がい福祉計画には現れない。"
          "逆向きにも片側のみでよい語を1つ置く",
  "katate": True},
 {"go": "65歳", "doko": "1-4・施策4-1・資9",
  "naze": "両計画の見込量を同じ年度に動かす起点である", "katate": False},
]


def touhou_text():
    """当方の素案の本文"""
    with open(TOUHOU, encoding="utf-8") as f:
        return f.read()


def aite_text(han=AITE_HAN):
    """相手の正本（docx）の本文。取り出せなければ空"""
    try:
        b = subprocess.run(["git", "show", f"{han}:{AITE_SEIHON}"], cwd=ROOT,
                           capture_output=True, timeout=120).stdout
    except Exception:
        return ""
    if not b:
        return ""
    try:
        z = zipfile.ZipFile(io.BytesIO(b))
        x = z.read("word/document.xml").decode("utf-8")
    except Exception:
        return ""
    return "".join(re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", x, re.S))


def kazu():
    """語ごとに（当方の件数, 相手の件数）を返す"""
    t, a = touhou_text(), aite_text()
    if not a:
        return None
    return [(g, t.count(g["go"]), a.count(g["go"])) for g in GO]


def katagawa():
    """片側のみの語を挙げる。片側でよいと決めた語（katate）は除く"""
    k = kazu()
    if k is None:
        return None
    warui = []
    for g, tn, an in k:
        if g["katate"]:
            continue
        if (tn == 0) != (an == 0):
            warui.append((g, tn, an))
    return warui


def main():
    k = kazu()
    if k is None:
        print("相手のブランチが取れないため数えられない。"
              f"git fetch origin（{AITE_HAN} が必要）")
        return 2
    only = "--katagawa" in sys.argv
    if not only:
        print(f"■ 両計画を語で数える　{len(GO)}語　"
              f"（相手の版 {AITE_HAN}・正本の本文）\n")
        print(f"  {'語':<32}{'本計画':>8}{'障がい計画':>12}　状態")
        print("  " + "─" * 70)
        for g, tn, an in k:
            if tn and an:
                st = "両計画にあり"
            elif tn or an:
                st = "★片側のみ（本計画のみ）" if tn else "★片側のみ（障がい計画のみ）"
                if g["katate"]:
                    st = "片側のみ（そう決めた）"
            else:
                st = "★両計画になし"
            print(f"  {g['go']:<32}{tn:>8}{an:>12}　{st}")
        print()
    w = katagawa()
    print(f"■ 片側のみの語　{len(w)}件\n")
    for g, tn, an in w:
        doko = g["doko"] or "（本計画には置かない）"
        print(f"  {g['go']}　本計画 {tn}件／障がい計画 {an}件")
        print(f"    なぜ数えるか {g['naze']}")
        print(f"    本計画での置き場所 {doko}")
        print()
    if not w:
        print("  片側のみの語はない\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
