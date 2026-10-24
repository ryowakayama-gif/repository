# -*- coding: utf-8 -*-
"""他団体・他社の固有名称が成果品の置き場に残っていないかの点検
（令和8年10月19日）

  python3 07_ソーススクリプト/check_tadantai_R8.10.19.py

なぜ要るか
  CLAUDE.md 第3節　**他団体の固有名称を成果品に残さない
  （比較材料として読んだ団体を含む）。**

  ⚠ 既にある `check_seiyaku_R8.10.6.py` は、**あらかじめ決めた
  9件の文書**しか見ない。成果品の置き場にはその何倍もの文書があり、
  ファイル名に他団体名が付いたものが現に3件残っていた
  （令和8年10月19日に 10_内部資料_非成果品/ へ移した）。

  本点検は**置き場をまるごと走査する。**
    ①ファイル名　②本文（docx・xlsx・pptx・md・txt）

  終了コード　0＝残っていない／1＝1件でも残っている

⚠ **計画素案の旧版は別に数える。** 旧版（Ver.1.6〜Ver.2.15）の本文には
  前期の策定業者名・他の計画の受託者名が残っているが、**現行版
  （Ver.2.16以降）では除去済み**であり、旧版は「版の連なり」として
  残しているものである。移すと版の連なりが切れ、リポジトリ健全性の
  点検（検査4）が通らなくなる。ここでは件数と版だけを示す。
"""
import os
import re
import sys
import zipfile

# 走査する置き場（町へお出しするもの、又はその素）
BASHO = ["01_第10期_最新版成果品", "02_計画素案", "03_委員会・説明資料",
         "05_試算・管理シート", "06_参考成果品", "08_図表"]

# 他団体・他社の固有名称
#   ⚠ **比較材料として読んだ団体を含む。**
#   ⚠ 有病率の調査地（久山町・中島町・中山町・海士町）は、
#     国の研究の出典の一部であり、比較材料ではないため挙げない。
#   kubun="団体"＝不適合（終了コード1）／kubun="他社"＝要判断
#   ⚠ **他社（前期の策定業者・他の計画の受託者）の扱いは、
#     令和8年10月19日時点で若山主担当のご判断待ちである。**
#     「他団体名を含む資料を移す」とのご指示は自治体名を指しており、
#     他社名は含まれていなかった。移すより、
#     「地域福祉計画・障害者計画の受託事業者」のような書き方に改める方が
#     筋と考える。決まるまでは**要判断**として件数で示し、
#     「適合」に丸めない。
NG = [
    ("他団体（比較材料として読んだ団体）",
     r"(大雪地区広域連合|大雪広域|東川町|東神楽町|上川町|美瑛町|"
     r"当麻町|比布町|愛別町|幌加内町|鷹栖町)", "団体"),
    ("他団体（別の案件）", r"(金ケ崎町|金ヶ崎町)", "団体"),
    ("他社（他の計画の受託者・前期の策定業者）",
     r"(ジャパン" r"総研|福祉工房|東京コンサルティング)", "他社"),
    # ⚠ **「川崎市と取り違えない」という注意書き自体を不適合にしない。**
    #   成果品には取り違えを戒める注記を意図して置いており、
    #   それを不適合にすると注記を消すことになって本末転倒である。
    #   注意の文脈（取り違え・混同・別団体・ではない）の近くに
    #   現れるものは数えない。
    ("川崎市との混同", r"川崎市", "団体"),
]

# 計画素案の旧版（版の連なりとして残すもの）
KYUHAN = re.compile(r"川崎町_計画書素案_v(\d+)\.(\d+)")
GENKO_IJO = (2, 16)      # Ver.2.16 以降が現行の系列


def honbun(p):
    e = os.path.splitext(p)[1].lower()
    try:
        if e in (".md", ".txt", ".csv"):
            return open(p, encoding="utf-8", errors="replace").read()
        if e in (".docx", ".xlsx", ".pptx"):
            z = zipfile.ZipFile(p)
            out = []
            for n in z.namelist():
                if n.endswith(".xml") or n.endswith(".rels"):
                    out.append(z.read(n).decode("utf-8", "replace"))
            return "\n".join(out)
    except Exception:
        return ""
    return ""


CHUI = re.compile(r"(取り違え|取違え|混同|別団体|別の団体|ではありません|"
                  r"ではない|と区別|と間違|混じ|紛れ|誤認)")


def chui(t, pos, haba=60):
    """その出現が「取り違えないように」という注意の文脈かどうか。"""
    return bool(CHUI.search(t[max(0, pos - haba):pos + haba]))


def kyuhan_ka(name):
    """計画素案の旧版なら (True, 版) を返す。"""
    m = KYUHAN.search(name)
    if not m:
        return False, None
    v = (int(m.group(1)), int(m.group(2)))
    return v < GENKO_IJO, f"Ver.{m.group(1)}.{m.group(2)}"


def main():
    futeki_na = []      # ファイル名
    futeki_hon = []     # 本文
    kyuhan = []         # 旧版（別に数える）
    n_file = 0

    for d in BASHO:
        if not os.path.isdir(d):
            continue
        for root, _, fs in os.walk(d):
            for f in sorted(fs):
                p = os.path.join(root, f)
                n_file += 1
                # ① ファイル名
                for mi, pat, ku in NG:
                    if re.search(pat, f):
                        futeki_na.append((p, mi, ku))
                # ② 本文
                t = honbun(p)
                if not t:
                    continue
                atari = []
                for mi, pat, ku in NG:
                    go = []
                    for m in re.finditer(pat, t):
                        if mi == "川崎市との混同" and chui(t, m.start()):
                            continue
                        go.append(m.group(0))
                    go = sorted(set(go))
                    if go:
                        atari.append((mi, "／".join(go), ku))
                if not atari:
                    continue
                ka, ban = kyuhan_ka(f)
                if ka:
                    kyuhan.append((p, ban, atari))
                else:
                    futeki_hon.append((p, atari))

    print("他団体・他社の固有名称の点検（成果品の置き場）")
    print(f"  走査した置き場 {len(BASHO)}件／ファイル {n_file:,}件")
    print()

    print("  ── ① ファイル名")
    if futeki_na:
        for p, mi, ku in futeki_na:
            print(f"   × {p}　（{mi}）")
    else:
        print("   ○ 他団体・他社の固有名称を含むファイル名はありません")

    print("  ── ② 本文（計画素案の旧版を除く）")
    dantai = [(p, a) for p, a in futeki_hon
              if any(k == "団体" for _, _, k in a)]
    kaisha = [(p, a) for p, a in futeki_hon
              if all(k == "他社" for _, _, k in a)]
    if dantai:
        for p, atari in dantai:
            print(f"   × {p}")
            for mi, go, ku in atari:
                print(f"        {mi}：{go}")
    else:
        print("   ○ 本文に**他団体**の固有名称はありません")
    print("  ── ②-2 本文（他社の固有名称）　⚠ **要判断**")
    if kaisha:
        print(f"   △ {len(kaisha)}件。ご判断待ちのため「適合」に"
              "丸めていません。")
        for p, atari in kaisha:
            go = "／".join(g for _, g, _ in atari)
            print(f"      {p}　{go}")
        print("      移すよりも、「地域福祉計画・障害者計画の受託事業者」の"
              "ような書き方に改める方が筋と考えます。")
        print("      ⚠ 議事録は発言の記録ですので、扱いを分ける必要が"
              "あるかもしれません。")
    else:
        print("   ○ 本文に他社の固有名称はありません")

    print("  ── ③ 計画素案の旧版（版の連なりとして残すもの）")
    if kyuhan:
        bans = sorted(set(b for _, b, _ in kyuhan),
                      key=lambda x: [int(y) for y in x[4:].split(".")])
        print(f"   △ {len(kyuhan)}件（{bans[0]}〜{bans[-1]}）")
        go_all = sorted(set(g for _, _, a in kyuhan for _, gs, _ in a
                            for g in gs.split("／")))
        print(f"      現れる語：{'／'.join(go_all)}")
        print(f"      ⚠ **現行版（Ver.{GENKO_IJO[0]}.{GENKO_IJO[1]}以降）"
              f"では除去済み**です。旧版は版の連なりとして残しています。")
        print("      移すと版の連なりが切れ、リポジトリ健全性の点検"
              "（検査4）が通らなくなります。")
    else:
        print("   ○ 旧版にも残っていません")

    print()
    print("  ── 走査できなかった範囲（未実施）")
    print("   図（png）の中の文字は機械では拾えません。目で見てください。")
    print("   pdf は本点検の対象外です（同名の docx・xlsx で見ています）。")

    print()
    n = len(futeki_na) + len(dantai)
    if n:
        print(f"   × 不適合 {n}件"
              f"（ファイル名 {len(futeki_na)}／本文 {len(dantai)}）")
        print("      他団体が主題の資料は 10_内部資料_非成果品/ へ移します。")
        print("      言及が1か所だけなら、その箇所を他団体名を出さない"
              "書き方に改めます。")
        sys.exit(1)
    print(f"   ○ 他団体の固有名称は0件です。"
          f"（要判断 他社名 {len(kaisha)}件／"
          f"旧版 {len(kyuhan)}件は版の連なりとして別に数えています）")


if __name__ == "__main__":
    main()
