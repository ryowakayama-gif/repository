"""計画書（成果品⑤「A4・4色刷り・80頁程度」）に向けた頁の配分計画。

素案のPDFから章ごとの頁数を数え、素案にしかない頁（協議用の注記・変更履歴・凡例・確定に必要な事項）を除いたうえで、
本編と資料編の配分と、80頁程度に収める整理案を示す。

出力
  10_納品・印刷/小野町_高齢者計画_計画書80頁への頁配分計画_YYYYMMDD.docx
"""

import pathlib
import re
import subprocess
import sys

import ono_style as S

HERE = pathlib.Path(__file__).parent
ROOT = HERE / "小野町_引継ぎ_整理済"
PDF = ROOT / "03_計画素案" / "PDF" / "小野町高齢者保健福祉計画_第10期介護保険事業計画_素案_第16版_20261022.pdf"
OUT = ROOT / "10_納品・印刷"
ASOF = "20261022"
ASOF_JP = "令和8年10月22日"
MOKUHYO = 80

# 章の見出し（素案の chapter() の題）。頁の先頭行からこれらを探して章の区切りとする
SHO = ["本素案について", "目次", "第1章　計画策定にあたって", "第2章　小野町の高齢者・介護等の状況",
       "第3章　第9期計画の実施状況と評価", "第4章　計画の基本的考え方", "第5章　施策の展開",
       "第6章　介護保険事業等の見込み", "資料編", "参考　本素案の確定に必要な事項"]
# 計画書での扱い：(扱い, 目安頁, 整理の方法)
ATSUKAI = {
    "本素案について": ("除く", 0, "協議用の注記・変更履歴・未確定の凡例。計画書には載せない"),
    "目次": ("残す", 2, "計画書の目次（2頁）"),
    "第1章　計画策定にあたって": ("残す", 9, "法改正の表は要点に絞る。関連計画は表1枚にまとめる"),
    "第2章　小野町の高齢者・介護等の状況": ("残す", 13, "図を優先し、同じ数値の表と図の重複を除く"),
    "第3章　第9期計画の実施状況と評価": ("残す", 5, "評価の表を中心に、説明文を短くする"),
    "第4章　計画の基本的考え方": ("残す", 4, "体系図を1頁にまとめる"),
    "第5章　施策の展開": ("残す", 22, "各施策を「現状と課題・施策の方向・主な事業・成果指標」の4項目、1施策1〜2頁にそろえる"),
    "第6章　介護保険事業等の見込み": ("残す", 13, "算定の注記を資料編に移し、見込量と保険料の表を残す"),
    "資料編": ("絞る", 12, "用語解説・策定経過・委員名簿・成果指標の定義を残し、図表番号一覧・地域分析の指標・事業所一覧はデータ（CD-R）に収める"),
    "参考　本素案の確定に必要な事項": ("除く", 0, "協議用。計画書には載せない"),
}


def pages():
    out = subprocess.run(["pdftotext", "-layout", str(PDF), "-"], capture_output=True, text=True).stdout
    return out.split("\f")


def main():
    ng = []
    pg = pages()
    n_all = len([p for p in pg if p.strip()])
    start = {}
    for i, p in enumerate(pg):
        # 柱（1行目）の次の行が章の題そのものである頁を章の先頭とする（目次の行は拾わない）
        ls = [x for x in p.strip().splitlines() if x.strip()]
        head = re.sub(r"[\s　]", "", ls[1]) if len(ls) > 1 else ""
        for s in SHO:
            key = re.sub(r"[\s　]", "", s)
            if s not in start and head == key:
                start[s] = i
    found = [s for s in SHO if s in start]
    if len(found) < 8:
        ng.append("章の区切りを%d個しか拾えない：%s" % (len(found), found))
    order = sorted(found, key=lambda s: start[s])
    cnt = {}
    for k, s in enumerate(order):
        e = start[order[k + 1]] if k + 1 < len(order) else len(pg)
        cnt[s] = len([p for p in pg[start[s]:e] if p.strip()])
    gen = sum(cnt.values())
    mok = sum(ATSUKAI[s][1] for s in order)
    hyoshi = 4   # 表紙・あいさつ・裏表紙等

    r = S.Report("小野町　高齢者計画　計画書の頁配分")
    r.cover(["小野町高齢者保健福祉計画・第10期介護保険事業計画"],
            ["計画書（80頁程度）への頁の配分計画", ASOF_JP, "ビズアップ公共コンサルティング株式会社"])
    r.toc_here()
    r.body_here()
    r.h1("1　結論")
    r.p(f"**素案（第16版）は{n_all}頁です。** 仕様書の計画書は「A4版、4色刷り、80頁程度」のため、"
        f"協議用の頁を除き、資料編の一部をデータ（CD-R）に移すと、本編・資料編で{mok}頁、"
        f"表紙・町長あいさつ等を加えて{mok + hyoshi}頁になります。")
    r.h1("2　章ごとの頁数と計画書での扱い")
    r.tbl([["章", "素案の頁数", "計画書での扱い", "目安の頁数", "整理の方法"]]
          + [[s, "%d" % cnt[s], ATSUKAI[s][0], "%d" % ATSUKAI[s][1], ATSUKAI[s][2]] for s in order]
          + [["合計", "%d" % gen, "", "%d" % mok, f"表紙・あいさつ等{hyoshi}頁を加えて{mok + hyoshi}頁"]],
          widths=[4.6, 1.8, 2.0, 1.8, 7.0], right=(1, 3), caption="章ごとの頁数と計画書での扱い")
    r.src("素案のPDF（LibreOfficeで変換。游ゴシックがない環境のため、町のWordとは頁がずれうる）から章の区切りを読み取ったもの")
    r.h1("3　進め方")
    r.tbl([["時期", "作業"],
           ["パブリックコメントの前", "協議用の注記を外した「公表用の素案」を作り、この配分で頁を確かめる"],
           ["計画書の原稿確定時", "章ごとの目安頁に収まっているかを確かめ、超える章は図表の統合で調整する"],
           ["組版", "4色刷りのため、図の色は計画書の配色にそろえる。資料編のうちデータに移すものをCD-Rに収める"]],
          widths=[4.4, 12.8], caption="進め方")
    p = OUT / ("小野町_高齢者計画_計画書80頁への頁配分計画_%s.docx" % ASOF)
    r.build_toc()
    r.save(p)
    print("■ 頁配分計画")
    print("   ", p.relative_to(HERE))
    print("   素案%d頁（章%d）→ 目安%d頁＋表紙等%d" % (n_all, len(order), mok, hyoshi), {s[:6]: cnt[s] for s in order})
    if abs(mok + hyoshi - MOKUHYO) > 6:
        ng.append("目安の合計%d頁が80頁程度から離れている" % (mok + hyoshi))
    print("\n■ 自己点検")
    for x in ng:
        print("   NG", x)
    print("   すべて適合" if not ng else "   **%d件**" % len(ng))
    return 1 if ng else 0


if __name__ == "__main__":
    sys.exit(main())
