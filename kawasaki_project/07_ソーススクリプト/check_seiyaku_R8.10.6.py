# -*- coding: utf-8 -*-
"""守るべき制約の走査（令和8年10月6日）

点検スキル `.claude/skills/plan-draft-check/` の3「守るべき制約の走査」を
機械で回せる形にする。**これまで走査のスクリプトがなく、目で見るほかなかった。**

Ver.2.10 に当てたところ、受託者を主語とする語が2か所残っていた
（1-5の注記と10-1の工程）。Ver.2.11 で是正している。

  走査するもの　本文・入れ子を含む表・ヘッダー・フッター・テキストボックス
  走査できないもの　画像の中の文字（件数で示す）

  python3 07_ソーススクリプト/check_seiyaku_R8.10.6.py [docx のパス ...]

引数を省くと、町へお送りする成果品（計画素案・委員会資料・想定問答集・
概要版）をまとめて見る。

判定は3つ。
  不適合　直さなければ送付できない
  要判断　規則には当たるが、文脈により適否が分かれる
  未実施　走査できなかった範囲

⚠ **協議用の版では「確認事項No.」の参照と［要確認］等を意図して残している。**
  公表版では取り除く（確認事項No.162）。本点検は、どちらの版として見るかを
  `--公表版` で切り替える。
"""
import re
import sys
import zipfile

import docx
from docx.oxml.ns import qn

SOAN = "01_第10期_最新版成果品/川崎町_計画書素案_v2.12_根拠整理版.docx"
# （パス, 区分）
#   「計画」＝川崎町が定めるもの。受託者を主語とする語を置かない。
#   「当方資料」＝当方が作成して町へお出しするもの。
#     訂正の責任の所在を示すために当方を主語とすることがあり、
#     受託者を主語とする語は**要判断**とする（委員に配る場合は外す）。
KITEI = [
    (SOAN, "計画"),
    ("01_第10期_最新版成果品/川崎町_計画素案_概要版_R8.10.5.docx", "計画"),
    ("03_委員会・説明資料/川崎町_第2回策定委員会資料_R8.11_v8.docx",
     "当方資料"),
    ("03_委員会・説明資料/川崎町_第2回策定委員会_想定問答集_R8.11_v4.docx",
     "当方資料"),
]

# ══════════════════════════ 不適合とするもの
# （名称, 正規表現, 例外の正規表現）
FUTEKI = [
    ("受託者を主語とする語",
     r"(受託者|当社|弊社|ビズアップ|スクリプト|再実行|固定値|判定しています)",
     # 奥付の策定支援の表示は計画書の慣行による（確認事項No.167）
     r"策定支援："),
    ("他団体の固有名称",
     r"(大雪地区広域連合|大雪広域|東川町|東神楽町|上川町|美瑛町|"
     r"金ヶ崎町|金ケ崎町|川崎市)", None),
    ("強調記号の残り", r"(\*\*|__)", None),
    ("未確定箇所を［　］以外の括弧で書いたもの",
     r"(【[^】]{0,12}(確認|協議|要|未定|最新値)[^】]{0,12}】|"
     r"〔[^〕]{0,20}〕|\[[^\]]{0,20}\])",
     # 【町確認】【委員会協議】は協議用の版で用いる定めの表記。
     # 「協議会」「委員会」などの組織の名は未確定箇所ではない。
     r"(【(町確認|委員会協議)】|【[^】]*(協議会|委員会|部会|課|所|局)】)"),
]

# ══════════════════════════ 要判断とするもの
YOHANDAN = [
    ("禁止表現",
     r"(に由来する|と整合する|1件も|有意差がないため|全国トップ級)", None),
    ("電話番号の形", r"0\d{1,4}-\d{2,4}-\d{3,4}", None),
    ("メールアドレスの形", r"[\w.+-]+@[\w-]+\.[\w.-]+", None),
]

# ══════════════════════════ 公表版でのみ不適合とするもの
KOHYO_NOMI = [
    ("確認事項の参照", r"確認事項No", None),
    ("未確定箇所の表記", r"(【町確認】|【委員会協議】|［要確認］|［要協議］|"
                        r"［要内訳］)", None),
]


def yomu(path):
    """本文・表・ヘッダー・フッター・テキストボックスの文字を集める。

    ⚠ python-docx の `paragraphs` は**テキストボックスの中を拾わない。**
      XML を直に辿って `w:t` をすべて集める。
    """
    doc = docx.Document(path)
    honbun = "".join(n.text or "" for n in doc.element.body.iter(qn("w:t")))
    z = zipfile.ZipFile(path)
    hf = ""
    for n in z.namelist():
        if re.match(r"word/(header|footer)\d*\.xml", n):
            hf += "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>",
                                     z.read(n).decode("utf-8")))
    # 走査できない範囲（画像）の件数
    n_img = len(re.findall(
        r"<w:drawing>", z.read("word/document.xml").decode("utf-8")))
    return honbun + "\n" + hf, n_img


def mireru(text, pat, nozoku):
    """当たった箇所を返す。例外に当たるものは除く。"""
    out = []
    for m in re.finditer(pat, text):
        s = m.group(0)
        mae = text[max(0, m.start() - 24):m.start()]
        ato = text[m.end():m.end() + 24]
        if nozoku and re.search(nozoku, mae + s + ato):
            continue
        out.append((s, (mae + "《" + s + "》" + ato).replace("\n", " ")))
    return out


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    kohyo = "--公表版" in sys.argv
    paths = [(a, "計画") for a in args] or KITEI

    print("守るべき制約の走査")
    print(f"  見る版　{'公表版' if kohyo else '協議用の版'}"
          f"（協議用の版では確認事項の参照と未確定箇所の表記を残します）")
    ng = yo = 0
    n_img_all = 0
    for path, kubun in paths:
        try:
            text, n_img = yomu(path)
        except FileNotFoundError:
            print(f"   × ファイルがない：{path}")
            ng += 1
            continue
        n_img_all += n_img
        name = path.split("/")[-1]
        hits_ng, hits_yo = [], []
        rules = list(FUTEKI) + (KOHYO_NOMI if kohyo else [])
        yo_rules = list(YOHANDAN)
        if kubun == "当方資料":
            # 当方が作成する資料では、訂正の責任の所在を示すために
            # 当方を主語とすることがある。不適合ではなく要判断とする。
            rules = [x for x in rules if x[0] != "受託者を主語とする語"]
            yo_rules = yo_rules + [x for x in FUTEKI
                                   if x[0] == "受託者を主語とする語"]
        for nm, pat, nozoku in rules:
            for s, around in mireru(text, pat, nozoku):
                hits_ng.append((nm, s, around))
        for nm, pat, nozoku in yo_rules:
            for s, around in mireru(text, pat, nozoku):
                hits_yo.append((nm, s, around))
        mark = "○" if not hits_ng else "×"
        print(f"   {mark} {name}［{kubun}］"
              f"（不適合{len(hits_ng)}／要判断{len(hits_yo)}／図{n_img}点）")
        for nm, s, around in hits_ng[:12]:
            print(f"      ［不適合］{nm}：{around[:84]}")
        for nm, s, around in hits_yo[:8]:
            print(f"      ［要判断］{nm}：{around[:84]}")
        ng += len(hits_ng)
        yo += len(hits_yo)

    print(f"  ── 走査できなかった範囲（未実施）")
    print(f"   図の中の文字 {n_img_all}点。"
          "団体名・個人情報・禁止表現は目で見てください。")
    print("   氏名・住所・自由記述は値の形で探せないため走査していません。"
          "委員名簿・事業所一覧・意見聴取の記録を加えるときは目で見てください。")
    print(f"  不適合 {ng}件／要判断 {yo}件")
    if ng:
        sys.exit(1)
    print("   ○ 受託者を主語とする語・他団体の固有名称・強調記号・"
          "括弧の誤りはありません")


if __name__ == "__main__":
    main()
