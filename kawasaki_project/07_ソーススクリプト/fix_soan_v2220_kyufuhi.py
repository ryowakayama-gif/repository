# -*- coding: utf-8 -*-
"""計画素案 Ver.2.21 → Ver.2.22（給付費の二重の値の是正ほか）
（令和8年10月21日）

  python3 07_ソーススクリプト/fix_soan_v2220_kyufuhi.py

なぜ要るか
  第2回策定委員会資料 v9 を起こすにあたって素案と突き合わせたところ、
  **同じ指標に2つの値**が置かれていた。

  ⚠ 7-2（p470）　「標準給付費は令和6年度の1,093,253千円から
    令和7年度の**1,085,203千円**へ**0.7％**減少しています」
  ⚠ 9-3（p597）　「保険給付費は令和6年度1,093,253千円から
    令和7年度**1,084,587千円**へ**0.8％**減少しており」

  原典（素案の表）で確かめた。
    表15  保険給付費 計　　　　1,084,587,307円（令和7年度の実績・円単位）
    表109 標準給付費見込額　　 令和7年度 1,084,587千円
  **令和7年度の標準給付費見込額＝保険給付費計＝1,084,587千円**であり、
  7-2 の 1,085,203千円 は Ver.1.11 の頃の値が残ったものである
  （Ver.1.14・1.17 で 1,084,587千円 に改めたが 7-2 だけ直し漏れていた）。
  1,093,253 → 1,084,587 は ▲0.79％であるから、率も 0.8％ が正しい。

あわせて直すもの
  ⚠ 巻末（p685）「事業所への照会の結果…を経て更新し」
    → 照会は行わない整理に改めた（Ver.2.21）。記述を合わせる。
  ⚠ 表34 行4 列3「B（事業者連絡会照会含む）」も同じ。

終了コード　0＝検査1〜6すべて適合／1＝1件でも不適合
"""
import os
import shutil
import sys

import docx

MOTO = "01_第10期_最新版成果品/川崎町_計画書素案_v2.21_公表情報把握版.docx"
SAKI = "01_第10期_最新版成果品/川崎町_計画書素案_v2.22_給付費整合版.docx"
BAN_MOTO = "Ver.2.21"
BAN = "Ver.2.22"

# (段落の番号, 旧の冒頭, 新しい本文)
DAN = [
    (470, "標準給付費は令和6年度の1,093,253千円から",
     "標準給付費は令和6年度の1,093,253千円から令和7年度の1,084,587千円へ"
     "0.8％減少しています。これが介護予防の成果によるものか、介護人材の"
     "不足によるサービス提供体制の制約によるものかの判別が、第10期の"
     "サービス見込量の前提となります。第1回策定委員会で報告された"
     "1ユニット分の人員削減に加え、町内の訪問介護の従業者が常勤換算"
     "4.5人にとどまること（介護サービス情報公表システム・令和8年1月記入）"
     "を踏まえると、供給の制約による可能性が高いと考えられます。"),
    # ⚠ 第3回の開催時期　第2回が令和8年11月24日に確定したことにより、
    #   パブリックコメント30日間を挟むと**第3回は令和9年1月では開けない。**
    #   回数と時期は第2回でお諮りする事項であるから、月を断定しない形に改める。
    (50, "被保険者代表・介護に関し学識又は経験を有する者・",
     "被保険者代表・介護に関し学識又は経験を有する者・介護サービスに"
     "関する事業に従事する者から構成される策定委員会において、"
     "令和8年9月2日（第1回）・令和8年11月24日（第2回）をはじめとする"
     "複数回の審議を経て策定します。第3回以降の開催時期及び回数は、"
     "策定委員会においてお諮りのうえ定めます［要協議］。"
     "なお、条例・要綱上の設置根拠及び正式名称は確認中です"
     "（確認事項No.30）。"),
    (645, "⑤　令和9年1月：第3回策定委員会で保険料基準額を決定し、",
     "⑤　第3回策定委員会で保険料基準額を決定し、答申の内容を確定"
     "（開催時期は策定委員会においてお諮りのうえ定めます）"),
    (648, "⚠ 具体的な保険料試算結果は、国民健康保険団体連合会の",
     "⚠ 具体的な保険料試算結果は、国民健康保険団体連合会の給付実績、"
     "所得段階別被保険者数、調整交付金の交付実績、及び令和9年度の"
     "介護報酬改定の内容を踏まえ、第3回策定委員会で確定します。"
     "確定値の信頼性確保のため、地域包括ケア「見える化」システムへの"
     "登録と、宮城県・近隣町との比較を経て協議に付します。"),
    (685, "本素案は、第9期の事業実績、関連施策評価調査の結果、",
     "本素案は、第9期の事業実績、関連施策評価調査の結果、介護サービス"
     "情報公表システムによる町内事業所の把握の結果、及び国通知の確認を"
     "経て更新し、令和8年11月24日の第2回策定委員会で計画骨子（素案）と"
     "して審議いただく予定です。第9章に掲げた保険料の試算値は、"
     "第3回策定委員会で確定値に置き換えます。"),
]

# (表の番号, 見出し, 行番号, 列番号, 旧の冒頭, 新)
SERU = [
    (34, "本節の項目", 4, 3, "B（事業者連絡会照会含む）",
     "B（介護サービス情報公表システムによる把握を含む）"),
]


# ⚠ 全体にかける語の置き換え（第3回の開催時期を断定しない）
#   第2回が令和8年11月24日に確定し、パブリックコメント30日間を挟むため、
#   **第3回は令和9年1月では開けない。** 回数と時期は第2回でお諮りする
#   事項であるから、月を断定している箇所を一括して改める。
OKIKAE = [
    ("第3回策定委員会（令和9年1月）", "第3回策定委員会"),
    ("第3回策定委員会（令和9年1月頃）", "第3回策定委員会"),
]


def okikae_all(d):
    """本文・表・ヘッダー・フッターの run に対して語を置き換える。"""
    n = 0

    def hashiru(ps):
        nonlocal n
        for p in ps:
            for r in p.runs:
                for kyu, atarashii in OKIKAE:
                    if kyu in r.text:
                        r.text = r.text.replace(kyu, atarashii)
                        n += 1

    hashiru(d.paragraphs)
    for t in d.tables:
        for row in t.rows:
            for c in row.cells:
                hashiru(c.paragraphs)
    for sec in d.sections:
        for part in (sec.header, sec.footer):
            hashiru(part.paragraphs)
    return n


def dan_kaku(p, s):
    if p.runs:
        p.runs[0].text = s
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(s)


def kaku(cell, s):
    ps = cell.paragraphs
    gyo = s.split("\n")
    for i, p in enumerate(ps):
        if i < len(gyo):
            dan_kaku(p, gyo[i])
        else:
            for r in p.runs:
                r.text = ""


def main():
    futeki = []

    def judge(ok, mi, setsu=""):
        print(("   ○ " if ok else "   × ") + mi + (("　" + setsu)
                                                   if setsu else ""))
        if not ok:
            futeki.append(mi)

    print(f"計画素案 {BAN_MOTO} → {BAN}（給付費の二重の値の是正）")
    print()
    if not os.path.exists(MOTO):
        print(f"   × 前の版がありません：{MOTO}")
        sys.exit(1)
    shutil.copy2(MOTO, SAKI)
    d = docx.Document(SAKI)
    mae = "\n".join([p.text for p in d.paragraphs]
                    + [c.text for t in d.tables for r in t.rows
                       for c in r.cells])

    print("  ── ① 段落の差し替え")
    n_dan = 0
    for i, atama, atarashii in DAN:
        p = d.paragraphs[i]
        if not p.text.startswith(atama[:18]):
            judge(False, f"p{i} の冒頭が合わない", f"実際『{p.text[:26]}』")
            continue
        dan_kaku(p, atarashii)
        n_dan += 1
        print(f"     p{i}　{atama[:28]}… → 差し替え")
    judge(n_dan == len(DAN), f"{len(DAN)}段落を差し替えた（実際 {n_dan}）")

    print("  ── ② 表のセルの差し替え")
    n_seru = 0
    for ban, mi, gyo, retsu, kyu, atarashii in SERU:
        t = d.tables[ban]
        if t.rows[0].cells[0].text.strip() != mi:
            judge(False, f"表{ban} の見出しが『{mi}』でない",
                  f"実際『{t.rows[0].cells[0].text[:16]}』")
            continue
        c = t.rows[gyo].cells[retsu]
        if not c.text.strip().startswith(kyu[:12]):
            judge(False, f"表{ban} 行{gyo} 列{retsu} の旧が合わない",
                  f"実際『{c.text[:40]}』")
            continue
        kaku(c, atarashii)
        n_seru += 1
    judge(n_seru == len(SERU), f"{len(SERU)}セルを差し替えた（実際 {n_seru}）")

    print("  ── ③ 第3回の開催時期を断定しない形に")
    n_oki = okikae_all(d)
    judge(n_oki >= 2, f"「第3回策定委員会（令和9年1月）」を改めた（{n_oki}か所）")

    print("  ── ④ 版表記")
    n_ban = 0
    for sec in d.sections:
        for part in (sec.header, sec.footer):
            for p in part.paragraphs:
                for r in p.runs:
                    if BAN_MOTO in r.text:
                        r.text = r.text.replace(BAN_MOTO, BAN)
                        n_ban += 1
    for p in d.paragraphs:
        for r in p.runs:
            if BAN_MOTO in r.text:
                r.text = r.text.replace(BAN_MOTO, BAN)
                n_ban += 1
    judge(n_ban >= 1, f"版表記を {BAN} に改めた（{n_ban}か所）")

    d.save(SAKI)

    d2 = docx.Document(SAKI)
    ato = "\n".join([p.text for p in d2.paragraphs]
                    + [c.text for t in d2.tables for r in t.rows
                       for c in r.cells])

    print("  ── 検査1　⚠ 同じ指標に2つの値が残っていないか")
    judge("1,085,203" not in ato, "「1,085,203」が0件（古い値）")
    judge(ato.count("1,084,587") >= 3,
          f"「1,084,587」が3件以上（実際 {ato.count('1,084,587')}件）")

    print("  ── 検査2　率が揃っているか")
    judge("1,084,587千円へ0.8％減少" in ato,
          "7-2 が「1,084,587千円へ0.8％減少」になっている")
    judge("0.7％減少" not in ato, "「0.7％減少」が0件")

    print("  ── 検査3　⚠ 事業所への照会の記述が残っていないか")
    judge("事業所への照会" not in ato, "「事業所への照会」が0件")
    judge("事業者連絡会照会" not in ato, "「事業者連絡会照会」が0件")
    judge("事業所ヒアリング" not in ato, "「事業所ヒアリング」が0件")

    print("  ── 検査4　第2回の開催日と第3回の時期")
    judge("令和8年11月24日の第2回策定委員会" in ato,
          "巻末に第2回の開催日（令和8年11月24日）がある")
    judge("令和8年11月24日（第2回）" in ato,
          "1-4 の策定の経過に第2回の開催日がある")
    # ⚠ 第3回を令和9年1月と断定していないか
    judge("第3回策定委員会（令和9年1月）" not in ato,
          "⚠ 「第3回策定委員会（令和9年1月）」が0件")
    judge("令和9年1月：第3回" not in ato, "⚠ 「令和9年1月：第3回」が0件")

    print("  ── 検査5　⚠ 既にあった数が消えていないか")
    import re
    # 意図した是正で消えるもの（誤った額と、それに対応する誤った率）
    IZU = {"1,085,203千円", "0.7％", "4回"}
    kesu = []
    for m in re.finditer(r"[0-9０-９][0-9０-９,.]*(?:人|％|千円|件|名|歳|回)",
                         mae):
        go = m.group(0)
        if mae.count(go) > ato.count(go):
            kesu.append(go)
    kesu = sorted(set(kesu) - IZU)
    judge(not kesu,
          "消えた数は誤った額（1,085,203千円）と率（0.7％）だけ",
          "／".join(kesu[:5]))

    print("  ── 検査6　受託者を主語とする語が入っていないか")
    NG = ("受託者", "当社", "ビズアップ", "スクリプト")
    aru = [w for w in NG if w in ato]
    judge(not aru, "受託者を主語とする語が0件", "／".join(aru))

    print()
    print(f"  {SAKI}")
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        for m in futeki:
            print(f"      {m}")
        sys.exit(1)
    print(f"   ○ 検査1〜6　不適合0件"
          f"（［要確認］{ato.count('［要確認］')}件）")


if __name__ == "__main__":
    main()
