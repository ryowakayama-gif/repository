# -*- coding: utf-8 -*-
"""第2回策定委員会資料 v9 → v10（宮城県の入所状況調の反映）
（令和8年10月22日）

  python3 07_ソーススクリプト/build_2kai_v10.py

v9 からの変更の理由
  計画素案を Ver.2.23 に上げた（fix_soan_v2230_kenchosa.py）。
  素案を直したら、同じ内容を載せている成果品も同じ版で直す（CLAUDE.md）。
  (1) 表6「公表の項目にないものは５件」→ 稼働率・休止床は県の入所状況調で
      推移が分かったため、残るものの整理を改める
  (2) 3か月ごとの入所率・入所希望者・在宅への退所の表を加える（表8 の後）
  (3) ⚠ 町内の軽費老人ホーム（定員70人）の休止・廃止を表6 に書く
  (4) 頁数の表と本文の素案の版を Ver.2.23 に改める
  (5) ⚠ v9 の表6 にあった「2025年09月01日から令和８年10月まで」
      （元号と西暦の混在）を正す。v9 を起こしたときの日付の変換の誤り

⚠ 県内他団体の前期計画は確かめていない（ご指示により方針Ｃ）。

数値の正本　09_元資料/その他データ/県入所状況調_川崎町_R8.10.22.csv
"""
import csv
import importlib.util
import os
import re
import shutil
import sys

import docx

HERE = os.path.dirname(os.path.abspath(__file__))


def load(name, fn):
    spec = importlib.util.spec_from_file_location(name, os.path.join(HERE, fn))
    m = importlib.util.module_from_spec(spec)
    sys.modules[name] = m
    spec.loader.exec_module(m)
    return m


V9 = load("build_2kai_v9", "build_2kai_v9.py")          # 表の操作の道具
S23 = load("fix_soan_v2230", "fix_soan_v2230_kenchosa.py")  # 数値

SRC = "03_委員会・説明資料/川崎町_第2回策定委員会資料_R8.11_v9.docx"
DST = "03_委員会・説明資料/川崎町_第2回策定委員会資料_R8.11_v10.docx"
CSV_IN = S23.CSV_IN
JUN = S23.JUN
A = S23.A


def kibou(shu, h):
    with open(CSV_IN, encoding="utf-8-sig", newline="") as f:
        return sum(S23.num(r["入所希望者"]) or 0 for r in csv.DictReader(f)
                   if r["種別"] == shu and r["時点"] == h)


def zen(s):
    """全角数字に（委員会資料は数字を全角で書いている箇所がある）。"""
    return s


HYOSOKU = [
    ["サービスの種類", f"入所率（{A['時点数']}時点の幅）",
     f"入所希望者（{S23.wa(JUN[0])}→{S23.wa(JUN[-1])}）",
     "在宅への退所（14時点の計）"],
    ["介護老人福祉施設（２施設）", S23.haba(A["特養"]),
     f"{kibou('特養', JUN[0]):.0f}人 → {kibou('特養', JUN[-1]):.0f}人",
     f"{A['在宅_特養']:.0f}人"],
    ["介護老人保健施設（１施設）", S23.haba(A["老健"]),
     f"{kibou('老健', JUN[0]):.0f}人 → {kibou('老健', JUN[-1]):.0f}人",
     f"{A['在宅_老健']:.0f}人"],
    ["認知症対応型共同生活介護（２事業所）", S23.haba(A["ＧＨ"]),
     f"{kibou('ＧＨ', JUN[0]):.0f}人 → {kibou('ＧＨ', JUN[-1]):.0f}人",
     f"{A['在宅_ＧＨ']:.0f}人"],
]

HYO6 = (
    "⚠ 事業所への照会は行いません。 照会でお尋ねする予定であった項目の"
    "うち、定員・在所者数・入所率・待機者数・職種別の従業者数・前年度の"
    "採用者数及び退職者数・指定の年月日・退所先は、厚生労働省「介護"
    "サービス情報公表システム」に公表されており、町内14の指定すべてに"
    "ついて把握できました（令和８年10月）。\n"
    "⚠ 入所率と退所理由の推移は、宮城県「老人福祉施設等入所状況調」"
    f"（3か月ごと・{S23.wa(JUN[0])}から{S23.wa(JUN[-1])}までの"
    f"{A['時点数']}時点）で把握しました（下の３つ目の表）。"
    "これは時点の在所者数によるもので、延べ利用日数による稼働率では"
    "ありません。\n"
    "⚠ いずれの資料にもないものは４件です。 入所前の居所・経営の状況・"
    "町への要望・短期入所の稼働。入所前の居所は、宮城県指定介護老人福祉"
    "施設入所指針が定める保険者の協力の場で集計値を求めることを検討し、"
    "経営の状況は計画に書かず、町への要望は地域ケア会議等の既存の場で"
    "聴きます。\n"
    f"⚠ 町内にあった軽費老人ホーム（定員{A['軽費定員']:.0f}人）は、"
    f"{S23.wa(S23.KYUSHI)}時点で休止中となり、{S23.wa(S23.NAI)}以降は"
    "県の入所状況調に掲載されていません。 老人福祉法の施設であり、"
    "介護サービス情報公表システムには載りません。経緯と入居されていた方の"
    "行き先を確認しています。\n"
    "⚠ 公表情報は事業所の年１回の報告によります。 記入日は事業所ごとに"
    "異なり、令和７年９月から令和８年10月までの幅があります。")

DAN = [
    (4, "計画骨子（案）とご選択いただきたい事項",
     "計画骨子（案）とご選択いただきたい事項　［第10版］"),
    (90, "計画書の事業一覧表（17表・138事業）に",
     None),     # 版だけ置き換える
]

SERU = [
    (6, "⚠ 事業所への照会は行いません。", 0, 0, "⚠ 事業所への照会は行いません。",
     HYO6),
    (37, "版", 2, 0, "Ver.2.22（現在・",
     "Ver.2.23（現在・認知症の人の数の推計、公表情報による供給の把握、"
     "県の入所状況調の推移を追加）"),
]


def main():
    futeki = []

    def judge(ok, mi, setsu=""):
        print(("   ○ " if ok else "   × ") + mi + (("　" + setsu)
                                                   if setsu else ""))
        if not ok:
            futeki.append(mi)

    print("第2回策定委員会資料 v9 → v10（宮城県の入所状況調の反映）")
    print()
    shutil.copy2(SRC, DST)
    d = docx.Document(DST)
    mae = "\n".join([p.text for p in d.paragraphs]
                    + [c.text for t in d.tables for r in t.rows
                       for c in r.cells])
    KYU = []

    print("  ── ① 段落")
    n = 0
    for i, atama, atarashii in DAN:
        p = d.paragraphs[i]
        if not p.text.strip().startswith(atama):
            judge(False, f"p{i} の冒頭が合わない", f"実際『{p.text[:24]}』")
            continue
        KYU.append(p.text)
        s = (p.text.replace("Ver.2.22", "Ver.2.23") if atarashii is None
             else atarashii)
        if p.runs:
            p.runs[0].text = V9.clean(s)
            for r in p.runs[1:]:
                r.text = ""
        n += 1
    judge(n == len(DAN), f"{len(DAN)}段落を差し替えた（実際 {n}）")

    print("  ── ② 表のセル")
    MID = {i: t.rows[0].cells[0].text.strip() for i, t in enumerate(d.tables)}
    n = 0
    for ban, mi, gyo, retsu, kyu, atarashii in SERU:
        if not MID[ban].startswith(mi):
            judge(False, f"表{ban} の見出しが合わない", f"実際『{MID[ban][:16]}』")
            continue
        c = d.tables[ban].rows[gyo].cells[retsu]
        if not c.text.strip().startswith(kyu):
            judge(False, f"表{ban} 行{gyo} 列{retsu} の旧が合わない",
                  f"実際『{c.text[:24]}』")
            continue
        KYU.append(c.text)
        V9.set_cell(c, atarashii)
        n += 1
    judge(n == len(SERU), f"{len(SERU)}セルを差し替えた（実際 {n}）")

    print("  ── ③ 入所状況の推移の表を足す（表8 の後）")
    t8 = d.tables[8]
    ok = t8.rows[0].cells[0].text.strip() == "施設の種類"
    judge(ok, "表8 が退所先の表である", f"実際『{t8.rows[0].cells[0].text[:10]}』")
    if ok:
        tn = V9.clone_table_after(t8, t8._tbl, len(HYOSOKU))
        for ri, gyo in enumerate(HYOSOKU):
            for ci, v in enumerate(gyo):
                V9.set_cell(tn.rows[ri].cells[ci], v)
    judge(len(d.tables) == 45, f"表が44 → 45（実際 {len(d.tables)}）")

    d.save(DST)

    d2 = docx.Document(DST)
    ato = "\n".join([p.text for p in d2.paragraphs]
                    + [c.text for t in d2.tables for r in t.rows
                       for c in r.cells])

    print("  ── 検査1　県調査の値と軽費老人ホーム")
    for go in ("老人福祉施設等入所状況調", S23.haba(A["特養"]),
               f"{A['在宅_老健']:.0f}人", "軽費老人ホーム"):
        judge(go in ato, f"「{go[:20]}」がある")
    print("  ── 検査2　古い前提が残っていないか")
    judge("公表の項目にないものは５件" not in ato,
          "「公表の項目にないものは５件」が0件")
    judge("Ver.2.22" not in ato, "「Ver.2.22」が0件")
    judge(not re.search(r"20\d\d年\d\d月\d\d日", ato),
          "⚠ 西暦の日付（元号との混在）が0件")
    judge("延べ利用日数による稼働率ではありません" in ato,
          "稼働率ではない旨がある")
    judge("短期入所の稼働" in ato, "短期入所の稼働が空白である旨がある")
    print("  ── 検査3　⚠ 既にあった数が消えていないか")
    kyu_zen = "\n".join(KYU)
    kieta = sorted(set(
        m.group(0) for m in re.finditer(
            r"[0-9０-９][0-9０-９,.]*(?:人|名|％|千円|件|頁|日|点)", mae)
        if mae.count(m.group(0)) > ato.count(m.group(0))
        and m.group(0) not in kyu_zen))
    judge(not kieta, "消えた数はすべて差し替えた旧の文に現れている",
          "／".join(kieta[:5]))
    print("  ── 検査4　⚠ 事業所名・他団体名・強調記号")
    for na in ("釜房みどりの園", "アルパイン", "蔵王グリーンホーム",
               "グループホームさわやか", "グループホームふれあい",
               "ながさか", "リリーハイム", "ケアハウスやまぶき"):
        judge(na not in ato, f"「{na}」が入っていない")
    judge("川崎市" not in ato and "福岡" not in ato,
          "川崎市・福岡県川崎町が混じらない")
    judge("**" not in ato, "強調記号（**）が0件")
    judge("［第10版］" in ato, "版が［第10版］")

    print()
    print(f"  {DST}")
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        for m in futeki:
            print(f"      {m}")
        sys.exit(1)
    print("   ○ 検査1〜4　不適合0件")


if __name__ == "__main__":
    main()
