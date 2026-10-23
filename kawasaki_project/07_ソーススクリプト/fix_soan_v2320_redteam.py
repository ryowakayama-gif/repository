# -*- coding: utf-8 -*-
"""協議用素案 Ver.2.31 → Ver.2.32（Red Team レビュー（令和8年10月24日）の指摘の是正）

  python3 07_ソーススクリプト/fix_soan_v2320_redteam.py

レビュー原本　09_元資料/レビュー/川崎町_第10期計画_RedTeamレビュー_R8.10.24受領.md
（レビューは素案 v2.29 を対象としている。v2.31 で解いた A02・A06（一部）・A07 は対象外）

  A10 1-1　「第1期から第9期までの25年間」→ 平成12年度〜令和8年度の27年間
  A05 1-5　基本指針案の別表11事項の件数を表から数え直す（本文「一部7・未対応3」→ 表は一部8・未対応2）
  A12 2-2　「詳細は第7章で論じます」→ 所得段階の乗率は 9-3 にある
  A06 3-3・5-4・10-4　交付金の要件を言い切る表現の残り3か所
       「そのまま要件イを満たします」「運用で確保すれば得点できる」「実績があれば要件を満たす」
  A11/B15 2-7・7-1　「8の事業所」→「8事業所」（概要版と同じ書き方。build_gaiyo の読み取りも改める）
  B11 生活支援コーディネーターの人数　25名 → 26人
       町提供の実績データ（データ入力フォーマット 02シート「生活支援ｻﾎﾟｰﾀｰ(SC)_人数」R6・R7見込・R8＝26）と
       国の評価（令和8年度 該当状況調査 26人）が一致。25名は地域福祉計画の資料の
       ユニバーサルサポーターの内訳（生活支援サポーター25名）の値で、時点・出所が異なる
  A12 9-3　「（9-4の感度表）」2か所 → 9-4 は存在しない。感度の表は 9-3 にある
       （check_sansho_R8.10.24.py の機械検索で発見）
  A03 5-5　「入所・居住系の利用者が1人増減すると月額約21円」に10人の場合（約211円）を書き添える
       （資料の選択２の根拠。表101 と 9-3 の調整交付金の見込交付割合から検算し、範囲を確かめる）

終了コード　0＝検査すべて適合／1＝1件でも不適合
"""
import re
import shutil
import sys
from decimal import Decimal

import docx

MOTO = "01_第10期_最新版成果品/川崎町_計画書素案_v2.31_協議用.docx"
SAKI = "01_第10期_最新版成果品/川崎町_計画書素案_v2.32_協議用.docx"
JISSEKI = "09_元資料/R7実績データ/★【川崎町様】データ入力フォーマット.xlsx"
KOFU = ("09_元資料/交付金評価/①令和８年度保険者機能強化推進交付金等（市町村）に係る全国集計結果/"
        "05-2_（支援）市町村評価指標（ⅱ活動指標群）.xlsx")

P_REP = [
    ("③　①の場における意見を", "一部の事業に活用した実績があれば要件を満たす（要件ウ・4点）",
     "一部の事業に活用した実績があれば要件を満たすとされている（評価指標の解説による。要件ウ・4点）", "A06"),
    # (段落の先頭, 旧, 新, 指摘)
    ("介護保険制度は平成12年4月に施行されて以降",
     "第1期から第9期までの25年間にわたり",
     "第1期（平成12年度）から第9期（令和8年度）までの27年間にわたり", "A10"),
    ("川崎町の所得段階は、令和5年度までの9段階から",
     "詳細は第7章で論じます。", "乗率の検討は第9章9-3に示します。", "A12"),
    ("⚠ 「評価結果の活用」の16点は令和6・7年度に満点であり",
     "そのまま要件イ（庁外の関係者の参画）を満たします。",
     "要件イ（庁外の関係者の参画）に当たる取組の一つとなり得ます。評価は、国の該当状況調査への"
     "回答と、実施日・参加者・議事の記録・反映の結果によります。", "A06"),
    ("町内のサービス提供体制は、介護老人福祉施設1施設", "これらを含む8の事業所が町内にあります",
     "これらを含む8事業所が町内にあります", "A11"),
    ("第2に、町内の8の事業所を運営しているのは3法人です。", "町内の8の事業所を",
     "町内の8事業所を", "A11"),
    ("令和8年度に0点であった指標は9つ", "生活支援コーディネーター25名の配置数",
     "生活支援コーディネーター26人の配置数", "B11"),
]
C_REP = [
    ("認知症の人と家族等が参画し意見を反映させる仕組みを整えたうえで策定に着手していることが要件。",
     "→ 第6章6-1の策定手続がそのまま要件となる",
     "→ 第6章6-1の策定手続がこの要件に対応する（評価は参画・反映・着手の実績と記録による）", "A06"),
    ("町内には8の事業所があり", "町内には8の事業所があり", "町内には8事業所があり", "A11"),
    ("配置数は県内1位だが会議への参加が3年連続0点。", "運用で確保すれば得点できる。",
     "運用で確保すれば評価の対象となり得る（評価は該当状況調査への回答と出席の記録による）。", "A06"),
    ("交：支援 目標Ⅰ「生活支援コーディネーター数」", "実：生活支援コーディネーター25名",
     "実：生活支援コーディネーター26人（令和7年度見込。町の実績データ）", "B11"),
    ("生活支援コーディネーター（SC）25名による地域づくり", "（SC）25名",
     "（SC）26人", "B11"),
    ("R7=SC25名", "R7=SC25名", "R7見込=SC26人", "B11"),
    ("（標準給付費見込額＋地域支援事業費）× 第1号被保険者負担割合", "（9-4の感度表）", "（9-3の感度表）", "A12"),
    ("後期高齢者割合・所得段階別人口により交付率を算定", "（9-4の感度表）", "（9-3の感度表）", "A12"),
    ("介護予防サロン84名、スマイル40名、レク29名、傾聴24名、SC25名", "SC25名",
     "生活支援サポーター25名", "B11"),
]


def kakikae(p, s):
    p.runs[0].text = s
    for r in p.runs[1:]:
        r.text = ""


def set_cell(c, s):
    kakikae(c.paragraphs[0], s)
    for q in c.paragraphs[1:]:
        q._p.getparent().remove(q._p)


def zenbun(d):
    return "\n".join([p.text for p in d.paragraphs]
                     + [c.text for t in d.tables for r in t.rows for c in r.cells])


def sc_genpyo():
    """町の実績データと国の評価の生活支援コーディネーター数。"""
    import openpyxl
    ws = openpyxl.load_workbook(JISSEKI, read_only=True, data_only=True)["02_総合事業_介護予防"]
    rows = list(ws.iter_rows(values_only=True))
    h = rows[3]
    r = [x for x in rows if x and any(v and "SC" in str(v) for v in x)][0]
    machi = dict(zip(h, r))
    ws = openpyxl.load_workbook(KOFU, read_only=True, data_only=True)["目標Ⅰ"]
    rows = list(ws.iter_rows(values_only=True))
    hd = rows[4]
    col = [j for j, x in enumerate(hd) if x and str(x).startswith("生活支援コーディネーター数\n（該当状況調査）")][0]
    kuni = [r for r in rows if r and str(r[1]) == "043240"][0]   # 宮城県川崎町（福岡県川崎町と取り違えない）
    return machi, kuni[col], kuni[2]


def main():
    futeki = []

    def judge(ok, mi, setsu=""):
        print(("   ○ " if ok else "   × ") + mi + (("　" + setsu) if setsu else ""))
        if not ok:
            futeki.append(mi)

    print("協議用素案 Ver.2.31 → Ver.2.32（Red Team レビューの指摘の是正）")
    shutil.copy2(MOTO, SAKI)
    d = docx.Document(SAKI)
    mae = zenbun(d)

    print("  ── ⓪ 原票で確かめる")
    machi, kuni, na = sc_genpyo()
    judge(na == "川崎町" and kuni == 26, f"国の評価（令和8年度・該当状況調査）{na} {kuni}人")
    judge(machi.get("R6") == 26 and machi.get("R7見込") == 26,
          f"町の実績データ R6 {machi.get('R6')}・R7見込 {machi.get('R7見込')}人")

    print("  ── ① 本文の段落")
    for atama, kyu, shin, ID in P_REP:
        h = [p for p in d.paragraphs if p.text.strip().startswith(atama) and kyu in p.text]
        judge(len(h) == 1, f"{ID}『{atama[:18]}』")
        for p in h:
            kakikae(p, p.text.replace(kyu, shin))

    print("  ── ② 表のセル")
    for atama, kyu, shin, ID in C_REP:
        seen, n = [], 0
        for t in d.tables:
            for r in t.rows:
                for c in r.cells:
                    if c._tc in seen:
                        continue
                    seen.append(c._tc)
                    if c.text.strip().startswith(atama) and kyu in c.text:
                        set_cell(c, c.text.replace(kyu, shin))
                        n += 1
        judge(n >= 1, f"{ID}『{atama[:18]}』（{n}か所）")
    # KPI の表の生活支援コーディネーター数の現状値
    n = 0
    for t in d.tables:
        if [c.text for c in t.rows[0].cells][:3] == ["指標", "区分", "現状値"]:
            for r in t.rows:
                if r.cells[0].text.strip() == "生活支援コーディネーター数" and r.cells[2].text.strip() == "25名":
                    set_cell(r.cells[2], "26人（令和7年度見込）")
                    n += 1
    judge(n == 1, "B11 10-4 KPI の現状値 25名 → 26人")

    print("  ── ③ A05 1-5 の11事項の件数を表から数える")
    t = [t for t in d.tables if [c.text.strip() for c in t.rows[0].cells][:3]
         == ["事項", "確認すべき指標・状況", "対応"]][0]
    k = [r.cells[2].text.strip() for r in t.rows[1:]]
    zumi, ichibu, mi = k.count("対応済"), k.count("一部"), k.count("未対応")
    judge(zumi + ichibu + mi == len(k) == 11, f"11事項＝対応済{zumi}・一部{ichibu}・未対応{mi}")
    h = [p for p in d.paragraphs if p.text.startswith("対応済1・一部7・未対応3です。未対応の3事項は")]
    judge(len(h) == 1, "1-5 の本文（旧の件数）")
    for p in h:
        kakikae(p, p.text.replace("対応済1・一部7・未対応3です。未対応の3事項は",
                                  f"対応済{zumi}・一部{ichibu}・未対応{mi}です。未対応の{mi}事項は"))

    print("  ── ④ A03 入所・居住系の10人の場合の月額（表101 から検算）")
    t101 = [t for t in d.tables if [c.text.strip() for c in t.rows[0].cells][:3]
            == ["記号", "項目", "第10期計"]][0]
    hit = [p for p in d.paragraphs if "入所・居住系の利用者が1人増減すると保険料基準額は月額約21円動きます。" in p.text]
    judge(len(hit) == 1, "5-5 の段落")
    tanka = Decimal(hit[0].text.split("1人1月あたり給付費は約")[1].split("万円")[0]) * 10  # 千円
    rates = [Decimal(x) / 100 for x in re.findall(r"令和\d+年度([\d.]+)％",
             [p.text for p in d.paragraphs if "見込交付割合は全国の仮置値による自動計算値" in p.text][0])]
    judge(len(rates) == 3, f"調整交付金の見込交付割合 {'・'.join(str(r * 100) for r in rates)}％")
    n_ = Decimal(re.sub(r"[^\d.]", "", [r.cells[2].text for r in t101.rows if r.cells[0].text == "③"][0]))
    d_j = sum(10 * tanka * 12 * (Decimal("0.23") + Decimal("0.05") - e) for e in rates)  # 千円
    d_m = d_j * 1000 / Decimal("0.96") / n_ / 12
    judge(Decimal(205) <= d_m <= Decimal(215), f"10人で月額 {d_m:.1f}円（約211円と矛盾しない）")
    for p in hit:
        kakikae(p, p.text.replace("月額約21円動きます。", "月額約21円、10人では約211円動きます（基金の取崩しの額によりません）。"))

    d.save(SAKI)
    ato = zenbun(docx.Document(SAKI))
    print("  ── 検査1　古い記述が残っていない")
    for go in ("9-4の感度表", "25年間", "第7章で論じます", "得点できる", "得点できます", "そのまま要件", "8の事業所",
               "生活支援コーディネーター25名", "SC25名", "対応済1・一部7・未対応3"):
        judge(go not in ato, f"「{go}」0件")
    print("  ── 検査2　未確定の印・他団体名")
    for w in ("［要確認］", "［要協議］", "［要内訳］", "【町確認】", "大雪", "北塩原", "金ヶ崎", "川崎市"):
        judge(ato.count(w) == mae.count(w), f"「{w}」増減なし")
    print("  ── 検査3　消えた数は旧の文の中のもの")
    kyu = "".join(x[1] for x in P_REP + C_REP) + "対応済1・一部7・未対応3です。未対応の3事項は25名"
    kieta = sorted({m.group(0) for m in re.finditer(r"[0-9][0-9,.]*(?:人|名|％|%|千円|件|頁|点|円|年間)", mae)
                    if mae.count(m.group(0)) > ato.count(m.group(0)) and m.group(0) not in kyu})
    judge(not kieta, "消えた数 0件", "／".join(kieta[:5]))
    print(f"  {SAKI}")
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        sys.exit(1)
    print("   ○ 不適合0件")


if __name__ == "__main__":
    main()
