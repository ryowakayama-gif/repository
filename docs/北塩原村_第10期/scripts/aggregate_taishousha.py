# -*- coding: utf-8 -*-
"""調査対象者データの集計（個人情報を出力しない）。

   ★取扱い（doc16 個人情報取扱管理ルール §5-2・§6）
     ・分析に用いるのは要介護度・年齢・性別・行政区の属性情報のみ。
       氏名・住所・生年月日・被保険者番号・管理番号は読み込んでも出力しない。
     ・被保険者番号は2つの名簿の重複件数を数えるためだけに用い、値は出力しない。
     ・該当数10未満の区分は「＜10」と伏せる（§6 少数値の扱い）。
     ・地区×性別×年齢の多重クロスは行わない（§6 特定リスク）。
     ・入力ファイルはリポジトリに複製しない。集計値のみを CSV に書き出す。

   入力（村から受領した対象者抽出データ。リポジトリ外に置く）
     DAHIHOKENSHA … 介護予防・日常生活圏域ニーズ調査の対象者
     DBJUKYUSHA   … 在宅介護実態調査の対象者（要介護度を含む）
"""
import csv, os, sys, collections

SRC_DIR = "/root/.claude/uploads/134138ca-61f7-57d3-9e9b-5f081a1a345d"
F_NEEDS = os.path.join(SRC_DIR, "3e0d9fcc-______.xlsx")
F_HOME = os.path.join(SRC_DIR, "d8ec6417-________.xlsx")
OUT = "/home/user/repository/docs/北塩原村_第10期/data/第10期_調査対象者の属性集計.csv"
MASK = 10          # この数未満の区分は伏せる

AGE = [(0, 64, "65歳未満"), (65, 74, "65〜74歳"), (75, 84, "75〜84歳"), (85, 200, "85歳以上")]
def age_band(v):
    try:
        a = int(v)
    except (TypeError, ValueError):
        return "不明"
    return next(lab for lo, hi, lab in AGE if lo <= a <= hi)

def cell(n):
    return str(n) if n >= MASK or n == 0 else f"＜{MASK}"


def load(path, sheet_prefix, cols):
    """指定した列だけを取り出す。氏名・住所・生年月日の列は読み込まない。"""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    name = next(n for n in wb.sheetnames if n.startswith(sheet_prefix))
    ws = wb[name]
    rows = ws.iter_rows(values_only=True)
    hdr = list(next(rows))
    idx = {k: hdr.index(v) for k, v in cols.items()}
    out = [{k: r[i] for k, i in idx.items()} for r in rows if r[0] is not None]
    wb.close()
    return out, name


def sheet_counts(path):
    """名簿シートの件数と管理番号の形式だけを確認する（番号の値は出力しない）"""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    out = {}
    for n in wb.sheetnames:
        ws = wb[n]
        rows = list(ws.iter_rows(min_row=1, max_row=3, values_only=True))
        if not rows or rows[0][0] != "管理番号":
            continue
        cnt = sum(1 for r in ws.iter_rows(min_row=2, values_only=True) if r[0] is not None)
        sample = [r[0] for r in rows[1:] if r and r[0] is not None]
        form = "数字のみ" if all(str(x).isdigit() for x in sample) else \
               ("英字＋数字" if sample else "―")
        out[n] = (cnt, form)
    wb.close()
    return out


def main():
    recs = []
    def add(chosa, kubun, ku, n, bikou=""):
        recs.append({"調査": chosa, "区分": kubun, "分類": ku, "人数": cell(n), "備考": bikou})

    # ── ニーズ調査の対象者 ──
    needs, sh1 = load(F_NEEDS, "DAHIHOKENSHA", {"age": "年齢", "sex": "性別", "ku": "行政区コード", "hino": "被保番号"})
    print(f"■ 介護予防・日常生活圏域ニーズ調査の対象者（{sh1}）　{len(needs)}人\n")
    for key, lab, f in (("age", "年齢階級", age_band), ("sex", "性別", lambda v: str(v))):
        c = collections.Counter(f(r[key]) for r in needs)
        print(f"  【{lab}】")
        for k in sorted(c):
            print(f"    {k:10s} {cell(c[k]):>6s}  {c[k]/len(needs):6.1%}")
            add("ニーズ調査", lab, k, c[k])
        print()
    c = collections.Counter(str(r["ku"]) for r in needs)
    print("  【行政区コード別】（23区分。個々の区は人数が少ないため地区に束ねて用いる）")
    for k in sorted(c):
        print(f"    {k:10s} {cell(c[k]):>6s}  {c[k]/len(needs):6.1%}")
        add("ニーズ調査", "行政区コード", k, c[k], "")
    print()

    # 行政区コードの先頭の桁が地区に対応するとみて、第9期の対象者数と照合する
    CHIKU = {"1": "北山", "2": "大塩", "3": "桧原", "4": "裏磐梯"}
    K9 = {"北山": 329, "大塩": 200, "桧原": 114, "裏磐梯": 240}   # 第9期の対象者数（doc01 §5-5）
    g = collections.Counter(CHIKU.get(str(r["ku"])[0], "不明") for r in needs)
    print("  【地区別】（行政区コードの先頭の桁による。第9期の対象者数と照合）")
    print(f"    {'地区':8s} {'第10期':>7s} {'第9期':>7s} {'差':>6s}")
    for k in ("北山", "大塩", "桧原", "裏磐梯", "不明"):
        if not g[k]:
            continue
        d = g[k] - K9.get(k, 0)
        print(f"    {k:8s} {cell(g[k]):>7s} {K9.get(k, 0):7d} {d:+6d}")
        add("ニーズ調査", "地区", k, g[k], f"第9期{K9.get(k, 0)}人（差{d:+d}）")
    print(f"    {'計':8s} {sum(g.values()):7d} {sum(K9.values()):7d} "
          f"{sum(g.values())-sum(K9.values()):+6d}")
    print("    ※ 4地区とも第9期の対象者数と近い。先頭の桁と地区の対応はこの照合による推定であり、")
    print("      村への確認を要する。")
    print()

    # ── 在宅介護実態調査の対象者 ──
    home, sh2 = load(F_HOME, "DBJUKYUSHA", {"age": "年齢", "sex": "性別", "ku": "行政区名",
                                            "yokaigo": "要介護度", "hino": "被保番号"})
    print(f"\n■ 在宅介護実態調査の対象者（{sh2}）　{len(home)}人\n")
    yc = collections.Counter(str(r["yokaigo"]).strip() for r in home)
    print("  【要介護度別】")
    for k in sorted(yc):
        print(f"    {k:12s} {cell(yc[k]):>6s}  {yc[k]/len(home):6.1%}")
        add("在宅調査", "要介護度", k, yc[k])
    print()
    for key, lab, f in (("age", "年齢階級", age_band), ("sex", "性別", lambda v: str(v))):
        c = collections.Counter(f(r[key]) for r in home)
        print(f"  【{lab}】")
        for k in sorted(c):
            print(f"    {k:12s} {cell(c[k]):>6s}  {c[k]/len(home):6.1%}")
            add("在宅調査", lab, k, c[k])
        print()
    # 行政区は23区分あり、ほぼすべてが10人未満である。伏せ字にしても「その集落に
    # 要介護認定を受けた方がいる」ことが分かるため、区分の一覧は出力しない（doc16 §6）。
    nku = len({str(r["ku"]) for r in home})
    over = sum(1 for _, v in collections.Counter(str(r["ku"]) for r in home).items() if v >= MASK)
    print(f"  【行政区】{nku}区分。うち{MASK}人以上は{over}区分のみ。")
    print(f"    小規模な集落が並ぶため区分ごとの人数は出力しない（doc16 §6 特定リスク）。")
    print(f"    地区（北山・大塩・桧原・裏磐梯）別に束ねるには行政区名と地区の対応表を要する。")
    add("在宅調査", "行政区", f"{nku}区分（うち{MASK}人以上は{over}区分）", 0,
        "個々の行政区の人数は特定リスクのため出力しない。地区別に束ねるには対応表が必要")


    # ── 在宅調査：要介護度の区分と回収見込み ──
    KEI = {"要介護１", "要介護２"}
    kei = sum(v for k, v in yc.items() if k in KEI)
    juu = len(home) - kei
    print("  【要介護度を2区分に束ねた場合と回収の見込み】")
    print(f"    {'区分':14s} {'対象者':>6s} {'回収見込（45.8%）':>16s} {'同（60%）':>10s}")
    for lab, n in (("要介護1〜2", kei), ("要介護3以上", juu)):
        print(f"    {lab:14s} {n:6d} {round(n*0.458):16d} {round(n*0.60):10d}")
        add("在宅調査", "要介護度2区分", lab, n,
            f"回収見込 45.8%で{round(n*0.458)}件／60%で{round(n*0.60)}件")
    print(f"    {'計':14s} {len(home):6d} {round(len(home)*0.458):16d} {round(len(home)*0.60):10d}")
    print("    ※ 45.8%は第9期の在宅介護実態調査の有効回収率。")
    print("    ※ 名簿に要支援の方は含まれていない（全員が要介護1以上）。")
    print()

    # ── 2つの名簿の重複（被保険者番号は値を出力しない） ──
    a = {str(r["hino"]).strip() for r in needs if r["hino"] is not None}
    b = {str(r["hino"]).strip() for r in home if r["hino"] is not None}
    dup = len(a & b)
    print(f"\n■ 2つの名簿の重複　{dup}人")
    print(f"  ニーズ調査の対象者のうち、在宅介護実態調査にも含まれる方の人数。")
    print(f"  重複があると同じ方に2種類の調査票が届く。宛名の重複と封入の誤りを防ぐため確認した。")
    add("両調査", "重複", "両方の名簿に含まれる方", dup, "被保険者番号で突合。番号の値は出力しない")

    # ── 名簿シートの件数と管理番号の形式 ──
    print("\n\n■ 名簿シートの件数と管理番号の形式")
    for path, nm in ((F_NEEDS, "ニーズ調査"), (F_HOME, "在宅調査")):
        for sheet, (cnt, form) in sheet_counts(path).items():
            print(f"  {nm:8s} シート「{sheet}」 {cnt:4d}件　管理番号の形式：{form}")

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8-sig", newline="") as fp:
        w = csv.DictWriter(fp, fieldnames=["調査", "区分", "分類", "人数", "備考"])
        w.writeheader(); w.writerows(recs)
    print(f"\n保存: {OUT}（集計値のみ・{len(recs)}行）")


if __name__ == "__main__":
    main()
