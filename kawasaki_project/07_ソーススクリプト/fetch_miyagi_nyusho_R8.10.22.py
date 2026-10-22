# -*- coding: utf-8 -*-
"""宮城県「老人福祉施設等入所状況調」から町内施設の推移を採録する
（令和8年10月22日）

  python3 07_ソーススクリプト/fetch_miyagi_nyusho_R8.10.22.py

なぜ要るか
  ⚠ **残る5件（稼働率・入所前の居所・休止床の数・経営の状況・町への要望）**
  のうち、稼働率と休止床は「公表の項目にない」として残していた。
  宮城県が**3か月ごとに施設別の入所状況を公表している**ことが分かった。

    入口　https://www.pref.miyagi.jp/soshiki/chouju/nyusyojyoukyou.html
    単位　施設別（特養・老健・養護・軽費・認知症GH・介護医療院）
    時点　3か月ごとの1日現在（報告月 4月・7月・10月・1月）
    項目　定員／入所現員／入所希望者／新たな入所者（人員・平均待機期間）／
          退所者（人員・平均利用期間）／退所理由（死亡・入院・在宅・他施設）

  公表システムが**年1回の報告の1時点**であるのに対し、こちらは
  **3か月ごとの時系列**である。令和5年4月から令和8年7月まで14期。

⚠ **定義（県の「入所状況調書の留意事項」による）**
  ・定員　　　　報告月1日現在の**長期（短期分を除く）**の定員
  ・入所現員　　報告月1日現在入所している長期の利用者数
  ・入所希望者　報告月1日現在の待機者の人数
  　⚠ **複数施設へ申し込めるため重複して計上されている**（県の注記）
  ・人員　　　　前回報告月1日から今回報告月の前月末までの入退所の人数
  ・平均待機期間　利用申込みから入所までの延べ月数÷人員
  ・平均利用期間　退所者の利用期間月数÷人員
  ・退所理由　　退所した時点での理由。退所者数と合計が一致する

⚠ **所在地の列は市町村ごとに1行だけ値が入る。**
  その市町村の施設は、所在地のある行から**次に所在地のある行の直前まで**。
  所在地のない行を直前の市町村に無条件に属させると取り違える。
  現にそれで別の市町村の同名施設（「ふれあい」）を拾った。

出すもの
  09_元資料/その他データ/県入所状況調_R8.10.22取得/*.xls（原典14件）
  09_元資料/その他データ/県入所状況調_川崎町_R8.10.22.csv（町内の施設）

  ⚠ **一度取った原典は取り直さない。** 既にあれば通信しない。

終了コード　0＝検査1〜8すべて適合／1＝1件でも不適合
"""
import csv
import os
import re
import subprocess
import sys
import tempfile
import time
import urllib.request

import openpyxl

MOTO = "https://www.pref.miyagi.jp/documents/8866/"
HOZON = "09_元資料/その他データ/県入所状況調_R8.10.22取得"
SAKI = "09_元資料/その他データ/県入所状況調_川崎町_R8.10.22.csv"
MACHI = "川崎町"

# 入口のページが示すファイル名と、その時点
#   ⚠ ファイル名に時点が入らないもの（令和5年度）があるため対応表で持つ。
#     対応は入口のページの表示（「令和○年○月１日現在」）による。
FILES = [
    ("nyuusyor050401_syuusei.xls", "R5.4.1"),
    ("nyuusyor050701_syuusei.xls", "R5.7.1"),
    ("nyuusyor051001_syuusei.xls", "R5.10.1"),
    ("nyuusyor060101_revise2.xls", "R6.1.1"),
    ("nyuusyo20240401_revise3.xls", "R6.4.1"),
    ("nyuusyo20240701_revise3.xls", "R6.7.1"),
    ("nyuusyo20241001_revise2.xls", "R6.10.1"),
    ("nyuusyo20250101_revise.xls", "R7.1.1"),
    ("nyuusyo20250401_revise.xls", "R7.4.1"),
    ("nyuusyo20250701_revise2.xls", "R7.7.1"),
    ("nyuusyo20251001_revise.xls", "R7.10.1"),
    ("nyuusyo20260101_revise.xls", "R8.1.1"),
    ("nyuusyo20260401_revise.xls", "R8.4.1"),
    ("nyuusyo20260701.xls", "R8.7.1"),
]
# 定義・様式（あわせて退避する）
SANKO = ["742313.pdf", "nyusyozyoukyou_yokuarushitumon.pdf",
         "nyusyozyoukyou_houkokuyoushiki_r060401kaisei_ver2.xls"]

JUN = [h for _, h in FILES]
SHICHO = re.compile(r"^[^\d]{2,6}[市町村]$")
KOUMOKU = ["定員", "入所現員", "入所希望者", "新入所人員", "平均待機期間月",
           "退所人員", "平均利用期間月", "死亡", "入院", "在宅", "他施設"]


def norm(s):
    return re.sub(r"\s+", "", str(s)) if s is not None else ""


def toru(name, saki):
    """原典を取る。既にあれば通信しない。"""
    if os.path.exists(saki) and os.path.getsize(saki) > 50000:
        return False
    req = urllib.request.Request(
        MOTO + name, headers={"User-Agent": "Mozilla/5.0 (kawasaki-town-plan)"})
    with urllib.request.urlopen(req, timeout=90) as r:
        b = r.read()
    os.makedirs(os.path.dirname(saki), exist_ok=True)
    open(saki, "wb").write(b)
    time.sleep(1.5)
    return True


def xlsx_ni(xls, tmpdir):
    """.xls を LibreOffice で .xlsx に直す（openpyxl で読むため）。"""
    out = os.path.join(tmpdir,
                       os.path.basename(xls)[:-4] + ".xlsx")
    if os.path.exists(out):
        return out
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx",
                    "--outdir", tmpdir, xls],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                   timeout=300)
    return out if os.path.exists(out) else None


def yomu(xlsx, hiduke):
    """町内の施設の行を返す。"""
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    out = []
    for sn in wb.sheetnames:
        if sn.startswith("000") or "仙台" in sn:
            continue
        ws = wb[sn]
        anchor = [(r, norm(ws.cell(r, 2).value))
                  for r in range(1, ws.max_row + 1)
                  if SHICHO.match(norm(ws.cell(r, 2).value))]
        for i, (r0, machi) in enumerate(anchor):
            if machi != MACHI:
                continue
            r1 = anchor[i + 1][0] if i + 1 < len(anchor) else ws.max_row + 1
            for r in range(r0, r1):
                na = norm(ws.cell(r, 3).value)
                if not na or na.isdigit():
                    continue
                v = [ws.cell(r, c).value for c in range(5, 16)]
                out.append([hiduke, sn.replace("・県分", "").replace("GH", "ＧＨ"),
                            na] + ["" if x is None else x for x in v])
    return out


def main():
    futeki = []
    mi = []

    def judge(ok, m, setsu=""):
        print(("   ○ " if ok else "   × ") + m + (("　" + setsu)
                                                  if setsu else ""))
        if not ok:
            futeki.append(m)

    print("宮城県「老人福祉施設等入所状況調」からの採録（令和8年10月22日）")
    print()
    print("  ── ① 原典を取る")
    tsushin = 0
    for name, hi in FILES + [(s, "") for s in SANKO]:
        saki = os.path.join(HOZON, name)
        try:
            atarashii = toru(name, saki)
        except Exception as e:
            judge(False, f"{name} が取れない", str(e)[:60])
            mi.append(f"{name}：取得できなかった（{e}）")
            continue
        tsushin += 1 if atarashii else 0
    print(f"     {HOZON}/　"
          f"{len([f for f in os.listdir(HOZON)]) if os.path.isdir(HOZON) else 0}件"
          f"（通信 {tsushin}件）")

    print("  ── ② 町内の施設を取り出す")
    dedata = []
    with tempfile.TemporaryDirectory() as td:
        for name, hi in FILES:
            xls = os.path.join(HOZON, name)
            if not os.path.exists(xls):
                mi.append(f"{hi}：原典がない")
                continue
            x = xlsx_ni(xls, td)
            if not x:
                judge(False, f"{hi} を xlsx に直せない")
                mi.append(f"{hi}：LibreOffice で読めなかった")
                continue
            gyo = yomu(x, hi)
            dedata += gyo
            print(f"     {hi}　{len(gyo)}施設")

    # ══════ 検査1　14の時点がそろっているか
    print("  ── 検査1　時点がそろっているか")
    aru = sorted(set(r[0] for r in dedata), key=JUN.index)
    judge(len(aru) == len(FILES),
          f"{len(FILES)}時点（実際 {len(aru)}）", "／".join(aru[:4]))

    # ══════ 検査2　⚠ 取り違えていないか（施設の顔ぶれ）
    print("  ── 検査2　⚠ 別の市町村の施設を拾っていないか")
    shisetsu = sorted(set(r[2] for r in dedata))
    BEKI = {"釜房みどりの園", "第二釜房みどりの園", "アルパイン川崎",
            "さわやか", "ふれあい"}
    yobun = [s for s in shisetsu
             if s not in BEKI and "蔵王グリーンホーム" not in s]
    judge(not yobun, f"町内の施設だけ（{len(shisetsu)}件）",
          "／".join(yobun[:4]))
    # 同じ時点に同じ施設が2件ないか（＝取り違えの印）
    kasanari = []
    for hi in aru:
        na = [r[2] for r in dedata if r[0] == hi]
        kasanari += [n for n in set(na) if na.count(n) > 1]
    judge(not kasanari, "同じ時点に同じ施設が重ねて出ていない",
          "／".join(sorted(set(kasanari))[:3]))

    # ══════ 検査3　入所現員が定員を超えていないか
    print("  ── 検査3　入所現員が定員を超えていないか")
    koe = []
    for r in dedata:
        try:
            tei, gen = float(r[3]), float(r[4])
        except (ValueError, TypeError):
            continue
        if tei and gen > tei:
            koe.append(f"{r[0]} {r[2]} {gen:.0f}/{tei:.0f}")
    judge(not koe, "入所現員が定員以内", "／".join(koe[:3]))

    # ══════ 検査4　⚠ **原典の内部の食い違い**を数える
    #   県の留意事項（7）は「退所者の人数と退所理由の合計人数は等しくなる」
    #   としている。⚠ **これは当方の読み取りの検査ではなく、
    #   公表値そのものの食い違いを数えるものである。**
    #   食い違いがあっても当方の不適合とはせず、件数と内容を残す。
    print("  ── 検査4　⚠ 原典の内部の食い違い（県の留意事項（7））")
    chigai = []
    for r in dedata:
        try:
            n = float(r[8] or 0)
            riyu = sum(float(r[i] or 0) for i in (10, 11, 12, 13))
        except (ValueError, TypeError):
            continue
        if abs(n - riyu) > 0.01:
            chigai.append(f"{r[0]} {r[2]}　退所{n:.0f}人・理由計{riyu:.0f}人")
    if chigai:
        print(f"     △ 公表値の食い違い {len(chigai)}件"
              f"（{len(dedata)}行中）。⚠ **当方の読み取りではなく"
              f"公表値そのものの食い違いです。**")
        for c in chigai:
            print(f"        {c}")
        mi.append(f"県の公表値に内部の食い違いが{len(chigai)}件"
                  f"（退所者数と退所理由の合計が合わない）。"
                  f"数値を用いる際は注記する")
    else:
        print("     ○ 公表値に内部の食い違いはありません")
    # 当方の読み取りが原典を再現しているかは、検査2・3・5で見る

    # ══════ 検査5　公表システムの定員と突き合わせる
    print("  ── 検査5　公表システムの定員と突き合わせる")
    #   ⚠ **県調査は広域型と地域密着型を合算している。**
    #     釜房みどりの園は 介護老人福祉施設30人＋地域密着型10人＝40人。
    saigo = [r for r in dedata if r[0] == JUN[-1]]
    teiin = {r[2]: r[3] for r in saigo}
    BEKI_TEIIN = {"釜房みどりの園": 40, "第二釜房みどりの園": 20,
                  "アルパイン川崎": 86, "さわやか": 9, "ふれあい": 18}
    chigai = [f"{k} 県={teiin.get(k)} 公表={v}"
              for k, v in BEKI_TEIIN.items() if teiin.get(k) != v]
    judge(not chigai, "最新時点の定員が公表システムと整合",
          "／".join(chigai[:3]))

    # ══════ 書き出し
    os.makedirs(os.path.dirname(SAKI), exist_ok=True)
    with open(SAKI, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["時点", "種別", "施設名"] + KOUMOKU)
        for r in sorted(dedata, key=lambda x: (x[2], JUN.index(x[0]))):
            w.writerow(r[:14])
    print()
    print("  ── ③ 書き出し")
    print(f"     {SAKI}　{len(dedata)}行")

    # ══════ 検査6　読み返せるか
    print("  ── 検査6　書き出したものが読み返せるか")
    with open(SAKI, encoding="utf-8-sig", newline="") as f:
        yomi = list(csv.DictReader(f))
    judge(len(yomi) == len(dedata),
          f"{len(dedata)}行が読み返せる（実際 {len(yomi)}）")

    # ══════ 検査7　⚠ 休止・廃止を見落としていないか
    print("  ── 検査7　⚠ 休止・廃止があるか")
    kyushi = sorted(set(r[2] for r in dedata if "休止" in r[2]))
    kieta = []
    saisho = set(r[2].split("（")[0] for r in dedata if r[0] == JUN[0])
    ima = set(r[2].split("（")[0] for r in dedata if r[0] == JUN[-1])
    kieta = sorted(saisho - ima)
    if kyushi or kieta:
        print(f"     △ 休止の記載 {len(kyushi)}件／"
              f"一覧から消えた施設 {len(kieta)}件")
        for s in kyushi:
            print(f"        休止：{s}")
        for s in kieta:
            print(f"        ⚠ {JUN[0]}にあり{JUN[-1]}にない：{s}")
    else:
        print("     ○ 休止・廃止はありません")

    # ══════ 検査8　⚠ 埋まらないものを数で残す
    print("  ── 検査8　⚠ この調査でも埋まらないもの（未実施）")
    nokoru = [
        "入所前の居所（在宅か医療機関か）：調査の項目にない"
        "（退所理由は死亡・入院・在宅・他施設の別がある）",
        "稼働率（延べ利用日数÷延べ定員日数）：調査は時点の在所者数であり、"
        "延べではない",
        "床単位の休止：調査は施設単位。床単位の休止は定員と入所現員の差の"
        "恒常性で代える",
        "経営の状況（収支・人件費率）：調査の項目にない",
        "町への要望：調査の項目にない",
        "短期入所の稼働：⚠ 定員・入所現員とも**長期（短期分を除く）**である",
    ]
    mi += nokoru
    print(f"     △ {len(nokoru)}件（「適合」に丸めていません）")
    for s in nokoru:
        print(f"        {s}")

    print()
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        for m in futeki:
            print(f"      {m}")
        sys.exit(1)
    print(f"   ○ 検査1〜8　不適合0件"
          f"（⚠ 埋まらないもの {len(mi)}件は別に数えています）")


if __name__ == "__main__":
    main()
