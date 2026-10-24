# -*- coding: utf-8 -*-
"""協議用素案 Ver.2.27 → Ver.2.28（見える化システムの事業所分析の反映）
（令和8年10月23日）

  python3 07_ソーススクリプト/fix_soan_v2280_jigyosho.py

ご指示「事業所データを協議用素案に反映して下さい」（令和8年10月23日）

反映するもの
  地域包括ケア「見える化」システムの事業所分析（事業所別指標・各年9月・令和3〜7年）
  定義は厚生労働省「地域分析の手引き」（令和8年9月改訂版）66頁。
    稼働率　1月あたり延べ利用者数÷（定員×30日）
            （原表の「定員充足率」＝延べ÷定員を、回転率と同じく9月の30日で割ったもの）
    回転率　1月あたり延べ利用者数÷（定員×30日）（通所系。手引きの定義どおり）

反映先
  2-7　データの出どころの段落・表26（把握の状況）・新たな小見出し「5年間の推移」と表
  別表（表4）事項八　「延べ利用日数による稼働率は把握しておらず」「短期入所の稼働は
       いずれの資料にもない」を改める
  7-1　提供体制の概況に、延べ利用日数による稼働率を加える
  7-2　給付費の平坦化に、介護職員（常勤換算）の推移を加える
  7-6　定点把握に、見える化システムの事業所分析を加える
  9-1　表102（介護施設整備の基本方針）と、地域密着型の空き床の扱い［要協議］

⚠ 事業所ごとの職員数は掲げない（町内は8事業所・3法人であり法人が識別され得る。
  素案2-7の方針）。介護職員は施設・居住系の合計（指定の重複を除く）と
  定員100人あたりで示す。
⚠ 同じ種類が2つあるもの（地域密着型介護老人福祉施設・認知症対応型共同生活介護）は
  延べと定員を合計して示す。
⚠ 宮城県「老人福祉施設等入所状況調」（時点の在所者数）とは定義が違う。並べない。

数値の正本　09_元資料/その他データ/事業所別5か年推移_川崎町_R8.10.23.csv
            （build_jigyosho_5nen_R8.10.23.py が宮城県分の原表から抜き出したもの）
終了コード　0＝検査1〜7すべて適合／1＝1件でも不適合
"""
import copy
import csv
import glob
import os
import re
import shutil
import statistics
import sys

import docx
import openpyxl
from docx.oxml.ns import qn

MOTO = "01_第10期_最新版成果品/川崎町_計画書素案_v2.27_協議用.docx"
SAKI = "01_第10期_最新版成果品/川崎町_計画書素案_v2.28_協議用.docx"
CSV_IN = "09_元資料/その他データ/事業所別5か年推移_川崎町_R8.10.23.csv"
KEN_DIR = "09_元資料/その他データ/事業所別5か年推移_R8.10.23受領/宮城県"
NEN = ["令和3年", "令和4年", "令和5年", "令和6年", "令和7年"]
HI = 30
SHUTTEN = "地域包括ケア「見える化」システム（事業所分析・各年9月）"
SHISETSU = ("51", "52", "54", "32")          # 施設・居住系（重複のない指定）


def yomu():
    with open(CSV_IN, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    v = {}
    for r in rows:
        x = float(r["値"]) if r["値"] not in ("", "None") else None
        v[(r["指標"], r["サービス種類コード"], r["事業所番号"], r["年"],
           r["シート"][:4])] = x
    return rows, v


def goukei(rows, shihyo, shu, nen, sheet=None):
    s = 0
    hit = False
    for r in rows:
        if (r["指標"] == shihyo and r["サービス種類コード"] == shu and r["年"] == nen
                and (sheet is None or r["シート"].startswith(sheet))):
            if r["値"] not in ("", "None"):
                s += float(r["値"])
                hit = True
    return s if hit else None


def ken_chuo(key, shu):
    """宮城県内の令和7年の中央値（施設・居住系の稼働率／通所の回転率）。"""
    f = glob.glob(os.path.join(KEN_DIR, f"04_宮城県_*{key}*.xlsx"))[0]
    ws = openpyxl.load_workbook(f, read_only=True, data_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    h = [i for i, r in enumerate(rows) if r and r[0] == "No."][0]
    vals = []
    for r in rows[h + 1:]:
        if r[4] == shu:
            try:
                n, t = float(r[18]), float(r[23])
            except (TypeError, ValueError):
                continue
            if t:
                vals.append(n / t / HI)
    return statistics.median(vals)


def pct(x):
    return f"{x * 100:.1f}％"


def kakikae(p, s):
    if p.runs:
        p.runs[0].text = s
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(s)


def main():
    futeki = []

    def judge(ok, mi, setsu=""):
        print(("   ○ " if ok else "   × ") + mi + (("　" + setsu)
                                                   if setsu else ""))
        if not ok:
            futeki.append(mi)

    print("協議用素案 Ver.2.27 → Ver.2.28（見える化システムの事業所分析の反映）")
    print()
    for f in (MOTO, CSV_IN):
        if not os.path.exists(f):
            print(f"   × 入力がありません：{f}")
            sys.exit(1)
    rows, _ = yomu()

    # ── 値の算出
    def kado(shu, nen):
        n = goukei(rows, "延べ利用者数（人日／月）", shu, nen)
        t = goukei(rows, "定員（人）", shu, nen, "施設別")
        return n / t / HI, t

    def kaiten(shu, nen):
        n = goukei(rows, "延べ利用者数（人回／月）", shu, nen)
        t = goukei(rows, "定員（人）", shu, nen, "事業所別回転")
        return n / t / HI, t

    K = {}
    for shu in SHISETSU:
        K[shu] = {nen: kado(shu, nen) for nen in NEN}
    T = {shu: {nen: kaiten(shu, nen) for nen in NEN} for shu in ("15", "16")}
    TAN = {nen: (goukei(rows, "利用者数（実数・人／月）", "21", nen) or 0)
           + (goukei(rows, "利用者数（実数・人／月）", "22", nen) or 0) for nen in NEN}
    KAI = {nen: sum(goukei(rows, "介護職員（常勤換算）", shu, nen,
                           "施設別従事") or 0 for shu in SHISETSU) for nen in NEN}
    TEI = {nen: sum(K[shu][nen][1] for shu in SHISETSU) for nen in NEN}
    C_KADO = {shu: ken_chuo("定員充足率", shu) for shu in SHISETSU}
    C_KAI = {shu: statistics.median([1]) if False else ken_chuo("回転率", shu)
             for shu in ("15", "16")}
    r3, r5, r7 = "令和3年", "令和5年", "令和7年"
    per100 = {nen: KAI[nen] / TEI[nen] * 100 for nen in NEN}
    mitsu7 = K["54"][r7][0]
    aki = K["54"][r7][1] * (1 - mitsu7)

    print("  ── ⓪ 値（CSV から算出）")
    judge(K["52"][r7][1] == 86 and K["51"][r7][1] == 30 and K["54"][r7][1] == 30
          and K["32"][r7][1] == 27, "定員 老健86・特養30・地域密着型30・GH27（既知の値と一致）")
    judge(K["54"][r3][1] == 20, f"地域密着型の定員 令和3年 {K['54'][r3][1]:.0f}→令和7年 30（令和5年に増床）")
    print(f"     地域密着型 稼働率 {pct(K['54'][r3][0])}→{pct(mitsu7)}・空き約{aki:.1f}床")
    print(f"     介護職員（施設・居住系計）{KAI[r3]:.1f}→{KAI[r7]:.1f}人・"
          f"定員100人あたり {per100[r3]:.1f}→{per100[r7]:.1f}人")

    # ── 文
    MIDASHI = "　▌ 5年間の推移（見える化システムの事業所分析）"
    HON = (f"{SHUTTEN}により、町内の指定ごとの利用者数・延べ利用者数・定員・従事者数の"
           f"令和3年から令和7年までの推移を把握しました。延べ利用日数でみた稼働率は、"
           f"令和7年で介護老人福祉施設{pct(K['51'][r7][0])}、介護老人保健施設"
           f"{pct(K['52'][r7][0])}と高い水準にあります。一方、地域密着型介護老人福祉施設"
           f"入所者生活介護は、令和5年に定員を20人から30人に増やした後も稼働率が"
           f"{pct(K['54'][r5][0])}（令和5年）、{pct(mitsu7)}（令和7年）にとどまり、"
           f"宮城県内の同じサービスの中央値{pct(C_KADO['54'])}を下回っています。"
           f"増やした定員のうち約{round(aki):.0f}人分が埋まっていないことになります。"
           f"短期入所（生活介護・療養介護の2事業所）の利用者は、令和3年の月{TAN[r3]:.0f}人から"
           f"令和7年の月{TAN[r7]:.0f}人へほぼ半減しました。施設・居住系の介護職員"
           f"（常勤換算・指定の重複を除く計）は{KAI[r3]:.1f}人から{KAI[r7]:.1f}人へ減り、"
           f"定員100人あたりでは{per100[r3]:.1f}人から{per100[r7]:.1f}人へ"
           f"{(per100[r7] / per100[r3] - 1) * 100:.1f}％減っています。")
    CHUI = ("⚠ 各年の値は9月分です（利用者数は介護保険総合データベースの9月サービス分、"
            "定員・従事者数は介護サービス情報公表システムの9月時点の情報）。稼働率は"
            "「1月あたり延べ利用者数÷（定員×30日）」、通所の回転率も同じ式によります"
            "（厚生労働省「地域包括ケア『見える化』システム等を活用した地域分析の手引き」"
            "令和8年9月改訂版）。予防給付の事業所は含みません。宮城県「老人福祉施設等"
            "入所状況調」の入所率（時点の在所者数）とは定義が違います。⚠ 地域密着型の"
            "増やした定員が埋まっていないことが、職員不足による受入れの制限によるものかは"
            "確かめていません［要確認］。")
    HYO = [
        ["区　分", "令和3年", "令和5年", "令和7年", "宮城県内の中央値（令和7年）"],
        ["介護老人福祉施設（定員30人）の稼働率"] + [pct(K["51"][n][0]) for n in (r3, r5, r7)]
        + [pct(C_KADO["51"])],
        ["介護老人保健施設（定員86人）の稼働率"] + [pct(K["52"][n][0]) for n in (r3, r5, r7)]
        + [pct(C_KADO["52"])],
        [f"地域密着型介護老人福祉施設入所者生活介護（2施設・定員{K['54'][r3][1]:.0f}→"
         f"{K['54'][r7][1]:.0f}人）の稼働率"] + [pct(K["54"][n][0]) for n in (r3, r5, r7)]
        + [pct(C_KADO["54"])],
        ["認知症対応型共同生活介護（2事業所・定員27人）の稼働率"]
        + [pct(K["32"][n][0]) for n in (r3, r5, r7)] + [pct(C_KADO["32"])],
        ["短期入所（2事業所）の利用者数（月）"] + [f"{TAN[n]:.0f}人" for n in (r3, r5, r7)]
        + ["―"],
        [f"通所介護（定員{T['15'][r3][1]:.0f}→{T['15'][r7][1]:.0f}人）の回転率"]
        + [pct(T["15"][n][0]) for n in (r3, r5, r7)] + [pct(C_KAI["15"])],
        [f"通所リハビリテーション（定員{T['16'][r7][1]:.0f}人）の回転率"]
        + [pct(T["16"][n][0]) for n in (r3, r5, r7)] + [pct(C_KAI["16"])],
        ["施設・居住系の介護職員（常勤換算・計）"] + [f"{KAI[n]:.1f}人" for n in (r3, r5, r7)]
        + ["―"],
        ["同 定員100人あたり"] + [f"{per100[n]:.1f}人" for n in (r3, r5, r7)] + ["―"],
    ]
    SHUTTEN_GYO = f"出典：{SHUTTEN}"

    shutil.copy2(MOTO, SAKI)
    d = docx.Document(SAKI)
    mae = "\n".join([p.text for p in d.paragraphs]
                    + [c.text for t in d.tables for r in t.rows for c in r.cells])
    P = d.paragraphs

    def p_hit(atama):
        h = [p for p in d.paragraphs if p.text.strip().startswith(atama)]
        return h[0] if len(h) == 1 else None

    print("  ── ① 2-7 データの出どころ")
    p = p_hit("本節に掲載する供給側のデータは次のとおりです。")
    KYU1 = "いずれにも項目のないもの（入所前の居所・経営の状況・短期入所の稼働）は、"
    ok = p is not None and KYU1 in p.text
    judge(ok, "p210 の旧の文がある")
    if ok:
        kakikae(p, p.text.replace(
            KYU1, f"また、{SHUTTEN}により、延べ利用日数による稼働率、短期入所・通所の"
                  "利用の推移、職員数の推移を令和3年から令和7年まで把握しました。"
                  "いずれにも項目のないもの（入所前の居所・経営の状況）は、"))

    print("  ── ② 表26（把握の状況）")
    t26 = [t for t in d.tables if t.rows[0].cells[0].text.strip() == "区分"
           and t.rows[0].cells[1].text.strip() == "把握する項目"]
    judge(len(t26) == 1, "表26 を特定")
    if t26:
        c = t26[0].rows[2].cells[2]
        KYU2 = "⚠ 延べ利用日数による稼働率ではありません。⚠ 短期入所の稼働はいずれの資料にもなく、把握していません"
        ok = KYU2 in c.text
        judge(ok, "表26 定員と稼働の旧の文がある")
        if ok:
            kakikae(c.paragraphs[0], c.text.replace(
                KYU2, f"⚠ これは時点の在所者数による入所率です。延べ利用日数による稼働率と"
                      f"短期入所の利用者数は、{SHUTTEN}による（本節の「5年間の推移」）"))
        c = t26[0].rows[3].cells[2]
        KYU3 = "また、町内は8事業所・3法人であるため、事業所ごとの職員数は本計画には掲げません"
        ok = KYU3 in c.text
        judge(ok, "表26 職員数の旧の文がある")
        if ok:
            kakikae(c.paragraphs[0], c.text.replace(
                KYU3, KYU3 + f"。施設・居住系の介護職員（常勤換算）の推移は、{SHUTTEN}により"
                             "指定の重複を除いた合計で示します（本節の「5年間の推移」）"))

    print("  ── ③ 2-7 に「5年間の推移」を置く（供給制約を示す事象の前）")
    mae_p = p_hit("　▌ 供給制約を示す事象") or p_hit("▌ 供給制約を示す事象")
    judge(mae_p is not None, "「供給制約を示す事象」の小見出しを特定")
    hon_tpl = p_hit("第2に、町内の8の事業所を運営しているのは3法人です。")
    mid_tpl = mae_p
    sh_tpl = p_hit("出典：第2章2-7　サービス供給体制の状況。")
    t_tpl = t26[0] if t26 else None
    if mae_p is not None and hon_tpl is not None and t_tpl is not None and sh_tpl is not None:
        anchor = mae_p._p
        els = []
        m = copy.deepcopy(mid_tpl._p)
        els.append((m, MIDASHI))
        h = copy.deepcopy(hon_tpl._p)
        els.append((h, HON))
        tb = copy.deepcopy(t_tpl._tbl)
        els.append((tb, HYO))
        sh = copy.deepcopy(sh_tpl._p)
        els.append((sh, SHUTTEN_GYO))
        ch = copy.deepcopy(hon_tpl._p)
        els.append((ch, CHUI))
        import docx.table as dt
        import docx.text.paragraph as dp
        for el, naiyo in els:
            anchor.addprevious(el)
            if isinstance(naiyo, str):
                kakikae(dp.Paragraph(el, hon_tpl._parent), naiyo)
            else:
                t = dt.Table(el, t_tpl._parent)
                while len(t.rows) > len(naiyo):
                    el.remove(t.rows[-1]._tr)
                while len(t.rows) < len(naiyo):
                    el.append(copy.deepcopy(t.rows[-1]._tr))
                # 列を5つに（表26は4列）。最後の列を複製して1列足す
                grid = el.find(qn("w:tblGrid"))
                cols = grid.findall(qn("w:gridCol"))
                if len(cols) == 4:
                    zentai = sum(int(c.get(qn("w:w"))) for c in cols)
                    nc = copy.deepcopy(cols[-1])
                    grid.append(nc)
                    for tr in el.findall(qn("w:tr")):
                        tr.append(copy.deepcopy(tr.findall(qn("w:tc"))[-1]))
                    haba = [int(zentai * x / 170) for x in (66, 24, 24, 24, 32)]
                    haba[-1] += zentai - sum(haba)
                    for c, w in zip(grid.findall(qn("w:gridCol")), haba):
                        c.set(qn("w:w"), str(w))
                    for tr in el.findall(qn("w:tr")):
                        for tc, w in zip(tr.findall(qn("w:tc")), haba):
                            tc.find(qn("w:tcPr")).find(qn("w:tcW")).set(qn("w:w"), str(w))
                t = dt.Table(el, t_tpl._parent)
                for ri, gyo in enumerate(naiyo):
                    for ci, v in enumerate(gyo):
                        cell = t.rows[ri].cells[ci]
                        kakikae(cell.paragraphs[0], v)
                        for q in cell.paragraphs[1:]:
                            for r in q.runs:
                                r.text = ""
        judge(True, "小見出し・本文・表・出典・注記を置いた")

    print("  ── ④ 別表（表4）事項八")
    KYU4 = ("⚠ 延べ利用日数による稼働率は把握しておらず、時点の入所率で代える。"
            "⚠ 短期入所の稼働はいずれの資料にもない（確認事項No.159）")
    hit4 = [c for t in d.tables for r in t.rows for c in r.cells if KYU4 in c.text]
    judge(len(hit4) >= 1, "表4 の旧の文がある")
    for c in {id(x._tc): x for x in hit4}.values():
        kakikae(c.paragraphs[0], c.text.replace(
            KYU4, f"延べ利用日数による稼働率と短期入所の利用者数・職員数の推移は"
                  f"{SHUTTEN}により令和3年から令和7年まで把握した（2-7）"))

    print("  ── ⑤ 7-1 提供体制の概況")
    p = p_hit("この提供体制は、厚生労働省「介護サービス情報公表システム」")
    KYU5 = "⚠ 入所率は「在所者数÷定員」であり、延べ利用日数による稼働率ではありません。"
    ok = p is not None and KYU5 in p.text
    judge(ok, "p467 の旧の文がある")
    if ok:
        kakikae(p, p.text.replace(
            KYU5, KYU5 + f"延べ利用日数による稼働率（{SHUTTEN}）は、令和7年で介護老人福祉施設"
                         f"{pct(K['51'][r7][0])}、介護老人保健施設{pct(K['52'][r7][0])}、"
                         f"地域密着型介護老人福祉施設入所者生活介護{pct(mitsu7)}、"
                         f"認知症対応型共同生活介護{pct(K['32'][r7][0])}です（2-7）。"))

    print("  ── ⑥ 7-2 給付費の平坦化")
    p = p_hit("標準給付費は令和6年度の1,093,253千円から令和7年度の1,084,587千円へ")
    KYU6 = "を踏まえると、供給の制約による可能性が高いと考えられます。"
    ok = p is not None and KYU6 in p.text
    judge(ok, "p471 の旧の文がある")
    if ok:
        kakikae(p, p.text.replace(
            KYU6, f"、及び施設・居住系の介護職員（常勤換算）が令和3年の{KAI[r3]:.1f}人から"
                  f"令和7年の{KAI[r7]:.1f}人へ減り、定員100人あたりで{per100[r3]:.1f}人から"
                  f"{per100[r7]:.1f}人になっていること（{SHUTTEN}・2-7）{KYU6}"))

    print("  ── ⑦ 7-6 定点把握")
    p = p_hit("①　年1回の公表情報の定点把握")
    KYU7 = "（職種別の従業者数と常勤換算、前年度の採用者数及び退職者数、入所率、待機者数）。"
    ok = p is not None and KYU7 in p.text
    judge(ok, "p495 の旧の文がある")
    if ok:
        kakikae(p, p.text.replace(
            KYU7, KYU7 + "あわせて、地域包括ケア「見える化」システムの事業所分析により、"
                         "稼働率・回転率・利用者数・従事者数の5年推移と「5年間連続減少」の"
                         "該当を毎年度確かめます。"))

    print("  ── ⑧ 9-1 表102 と地域密着型の空き床")
    t102 = [t for t in d.tables if t.rows[0].cells[0].text.strip() == "判断項目"]
    judge(len(t102) == 1, "表102 を特定")
    if t102:
        c = t102[0].rows[1].cells[1]
        KYU8 = "⚠ 延べ利用日数による稼働率・短期入所の稼働・入所前の居所は公表資料にない"
        ok = KYU8 in c.text
        judge(ok, "表102 既存資源の旧の文がある")
        if ok:
            kakikae(c.paragraphs[0], c.text.replace(
                KYU8, "延べ利用日数による稼働率と短期入所の利用者数（2-7）。⚠ 入所前の居所は"
                      "公表資料にない"))
        c = t102[0].rows[1].cells[3]
        if "見える化" not in c.text:
            kakikae(c.paragraphs[0], c.text + f"。{SHUTTEN}（稼働率・短期入所の利用者数）")
        c = t102[0].rows[3].cells[3]
        if "見える化" not in c.text:
            kakikae(c.paragraphs[0], c.text + f"。{SHUTTEN}（従事者数の5年推移）")
    p = p_hit("⚠ 見込量は、川崎町の被保険者が利用する量であり、町内の事業所の定員とは")
    judge(p is not None, "必要利用定員総数の注記を特定")
    if p is not None:
        n = copy.deepcopy(p._p)
        p._p.addnext(n)
        import docx.text.paragraph as dp
        kakikae(dp.Paragraph(n, p._parent),
                f"⚠ 町内の地域密着型介護老人福祉施設入所者生活介護は、令和5年に定員を"
                f"30人とした後も延べ利用日数による稼働率が{pct(mitsu7)}（令和7年）であり、"
                f"約{round(aki):.0f}人分の受入れの余地があります（2-7）。これが職員不足による"
                f"受入れの制限によるものかを確かめたうえで、第10期の見込量と必要利用定員総数"
                f"（月27.2人）を据え置くか、町内の空き床の活用を見込量に反映するかを定めます"
                f"［要協議］。")

    d.save(SAKI)
    d2 = docx.Document(SAKI)
    ato = "\n".join([p.text for p in d2.paragraphs]
                    + [c.text for t in d2.tables for r in t.rows for c in r.cells])

    print("  ── 検査1　「公表資料にない」の古い記述が残っていない")
    for go in ("短期入所の稼働はいずれの資料にもない", "延べ利用日数による稼働率は把握しておらず",
               "短期入所の稼働・入所前の居所は公表資料にない",
               "（入所前の居所・経営の状況・短期入所の稼働）"):
        judge(go not in ato, f"「{go[:22]}」0件")
    print("  ── 検査2　表と本文の値が同じ")
    judge(pct(mitsu7) in HON and HYO[3][3] == pct(mitsu7), f"地域密着型の稼働率 {pct(mitsu7)}")
    judge(ato.count(pct(mitsu7)) >= 4, f"地域密着型 {pct(mitsu7)} が2-7・7-1・9-1・表に出る")
    print("  ── 検査3　出典を明記")
    judge(ato.count(SHUTTEN) >= 6, f"出典「{SHUTTEN}」{ato.count(SHUTTEN)}か所")
    print("  ── 検査4　⚠ 事業所ごとの職員数を掲げていない")
    kaigo = {f"{v:.1f}人" for (k, s, b, n, sh), v in yomu()[1].items()
             if k == "介護職員（常勤換算）" and v and s in ("51", "52", "54", "32")}
    goukei_s = {f"{KAI[n]:.1f}人" for n in NEN}
    kobetsu = [x for x in kaigo - goukei_s if x in HON or any(x in g for g in sum(HYO, []))]
    judge(not kobetsu, "事業所ごとの介護職員数を本文・表に書いていない", "／".join(kobetsu))
    print("  ── 検査5　事業所名・法人名・他団体名")
    meisho = {r["事業所名"] for r in rows}
    # 社会福祉協議会は事業の実施主体として以前から本文にある。今回の反映で増えていないことを見る
    fueta = [m for m in meisho if m and ato.count(m) > mae.count(m)]
    judge(not fueta, "今回の反映で事業所名・法人名が増えていない", "／".join(fueta))
    judge(not any(w in ato for w in ("北塩原", "金ヶ崎", "大雪", "川崎市")), "他団体名 0件")
    print("  ── 検査6　未確定の印")
    judge(ato.count("［要確認］") == mae.count("［要確認］") + 1, "［要確認］+1（空き床の理由）")
    judge(ato.count("［要協議］") == mae.count("［要協議］") + 1, "［要協議］+1（見込量の扱い）")
    print("  ── 検査7　⚠ 既にあった数が消えていない")
    kesu = sorted({m.group(0) for m in re.finditer(
        r"[0-9０-９][0-9０-９,.]*(?:人|％|千円|件|名|歳|回|円|点)", mae)
        if mae.count(m.group(0)) > ato.count(m.group(0))})
    judge(not kesu, "消えた数が0件", "／".join(kesu[:6]))

    print()
    print(f"  {SAKI}")
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        for m in futeki:
            print(f"      {m}")
        sys.exit(1)
    print(f"   ○ 検査1〜7　不適合0件（［要確認］{ato.count('［要確認］')}件・"
          f"［要協議］{ato.count('［要協議］')}件）")


if __name__ == "__main__":
    main()
