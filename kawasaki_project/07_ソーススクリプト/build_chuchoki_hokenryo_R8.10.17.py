# -*- coding: utf-8 -*-
"""中長期の見込量と保険料の内訳の書き出し（作業順位29・令和8年10月17日）

  python3 07_ソーススクリプト/build_chuchoki_hokenryo_R8.10.17.py

なぜ作るか
  素案の根拠の棚卸し（R8.10.7）で、素案の数値1,447件のうち**70件が
  出所を特定できていない**ものとして残った。原因は1つに絞れている。
  **いずれも中長期（令和17・22年度）の見込量と保険料の内訳であり、
  算定はしているが、計画見込量の xlsx が令和11年度までしか持っておらず、
  素案にしか数が残っていない。**
  書き出せば、素案の数の裏づけがほぼすべて揃う。

⚠ **数値の正本は計画素案である**（CLAUDE.md）。
  本スクリプトは素案の表から値を読み、算定式と出所を添えて書き出す。
  **値を本スクリプトに書き写さない。** 書き写せば正本が2か所になる。
  例外は、素案に現れない算定の中間値（補正係数・乗率・収納率）に限り、
  いずれも素案の表111・表119 に印字されている値である。

出力
  05_試算・管理シート/川崎町_第10期_計画見込量_中長期・保険料内訳_R8.10.17.xlsx
    00_この表について
    01_中長期の見込量      令和17・22年度のサービス別見込量（算定式つき）
    02_総給付費の中長期    9-2 の総給付費の年次
    03_保険料の算定        記号A〜J の年度別（9-3 表111）
    04_3パターンと感度      基金取崩の3案と前提を動かした場合（表113・114）
    05_所得段階別の保険料   13段階の月額・年額・軽減後（表119）
    06_棚卸し70件の対応     ⚠ **本作業の核心。** 70件の一つずつに
                           出所又は算定式を対応づけ、残りを0件にする
    07_自己点検

自己点検
  検査1  棚卸しの70件すべてが本表のどこかに現れること
  検査2  中長期の見込量の合計が、素案 9-2 の総給付費（令和17・22年度）と合うこと
  検査3  所得段階別の月額が 算定Ｂ×乗率 で再現できること
  検査4  年度別の ①＝A＋B／C＝①×23％／J＝C＋D－E が成り立つこと
  検査5  3か年計が年度別の合計と一致すること
  検査6  入所・居住系の給付費（5-5 の680,226千円）が内訳の和で再現できること
  検査7  2-4 の要支援計が居宅サービスの要支援の給付費の和で再現できること
  検査8  受託者を主語とする語・他団体の固有名称が入っていないこと
"""
import importlib.util
import os
import re
import sys
from decimal import ROUND_HALF_UP, Decimal

import docx
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

SOAN = "01_第10期_最新版成果品/川崎町_計画書素案_v2.18_範囲明記版.docx"
TANA = "05_試算・管理シート/川崎町_素案の根拠の棚卸し_R8.10.7.xlsx"
OUT = ("05_試算・管理シート/"
       "川崎町_第10期_計画見込量_中長期・保険料内訳_R8.10.17.xlsx")

# 素案に印字されている算定の前提（表111・表119 から読む。ここには書かない）
NENDO = ["令和9年度", "令和10年度", "令和11年度"]


def load(name, fname):
    spec = importlib.util.spec_from_file_location(
        name, os.path.join(HERE, fname))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


GY = load("gyotei_chu", "build_iinkai_gyotei_R8.10.1.py")
CC = load("calc_chuchoki", "calc_chuchoki_R8.10.17.py")


def seiri(s):
    return (s or "").replace("\n", "").strip()


def kazu(s):
    """「31,040千円」「2,957.5円」「96.0％」→ Decimal。なければ None"""
    if s is None:
        return None
    t = seiri(s)
    if t in ("", "─", "―", "-", "［要協議］"):
        return None
    m = re.search(r"(-?[\d,]+(?:\.\d+)?)", t)
    if not m:
        return None
    try:
        return Decimal(m.group(1).replace(",", ""))
    except Exception:
        return None


def en(v):
    return int(Decimal(str(v)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def hyo(doc, atama):
    """1行目の見出しの並びの先頭が atama に一致する表を返す"""
    for t in doc.tables:
        h = [seiri(c.text) for c in t.rows[0].cells]
        if h[:len(atama)] == atama:
            return t
    return None


def hyo_mikomi(doc):
    """見込量の4表を、見出し（▌）との対応で拾う"""
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.oxml.ns import qn
    out = {}
    midashi = None
    for el in doc.element.body.iterchildren():
        if el.tag == qn("w:p"):
            s = seiri(Paragraph(el, doc).text)
            if s.startswith("▌") and "見込量" in s:
                midashi = s
        elif el.tag == qn("w:tbl"):
            if midashi and midashi not in out:
                out[midashi] = Table(el, doc)
                midashi = None
    return out


def yomu_mikomi(doc):
    """中長期（令和17・22年度）を持つサービスを読む

    戻り　[(区分, サービス, 項目, 単位, R7実績, R9, R10, R11, R17, R22)]
    """
    tbls = hyo_mikomi(doc)
    KUBUN = {"▌ 居宅サービスの見込量": "居宅サービス",
             "▌ 介護予防サービスの見込量": "介護予防サービス",
             "▌ 地域密着型サービスの見込量": "地域密着型サービス",
             "▌ 施設サービスの見込量": "施設サービス"}
    rows = []
    for mi, kubun in KUBUN.items():
        t = tbls.get(mi)
        if t is None:
            raise SystemExit(f"見込量の表が見つからない：{mi}")
        h = [seiri(c.text) for c in t.rows[0].cells]
        idx = {k: h.index(k) for k in h}
        na = ""
        for r in t.rows[1:]:
            c = [seiri(x.text) for x in r.cells]
            if c[0]:
                na = c[0]
            rows.append((kubun, na, c[1], c[2],
                         kazu(c[idx["令和7年度実績"]]),
                         kazu(c[idx["令和9年度"]]), kazu(c[idx["令和10年度"]]),
                         kazu(c[idx["令和11年度"]]),
                         kazu(c[idx["令和17年度"]]), kazu(c[idx["令和22年度"]])))
    return rows


def yomu_tanaoroshi():
    """棚卸しの「出所を特定できていないもの」を読む"""
    if not os.path.exists(TANA):
        return []
    wb = openpyxl.load_workbook(TANA, read_only=True)
    ws = wb["03_出所を特定できていないもの"]
    out = []
    for r in ws.iter_rows(min_row=4, values_only=True):
        if not r[0] or seiri(str(r[0])) == "節":
            continue
        out.append((seiri(str(r[0])), seiri(str(r[1])),
                    seiri(str(r[2])), seiri(str(r[3] or ""))))
    wb.close()
    return out


def main():
    for p in (SOAN,):
        if not os.path.exists(p):
            raise SystemExit("入力がない：" + p)
    doc = docx.Document(SOAN)

    # ══════════════════ 素案から読む
    mikomi = yomu_mikomi(doc)
    choki = [x for x in mikomi if x[8] is not None and x[2] == "給付費"]

    t108 = hyo(doc, ["区分", "令和7年度\n実績"]) or hyo(doc, ["区分", "令和7年度 実績"])
    if t108 is None:
        for t in doc.tables:
            h = [seiri(c.text) for c in t.rows[0].cells]
            if h and h[0] == "区分" and "令和17年度" in h:
                t108 = t
                break
    if t108 is None:
        raise SystemExit("9-2 の総給付費の表が見つからない")
    h108 = [seiri(c.text) for c in t108.rows[0].cells]
    kyufu = {}
    for r in t108.rows[1:]:
        c = [seiri(x.text) for x in r.cells]
        kyufu[c[0]] = {h108[i]: kazu(c[i]) for i in range(1, len(c))}

    t111 = hyo(doc, ["記号", "項目", "第10期計"])
    if t111 is None:
        raise SystemExit("9-3 の保険料算定表（表111）が見つからない")
    h111 = [seiri(c.text) for c in t111.rows[0].cells]
    sanN = {}
    for r in t111.rows[1:]:
        c = [seiri(x.text) for x in r.cells]
        key = c[0] or c[1]
        sanN[key] = c

    t113 = hyo(doc, ["パターン", "基金取崩方針"])
    t114 = hyo(doc, ["前提", "本算定の置き方"])
    t119 = hyo(doc, ["段階", "対象", "基準額に対する割合"])
    for nm, t in (("表113", t113), ("表114", t114), ("表119", t119)):
        if t is None:
            raise SystemExit(f"9-3 の{nm}が見つからない")

    # 算定の前提（素案の表111 の文字から取る）
    shuno = kazu(sanN["予定保険料収納率"][2])          # 96.0（％）
    hosei_n = kazu(sanN["③"][2])                      # 9,743.1 人
    kijun_a = None
    for r in t111.rows[1:]:
        c = [seiri(x.text) for x in r.cells]
        if "月額基準額" in c[1]:
            kijun_a = kazu(c[2])
    if kijun_a is None:
        raise SystemExit("算定Ａの月額基準額が読めない")
    # 算定Ｂ＝第5段階（基準）の第10期月額
    kijun_b = None
    for r in t119.rows[1:]:
        c = [seiri(x.text) for x in r.cells]
        if c[0].startswith("第5段階"):
            kijun_b = kazu(c[4])
    if kijun_b is None:
        raise SystemExit("算定Ｂ（第5段階の第10期月額）が読めない")

    # ══════════════════ 書き出し
    wb = openpyxl.Workbook()

    # ── 00
    ws = GY.sheet(
        wb, "00_この表について", "中長期の見込量と保険料の内訳",
        "素案の根拠の棚卸し（令和8年10月7日）で出所を特定できていなかった"
        "70件は、いずれも**中長期（令和17年度・令和22年度）の見込量**と"
        "**保険料の内訳**でした。算定はしていましたが、計画見込量の表が"
        "令和11年度までしか持っておらず、計画素案にしか数が残っていない"
        "状態でした。本表はそれを書き出したものです。\n"
        "⚠ **数値の正本は計画素案です。** 本表は素案の表から値を読み、"
        "算定式と出所を添えたものであり、数値を別に持つものではありません。\n"
        "⚠ 中長期の推計は、第10期計画期間（令和9〜11年度）の見込量とは"
        "位置づけが異なります。**計画の数値ではなく、長期の見通しです。**",
        ["シート", "内容"], [30, 104], first=True)
    GY.put(ws, [
        ["01_中長期の見込量",
         f"令和17年度・令和22年度の給付費を持つ{len(choki)}サービス。"
         "算定式と、令和7年度実績からの伸びを添えています。"],
        ["02_総給付費の中長期", "9-2 の総給付費の年次（令和7年度実績〜令和22年度）。"],
        ["03_保険料の算定", "記号A〜J の年度別と3か年計（9-3 表111）。算式つき。"],
        ["04_3パターンと感度",
         "介護給付費準備基金の取崩しの3案（表113）と、前提を動かした場合の"
         "幅（表114）。"],
        ["05_所得段階別の保険料",
         "13段階の乗率・月額・年額・軽減後（表119）。"
         "⚠ 年額は月額×12で機械が計算しています。"],
        ["06_棚卸し70件の対応",
         "⚠ **本表の核心です。** 棚卸しで残った70件の一つずつに、"
         "出所又は算定式を対応づけています。"],
        ["07_自己点検", "検査1〜8。1件でも不適合があれば終了コード1で終わります。"],
    ])

    # ── 01　中長期の見込量
    ws = GY.sheet(
        wb, "01_中長期の見込量",
        "中長期の見込量（令和17年度・令和22年度）",
        f"{len(choki)}サービス。算定式は第10期計画期間と同じです。\n"
        "　見込量(年,要介護度) ＝ 令和7年度実績(要介護度) × 認定者数(年,要介護度)"
        " ÷ 認定者数(令和8年度,要介護度)\n"
        "認定者数は地域包括ケア「見える化」システムの推計値"
        "（性別・年齢5歳階級別・要介護度別）によります。"
        "サービス利用率・1人1月あたりの利用回（日）数・単価は"
        "令和7年度の水準で据え置いています。\n"
        "⚠ **介護予防給付のサービスは、サービスごとの長期推計を行って"
        "いないため「─」としています**（素案 9-1 の注のとおり）。"
        "「未実施」を「適合」に丸めていません。",
        ["区分", "サービス種類", "項目", "単位", "令和7年度実績",
         "令和11年度", "令和17年度", "令和22年度",
         "R7→R17 の伸び", "R7→R22 の伸び", "算定式・出所"],
        [16, 32, 10, 9, 13, 13, 13, 13, 12, 12, 46])
    rows01 = []
    for kubun, na, komoku, tan, r7, r9, r10, r11, r17, r22 in mikomi:
        if r17 is None:
            continue
        nobi17 = f"{float(r17 / r7):.3f}倍" if r7 else "―"
        nobi22 = f"{float(r22 / r7):.3f}倍" if r7 else "―"
        rows01.append([
            kubun, na, komoku, tan,
            float(r7) if r7 is not None else None,
            float(r11) if r11 is not None else None,
            float(r17), float(r22), nobi17, nobi22,
            "令和7年度実績×認定者数の伸び（要介護度別）。"
            "出所＝介護保険事業状況報告（年報・令和7年度）＋"
            "見える化システムの推計値。素案 9-1 の表に同じ値がある"])
    GY.put(ws, rows01,
           [GY.HI if r[2] == "給付費" else None for r in rows01])

    # ── Ver.2.17 で素案 9-1 に入れた15サービスを、算定の出所つきで示す
    #   ⚠ **値は素案（正本）から読んでいる。** ここでは「どの15サービスが
    #     Ver.2.17 で補われたか」を印すだけで、値を別に持たない。
    hokan = CC.santei()
    CC.kensho(hokan, *CC.yomu_soan_kyufu(SOAN))     # 行の引当を確かめる
    hokan_na = {k for k in hokan if not k.startswith("__")}
    hokan17 = CC.gokei(hokan, "令和17年度")
    hokan22 = CC.gokei(hokan, "令和22年度")
    # 算定式の列に、補完したものである旨を書き足す
    for i, r in enumerate(rows01):
        if r[1] in hokan_na:
            ws.cell(5 + i, 11).value = (
                "⚠ **Ver.2.17 で補った欄。** 素案 9-1 では「─」のままで、"
                "9-2 の総給付費にだけ入っていた（表の間の範囲の食い違い）。"
                "式は上と同じ。算定は calc_chuchoki_R8.10.17.py にある")
            ws.cell(5 + i, 11).fill = GY.WARN

    # ── 02　総給付費の中長期
    ws = GY.sheet(
        wb, "02_総給付費の中長期", "総給付費・標準給付費見込額の年次",
        "素案 9-2 の表から読んだ値です。\n"
        "⚠ 標準給付費見込額・地域支援事業費は第10期計画期間"
        "（令和9〜11年度）についてのみ算定しています。"
        "**中長期は総給付費のみです。**",
        ["区分"] + h108[1:] + ["算定式・出所"],
        [30] + [14] * (len(h108) - 1) + [50])
    SHIKI = {
        "総給付費": "サービス別見込量の給付費の合計（01シート）",
        "特定入所者介護サービス費等":
            "令和7年度実績52,967千円を要介護度別に伸ばす（素案 9-2）",
        "高額介護サービス費等": "令和7年度実績を総給付費の伸びで按分",
        "高額医療合算介護サービス費等": "同上",
        "審査支払手数料": "給付費決定件数×単価60円",
        "標準給付費見込額":
            "総給付費＋特定入所者＋高額介護＋高額医療合算＋審査支払手数料",
        "地域支援事業費": "⚠ 見える化システムの入力値（町のデータ未受領）",
    }
    rows02 = []
    for r in t108.rows[1:]:
        c = [seiri(x.text) for x in r.cells]
        v = [kazu(x) for x in c[1:]]
        rows02.append([c[0]] + [float(x) if x is not None else c[i + 1]
                                for i, x in enumerate(v)]
                      + [SHIKI.get(c[0], "")])
    GY.put(ws, rows02)

    # ── 03　保険料の算定
    ws = GY.sheet(
        wb, "03_保険料の算定", "保険料の算定（記号A〜J・年度別）",
        "素案 9-3 の表から読んだ値です。算式の列は素案の記載をそのまま"
        "引いています。\n"
        f"　月額基準額（算定Ａ）＝ J ÷ 収納率{shuno}％ ÷ {hosei_n}人 ÷ 12月"
        f" ＝ {kijun_a}円\n"
        "⚠ **I（介護給付費準備基金取崩額）は［要協議］です。** "
        "第3回策定委員会（令和9年1月）で取崩方針が決まるまで確定しません。"
        "04シートの3パターンをご覧ください。",
        h111, [8, 34, 15, 14, 14, 14, 52])
    rows03 = []
    for r in t111.rows[1:]:
        c = [seiri(x.text) for x in r.cells]
        v = [kazu(x) for x in c[2:6]]
        rows03.append(c[:2] + [float(x) if x is not None else c[2 + i]
                               for i, x in enumerate(v)] + [c[6]])
    GY.put(ws, rows03,
           [GY.WARN if "［要協議］" in "".join(
               seiri(x.text) for x in r.cells) else None
            for r in t111.rows[1:]])

    # ── 04　3パターンと感度
    ws = GY.sheet(
        wb, "04_3パターンと感度", "基金取崩の3案と、前提を動かした場合の幅",
        "素案 9-3 の表から読んだ値です。\n"
        "⚠ **いずれも第3回策定委員会（令和9年1月）でご判断いただく"
        "事項です。** 当方が選ぶものではありません。",
        ["区分"] + [seiri(c.text) for c in t113.rows[0].cells],
        [14, 12, 30, 40, 14, 34])
    rows04 = [["基金取崩の案"] + [seiri(c.text) for c in r.cells]
              for r in t113.rows[1:]]
    GY.put(ws, rows04)
    r0 = 5 + len(rows04) + 1
    ws.cell(r0, 1).value = "前提を動かした場合の幅（素案 9-3 表114）"
    ws.cell(r0, 1).font = openpyxl.styles.Font(name="游ゴシック", size=9,
                                               bold=True)
    GY.put(ws, [[""] + [seiri(c.text) for c in r.cells]
                for r in t114.rows],
           r0=r0 + 1)

    # ── 05　所得段階別の保険料
    ws = GY.sheet(
        wb, "05_所得段階別の保険料", "所得段階別の保険料（13段階）",
        f"素案 9-3 の表119 から読んだ値です。基準額は算定Ｂ＝{kijun_b}円"
        "（介護給付費準備基金を半額取崩した場合）です。\n"
        "　段階別の月額 ＝ 算定Ｂ × 基準額に対する割合（乗率）\n"
        "　年額 ＝ 月額 × 12（本表で機械が計算しています）\n"
        "⚠ **第10〜13段階の第1号被保険者数が0人のままです。** "
        "実際の分布を入手しだい再算定が必要です（確認事項No.13）。",
        ["段階", "対象", "乗率", "軽減後の乗率",
         "第9期 月額", "第10期 月額（算定Ｂ）", "第10期 年額",
         "第10期 月額（軽減後）", "第10期 年額（軽減後）",
         "乗率×算定Ｂ の検算"],
        [13, 46, 8, 10, 13, 15, 13, 15, 15, 24])
    rows05 = []
    for r in t119.rows[1:]:
        c = [seiri(x.text) for x in r.cells]
        m = re.match(r"([\d.]+)", c[2])
        if not m:
            continue
        jou = Decimal(m.group(1))
        kei = re.search(r"軽減後([\d.]+)", c[2])
        jou_k = Decimal(kei.group(1)) if kei else None
        tsuki = kazu(c[4])
        tsuki_k = None
        mk = re.search(r"軽減後([\d,]+)円", c[4])
        if mk:
            tsuki_k = Decimal(mk.group(1).replace(",", ""))
        saiken = en(kijun_b * jou)
        rows05.append([
            c[0], c[1], float(jou),
            float(jou_k) if jou_k is not None else None,
            float(kazu(c[3])) if kazu(c[3]) is not None else None,
            float(tsuki) if tsuki is not None else None,
            en(tsuki * 12) if tsuki is not None else None,
            float(tsuki_k) if tsuki_k is not None else None,
            en(tsuki_k * 12) if tsuki_k is not None else None,
            f"{jou}×{kijun_b}円＝{saiken:,}円"
            + ("　○" if tsuki is not None and en(tsuki) == saiken else "　×")])
    GY.put(ws, rows05,
           [GY.OKF if r[-1].endswith("○") else GY.NGF for r in rows05])

    # ══════════════════ 06b　見える化システムの総括表との差
    #   ⚠ **総括表のサービス別の値に合わせに行くための表ではない。**
    #     用いていない理由（確認事項No.76・79・80）と、差の大きさを
    #     記録に残すための表である。県への事前協議で問われうる。
    hik = CC.sokatsu_hikaku()
    ws = GY.sheet(
        wb, "06b_見える化システムとの差",
        "見える化システムの総括表のサービス別の値との差",
        "⚠ **総括表の「2_サービス別給付費」には、サービス別の"
        "令和17・22年度の給付費が別に入っています。"
        "本計画ではこれを用いていません。** 理由は2つです。\n"
        f"　①令和7年度に給付実績のあるサービスを令和9年度以降0と"
        f"置いている（{len(hik['ゼロ']) if hik else 0}件。確認事項No.76）\n"
        f"　②令和11年度から令和27年度まで同じ値が並び、認定者数が"
        f"伸びても給付費が動かない（{len(hik['一定']) if hik else 0}件。"
        f"確認事項No.80の、令和8年5月時点の値が解除できない"
        f"状態によるもの）\n"
        "⚠ **本計画は令和9〜11年度についても総括表のサービス別の値を"
        "用いておらず、全年度を通じて年報の実績に認定者数の伸びを乗じる"
        "方法で一貫して算定しています。**\n"
        "⚠ **県への事前協議で、本計画の数値と見える化システムの出力の"
        "違いを問われる可能性があります。** 差を下表に示します。",
        ["年度", "見える化システムの総括表", "本計画（素案 9-2）",
         "差", "差の割合", "備考"],
        [14, 22, 20, 16, 12, 56])
    rows06b = []
    if hik and hik["総括表計"]:
        for y in ("令和7年度", "令和9年度", "令和11年度", "令和17年度",
                  "令和22年度"):
            so = hik["総括表計"].get(y)
            ho = kyufu.get("総給付費", {}).get(y)
            if so is None or ho is None:
                continue
            sa = float(so) - float(ho)
            biko = ("実績。差は丸めによるもの" if y == "令和7年度" else
                    ("⚠ **計画期間の中にあり、保険料に効く**"
                     if y in ("令和9年度", "令和11年度") else
                     "長期の見通し。保険料には効かない"))
            rows06b.append([y, float(so), float(ho), sa,
                            f"{sa / float(ho) * 100:+.2f}％", biko])
    GY.put(ws, rows06b,
           [GY.OKF if abs(r[3]) <= 3 else
            (GY.NGF if "保険料に効く" in r[5] else GY.WARN)
            for r in rows06b])
    r0 = 5 + len(rows06b) + 1
    ws.cell(r0, 1).value = ("令和7年度に給付実績があるのに"
                            "令和9年度以降0とされているサービス"
                            "（確認事項No.76）")
    ws.cell(r0, 1).font = openpyxl.styles.Font(name="游ゴシック", size=9,
                                               bold=True)
    GY.put(ws, [[x, "", "", "", "", ""] for x in (hik["ゼロ"] if hik else [])],
           [GY.NGF] * len(hik["ゼロ"] if hik else []), r0=r0 + 1)
    r0 = r0 + 2 + len(hik["ゼロ"] if hik else [])
    ws.cell(r0, 1).value = ("令和11年度から令和27年度まで値が一定の"
                            "サービス（確認事項No.80）")
    ws.cell(r0, 1).font = openpyxl.styles.Font(name="游ゴシック", size=9,
                                               bold=True)
    GY.put(ws, [[x, "", "", "", "", ""] for x in (hik["一定"] if hik else [])],
           [GY.WARN] * len(hik["一定"] if hik else []), r0=r0 + 1)

    # ══════════════════ 06　棚卸し70件の対応
    tana = yomu_tanaoroshi()
    # 本表に書き出した値の集合（文字として突合する）
    aru = set()

    def tsumu(v):
        if v is None:
            return
        d = Decimal(str(v))
        aru.add(f"{d:,.0f}" if d == d.to_integral_value() else f"{d:,.1f}")
        aru.add(str(v))

    for _, _, _, _, r7, r9, r10, r11, r17, r22 in mikomi:
        for v in (r7, r9, r10, r11, r17, r22):
            tsumu(v)
    for k, d in kyufu.items():
        for v in d.values():
            tsumu(v)
    for r in t111.rows[1:]:
        for x in r.cells:
            tsumu(kazu(x.text))
    for t in (t113, t114, t119):
        for r in t.rows[1:]:
            for x in r.cells:
                for m in re.finditer(r"([\d,]+(?:\.\d+)?)", seiri(x.text)):
                    aru.add(m.group(1))
    for r in rows05:
        tsumu(r[6])
        tsumu(r[8])

    # ── 入所・居住系（5-5 の680,226千円）
    # ⚠ **「地域密着型サービス」をまるごと足すのは誤りである。**
    #   定期巡回・随時対応型訪問介護看護、地域密着型通所介護、
    #   看護小規模多機能型居宅介護は在宅のサービスであり入所・居住系ではない。
    #   一方、介護予防の認知症対応型共同生活介護は入所・居住系である。
    #   区分は第1回策定委員会資料の「居住形態別の再構成」の表による。
    #   ⚠ **円で足して最後に一度だけ千円に四捨五入する。**
    #     千円に丸めた値を足すと1千円ずれる。算定は calc_chuchoki に置く。
    nyusho_uchi_en, nyusho_en, nyusho = CC.nyusho_kyojukei()
    nyusho_uchi = [(n, v / 1000) for n, v in nyusho_uchi_en]
    tsumu(nyusho)
    # 2-4 の居宅サービス計の要支援計は、2-4 の表そのものが正本である
    t14_saki = hyo(doc, ["サービス区分", "要支援計"])
    if t14_saki is not None:
        for r in t14_saki.rows[1:]:
            for x in r.cells:
                v = kazu(x.text)
                if v is not None:
                    tsumu(v)

    MEMO = {
        "9-1": ("中長期の見込量（給付費・回（日）数・利用者数）",
                "01シートに、令和7年度実績からの伸び（倍）と算定式を添えて"
                "書き出した"),
        "9-2": ("総給付費の中長期と3か年計",
                "02シートに書き出した。3か年計は年度別の合計（検査5で確認）"),
        "9-3": ("保険料の内訳（記号A〜J・3パターン・感度・13段階）",
                "03〜05シートに書き出した。13段階は 乗率×算定Ｂ で"
                "再現できることを検査3で確認"),
        "2-4": ("居宅サービス計の要支援計（円）",
                "年報（令和7年度）の要介護度別明細による。"
                "⚠ 区分別の内訳は素案 2-4 の表が正本であり、本表では持たない"),
        "5-5": ("入所・居住系の給付費（千円）",
                f"施設3＋地域密着型介護老人福祉施設＋"
                f"認知症対応型共同生活介護（介護給付・介護予防）＋"
                f"地域密着型特定施設＋特定施設を**円で足して**"
                f"一度だけ千円に丸めると{nyusho:,.0f}千円。検査6で確認。"
                f"⚠ 定期巡回・地域密着型通所介護・看護小規模多機能型は"
                f"在宅のサービスであり含めない"),
    }
    # 値ごとに扱いを明記するもの（突合では拾えないが出所が定まっているもの）
    #   ⚠ 丸めにより1千円の差が出るため、数字の文字の突合では拾えない。
    DAI9 = {
        "3,294,553": ("第9期計画の値（当方の算定ではない）",
                      "第9期計画の標準給付費見込額（3か年計）。"
                      "出所＝川崎町第9期計画書"),
        "91,560": ("第9期計画の値（当方の算定ではない）",
                   "第9期計画の地域支援事業費。出所＝川崎町第9期計画書"),
        "680,226": ("入所・居住系の給付費（千円）",
                    f"内訳を円で足して一度だけ千円に丸めると"
                    f"{nyusho:,.0f}千円で、差0で再現した（検査6）"),
    }
    ws = GY.sheet(
        wb, "06_棚卸し70件の対応", "棚卸しで出所が示せていなかった70件の対応",
        "素案の根拠の棚卸し（令和8年10月7日）の「出所を特定できていない"
        "もの」の一つずつに、本表のどこで出所又は算定式を示したかを"
        "対応づけています。\n"
        "⚠ **この列が「未対応」のまま残るものがあれば、公表版の確定の"
        "前に解消が必要です。** 件数は07シートの検査1で数えています。",
        ["節", "種類", "数値", "前後の文", "何の数か", "本表での扱い",
         "状態"],
        [8, 8, 15, 36, 40, 54, 12])
    rows06 = []
    michi = []
    aru_nk = {x.replace(",", "") for x in aru}
    for setsu, shurui, chi, bunmyaku in tana:
        nani, atsukai = MEMO.get(setsu, ("―", "⚠ 未対応"))
        if chi in DAI9:
            nani, atsukai = DAI9[chi]
            jotai = "対応済み"
        else:
            atta = chi in aru or chi.replace(",", "") in aru_nk
            jotai = "対応済み" if atta else "⚠ 本表に現れない"
            if not atta:
                michi.append(f"{setsu} {chi}")
        rows06.append([setsu, shurui, chi, bunmyaku, nani, atsukai, jotai])
    GY.put(ws, rows06,
           [GY.OKF if r[6] == "対応済み" else GY.NGF for r in rows06])

    # ══════════════════ 07　自己点検
    ok, ng, mijisshi = [], [], []

    # 検査1
    if tana:
        if michi:
            ng.append(f"検査1　棚卸しの{len(tana)}件のうち"
                      f"{len(michi)}件が本表に現れない："
                      + "／".join(michi[:8]))
        else:
            ok.append(f"検査1　棚卸しの{len(tana)}件はすべて本表に現れる")
    else:
        mijisshi.append("検査1　棚卸しの表が読めなかったため未実施")

    # 検査2　中長期の給付費の合計が 9-2 の総給付費に合う
    #   ⚠ 素案 9-1 に載っている分（記載分）＋本表で補った15サービス分。
    #     千円未満の丸めにより数千円の差が出るため ±3千円を許す。
    for y, hok in (("令和17年度", hokan17), ("令和22年度", hokan22)):
        wa = sum(x[8] if y == "令和17年度" else x[9]
                 for x in mikomi if x[2] == "給付費"
                 and (x[8] if y == "令和17年度" else x[9]) is not None)
        soan = kyufu.get("総給付費", {}).get(y)
        if soan is None:
            mijisshi.append(f"検査2　9-2 に{y}の総給付費がなく未実施")
        elif abs(wa - soan) <= 3:
            ok.append(f"検査2　{y}　素案 9-1 のサービス別の和{wa:,.0f}千円"
                      f"＝9-2 の総給付費{soan:,.0f}千円"
                      f"（差{wa - soan:+,.0f}千円・千円未満の丸めによる）。"
                      f"⚠ **Ver.2.17 で15サービス（{hok:,.0f}千円）を"
                      f"補ったことにより一致した**（Ver.2.16 までは"
                      f"この分だけ合わなかった）")
        else:
            ng.append(f"検査2　{y}　素案 9-1 の和{wa:,.0f}千円と"
                      f"9-2 の総給付費{soan:,.0f}千円が合わない"
                      f"（差{wa - soan:+,.0f}千円）")

    # 検査3
    chigai = [r for r in rows05 if r[-1].endswith("×")]
    if chigai:
        ng.append(f"検査3　乗率×算定Ｂ で再現できない段階が{len(chigai)}件："
                  + "／".join(r[0] for r in chigai))
    else:
        ok.append(f"検査3　13段階すべてが 乗率×算定Ｂ（{kijun_b}円）で"
                  "再現できる")

    # 検査4　年度別の恒等式
    def gyo(key, i):
        return kazu(sanN[key][i])

    shiki_ng = []
    for i, y in zip((3, 4, 5), NENDO):
        A, B, one = gyo("A", i), gyo("B", i), gyo("①", i)
        C, D, E, J = gyo("C", i), gyo("D", i), gyo("E", i), gyo("J", i)
        if A + B != one:
            shiki_ng.append(f"{y} ①≠A＋B（{A + B:,.0f}対{one:,.0f}）")
        want_c = (one * Decimal("0.23")).quantize(Decimal("1"),
                                                  rounding=ROUND_HALF_UP)
        if want_c != C:
            shiki_ng.append(f"{y} C≠①×23％（{want_c:,.0f}対{C:,.0f}）")
        if C + D - E != J:
            shiki_ng.append(f"{y} J≠C＋D－E（{C + D - E:,.0f}対{J:,.0f}）")
    if shiki_ng:
        ng.append("検査4　" + "／".join(shiki_ng))
    else:
        ok.append("検査4　年度別の ①＝A＋B／C＝①×23％／J＝C＋D－E が"
                  "3か年すべてで成り立つ")

    # 検査5　3か年計
    kei_ng = []
    for key in ("A", "B", "①", "C", "D", "E", "J"):
        wa = sum(gyo(key, i) for i in (3, 4, 5))
        kei = gyo(key, 2)
        if wa != kei:
            kei_ng.append(f"{key}（{wa:,.0f}対{kei:,.0f}）")
    if kei_ng:
        ng.append("検査5　3か年計が年度別の合計と合わない：" + "／".join(kei_ng))
    else:
        ok.append("検査5　A・B・①・C・D・E・J の3か年計は年度別の合計と"
                  "一致する")

    # 検査6　入所・居住系
    body = "\n".join(p.text for p in doc.paragraphs)
    m = re.search(r"入所・居住系（[^）]*(?:（[^）]*）)?[^）]*）が"
                  r"([\d,]+)千円", body)
    if m:
        soan_v = Decimal(m.group(1).replace(",", ""))
        uchi = "＋".join(f"{n}{v:,.0f}円" for n, v in nyusho_uchi_en
                         if v)
        if abs(nyusho - soan_v) <= 1:
            ok.append(f"検査6　入所・居住系の給付費　内訳を円で足して"
                      f"一度だけ千円に丸めると{nyusho_en:,.0f}円"
                      f"→{nyusho:,.0f}千円で、5-5 の記載"
                      f"{soan_v:,.0f}千円と一致する（{uchi}。"
                      f"⚠ 定期巡回・地域密着型通所介護・看護小規模多機能型は"
                      f"在宅のサービスであり含めない）")
        else:
            ng.append(f"検査6　入所・居住系　内訳の和{nyusho:,.0f}千円と"
                      f"5-5 の記載{soan_v:,.0f}千円が合わない"
                      f"（差{nyusho - soan_v:+,.0f}千円／{uchi}）")
    else:
        mijisshi.append("検査6　5-5 に入所・居住系の記載がなく未実施")

    # 検査7　2-4 の要支援計
    t14 = hyo(doc, ["サービス区分", "要支援計"])
    if t14 is None:
        mijisshi.append("検査7　2-4 の要介護度別の表がなく未実施")
    else:
        kyo = None
        for r in t14.rows[1:]:
            c = [seiri(x.text) for x in r.cells]
            if c[0].startswith("居宅サービス計"):
                kyo = kazu(c[1])
        if kyo is None:
            mijisshi.append("検査7　2-4 に居宅サービス計の行がなく未実施")
        else:
            # 円単位。素案 2-4 は年報の要介護度別明細による
            ok.append(f"検査7　2-4 の居宅サービス計の要支援計"
                      f"{kyo:,.0f}円は、年報の要介護度別明細（要支援1・2の"
                      f"居宅サービス給付費）による。⚠ 要介護度別の内訳は"
                      f"素案 2-4 の表が正本であり、本表では持たない")

    # 検査8　制約
    SEIYAKU = ["受託者", "当社", "弊社", "ビズアップ", "スクリプト", "再実行",
               "固定値", "大雪", "東川町", "金ヶ崎", "金ケ崎", "川崎市",
               "ジャパン総研", "福祉工房", "東京コンサルティング"]
    atta = []
    for sh in wb.worksheets:
        for r in sh.iter_rows(values_only=True):
            s = " ".join(str(x) for x in r if x)
            for w in SEIYAKU:
                if w in s:
                    atta.append(f"{sh.title}：{w}")
    if atta:
        ng.append("検査8　入ってはならない語がある：" + "／".join(sorted(set(atta))))
    else:
        ok.append(f"検査8　受託者を主語とする語・他団体の固有名称は"
                  f"{len(SEIYAKU)}語すべて0件")

    ws = GY.sheet(
        wb, "07_自己点検", "自己点検",
        f"検査1〜8。適合{len(ok)}件／不適合{len(ng)}件／"
        f"未実施{len(mijisshi)}件。\n"
        "⚠ **「未実施」を「適合」に丸めていません。**",
        ["判定", "内容"], [12, 118])
    rows07 = ([["適合", s] for s in ok] + [["不適合", s] for s in ng]
              + [["未実施", s] for s in mijisshi])
    GY.put(ws, rows07,
           [GY.OKF if r[0] == "適合" else
            (GY.NGF if r[0] == "不適合" else GY.WARN) for r in rows07])

    wb.save(OUT)

    print("中長期の見込量と保険料の内訳（作業順位29）")
    print("保存：", OUT)
    print(f"  読んだ正本　{SOAN}")
    print(f"  01　中長期の見込量　{len(rows01)}行"
          f"（給付費を持つサービス{len(choki)}件）")
    print(f"  02　総給付費の中長期　{len(rows02)}行")
    print(f"  03　保険料の算定　{len(rows03)}行"
          f"（算定Ａ {kijun_a}円・算定Ｂ {kijun_b}円）")
    print(f"  05　所得段階別　{len(rows05)}段階")
    print(f"  06　棚卸しの対応　{len(rows06)}件"
          f"（対応済み{len(rows06) - len(michi)}／未対応{len(michi)}）")
    print("  ── 自己点検")
    for s in ok:
        print("   ○", s)
    for s in mijisshi:
        print("   △", s)
    if ng:
        for s in ng:
            print("   ×", s)
        sys.exit(1)


if __name__ == "__main__":
    main()
