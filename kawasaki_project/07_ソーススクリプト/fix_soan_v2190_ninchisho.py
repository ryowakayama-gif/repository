# -*- coding: utf-8 -*-
"""計画素案 Ver.2.19（認知症の人の数の推計・令和8年10月19日）

  入力　01_第10期_最新版成果品/川崎町_計画書素案_v2.18_範囲明記版.docx
  出力　01_第10期_最新版成果品/川崎町_計画書素案_v2.19_認知症推計版.docx

なぜ要るか
  第10期基本指針（案）別表一 事項十一「認知症の人の数及び関連施策の
  状況」は、確認すべき指標・状況として「自治体内の認知症の人の数**や
  推計値**」を掲げ、確認の観点として「地域における認知症の人の数や
  推計値を**算出し**」と定めている。
  ⚠ **Ver.2.18 までの素案には、認知症の人の数が1か所もなかった。**
  方針（認定調査の日常生活自立度Ⅱ以上により把握する）を 1-5 の対比表に
  書いていただけであり、**数値がなかった。**

  本版で、第6章6-1 に**推計値**を掲げる。
  算定は 07_ソーススクリプト/calc_ninchisho_R8.10.19.py にあり、
  本スクリプトは数値を持たず、そこから読む。

⚠ **認定調査の日常生活自立度Ⅱ以上の実数は、引き続き町のご提供を待つ
  （確認事項No.157）。** 推計値と実数は測るものが違い、
  **両者の差が「認定を受けていない認知症の方の規模」になる。**
  本版は推計値の側を掲げるものであり、実数を置き換えるものではない。

⚠ **将来（令和12・17・22年）の推計値は置いていない。**
  年齢階級別の**将来**人口が要る。本文に［要確認］として残す。

══════ 本版で加えるもの（第6章6-1「▌ 本町の認知症の人の数」）
  段落5・注3・表1（5列7行）

══════ 本版で直すもの
  ① 1-5 の別表一 対比表 事項十一の「本計画での記載箇所」
  ② 版表記 Ver.2.18 → Ver.2.19（本文・ヘッダー）
"""
import copy
import os
import re
import shutil
import sys
from decimal import Decimal, ROUND_HALF_UP

import docx
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from fix_soan_v111 import clone_table, set_el, set_tc        # noqa: E402

import importlib.util                                        # noqa: E402

spec = importlib.util.spec_from_file_location(
    "calc_ninchisho", os.path.join(HERE, "calc_ninchisho_R8.10.19.py"))
C = importlib.util.module_from_spec(spec)
sys.modules["calc_ninchisho"] = C
spec.loader.exec_module(C)

SRC = "01_第10期_最新版成果品/川崎町_計画書素案_v2.18_範囲明記版.docx"
DST = "01_第10期_最新版成果品/川崎町_計画書素案_v2.19_認知症推計版.docx"
KYU, SHIN = "Ver.2.18", "Ver.2.19"

ANCHOR = "▌ 川崎町の対応方針"        # この見出しの前に差し込む
TPL_SUB = "▌ 本町が目指す姿と「新しい認知症観」"
TPL_BODY = "この計画は、認知症があっても住み慣れた地域で"
TPL_WARN = "⚠ 基本法第13条第3項において準用する第12条第5項は"
TPL_SHUTTEN = "出典：介護保険事業状況報告"   # 出典は8.5pt（check_font 検査4）
TPL_TBL_HEAD = "令和3年度末"           # 2-1 の5列の表をひな形にする

SHUTTEN = (
    "出典：有病率は厚生労働省「認知症および軽度認知障害（ＭＣＩ）の"
    "高齢者数と有病率の将来推計」（令和5年度老人保健事業推進費等補助金"
    "「認知症及び軽度認知障害の有病率調査並びに将来推計に関する研究」"
    "（研究代表者　九州大学　二宮利治）より厚生労働省にて作成）の"
    "令和4年（2022年）の値。人口は川崎町住民基本台帳（令和8年6月末現在）")


def maru(x):
    """四捨五入して整数にする（銀行丸めを避ける）。"""
    return int(Decimal(str(x)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def hon(el):
    return "".join(x.text or "" for x in el.iter(qn("w:t")))


def honbun(jinko, ninchi, mci, korei):
    """差し込む本文を組む。戻り　[(種類, 中身), …]"""
    r = C.YUBYO["認知症"]
    ritsu = "、".join(f"{na}{r['全体'][i]}％"
                     for i, na in enumerate(C.KAIKYU))
    n_sei, m_sei = maru(ninchi["sei"]), maru(mci["sei"])
    kani = maru(ninchi["kani"])
    nana5 = sum(jinko[na][2] for na in C.KAIKYU[2:])

    rows = [["年齢階級", "人口（男）", "人口（女）",
             "認知症の人の数\n（推計）", "軽度認知障害（ＭＣＩ）\nの人の数（推計）"]]
    for i, na in enumerate(C.KAIKYU):
        o, n, k = jinko[na]
        rows.append([na, f"{o:,}人", f"{n:,}人",
                     f"{maru(ninchi['meisai'][i][6]):,}人",
                     f"{maru(mci['meisai'][i][6]):,}人"])
    rows.append(["計",
                 f"{sum(v[0] for v in jinko.values()):,}人",
                 f"{sum(v[1] for v in jinko.values()):,}人",
                 f"{n_sei:,}人", f"{m_sei:,}人"])

    return [
        ("sub", "　▌ 本町の認知症の人の数"),
        ("body",
         "認知症施策を進めるにあたっては、本町にどれくらいの認知症の方が"
         "おられるかをまず押さえる必要があります。国の第10期基本指針（案）"
         "も、計画の策定に当たって確認すべき指標の一つとして"
         "「自治体内の認知症の人の数や推計値」を掲げ、"
         "「地域における認知症の人の数や推計値を算出し、その上で"
         "認知症疾患医療センターを始めとする地域の医療資源の機能や"
         "利用者等を確認すること」としています。"),
        ("body",
         "そこで、国が全国の認知症高齢者数を推計する際と同じ方法により、"
         "本町の認知症の人の数を推計しました。年齢階級別・男女別の"
         "有病率を、本町の年齢階級別・男女別の人口に乗じて積み上げる"
         "ものです。有病率は、令和4年（2022年）に地域悉皆調査"
         "（調査率80％以上）を実施した4地域の結果に基づく国の推計値に"
         "よります。"),
        ("body",
         f"年齢階級別の有病率（男女計）は、{ritsu}です。"
         "男女別では、85歳以上で女性が男性を大きく上回ります"
         "（85〜89歳　女性37.2％・男性25.2％、"
         "90歳以上　女性55.1％・男性36.6％）。"),
        ("table", rows),
        ("shutten", SHUTTEN + "。四捨五入により、内訳の合計と計とは"
                              "一致しない場合があります。"),
        ("body",
         f"本町の65歳以上人口{korei:,}人のうち、認知症の方は"
         f"{n_sei}人（{ninchi['sei'] / korei * 100:.1f}％）、"
         f"軽度認知障害（ＭＣＩ）の方は{m_sei}人"
         f"（{mci['sei'] / korei * 100:.1f}％）、"
         f"合わせて{n_sei + m_sei}人"
         f"（{(ninchi['sei'] + mci['sei']) / korei * 100:.1f}％）と"
         f"見込まれます。全国では、令和4年時点で65歳以上の"
         f"認知症12.3％・ＭＣＩ15.5％、合わせて約28％とされています。"
         f"この{n_sei + m_sei}人という規模は、本章に掲げる"
         f"認知症サポーターの養成、認知症カフェ、チームオレンジ、"
         f"もの忘れ相談といった施策が向き合う相手の大きさを示すものです。"
         f"本町は75歳以上が{nana5:,}人と65歳以上の"
         f"{nana5 / korei * 100:.0f}％を占め、有病率が急に高くなる"
         f"85歳以上が{jinko['85〜89歳'][2] + jinko['90歳以上'][2]:,}人"
         f"おられます。"),
        ("warn",
         "⚠ この推計値は、要介護認定を受けておられるかどうかにかかわらず、"
         "地域にどれくらいの認知症の方がおられるかを示すものです。"
         "要介護認定の認定調査における認知症高齢者の日常生活自立度Ⅱ以上の"
         "人数は、認定を受けておられる方の中での把握であり、"
         "両者は測るものが異なります。自立度Ⅱ以上の人数は［要確認］であり、"
         "把握でき次第、推計値と並べて掲げます。"
         "両者の差が、認定を受けておられない認知症の方のおおよその規模に"
         "当たります。"),
        ("warn",
         f"⚠ 65歳以上人口{korei:,}人に全国の有病率12.3％を一律に乗じると"
         f"{kani}人となりますが、この12.3％は性年齢を調整した全国値で"
         f"あり、本町の年齢構成を反映しません。本町は65歳以上のうち"
         f"70歳代が厚く80歳代前半が薄いため、一律に乗じると"
         f"{kani - n_sei}人多く出ます。本計画では年齢階級別・男女別の"
         f"積み上げによります。"),
        ("body",
         "将来の認知症の人の数は、同じ有病率を年齢階級別の将来推計人口に"
         "乗じて算出します。令和12年・令和17年・令和22年の推計値は、"
         "年齢階級別の将来推計人口を整えたうえで掲げます［要確認］。"
         "国は、令和4年時点の性・年齢階級別の有病率が今後も一定と"
         "仮定して全国の将来推計を行っており、本町もこれに倣います。"),
    ]


def main():
    if not os.path.exists(SRC):
        raise SystemExit("入力がない：" + SRC)
    shutil.copy(SRC, DST)
    doc = docx.Document(DST)
    body = doc.element.body

    jinko = C.machi_kaikyu()
    korei = sum(v[2] for v in jinko.values())
    ninchi = C.santei(jinko, "認知症")
    mci = C.santei(jinko, "ＭＣＩ")

    # ── ひな形を拾う
    tpl = {}
    for el in body.iterchildren():
        if el.tag == qn("w:p"):
            s = hon(el)
            if TPL_SUB in s and "sub" not in tpl:
                tpl["sub"] = el
            elif s.startswith(TPL_BODY) and "body" not in tpl:
                tpl["body"] = el
            elif s.startswith(TPL_WARN) and "warn" not in tpl:
                tpl["warn"] = el
            elif s.startswith(TPL_SHUTTEN) and "shutten" not in tpl:
                tpl["shutten"] = el
        elif el.tag == qn("w:tbl") and "table" not in tpl:
            if TPL_TBL_HEAD in hon(el) and len(el.findall(qn("w:tblGrid"))[0]
                                               .findall(qn("w:gridCol"))) == 5:
                tpl["table"] = el
    nai = [k for k in ("sub", "body", "warn", "shutten", "table")
           if k not in tpl]
    if nai:
        raise SystemExit("ひな形が見つからない：" + "／".join(nai))

    # ── 差し込む位置
    oku = None
    for el in body.iterchildren():
        if el.tag == qn("w:p") and ANCHOR in hon(el):
            oku = el
            break
    if oku is None:
        raise SystemExit("差し込む位置が見つからない：" + ANCHOR)

    # ── 組んで差し込む
    items = honbun(jinko, ninchi, mci, korei)
    n_p = n_t = 0
    for kind, mono in items:
        if kind == "table":
            e = clone_table(tpl["table"], mono, widths=[22, 16, 16, 23, 23])
            n_t += 1
        else:
            e = copy.deepcopy(tpl[kind])
            set_el(e, mono)
            n_p += 1
        oku.addprevious(e)

    # ── 1-5 の別表一 対比表 事項十一
    n_hyo = 0
    FURUI = "認知症の人の数は認定調査の日常生活自立度Ⅱ以上により把握する"
    ATARASHII = (
        f"第6章6-1〜6-4。認知症の人の数の推計値"
        f"（{maru(ninchi['sei'])}人）及び軽度認知障害（ＭＣＩ）"
        f"（{maru(mci['sei'])}人）を6-1に掲げた。認定調査の"
        f"日常生活自立度Ⅱ以上の実数（確認事項No.157）と、"
        f"年齢階級別の将来推計人口による将来の推計値は未了")
    for tbl in body.iter(qn("w:tbl")):
        for tc in tbl.iter(qn("w:tc")):
            s = hon(tc)
            if FURUI in s:
                set_tc(tc, s.replace(
                    FURUI + "（確認事項No.157）。認知症疾患医療センター・"
                    "認知症サポート医は確認事項No.138", "")
                    .replace("第6章6-1〜6-4。", "")
                    + ATARASHII +
                    "。認知症疾患医療センター・認知症サポート医は"
                    "確認事項No.138")
                n_hyo += 1

    # ── 版表記
    n_ver = 0
    for el in body.iter(qn("w:t")):
        if el.text and KYU in el.text:
            el.text = el.text.replace(KYU, SHIN)
            n_ver += 1
    for sec in doc.sections:
        for hf in (sec.header, sec.footer, sec.even_page_header,
                   sec.first_page_header):
            for el in hf._element.iter(qn("w:t")):
                if el.text and KYU in el.text:
                    el.text = el.text.replace(KYU, SHIN)
                    n_ver += 1

    doc.save(DST)
    print(f"作成　{DST}")
    print(f"  段落 {n_p}／表 {n_t}／対比表のセル {n_hyo}／版表記 {n_ver}")

    # ══════════════════════════ 点検
    print()
    futeki = []

    def judge(ok, mi, setsu=""):
        print(("   ○ " if ok else "   × ") + mi + (("　" + setsu)
                                                   if setsu else ""))
        if not ok:
            futeki.append(mi)

    d2 = docx.Document(DST)
    zen = "\n".join([p.text for p in d2.paragraphs] +
                    [c.text for t in d2.tables for r in t.rows
                     for c in r.cells])
    d0 = docx.Document(SRC)
    zen0 = "\n".join([p.text for p in d0.paragraphs] +
                     [c.text for t in d0.tables for r in t.rows
                      for c in r.cells])

    print("  ── 検査1　加えたものが入っているか")
    judge("▌ 本町の認知症の人の数" in zen, "小見出しが入っている")
    judge(f"{maru(ninchi['sei'])}人" in zen,
          f"認知症の推計値 {maru(ninchi['sei'])}人 が入っている")
    judge(f"{maru(mci['sei'])}人" in zen,
          f"ＭＣＩの推計値 {maru(mci['sei'])}人 が入っている")
    judge(len(d2.tables) == len(d0.tables) + 1,
          f"表が1つ増えた（{len(d0.tables)}→{len(d2.tables)}）")

    print("  ── 検査2　表の中の数が算定と合うか")
    hyo = None
    for t in d2.tables:
        if t.rows[0].cells[0].text == "年齢階級":
            hyo = t
            break
    judge(hyo is not None, "加えた表を見つけた")
    if hyo is not None:
        ok = True
        for i, na in enumerate(C.KAIKYU):
            c = hyo.rows[i + 1].cells
            if (c[0].text != na
                    or c[1].text != f"{jinko[na][0]:,}人"
                    or c[2].text != f"{jinko[na][1]:,}人"
                    or c[3].text != f"{maru(ninchi['meisai'][i][6]):,}人"
                    or c[4].text != f"{maru(mci['meisai'][i][6]):,}人"):
                ok = False
                print(f"      × {na}")
        judge(ok, "6階級すべてが算定の値と一致")
        c = hyo.rows[7].cells
        judge(c[3].text == f"{maru(ninchi['sei']):,}人"
              and c[4].text == f"{maru(mci['sei']):,}人",
              "計の行が算定の値と一致")
        uchi = sum(maru(mci["meisai"][i][6]) for i in range(6))
        print(f"      （ＭＣＩの内訳の和 {uchi}人／計 "
              f"{maru(mci['sei'])}人。四捨五入による差を注記済み）")

    print("  ── 検査3　版表記")
    nokori = re.findall(r"Ver\.\d+\.\d+[a-z]?", zen)
    judge(all(v == SHIN for v in nokori),
          f"本文の版表記がすべて {SHIN}",
          f"（{len(nokori)}か所）" if nokori else "")
    hdr = "\n".join(p.text for s in d2.sections for p in s.header.paragraphs)
    judge(KYU not in hdr and SHIN in hdr, "ヘッダーの版表記も改めた")

    print("  ── 検査4　既にあった数値を動かしていないか")
    kazu0 = re.findall(r"\d[\d,]*", zen0)
    kazu2 = re.findall(r"\d[\d,]*", zen)
    from collections import Counter
    c0, c2 = Counter(kazu0), Counter(kazu2)
    kieta = {k: v - c2.get(k, 0) for k, v in c0.items() if v > c2.get(k, 0)}
    # 版表記（2.18→2.19）の分だけは減る
    kieta.pop("2.18", None)
    kieta.pop("18", None)
    judge(not kieta, "元からあった数が1つも消えていない",
          f"（消えた {kieta}）" if kieta else "")

    print("  ── 検査5　制約")
    for mi, pat in (("受託者を主語とする語",
                     r"(受託者|当社|弊社|ビズアップ|スクリプト|再実行)"),
                    ("強調記号の残り", r"(\*\*|__)"),
                    ("他団体の固有名称",
                     r"(大雪地区広域連合|東川町|東神楽町|上川町|金ヶ崎町)"),
                    ("川崎市との混同", r"川崎市")):
        atari = [s[:50] for s in zen.split("\n") if re.search(pat, s)]
        judge(not atari, f"{mi} が無い",
              f"→ {atari[:2]}" if atari else "")
    judge("久山町" not in zen,
          "有病率の調査地の町名を本文に残していない",
          "（他団体の固有名称を置かない。出典は研究名で示す）")

    print("  ── 検査6　未確定を推測で埋めていないか")
    judge("［要確認］" in zen, "将来の推計値を［要確認］として残している")
    judge("日常生活自立度Ⅱ以上" in zen,
          "自立度Ⅱ以上との違いを本文に書いている")

    print()
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        for m in futeki:
            print(f"      {m}")
        sys.exit(1)
    print("   ○ 検査1〜6　不適合0件")


if __name__ == "__main__":
    main()
