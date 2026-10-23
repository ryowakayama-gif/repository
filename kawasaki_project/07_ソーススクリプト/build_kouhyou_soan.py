# -*- coding: utf-8 -*-
"""協議用素案から公表用の計画書（案）を組む
（令和8年10月23日。確認事項No.187 のご指示による運用）

  python3 07_ソーススクリプト/build_kouhyou_soan.py

2つの版の切り分け
  **内容の正本は協議用素案である。** 本スクリプトは協議用素案の docx を読み、
  協議のためだけに置いた記述を落として公表用を組む。本文・数値を別に持たない
  ため、協議用素案を改めれば公表用も追随する。

  |            | 協議用素案                     | 公表用の計画書（案）            |
  |------------|--------------------------------|---------------------------------|
  | 読み手     | 町（保健福祉課）               | 町民・策定委員会の委員          |
  | 表紙       | 協議用素案（令和N年N月時点）   | 計画書（案）。未確定が残る間は（体裁確認用） |
  | ヘッダー   | …（協議用素案）                | 計画名のみ                      |
  | ⚠ の注記   | 残す（町へのお尋ね・途中経過） | 協議の文を落とし、残る文は ※ の注記に |
  | 確認事項No | 残す                           | 落とす                          |
  | 【町確認】 | 残す                           | ［要確認］に揃える              |
  | ●の要点    | 残す                           | 法令・交付金への対応の注記は落とす |
  | 見出し・図表の題・数値 | ―                  | 協議用と同じ（変えない）        |

落とす文の決め方
  ⚠ の段落・表の中の ⚠ の文・● の段落を文（「。」）に分け、協議の語
  （確認事項No・お諮り・ご確認・確認中・入手次第・追記・照合・ご提供・ご記入・
  お願い・町と協議・予定しています・協議事項・作業用のファイル名 等）を含む文を
  落とす。**［要確認］［要協議］［要内訳］を含む文は落とさない**（未確定を隠さない）。
  落とした文は 05_試算・管理シート/川崎町_公表用で落とした文_<版>.xlsx に残す。

出力　01_第10期_最新版成果品/川崎町_計画書案_公表用_v<協議用素案の版>.docx
終了コード　0＝検査1〜8すべて適合／1＝1件でも不適合
"""
import os
import re
import shutil
import sys

import docx
import openpyxl
from docx.oxml.ns import qn
from openpyxl.styles import Alignment, Font

KYOGI = "01_第10期_最新版成果品/川崎町_計画書素案_v2.30_協議用.docx"
BAN = re.search(r"_v(\d+\.\d+)_", KYOGI).group(1)
SAKI = f"01_第10期_最新版成果品/川崎町_計画書案_公表用_v{BAN}.docx"
KIROKU = f"05_試算・管理シート/川崎町_公表用で落とした文_v{BAN}.xlsx"

HYOSHI_KYOGI = re.compile(r"^協議用素案（.+時点）$")
HEADER = "川崎町高齢者保健福祉計画・第10期介護保険事業計画"

# 段落ごと落とすもの（協議の運用の説明）
OTOSU_ATAMA = ("⚠ 頁番号は、現在の組版", "※この素案は", "この素案は")
# 協議の語（文を落とす）
KYOGI_GO = ("確認事項No", "お諮り", "ご確認", "確認中", "入手次第", "追記します",
            "照合します", "照合を行った", "ご提供を待", "ご提供により把握", "ご記入",
            "お願いし", "町と協議", "予定しています", "受領していない",
            "記述を確定", "協議事項", "でご審議いただき", ".xlsx", "別添『",
            "全国一覧表（", "宮城県を通じて確認")
# ● の要点のうち落とすもの（法令・交付金への対応の注記）
MARU_OTOSU = re.compile(r"(に対応する(計画の核心部分である)?。$|回復|\.xlsx)")
RE_KAKUNIN = re.compile(r"（[^（）]*確認事項No\.[^（）]*）")
MIKAKUTEI = ("［要確認］", "［要協議］", "［要内訳］")


HENKAN = {"要確認": 0, "要協議": 0}


def seiki(s):
    """確認事項No を伴う未確定の印を、未確定の印だけの形にそろえる。
       ［要協議・確認事項No.146］→［要協議］
       「…は確認事項No.138」（文末）→「…は［要確認］」"""
    def f1(m):
        HENKAN[m.group(1)] = HENKAN.get(m.group(1), 0) + 1
        return f"［{m.group(1)}］"
    s = re.sub(r"［(要確認|要協議|要内訳)・確認事項No\.[^］]*］", f1, s)

    def f2(m):
        HENKAN["要確認"] += 1
        return "は［要確認］"
    s = re.sub(r"は確認事項No\.[0-9０-９・〜～]+(?=$|[。、])", f2, s)
    return s


def bunkatsu(s):
    return [b for b in re.split(r"(?<=。)", s) if b.strip()]


def kyogi_no_bun(b):
    if "［要" in b:
        return False
    return any(g in b for g in KYOGI_GO)


def kakikae(p, s):
    if p.runs:
        p.runs[0].text = s
        for r in p.runs[1:]:
            r.text = ""
    elif s:
        p.add_run(s)


def seiri(s, kigou):
    """⚠ の文を整える。協議の文を落とし、残りを返す（落とした文も返す）。"""
    s = RE_KAKUNIN.sub("", seiki(s))
    nokosu, otoshi = [], []
    for b in bunkatsu(s):
        (otoshi if kyogi_no_bun(b) else nokosu).append(b.strip())
    t = "".join(nokosu)
    t = re.sub(r"^\s*⚠\s*", kigou, t)
    t = re.sub(r"\s*⚠\s*", "", t)
    return t.strip(), otoshi


def tsujo(s):
    """⚠ のない段落：確認事項No の括弧と、それを含む文だけを除く。"""
    s = RE_KAKUNIN.sub("", seiki(s))
    nokosu, otoshi = [], []
    for b in bunkatsu(s) or [s]:
        if (("確認事項No" in b or ".xlsx" in b or "別添『" in b)
                and not any(m in b for m in MIKAKUTEI)):
            otoshi.append(b.strip())
        else:
            nokosu.append(b)
    return "".join(nokosu), otoshi


def zenbun(d):
    out = [p.text for p in d.paragraphs]
    out += [c.text for t in d.tables for r in t.rows for c in r.cells]
    for s in d.sections:
        for part in (s.header, s.footer):
            out += [p.text for p in part.paragraphs]
    return "\n".join(out)


def main():
    futeki = []

    def judge(ok, mi, setsu=""):
        print(("   ○ " if ok else "   × ") + mi + (("　" + setsu)
                                                   if setsu else ""))
        if not ok:
            futeki.append(mi)

    print(f"公表用の計画書（案）を組む（協議用素案 Ver.{BAN} から）")
    print()
    if not os.path.exists(KYOGI):
        print(f"   × 協議用素案がありません：{KYOGI}")
        sys.exit(1)
    shutil.copy2(KYOGI, SAKI)
    d = docx.Document(SAKI)
    mae = zenbun(d)
    n_sect = len(list(d.element.body.iter(qn("w:sectPr"))))
    kiroku = []          # (場所, 区分, 落とした文)

    print("  ── ① 本文の段落")
    kesu = []
    for i, p in enumerate(d.paragraphs):
        t = p.text.strip()
        if not t:
            continue
        if list(p._p.iter(qn("w:sectPr"))):
            continue                         # 節の区切りを持つ段落は落とさない
        if t.startswith(OTOSU_ATAMA):
            kesu.append(p._p)
            kiroku.append((f"p{i}", "協議の運用の説明", t))
            continue
        if t.startswith("⚠"):
            s, ot = seiri(t, "※")
            for b in ot:
                kiroku.append((f"p{i}", "⚠ の協議の文", b))
            if s.strip("※ "):
                kakikae(p, s)
            else:
                kesu.append(p._p)
            continue
        if t.startswith("●") and MARU_OTOSU.search(t) and not any(
                m in t for m in MIKAKUTEI):
            kesu.append(p._p)
            kiroku.append((f"p{i}", "● 法令・交付金への対応の注記", t))
            continue
        if "確認事項No" in t or ".xlsx" in t or "別添『" in t or "⚠" in t:
            s, ot = tsujo(t)
            s = re.sub(r"\s*⚠\s*", "", s)
            for b in ot:
                kiroku.append((f"p{i}", "確認事項No・作業用の資料を含む文", b))
            if s.strip():
                kakikae(p, s)
            else:
                kesu.append(p._p)
    for el in kesu:
        el.getparent().remove(el)
    judge(True, f"落とした段落 {len(kesu)}")

    print("  ── ② 表の中")
    n_cell = 0
    for ti, tbl in enumerate(d.tables):
        seen = set()
        for row in tbl.rows:
            for c in row.cells:
                if c._tc in seen:
                    continue
                seen.add(c._tc)
                for p in c.paragraphs:
                    t = p.text
                    if not ("⚠" in t or "確認事項No" in t or ".xlsx" in t):
                        continue
                    s, ot = seiri(t, "※") if "⚠" in t else tsujo(t)
                    for b in ot:
                        kiroku.append((f"表{ti}", "表の中の協議の文", b))
                    kakikae(p, s)
                    n_cell += 1
    judge(True, f"書き改めた表の中の段落 {n_cell}")

    print("  ── ③ 未確定の印・語")
    n_mk = 0
    for p in list(d.paragraphs) + [p for t in d.tables for r in t.rows
                                   for c in r.cells for p in c.paragraphs]:
        if "【町確認】" in p.text or "本素案" in p.text:
            n_mk += p.text.count("【町確認】")
            kakikae(p, p.text.replace("【町確認】", "［要確認］")
                    .replace("本素案", "本計画"))
    judge(True, f"【町確認】→［要確認］ {n_mk}")

    print("  ── ④ 表紙・ヘッダー")
    zan = sum(zenbun(d).count(m) for m in MIKAKUTEI)
    hyoshi = "計画書（案）" + ("（体裁確認用）" if zan else "")
    hit = [p for p in d.paragraphs if HYOSHI_KYOGI.match(p.text.strip())]
    judge(len(hit) == 1, f"表紙の名称 → {hyoshi}（未確定 {zan}箇所）")
    for p in hit:
        kakikae(p, hyoshi)
    n_hd = 0
    for s in d.sections:
        for p in s.header.paragraphs:
            if "協議用素案" in p.text:
                kakikae(p, HEADER)
                n_hd += 1
    judge(n_hd >= 1, f"ヘッダー {n_hd}か所")
    d.save(SAKI)

    # 落とした文の記録
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "落とした文"
    ws.append(["No", "場所（協議用素案）", "区分", "落とした文"])
    for k, (basho, kubun, bun) in enumerate(kiroku, 1):
        ws.append([k, basho, kubun, bun])
    for col, w in zip("ABCD", (6, 14, 28, 110)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows():
        for c in row:
            c.font = Font(name="游ゴシック", size=9, bold=c.row == 1)
            c.alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(KIROKU)

    d2 = docx.Document(SAKI)
    ato = zenbun(d2)
    k = docx.Document(KYOGI)
    print("  ── 検査1　協議の印・語が残っていない")
    for go in ("⚠", "確認事項No", "【町確認】", "本素案", "この素案", "協議用",
               "受託者", "当社", "ビズアップ", "スクリプト", ".xlsx", "別添『"):
        judge(go not in ato, f"「{go}」0件", str(ato.count(go)))
    print("  ── 検査2　⚠ 未確定を隠していない")
    # 落とした文の【町確認】は、印そのものを説明する文（「【町確認】は…」）に限る。
    # 値に付いた印は落とさない
    ot_mk = [b for _, _, b in kiroku if "【町確認】" in b]
    judge(all("【町確認】は" in b for b in ot_mk),
          f"落とした文の【町確認】{len(ot_mk)}件はいずれも印の説明の文")
    judge(ato.count("［要確認］") == mae.count("［要確認］") + mae.count("【町確認】")
          - sum(b.count("【町確認】") for b in ot_mk) + HENKAN["要確認"],
          f"［要確認］{ato.count('［要確認］')}＝協議用の［要確認］"
          f"{mae.count('［要確認］')}＋【町確認】{mae.count('【町確認】')}"
          f"−印の説明の文 {len(ot_mk)}＋確認事項Noから {HENKAN['要確認']}")
    judge(ato.count("［要協議］") == mae.count("［要協議］") + HENKAN["要協議"],
          f"［要協議］{ato.count('［要協議］')}＝協議用{mae.count('［要協議］')}"
          f"＋［要協議・確認事項No］から {HENKAN['要協議']}")
    judge(ato.count("［要内訳］") == mae.count("［要内訳］"), "［要内訳］同数")
    judge("［要協議・" not in ato and "［要確認・" not in ato, "印に確認事項Noが残っていない")
    judge(("体裁確認用" in ato) == (zan > 0), "未確定が残る間は表紙に（体裁確認用）")
    print("  ── 検査3　節の区切り・表・図")
    judge(len(list(d2.element.body.iter(qn("w:sectPr")))) == n_sect,
          f"節の区切り {n_sect}（ヘッダー・頁番号を失っていない）")
    judge(len(d2.tables) == len(k.tables), f"表 {len(d2.tables)}（協議用と同じ）")
    judge(len(d2.inline_shapes) == len(k.inline_shapes),
          f"図 {len(d2.inline_shapes)}（協議用と同じ）")
    print("  ── 検査4　見出し・図表の題は協議用と同じ")
    for rx, mi in ((r"^第\s*\d+\s*章$", "章"), (r"^\d+-\d+　", "節"),
                   (r"^図\d+-\d+", "図の題"), (r"^　?▌", "小見出し")):
        a = [p.text for p in k.paragraphs if re.match(rx, p.text.strip())]
        b = [p.text for p in d2.paragraphs if re.match(rx, p.text.strip())]
        judge(a == b, f"{mi} {len(b)}件が同じ")
    print("  ── 検査5　表の数値は協議用と同じ")
    kazu = re.compile(r"[0-9０-９][0-9０-９,.]*(?:人|％|千円|円|件|点|頁|位|か所)")
    ta = sorted(m.group(0) for t in k.tables for r in t.rows for c in r.cells
                for m in kazu.finditer(c.text))
    tb = sorted(m.group(0) for t in d2.tables for r in t.rows for c in r.cells
                for m in kazu.finditer(c.text))
    from collections import Counter
    kie = Counter(ta) - Counter(tb)
    ot_zen = "\n".join(b for _, _, b in kiroku)
    judge(all(x in ot_zen for x in kie), "表の中で消えた数はすべて落とした文に現れる",
          "／".join(list(kie)[:5]))
    print("  ── 検査6　本文で消えた数は落とした文に現れる")
    hb = Counter(m.group(0) for m in kazu.finditer("\n".join(p.text for p in k.paragraphs)))
    ha = Counter(m.group(0) for m in kazu.finditer("\n".join(p.text for p in d2.paragraphs)))
    kie2 = [x for x in (hb - ha) if x not in ot_zen]
    judge(not kie2, "本文で消えた数はすべて落とした文に現れる", "／".join(kie2[:5]))
    print("  ── 検査7　他団体名・施設名")
    for na in ("北塩原", "金ヶ崎", "大雪", "川崎市", "蔵王グリーンホーム", "鶴寿会"):
        judge(na not in ato, f"「{na}」0件")
    print("  ── 検査8　落とした文の記録")
    judge(len(kiroku) > 0 and os.path.exists(KIROKU), f"{len(kiroku)}件 → {KIROKU}")

    print()
    print(f"  {SAKI}")
    print(f"  段落 協議用 {len(k.paragraphs)} → 公表用 {len(d2.paragraphs)}")
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        for m in futeki:
            print(f"      {m}")
        sys.exit(1)
    print("   ○ 検査1〜8　不適合0件")


if __name__ == "__main__":
    main()
