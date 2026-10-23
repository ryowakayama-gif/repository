# -*- coding: utf-8 -*-
"""計画素案 Ver.2.25 → Ver.2.26（版の表記の整理と、紙面の見え方の是正）
（令和8年10月23日）

  python3 07_ソーススクリプト/fix_soan_v2260_hyoki.py

ご指示（令和8年10月23日）
  「素案側の版表記も整理して下さい。北塩原村・大雪広域のブランチを確認の上、
   前回計画のフォント等を確認の上、計画の見え方との整合を踏まえて確認を」

(1) 版の表記を除く（素案は策定委員会に未提出。修正履歴は書かない）
    表紙「計画書素案 Ver.2.25」「（範囲明記版）」／ヘッダーの「Ver.2.25」／
    目次の注記「本素案Ver.2.25の組版」／表紙裏「本素案Ver.2.25は…反映した版」
    「本版では…新たに加えました」／3-? の「本版で確定しました」／
    出典の「訂正後」
    → 表紙は「計画書素案（令和8年10月時点）」とする。版の番号はファイル名で
      管理し、紙面には出さない（他の案件でも「協議用素案（令和8年8月時点）」の形）

(2) 行送りの指定に lineRule を付ける（他の案件の点検で見つかった型）
    ⚠ 行送りを指定している 5,621段落（本文654・表の中4,967）に lineRule が
    なかった。OOXML の既定は auto で、Word は w:line を「1行の240分の1倍」と
    読む。1行の高さは書体の字面で決まるため、**游ゴシックの入った町の Word では
    行送りが変わり、頁数と目次の頁番号が当方の実測と合わなくなる。**
    → 文字が行送りに収まる段落は exact（固定値）、見出しなど文字が行送り
      より大きい段落は atLeast（最小値）とする。値そのものは変えない。

(3) 表の列幅を実寸にする
    ⚠ 129表すべてで列幅の指定が壊れていた。
      87表　比率のつもりの値（24・136・32・208 等）が実寸（dxa＝1/20pt）として
            書かれている（列幅0.4〜3.7mm）
      38表　セルは割合（18%・36%…）だが、表の列の格子は均等割り（100・100・100）
    Word は自動調整で見かけ上収めることが多いが、環境によって列が細くなり
    **文字が1字ずつ縦に折れる**（実測で第1章の別表）。
    → 各表の比率を保ったまま本文幅（170mm）に掛けて実寸を出し、格子と各セルの
      幅を揃え、tblLayout を fixed とする。

(4) 表の中の長い文を左揃えに
    表の本文の行で、13字以上の文が中央揃えになっていた段落が441あった。
    第9期計画の表は文を左揃え・数を右寄せとしており、委員会資料でも
    「中央揃えで表中の文言が見えにくい」とのご指示を頂いた（令和8年10月22日）。
    → 見出しの行と、数・短い語（●・一部・―等）はそのままとし、長い文だけ左揃えに。

数値・文言は (1) のほかは変えない。
終了コード　0＝検査1〜6すべて適合／1＝1件でも不適合
"""
import os
import re
import shutil
import sys

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

MOTO = "01_第10期_最新版成果品/川崎町_計画書素案_v2.25_軽費老人ホーム確認版.docx"
SAKI = "01_第10期_最新版成果品/川崎町_計画書素案_v2.26_版表記整理版.docx"
JITEN = "計画書素案（令和8年10月時点）"
HEADER = "川崎町高齢者保健福祉計画・第10期介護保険事業計画（素案）"

# (冒頭, 新) 新が None なら段落を空にする
DAN = [
    ("計画書素案 Ver.2.25", JITEN),
    ("（範囲明記版）", None),
    ("⚠ 頁番号は、本素案Ver.2.25の組版によるものです。",
     "⚠ 頁番号は、現在の組版によるものです。図表の差替え・加筆により頁の割付けが"
     "動いた場合は、目次の頁番号も併せて更新します。"),
]
# 段落の中の語の置き換え（旧, 新）
OKIKAE = [
    ("※本素案Ver.2.25は、", "※この素案は、"),
    ("第1回策定委員会でのご意見を反映した版です。",
     "第1回策定委員会でのご意見を踏まえて作成したものです。"),
    ("本版では、介護保険制度改正の主な内容（第1章1-5）、日常生活圏域の設定"
     "（同1-6）、24時間対応サービスの確保方策及び介護施設整備に係る基本方針"
     "（第9章9-1）、低所得者支援（同9-3）を新たに加えました。", ""),
    ("現状値は本版で確定しました。", "現状値は確定しています。"),
    ("（令和8年9月11日出力・訂正後）", "（令和8年9月11日出力）"),
]
# 比率そのものが中身に合っていない表（1列目10mm・3列目14mmに16〜19字の
# 事項名が入り、1字ずつ縦に折れていた）。字数に見合う幅を mm で定める。
# 見出しの行の文字で当てる。合計は本文幅 170mm。
HABA_SHITEI = {
    ("手　続", "根　拠", "時　期"): [40, 34, 30, 66],
    ("事　項", "内　容", "施行の時期"): [38, 50, 26, 56],
}

RE_HAN = re.compile(r"Ver\.\s?\d|本版|版です|第\d+版|範囲明記版")

TBLPR = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
         "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
         "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout",
         "tblCellMar", "tblLook", "tblCaption", "tblDescription"]
TCPR = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders",
        "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign",
        "hideMark"]


def junban(el, order):
    idx = {qn("w:" + n): i for i, n in enumerate(order)}
    for k in sorted(list(el), key=lambda c: idx.get(c.tag, len(order))):
        el.append(k)


def dan_kaku(p, s):
    if p.runs:
        p.runs[0].text = s
        for r in p.runs[1:]:
            r.text = ""
    elif s:
        p.add_run(s)


def gyookuri(d):
    """lineRule のない行送りに exact／atLeast を付ける。"""
    kijun = d.styles["Normal"].font.size.pt
    n_ex = n_al = 0
    for p in d.element.iter(qn("w:p")):
        pPr = p.find(qn("w:pPr"))
        sp = pPr.find(qn("w:spacing")) if pPr is not None else None
        if sp is None or sp.get(qn("w:line")) is None or sp.get(qn("w:lineRule")):
            continue
        line = int(sp.get(qn("w:line")))
        if list(p.iter(qn("w:drawing"))):
            continue                      # 画像の段落は触らない（該当0件）
        szs = [int(s.get(qn("w:val"))) / 2 for s in p.iter(qn("w:sz"))] or [kijun]
        if line >= max(szs) * 20 * 1.15:
            sp.set(qn("w:lineRule"), "exact")
            n_ex += 1
        else:
            sp.set(qn("w:lineRule"), "atLeast")
            n_al += 1
    return n_ex, n_al


def hyohaba(d):
    """表の列幅を本文幅に掛けた実寸に揃え、固定する。"""
    sec = d.sections[0]
    honbun = int((sec.page_width - sec.left_margin - sec.right_margin) / 635)
    n = n_ire = 0
    for tbl in d.element.body.iter(qn("w:tbl")):
        if any(a.tag == qn("w:tbl") for a in tbl.iterancestors()):
            n_ire += 1                    # 入れ子の表は外の表のセルに従う
            continue
        tblPr = tbl.find(qn("w:tblPr"))
        grid = tbl.find(qn("w:tblGrid"))
        cols = grid.findall(qn("w:gridCol"))
        k = len(cols)
        # 比率：格子と同じ数のセルを持つ行の tcW を採る。なければ格子
        hiritsu = None
        for tr in tbl.findall(qn("w:tr")):
            tcs = tr.findall(qn("w:tc"))
            if len(tcs) != k:
                continue
            ws = []
            for tc in tcs:
                w = tc.find(qn("w:tcPr")).find(qn("w:tcW")) if tc.find(
                    qn("w:tcPr")) is not None else None
                v = (w.get(qn("w:w")) or "") if w is not None else ""
                v = v.rstrip("%")
                ws.append(float(v) if v.replace(".", "").isdigit() else 0)
            if all(x > 0 for x in ws):
                hiritsu = ws
                break
        if hiritsu is None:
            hiritsu = [float(c.get(qn("w:w"))) for c in cols]
        tw = tblPr.find(qn("w:tblW"))
        wari = 1.0
        if tw is not None and tw.get(qn("w:type")) == "pct":
            v = tw.get(qn("w:w")) or "100%"
            wari = (float(v.rstrip("%")) / 100 if v.endswith("%")
                    else float(v) / 5000)
        zentai = int(honbun * wari)
        goukei = sum(hiritsu)
        haba = [int(round(zentai * x / goukei)) for x in hiritsu]
        haba[-1] += zentai - sum(haba)
        for c, w in zip(cols, haba):
            c.set(qn("w:w"), str(w))
        if tw is None:
            tw = OxmlElement("w:tblW")
            tblPr.append(tw)
        tw.set(qn("w:type"), "dxa")
        tw.set(qn("w:w"), str(zentai))
        lay = tblPr.find(qn("w:tblLayout"))
        if lay is None:
            lay = OxmlElement("w:tblLayout")
            tblPr.append(lay)
        lay.set(qn("w:type"), "fixed")
        junban(tblPr, TBLPR)
        for tr in tbl.findall(qn("w:tr")):
            ichi = 0
            trPr = tr.find(qn("w:trPr"))
            if trPr is not None and trPr.find(qn("w:gridBefore")) is not None:
                ichi = int(trPr.find(qn("w:gridBefore")).get(qn("w:val")))
            for tc in tr.findall(qn("w:tc")):
                tcPr = tc.find(qn("w:tcPr"))
                if tcPr is None:
                    tcPr = OxmlElement("w:tcPr")
                    tc.insert(0, tcPr)
                gs = tcPr.find(qn("w:gridSpan"))
                span = int(gs.get(qn("w:val"))) if gs is not None else 1
                w = sum(haba[ichi:ichi + span])
                ichi += span
                tcW = tcPr.find(qn("w:tcW"))
                if tcW is None:
                    tcW = OxmlElement("w:tcW")
                    tcPr.append(tcW)
                tcW.set(qn("w:type"), "dxa")
                tcW.set(qn("w:w"), str(w))
                junban(tcPr, TCPR)
        n += 1
    return n, n_ire, honbun


RE_KAZU = re.compile(r"^[\s▲△＋+\-−―－0-9０-９,.，．％%人円千件点位年度月日回倍歳"
                     r"ヵか所床～〜/／（）()]*$")


def hidari(d):
    """表の本文の行の、中央揃えの長い文を左揃えにする。"""
    from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
    n = 0
    for t in d.tables:
        seen = set()
        for ri, row in enumerate(t.rows):
            if ri == 0:
                continue
            for c in row.cells:
                if c._tc in seen:
                    continue
                seen.add(c._tc)
                for p in c.paragraphs:
                    s = p.text.strip()
                    if (p.alignment == AL.CENTER and len(s) > 12
                            and not RE_KAZU.match(s)):
                        p.alignment = AL.LEFT
                        n += 1
    return n


def shitei(d):
    """HABA_SHITEI の表の列幅を mm で置き直す。"""
    n = 0
    for t in d.tables:
        mi = tuple(c.text.strip() for c in t.rows[0].cells)[:3]
        if mi not in HABA_SHITEI:
            continue
        mm = HABA_SHITEI[mi]
        cols = t._tbl.tblGrid.findall(qn("w:gridCol"))
        if len(cols) != len(mm):
            continue
        zentai = sum(int(c.get(qn("w:w"))) for c in cols)
        haba = [int(round(zentai * x / sum(mm))) for x in mm]
        haba[-1] += zentai - sum(haba)
        for c, w in zip(cols, haba):
            c.set(qn("w:w"), str(w))
        for tr in t._tbl.findall(qn("w:tr")):
            ichi = 0
            for tc in tr.findall(qn("w:tc")):
                gs = tc.tcPr.find(qn("w:gridSpan"))
                span = int(gs.get(qn("w:val"))) if gs is not None else 1
                tc.tcPr.find(qn("w:tcW")).set(qn("w:w"),
                                              str(sum(haba[ichi:ichi + span])))
                ichi += span
        n += 1
    return n


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

    print("計画素案 Ver.2.25 → Ver.2.26（版の表記の整理・紙面の見え方の是正）")
    print()
    if not os.path.exists(MOTO):
        print(f"   × 前の版がありません：{MOTO}")
        sys.exit(1)
    shutil.copy2(MOTO, SAKI)
    d = docx.Document(SAKI)
    mae = zenbun(d)

    print("  ── ① 版の表記")
    n = 0
    for atama, atarashii in DAN:
        hit = [p for p in d.paragraphs if p.text.strip().startswith(atama)]
        if len(hit) != 1:
            judge(False, f"『{atama[:20]}』が{len(hit)}件")
            continue
        dan_kaku(hit[0], atarashii or "")
        n += 1
    for kyu, atarashii in OKIKAE:
        hit = [p for p in d.paragraphs if kyu in p.text]
        if len(hit) != 1:
            judge(False, f"『{kyu[:20]}』が{len(hit)}件")
            continue
        dan_kaku(hit[0], hit[0].text.replace(kyu, atarashii))
        n += 1
    n_hd = 0
    for s in d.sections:
        for p in s.header.paragraphs:
            if "計画書素案" in p.text:
                dan_kaku(p, HEADER)
                n_hd += 1
    judge(n == len(DAN) + len(OKIKAE) and n_hd >= 1,
          f"本文 {n}か所・ヘッダー {n_hd}か所")

    print("  ── ② 行送りに lineRule を付ける")
    n_ex, n_al = gyookuri(d)
    judge(n_ex + n_al > 5000, f"exact {n_ex}段落・atLeast {n_al}段落")

    print("  ── ③ 表の列幅を実寸に")
    n_t, n_ire, honbun = hyohaba(d)
    judge(n_t == len(d.tables), f"{n_t}表（本文幅 {honbun} twip＝"
          f"{honbun / 1440 * 25.4:.0f}mm、入れ子 {n_ire}）")
    print("  ── ④ 表の中の長い文を左揃えに")
    n_h = hidari(d)
    judge(n_h > 0, f"{n_h}段落")
    n_s = shitei(d)
    judge(n_s == len(HABA_SHITEI), f"比率を定め直した表 {n_s}（制度改正・手続の表）")
    d.save(SAKI)

    d2 = docx.Document(SAKI)
    ato = zenbun(d2)
    print("  ── 検査1　版の表記が残っていない")
    judge(not RE_HAN.search(ato), "Ver.・本版・第N版・範囲明記版 が0件",
          "／".join(m.group(0) for m in RE_HAN.finditer(ato))[:60])
    judge(JITEN in ato and HEADER in ato, "表紙は時点、ヘッダーは計画名（素案）")
    print("  ── 検査2　行送り")
    nashi = sum(1 for p in d2.element.iter(qn("w:p"))
                if p.find(qn("w:pPr")) is not None
                and p.find(qn("w:pPr")).find(qn("w:spacing")) is not None
                and p.find(qn("w:pPr")).find(qn("w:spacing")).get(qn("w:line"))
                and not p.find(qn("w:pPr")).find(qn("w:spacing"))
                .get(qn("w:lineRule")))
    judge(nashi == 0, f"lineRule のない行送りが0件（実際 {nashi}）")
    print("  ── 検査3　表の列幅")
    hosoi = []
    for ti, t in enumerate(d2.tables):
        ws = [int(g.get(qn("w:w"))) for g in
              t._tbl.tblGrid.findall(qn("w:gridCol"))]
        if min(ws) < 300 or abs(sum(ws) - honbun) > honbun * 0.05:
            hosoi.append(ti)
    judge(not hosoi, "列幅5mm未満・本文幅と食い違う表が0件", str(hosoi[:8]))
    print("  ── 検査4　⚠ 数・文言が消えていない（版の表記を除く）")
    kesu = sorted({m.group(0) for m in re.finditer(
        r"[0-9０-９][0-9０-９,.]*(?:人|％|千円|件|名|歳|回|円|点)", mae)
        if mae.count(m.group(0)) > ato.count(m.group(0))})
    judge(not kesu, "消えた数が0件", "／".join(kesu[:6]))
    judge(len(d2.paragraphs) == len(d.paragraphs)
          and len(d2.tables) == len(d.tables), "段落数・表数が変わらない")
    print("  ── 検査5　［要確認］［要協議］の件数")
    for w in ("［要確認］", "［要協議］"):
        judge(mae.count(w) == ato.count(w), f"{w} {ato.count(w)}件（不変）")
    print("  ── 検査6　受託者を主語とする語")
    aru = [w for w in ("受託者", "当社", "ビズアップ", "スクリプト")
           if w in ato]
    judge(not aru, "0件", "／".join(aru))

    print()
    print(f"  {SAKI}")
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        for m in futeki:
            print(f"      {m}")
        sys.exit(1)
    print("   ○ 検査1〜6　不適合0件")


if __name__ == "__main__":
    main()
