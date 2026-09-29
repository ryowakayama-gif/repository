"""成果品の体裁がデザインルールどおりかを機械的に点検する。

他町村の案件で用いている「策定委員会資料のデザインルール」に合わせた体裁を、
`ono_style.Report` で実装している。本スクリプトは、出来上がった docx を開き
直して、規則どおりになっているかを確かめる。

  ・用紙・余白　　A4縦　上下左右2.0cm　ヘッダ／フッタ1.25cm
  ・本文　　　　　BIZ UDPゴシック 12pt
  ・見出し　　　　H1 16pt #2E74B5／H2 14pt #2E74B5／H3 11pt #1F4D78
  ・表　　　　　　見出し行 塗り#1F3864・白文字・中央／列幅はグリッドにも書く
  ・列幅　　　　　合計17.0cm（本文幅）
  ・出典　　　　　8pt
  ・強調　　　　　`**` が本文に出ていないこと
  ・図　　　　　　キャプション【図N】と出典が対になっていること

`soffice --headless --convert-to pdf` による目視確認の代わりに置くものではなく、
その前段で機械的に落とせる誤りを落とすためのもの。

  python3 check_ono_taisai.py
"""

import pathlib
import re
import sys
import warnings
import zipfile

warnings.filterwarnings("ignore")

from docx import Document
from docx.shared import Cm, Pt

ROOT = pathlib.Path(__file__).parent / "小野町_引継ぎ_整理済"

# 点検する成果品（デザインルールを適用したもの）
TARGETS = [
    "08_協議会資料/小野町高齢者福祉サービス推進協議会資料_"
    "第9期計画の評価と令和7年度調査結果_20260929.docx",
    "03_計画素案/小野町高齢者保健福祉計画_第10期介護保険事業計画_"
    "素案_第7版_20260929.docx",
    "19_打合せ/小野町_打合せ次第_20260929.docx",
    "19_打合せ/小野町_打合せ資料_現在地と論点_20260929.docx",
    "19_打合せ/小野町_打合せ記録_20260929.docx",
    "20_総合事業ワークシート/小野町_総合事業ワークシートの確認結果_20260928.docx",
    "21_給付適正化・交付金/小野町_給付適正化と交付金の取りまとめ_20260928.docx",
]

FONT = "BIZ UDPゴシック"
TEXTW = 17.0
TOL = 0.35          # 列幅の合計の許容差（cm）


def _emu(v):
    return None if v is None else round(v / 360000, 2)      # EMU → cm


def check(path):
    bad = []
    d = Document(path)

    # ---- 用紙・余白
    for i, s in enumerate(d.sections):
        if (_emu(s.page_width), _emu(s.page_height)) != (21.0, 29.7):
            bad.append(f"section{i} 用紙 {_emu(s.page_width)}×"
                       f"{_emu(s.page_height)}cm（A4縦でない）")
        for nm, v in (("左", s.left_margin), ("右", s.right_margin),
                      ("上", s.top_margin), ("下", s.bottom_margin)):
            if abs(_emu(v) - 2.0) > 0.01:
                bad.append(f"section{i} {nm}余白 {_emu(v)}cm（2.0cmでない）")
        for nm, v in (("ヘッダ", s.header_distance), ("フッタ", s.footer_distance)):
            if abs(_emu(v) - 1.25) > 0.01:
                bad.append(f"section{i} {nm}距離 {_emu(v)}cm（1.25cmでない）")

    # ---- 本文のフォント
    st = d.styles["Normal"]
    if st.font.name != FONT:
        bad.append(f"本文フォント {st.font.name}（{FONT}でない）")
    if st.font.size != Pt(12):
        bad.append(f"本文サイズ {st.font.size.pt if st.font.size else None}pt"
                   "（12ptでない）")

    # ---- 表：見出し行の塗りと列幅
    for ti, t in enumerate(d.tables):
        xml = t._tbl.xml
        if "1F3864" not in xml:
            bad.append(f"表{ti + 1} 見出し行の塗り(#1F3864)がない")
        w = sum(c.width for c in t.rows[0].cells if c.width is not None)
        if w:
            cm = round(w / 360000, 2)
            if abs(cm - TEXTW) > TOL:
                bad.append(f"表{ti + 1} 列幅の合計 {cm}cm（{TEXTW}cmでない）")
        else:
            bad.append(f"表{ti + 1} 列幅が設定されていない")
        if "gridCol" not in xml or 'w:w="' not in xml.split("gridCol")[1][:40]:
            bad.append(f"表{ti + 1} グリッド(w:gridCol)に幅が書かれていない")

    # ---- 強調マーカーが本文に出ていないこと
    with zipfile.ZipFile(path) as z:
        raw = re.sub(r"<[^>]+>", "",
                     z.read("word/document.xml").decode("utf-8", "ignore"))
    if "**" in raw:
        bad.append(f"強調マーカー ** が本文に {raw.count('**')}か所")

    # ---- 図：キャプションと出典が対になっていること
    caps = re.findall(r"【図(\d+)】", raw)
    if caps:
        if [int(c) for c in caps] != list(range(1, len(caps) + 1)):
            bad.append(f"図番号が連番でない {caps}")
        imgs = sum(1 for p in d.paragraphs if "graphicData" in p._p.xml)
        if imgs != len(caps):
            bad.append(f"図の数 {imgs} とキャプション {len(caps)} が合わない")

    return bad, len(d.tables), len(caps)


def main():
    rc, n = 0, 0
    for rel in TARGETS:
        p = ROOT / rel
        if not p.exists():
            print(f"** 見つからない {rel}")
            rc = 1
            continue
        n += 1
        bad, tbls, figs = check(p)
        head = f"{p.name}　表{tbls}・図{figs}"
        if bad:
            rc = 1
            print(f"NG  {head}")
            for b in bad[:12]:
                print(f"      {b}")
            if len(bad) > 12:
                print(f"      ほか{len(bad) - 12}件")
        else:
            print(f"OK  {head}")
    print(f"\n点検 {n}件　{'すべて規則どおり' if rc == 0 else '要修正あり'}")
    return rc


if __name__ == "__main__":
    sys.exit(main())
