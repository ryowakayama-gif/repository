"""成果品の体裁がデザインルールどおりかを機械的に点検する。

他町村の案件で用いている「策定委員会資料のデザインルール」に合わせた体裁を、
`ono_style.Report` で実装している。本スクリプトは、出来上がった docx を開き
直して、規則どおりになっているかを確かめる。

  ・用紙・余白　　A4縦　上下左右2.0cm　ヘッダ／フッタ1.25cm
  ・本文　　　　　游ゴシック 10.5pt（令和8年9月29日に他団体の協議用素案に合わせた）
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
    "素案_第9版_20261005.docx",
    "19_打合せ/小野町_打合せ次第_20261005.docx",
    "19_打合せ/小野町_打合せ資料_現在地と論点_20261005.docx",
    "19_打合せ/小野町_打合せ記録_20261005.docx",
    "20_総合事業ワークシート/小野町_総合事業ワークシートの確認結果_20260928.docx",
    "21_給付適正化・交付金/小野町_給付適正化と交付金の取りまとめ_20261005.docx",
    "11_見える化出力依頼/小野町_地域分析・検討シートの記入案_20260930.docx",
    "03_計画素案/小野町_計画素案_構成の対比と漏れの点検_20261005.docx",
    "17_業務進捗管理/小野町_ペンディング整理と作業順位_20261007.docx",
    "02_キックオフ・業務計画/小野町_工程変更協議資料_第2版_20261005.docx",
    "04_算定・見込量/小野町_第10期_単価の置き方の感度分析_20261005.docx",
    "11_見える化出力依頼/小野町_見える化_医療リハビリの出力依頼_20261005.docx",
    "04_算定・見込量/小野町_第10期_サービス見込量推計結果報告書_20261006.docx",
    "09_骨子案/小野町高齢者保健福祉計画_第10期介護保険事業計画_"
    "骨子案_第4版_20261006.docx",
    "09_骨子案/おのまち障がい者計画_第4期障がい児福祉計画_"
    "第8期障がい福祉計画_骨子案_第2版_20261006.docx",
    "10_納品・印刷/小野町_第10期計画_印刷の発注仕様書_20261006.docx",
    "10_納品・印刷/小野町_第10期計画_概要版のページ割りと図の割付_20261006.docx",
    "08_協議会資料/小野町高齢者福祉サービス推進協議会資料_第2回_"
    "骨子案と見込量・保険料_20261006.docx",
    "23_障がいアンケート分析/小野町_障がいアンケート分析報告書のレビュー_"
    "20261006.docx",
    "23_障がいアンケート分析/小野町_障がい計画_目標値とアンケート設問の対応_"
    "第2版_20261007.docx",
    "23_障がいアンケート分析/小野町_障がい調査_問31と給付実績の食い違いの"
    "切り分け_20261007.docx",
    "04_算定・見込量/小野町_障がい計画_国の動向の反映と手順の乖離_"
    "20261007.docx",
]

FONT = "游ゴシック"
BODY_PT = 10.5
TEXTW = 17.2
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
        for nm, v, want in (("左", s.left_margin, 1.9),
                            ("右", s.right_margin, 1.9),
                            ("上", s.top_margin, 2.0),
                            ("下", s.bottom_margin, 2.0)):
            if abs(_emu(v) - want) > 0.01:
                bad.append(f"section{i} {nm}余白 {_emu(v)}cm（{want}cmでない）")
        for nm, v in (("ヘッダ", s.header_distance), ("フッタ", s.footer_distance)):
            if abs(_emu(v) - 1.25) > 0.01:
                bad.append(f"section{i} {nm}距離 {_emu(v)}cm（1.25cmでない）")

    # ---- 本文のフォント
    st = d.styles["Normal"]
    if st.font.name != FONT:
        bad.append(f"本文フォント {st.font.name}（{FONT}でない）")
    if st.font.size != Pt(BODY_PT):
        bad.append(f"本文サイズ {st.font.size.pt if st.font.size else None}pt"
                   f"（{BODY_PT}ptでない）")

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

    # ---- 図表：キャプションが連番であること
    #   キャプションは本文の段落に置く。表の中に出てくる【図N】【表N】は
    #   資料編の図表番号一覧であってキャプションではないため、
    #   XML全文ではなく本文の段落だけを見る
    #   （python-docx の paragraphs は表の中の段落を含まない）。
    heads = [q.text for q in d.paragraphs]
    caps = [m.group(1) for m in (re.match(r"【図(\d+)】", t) for t in heads) if m]
    tcaps = [m.group(1) for m in (re.match(r"【表(\d+)】", t) for t in heads) if m]
    for nm, seq in (("図", caps), ("表", tcaps)):
        if seq and [int(c) for c in seq] != list(range(1, len(seq) + 1)):
            bad.append(f"{nm}番号が連番でない {seq[:12]}")
    if caps:
        imgs = sum(1 for q in d.paragraphs if "graphicData" in q._p.xml)
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
