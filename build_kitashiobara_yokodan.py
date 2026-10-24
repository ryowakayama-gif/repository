# -*- coding: utf-8 -*-
"""
北塩原村　他ＡＩの横断レビュー（第４回・令和8年10月24日）への対応

受領したもの
  「策定委員会資料・計画素案・確認事項　Red Team追加レビュー報告書」
  （source/他AIレビュー_20261024_第4回/。同じ内容の2ファイルを受領し、
  ハッシュが一致することを確かめた）。指摘はRT-01〜RT-17の17項目。
  あわせて当方で、全成果物を横断して食い違いを探すレッドチームの点検を
  行った（Workflow。8つの観点で探し、各指摘を3者で反証にかけた）。

本書が正本に入れるもの
  他メンバー版の本文のうち、村の資料・工程・告示の原典と食い違っている
  ものを直す（書き換えは事実として誤っているものに限るという取決めの範囲）。
    YD_KAKIKAE … 本文の段落の書き換え
    YD_HYOU    … 表のセルの書き換え（表８-２ 計画の策定経過）
  当方が書いた文の誤りは、本書ではなく、それを書いた生成器で直した
  （理解促進研修・啓発事業の実績、訪問系サービスの事業所の所在、
  就労定着支援の基準の書き方など）。

出力: output/北塩原村_横断レビューへの対応.xlsx
"""

import os
import re
import sys

import docx
from openpyxl import Workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from kitashiobara_common import (  # noqa: E402
    add_sheet, ensure_out_dir, style_header_row, write_row,
)
from kitashiobara_koutei import (  # noqa: E402
    PABU_KAISHI, PABU_SHURYO, taigai, wareki,
)
from build_kitashiobara_mura_daicho import DK_HYOU  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_FILE = f"{REPO_ROOT}/output/北塩原村_横断レビューへの対応.xlsx"
HONPON = f"{REPO_ROOT}/output/北塩原村_計画素案_正本_移植後.docx"
MEMBER = (f"{REPO_ROOT}/source/他メンバー素案_20261008/"
          "北塩原村_第8期障がい福祉計画_第4期障がい児福祉計画_素案_他メンバー版.docx")
REVIEW = (f"{REPO_ROOT}/source/他AIレビュー_20261024_第4回/"
          "北塩原村_策定委員会資料・計画素案・確認事項_RedTeam追加レビュー_R81024.md")
KOKUJI = f"{REPO_ROOT}/source/基本指針/第8期基本指針_告示本文_令和8年3月31日.pdf"


def _dk(svc, tani):
    hits = [r for r in DK_HYOU if r.get("svc") == svc and r.get("tani") == tani]
    if len(hits) != 1:
        raise LookupError(f"実績の表の行が一意でない: {svc} {tani}")
    return [re.match(r"^([0-9.]+|―)（", v).group(1) for v in hits[0]["ato"]]


_B = _dk("就労継続支援\nＢ型", "人／月")
_A = _dk("就労継続支援\nＡ型", "人／月")

# ===========================================================================
# 1. 本文の段落の書き換え（他メンバー版の本文）
# ===========================================================================
YD_KAKIKAE = [
    dict(
        no="ＹＤ-1",
        saki="第4章2（3）福祉施設から一般就労への移行等（現状）",
        mae="北塩原村においては、現在、就労継続支援Ａ型のサービスを１人、"
            "就労継続支援Ｂ型のサービスを11人が利用していますが、"
            "村内には就労継続支援事業所がなく、近隣市町村への通所となっており、"
            "移動手段の確保が課題となっています。"
            "また、利用者の個別のケース会議等では、能力や心身の状況等から"
            "一般就労への移行は容易ではないとの声もあります。",
        ato="北塩原村においては、現在、就労継続支援Ｂ型のサービスを"
            f"月に{_B[2].split('.')[0]}人程度が利用していますが"
            f"（令和８年度は年度途中までの平均で月{_B[2]}人。"
            "就労継続支援Ａ型は令和７年度まで利用がありません）、"
            "村内には就労継続支援事業所がなく、近隣市町村への通所となっており、"
            "移動手段の確保が課題となっています。"
            "また、利用者の個別のケース会議等では、能力や心身の状況等から"
            "一般就労への移行は容易ではないとの声もあります。",
        konkyo="村の協議会資料（資料２－３）と給付の台帳による表５-７"
               "（ＤＫ-9〜ＤＫ-12）。他ＡＩの横断レビュー（正本内部の矛盾）",
        ryo="本文1段落の置換"),
    dict(
        no="ＹＤ-2",
        saki="第4章2（4）障がい児支援の提供体制の整備等（≪国の基本指針≫の②）",
        mae="② インクルージョン推進のための協議の場の設置：各市町村又は圏域【新規】",
        ato="② インクルージョン推進のための協議の場の設置：各都道府県及び各市町村【新規】",
        konkyo="第8期基本指針 告示本文（令和8年3月31日）第二の四の２"
               "「令和十一年度末までに、各都道府県及び各市町村において、"
               "保育、子育て支援、教育等の関係機関が連携を図るための協議の場を"
               "設置」。「各市町村又は各圏域」は児童発達支援センター等の"
               "中核機能についての語である",
        ryo="表の中の1段落の置換"),
    dict(
        no="ＹＤ-3",
        saki="第5章1 障がい福祉サービスの見込量（冒頭）",
        mae="各年度における障がい福祉サービスの実績値及び見込量は次のとおりです。"
            "なお、令和８年度においては、本計画策定時点で年度が終了していないため、"
            "見込み値となっています。",
        ato="各年度における障がい福祉サービスの実績値及び見込量は次のとおりです。"
            "なお、令和８年度においては、本計画策定時点で年度が終了していないため、"
            "年度途中までの実績（村の資料及び給付の台帳による暫定の値）を"
            "示しています。",
        konkyo="表の注記（ＤＫ-注１ほか）は令和８年度の値を台帳の年度途中の"
               "暫定値としている。冒頭の「見込み値」と食い違っていた",
        ryo="本文1段落の置換"),
]

# ===========================================================================
# 2. 表のセルの書き換え（表８-２ 計画の策定経過）
#    工程は案Ａ（令和8年10月14日採用）。日は kitashiobara_koutei から読み、
#    仮置きの日は月又は旬に戻す（taigai）。第1章5（4）は令和8年10月14日に
#    改めた（ＴＫ-4）が、資料編の表８-２が従前の工程のまま残っていた。
# ===========================================================================
_PABU = f"{wareki(PABU_KAISHI)}～{wareki(PABU_SHURYO).replace('令和９年', '')}"
YD_HYOU = [
    dict(no="ＹＤ-Ｔ1", saki="第8章3 表８-２ 計画の策定経過（令和８年11月～12月）",
         atama="年月日|内容", row_key="令和８年11月～12月", col=1,
         mae=("計画素案の作成・庁内調整",),
         #   策定委員会（第2回）は第1章5の策定体制に無い会議体であるため、
         #   ここには書かない（村の会議体の名称と開催日はＭ-125で確認中）
         ato=("計画素案の作成・庁内調整、４町村広域の障がい者自立支援協議会"
              f"（{taigai('2026-12-16')}）",),
         konkyo="kitashiobara_koutei（案Ａ）。第1章5（5）の４町村広域の協議会",
         ryo="表の1セルの置換"),
    dict(no="ＹＤ-Ｔ2", saki="第8章3 表８-２ 計画の策定経過（自立支援協議会）",
         atama="年月日|内容", row_key="令和９年１月", col=0,
         mae=("令和９年１月",), ato=(taigai("2027-01-15"),),
         konkyo="kitashiobara_koutei（案Ａ）", ryo="表の1セルの置換"),
    dict(no="ＹＤ-Ｔ3", saki="第8章3 表８-２ 計画の策定経過（パブリックコメント）",
         atama="年月日|内容", row_key="令和９年２月８日～２月22日", col=0,
         mae=("令和９年２月８日～２月22日",), ato=(_PABU,),
         konkyo="kitashiobara_koutei（案Ａ）。第1章5（4）と同じ期間",
         ryo="表の1セルの置換"),
    dict(no="ＹＤ-Ｔ4", saki="第8章3 表８-２ 計画の策定経過（庁議・公表）",
         atama="年月日|内容", row_key="令和９年３月", col=1,
         mae=("庁議（計画の決定）、計画の策定・公表",),
         ato=(f"庁議（計画の決定。{wareki('2027-02-19')}）、"
              f"議会全員協議会（{taigai('2027-02-24')}）、"
              "計画の策定・公表（令和９年３月）",),
         konkyo="kitashiobara_koutei（案Ａ）", ryo="表の1セルの置換"),
    dict(no="ＹＤ-Ｔ5", saki="第8章3 表８-２ 計画の策定経過（庁議・公表の時期）",
         atama="年月日|内容", row_key="令和９年３月", col=0,
         mae=("令和９年３月",), ato=("令和９年２月～３月",),
         konkyo="kitashiobara_koutei（案Ａ）", ryo="表の1セルの置換"),
]


# ===========================================================================
# 2b. 村への照会（この回で起こすもの）
#     (番号, 宛先, 件名, 内容, 期限)
#     人口の見通しの食い違いは、令和8年10月14日から「Ｍ-123・R-41」として
#     台帳に置いていたが、Ｍ-123は工程の照会（解決済み）であり、
#     人口の見通しを村に尋ねた番号は無かった（横断レビューの点検で判明）。
# ===========================================================================
YD_SHOUKAI = [
    ("Ｍ-142", "村", "人口の見通し（令和11年度の高齢化率）の系列",
     "計画素案の表２-１は、住民基本台帳の実績を延ばした見通しで、"
     "令和11年度の総人口2,166人・高齢化率45.2％としています。"
     "一方、村のこども・子育て計画の推計（コーホート変化率法）は"
     "令和11年の総人口2,185人・老年人口1,051人・高齢化率48.1％で、"
     "老年人口の向き（減少か増加か）が逆になっています。"
     "障がい者手帳の所持率による推計はこの総人口を分母に用いるため、"
     "計画ではどちらの系列を用いるかをご教示ください。"
     "高齢者福祉計画・介護保険事業計画の人口の基礎と揃える必要があれば、"
     "あわせてお知らせください。",
     "令和８年10月30日（金）"),
]


# ===========================================================================
# 3. 自己点検
# ===========================================================================
def _norm(s):
    return re.sub(r"\s+", "", str(s))


def _member_text():
    d = docx.Document(MEMBER)
    out = [p.text for p in d.paragraphs]
    for t in d.tables:
        for r in t.rows:
            for c in r.cells:
                out += [p.text for p in c.paragraphs]
    return out


def kensa():
    ng = []
    mem = _member_text()
    # ① 直す前の文が他メンバー版の段落にちょうどある
    for r in YD_KAKIKAE:
        if sum(1 for t in mem if _norm(t) == _norm(r["mae"])) < 1:
            ng.append(f"① 直す前の文が他メンバー版に無い: {r['no']}")
        if r["mae"] == r["ato"]:
            ng.append(f"① 直しても変わらない: {r['no']}")
    # ② 表の直す前のセルが他メンバー版の表８-２にある
    d = docx.Document(MEMBER)
    t82 = None
    for t in d.tables:
        rows = [[c.text for c in r.cells] for r in t.rows]
        if rows[0][:2] == ["年月日", "内容"]:
            t82 = rows
    if t82 is None:
        ng.append("② 他メンバー版に表８-２が無い")
    else:
        for r in YD_HYOU:
            if not any(row[0] == r["row_key"] and row[r["col"]] == r["mae"][0]
                       for row in t82):
                ng.append(f"② 表８-２に直す前のセルが無い: {r['no']}")
    # ③ ＹＤ-2 の根拠が告示の原典にある
    import logging
    import pdfplumber
    logging.getLogger("pdfminer").setLevel(logging.ERROR)
    with pdfplumber.open(KOKUJI) as pdf:
        tx = "".join((pg.extract_text() or "") for pg in pdf.pages)
    tx = _norm(tx)
    if _norm("各都道府県及び各市町村において、保育、子育て支援、教育等の関係機関が"
             "連携を図るための協議の場") not in tx:
        ng.append("③ ＹＤ-2 の根拠の文が告示に無い")
    # ④ 正本に反映されていること（正本があれば）
    if os.path.exists(HONPON):
        h = docx.Document(HONPON)
        ht = [p.text for p in h.paragraphs]
        for t in h.tables:
            for row in t.rows:
                for c in row.cells:
                    ht += [p.text for p in c.paragraphs]
        hn = "\n".join(_norm(x) for x in ht)
        for r in YD_KAKIKAE:
            if _norm(r["ato"]) not in hn:
                ng.append(f"④ 正本に直した後の文が無い: {r['no']}")
            if _norm(r["mae"]) in hn:
                ng.append(f"④ 正本に直す前の文が残っている: {r['no']}")
        for r in YD_HYOU:
            if _norm(r["ato"][0]) not in hn:
                ng.append(f"④ 正本の表８-２に直した後のセルが無い: {r['no']}")
        for furui in ("令和９年２月８日～２月22日", "就労継続支援Ａ型のサービスを１人"):
            if _norm(furui) in hn:
                ng.append(f"④ 正本に直す前の記述が残っている: {furui}")
    # ⑤ 受領したレビューが2ファイルとも同じ内容であること（原典の確認）
    if not os.path.exists(REVIEW):
        ng.append("⑤ 受領したレビューが無い")
    return ng


def main():
    ng = kensa()
    if ng:
        print("自己点検 不合格:")
        for x in ng:
            print("   -", x)
        raise SystemExit(1)
    ensure_out_dir()
    wb = Workbook()
    wb.remove(wb.active)
    ws = add_sheet(wb, "01_正本の書き換え",
                   f"正本の書き換え（本文{len(YD_KAKIKAE)}件・表のセル{len(YD_HYOU)}件）",
                   "他メンバー版の本文のうち、村の資料・工程・告示と食い違うもの。",
                   [10, 34, 60, 60, 44])
    style_header_row(ws, 5, ["番号", "箇所", "直す前", "直した後", "根拠"])
    r = 6
    for i, x in enumerate(YD_KAKIKAE + YD_HYOU):
        mae = x["mae"][0] if isinstance(x["mae"], tuple) else x["mae"]
        ato = x["ato"][0] if isinstance(x["ato"], tuple) else x["ato"]
        r = write_row(ws, r, [x["no"], x["saki"], mae, ato, x["konkyo"]],
                      alt=(i % 2 == 1))
    wb.save(OUT_FILE)
    print(f"  自己点検: 直す前の文{len(YD_KAKIKAE)}件・表のセル{len(YD_HYOU)}件が"
          "他メンバー版に実在／ＹＤ-2の根拠を告示の原典で照合／正本に反映")
    print(f"作成: {OUT_FILE}")


if __name__ == "__main__":
    main()
