"""アンケートの回収数の変更の影響ブック ジェネレータ

ご指示（令和8年10月21日）
  「アンケートの回収数の変更から添付ファイルを確認してください。
   合わせて計画草案への影響範囲についてご確認お願いします」

受領したもの
  報告書用集計表（令和8年10月21日版。15シート）
  アンケート調査結果報告書（同版。2,141段落）
  いずれも source/アンケート結果 に収めた。

何が変わったか
  障がい者調査の有効回答数が38件から40件になった。
  有効回答率は22.1％から23.3％である。
  これに伴い、集計_障がい者シートの449行のうち446行が動いた
  （件数そのものと、回答者数を基数とする割合の両方）。

いちばん大事なこと ― 結論は動かない
  設問ごとに最も多い選択肢を新旧で突き合わせたところ、
  入れ替わった設問は0件であった。
  課題の8区分（6章_課題）も1行も動いていない。
  障がい児調査（n＝3）と障がい別の集計も動いていない。
  動いたのは数値だけである。

  障がい別が動かない理由も現物で確かめた。
  増えた2件はいずれも身体・療育・精神のどの手帳も持たない方
  （自立支援医療受給者証1件、手帳等は持っていないが障がい福祉
  サービスを利用している1件）であり、障がい別の集計は手帳で
  区分しているため影響しない。新しい集計表は内的に整合している。

誰が直すかを分ける
  正本で割合を含む段落・セルを、他メンバー版の現物にそのまま
  あるものと、当方が足した・書き換えたものに分ける。
  前者は他メンバーが素案を作り直す範囲であり、
  後者は当方が直す範囲である。
  この切り分けは、当方が他メンバー版の本文を大量に書き換えないため
  （書き換えは事実として誤っているものに限るという取決め）に要る。

出力: output/北塩原村_アンケート回収数の変更の影響.xlsx
"""

import os
import re
import sys
import warnings

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from kitashiobara_common import (
    COLORS, FONT, OUT_DIR, SRC_DIR, add_sheet, ensure_out_dir,
    style_header_row, style_title, write_row,
)

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

OUT_FILE = f"{OUT_DIR}/北塩原村_アンケート回収数の変更の影響.xlsx"
KYU = f"{SRC_DIR}/アンケート結果/報告書用集計表.xlsx"
SHIN = (f"{SRC_DIR}/アンケート結果/"
        "報告書用集計表_令和8年10月21日版_回収40.xlsx")
HOUKOKU = (f"{SRC_DIR}/アンケート結果/"
           "アンケート調査結果報告書_令和8年10月21日版.docx")
MEMBER = (f"{SRC_DIR}/他メンバー素案_20261008/"
          "北塩原村_第8期障がい福祉計画_第4期障がい児福祉計画_"
          "素案_他メンバー版.docx")
HONPON = f"{OUT_DIR}/北塩原村_計画素案_正本_移植後.docx"
KIJUNBI = "基準日：令和8年10月21日"

SH = "集計_障がい者"
KARE = "他メンバーが直す"
WARE = "当方が直す"


# ===========================================================================
# 1. 集計表を読む
# ===========================================================================
def tally(path, sheet=SH):
    """(問, 選択肢) → (件数, 割合, 回答者数)"""
    wb = load_workbook(path, read_only=True, data_only=True)
    out = {}
    for r in wb[sheet].iter_rows(values_only=True):
        if not r or not r[0] or not isinstance(r[3], (int, float)):
            continue
        out[(str(r[0]).strip(), str(r[2]).strip())] = (
            r[3], r[4], r[5] if len(r) > 5 else None)
    wb.close()
    return out


def _kazu(v):
    """集計表の値は文字列で入っていることがある。数に直す。"""
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return v


def gaiyou(path):
    """表1-1 調査の概要から、配布数・有効回答数・有効回答率・対象の内訳。

    項目名の列は左に空の列があるため固定の位置で読まない。
    行の中から項目名を探し、その右の2つ（障がい者・障がい児）を取る。
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["1-2章_概要・要点"]
    out = {}
    for r in ws.iter_rows(min_row=1, max_row=12, values_only=True):
        if not r:
            continue
        for i, v in enumerate(r):
            if v is None:
                continue
            key = str(v).strip()
            if key in ("調査対象", "配布数", "有効回答数", "有効回答率"):
                ato = [x for x in r[i + 1:] if x is not None]
                if len(ato) >= 2:
                    out[key] = (_kazu(ato[0]), _kazu(ato[1]))
                elif ato:
                    out[key] = (_kazu(ato[0]), None)
                break
    wb.close()
    return out


def pc(v):
    return f"{round(v * 100, 1):.1f}％" if isinstance(v, (int, float)) \
        else "―"


def sheet_rows(path, sheet):
    wb = load_workbook(path, read_only=True, data_only=True)
    g = [[x for x in r] for r in wb[sheet].iter_rows(values_only=True)]
    wb.close()
    return g


# ===========================================================================
# 2. 正本と他メンバー版を読む
# ===========================================================================
def paras(path):
    import docx
    d = docx.Document(path)
    out = [p.text for p in d.paragraphs]
    for tb in d.tables:
        for r in tb.rows:
            for c in r.cells:
                out.append(c.text)
    return out


def _norm(s):
    return "".join(str(s).split())


PAT = re.compile(r"[0-9]+\.[0-9]％")


def wakeru():
    """正本で割合を含む段落・セルを、他メンバー版由来と当方由来に分ける。"""
    mem = set(_norm(x) for x in paras(MEMBER))
    kare, ware = [], []
    for p in paras(HONPON):
        if not PAT.search(p):
            continue
        (kare if _norm(p) in mem else ware).append(p)
    return kare, ware


# ===========================================================================
# 3. 当方が直す記述
#    (番号, 箇所, いまの文の要点, 直す内容, 旧→新)
#    値は集計表から読む。ここに数値を書き写さない。
# ===========================================================================
WARE_NAOSU = [
    ("ＡＫ-2", "第2章7（5）調査結果の概要（当方の注記）",
     "有効回答率22.1％、回答者の65.8％が65歳以上、"
     "68.4％が身体障がい",
     "３つとも新しい値に改める",
     [("―", "有効回答率"), ("問2", "65歳以上"), ("問6", "身体障がい")]),
    ("ＡＫ-3", "第2章7 住まいについての読み方（当方の追記）",
     "「今の自宅で暮らしたい」39.5％と「家族と一緒に暮らしたい」"
     "15.8％を合わせた55.3％が在宅を希望",
     "３つとも新しい値に改める。合計は足し算で出す",
     [("問11", "今の自宅で暮らしたい"), ("問11", "家族と一緒に暮らしたい")]),
    ("ＡＫ-4", "第2章7 施設入所支援の見込量との関係（当方の追記）",
     "今後３年間に施設入所支援を挙げた方が23.7％",
     "新しい値に改める。23.7％は２カ所に出る",
     [("問27-3", "施設入所支援")]),
    ("ＡＫ-5", "第3章3 個別避難計画の指標（当方の追記）",
     "名簿及び個別避難計画を「知らない」と回答した方が68.4％",
     "新しい値に改める",
     [("問39", "知らない")]),
    ("ＡＫ-6", "第4章2（2）退院の見込みを置く根拠（当方の追記）",
     "在宅での生活の継続を望む回答が55.3％"
     "（精神障がいのある方では50.0％）",
     "合計を新しい値に改める。精神障がいのある方の50.0％は"
     "障がい別の集計によるもので動かない",
     [("問11", "今の自宅で暮らしたい"), ("問11", "家族と一緒に暮らしたい")]),
    ("ＡＫ-7", "第2章7（6）表2-8 課題の8区分目（ＡＫ-1で足した行）",
     "定期的に通院している方が71.1％",
     "新しい値に改める",
     [("問15", "通院している")]),
    ("ＡＫ-8", "第2章7 調査結果と実績の食い違い（当方の追記）",
     "外出に困っている方が34.2％、制度をよく知らない方が23.7％",
     "２つとも新しい値に改める",
     [("問14", "冬季・積雪時に外出しにくい"),
      ("問27-2", "重度訪問介護")]),
    ("ＡＫ-9", "第2章7 村が力を入れるべきこと（当方の追記）",
     "相談支援体制と災害時支援が各34.2％、"
     "移動支援とサービス事業所の確保が各26.3％",
     "値を改めるだけでなく、言い回しを改める必要がある。"
     "新しい集計では相談支援体制と災害時支援が同率でなくなり、"
     "相談支援体制が単独で最も多い",
     [("問44", "相談支援体制"), ("問44", "災害時支援"),
      ("問44", "移動支援"), ("問44", "サービス事業所の確保")]),
]


# ===========================================================================
# 4. 変わらないもの（確かめた結果）
# ===========================================================================
def kawaranai():
    a, b = tally(KYU), tally(SHIN)
    #   設問ごとに最も多い選択肢が入れ替わるか
    byq_a, byq_b = {}, {}
    for (q, sel), v in a.items():
        if sel != "無回答":
            byq_a.setdefault(q, []).append((sel, v[0]))
    for (q, sel), v in b.items():
        if sel != "無回答":
            byq_b.setdefault(q, []).append((sel, v[0]))
    irekawari = []
    for q in byq_a:
        if q not in byq_b:
            continue
        ta = sorted(byq_a[q], key=lambda x: -x[1])[0][0]
        tb = sorted(byq_b[q], key=lambda x: -x[1])[0][0]
        if ta != tb:
            irekawari.append((q, ta, tb))
    #   シートごとの差のある行数
    sa = {}
    for sh in ("6章_課題", "集計_障がい児", "集計_障がい別"):
        ga, gb = sheet_rows(KYU, sh), sheet_rows(SHIN, sh)
        n = sum(1 for i in range(max(len(ga), len(gb)))
                if (ga[i] if i < len(ga) else None)
                != (gb[i] if i < len(gb) else None))
        sa[sh] = n
    return irekawari, sa


def fueta_nimai():
    """増えた2件がどの選択肢に入ったか。"""
    a, b = tally(KYU), tally(SHIN)
    out = []
    for k in sorted(b):
        if k not in a:
            out.append((k[0], k[1], 0, b[k][0]))
        elif b[k][0] != a[k][0]:
            out.append((k[0], k[1], a[k][0], b[k][0]))
    return out


# ===========================================================================
# シート
# ===========================================================================
def sheet_ketsuron(wb):
    ga, gb = gaiyou(KYU), gaiyou(SHIN)
    kare, ware = wakeru()
    irekawari, sa = kawaranai()
    a, b = tally(KYU), tally(SHIN)
    ugoita = sum(1 for k in a if k in b and pc(a[k][1]) != pc(b[k][1]))
    ws = add_sheet(
        wb, "00_結論", "アンケートの回収数の変更の影響",
        "障がい者調査の有効回答数が38件から40件になりました。"
        "集計の件数そのものが動いているため、割合は分母だけでなく"
        "分子も変わります。"
        "ただし設問ごとに最も多い選択肢は1件も入れ替わっておらず、"
        "課題の8区分も動いていません。"
        f"動いたのは数値だけです。{KIJUNBI}。",
        [28, 26, 26, 64])
    r = 5
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    style_title(ws.cell(row=r, column=1), "調査の概要の新旧",
                fill=COLORS["subhead"], size=11)
    r += 1
    style_header_row(ws, r, ["項目", "旧（回収38）", "新（回収40）", "差"])
    r += 1
    for key in ("配布数", "有効回答数", "有効回答率", "調査対象"):
        va = ga.get(key, ("―", "―"))[0]
        vb = gb.get(key, ("―", "―"))[0]
        if key == "有効回答率":
            va, vb = pc(va), pc(vb)
        sa_ = "変わらない" if str(va) == str(vb) else "変わる"
        write_row(ws, r, [f"{key}（障がい者調査）", str(va)[:60],
                          str(vb)[:60], sa_],
                  aligns=["left", "left", "left", "center"])
        if sa_ == "変わる":
            ws.cell(row=r, column=3).fill = PatternFill(
                "solid", fgColor="FFF2CC")
        ws.row_dimensions[r].height = 40 if key == "調査対象" else 20
        r += 1
    r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    style_title(ws.cell(row=r, column=1), "変わるもの",
                fill=COLORS["subhead"], size=11)
    r += 1
    style_header_row(ws, r, ["", "件数", "", "内容"])
    r += 1
    for nm, n, naka in (
            ("集計_障がい者の行",
             f"{sum(1 for i in range(len(sheet_rows(SHIN, SH))) if True)}行中"
             f"{ugoita}組",
             "(問, 選択肢) の組のうち割合が動くもの。件数そのものも動く"),
            ("正本で割合を含む段落・セル", f"{len(kare) + len(ware)}",
             f"他メンバー版の現物にあるもの{len(kare)}／"
             f"当方が足した・書き換えたもの{len(ware)}"),
            ("当方が直す記述", f"{len(WARE_NAOSU)}",
             "うち1件（ＡＫ-9）は値だけでなく言い回しを改める必要がある"),
    ):
        write_row(ws, r, [nm, n, "", naka],
                  aligns=["left", "center", "left", "left"])
        ws.row_dimensions[r].height = 24
        r += 1
    r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    style_title(ws.cell(row=r, column=1), "変わらないもの（確かめた結果）",
                fill=COLORS["subhead"], size=11)
    r += 1
    style_header_row(ws, r, ["", "結果", "", "なぜ変わらないか"])
    r += 1
    for nm, kekka, naze in (
            ("設問ごとに最も多い選択肢", f"入れ替わり{len(irekawari)}件",
             "増えた2件は特定の選択肢に偏っておらず、順位を動かさない。"
             "「最も高い」「最多」といった本文の言い回しはそのままでよい"),
            ("課題の8区分（6章_課題）", f"差のある行{sa['6章_課題']}行",
             "表2-8 の8区分（ＡＫ-1で足した専門的支援を含む）は動かない"),
            ("障がい児調査", f"差のある行{sa['集計_障がい児']}行",
             "有効回答数3件のまま。差の1行は入力値の色の説明文のみ"),
            ("障がい別の集計（表2-7）", f"差のある行{sa['集計_障がい別']}行",
             "増えた2件はいずれも身体・療育・精神のどの手帳も持たない方"
             "（自立支援医療受給者証1件、手帳等は持っていないが"
             "障がい福祉サービスを利用している1件）であり、"
             "手帳で区分する障がい別の集計に入らない"),
    ):
        write_row(ws, r, [nm, kekka, "", naze],
                  aligns=["left", "center", "left", "left"])
        ws.cell(row=r, column=2).fill = PatternFill(
            "solid", fgColor="E2EFDA")
        ws.row_dimensions[r].height = 40
        r += 1
    r += 2

    ws.merge_cells(start_row=r, start_column=1, end_row=r + 4, end_column=4)
    c = ws.cell(row=r, column=1)
    c.value = (
        "※ 調査対象の内訳が「身体118名・精神1名・療育8名・入所等12名」"
        "（合計139名）から「身体118名・精神34名・療育8名・入所等12名」"
        "（合計172名）に改められ、配布数172と一致しました。"
        "当方がＭ-128でお尋ねしていた食い違い"
        "（内訳139名と配布数172の差33名、精神1名の数え方）は"
        "これで解消します。"
        "増えた2件が自立支援医療受給者証の方と手帳を持たずサービスを"
        "利用している方であることは、Ｍ-128の①"
        "（内訳の4区分のほかに対象とした方がいるか）のお答えにも"
        "当たります。")
    c.alignment = Alignment(wrap_text=True, vertical="top")
    c.font = Font(name=FONT, size=9)
    return ws


def sheet_nimai(wb):
    rows = fueta_nimai()
    ws = add_sheet(
        wb, "01_増えた2件の入り方",
        f"件数が動いた選択肢 {len(rows)}件",
        "増えた2件がどの選択肢に入ったかです。"
        "設問ごとの件数の合計は38から40へ2件増えています"
        "（複数回答の設問を除く）。"
        "この2件は身体・療育・精神のどの手帳も持たない方であり、"
        "障がい別の集計には入りません。",
        [12, 44, 12, 12, 12])
    style_header_row(ws, 5, ["問", "選択肢", "旧", "新", "差"])
    r = 6
    for i, (q, sel, ka, kb) in enumerate(rows):
        write_row(ws, r, [q, sel, ka, kb, f"{kb - ka:+d}"],
                  alt=(i % 2 == 1),
                  aligns=["center", "left", "center", "center", "center"])
        ws.row_dimensions[r].height = 18
        r += 1
    return ws


def sheet_ware(wb):
    a, b = tally(KYU), tally(SHIN)
    ga, gb = gaiyou(KYU), gaiyou(SHIN)
    ws = add_sheet(
        wb, "02_当方が直す分",
        f"{len(WARE_NAOSU)}件",
        "正本のうち、当方が足した・書き換えた記述で"
        "回収数の変更の影響を受けるものです。"
        "新しい値は集計表から読んでおり、このブックで書き写して"
        "いません。",
        [10, 34, 44, 44, 44])
    style_header_row(ws, 5, ["番号", "箇所", "いまの文の要点",
                             "直す内容", "旧→新"])
    r = 6
    for i, (no, saki, ima, naosu, keys) in enumerate(WARE_NAOSU):
        shin = []
        for q, sel in keys:
            if q == "―":
                shin.append(f"有効回答率 {pc(ga['有効回答率'][0])}"
                            f"→{pc(gb['有効回答率'][0])}")
                continue
            ka, kb = a.get((q, sel)), b.get((q, sel))
            if ka and kb:
                shin.append(f"{q} {sel} {pc(ka[1])}→{pc(kb[1])}"
                            f"（{ka[0]}人→{kb[0]}人）")
            else:
                shin.append(f"{q} {sel} ―")
        write_row(ws, r, [no, saki, ima, naosu, "／".join(shin)],
                  alt=(i % 2 == 1), aligns=["center"] + ["left"] * 4)
        if no == "ＡＫ-9":
            ws.cell(row=r, column=4).fill = PatternFill(
                "solid", fgColor="FCE4E4")
        ws.row_dimensions[r].height = 56
        r += 1
    return ws


def sheet_kare(wb):
    kare, _ware = wakeru()
    a, b = tally(KYU), tally(SHIN)
    ugoku = {pc(a[k][1]) for k in a if k in b and pc(a[k][1]) != pc(b[k][1])}
    atari = [p for p in kare
             if any(w in ugoku for w in PAT.findall(p))]
    ws = add_sheet(
        wb, "03_他メンバーが直す分",
        f"他メンバー版の現物にある段落・セル {len(kare)}件"
        f"（うち値が動くもの {len(atari)}件）",
        "正本 第2章7 の割合は、他メンバーが作成された素案の本文に"
        "あります。回収数が変わったことによる差し替えは、"
        "素案を作り直していただく範囲です。"
        "当方が本文を書き換えるのは事実として誤っているものに限る"
        "という取決めによるものです（Ｓ-39でお願いします）。",
        [14, 120])
    style_header_row(ws, 5, ["動く割合", "段落・セル（先頭120字）"])
    r = 6
    for i, p in enumerate(atari):
        w = sorted(set(x for x in PAT.findall(p) if x in ugoku))
        write_row(ws, r, ["・".join(w), p.replace("\n", " ")[:120]],
                  alt=(i % 2 == 1), aligns=["left", "left"])
        ws.row_dimensions[r].height = 22
        r += 1
    return ws


def sheet_shoukai(wb):
    ws = add_sheet(
        wb, "04_照会", "2件",
        "Ｍ-128 は解消しました。"
        "新たにＳ-39（他メンバーへ）を起こします。",
        [10, 20, 34, 76, 16])
    style_header_row(ws, 5, ["番号", "相手", "件名", "内容", "期限"])
    r = 6
    for i, rec in enumerate(SHOUKAI):
        write_row(ws, r, list(rec), alt=(i % 2 == 1),
                  aligns=["center", "left", "left", "left", "center"])
        ws.row_dimensions[r].height = 110
        r += 1
    return ws


SHOUKAI = [
    ("Ｓ-39", "他メンバー",
     "回収数40件への差し替え（第2章7の割合）",
     "報告書用集計表の令和8年10月21日版（有効回答数40件）を"
     "拝見しました。設問ごとに最も多い選択肢は1件も入れ替わって"
     "おらず、課題の8区分・障がい児調査・障がい別の集計も"
     "動いていません。動いたのは数値だけです。"
     "正本 第2章7 の割合は素案の本文にあるものですので、"
     "新しい集計に合わせて差し替えていただけますでしょうか。"
     "当方が足した記述（9件）は当方で直します。"
     "なお、村が力を入れるべきこと（問44）は、"
     "新しい集計では相談支援体制37.5％・災害時支援35.0％となり"
     "同率でなくなります。「各34.2％」という書き方をされている"
     "箇所は言い回しの修正が要ります。"
     "なお、差し替えは報告書の直し残しの修正（Ｓ-40）の後に"
     "お願いいたします。直し残し11か所のうち4か所は、"
     "同じ文が正本の第2章7にあるためです"
     "（「北塩原村_報告書の修正方針と計画への影響.xlsx」の05シート）。",
     "令和８年11月14日（金）"),
    ("Ｍ-128", "村（解消）",
     "アンケート調査の対象者の内訳と配布数（解消）",
     "調査対象の内訳が「身体118名・精神1名・療育8名・入所等12名」"
     "（合計139名）から「身体118名・精神34名・療育8名・入所等12名」"
     "（合計172名）に改められ、配布数172と一致しました。"
     "精神1名の数え方と、内訳139名と配布数172の差33名について"
     "お尋ねしていた事項は解消します。"
     "また、増えた2件が自立支援医療受給者証の方と、手帳等は"
     "持たないが障がい福祉サービスを利用している方であることは、"
     "①（内訳の4区分のほかに対象とした方がいるか）のお答えにも"
     "当たります。入所等12名の範囲（施設入所か入院を含むか）のみ"
     "残ります。",
     "解消（令和８年10月21日）"),
]


# ===========================================================================
# 自己点検
# ===========================================================================
def verify():
    ng = []
    for p, nm in ((KYU, "旧の集計表"), (SHIN, "新の集計表"),
                  (HOUKOKU, "新の報告書")):
        if not os.path.exists(p):
            ng.append(f"{nm}が無い: {p}")
    if ng:
        for x in ng:
            print(f"  × {x}")
        raise SystemExit("回収数の変更の自己点検で止めました")

    ga, gb = gaiyou(KYU), gaiyou(SHIN)

    # ① 回収数が38から40に変わっていること。
    #    旧の集計表が差し替わっていればここで気づく。
    if ga["有効回答数"][0] != 38:
        ng.append(f"旧の集計表の有効回答数が38でない: "
                  f"{ga['有効回答数'][0]}")
    if gb["有効回答数"][0] != 40:
        ng.append(f"新の集計表の有効回答数が40でない: "
                  f"{gb['有効回答数'][0]}")
    #   障がい児は3件のまま
    if ga["有効回答数"][1] != gb["有効回答数"][1]:
        ng.append("障がい児調査の有効回答数が変わっている")

    # ② 有効回答率が配布数と有効回答数から出ること（書き写していない）
    for g, nm in ((ga, "旧"), (gb, "新")):
        machi = g["有効回答数"][0] / g["配布数"][0]
        if abs(g["有効回答率"][0] - machi) > 1e-9:
            ng.append(f"{nm}の有効回答率が配布数と合わない")

    # ③ 新の調査対象の内訳の合計が配布数と一致すること（Ｍ-128の解消）
    uchi = str(gb["調査対象"][0])
    kazu = [int(x) for x in re.findall(r"（([0-9]+)名）", uchi)]
    if sum(kazu) != gb["配布数"][0]:
        ng.append(f"新の内訳の合計が配布数と合わない: "
                  f"{sum(kazu)} ≠ {gb['配布数'][0]}")
    #   旧は合わなかったこと（解消したと言える裏づけ）
    uchi_k = str(ga["調査対象"][0])
    kazu_k = [int(x) for x in re.findall(r"（([0-9]+)名）", uchi_k)]
    if sum(kazu_k) == ga["配布数"][0]:
        ng.append("旧の内訳も配布数と合っている（Ｍ-128の前提が崩れる）")

    # ④ 増えた2件が3手帳のいずれも持たないこと。
    #    障がい別が動かない理由の裏づけである。
    a, b = tally(KYU), tally(SHIN)
    for sel in ("身体障害者手帳", "療育手帳", "精神障がい者保健福祉手帳"):
        ka, kb = a.get(("問5", sel)), b.get(("問5", sel))
        if ka and kb and ka[0] != kb[0]:
            ng.append(f"問5 {sel} の件数が動いている: "
                      f"{ka[0]}→{kb[0]}（障がい別が動かない説明と矛盾）")

    # ⑤ 設問ごとに最も多い選択肢が入れ替わっていないこと。
    #    入れ替わっていれば本文の言い回しを直す必要がある。
    irekawari, sa = kawaranai()
    if irekawari:
        ng.append(f"最も多い選択肢が入れ替わる設問がある: {irekawari}")
    #   課題・障がい児・障がい別が動いていないこと（注記1行を除く）
    for sh, kagiri in (("6章_課題", 0), ("集計_障がい児", 1),
                       ("集計_障がい別", 1)):
        if sa[sh] > kagiri:
            ng.append(f"{sh} の差が想定より多い: {sa[sh]}行")

    # ⑥ 当方が直す記述が、正本の現物にあって他メンバー版に無いこと
    if os.path.exists(HONPON) and os.path.exists(MEMBER):
        kare, ware = wakeru()
        ware_t = "\n".join(ware)
        for no, saki, ima, _n, keys in WARE_NAOSU:
            #   旧の値が当方の段落に実在すること
            atta = False
            for q, sel in keys:
                if q == "―":
                    atta = atta or pc(ga["有効回答率"][0]) in ware_t
                    continue
                ka = a.get((q, sel))
                if ka and pc(ka[1]) in ware_t:
                    atta = True
            if not atta:
                ng.append(f"当方が直すとした記述の旧の値が"
                          f"当方の段落に無い: {no}（{saki}）")
            #   引いている (問, 選択肢) が集計表に実在すること
            for q, sel in keys:
                if q != "―" and (q, sel) not in b:
                    ng.append(f"集計表に無い選択肢を引いている: "
                              f"{no} {q}「{sel}」")
        if not kare:
            ng.append("他メンバー版由来の段落が0件（切り分けが働いていない)")
    else:
        print("  ※ 正本又は他メンバー版が無いため⑥は省略")

    # ⑦ ＡＫ-9 の言い回しの問題が実際に起きていること。
    #    同率でなくなることを現物で確かめる。
    s_a = a.get(("問44", "相談支援体制")), a.get(("問44", "災害時支援"))
    s_b = b.get(("問44", "相談支援体制")), b.get(("問44", "災害時支援"))
    if all(s_a) and all(s_b):
        if s_a[0][0] != s_a[1][0]:
            ng.append("旧で相談支援体制と災害時支援が同率でない"
                      "（ＡＫ-9の前提が崩れる）")
        if s_b[0][0] == s_b[1][0]:
            ng.append("新でも相談支援体制と災害時支援が同率である"
                      "（ＡＫ-9の言い回しの修正は不要になる）")
    else:
        ng.append("問44 の相談支援体制・災害時支援が集計表に無い")

    # ⑧ 報告書の本文が回収40件になっていること
    try:
        h = "\n".join(paras(HOUKOKU))
    except Exception:
        print("  ※ 報告書を読めないため⑧は省略")
    else:
        if "有効回答数はそれぞれ40件" not in _norm(h).replace("、", "") \
                and "40件" not in h:
            ng.append("新の報告書に有効回答数40件の記述がない")

    # ⑨ 他団体の名が入っていないこと
    zen = "\n".join(str(v) for rec in WARE_NAOSU for v in
                    (rec[0], rec[1], rec[2], rec[3]))
    zen += "\n" + "\n".join(str(v) for rec in SHOUKAI for v in rec)
    for hoka in ("小野町", "金ヶ崎", "金ケ崎", "阿蘇", "札幌", "川崎町"):
        if hoka in zen:
            ng.append(f"他団体の名が入っている: {hoka}")

    if ng:
        for x in ng:
            print(f"  × {x}")
        raise SystemExit("回収数の変更の自己点検で止めました")

    kare, ware = wakeru()
    ugoita = sum(1 for k in a if k in b and pc(a[k][1]) != pc(b[k][1]))
    return (f"有効回答数 {ga['有効回答数'][0]}→{gb['有効回答数'][0]}件"
            f"（有効回答率 {pc(ga['有効回答率'][0])}→"
            f"{pc(gb['有効回答率'][0])}）／"
            f"割合が動く組 {ugoita}／"
            f"最も多い選択肢の入れ替わり {len(irekawari)}件・"
            f"課題の8区分の差 {sa['6章_課題']}行／"
            f"正本の段落 他メンバー版{len(kare)}・当方{len(ware)}／"
            f"当方が直す記述 {len(WARE_NAOSU)}件／"
            f"内訳の合計 {sum(kazu)}＝配布数（Ｍ-128 解消）")


def main():
    ensure_out_dir()
    msg = verify()
    wb = Workbook()
    wb.remove(wb.active)
    sheet_ketsuron(wb)
    sheet_nimai(wb)
    sheet_ware(wb)
    sheet_kare(wb)
    sheet_shoukai(wb)
    wb.save(OUT_FILE)
    print(f"  自己検証: {msg}")
    print(f"作成: {OUT_FILE}")
    print(f"  シート{len(wb.sheetnames)}／当方が直す{len(WARE_NAOSU)}件")


if __name__ == "__main__":
    sys.exit(main())
