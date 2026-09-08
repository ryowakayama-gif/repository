# -*- coding: utf-8 -*-
"""調査票のdocxと集計仕様（shukei_data.py）を突き合わせてコードブックを作る。

コードブック１件の構造
  {"票": "ニーズ"/"在宅A"/"在宅B",
   "番号": "問1(1)",           # 調査票の設問番号
   "変数": "N_問1_1",           # データの列名（単一回答・数値・自由記述）
   "設問文": "...",
   "形式": "単一"/"複数"/"数値"/"自由"/"マトリクス",
   "区分": "必須"/"オプション"/"村独自"/"基本",
   "クロス軸": "A,B,C,D",
   "選択肢": [[1,"1人暮らし"], ...],
   "列": ["N_問1_1"] または複数回答なら選択肢ごとの列}

出力: data/集計_コードブック.json
"""
import os, sys, json, re, unicodedata

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import shukei_data as SD
from shukei_parse_docx import parse

UPLOAD = "/root/.claude/uploads/134138ca-61f7-57d3-9e9b-5f081a1a345d"
DOCX_NEEDS = os.path.join(UPLOAD, "20085979-0904__10____________________________.docx")
DOCX_ZAI = os.path.join(UPLOAD, "fdc7d3d0-0904__10____________________1.docx")
OUT = os.path.join(BASE, "data", "集計_コードブック.json")

# 在宅調査はA票とB票で設問番号が重複するため、B票の開始位置で切り分ける
ZAI_B_START = "ご家族やご親族の中で"

# 表の見出し行にしか選択肢がないマトリクス設問などを補う。
#   docxの本文からは拾えないため、調査票の表を見て手で定義する。
SANKA = [[1, "週4回以上"], [2, "週2〜3回"], [3, "週1回"],
         [4, "月1〜3回"], [5, "年に数回"], [6, "参加していない"]]
KAISU = [[1, "利用していない"], [2, "週1回程度"], [3, "週2回程度"],
         [4, "週3回程度"], [5, "週4回程度"], [6, "週5回以上"]]
UMU = [[1, "利用していない"], [2, "利用した"]]
NISSU = [[1, "利用していない"], [2, "月1〜7日程度"], [3, "月8〜14日程度"],
         [4, "月15〜21日程度"], [5, "月22日以上"]]
TSUKI = [[1, "利用していない"], [2, "月1回程度"], [3, "月2回程度"],
         [4, "月3回程度"], [5, "月4回程度"]]

OVERRIDES = {
 # ニーズ 問5(1) 会・グループへの参加頻度（8項目）
 ("ニーズ", "問5(1)①"): dict(選択肢=SANKA, 短縮="ボランティアのグループ"),
 ("ニーズ", "問5(1)②"): dict(選択肢=SANKA, 短縮="スポーツ関係のグループやクラブ"),
 ("ニーズ", "問5(1)③"): dict(選択肢=SANKA, 短縮="趣味関係のグループ"),
 ("ニーズ", "問5(1)④"): dict(選択肢=SANKA, 短縮="学習・教養サークル"),
 ("ニーズ", "問5(1)⑤"): dict(選択肢=SANKA, 短縮="介護予防のための通いの場"),
 ("ニーズ", "問5(1)⑥"): dict(選択肢=SANKA, 短縮="老人クラブ"),
 ("ニーズ", "問5(1)⑦"): dict(選択肢=SANKA, 短縮="町内会・自治会"),
 ("ニーズ", "問5(1)⑧"): dict(選択肢=SANKA, 短縮="収入のある仕事"),
 # ニーズ 問8(2) 主観的幸福感（0〜10点）
 ("ニーズ", "問8(2)"): dict(形式="数値", 数値範囲=[0, 10], 短縮="主観的幸福感（点）"),
 # ニーズ 問3(1) 身長・体重
 ("ニーズ", "問3(1)"): dict(形式="数値", 数値範囲=None, 短縮="身長・体重"),
}

# 在宅A 問9（介護保険サービスの利用状況）はサービスごとに選択肢が異なる
ZAI_Q9 = [
 ("A", "訪問介護", KAISU), ("B", "訪問入浴介護", KAISU), ("C", "訪問看護", KAISU),
 ("D", "訪問リハビリテーション", KAISU), ("E", "通所介護", KAISU),
 ("F", "通所リハビリテーション", KAISU), ("G", "夜間対応型訪問介護", KAISU),
 ("H", "定期巡回・随時対応型訪問介護看護", UMU), ("I", "小規模多機能型居宅介護", UMU),
 ("J", "看護小規模多機能型居宅介護", UMU),
 ("K", "ショートステイ", NISSU), ("L", "居宅療養管理指導", TSUKI),
]

# 第2稿の番号の誤り（doc27）。データ設計上は正しい番号に読み替える。
ZAI_RENUM = {
    # A票の問12が2つあり、後者（訪問診療）は問13が正しい
    ("在宅A", "問12", "現在、訪問診療"): "問13",
}


def varname(prefix, no):
    """設問番号を列名に変換する。問1(2)① -> N_問1_2_1"""
    s = no
    for a, b in (("(", "_"), (")", ""), ("（", "_"), ("）", "")):
        s = s.replace(a, b)
    s = re.sub(r"[①-⑳]", lambda m: "_" + str(ord(m.group()) - ord("①") + 1), s)
    s = s.replace("-", "_")
    return f"{prefix}_{s}"


def guess_format(item, spec_fmt):
    if spec_fmt:
        return spec_fmt
    t = item["設問文"]
    if "いくつでも" in t or "複数選択可" in t or "複数回答可" in t or "３つまで" in t:
        return "複数"
    if not item["選択肢"]:
        return "自由"
    return "単一"


def build():
    # ── 集計仕様の索引 ─────────────────────────
    spec_n = {}
    for (_hyo, q, e, text, kb, fm, sh, cr, mr, note) in SD.N:
        spec_n[q + e] = dict(区分=kb, 形式=fm, 集計方法=sh, クロス軸=cr, 見える化=mr, 備考=note)
    spec_z = {}
    for (_hyo, hyo, q, text, kb, fm, sh, cr, mr, note) in SD.Z:
        spec_z[(hyo, q)] = dict(区分=kb, 形式=fm, 集計方法=sh, クロス軸=cr, 見える化=mr, 備考=note)

    book = []

    # ── ニーズ調査 ─────────────────────────
    parent = None
    for it in parse(DOCX_NEEDS):
        if not it["設問文"] and not it["選択肢"]:
            continue                      # 「問1」だけの見出し行
        sp = spec_n.get(it["番号"], {})
        if not sp and parent and it["番号"].startswith(parent[0]):
            # マトリクス設問の行（問5(1)①〜⑧）は親の仕様を引き継ぐ
            sp = dict(parent[1], 形式="単一")
        fmt = guess_format(it, sp.get("形式"))
        rec = make("ニーズ", "N", it, fmt, sp)
        if fmt == "マトリクス":
            parent = (it["番号"], sp)
            rec["列"] = []                # 親自体はデータ列を持たない
        book.append(rec)

    # ── 在宅介護実態調査 ───────────────────────
    hyo = "在宅A"
    for it in parse(DOCX_ZAI):
        if ZAI_B_START in it["設問文"]:
            hyo = "在宅B"
        if not it["設問文"] and not it["選択肢"]:
            continue
        no = it["番号"]
        if hyo == "在宅A" and no == "問6" and it["設問文"].lstrip().startswith("-1"):
            no = "問6-1"
            it = dict(it, 設問文=it["設問文"].lstrip()[2:].strip())
        for (h, n0, marker), n1 in ZAI_RENUM.items():
            if h == hyo and no == n0 and marker in it["設問文"]:
                no = n1
        it = dict(it, 番号=no)
        sp = spec_z.get(("A票" if hyo == "在宅A" else "B票", no), {})
        if not sp and not it["選択肢"]:
            sp = {"区分": "村独自", "形式": "自由", "集計方法": "自由記述の分類",
                  "クロス軸": "A", "見える化": "－", "備考": "自由記述欄"}
        fmt = guess_format(it, sp.get("形式"))
        rec = make(hyo, "ZA" if hyo == "在宅A" else "ZB", it, fmt, sp)
        if hyo == "在宅A" and no == "問9":
            rec["列"] = []
            book.append(rec)
            for code, name, opts in ZAI_Q9:
                book.append({"票": hyo, "番号": f"問9{code}", "変数": f"ZA_問9_{code}",
                             "設問文": name, "形式": "単一", "区分": sp.get("区分", "基本"),
                             "集計方法": "単純集計", "クロス軸": sp.get("クロス軸", "A,B,I,J"),
                             "見える化": "○", "備考": "問9のマトリクスの行",
                             "選択肢": opts, "列": [f"ZA_問9_{code}"]})
            continue
        book.append(rec)

    json.dump(book, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return book


def make(hyo, prefix, it, fmt, sp):
    v = varname(prefix, it["番号"])
    ov = OVERRIDES.get((hyo, it["番号"]), {})
    opts = ov.get("選択肢", it["選択肢"])
    fmt = ov.get("形式", fmt)
    if fmt.startswith("複数"):
        cols = [f"{v}_c{n}" for n, _ in opts] or [v]
    elif fmt == "数値" and "身長" in it["設問文"]:
        cols = [v + "_身長", v + "_体重"]
    else:
        cols = [v]
    if fmt.endswith("＋自由"):
        cols = cols + [v + "_自由"]
    rec = {"票": hyo, "番号": it["番号"], "変数": v,
           "設問文": ov.get("短縮", it["設問文"]),
           "形式": fmt, "区分": sp.get("区分", "―"),
           "集計方法": sp.get("集計方法", "単純集計"),
           "クロス軸": sp.get("クロス軸", "A"),
           "見える化": sp.get("見える化", "―"),
           "備考": sp.get("備考", ""),
           "選択肢": opts, "列": cols}
    if "数値範囲" in ov:
        rec["数値範囲"] = ov["数値範囲"]
    return rec


if __name__ == "__main__":
    b = build()
    import collections
    print(f"設問 {len(b)}件 -> {os.path.relpath(OUT, BASE)}")
    print(" 票別:", collections.Counter(x["票"] for x in b).most_common())
    print(" 形式:", collections.Counter(x["形式"] for x in b).most_common())
    print(" 区分:", collections.Counter(x["区分"] for x in b).most_common())
    print(" 列数:", sum(len(x["列"]) for x in b))
    miss = [x["番号"] for x in b if x["区分"] == "―"]
    if miss:
        print(f" 集計仕様と突合できなかった設問 {len(miss)}件: {miss}")
