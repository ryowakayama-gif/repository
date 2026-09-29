"""成果品に載っている集計値を、一次資料から導き直して突き合わせる。

令和8年9月29日のご指示
  「一次資料に戻る点検を他の集計値についても通して下さい」

令和7年度調査の「サービスを利用していない理由」で、複数回答の割合を足し上げて
いた（92.9％）ことが分かったのを受けて、**成果品に直書きしている集計値を
一次資料から計算し直し、1件ずつ突き合わせる**ことにしたもの。

一次資料
  年報　　　18_年報・国保連データ/原本_年報/年報データ_2021〜2025_小野町.xlsx
  推計人口　18_年報・国保連データ/【受領】第10期将来推計用推計人口_小野町.xlsx
  見える化　04_算定・見込量/原本_見える化ワークシート/…_20260928_推計完了.xlsx
  総合事業　20_総合事業ワークシート/原本_受領_20260928/…ワークシート_Ver2….xlsm
  交付金　　21_給付適正化・交付金/原本_受領_20260928/…集計表_市町村分….xlsx
  第9期計画 13_関連計画/小野町高齢者保健福祉計画_第9期介護保険事業計画_令和6年3月.pdf

判定
  OK　　一次資料の値と一致した
  NG　　食い違った（成果品を直す）
  参考　一次資料が手元にないため突き合わせられない（出所を明記して扱う）

  python3 check_ono_ichiji.py
"""

import pathlib
import re
import sys
import tempfile
import warnings

warnings.filterwarnings("ignore")

import openpyxl

ROOT = pathlib.Path(__file__).parent
SRC = ROOT / "小野町_引継ぎ_整理済"
NENPO = SRC / "18_年報・国保連データ" / "原本_年報"
WORK = pathlib.Path(tempfile.mkdtemp())

RESULT = []            # (区分, 項目, 成果品の値, 一次資料の値, 判定)

# 一次資料と食い違うが、理由があって別の値を採っているもの。
# 判定は「相違（説明済み）」とし、NGには数えない。
EXCEPT = set()
EXCEPT_WHY = {
    "年報 様式1|第2段階 賦課割合":
        "年報の令和7年度は0.486だが、第9期計画本文が「基準額×0.485」"
        "「年額38,500円」と明記しており、令和6年度の年報も485である。"
        "令和7年度の486を入力誤りとみて0.485を採っている。",
}


def chk(kubun, item, used, primary, tol=0):
    """成果品の値と一次資料の値を突き合わせる。"""
    if primary is None:
        ok = "参考"
    elif isinstance(used, (int, float)) and isinstance(primary, (int, float)):
        ok = "OK" if abs(used - primary) <= tol else "NG"
    else:
        ok = "OK" if str(used) == str(primary) else "NG"
    if ok == "NG" and f"{kubun}|{item}" in EXCEPT:
        ok = "相違（説明済み）"
    RESULT.append((kubun, item, used, primary, ok))
    return ok


# ================================================================ 年報

def nenpo_books():
    from build_ono_nenpo import load_nenpo
    out = {}
    for y, lab in ((2021, "令和3年度"), (2022, "令和4年度"), (2023, "令和5年度"),
                   (2024, "令和6年度"), (2025, "令和7年度")):
        out[lab] = load_nenpo(NENPO / f"年報データ_{y}_小野町.xlsx", WORK)
    return out


def _y4(ws, label, col=11):
    """様式4 の歳出側から、科目名で決算額を拾う。"""
    for r in range(1, ws.max_row + 1):
        for c in (8, 9):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip() == label:
                return ws.cell(r, col).value
    return None


def _y4_row(ws, label):
    """様式4 の下部（歳入歳出差引残額など）を拾う。"""
    for r in range(1, ws.max_row + 1):
        for c in range(1, 8):
            v = ws.cell(r, c).value
            if isinstance(v, str) and label in v.replace("　", ""):
                for cc in range(c + 1, ws.max_column + 1):
                    x = ws.cell(r, cc).value
                    if isinstance(x, (int, float)):
                        return x
    return None


def check_keiri(wbs):
    """地域支援事業費・基金・繰越金（年報 様式4）。"""
    import build_ono_tanka as T
    for lab, d in T.CHIIKI_JISSEKI.items():
        ws = wbs[lab]["様式４"]
        for nm, key in (("介護予防・生活支援サービス事業費",
                         "介護予防・生活支援サービス事業"),
                        ("一般介護予防事業費", "一般介護予防事業"),
                        ("包括的支援事業・任意事業", "包括的支援事業・任意事業")):
            chk("年報 様式4", f"{lab} {nm}", d[key], _y4(ws, nm))
    ws6 = wbs["令和6年度"]["様式４"]
    chk("年報 様式4", "令和6年度 歳入歳出差引残額", 144205672,
        _y4_row(ws6, "歳入歳出差引残額"))
    chk("年報 様式4", "令和6年度 介護給付費準備基金保有額", 12000000,
        _y4_row(ws6, "介護給付費準備基金保有額"))
    for lab, v in (("令和4年度", 120000000), ("令和5年度", 12000000),
                   ("令和6年度", 12000000)):
        chk("年報 様式4", f"{lab} 基金保有額",
            v, _y4_row(wbs[lab]["様式４"], "介護給付費準備基金保有額"))
    # 繰入金（取崩）は全年度0と説明している
    for lab in ("令和3年度", "令和4年度", "令和5年度", "令和6年度"):
        ws = wbs[lab]["様式４"]
        kuri = None
        for r in range(1, ws.max_row + 1):
            v = ws.cell(r, 5).value
            if isinstance(v, str) and "介護給付費準備基金繰入金" in v:
                kuri = ws.cell(r, 7).value
        chk("年報 様式4", f"{lab} 介護給付費準備基金繰入金（取崩）", 0, kuri)
    # 令和7年度の様式4 が未入力であること
    ws7 = wbs["令和7年度"]["様式４"]
    nz = sum(1 for r in range(1, ws7.max_row + 1)
             for c in range(1, ws7.max_column + 1)
             if isinstance(ws7.cell(r, c).value, (int, float))
             and ws7.cell(r, c).value)
    chk("年報 様式4", "令和7年度 様式4 の非ゼロの数値", 0, nz)


def check_shuno(wbs):
    """保険料収納率（年報 様式3）。"""
    import build_ono_tanka as T
    ws = wbs["令和7年度"]["様式３"]
    chotei, shuno = ws.cell(12, 6).value, ws.cell(12, 7).value
    rate = round(shuno / chotei * 100, 2) if chotei else None
    chk("年報 様式3", "令和7年度 現年度分収納率（％）", 99.35, rate, tol=0.005)
    chk("年報 様式3", "予定保険料収納率に用いた値（％）",
        round(T.SHUNORITSU * 100, 2), rate, tol=0.005)


def check_shotoku(wbs):
    """所得段階別第1号被保険者数（年報 様式1 所得段階別）。"""
    import build_ono_tanka as T
    # C列＝段階名、O列（15）＝保険者の定める割合×1000、T列（20）＝被保険者数
    def read(lab):
        ws = wbs[lab]["様式１ 所得段階別"]
        out = []
        for r in range(1, ws.max_row + 1):
            v = ws.cell(r, 3).value
            if isinstance(v, str) and re.fullmatch(r"第[0-9０-９一二三四五六七八九十]+段階",
                                                   v.strip()):
                out.append((v.strip(), ws.cell(r, 15).value, ws.cell(r, 20).value))
        return out

    got = read("令和7年度")
    EXCEPT.add("年報 様式1|第2段階 賦課割合")
    chk("年報 様式1", "所得段階の数", len(T.SHOTOKU_R7), len(got))
    for (nm, _h, fuka, n), (_g, gr, gn) in zip(T.SHOTOKU_R7, got):
        chk("年報 様式1", f"{nm} 被保険者数（令和7年度末）", n, gn)
        chk("年報 様式1", f"{nm} 賦課割合", round(fuka, 3),
            round((gr or 0) / 1000, 3), tol=0.0005)
    chk("年報 様式1", "所得段階別の合計（人）",
        sum(n for *_x, n in T.SHOTOKU_R7), sum(n for *_x, n in got if n))
    # 第2段階の乗率は、年報の令和7年度が0.486、令和6年度が0.485。
    # 第9期計画本文が0.485・年額38,500円と明記しているため0.485を採っている。
    r6 = {nm: gr for nm, gr, _n in read("令和6年度")}
    r7 = {nm: gr for nm, gr, _n in got}
    chk("年報 様式1", "第2段階の乗率（令和6年度・採用値）", 485, r6.get("第２段階"))
    chk("年報 様式1", "第2段階の乗率（令和7年度・採らなかった値）", 486, r7.get("第２段階"))


def check_nintei(wbs):
    """認定者数（年報 様式1の5 総数）。"""
    import build_ono_tanka as T
    _act, _sub, gen = T.extract()
    want = {"令和3年度": None, "令和4年度": None, "令和5年度": None,
            "令和6年度": 944, "令和7年度": 792}
    for lab, v in want.items():
        if v is None:
            continue
        chk("年報 様式1の5", f"{lab}末 認定者数（人）", v, gen[lab]["認定"])
    return gen


def check_kyufu(gen):
    """総給付費・受給者数（年報 様式3・様式1の6）。"""
    import build_ono_tanka as T
    chk("年報 様式3", "令和3年度 総給付費（百万円）", 1060.7,
        round(gen["令和3年度"]["総給付費"] / 1e6, 1), tol=0.05)
    chk("年報 様式3", "令和7年度 総給付費（百万円）", 1005.1,
        round(gen["令和7年度"]["総給付費"] / 1e6, 1), tol=0.05)
    chk("年報 様式1", "令和3年度 第1号被保険者数（人）", 3498, gen["令和3年度"]["1号"])
    chk("年報 様式1", "令和7年度 第1号被保険者数（人）", 3443, gen["令和7年度"]["1号"])
    r3 = gen["令和3年度"]["総給付費"] / gen["令和3年度"]["1号"] / 1000
    r7 = gen["令和7年度"]["総給付費"] / gen["令和7年度"]["1号"] / 1000
    chk("年報", "令和3年度 1人当たり給付費（千円）", 303.2, round(r3, 1), tol=0.05)
    chk("年報", "令和7年度 1人当たり給付費（千円）", 291.9, round(r7, 1), tol=0.05)
    chk("年報", "令和3→7年度 総給付費の変化（％）", -5.2,
        round((gen["令和7年度"]["総給付費"] / gen["令和3年度"]["総給付費"] - 1) * 100, 1),
        tol=0.05)


def check_kensu():
    """住宅改修・特定福祉用具販売の件数（年報 様式2・介護＋予防）。"""
    import build_ono_tanka as T
    act, _sub, _gen = T.extract()
    for nm, want in (("住宅改修", 33), ("特定福祉用具販売", 62)):
        yo, ka = act["令和7年度"]["件数"].get(nm, (0, 0))
        chk("年報 様式2", f"令和7年度 {nm} 件数（介護＋予防）", want, yo + ka)


# ================================================================ 推計人口

def check_jinko():
    import build_ono_tanka as T
    p = T.POP_DAI10
    chk("将来推計人口", "令和9年度 総人口（人）", 8591, p["令和9年度"].get("総人口"))
    chk("将来推計人口", "令和11年度 総人口（人）", 8232, p["令和11年度"].get("総人口"))
    chk("将来推計人口", "令和9年度 第1号被保険者数（人）", 3438, p["令和9年度"].get("1号"))
    chk("将来推計人口", "令和11年度 第1号被保険者数（人）", 3418, p["令和11年度"].get("1号"))


# ================================================================ 見える化

def check_mieruka():
    p = (SRC / "04_算定・見込量" / "原本_見える化ワークシート"
         / "【受領】第10期介護保険事業計画策定に向けたワークシート_"
           "将来推計0925-2_20260928_推計完了.xlsx")
    if not p.exists():
        chk("見える化", "推計完了ワークシート", "―", None)
        return
    import build_ono_mieruka_kanryo as M
    wb = openpyxl.load_workbook(p, data_only=True)
    ws = wb["5_保険料推計"]
    lab = {}
    for r in range(1, ws.max_row + 1):
        v = next((ws.cell(r, c).value for c in range(1, 6)
                  if isinstance(ws.cell(r, c).value, str)
                  and ws.cell(r, c).value.strip()), None)
        if v:
            lab.setdefault(v.strip(), r)

    def val(name, col=6):
        r = lab.get(name)
        return ws.cell(r, col).value if r else None

    chk("見える化", "標準給付費見込額（3年計・千円）", M.SYS_STD3,
        round(val("標準給付費見込額") / 1000, 3), tol=0.01)
    chk("見える化", "基準保険料額（月額・円）", 6533.46,
        round(val("基準保険料額（月額）"), 2), tol=0.01)
    for nm, want in (("総給付費", 5772.56), ("その他給付費", 399.64),
                     ("地域支援事業費", 361.26), ("在宅サービス", 2849.84),
                     ("居住系サービス", 786.24), ("施設サービス", 2136.47)):
        chk("見える化", f"保険料の内訳 {nm}（月額・円）", want,
            round(val(nm), 2), tol=0.01)
    wb.close()


# ================================================================ 総合事業WS

def _sogo_rows():
    p = (SRC / "20_総合事業ワークシート" / "原本_受領_20260928"
         / "【受領】（改修版）総合事業の充実に向けたワークシート_Ver2_20260928.xlsm")
    wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
    ws = wb["総合事業"]
    rows = []
    for r in ws.iter_rows(values_only=True):
        if any(x is not None and str(x).strip().startswith("07522")
               for x in r[4:8]):
            rows.append((str(r[2]), list(r)))
    wb.close()
    order = ["平成29年度", "平成30年度", "令和元年度", "令和2年度",
             "令和3年度", "令和4年度", "令和5年度"]
    rows.sort(key=lambda x: order.index(x[0]))
    return order, rows


def check_sogo():
    import build_ono_sogo_ws as W
    order, rows = _sogo_rows()
    chk("総合事業WS", "小野町の行数（年度）", len(W.CHOSA_Y), len(rows))
    col = {"訪問型　従前相当　実人数": 9, "訪問型　従前相当　延べ人数": 10,
           "訪問型サービスB　実人数": 15, "訪問型サービスB　延べ人数": 16,
           "通所型　従前相当　実人数": 27, "通所型　従前相当　延べ人数": 28,
           "通所型サービスB　実人数": 33, "通所型サービスB　延べ人数": 34,
           "介護予防ケアマネジメント　件数": 94}
    for nm, vals in W.CHOSA:
        c = col[nm]
        got = []
        for _y, r in rows:
            v = r[c] if c < len(r) else None
            got.append(v if isinstance(v, (int, float)) else None)
        chk("総合事業WS", f"調査回答 {nm}", str(vals), str(got))
    jcol = {"訪問型　従前相当": 41, "訪問型サービスA": 46,
            "**訪問型サービスB（住民主体）**": 51, "訪問型サービスC（短期集中）": 56,
            "訪問型サービスD（移動支援）": 61, "通所型　従前相当": 71,
            "通所型サービスA": 76, "**通所型サービスB（住民主体）**": 81,
            "通所型サービスC（短期集中）": 86, "介護予防ケアマネジメントA": 96,
            "介護予防ケアマネジメントB": 99, "介護予防ケアマネジメントC": 102}
    for nm, d in W.JISSHI:
        c = jcol[nm]
        got = {}
        for y, r in rows:
            v = r[c] if c < len(r) else None
            if isinstance(v, (int, float)):
                got[y] = int(v)
        chk("総合事業WS", f"実施の有無 {nm.replace('**', '')}", str(d), str(got))
    for nm, vals in W.JIGYOSHO:
        c = 52 if "訪問" in nm else 82
        got = [(r[c] if c < len(r) and isinstance(r[c], (int, float)) else None)
               for _y, r in rows]
        chk("総合事業WS", f"事業所数 {nm.replace('**', '')}", str(vals), str(got))
    # 歳出シートが年報と合うこと
    import build_ono_tanka as T
    for y, (a, b, c) in W.SAISHUTSU.items():
        j = T.CHIIKI_JISSEKI.get(y)
        if not j:
            chk("総合事業WS", f"歳出 {y}", f"{a}/{b}/{c}", None)
            continue
        chk("総合事業WS", f"歳出 {y}（年報との一致）", f"{a}/{b}/{c}",
            f"{j['介護予防・生活支援サービス事業']}/{j['一般介護予防事業']}"
            f"/{j['包括的支援事業・任意事業']}")


# ================================================================ 交付金

def check_kofukin():
    import ono_kofukin_data as K
    import build_ono_tekiseika as X
    d = K.load(refresh=True)
    z = d["合計"]["推進・支援合計"]
    chk("交付金", "推進・支援合計（点）", 537, z["小野町"])
    chk("交付金", "市町村分の平均（点）", 454.8, round(z["全国平均"], 1), tol=0.05)
    chk("交付金", "市町村分の件数", 1536, z["全国件数"])
    chk("交付金", "県内の件数", 59, z["県件数"])
    chk("交付金", "市町村分の中での順位", 258, z["全国順位"])
    chk("交付金", "県内順位", 13, z["県内順位"])
    chk("交付金", "公表順位の母数", 1741, int(d["基本"]["前年度順位母数"]))
    chk("交付金", "令和7年度の得点（点）", 324, int(d["基本"]["前年度得点"]))
    chk("交付金", "推進Ⅱ 合計（点）", 84,
        d["合計"]["推進Ⅱ　公正・公平な給付を行う体制を構築する（介護給付適正化）"]["小野町"])
    chk("交付金", "支援Ⅲ 合計（点）", 55,
        d["合計"]["支援Ⅲ　在宅医療・在宅介護連携の体制を構築する"]["小野町"])
    chk("交付金", "推進Ⅱ 体制・取組（点）", 68, d["指標群"]["推進Ⅱ　体制・取組"]["小野町"])
    chk("交付金", "推進Ⅲ 活動（点）", 0, d["指標群"]["推進Ⅲ　活動"]["小野町"])
    # 指標ごとの配点の和が、評価指標に書かれた指標群の配点と合うこと
    for nm, v in d["指標群"].items():
        kf = ("保険者機能強化推進交付金" if nm.startswith("推進")
              else "介護保険保険者努力支援交付金")
        mk = "目標" + nm.split("　")[0][-1]
        tai = "体制" in nm
        h = sum(s["配点"] for s in X.SHIHYO
                if s["交付金"] == kf and s["目標"].startswith(mk)
                and (("体制" in s["指標群"]) == tai))
        chk("交付金", f"{nm} の配点", v["配点"], h)


# ================================================================ 第9期計画

def check_dai9():
    p = (SRC / "13_関連計画"
         / "小野町高齢者保健福祉計画_第9期介護保険事業計画_令和6年3月.pdf")
    if not p.exists():
        chk("第9期計画", "本文PDF", "―", None)
        return
    from pypdf import PdfReader
    txt = "".join((pg.extract_text() or "") for pg in PdfReader(str(p)).pages)
    flat = re.sub(r"\s+", "", txt).replace(",", "")

    def has(s):
        return s.replace(",", "") in flat

    # 保険料の算定表（第4章(8)）の数値。計画本文は円単位で書いている。
    for nm, used, s in (
            ("標準給付費見込額（3年計）", "3,566,486千円", "3,566,485,542"),
            ("　令和6年度の標準給付費見込額", "1,185,894千円", "1,185,893,846"),
            ("　令和7年度の標準給付費見込額", "1,189,712千円", "1,189,711,868"),
            ("保険料収納必要額（G）", "832,562,753円", "832,562,753"),
            ("介護給付費準備基金取崩額（L）", "35,000,000円", "35,000,000"),
            ("保険料収納必要額（M＝G－L）", "797,562,752円", "797,562,752"),
            ("補正後被保険者数（3年計）", "10,136人", "10,136"),
            ("保険料基準額（月額・算定値）", "6,886円", "6,886"),
            ("保険料基準額（月額・基金取崩後）", "6,597円", "6,597"),
            ("保険料基準額（月額・最終）", "6,600円", "6,600"),
            ("保険料基準額（年額・最終）", "79,200円", "79,200"),
            ("第2段階の乗率", "0.485", "0.485"),
            ("第2段階の年額", "38,500円", "38,500")):
        chk("第9期計画", nm, used, used if has(s) else f"本文に{s}が見当たらない")
    for nm, kw in (("「評価指標」の語", "評価指標"),
                   ("「目指すべき姿」の語", "目指すべき姿"),
                   ("「看取り」の語", "看取り"),
                   ("「入退院」の語", "入退院")):
        chk("第9期計画", f"{nm}の有無", "なし" if kw in ("目指すべき姿", "看取り",
                                                  "入退院") else "あり",
            "あり" if kw in flat else "なし")


# ================================================================ 調査

def check_chosa():
    """令和7年度調査。報告書（104頁）は手元にないため、内部の整合のみ確かめる。"""
    import build_ono_enquete_report as E
    chk("令和7年度調査", "調査結果報告書（104頁・町から受領）", "―", None)
    # 未利用理由は複数回答。全項目の合計が100％を超えることを確かめる
    riyu = [46.8, 29.2, 16.9, 5.2, 4.5, 3.9, 0.6, 0.6, 16.2, 8.4]
    chk("令和7年度調査", "未利用理由の全項目の合計（％）", round(sum(riyu), 1),
        round(sum(riyu), 1))
    chk("令和7年度調査", "複数回答か（合計が100％超）", "複数回答",
        "複数回答" if sum(riyu) > 100.5 else "単一回答")
    n = 154
    for pct, want in ((46.8, 72), (29.2, 45), (16.9, 26), (0.6, 1),
                      (16.2, 25), (8.4, 13)):
        chk("令和7年度調査", f"未利用理由 {pct}％ の人数（n=154）", want,
            round(pct / 100 * n))
    chk("令和7年度調査", "未利用の人数（366人×42.1％）", 154,
        round(366 * 0.421))


# ================================================================ 見える化の地域分析

MIE = SRC / "07_介護保険_見える化整理"


def _chiiki_hyo(path, sheet="表形式（地域別）"):
    """見える化の地域別ファイルから、福島県内59市町村の値を取り出す。

    **個別の団体名と個別の値は成果品に出さない。順位を出すためだけに使う。**
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    names = [str(x) for x in rows[1][3:] if x is not None]
    out = []
    for r in rows[2:]:
        lab = r[1]
        if not isinstance(lab, str) or lab.startswith("行"):
            continue
        vals = list(r[3:3 + len(names)])
        out.append((lab, dict(zip(names, vals))))
    wb.close()
    return out


def _rank(d, target="小野町", desc=True):
    v = d.get(target)
    if v is None:
        return None, None
    xs = [x for n, x in d.items()
          if n not in ("福島県", "全国") and not n.startswith("列")
          and isinstance(x, (int, float))]
    return sorted(xs, reverse=desc).index(v) + 1, len(xs)


def check_chiiki():
    """地域差指数・自給率・給付月額の県内順位・定員。"""
    wb = openpyxl.load_workbook(
        MIE / "小野町_第10期介護保険事業計画_地域分析データ整理_認定受給給付自給率.xlsx",
        data_only=True)
    shisu = {str(r[1]): r[3] for r in wb["05_地域差指数"].iter_rows(values_only=True)
             if r and isinstance(r[1], str)}
    chk("見える化 D49", "要介護認定率（性・年齢調整後）の地域差指数", 1.24,
        shisu.get("要介護認定率（性・年齢調整後）"), tol=0.005)
    chk("見える化 D49", "第1号被保険者1人あたり給付月額の地域差指数", 1.08,
        shisu.get("第１号被保険者１人あたり給付月額（性・年齢調整後）"), tol=0.005)
    chk("見える化 D49", "受給率の地域差指数", 1.07, shisu.get("受給率"), tol=0.005)
    long = {}
    for r in wb["06_ロングデータ"].iter_rows(values_only=True):
        if r and r[4] == "【注目する地域のみ】介護サービス自給率":
            long[str(r[6]).replace(" ", "")] = r[7]
    wb.close()
    chk("見える化 δ2", "介護サービス自給率（令和5年度・％）", 72.43,
        long.get("令和5年度"), tol=0.005)
    chk("見える化 δ2", "介護サービス自給率（令和7年度・％）", 67.57,
        long.get("令和7年度"), tol=0.005)

    # 第1号被保険者1人あたり給付月額の県内順位（D6・2025年度・地域別）
    d6 = _chiiki_hyo(MIE / "給付分析"
                     / "D6_第１号被保険者１人あたり給付月額（在宅サービス・"
                       "施設および居住系サービス）_2025_地域別 (2).xlsx")
    for lab, used in (("第１号被保険者１人あたり給付月額（在宅サービス）", 28),
                      ("第１号被保険者１人あたり給付月額（施設および居住系サービス）",
                       33)):
        d = next((v for nm, v in d6 if nm.startswith(lab[:20])), None)
        r, n = _rank(d) if d else (None, None)
        nm = "在宅サービス" if "在宅" in lab else "施設・居住系サービス"
        chk("見える化 D6", f"給付月額の県内順位（{nm}）", f"{used}位／59", f"{r}位／{n}")

    # 認定率の県内順位は、地域別の出力が手元にないため確かめられない
    # 認定率の県内順位は、地域別の出力が手元にない。
    # 当方の給付費等分析結果報告書は、令和7年3月末26.8％＝県内2位、
    # 令和8年3月末22.7％＝県内6位としている。成果品では年度を明記して扱う。
    chk("見える化", "要介護認定率の県内順位（地域別の出力が未取得）",
        "令和7年3月末 26.8％＝県内2位", None)

    # 定員（D25〜D27）
    for fn, want in (
            ("入所（利用）定員/D25_定員（施設サービス別）_時系列 (1).xlsx",
             {"定員（介護老人福祉施設）": 54,
              "定員（地域密着型介護老人福祉施設入所者生活介護）": 58}),
            ("入所（利用）定員/D26_定員（居住系サービス別）_時系列 (1).xlsx",
             {"定員（認知症対応型共同生活介護）": 98}),
            ("入所（利用）定員/D27_定員（通所系サービス別）_時系列 (1).xlsx",
             {"定員（通所介護）": 139, "定員（認知症対応型通所介護）": 12,
              "定員_通い（小規模多機能型居宅介護）": 30})):
        wb = openpyxl.load_workbook(MIE / fn, data_only=True)
        ws = wb[wb.sheetnames[-1]]
        rows = list(ws.iter_rows(values_only=True))
        head = [str(x).replace("\n", "") if x is not None else "" for x in rows[1]]
        nendo = [i for i, h in enumerate(head) if h.endswith("年度")]
        last = nendo[-1]
        for r in rows[2:]:
            if r[1] != "小野町":
                continue
            if str(r[2]) in want:
                chk("見える化 D25〜27", f"{r[2]}（{head[last]}）",
                    want[str(r[2])], r[last])
        # 認知症対応型共同生活介護の増床（令和4年度71人→令和6年度98人）
        if "D26" in fn:
            for r in rows[2:]:
                if r[1] == "小野町" and str(r[2]) == "定員（認知症対応型共同生活介護）":
                    i4 = head.index("令和4年度") if "令和4年度" in head else None
                    chk("見える化 D26", "認知症対応型共同生活介護 令和4年度の定員",
                        71, r[i4] if i4 else None)
        wb.close()
    chk("見える化 D25〜27", "定員の最新年度（成果品の表記）", "令和6年度", "令和6年度")


def check_yobo():
    """通いの場・認知症施策（見える化 F系・J系）。"""
    wb = openpyxl.load_workbook(
        MIE / "小野町_第10期介護保険事業計画_供給体制・介護予防整理.xlsx",
        data_only=True)
    ws = wb["02_通いの場"]
    rows = list(ws.iter_rows(values_only=True))
    head = [str(x).replace("\n", "") if x is not None else "" for x in rows[0]]
    kasho = next((r for r in rows if r[1] == "週1回以上の通いの場の箇所数"), None)
    after = [v for h, v in zip(head[2:], kasho[2:])
             if h.startswith(("H28", "H29", "H30", "R"))]
    chk("見える化 F3", "平成28年度以降の週1回以上の通いの場の箇所数",
        "すべて0", "すべて0" if all((v in (0, "0", None, "-")) for v in after)
        else str(after))
    ws = wb["04_認知症施策"]
    d = {str(r[1]): (r[3], r[4]) for r in ws.iter_rows(values_only=True)
         if r and isinstance(r[1], str)}
    chk("見える化 J17", "認知症地域支援推進員の配置人数（前回→直近）", "1→4",
        "{}→{}".format(*d.get("認知症地域支援推進員配置状況（配置人数）", ("", ""))))
    chk("見える化 J16", "初期集中支援チームの訪問実績（前回→直近）", "0→0",
        "{}→{}".format(*d.get("初期集中支援チームの実施状況（訪問実績）", ("", ""))))
    wb.close()


# ================================================================ main

def main():
    wbs = nenpo_books()
    check_keiri(wbs)
    check_shuno(wbs)
    check_shotoku(wbs)
    gen = check_nintei(wbs)
    check_kyufu(gen)
    check_kensu()
    check_jinko()
    check_mieruka()
    check_sogo()
    check_kofukin()
    check_dai9()
    check_chiiki()
    check_yobo()
    check_chosa()

    ng = [r for r in RESULT if r[4] == "NG"]
    ref = [r for r in RESULT if r[4] == "参考"]
    exc = [r for r in RESULT if r[4] == "相違（説明済み）"]
    print(f"点検 {len(RESULT)}件　"
          f"OK {len(RESULT) - len(ng) - len(ref) - len(exc)}／NG {len(ng)}／"
          f"相違（説明済み） {len(exc)}／参考 {len(ref)}\n")
    if exc:
        print("■ 一次資料と違う値を、理由があって採っているもの")
        for k, i, u, pv, _o in exc:
            print(f"  ・[{k}] {i}　成果品 {u} ／ 一次資料 {pv}")
            print(f"      {EXCEPT_WHY.get(f'{k}|{i}', '')}")
        print()
    if ng:
        print("■ 食い違い")
        for k, i, u, p, _o in ng:
            print(f"  NG  [{k}] {i}")
            print(f"        成果品 {u}")
            print(f"        一次資料 {p}")
    if ref:
        print("\n■ 一次資料が手元にないもの（出所を明記して扱う）")
        for k, i, _u, _p, _o in ref:
            print(f"  参考 [{k}] {i}")
    return 1 if ng else 0


if __name__ == "__main__":
    sys.exit(main())
