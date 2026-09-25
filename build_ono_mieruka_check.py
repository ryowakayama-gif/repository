"""小野町 見える化システムのワークシート（将来推計1）を自前の算定と突き合わせる。

町から受領した「第10期介護保険事業（支援）計画策定に向けたワークシート」（出力日 令和8年9月15日・推計パターン名「将来推計1」）は
**何も調整していない自然体推計**である。町から「確認してほしい」との依頼があったため、
(1) 使える値、(2) 直さないと使えない値、(3) 入力が要る値 の3つに仕分ける。

**見える化と自前算定の作りの違い**

  見える化  基準年度＝令和8年度。令和8年度の月平均実績を12倍した値を置き、
            利用率・1人1月あたり回数の伸びを0として第10期3年度に延ばす。
  自前算定  基準期間＝令和5〜7年度（サービスにより令和6〜7年度・令和7年度単年）。
            第1号被保険者1人あたりの利用率を平均し、将来推計人口を乗じる。

  令和8年度は本稿執筆時点で審査4か月分しか確定していない。見える化の令和8年度は
  その月平均を年換算したものなので、**サービスによっては年間実績と大きく食い違う**。
  総給付費では相殺されるが、サービス別では採れない値がある。

**見える化のワークシートで未入力のもの**

  ・地域支援事業費（3シート全体がゼロ）
  ・特定入所者介護サービス費・高額介護サービス費・高額医療合算・審査支払手数料
  ・準備基金残高と取崩額、予定保険料収納率、所得段階別加入割合
  ・調整交付金見込額

  このため「5_保険料推計」は #DIV/0! のままで、保険料は出ていない。

出力:
  04_算定・見込量/小野町_第10期_見える化ワークシートとの突合_YYYYMMDD.xlsx
  04_算定・見込量/小野町_見える化ワークシートの確認結果_YYYYMMDD.docx
"""

import pathlib
import tempfile
import warnings

warnings.filterwarnings("ignore")

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.styles.borders import Border, Side

import build_ono_nenpo as N
import build_ono_tanka as T
from build_ono_tanka import CHIIKI_JISSEKI  # noqa: F401
import ono_shizentai as SZ

# **標準は自然体推計（令和8年度基点）である。**
# ワークシートとの突合は、素案に載る値で行う。
plan_rows = SZ.plan_rows

ROOT = pathlib.Path(__file__).parent / "小野町_引継ぎ_整理済"
OUT = ROOT / "04_算定・見込量"
WS_FILE = (OUT / "原本_見える化ワークシート"
           / "【受領】第10期介護保険事業計画策定に向けたワークシート_将来推計1_20260914.xlsx")
NENPO = ROOT / "18_年報・国保連データ" / "原本_年報"
ASOF = "20260925"
ASOF_JP = "令和8年9月25日"

HEAD = PatternFill("solid", fgColor="1F3864")
SUB = PatternFill("solid", fgColor="D9E2F3")
OK = PatternFill("solid", fgColor="C6EFCE")
WARN = PatternFill("solid", fgColor="FFE0B2")
NG = PatternFill("solid", fgColor="FFC7CE")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

YEARS = ["令和9年度", "令和10年度", "令和11年度"]

# 見える化のサービス名 → 自前算定の (区分, 表示名)
MAP_KAIGO = {
    "訪問介護": ("居宅", "訪問介護"),
    "訪問入浴介護": ("居宅", "訪問入浴介護"),
    "訪問看護": ("居宅", "訪問看護"),
    "訪問リハビリテーション": ("居宅", "訪問リハビリテーション"),
    "居宅療養管理指導": ("居宅", "居宅療養管理指導"),
    "通所介護": ("居宅", "通所介護"),
    "通所リハビリテーション": ("居宅", "通所リハビリテーション"),
    "短期入所生活介護": ("居宅", "短期入所生活介護"),
    "短期入所療養介護（老健）": ("居宅", "短期入所療養介護（介護老人保健施設）"),
    "短期入所療養介護（病院等）": ("居宅", "短期入所療養介護（病院等）"),
    "短期入所療養介護(介護医療院)": ("居宅", "短期入所療養介護（介護医療院）"),
    "福祉用具貸与": ("居宅", "福祉用具貸与"),
    "特定福祉用具購入費": ("居宅", "特定福祉用具販売"),
    "住宅改修費": ("居宅", "住宅改修"),
    "特定施設入居者生活介護": ("居宅", "特定施設入居者生活介護"),
    "定期巡回・随時対応型訪問介護看護": ("地域密着型", "定期巡回・随時対応型訪問介護看護"),
    "夜間対応型訪問介護": ("地域密着型", "夜間対応型訪問介護"),
    "地域密着型通所介護": ("地域密着型", "地域密着型通所介護"),
    "認知症対応型通所介護": ("地域密着型", "認知症対応型通所介護"),
    "小規模多機能型居宅介護": ("地域密着型", "小規模多機能型居宅介護"),
    "認知症対応型共同生活介護": ("地域密着型", "認知症対応型共同生活介護"),
    "地域密着型特定施設入居者生活介護": ("地域密着型", "地域密着型特定施設入居者生活介護"),
    "地域密着型介護老人福祉施設入所者生活介護":
        ("地域密着型", "地域密着型介護老人福祉施設入所者生活介護"),
    "看護小規模多機能型居宅介護": ("地域密着型", "看護小規模多機能型居宅介護"),
    "介護老人福祉施設": ("施設", "介護老人福祉施設"),
    "介護老人保健施設": ("施設", "介護老人保健施設"),
    "介護医療院": ("施設", "介護医療院"),
    "介護療養型医療施設": ("施設", "介護療養型医療施設"),
    "居宅介護支援": ("居宅", "居宅介護支援"),
}
MAP_YOBO = {
    "介護予防訪問入浴介護": ("居宅", "介護予防訪問入浴介護"),
    "介護予防訪問看護": ("居宅", "介護予防訪問看護"),
    "介護予防訪問リハビリテーション": ("居宅", "介護予防訪問リハビリテーション"),
    "介護予防居宅療養管理指導": ("居宅", "介護予防居宅療養管理指導"),
    "介護予防通所リハビリテーション": ("居宅", "介護予防通所リハビリテーション"),
    "介護予防短期入所生活介護": ("居宅", "介護予防短期入所生活介護"),
    "介護予防短期入所療養介護（老健）":
        ("居宅", "介護予防短期入所療養介護（介護老人保健施設）"),
    "介護予防短期入所療養介護（病院等）": ("居宅", "介護予防短期入所療養介護（病院等）"),
    "介護予防短期入所療養介護(介護医療院)": ("居宅", "介護予防短期入所療養介護（介護医療院）"),
    "介護予防福祉用具貸与": ("居宅", "介護予防福祉用具貸与"),
    "特定介護予防福祉用具購入費": ("居宅", "特定介護予防福祉用具販売"),
    "介護予防住宅改修": ("居宅", "介護予防住宅改修"),
    "介護予防特定施設入居者生活介護": ("居宅", "介護予防特定施設入居者生活介護"),
    "介護予防認知症対応型通所介護": ("地域密着型", "介護予防認知症対応型通所介護"),
    "介護予防小規模多機能型居宅介護": ("地域密着型", "介護予防小規模多機能型居宅介護"),
    "介護予防認知症対応型共同生活介護": ("地域密着型", "介護予防認知症対応型共同生活介護"),
    "介護予防支援": ("居宅", "介護予防支援"),
}
SUBROWS = {"給付費（千円）": "給付費", "回数（回）": "量", "日数（日）": "量",
           "人数（人）": "人数"}


# ================================================ 見える化ワークシートの読み取り

def read_worksheet():
    """ワークシートから {種別: {サービス: {指標: {年度: 値}}}} を返す。"""
    wb = openpyxl.load_workbook(WS_FILE, data_only=True)
    ws = wb["2_サービス別給付費"]
    # 年度の列。D=令和6年度…I=令和11年度
    cols = {"令和6年度": 4, "令和7年度": 5, "令和8年度": 6,
            "令和9年度": 7, "令和10年度": 8, "令和11年度": 9}
    out = {"予防": {}, "介護": {}}
    kind, cur = None, None
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        b = ws.cell(r, 2).value
        c = ws.cell(r, 3).value
        if a and "１．介護予防サービス見込量" in str(a):
            kind = "予防"
        elif a and "２．介護サービス見込量" in str(a):
            kind = "介護"
        elif a and "３．総給付費" in str(a):
            kind = None
        if kind is None:
            continue
        if a and str(a).strip().startswith("合計"):
            cur = None          # 合計行をサービスの続きと読まないようにする
            continue
        name = None
        if b and str(b).strip() not in ("", "未使用"):
            name = str(b).strip()
        elif a and str(a).strip() in ("（３）介護予防支援", "（４）居宅介護支援"):
            name = str(a).strip().split("）")[1]
        if name:
            cur = name
            out[kind].setdefault(cur, {})
        if cur and c and str(c).strip() in SUBROWS:
            key = SUBROWS[str(c).strip()]
            d = out[kind][cur].setdefault(key, {})
            for y, col in cols.items():
                v = ws.cell(r, col).value
                d[y] = v if isinstance(v, (int, float)) else None
    # 総括
    s1 = wb["1_推計値サマリ"]
    tot = {}
    for r, key in ((11, "総人口"), (12, "1号"), (13, "2号"), (19, "認定者数")):
        tot[key] = {y: s1.cell(r, 4 + i).value
                    for i, y in enumerate(["令和6年度", "令和7年度", "令和8年度",
                                           "令和9年度", "令和10年度", "令和11年度"])}
    s2 = wb["2_サービス別給付費"]
    ku = {}
    for r, key in ((146, "総給付費"), (147, "在宅"), (148, "居住系"), (149, "施設")):
        ku[key] = {y: s2.cell(r, 4 + i).value
                   for i, y in enumerate(["令和6年度", "令和7年度", "令和8年度",
                                          "令和9年度", "令和10年度", "令和11年度"])}
    s5 = wb["5_保険料推計"]
    hoken = {
        "第9期保険料基準額": s5.cell(8, 6).value,
        "標準給付費見込額": s5.cell(68, 6).value,
        "地域支援事業費": s5.cell(86, 6).value,
        "第1号被保険者負担分相当額": s5.cell(90, 6).value,
        "調整交付金相当額": s5.cell(91, 6).value,
        "調整交付金見込額": s5.cell(92, 6).value,
        "保険料収納必要額": s5.cell(103, 6).value,
        "第1号被保険者数": s5.cell(109, 6).value,
        "後期加入割合補正係数": [s5.cell(96, 7 + i).value for i in range(3)],
    }
    s6 = wb["(参考)保険料の推計に要する係数"]
    hoken["第1号負担割合"] = s6.cell(5, 4).value
    return out, tot, ku, hoken


# ================================================ 年報 様式4（地域支援事業費の決算）

OUT_LABELS = ["介護予防・生活支援サービス事業費", "一般介護予防事業費",
              "包括的支援事業・任意事業", "重層的支援体制整備事業保険料操出金",
              "保健福祉事業費", "基金積立金", "総務費"]


def read_chiiki():
    """年報 様式4 の歳出から地域支援事業費の決算額（円）を年度別に返す。"""
    work = pathlib.Path(tempfile.mkdtemp())
    out = {}
    for code, jp in N.YEARS:
        wb = N.load_nenpo(NENPO / f"年報データ_{code}_小野町.xlsx", work)
        ws = wb["様式４"]
        d = {}
        for r in range(1, ws.max_row + 1):
            lab = ws.cell(r, 9).value or ws.cell(r, 8).value
            v = ws.cell(r, 11).value
            if lab and isinstance(v, (int, float)):
                t = str(lab).strip()
                if t in OUT_LABELS:
                    d[t] = v
                elif t.startswith("合"):
                    d["歳出合計"] = v
            lab2 = ws.cell(r, 5).value or ws.cell(r, 4).value
            v2 = ws.cell(r, 7).value
            if lab2 and isinstance(v2, (int, float)):
                t2 = str(lab2).strip()
                if t2 == "介護給付費準備基金保有額":
                    d["準備基金保有額"] = v2
                elif t2 == "介護給付費準備基金繰入金":
                    d["準備基金繰入金"] = v2
                elif t2.startswith("合"):
                    d["歳入合計"] = v2
        out[jp] = d
    return out


# ================================================ 突合

def compare():
    mie, tot, ku, hoken = read_worksheet()
    P = plan_rows()
    mine = {"介護": {}, "予防": {}}
    for kind in ("介護", "予防"):
        for cat, name, unit, nin, ryo, kyu in P[kind]:
            mine[kind][name] = {"区分": cat, "単位": unit, "人数": nin,
                                "量": ryo, "給付費": kyu}
    rows = []
    for kind, mp in (("介護", MAP_KAIGO), ("予防", MAP_YOBO)):
        for mname, (cat, myname) in mp.items():
            m = mie[kind].get(mname, {})
            o = mine[kind].get(myname)
            if o is None:
                continue
            mk = [(m.get("給付費", {}) or {}).get(y) or 0.0 for y in YEARS]
            ok = list(o["給付費"])
            rows.append({
                "種別": kind, "区分": cat, "見える化名": mname, "算定名": myname,
                "単位": o["単位"],
                "見_R6": (m.get("給付費", {}) or {}).get("令和6年度") or 0.0,
                "見_R7": (m.get("給付費", {}) or {}).get("令和7年度") or 0.0,
                "見_R8": (m.get("給付費", {}) or {}).get("令和8年度") or 0.0,
                "見_給付": mk, "自_給付": ok,
                "見_量": [(m.get("量", {}) or {}).get(y) for y in
                        ["令和7年度", "令和8年度", "令和9年度"]],
                "自_量": o["量"], "自_人数": o["人数"],
                "見_人数": [(m.get("人数", {}) or {}).get(y) for y in
                          ["令和7年度", "令和8年度", "令和9年度"]],
            })
    return rows, tot, ku, hoken, P


# ================================================ 書式

def style_head(ws, row=1):
    for c in ws[row]:
        if c.value is None:
            continue
        c.fill = HEAD
        c.font = Font(bold=True, color="FFFFFF", size=9)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def body(ws, first=2, wrap=()):
    for r in ws.iter_rows(min_row=first):
        for c in r:
            if c.value is None and not c.fill.fgColor.rgb:
                continue
            c.border = BORDER
            if c.font.size is None or c.font.size > 9:
                c.font = Font(bold=c.font.bold, size=9, color=c.font.color)
            c.alignment = Alignment(vertical="center",
                                    wrap_text=c.column in wrap)


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w


def num(ws, cols, first=2, fmt="#,##0"):
    for r in range(first, ws.max_row + 1):
        for c in cols:
            ws.cell(r, c).number_format = fmt


# ================================================ 月報による令和8年度の実測

# 月報のサービスコード → 見える化のサービス名
GEPPO_CODE = {
    "11訪問介護": "訪問介護", "12訪問入浴介護": "訪問入浴介護", "13訪問看護": "訪問看護",
    "14訪問リハビリテーション": "訪問リハビリテーション",
    "15通所介護": "通所介護", "16通所リハビリテーション": "通所リハビリテーション",
    "17福祉用具貸与": "福祉用具貸与", "21短期入所生活介護": "短期入所生活介護",
    "22短期入所療養介護（老健）": "短期入所療養介護（老健）",
    "31居宅療養管理指導": "居宅療養管理指導",
    "33特定施設入居者生活介護": "特定施設入居者生活介護",
    "43居宅介護支援": "居宅介護支援",
    "72認知症対応型通所介護": "認知症対応型通所介護",
    "73小規模多機能型居宅介護": "小規模多機能型居宅介護",
    "32認知症対応型共同生活介護": "認知症対応型共同生活介護",
    "54地域密着型介護老人福祉施設入所者生活介護":
        "地域密着型介護老人福祉施設入所者生活介護",
    "51介護福祉施設サービス": "介護老人福祉施設",
    "52介護保健施設サービス": "介護老人保健施設",
    "55介護医療院サービス": "介護医療院",
    "63介護予防訪問看護": "介護予防訪問看護",
    "66介護予防通所リハビリテーション": "介護予防通所リハビリテーション",
    "67介護予防福祉用具貸与": "介護予防福祉用具貸与",
    "75介護予防小規模多機能型居宅介護": "介護予防小規模多機能型居宅介護",
    "37介護予防認知症対応型共同生活介護": "介護予防認知症対応型共同生活介護",
    "46介護予防支援": "介護予防支援",
}


def read_geppo_year():
    """月報から {見える化名: {年度: (月平均件数, 月平均給付額)}} を返す。"""
    import build_ono_kaisu as K
    data = K.read_geppo()
    ms = sorted({ym for (ym, _i) in data})
    codes = {c: n for c, n in GEPPO_CODE.items()}
    out = {}
    for code, name in codes.items():
        by = {}
        for ind, slot in (("件数", 0), ("給付額", 1)):
            for ym in ms:
                d = data.get((ym, ind), {})
                v = sum((row.get("計", 0) or 0) for nm, row in d.items() if nm == code)
                y = K.fiscal(ym)
                by.setdefault(y, [[], []])[slot].append(v)
        out[name] = {y: (sum(v[0]) / len(v[0]) if v[0] else 0,
                         sum(v[1]) / len(v[1]) if v[1] else 0)
                     for y, v in by.items()}
        out[name]["月数"] = {y: len(v[0]) for y, v in by.items()}
    return out


# ================================================ 保険料の試算（比較用）

# 地域支援事業費は令和6年度決算（年報 様式4）を3年度とも据え置いた仮置き。
CHIIKI_R6_SOGO = CHIIKI_JISSEKI["令和6年度"]["総合事業"]          # 円
CHIIKI_R6_HOKATSU = CHIIKI_JISSEKI["令和6年度"]["包括的支援事業・任意事業"]
CHOSEI_MIKOMI = T.CHOSEI_MIKOMI * 100   # 調整交付金見込交付割合（％）算式による
KIKIN_TORIKUZUSHI = 35000  # 準備基金取崩額（千円）第9期と同額の仮置き
SHUNO = 0.9935             # 予定保険料収納率


def hosei_ninzu():
    """所得段階別加入割合補正後の第1号被保険者数（第10期3年計）。"""
    pop = sum(T.POP_DAI10[y]["1号"] for y in YEARS)
    tot = sum(n for *_x, n in T.SHOTOKU_R7)
    # 割合は軽減前の標準割合（見える化システムの「標準段階区分・割合」と同じ）。
    return sum(hyo * (n / tot) * pop for _s, hyo, _fu, n in T.SHOTOKU_R7), pop


def premium(hyojun_sen, sogo_sen=None, hokatsu_sen=None):
    """標準給付費（千円・3年計）から保険料基準額（円／月）を出す。"""
    sogo = sogo_sen if sogo_sen is not None else CHIIKI_R6_SOGO * 3 / 1000
    hok = hokatsu_sen if hokatsu_sen is not None else CHIIKI_R6_HOKATSU * 3 / 1000
    chiiki = sogo + hok
    dai1 = (hyojun_sen + chiiki) * 0.23
    chosei_sou = (hyojun_sen + sogo) * 0.05
    # 調整交付金見込額の基数は標準給付費見込額のみ（見える化システムの算式）。
    # 調整交付金相当額（5％）が総合事業費を含むのとは非対称。
    chosei_mi = hyojun_sen * CHOSEI_MIKOMI / 100
    shuno_hitsuyo = dai1 + chosei_sou - chosei_mi - KIKIN_TORIKUZUSHI
    hosei, _pop = hosei_ninzu()
    return {
        "標準給付費": hyojun_sen, "総合事業費": sogo, "包括的支援等": hok,
        "地域支援事業費": chiiki, "第1号負担分相当額": dai1,
        "調整交付金相当額": chosei_sou, "調整交付金見込額": chosei_mi,
        "保険料収納必要額": shuno_hitsuyo,
        "補正後被保険者数": hosei,
        "基準額": shuno_hitsuyo * 1000 / SHUNO / hosei / 12,
    }


# ================================================ シートの組み立て

def sheet_yoten(wb, facts):
    ws = wb.create_sheet("00_要点")
    ws.append(["#", "区分", "確認したこと", "根拠", "取扱い"])
    for i, (ku, txt, konkyo, atsukai) in enumerate(facts, start=1):
        ws.append([i, ku, txt, konkyo, atsukai])
    style_head(ws)
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, 5).value or ""
        ws.cell(r, 5).fill = (OK if v.startswith("採用") else
                              NG if v.startswith("不採用") else WARN)
    body(ws, wrap=(3, 4, 5))
    widths(ws, [4, 16, 62, 46, 34])
    ws.freeze_panes = "A2"
    return ws


def sheet_soukatsu(wb, tot, ku, P):
    ws = wb.create_sheet("01_総括の突合")
    ws.append(["項目", "出所", "令和6年度", "令和7年度", "令和8年度",
               "令和9年度", "令和10年度", "令和11年度", "第10期3年計", "備考"])
    ys6 = ["令和6年度", "令和7年度", "令和8年度"] + YEARS

    def line(name, src, vals, note, three=True):
        s = sum(v for v in vals[3:] if isinstance(v, (int, float))) if three else None
        ws.append([name, src] + list(vals) + [s, note])

    line("総給付費（千円）", "見える化", [ku["総給付費"][y] for y in ys6],
         "令和8年度を基準年に置き、利用率の伸び0で延ばしたもの")
    mine = ["", "", ""] + list(P["集計"]["計"])
    line("総給付費（千円）", "自前算定", mine,
         "令和5〜7年度の1号1人当たり利用率×将来推計人口")
    ws.append(["差（見える化－自前）", "", "", "", "",
               f"=D2-D3", f"=E2-E3", f"=F2-F3", f"=I2-I3", "3年計で+0.4％程度"])
    for c in (6, 7, 8, 9):
        ws.cell(ws.max_row, c).value = None
    r = ws.max_row
    for i, k in enumerate(range(6, 9)):
        ws.cell(r, k).value = (ku["総給付費"][YEARS[i]] - P["集計"]["計"][i])
    ws.cell(r, 9).value = sum(ku["総給付費"][y] for y in YEARS) - sum(P["集計"]["計"])
    ws.append([])
    ws.append(["【区分別（見える化の3区分）】"])
    for key in ("在宅", "居住系", "施設"):
        line(f"{key}サービス（千円）", "見える化", [ku[key][y] for y in ys6], "")
    ws.append([])
    ws.append(["【区分別（計画書の3区分）】"])
    for key, v in P["集計"]["区分"].items():
        line(f"{key}サービス（千円）", "自前算定", ["", "", ""] + list(v),
             "居住系を地域密着型・居宅に含めるため見える化とは区分が違う")
    ws.append([])
    ws.append(["【被保険者数・認定者数】"])
    line("第1号被保険者数（人）", "見える化", [tot["1号"][y] for y in ys6],
         "令和8年3月末実績からの推計", three=False)
    line("第1号被保険者数（人）", "自前算定",
         [T.POP_DAI10.get(y, {}).get("1号", "") for y in ys6],
         "第10期将来推計用推計人口（コーホート変化率法）", three=False)
    line("要介護（支援）認定者数（人）", "見える化", [tot["認定者数"][y] for y in ys6],
         "性別・5歳階級別・要介護度別に認定率を推計", three=False)
    style_head(ws)
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").startswith("【"):
            ws.cell(r, 1).font = Font(bold=True, size=9)
            ws.cell(r, 1).fill = SUB
    num(ws, range(3, 10))
    body(ws, wrap=(10,))
    widths(ws, [26, 12, 13, 13, 13, 13, 13, 13, 14, 46])
    ws.freeze_panes = "C2"
    return ws


def sheet_service(wb, rows):
    ws = wb.create_sheet("02_サービス別給付費の突合")
    ws.append(["種別", "区分", "サービス",
               "見 令和9", "見 令和10", "見 令和11", "見 3年計",
               "自 令和9", "自 令和10", "自 令和11", "自 3年計",
               "差（見－自）", "差率", "見 令和8基準／令和7実績", "所見"])
    for r in sorted(rows, key=lambda x: -abs(sum(x["見_給付"]) - sum(x["自_給付"]))):
        m, o = sum(r["見_給付"]), sum(r["自_給付"])
        if max(m, o) < 500:
            continue
        rat = (r["見_R8"] / r["見_R7"]) if r["見_R7"] else None
        note = ""
        if rat is not None and (rat > 1.4 or rat < 0.7):
            note = "見える化の令和8年度基準値が令和7年度実績と大きく違う。04参照"
        elif abs(m - o) > 20000:
            note = "基準期間の置き方の違いが効いている。02の下の注参照"
        ws.append([r["種別"], r["区分"], r["算定名"]] + r["見_給付"] + [m]
                  + r["自_給付"] + [o, m - o, (m / o - 1) if o else None, rat, note])
    style_head(ws)
    for r in range(2, ws.max_row + 1):
        d = ws.cell(r, 12).value
        if isinstance(d, (int, float)) and abs(d) > 20000:
            ws.cell(r, 12).fill = NG
        elif isinstance(d, (int, float)) and abs(d) > 5000:
            ws.cell(r, 12).fill = WARN
    num(ws, list(range(4, 13)))
    num(ws, [13, 14], fmt="0.0%")
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 14).number_format = "0.00"
    body(ws, wrap=(15,))
    widths(ws, [5, 10, 30] + [11] * 9 + [9, 13, 52])
    ws.freeze_panes = "D2"
    ws.append([])
    ws.append(["※ 「見」＝見える化ワークシート（将来推計1）。「自」＝自前算定（回数日数ベース見込量）。"])
    ws.append(["※ 認知症対応型共同生活介護・介護予防福祉用具貸与は、この突合にあわせて"
               "基準期間を令和7年度単年に改めた。見える化との差はその分縮んでいる。"])
    return ws


def sheet_ryo(wb, rows):
    ws = wb.create_sheet("03_量と利用者数の突合")
    ws.append(["種別", "サービス", "単位",
               "見 令和7 量", "自 令和7 量", "見／自",
               "見 令和9 量", "自 令和9 量",
               "見 令和7 人数", "自 令和7 件数", "見 令和9 人数", "自 令和9 件数",
               "所見"])
    for r in rows:
        mv7, mv9 = r["見_量"][0], r["見_量"][2]
        ov7 = r["自_量"][0] if r["自_量"] else None
        ov9 = r["自_量"][1] if r["自_量"] else None
        mn7, mn9 = r["見_人数"][0], r["見_人数"][2]
        on7, on9 = r["自_人数"][0], r["自_人数"][1]
        if not any(isinstance(v, (int, float)) and v for v in (mv7, ov7, mn7, on7)):
            continue
        ratio = (mv7 / ov7) if (mv7 and ov7) else None
        note = ""
        if ratio and ratio > 1.15:
            note = ("見える化は給付実績の回数、自前は月報の実日数。"
                    "訪問系は1日に複数回入るため回数が日数を上回る")
        elif ratio and ratio < 0.85:
            note = "量の取り方が違う。04で令和8年度の月報と照合"
        ws.append([r["種別"], r["算定名"], r["単位"], mv7, ov7, ratio,
                   mv9, ov9, mn7, on7, mn9, on9, note])
    style_head(ws)
    num(ws, [4, 5, 7, 8, 9, 10, 11, 12], fmt="#,##0.0")
    num(ws, [6], fmt="0.00")
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, 6).value
        if isinstance(v, (int, float)) and (v > 1.15 or v < 0.85):
            ws.cell(r, 6).fill = WARN
    body(ws, wrap=(13,))
    widths(ws, [5, 30, 6] + [12] * 9 + [56])
    ws.freeze_panes = "C2"
    ws.append([])
    ws.append(["※ 量・人数はいずれも1月当たり。自前の「件数」は年報の年間件数÷12で、"
               "見える化の「人数（1月当たり利用者数）」とほぼ同義。"])
    ws.append(["※ **回数の出所は見える化の方が正しい。**自前は国保連月報の実日数を回数に"
               "読み替えていたため、訪問介護で2割強、訪問看護で1割強、回数を少なく見ていた。"])
    return ws


def sheet_base(wb, rows, geppo):
    ws = wb.create_sheet("04_令和8年度基準値の点検")
    ws.append(["種別", "サービス", "見 令和6", "見 令和7", "見 令和8",
               "見 令和8／令和7", "月報 令和7（月平均給付額・千円）",
               "月報 令和8（同・3か月）", "月報 令和8／令和7", "見／月報のずれ",
               "見える化の令和8年度基準値", "月報が示す令和8年度の動き", "所見"])
    for r in sorted(rows, key=lambda x: -max(x["見_R7"], x["見_R8"])):
        if max(r["見_R6"], r["見_R7"], r["見_R8"]) < 1000:
            continue
        g = geppo.get(r["見える化名"]) or geppo.get(r["算定名"])
        g7 = (g.get(2025, (0, 0))[1] / 1000) if g else None
        g8 = (g.get(2026, (0, 0))[1] / 1000) if g else None
        gr = (g8 / g7) if (g7 and g8) else None
        mr = (r["見_R8"] / r["見_R7"]) if r["見_R7"] else None
        dev = (mr / gr - 1) if (mr and gr) else None
        if mr is None:
            hantei, ugoki, shoken = "", "", "令和7年度に実績がない"
        elif gr is None:
            hantei = "要確認" if (mr > 1.4 or mr < 0.7) else "使える"
            ugoki = "月報で照合できない"
            shoken = ("住宅改修・福祉用具購入は月報の対象外。"
                      "件数が少なく年によって振れるため、町の実績で確かめる")
        else:
            hantei = "使えない" if abs(dev) > 0.20 else "使える"
            if gr > 1.20:
                ugoki = f"増えている（{gr:.2f}倍）"
            elif gr < 0.85:
                ugoki = f"減っている（{gr:.2f}倍）"
            else:
                ugoki = f"横ばい（{gr:.2f}倍）"
            if hantei == "使えない":
                shoken = (f"見える化は{mr:.2f}倍だが月報は{gr:.2f}倍。"
                          f"{dev * 100:+.0f}％ずれている")
            else:
                shoken = "月報と整合する"
            if hantei == "使える" and ugoki.startswith(("増", "減")):
                shoken += "。ただし令和8年度に入って実勢が動いているので町に照会する"
        ws.append([r["種別"], r["算定名"], r["見_R6"], r["見_R7"], r["見_R8"],
                   mr, g7, g8, gr, dev, hantei, ugoki, shoken])
    style_head(ws)
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, 11).value
        ws.cell(r, 11).fill = (NG if v == "使えない" else
                               WARN if v == "要確認" else
                               OK if v == "使える" else PatternFill())
        u = str(ws.cell(r, 12).value or "")
        if u.startswith(("増", "減")):
            ws.cell(r, 12).fill = WARN
    num(ws, [3, 4, 5, 7, 8])
    num(ws, [6, 9], fmt="0.00")
    num(ws, [10], fmt="+0.0%;-0.0%")
    body(ws, wrap=(13,))
    widths(ws, [5, 30, 12, 12, 12, 10, 16, 16, 12, 11, 13, 16, 50])
    ws.freeze_panes = "C2"
    ws.append([])
    ws.append(["※ 月報の令和8年度は審査202605〜202607の3か月平均（給付額）。"
               "見える化の令和8年度も年度途中の月平均を年換算したものなので、"
               "同じ土俵で比べられる。"])
    ws.append(["※ 「使えない」は、見える化の令和8年度／令和7年度の比が、"
               "月報の同じ比と20％以上ずれているもの。"
               "とくに通所リハビリテーションは月報が1.01倍（横ばい）なのに"
               "見える化は2.65倍で、そのまま第10期の見込量にすると"
               "3年間で約8,500万円の過大になる。"])
    ws.append(["※ 「月報が示す令和8年度の動き」が横ばいでないものは、"
               "見える化の基準値が使えるかどうかとは別に、"
               "実勢が動いているので町への照会事項とする。"
               "介護老人保健施設（0.75倍）、特定施設入居者生活介護（0.53倍）、"
               "認知症対応型通所介護（0.71倍）、短期入所生活介護（1.29倍）、"
               "介護予防福祉用具貸与（1.23倍）。"])
    return ws


def sheet_hoken(wb, hoken, P):
    ws = wb.create_sheet("05_保険料算定の入力値")
    ws.append(["項目", "見える化ワークシート", "自前算定", "差", "扱い"])
    h = sum(P["集計"]["標準給付費"])
    g = sum(P["集計"]["計"])
    rows = [
        ("総給付費（3年計・千円）", 3101295, g, "見える化の方が1.9％高い"),
        ("特定入所者介護サービス費等（千円）", 0, sum(P["集計"]["特定入所"]),
         "**見える化は未入力**。総給付費の4.7％（令和5〜7年度実績）"),
        ("高額介護サービス費等（千円）", 0, sum(P["集計"]["高額"]),
         "**見える化は未入力**。同2.3％"),
        ("高額医療合算介護サービス費等（千円）", 0, sum(P["集計"]["高額合算"]),
         "**見える化は未入力**。同0.25％"),
        ("算定対象審査支払手数料（千円）", 0, sum(P["集計"]["手数料"]),
         "**見える化は未入力**。件数×60円の仮置き。"
         "見える化の参考係数では単価の上限が95円とされている"),
        ("標準給付費見込額（千円）", 3101295, h,
         "見える化は総給付費だけ。その他給付費を入れると自前と1.9％差になる"),
        ("地域支援事業費（千円）", 0, CHIIKI_R6_SOGO * 3 / 1000
         + CHIIKI_R6_HOKATSU * 3 / 1000,
         "**見える化は3シートとも全部ゼロ**。年報 様式4 の令和6年度決算を置いた"),
        ("　うち総合事業費（千円）", 0, CHIIKI_R6_SOGO * 3 / 1000,
         "調整交付金の算定対象になる"),
        ("　うち包括的支援事業・任意事業費（千円）", 0, CHIIKI_R6_HOKATSU * 3 / 1000,
         "調整交付金の算定対象にならない"),
        ("第1号被保険者負担割合", 0.23, 0.23, "一致。国の基本指針による"),
        ("第1号被保険者数（3年計・人）", hoken["第1号被保険者数"],
         sum(T.POP_DAI10[y]["1号"] for y in YEARS),
         "見える化10,275人、自前10,284人。差9人（0.09％）"),
        ("所得段階別加入割合補正係数", 0, hosei_ninzu()[0] / sum(
            T.POP_DAI10[y]["1号"] for y in YEARS),
         "**見える化は未入力**。年報 様式1 所得段階別（令和7年度末）の構成比による"),
        ("調整交付金見込交付割合（％）", 0, CHOSEI_MIKOMI,
         "**見える化は未入力**。（第1号負担割合＋5％）－第1号負担割合"
         "×後期高齢者加入割合補正係数×所得段階別加入割合補正係数で求めた値。"
         "後期高齢者加入割合補正係数（令和9年度0.9558ほか）は見える化に入っており、"
         "**所得段階別加入割合補正係数を入れれば見える化が自動で計算する**"),
        ("準備基金取崩額（千円）", 0, KIKIN_TORIKUZUSHI,
         "**見える化は未入力**。年報 様式4 の保有額は12,000千円"),
        ("予定保険料収納率（％）", 0, SHUNO * 100,
         "**見える化は未入力**。年報 様式3 の現年度分99.35％"),
        ("第9期保険料基準額（円／月）", hoken["第9期保険料基準額"], 6609,
         "一致。見える化に入っている6,608.9円は町の第9期の値"),
    ]
    for name, m, o, note in rows:
        d = (m - o) if isinstance(m, (int, float)) and isinstance(o, (int, float)) else None
        ws.append([name, m, o, d, note])
    ws.append([])
    ws.append(["【保険料基準額（円／月）の試算】"])
    a = premium(h)
    b = premium(3101295 * (1 + (h - g) / g))
    ws.append(["自前算定による第10期保険料基準額", "", round(a["基準額"]), "",
               "標準給付費＋地域支援事業費（年報 様式4 の令和6年度決算）、"
               "基金取崩35,000千円、収納率99.35％"])
    ws.append(["見える化の総給付費に置き換えた場合", round(b["基準額"]), "",
               round(b["基準額"] - a["基準額"]),
               "その他給付費・地域支援事業費・基金・収納率は自前と同じものを使った"])
    ws.append(["見える化ワークシートのまま", "算定不能", "", "",
               "未入力が多く「5_保険料推計」は #DIV/0! のまま"])
    style_head(ws)
    for r in range(2, ws.max_row + 1):
        if "未入力" in str(ws.cell(r, 5).value or ""):
            ws.cell(r, 2).fill = NG
        if str(ws.cell(r, 1).value or "").startswith("【"):
            ws.cell(r, 1).font = Font(bold=True, size=9)
            ws.cell(r, 1).fill = SUB
    num(ws, [2, 3, 4])
    body(ws, wrap=(5,))
    widths(ws, [38, 18, 18, 14, 62])
    ws.freeze_panes = "B2"
    return ws


def sheet_chiiki(wb, chiiki):
    ws = wb.create_sheet("06_地域支援事業費_年報様式4")
    ws.append(["科目"] + [jp for _c, jp in N.YEARS] + ["令和4〜6年度3年計", "備考"])
    keys = [("介護予防・生活支援サービス事業費", "総合事業。訪問型・通所型・その他・"
                                  "介護予防ケアマネジメントの合計"),
            ("一般介護予防事業費", "**令和6年度に993千円から5,305千円へ5.3倍。"
                            "通いの場の展開が始まったとみられる。町に内訳を確認する**"),
            ("包括的支援事業・任意事業", "地域包括支援センター運営・在宅医療介護連携・"
                              "生活支援体制整備・認知症総合支援・地域ケア会議・任意事業の合計"),
            ("重層的支援体制整備事業保険料操出金", "全年度ゼロ。小野町は重層事業を実施していない"),
            ("保健福祉事業費", "令和5年度から計上。地域支援事業費には含まない"),
            ("歳出合計", "介護保険特別会計（保険事業勘定）の歳出決算額"),
            ("歳入合計", ""),
            ("準備基金保有額", "**令和3年度末0円→令和4年度末120,000千円→令和5年度末12,000千円。"
                        "繰入金は全年度ゼロなのに残高が10分の1になっており内部矛盾がある。"
                        "基金運用状況調書での確認を要する**"),
            ("基金積立金", "令和3年度・令和5年度に各30,000千円"),
            ]
    for k, note in keys:
        vals = [chiiki[jp].get(k) for _c, jp in N.YEARS]
        three = None
        if k not in ("準備基金保有額",):
            v3 = [chiiki[jp].get(k) for jp in ("令和4年度", "令和5年度", "令和6年度")]
            three = sum(v for v in v3 if isinstance(v, (int, float)))
        ws.append([k] + vals + [three, note])
    ws.append([])
    ws.append(["地域支援事業費 計（総合事業＋包括的支援事業・任意事業）"]
              + [(chiiki[jp].get("介護予防・生活支援サービス事業費", 0)
                  + chiiki[jp].get("一般介護予防事業費", 0)
                  + chiiki[jp].get("包括的支援事業・任意事業", 0))
                 for _c, jp in N.YEARS]
              + [sum(chiiki[jp].get("介護予防・生活支援サービス事業費", 0)
                     + chiiki[jp].get("一般介護予防事業費", 0)
                     + chiiki[jp].get("包括的支援事業・任意事業", 0)
                     for jp in ("令和4年度", "令和5年度", "令和6年度")),
                 "第10期の見込みはこの水準を出発点に置く"])
    style_head(ws)
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").startswith("地域支援事業費 計"):
            for c in range(1, 8):
                ws.cell(r, c).font = Font(bold=True, size=9)
                ws.cell(r, c).fill = SUB
        v = ws.cell(r, 6).value      # 令和7年度
        if v == 0:
            ws.cell(r, 6).fill = NG
    num(ws, range(2, 8))
    body(ws, wrap=(8,))
    widths(ws, [34, 15, 15, 15, 15, 15, 17, 60])
    ws.freeze_panes = "B2"
    ws.append([])
    ws.append(["※ 単位：円。年報 様式4「５．介護保険特別会計経理状況（1）保険事業勘定」の"
               "歳出の決算額。**令和7年度は年報そのものが未入力のため全科目ゼロ。**"])
    ws.append(["※ 一般介護予防事業費と包括的支援事業・任意事業費は、この様式4で"
               "令和3年度から令和6年度まで確認できる。見える化システムの"
               "「3_地域支援事業費」シートは全部ゼロで、出所にならない。"])
    ws.append(["※ 総合事業の事業別内訳（訪問型サービスA・通所型サービスA など）は"
               "様式4にはない。事業別に置くには町の事業実績が要る。"])
    return ws


def sheet_input(wb, P):
    A = P["集計"]
    ws = wb.create_sheet("07_見える化への入力事項")
    ws.append(["#", "シート", "入力する項目", "入れる値", "出所", "優先"])
    items = [
        ("3_地域支援事業費", "介護予防・日常生活支援総合事業（1〜34行）",
         "令和6年度決算34,581,047円を事業別に分解した額",
         "年報 様式4＋町の事業実績（事業別内訳は町にしかない）", "高"),
        ("3_地域支援事業費", "包括的支援事業（地域包括支援センターの運営）及び任意事業",
         "令和6年度決算29,335,189円の内訳", "年報 様式4＋町の事業実績", "高"),
        ("3_地域支援事業費", "包括的支援事業（社会保障充実分）の6事業",
         "在宅医療・介護連携／生活支援体制整備／認知症初期集中／認知症地域支援／"
         "認知症サポーター／地域ケア会議の事業費",
         "町の事業実績。年報 様式4 では包括的支援事業・任意事業に合算されている", "高"),
        ("5_保険料推計", "特定入所者介護サービス費等給付額（76行）",
         f"第10期3年計 {sum(A['特定入所']):,.0f}千円",
         f"自前算定（総給付費×{sum(A['特定入所']) / sum(A['計']) * 100:.3f}％）", "高"),
        ("5_保険料推計", "高額介護サービス費等給付額（79行）",
         f"第10期3年計 {sum(A['高額']):,.0f}千円",
         f"自前算定（同{sum(A['高額']) / sum(A['計']) * 100:.3f}％）", "高"),
        ("5_保険料推計", "高額医療合算介護サービス費等給付額（81行）",
         f"{sum(A['高額合算']):,.0f}千円",
         f"自前算定（同{sum(A['高額合算']) / sum(A['計']) * 100:.3f}％）", "高"),
        ("5_保険料推計", "審査支払手数料 単価・件数（83・84行）",
         "単価60円（年報に計上欄がないため仮置き。上限95円）×第10期の件数",
         "国保連への支払実績。**町の決算で確認を要する**", "中"),
        ("5_保険料推計", "準備基金の残高・取崩額（14・15行）",
         "残高12,000千円（年報 様式4）。取崩額は町の方針次第",
         "**残高の系列に内部矛盾があるため基金運用状況調書で要確認**", "高"),
        ("5_保険料推計", "予定保険料収納率（104行）", "99.35％",
         "年報 様式3 の現年度分", "高"),
        ("5_保険料推計", "所得段階別加入割合（115行以下）",
         "第1〜13段階の構成比（令和7年度末実績）", "年報 様式1 所得段階別", "高"),
        ("5_保険料推計", "調整交付金見込額・調整率（92・93行）",
         f"見込交付割合{CHOSEI_MIKOMI:.3f}％（算式。基数は標準給付費見込額のみ）",
         "後期高齢者加入割合補正係数はワークシートに入っている。"
         "**所得段階別加入割合補正係数（99行）を入れれば自動で計算される**", "高"),
        ("2_サービス別給付費", "令和8年度の基準値（通所リハビリテーションほか）",
         "令和7年度実績に置き換える",
         "**04シート参照。通所リハは月報で横ばいなのに見える化は2.65倍**", "高"),
        ("4_施策反映の解説", "施策反映の全体方針ほか（8行以下すべて）",
         "計画素案の基本目標に沿った記述",
         "現在は「出力後に入力」のまま", "中"),
    ]
    for i, it in enumerate(items, start=1):
        ws.append([i] + list(it))
    style_head(ws)
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 6).fill = NG if ws.cell(r, 6).value == "高" else WARN
    body(ws, wrap=(3, 4, 5))
    widths(ws, [4, 22, 40, 44, 50, 6])
    ws.freeze_panes = "A2"
    return ws


# ================================================ 確認結果メモ（docx）

def build_doc(rows, tot, ku, hoken, P, chiiki, geppo):
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt

    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "游明朝"
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(
        __import__("docx").oxml.ns.qn("w:eastAsia"), "游明朝")

    def h(text, level=1):
        p = doc.add_heading(text, level=level)
        for r in p.runs:
            r.font.name = "游ゴシック"
            r._element.rPr.rFonts.set(
                __import__("docx").oxml.ns.qn("w:eastAsia"), "游ゴシック")
        return p

    def para(text):
        return doc.add_paragraph(text)

    def table(header, body_rows, widths_cm=None):
        t = doc.add_table(rows=1, cols=len(header))
        t.style = "Table Grid"
        for i, v in enumerate(header):
            c = t.rows[0].cells[i]
            c.text = str(v)
            for p in c.paragraphs:
                for r in p.runs:
                    r.font.bold = True
                    r.font.size = Pt(9)
        for row in body_rows:
            cells = t.add_row().cells
            for i, v in enumerate(row):
                cells[i].text = "" if v is None else str(v)
                for p in cells[i].paragraphs:
                    if i > 0:
                        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    for r in p.runs:
                        r.font.size = Pt(9)
        return t

    h(f"見える化システムのワークシート（将来推計1）の確認結果", 0)
    para(f"小野町　第10期介護保険事業計画　／　{ASOF_JP}")
    para("町から受領した「第10期介護保険事業（支援）計画策定に向けたワークシート」"
         "（出力日 令和8年9月15日・推計パターン名「将来推計1」）を、"
         "こちらで組んだ見込量算定と突き合わせた結果です。"
         "受領時のお話のとおり、何も調整していない自然体推計の状態でした。")

    h("1　結論", 1)
    para("総給付費の水準は、こちらの算定と1.9％の差で収まっています。"
         "別々の方法で組んだ推計が近い値になったので、どちらも大きくは外していないと"
         "見てよいと思います。")
    para("ただし、そのまま計画に使える状態ではありません。理由は3つです。")
    para("第一に、地域支援事業費が3シートとも全部ゼロです。"
         "保険料の算定には地域支援事業費が要るため、このままでは保険料が出ません。"
         "実際「5_保険料推計」は #DIV/0! のままです。")
    sonota = sum(P["集計"]["標準給付費"]) - sum(P["集計"]["計"])
    para("第二に、標準給付費のうち総給付費以外の部分"
         "（特定入所者介護サービス費・高額介護サービス費・高額医療合算・審査支払手数料）"
         f"が未入力です。こちらの算定では3年計で{sonota:,.0f}千円、標準給付費の"
         f"{sonota / sum(P['集計']['標準給付費']) * 100:.1f}％にあたります。")
    para("第三に、令和8年度を基準年に置いているため、サービスによって基準値が"
         "年間実績と大きく食い違っています。とくに通所リハビリテーションは"
         "令和7年度の2.65倍になっていますが、国保連の月報で確かめると"
         "令和8年度の給付額は月130万円で令和7年度とほぼ同じです。"
         "この基準値をそのまま延ばすと、第10期の3年間で約8,500万円の過大になります。")

    h("2　総給付費の突合", 1)
    body_rows = []
    for i, y in enumerate(YEARS):
        m, o = ku["総給付費"][y], P["集計"]["計"][i]
        body_rows.append([y, f"{m:,.0f}", f"{o:,.0f}", f"{m - o:+,.0f}",
                          f"{(m / o - 1) * 100:+.1f}％"])
    m3 = sum(ku["総給付費"][y] for y in YEARS)
    o3 = sum(P["集計"]["計"])
    body_rows.append(["3年計", f"{m3:,.0f}", f"{o3:,.0f}", f"{m3 - o3:+,.0f}",
                      f"{(m3 / o3 - 1) * 100:+.1f}％"])
    table(["年度", "見える化（千円）", "こちらの算定（千円）", "差", "差率"], body_rows)
    para("")
    para("第1号被保険者数も、見える化の第10期3年計10,275人に対し、"
         "こちらが用いている第10期将来推計用推計人口では10,284人で、差は9人（0.09％）です。"
         "分母となる人口には差がありません。")
    para("いっぽう認定者数は8％前後ずれています。"
         "見える化は令和9年度849人・令和10年度850人・令和11年度841人、"
         "こちらは令和9年度782人・令和10年度781人・令和11年度780人です。")
    table(["年度", "見える化（人）", "こちらの算定（人）", "差", "差率"],
          [[y, f"{m:,}", f"{o:,}", f"{m - o:+,}", f"{(m / o - 1) * 100:+.1f}％"]
           for y, m, o in (("令和9年度", 849, 782), ("令和10年度", 850, 781),
                           ("令和11年度", 841, 780))]
          + [["3年計", "2,540", "2,343", "+197", "+8.4％"]])
    para("")
    para("差の原因は基点の置き方です。見える化の自然体推計は"
         "直近年度の認定率を基点とし、そこに変化分（初期値は令和3年度→令和4年度の変化）を"
         "加えていきます。この直近年度の認定率は令和7年9月の865人によっています。"
         "令和7年9月は、国保連の月報で認定者数が961人から862人へ一段下がった月の直後であり、"
         "その後さらに866人から794人へ減っていく途中の点です。"
         "見える化はこの後半の減少を織り込んでいません。")
    para("こちらの算定は、年齢階層別の認定率（令和7年度実績）を"
         "第10期将来推計用推計人口の年齢階層別人口に乗じており、"
         "令和8年3月末の第1号被保険者の認定者783人と接続します。"
         "認定者数の絶対水準としては、こちらの782人前後が令和8年3月末の実績に沿います。")
    para("ただし、この差が見込量にそのまま出るわけではありません。"
         "見える化の見込量は「第1号被保険者数×認定率×利用率」の3段の掛け算で、"
         "利用率の分母は認定者数（在宅サービスは認定者数から施設・居住系の利用者数を引いたもの）です。"
         "認定率が高く出れば利用率が同じだけ低く出るため、"
         "認定率と利用率を断層の同じ側からそろえて採る限り、積は変わりません。"
         "したがって認定者数の差は、見込量そのものよりも、"
         "計画書に載せる認定者数の表と、認定率を使う指標の側に影響します。"
         "**計画書に載せる認定者数は、令和8年3月末の実績と接続するこちらの系列を採ります。**")
    para("")
    para("なお、見える化のワークシートの令和6年度の列は、"
         "指標によって年度がそろっていません。"
         "給付費の欄は14サービスすべてで年報の令和5年度と比が1.000で一致し、"
         "実体は令和5年度の実績です。"
         "同じ列の認定者数937人は、国保連の令和6年10月審査分（令和6年9月提供分）と一致します。"
         "列見出しの年度をそのまま読むと1年ずれますので、"
         "ワークシートから実績を読み取る際は指標ごとに年度を確かめる必要があります。")

    h("3　サービス別に見た差", 1)
    para("金額の大きい順に、差の理由を整理します。")
    big = sorted(rows, key=lambda x: -abs(sum(x["見_給付"]) - sum(x["自_給付"])))[:8]
    br = []
    for r in big:
        m, o = sum(r["見_給付"]), sum(r["自_給付"])
        br.append([r["算定名"][:22], f"{m:,.0f}", f"{o:,.0f}", f"{m - o:+,.0f}"])
    table(["サービス（第10期3年計）", "見える化（千円）", "こちら（千円）", "差"], br)
    para("")
    para("差が大きいものは、ほとんどが基準期間の置き方の違いによるものです。"
         "見える化は令和8年度の数か月を年換算した1年分だけを見ています。"
         "こちらは令和5〜7年度の3年平均を使っています。")

    h("4　こちらの算定を直したところ", 1)
    para("突合の過程で、こちらの算定に直すべき点が見つかりましたので、"
         "先に見込量算定を改めました。")
    para("令和5年度から令和7年度まで一方向に動いていて、"
         "令和8年度の国保連月報でも同じ向きが続いているサービスは、"
         "3年平均だと直近の実績から離れてしまいます。"
         "次の3つについて、基準期間を令和7年度単年に改めました。")
    table(["サービス", "令和5→7年度", "令和8年度の月報", "3年平均による令和9年度／令和7年度実績"],
          [["認知症対応型共同生活介護", "▲13.6％（単調減）", "0.93倍（減少が継続）", "1.098"],
           ["介護予防認知症対応型共同生活介護", "令和6年度3.1件→令和7年度1.2件",
            "0.86倍（減少が継続）", "1.876"],
           ["介護予防福祉用具貸与", "＋24.6％（単調増）", "1.09倍（増加が継続）", "0.867"]])
    para("")
    para("いっぽう、訪問介護（令和5→7年度▲17.2％の単調減）と"
         "介護予防小規模多機能型居宅介護（同＋41.4％の単調増）は、"
         "令和8年度の月報で向きが反転していますので、3年平均のまま置きました。")
    para(f"この修正で、総給付費の第10期3年計は"
         f"3,089,383千円から{o3:,.0f}千円へ、46,317千円（1.5％）下がりました。")

    h("5　地域支援事業費（年報 様式4 で確認できます）", 1)
    para("一般介護予防事業費と包括的支援事業・任意事業費の決算額は、"
         "介護保険事業状況報告（年報）の様式4「５．介護保険特別会計経理状況」の"
         "歳出側に、独立した科目として計上されています。"
         "受領済みの年報データで令和3年度から令和6年度まで確認できました。"
         "見える化システムからダウンロードする必要はありません。")
    ck = []
    for k in ("介護予防・生活支援サービス事業費", "一般介護予防事業費",
              "包括的支援事業・任意事業"):
        ck.append([k] + [f"{chiiki[jp].get(k, 0):,.0f}" for _c, jp in N.YEARS])
    ck.append(["地域支援事業費 計"]
              + [f"{sum(chiiki[jp].get(k, 0) for k in ('介護予防・生活支援サービス事業費', '一般介護予防事業費', '包括的支援事業・任意事業')):,.0f}"
                 for _c, jp in N.YEARS])
    table(["科目（単位：円）"] + [jp for _c, jp in N.YEARS], ck)
    para("")
    para("令和7年度だけは年報そのものが未入力で、様式4の全科目がゼロです。"
         "したがって町にお願いするのは令和7年度の決算額だけで足ります。")
    para("なお、一般介護予防事業費が令和5年度の993,702円から"
         "令和6年度の5,304,726円へ5.3倍に増えています。"
         "計画素案では「令和5年度で1.0百万円と低い水準」と書いていましたので、"
         "令和6年度の実績に置き換えます。何に使われたのかを教えていただけると、"
         "第10期の一般介護予防事業の組み立てに反映できます。")
    para("それから、様式4の介護給付費準備基金保有額が、"
         "令和3年度末0円・令和4年度末120,000千円・令和5年度末12,000千円・"
         "令和6年度末12,000千円となっています。"
         "繰入金（取崩し）は全年度ゼロですので、残高が10分の1になるのは説明がつきません。"
         "基金運用状況調書などで正しい残高をご確認いただけますでしょうか。"
         "準備基金の取崩額は保険料の基準額に直接効きます。")

    h("6　保険料の試算", 1)
    para("上の修正を反映し、地域支援事業費に年報 様式4 の令和6年度決算を置いて算定すると、"
         "第10期の保険料基準額は次のようになります。"
         "第1号被保険者負担割合と調整交付金の見込交付割合はまだ国から示されていないため、"
         "いずれも第9期と同じ値で置いた暫定の数字です。")
    aa = premium(sum(P["集計"]["標準給付費"]))
    tbl = []
    for kikin, lab in ((0, "準備基金を取り崩さない場合"),
                       (12000, "準備基金12,000千円を全額取り崩す場合"),
                       (35000, "第9期と同じ35,000千円を取り崩す場合")):
        r = T.premium(sum(P["集計"]["標準給付費"]), CHIIKI_R6_SOGO * 3 / 1000
                      + CHIIKI_R6_HOKATSU * 3 / 1000, CHIIKI_R6_SOGO * 3 / 1000,
                      sum(T.POP_DAI10[y]["1号"] for y in YEARS), kikin=kikin)
        tbl.append([lab, f"{r['月額']:,.0f}円",
                    f"{r['月額'] - T.DAI9_KIJUN:+,.0f}円"])
    table(["ケース", "保険料基準額（月額）", "第9期6,600円との差"], tbl)
    para("")
    para("いずれも第9期の6,600円を下回ります。"
         "ただし第1号被保険者負担割合が24％になると"
         f"{T.premium(sum(P['集計']['標準給付費']), CHIIKI_R6_SOGO * 3 / 1000 + CHIIKI_R6_HOKATSU * 3 / 1000, CHIIKI_R6_SOGO * 3 / 1000, sum(T.POP_DAI10[y]['1号'] for y in YEARS), futan=0.24)['月額']:,.0f}円となり、"
         "6,600円を上回ります。保険料の水準を決めるのは給付費の精度ではなく、"
         "第1号被保険者負担割合と調整交付金の見込交付割合という制度側の2つです。")

    h("7　見える化のワークシートに入力していただきたいもの", 1)
    para("このワークシートで保険料まで出すには、次の入力が要ります。"
         "町のほうで入力される場合は、こちらで用意した値をお渡しできます。")
    table(["シート", "入力する項目", "こちらで用意できる値"],
          [["3_地域支援事業費", "総合事業・包括的支援事業・任意事業の各事業費",
            "年報 様式4 の令和6年度決算（事業別の内訳は町の資料が要ります）"],
           ["5_保険料推計", "特定入所者・高額・高額医療合算・審査支払手数料",
            f"第10期3年計で{sonota:,.0f}千円"],
           ["5_保険料推計", "準備基金の残高と取崩額", "残高12,000千円（要確認）"],
           ["5_保険料推計", "予定保険料収納率", "99.35％（年報 様式3）"],
           ["5_保険料推計", "所得段階別加入割合", "令和7年度末の第1〜13段階の構成比"],
           ["5_保険料推計", "調整交付金見込額",
            f"見込交付割合{CHOSEI_MIKOMI:.3f}％"],
           ["4_施策反映の解説", "施策反映の全体方針ほか", "計画素案の基本目標に沿って記述"]])

    h("8　お願いしたい確認事項", 1)
    for i, t in enumerate([
        "通所リハビリテーションについて、令和8年度に事業所や利用の状況が変わったことは"
        "ありますか。見える化の令和8年度が令和7年度の2.65倍になっていますが、"
        "国保連の月報では給付額も件数も横ばいです。",
        "短期入所生活介護が令和8年度に入って増えています。"
        "月報で令和7年度の月平均2,673千円が令和8年度は3,456千円（＋29％）、"
        "日数も月300日から380日に伸びています。"
        "特別養護老人ホームの入所待機が増えているなど、心当たりはありますか。"
        "この水準が続く場合、第10期の給付費は年1,280万円ほど上振れします。",
        "認知症対応型共同生活介護の利用が減り続けています"
        "（令和6年度50.1件／月→令和7年度44.6件／月→令和8年度41.3件／月）。"
        "空床が出ているのか、待機があるのに入れていないのか、状況を教えてください。",
        "介護老人保健施設の給付額が、月報で令和7年度の月8,006千円から"
        "令和8年度は月5,973千円（▲25％）に落ちています。"
        "退所が続いているのか、入所先が変わっているのか、お分かりになりますか。",
        "特定施設入居者生活介護も月報で月766千円から月409千円（▲47％）に落ちています。"
        "利用者数が5人から半減した形です。",
        "認知症対応型通所介護が月報で月990千円から月707千円（▲29％）に落ちています。",
        "介護予防福祉用具貸与は逆に増え続けています"
        "（令和6年度50.6件／月→令和7年度58.1件／月→令和8年度63.0件／月）。",
        "令和6年度の一般介護予防事業費が前年の5.3倍（993千円→5,305千円）です。"
        "何の事業を始められたのでしょうか。",
        "介護給付費準備基金の年度末残高の推移をご教示ください（年報の値に矛盾があります）。",
        "令和7年度の年報（介護保険事業状況報告）の入力予定をご教示ください。",
    ], start=1):
        doc.add_paragraph(f"{i}　{t}")

    path = OUT / f"小野町_見える化ワークシートの確認結果_{ASOF}.docx"
    doc.save(path)
    return path


# ================================================ main

def main():
    rows, tot, ku, hoken, P = compare()
    chiiki = read_chiiki()
    geppo = read_geppo_year()

    facts = [
        ("総給付費", "見える化の第10期3年計3,101,295千円に対し、こちらの算定は"
                 f"{sum(P['集計']['計']):,.0f}千円。差は1.9％",
         "見える化ワークシート 1_推計値サマリ／自前算定", "採用（相互に検証できた）"),
        ("地域支援事業費", "見える化の3_地域支援事業費シートは全項目ゼロ。出所にならない",
         "見える化ワークシート 3_地域支援事業費", "不採用（年報 様式4 を使う）"),
        ("地域支援事業費", "一般介護予防事業費・包括的支援事業・任意事業費とも、"
                    "年報 様式4 の歳出で令和3〜6年度が確認できる",
         "年報 様式4 ５．介護保険特別会計経理状況", "採用"),
        ("地域支援事業費", "令和6年度の一般介護予防事業費が5,304,726円。前年の5.3倍",
         "年報 様式4（令和6年度）", "採用（素案の記述を差し替える）"),
        ("標準給付費", "見える化の標準給付費見込額は総給付費と同額。"
                  "特定入所者・高額・高額医療合算・審査支払手数料が未入力",
         "見える化ワークシート 5_保険料推計 68〜85行", "不採用（自前の241,112千円を使う）"),
        ("保険料", "見える化の5_保険料推計は #DIV/0!。基金・収納率・所得段階別加入割合が未入力",
         "見える化ワークシート 5_保険料推計 12行", "不採用"),
        ("基準年", "見える化の令和8年度は数か月の月平均を年換算したもの。"
                "通所リハビリテーションは令和7年度の2.65倍だが、月報では横ばい",
         "見える化 2_サービス別給付費 83行／国保連月報 THRK0401", "不採用（令和7年度で置く）"),
        ("基準年", "短期入所生活介護は見える化が1.65倍。月報でも1.29倍で、実勢の増加",
         "国保連月報 THRK0401・THRK0201", "要確認（町に照会）"),
        ("量の単位", "回数は見える化の方が正しい。自前は月報の実日数を回数に読み替えていたため、"
                 "訪問介護で2割強少なく見ていた",
         "見える化 2_サービス別給付費／国保連月報 THRK0201", "採用（見える化の回数を使う）"),
        ("見込量算定", "認知症対応型共同生活介護・介護予防認知症対応型共同生活介護・"
                  "介護予防福祉用具貸与の基準期間を令和7年度単年に改めた",
         "年報 様式2／国保連月報の令和8年度3か月", "採用（自前算定を修正済み）"),
        ("人口", "見える化の第1号被保険者数（3年計10,275人）と"
               "第10期将来推計用推計人口（同10,284人）の差は9人",
         "見える化 5_保険料推計 109行／推計人口原本", "採用（推計人口原本で統一）"),
        ("準備基金", "年報 様式4 の保有額が令和4年度末120,000千円→令和5年度末12,000千円。"
                 "繰入金は全年度ゼロで、内部矛盾がある",
         "年報 様式4（令和4〜6年度）", "要確認（基金運用状況調書）"),
    ]

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    sheet_yoten(wb, facts)
    sheet_soukatsu(wb, tot, ku, P)
    sheet_service(wb, rows)
    sheet_ryo(wb, rows)
    sheet_base(wb, rows, geppo)
    sheet_hoken(wb, hoken, P)
    sheet_chiiki(wb, chiiki)
    sheet_input(wb, P)
    xp = OUT / f"小野町_第10期_見える化ワークシートとの突合_{ASOF}.xlsx"
    wb.save(xp)

    dp = build_doc(rows, tot, ku, hoken, P, chiiki, geppo)
    print("出力:", xp)
    print("出力:", dp)


if __name__ == "__main__":
    main()
