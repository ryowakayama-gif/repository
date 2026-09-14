"""小野町 第10期 サービス見込量を回数・日数ベースで算定する。

令和8年9月14日に受領した国保連月報のExcel版により、サービス種類別×要介護度別の
日数が月次で取れるようになった。これを用いて、計画書の見込量表に載せる
回数・日数ベースの見込量を組む。

**量の出し方**

  見込量(年) ＝ 見込件数(年) × 基準年度の1件当たり日数

  見込件数は年報から組んだ試算（build_ono_tanka）による。第1号被保険者1人当たりの
  利用率を令和5〜7年度の平均で置き、第10期将来推計用推計人口を乗じたもの。
  1件当たり日数は国保連月報の令和7年度（審査202505〜202604の12か月）による。

**単位の考え方**

  国保連月報の「日数」は各サービス種類ごとの実日数である（帳票の注記）。
  計画書が回数で表記するサービスについて、日数がそのまま回数になるかを
  1日当たり単位数で確かめた。

    訪問介護      429単位/日（身体介護30分以上1時間未満396単位）
    訪問入浴介護 1,373単位/日（1,266単位/回）
    訪問看護      876単位/日（訪問看護ステーション30分以上1時間未満823単位）
    通所介護      941単位/日（通常規模7〜8時間・要介護3で883単位）

  いずれも1回分の報酬をわずかに上回る水準にとどまり、1日1回が主体である。
  **したがって訪問系は日数を回数とみなせるが、厳密には回数は日数を1割程度上回る。**
  計画書に回数で載せる場合はこの点を注記する。

**給付費は年報を使う**

  月報の3区分計（居宅・地域密着型・施設）は年報の様式2 総計に対し、令和6年度・
  令和7年度とも▲1.0％である。差が2年度とも同じ幅であり系統的だが、理由が
  特定できていない。このため月報は1件当たり日数を出すためだけに用い、
  量の水準と給付費は年報に合わせる。

出力: 04_算定・見込量/小野町_第10期_回数日数ベース見込量_YYYYMMDD.xlsx
"""

import collections
import io
import pathlib
import re
import warnings
import zipfile

warnings.filterwarnings("ignore")

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill
from openpyxl.styles.borders import Side
from openpyxl.utils import get_column_letter

import build_ono_tanka as T

ROOT = pathlib.Path(__file__).parent / "小野町_引継ぎ_整理済"
ZDIR = ROOT / "18_年報・国保連データ" / "原本_国保連月報"
OUT = ROOT / "04_算定・見込量"
ASOF = "20260914"
ASOF_JP = "令和8年9月14日"

HEAD = PatternFill("solid", fgColor="1F3864")
KEY = PatternFill("solid", fgColor="FCE4E4")
CALC = PatternFill("solid", fgColor="EAF1FB")
WARN = PatternFill("solid", fgColor="FFE0B2")
OK = PatternFill("solid", fgColor="C6EFCE")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FORMS = {"THRK0101": "件数", "THRK0201": "日数", "THRK0301": "単位数",
         "THRK0401": "給付額", "THRK0501": "費用額"}
COLS = {"計": 6, "事業対象者": 7, "要支援1": 8, "要支援2": 9,
        "要介護1": 10, "要介護2": 11, "要介護3": 12, "要介護4": 13, "要介護5": 14}

# 月報のサービスコード → 試算の正規化名と給付の別
#   同じ区分に複数コードが対応するもの（短期利用など）は合算する。
MAP = {
    "11訪問介護": ("訪問介護", "介護"),
    "61介護予防訪問介護": ("訪問介護", "予防"),
    "12訪問入浴介護": ("訪問入浴介護", "介護"),
    "62介護予防訪問入浴介護": ("訪問入浴介護", "予防"),
    "13訪問看護": ("訪問看護", "介護"),
    "63介護予防訪問看護": ("訪問看護", "予防"),
    "14訪問リハビリテーション": ("訪問リハビリテーション", "介護"),
    "64介護予防訪問リハビリテーション": ("訪問リハビリテーション", "予防"),
    "15通所介護": ("通所介護", "介護"),
    "65介護予防通所介護": ("通所介護", "予防"),
    "16通所リハビリテーション": ("通所リハビリテーション", "介護"),
    "66介護予防通所リハビリテーション": ("通所リハビリテーション", "予防"),
    "17福祉用具貸与": ("福祉用具貸与", "介護"),
    "67介護予防福祉用具貸与": ("福祉用具貸与", "予防"),
    "21短期入所生活介護": ("短期入所生活介護", "介護"),
    "24介護予防短期入所生活介護": ("短期入所生活介護", "予防"),
    "22短期入所療養介護（老健）": ("短期入所療養介護（介護老人保健施設）", "介護"),
    "25介護予防短期入所療養介護（老健）": ("短期入所療養介護（介護老人保健施設）", "予防"),
    "2A短期入所療養介護（介護医療院）": ("短期入所療養介護（介護医療院）", "介護"),
    "2B介護予防短期入所療養介護（介護医療院）": ("短期入所療養介護（介護医療院）", "予防"),
    "23短期入所療養介護（病院等）": ("短期入所療養介護（病院等）", "介護"),
    "26介護予防短期入所療養介護（病院等）": ("短期入所療養介護（病院等）", "予防"),
    "31居宅療養管理指導": ("居宅療養管理指導", "介護"),
    "34介護予防居宅療養管理指導": ("居宅療養管理指導", "予防"),
    "33特定施設入居者生活介護": ("特定施設入居者生活介護", "介護"),
    "27特定施設入居者生活介護（短期）": ("特定施設入居者生活介護", "介護"),
    "35介護予防特定施設入居者生活介護": ("特定施設入居者生活介護", "予防"),
    "43居宅介護支援": ("居宅介護支援・介護予防支援", "介護"),
    "46介護予防支援": ("居宅介護支援・介護予防支援", "予防"),
    "76定期巡回・随時対応型訪問介護看護": ("定期巡回・随時対応型訪問介護看護", "介護"),
    "71夜間対応型訪問介護": ("夜間対応型訪問介護", "介護"),
    "78地域密着型通所介護": ("地域密着型通所介護", "介護"),
    "72認知症対応型通所介護": ("認知症対応型通所介護", "介護"),
    "74介護予防認知症対応型通所介護": ("認知症対応型通所介護", "予防"),
    "73小規模多機能型居宅介護": ("小規模多機能型居宅介護", "介護"),
    "68小規模多機能型居宅介護（短期）": ("小規模多機能型居宅介護", "介護"),
    "75介護予防小規模多機能型居宅介護": ("小規模多機能型居宅介護", "予防"),
    "69介護予防小規模多機能型居宅介護（短期）": ("小規模多機能型居宅介護", "予防"),
    "32認知症対応型共同生活介護": ("認知症対応型共同生活介護", "介護"),
    "38認知症対応型共同生活介護(短期)": ("認知症対応型共同生活介護", "介護"),
    "37介護予防認知症対応型共同生活介護": ("認知症対応型共同生活介護", "予防"),
    "39介護予防認知症型共同生活介護(短期)": ("認知症対応型共同生活介護", "予防"),
    "36地域密着型特定施設入居者生活介護": ("地域密着型特定施設入居者生活介護", "介護"),
    "28地域密着型特定施設入居者生活介護（短期）": ("地域密着型特定施設入居者生活介護", "介護"),
    "54地域密着型介護老人福祉施設入所者生活介護":
        ("地域密着型介護老人福祉施設入所者生活介護", "介護"),
    "77複合型サービス（看護小規模多機能型居宅介護）": ("看護小規模多機能型居宅介護", "介護"),
    "79複合型サービス（看護小規模多機能型居宅介護・短期）": ("看護小規模多機能型居宅介護", "介護"),
    "51介護福祉施設サービス": ("介護老人福祉施設", "介護"),
    "52介護保健施設サービス": ("介護老人保健施設", "介護"),
    "55介護医療院サービス": ("介護医療院", "介護"),
    "53介護療養施設サービス": ("介護療養型医療施設", "介護"),
}

# 計画書に載せるときの量の単位。
#   回 … 訪問系（月報は実日数。1日1回が主体で日数≒回数）
#   日 … 通所・短期入所（1日1回のため日数＝回数でもある）
#   人 … 貸与・購入・改修・居住系・施設・支援（量は人数のみ）
UNIT = {
    "訪問介護": "回", "訪問入浴介護": "回", "訪問看護": "回",
    "訪問リハビリテーション": "回",
    "通所介護": "日", "通所リハビリテーション": "日",
    "短期入所生活介護": "日",
    "短期入所療養介護（介護老人保健施設）": "日",
    "短期入所療養介護（病院等）": "日",
    "短期入所療養介護（介護医療院）": "日",
    "居宅療養管理指導": "人", "福祉用具貸与": "人",
    "特定福祉用具販売": "人", "住宅改修": "人",
    "特定施設入居者生活介護": "人", "居宅介護支援・介護予防支援": "人",
    "定期巡回・随時対応型訪問介護看護": "人", "夜間対応型訪問介護": "人",
    "地域密着型通所介護": "日", "認知症対応型通所介護": "日",
    "小規模多機能型居宅介護": "人", "認知症対応型共同生活介護": "人",
    "地域密着型特定施設入居者生活介護": "人",
    "地域密着型介護老人福祉施設入所者生活介護": "人",
    "看護小規模多機能型居宅介護": "人",
    "介護老人福祉施設": "人", "介護老人保健施設": "人",
    "介護療養型医療施設": "人", "介護医療院": "人",
}

# 計画書に載せるときの正式なサービス名。予防給付は「介護予防」が付く。
YOBO_NAME = {
    "特定福祉用具販売": "特定介護予防福祉用具販売",
    "居宅介護支援・介護予防支援": "介護予防支援",
}
KAIGO_NAME = {"居宅介護支援・介護予防支援": "居宅介護支援"}


def disp(svc, kind):
    """計画書に載せるサービス名を返す。"""
    if kind == "介護":
        return KAIGO_NAME.get(svc, svc)
    if svc in YOBO_NAME:
        return YOBO_NAME[svc]
    return "介護予防" + svc


# 町内の定員（令和7年度。見える化システム Ｄ25〜Ｄ27）
# 「人」は入所・登録の定員、「人/日」は1日あたりの利用定員。
TEIIN = {
    "介護老人福祉施設": (54, "人"),
    "介護老人保健施設": (0, "人"),
    "介護療養型医療施設": (0, "人"),
    "介護医療院": (0, "人"),
    "地域密着型介護老人福祉施設入所者生活介護": (58, "人"),
    "認知症対応型共同生活介護": (98, "人"),
    "特定施設入居者生活介護": (0, "人"),
    "地域密着型特定施設入居者生活介護": (0, "人"),
    "通所介護": (139, "人/日"),
    "地域密着型通所介護": (0, "人/日"),
    "通所リハビリテーション": (0, "人/日"),
    "認知症対応型通所介護": (12, "人/日"),
    "小規模多機能型居宅介護": (30, "人/日（通い）"),
    "看護小規模多機能型居宅介護": (0, "人/日（通い）"),
}
EIGYO = 26          # 通所系の月あたり営業日数の目安
TESURYO_YEN = T.TESURYO   # 審査支払手数料（円／件）

# 1回あたりの報酬（単位）の目安。日数が回数とみなせるかの確認に用いる。
TANI_ME = {
    "訪問介護": (396, "身体介護30分以上1時間未満"),
    "訪問入浴介護": (1266, "訪問入浴介護1回"),
    "訪問看護": (823, "訪問看護ステーション30分以上1時間未満"),
    "通所介護": (883, "通常規模7時間以上8時間未満・要介護3"),
    "通所リハビリテーション": (762, "通常規模6時間以上7時間未満・要介護3"),
}


# ================================================================ 月報の読み込み

def _names(z):
    out = {}
    for it in z.infolist():
        raw = (it.filename.encode() if it.flag_bits & 0x800
               else it.filename.encode("cp437", "ignore"))
        for e in ("cp932", "utf-8"):
            try:
                out[raw.decode(e)] = it.filename
                break
            except Exception:
                continue
    return out


def fiscal(ym):
    """審査年月から、サービス提供のあった年度を返す（審査は提供の翌月）。"""
    y, m = int(ym[:4]), int(ym[4:])
    y2, m2 = (y, m - 1) if m > 1 else (y - 1, 12)
    return (y2 - 1) if m2 <= 3 else y2


def read_geppo():
    """{(審査年月, 指標): {帳票の行名: {度: 値}}} を返す。"""
    data = {}
    for zp in sorted(ZDIR.glob("*.zip")):
        z = zipfile.ZipFile(zp)
        for shown, real in _names(z).items():
            m = re.search(r"(THRK0[1-5]01)_XLS(\d{6})_", shown)
            if not m:
                continue
            ws = openpyxl.load_workbook(io.BytesIO(z.read(real)),
                                        data_only=True)["印刷帳票1"]
            d = {}
            for r in range(8, ws.max_row + 1):
                lab = next((ws.cell(r, c).value.strip() for c in (2, 3, 4)
                            if isinstance(ws.cell(r, c).value, str)
                            and ws.cell(r, c).value.strip()), None)
                if not lab or "再掲" in lab or lab.startswith("・"):
                    continue
                vals = {k: ws.cell(r, c).value for k, c in COLS.items()}
                if not any(isinstance(v, (int, float)) for v in vals.values()):
                    continue
                d[lab] = {k: (v if isinstance(v, (int, float)) else 0)
                          for k, v in vals.items()}
            data[(m.group(2), FORMS[m.group(1)])] = d
    return data


def aggregate(data, fy):
    """年度分を正規化名×給付の別に集計する。"""
    yms = sorted(ym for ym in {k[0] for k in data} if fiscal(ym) == fy)
    out = collections.defaultdict(lambda: collections.defaultdict(float))
    raw = collections.defaultdict(lambda: collections.defaultdict(float))
    for ym in yms:
        for metric in ("件数", "日数", "単位数", "給付額", "費用額"):
            d = data.get((ym, metric), {})
            for code, (name, kind) in MAP.items():
                v = d.get(code, {}).get("計", 0)
                out[(name, kind)][metric] += v
                raw[code][metric] += v
        for k in ("合計", "居宅サービス計", "地域密着型サービス計",
                  "施設サービス計", "総合事業計"):
            raw[k]["給付額"] += data.get((ym, "給付額"), {}).get(k, {}).get("計", 0)
    return len(yms), out, raw


# ================================================================ 体裁

def style_header(ws, row=1):
    for c in ws[row]:
        if c.value is not None:
            c.font = Font(bold=True, size=9, color="FFFFFF")
            c.fill = HEAD
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = BORDER


def body(ws, first=2, wrap=()):
    for row in ws.iter_rows(min_row=first):
        for c in row:
            if c.value is not None:
                c.border = BORDER
            c.font = Font(size=c.font.size or 9, bold=bool(c.font.bold))
            c.alignment = Alignment(vertical="top", wrap_text=(c.column in wrap))


def widths(ws, ws_w):
    for i, w in enumerate(ws_w, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def notes(ws, lines):
    ws.append([])
    for t in lines:
        ws.append([t])
        ws.cell(ws.max_row, 1).font = Font(size=9, italic=True)


# ================================================================ シート

def sheet_intro(wb, a6, a7, chk):
    ws = wb.create_sheet("00_この試算の読み方")
    rows = [
        ["小野町 第10期介護保険事業計画　サービス見込量（回数・日数ベース）"],
        [f"作成 {ASOF_JP}／出典 国保連合会業務統計表（月報）令和6・7年度、"
         "介護保険事業状況報告（年報）令和3〜7年度"],
        [],
        ["■ なぜ今これができるようになったか"],
        ["", "令和8年9月14日に国保連月報のExcel版を受領した。"
             "既に受領していたCSV版は見出しを持たず行番号で読む必要があったため、"],
        ["", "**日数の帳票（THRK0201 サービス種類別給付状況その2）を使えていなかった。**"],
        ["", "Excel版にはサービス種類別×要介護度別の日数が月次で入っており、"
             "令和6・7年度がいずれも12か月そろっている。"],
        [],
        ["■ 量の出し方"],
        ["", "見込量（年） ＝ 見込件数（年） × 基準年度の1件当たり日数"],
        ["", "見込件数は年報から組んだ試算による（第1号被保険者1人当たりの利用率を"
             "令和5〜7年度平均で置き、第10期将来推計用推計人口を乗じたもの）。"],
        ["", "1件当たり日数は国保連月報の令和7年度（審査202505〜202604の12か月）による。"],
        [],
        ["■ 単位の考え方"],
        ["", "国保連月報の「日数」は各サービス種類ごとの実日数である（帳票の注記）。"],
        ["", "計画書が回数で表記するサービスについて、日数をそのまま回数とみなせるかを"
             "1日当たり単位数で確かめた。"],
        ["", "サービス", "1日当たり単位数", "1回あたりの報酬（目安）", "比", "判定"],
    ]
    for svc, (tani, memo) in TANI_ME.items():
        d = a7.get((svc, "介護"), {})
        if not d.get("日数"):
            continue
        per = d["単位数"] / d["日数"]
        rows.append(["", svc, round(per), f"{tani}単位（{memo}）",
                     f"{per / tani:.2f}", "1日1回が主体" if per / tani < 1.3 else "要確認"])
    rows += [
        ["", "**いずれも1回分の報酬をわずかに上回る水準にとどまる。1日1回が主体である。**"],
        ["", "**したがって訪問系は日数を回数とみなせるが、厳密には回数は日数を1割程度上回る。**"],
        ["", "計画書に回数で載せる場合は、この点を注記する。"],
        [],
        ["■ 給付費は年報を使う"],
        ["", "月報の3区分計（居宅・地域密着型・施設）を年報の様式2 総計と突き合わせた。"],
        ["", "年度", "月報の3区分計", "年報 様式2 総計", "差", "差の率"],
    ]
    for fy, lab in ((2024, "令和6年度"), (2025, "令和7年度")):
        a, b = chk[fy]
        rows.append(["", lab, a, b, a - b, f"{(a / b - 1) * 100:+.2f}％"])
    rows += [
        ["", "**両年度とも▲1.0％で、差の幅がそろっている。系統的な差だが理由が特定できていない。**"],
        ["", "このため月報は1件当たり日数を出すためだけに用い、量の水準と給付費は年報に合わせる。"],
        ["", "差の理由（過誤調整の反映時期、住所地特例の扱い等）は国保連に照会する。"],
        [],
        ["■ 趨勢のあるサービスの扱い（判断を要する2件）"],
        ["", "見込件数は令和5〜7年度の3年平均の利用率で置いている。"
             "一方向に動いてきたサービスでは、平均に引き戻される。"],
        ["", "令和7年度実績から15％を超えて動くのは次の2件だけで、残りは±15％に収まっている。"],
        ["", "サービス", "R5", "R6", "R7", "R11見込（3年平均）", "R11見込（R7単年）", "判断"],
    ]
    for svc, kind in (("訪問介護", "介護"), ("訪問看護", "予防")):
        i = 1 if kind == "介護" else 0
        v = [chk["act"][y]["件数"].get(svc, (0, 0))[i]
             for y in ("令和5年度", "令和6年度", "令和7年度")]
        r11 = chk["proj"][svc, kind]
        alt = v[2] / chk["pop7"] * T.POP_EST["令和11年度"]["1号"]
        rows.append(["", f"{svc}（{kind}給付）", *v, round(r11), round(alt),
                     "**一方向に動いている。どちらで置くかを協議会に諮る**"])
    rows += [
        ["", "**訪問介護は958件→843件→628件と減り続けており、"
             "3年平均だと令和11年度に795件（＋26.6％）となる。**"
             "令和7年度単年で置けば623件（▲0.8％）である。"],
        ["", "**介護予防訪問看護は48件→66件→88件と増え続けており、"
             "3年平均だと令和11年度に66件（▲24.7％）となる。**"
             "令和7年度単年で置けば87件（▲0.8％）である。"],
        ["", "いずれも給付費に占める割合は小さい（訪問介護3.1％、介護予防訪問看護0.2％）が、"
             "計画書に載せる見込量としては説明を要する。"],
        [],
        ["■ この試算で決めていないこと"],
        ["", "1. 訪問系を回数で載せるか日数で載せるか。回数で載せる場合の割増の扱い"],
        ["", "2. 令和9年度の介護報酬改定率。給付費は改定率0.0％で置いている"],
        ["", "3. 施設・居住系の整備方針。整備があれば人数が段差で増える"],
        [],
        ["■ 限界"],
        ["", "**1件当たり日数は令和6・7年度の2年分しかない。**"
             "国保連月報が令和6年4月審査分からしかないためである。"],
        ["", "件数・給付費は年報で令和3〜7年度の5年分あるが、日数はこの2年で置くほかない。"],
        ["", "**福祉用具販売と住宅改修は月報のサービス種類にない。**"
             "償還払いのため給付統計に計上されない。人数は年報の件数による。"],
    ]
    for r in rows:
        ws.append(r)
    ws["A1"].font = Font(bold=True, size=13)
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith("■"):
            ws.cell(r, 1).font = Font(bold=True, size=10)
        for c in range(3, 7):
            cell = ws.cell(r, c)
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0"
                cell.fill = CALC
    widths(ws, [3, 34, 20, 30, 14, 16])
    for row in ws.iter_rows():
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    return ws


def sheet_jisseki(wb, name, a6, a7, kind, label, n6, n7):
    ws = wb.create_sheet(name)
    ws.append(["区分", "サービス", "計画表記の単位",
               "R6 件数", "R6 日数", "R6 日数/件",
               "R7 件数", "R7 日数", "R7 日数/件", "R7 単位数/日",
               "日数/件の R6→R7", "備考"])
    order = [(c, s) for c, s, _u in T.ORDER_KAIGO]
    seen = set()
    for cat, svc in order:
        if svc in seen:
            continue
        seen.add(svc)
        d6, d7 = a6.get((svc, kind), {}), a7.get((svc, kind), {})
        k6, h6 = d6.get("件数", 0), d6.get("日数", 0)
        k7, h7 = d7.get("件数", 0), d7.get("日数", 0)
        if not (k6 or k7) and svc not in UNIT:
            continue
        p6 = (h6 / k6) if k6 else None
        p7 = (h7 / k7) if k7 else None
        chg = f"{(p7 / p6 - 1) * 100:+.1f}％" if (p6 and p7) else "―"
        memo = ""
        if svc in ("特定福祉用具販売", "住宅改修"):
            memo = "**月報のサービス種類にない（償還払い）。人数は年報による**"
        elif UNIT.get(svc) == "人":
            memo = "量は人数のみ。日数は在所日数・貸与日数であり計画書には載せない"
        ws.append([cat, svc, UNIT.get(svc, ""), k6, h6, p6, k7, h7, p7,
                   (d7["単位数"] / h7) if h7 else None, chg, memo])
    style_header(ws)
    for r in range(2, ws.max_row + 1):
        for c in (4, 5, 7, 8, 10):
            ws.cell(r, c).number_format = "#,##0"
            ws.cell(r, c).fill = CALC
        for c in (6, 9):
            ws.cell(r, c).number_format = "#,##0.00"
            ws.cell(r, c).fill = KEY
    body(ws, wrap=(12,))
    widths(ws, [11, 32, 8, 10, 10, 10, 10, 10, 10, 11, 12, 42])
    ws.freeze_panes = "D2"
    notes(ws, [
        f"※ {label}。国保連月報による。"
        f"令和6年度は審査{n6}か月、令和7年度は審査{n7}か月（いずれも提供月ベースで12か月）。",
        "※ 「日数/件」は1件（＝1人が1か月そのサービスを使ったとき）当たりの実日数。"
        "**これに見込件数を乗じて見込量を出す。**",
        "※ 「単位数/日」は1日当たりの報酬単位数。1回分の報酬と比べて1日何回使っているかの手がかり。",
        "※ 短期利用の区分（認知症対応型共同生活介護（短期）など）は本体に合算した。",
    ])
    return ws


def sheet_mikomi(wb, name, a7, proj, kind, label):
    ws = wb.create_sheet(name)
    hdr = ["区分", "サービス", "単位", "R7 日数/件", "R7実績 人/月", "R7実績 量/月",
           "R7 単価(円/件)", "見込単価(円/件)", "単価の乖離"]
    for y in T.PLAN_YEARS:
        hdr += [f"{y[2:]} 人/月", f"{y[2:]} 量/月", f"{y[2:]} 給付費(千円)"]
    hdr += ["R7→R11", "備考"]
    ws.append(hdr)
    by = {r["サービス"]: r for r in proj[kind]}
    i = 1 if kind == "介護" else 0
    seen = set()
    for cat, svc, _u in T.ORDER_KAIGO:
        if svc in seen or svc not in by:
            continue
        seen.add(svc)
        row = by[svc]
        d7 = a7.get((svc, kind), {})
        k7, h7 = d7.get("件数", 0), d7.get("日数", 0)
        per = (h7 / k7) if k7 else 0
        unit = UNIT.get(svc, "人")
        ken7 = row["_R7件数"]
        # 実績の量は月報の実測日数、実績の人数は年報の件数による（出所が異なる）
        g7 = row["_R7給付費"]
        t7 = (g7 / ken7) if ken7 else 0
        ken9, kyu9 = row["見込"][1]
        ti = (kyu9 * 1000 / ken9) if ken9 else 0
        dev = (ti / t7 - 1) * 100 if t7 else None
        r = [cat, disp(svc, kind), unit, per if unit != "人" else None,
             ken7 / 12, (h7 / 12) if unit != "人" else None,
             t7 or None, ti or None,
             f"{dev:+.1f}％" if dev is not None else "―"]
        last = None
        for k in range(1, len(T.EST_YEARS)):
            ken, kyu = row["見込"][k]
            r += [ken / 12, (ken * per / 12) if unit != "人" else None, kyu]
            last = ken * per / 12 if unit != "人" else ken / 12
        base = (h7 / 12) if unit != "人" else ken7 / 12
        chg = (last / base - 1) * 100 if base else None
        memo = ""
        if unit != "人" and not k7:
            memo = "**月報に実績がないため量は出せない。人数のみ**"
        elif dev is not None and abs(dev) > 10:
            memo = ("**見込単価が令和7年度実績から10％を超えて離れる。"
                    "単価も一方向に動いてきたため、3年平均が実績から離れる。"
                    "件数と単価が逆方向に動くと給付費では打ち消し合う**")
        elif chg is not None and abs(chg) > 15:
            memo = ("**令和7年度実績から15％を超えて動く。"
                    "基準期間を令和5〜7年度の3年平均としているため、"
                    "一方向に動いてきたサービスでは平均に引き戻される。"
                    "00シートの「趨勢のあるサービスの扱い」を参照**")
        r += [f"{chg:+.1f}％" if chg is not None else "―", memo]
        ws.append(r)
        if memo.startswith("**令和7"):
            for c in range(1, len(hdr) + 1):
                ws.cell(ws.max_row, c).fill = WARN
    style_header(ws)
    n = len(hdr)
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 4).number_format = "#,##0.00"
        ws.cell(r, 4).fill = KEY
        for c in (5, 6):
            ws.cell(r, c).number_format = "#,##0.0"
            ws.cell(r, c).fill = CALC
        for c in (7, 8):
            ws.cell(r, c).number_format = "#,##0"
            ws.cell(r, c).fill = CALC
        for k in range(len(T.PLAN_YEARS)):
            ws.cell(r, 10 + k * 3).number_format = "#,##0.0"
            ws.cell(r, 11 + k * 3).number_format = "#,##0.0"
            ws.cell(r, 12 + k * 3).number_format = "#,##0"
            for c in (10 + k * 3, 11 + k * 3, 12 + k * 3):
                ws.cell(r, c).fill = CALC
    body(ws, wrap=(n,))
    widths(ws, [11, 34, 6, 11, 11, 11, 12, 12, 11]
           + [10, 11, 13] * len(T.PLAN_YEARS) + [10, 38])
    ws.freeze_panes = "D2"
    notes(ws, [
        f"※ {label}。「量/月」は単位欄の単位（回・日）による月あたりの量。"
        "単位が「人」のサービスは人数のみで、量は出さない。",
        "※ 量/月 ＝ 人/月 × R7の日数/件。人/月は年報から組んだ見込件数÷12。",
        "**※ 訪問系の「回」は国保連月報の実日数による。**"
        "1日当たり単位数からみて1日1回が主体だが、厳密には回数は日数を1割程度上回る。",
        "※ 給付費は年報から組んだ試算による。報酬改定率は0.0％で置いている。",
        "**※ 「見込単価」は見込給付費÷見込件数。件数と給付費をそれぞれ"
        "令和5〜7年度の平均で置いているため、両者の比である単価は令和7年度実績と一致しない。**"
        "10％を超えて離れるのは訪問介護（▲13.8％）と短期入所療養介護（老健）（▲12.9％）の2件。",
    ])
    return ws


def sheet_keikaku(wb, a7, proj):
    """計画書の見込量表の形にそろえたもの。"""
    ws = wb.create_sheet("05_計画書掲載用")
    ws.append(["区分", "サービス", "単位", "令和7年度（実績）"]
              + [f"令和{y[2:]}" for y in T.PLAN_YEARS] + ["備考"])
    ncol = 4 + len(T.PLAN_YEARS)
    for kind, head in (("介護", "■ 介護給付（要介護1〜5）"),
                       ("予防", "■ 予防給付（要支援1・2）")):
        ws.append([head])
        ws.cell(ws.max_row, 1).font = Font(bold=True, size=10)
        by = {r["サービス"]: r for r in proj[kind]}
        tot = [0.0] * (1 + len(T.PLAN_YEARS))
        seen = set()
        for cat, svc, _u in T.ORDER_KAIGO:
            if svc in seen or svc not in by:
                continue
            seen.add(svc)
            row = by[svc]
            d7 = a7.get((svc, kind), {})
            k7, h7 = d7.get("件数", 0), d7.get("日数", 0)
            per = (h7 / k7) if k7 else 0
            unit = UNIT.get(svc, "人")
            ken7 = row["_R7件数"]
            nin = [ken7 / 12] + [row["見込"][k][0] / 12
                                 for k in range(1, len(T.EST_YEARS))]
            for i, v in enumerate(nin):
                tot[i] += v
            name = disp(svc, kind)
            if unit == "人":
                ws.append([cat, name, "人/月"] + [round(v, 1) for v in nin] + [""])
            elif per:
                # 量は月報の実測日数、見込は見込件数×1件当たり日数
                ryo = [h7 / 12] + [row["見込"][k][0] * per / 12
                                   for k in range(1, len(T.EST_YEARS))]
                ws.append([cat, name, f"{unit}/月"] + [round(v) for v in ryo] + [""])
                ws.append(["", "", "人/月"] + [round(v, 1) for v in nin] + [""])
            else:
                # 月報に実績がないサービスも、単位は計画書の表記のままにする
                ws.append([cat, name, f"{unit}/月"] + [0] * (1 + len(T.PLAN_YEARS))
                          + ["**月報に実績がない。量は0**"])
                ws.append(["", "", "人/月"] + [round(v, 1) for v in nin] + [""])
        ws.append([f"{kind}給付　利用者数の合計", "", "人/月"]
                  + [round(v, 1) for v in tot] + ["延べ利用者数。重複を含む"])
        for c in range(1, ncol + 2):
            ws.cell(ws.max_row, c).font = Font(bold=True, size=9)
            ws.cell(ws.max_row, c).fill = KEY
    style_header(ws)
    for r in range(2, ws.max_row + 1):
        for c in range(4, ncol + 1):
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)):
                ws.cell(r, c).number_format = ("#,##0.0" if isinstance(v, float)
                                               else "#,##0")
                if not ws.cell(r, c).fill.fgColor.rgb == KEY.fgColor.rgb:
                    ws.cell(r, c).fill = CALC
    body(ws, wrap=(ncol + 1,))
    widths(ws, [11, 34, 9, 14, 12, 12, 12, 34])
    ws.freeze_panes = "D2"
    notes(ws, [
        "※ 計画書の見込量表にそのまま載せられる形にしたもの。"
        "回・日で表記するサービスは、量と人数の2行で示している。",
        "**※ 訪問系の「回/月」は国保連月報の実日数による。**"
        "計画書に載せる前に、回数として表記してよいかを町と確認する。",
        "**※ 令和7年度（実績）は、量が国保連月報の実測日数÷12、"
        "人数が年報の件数÷12である。**出所が異なるため、"
        "量÷人数は1件当たり日数と厳密には一致しない（件数が年報628件 対 月報642件で2.2％違う）。",
        "※ サービス名は計画書の正式名称による（予防給付は「介護予防」を冠する）。",
    ])
    return ws


def sheet_teiin(wb, a7, proj):
    """見込量と町内の定員を突き合わせる（手引き手順9 供給実現性）。"""
    ws = wb.create_sheet("06_定員との突合")
    ws.append(["区分", "サービス", "町内定員", "単位",
               "R11見込（人/月）", "R11見込（1日あたり）", "充足率", "判定", "見方"])
    by = {(r["サービス"], k): r for k in ("介護", "予防") for r in proj[k]}
    for svc, (cap, unit) in TEIIN.items():
        nin = sum(by[(svc, k)]["見込"][-1][0] / 12
                  for k in ("介護", "予防") if (svc, k) in by)
        if not nin and not cap:
            continue
        if unit == "人":
            per_day, base = None, nin
        else:
            # 通所系は延べ日数を営業日数で割って1日あたりの利用者数にする
            tot = 0.0
            for k in ("介護", "予防"):
                if (svc, k) not in by:
                    continue
                d = a7.get((svc, k), {})
                p_ = (d["日数"] / d["件数"]) if d.get("件数") else 0
                tot += by[(svc, k)]["見込"][-1][0] * p_ / 12
            per_day, base = tot / EIGYO, tot / EIGYO
        rate = (base / cap * 100) if cap else None
        if not cap:
            judge, how = "町内になし", "町外の事業所の利用による。見込量は残す"
        elif rate > 100:
            judge, how = "**超過**", "町外の事業所の利用による。自給率の分析と整合させる"
        elif rate > 85:
            judge, how = "ほぼ満床", "**整備がなければこれ以上増やせない。見込量の上限になる**"
        else:
            judge, how = "余力あり", "定員は上限として働いていない"
        ws.append(["施設・居住系" if unit == "人" else "通所系", svc, cap or None, unit,
                   nin, per_day, rate, judge, how])
    style_header(ws)
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 3).number_format = "#,##0"
        for c in (5, 6):
            ws.cell(r, c).number_format = "#,##0.0"
            ws.cell(r, c).fill = CALC
        ws.cell(r, 7).number_format = "0.0％"
        j = ws.cell(r, 8).value
        f = WARN if j in ("**超過**", "ほぼ満床") else CALC
        ws.cell(r, 8).fill = f
        ws.cell(r, 7).fill = f
    body(ws, wrap=(9,))
    widths(ws, [14, 34, 10, 14, 15, 17, 10, 12, 44])
    notes(ws, [
        "※ 定員は見える化システムのＤ25〜Ｄ27（令和7年度）による。出典は介護サービス情報公表システム。",
        f"※ 通所系は、見込延べ日数を月{EIGYO}日の営業日で割って1日あたりの利用者数にしている。",
        "**※ 介護老人福祉施設は町内定員54人に対し見込み67.6人/月で、13.6人が町外の施設の利用にあたる。**"
        "第9期の自給率の分析と整合させる。",
        "**※ 小規模多機能型居宅介護は登録定員が29人以下と定められている。**"
        "見込み31.3人/月は登録者数にあたるため、1事業所であれば上限を超える。"
        "事業所数と登録定員を町に確認する。",
        "※ 認知症対応型共同生活介護の定員は令和5年度71人から令和6年度98人へ増えている。"
        "第9期が「新規整備を原則行わない」方針であったこととの整合を、第9期評価で整理する。",
    ])
    return ws


def sheet_check(wb, chk, a7, proj):
    ws = wb.create_sheet("06_突合結果")
    ws.append(["№", "突合の内容", "値A", "値B", "差", "差の率", "判定", "扱い"])
    items = []
    for fy, lab in ((2024, "令和6年度"), (2025, "令和7年度")):
        a, b = chk[fy]
        items.append((f"{lab}　月報の3区分計 対 年報 様式2 総計（給付額・円）", a, b,
                      "要確認",
                      "**両年度とも▲1.0％で幅がそろう。系統的な差だが理由が特定できていない。"
                      "月報は1件当たり日数の算出にのみ用い、量と給付費は年報に合わせる**"))
    # 件数の突合（年報 対 月報・令和7年度）
    for kind in ("介護", "予防"):
        nk = sum(r["_R7件数"] for r in proj[kind])
        gk = sum(v["件数"] for (s, k), v in a7.items() if k == kind)
        items.append((f"令和7年度 件数の合計　年報 対 月報（{kind}給付）", nk, gk,
                      "要確認" if abs(nk / gk - 1) > 0.02 else "軽微",
                      "年報は決定ベース、月報は審査ベース。"
                      "**見込量の人数は年報を用いる**"))
    for i, (name, a, b, judge, how) in enumerate(items, 1):
        ws.append([i, name, a, b, a - b, f"{(a / b - 1) * 100:+.2f}％", judge, how])
    style_header(ws)
    for r in range(2, ws.max_row + 1):
        for c in (3, 4, 5):
            ws.cell(r, c).number_format = "#,##0"
        ws.cell(r, 7).fill = {"一致": OK, "軽微": CALC, "要確認": WARN}[ws.cell(r, 7).value]
    body(ws, wrap=(2, 8))
    widths(ws, [4, 44, 16, 16, 14, 10, 9, 56])
    notes(ws, [
        "※ 月報と年報は出所が異なる（月報は国保連の審査ベース、年報は保険者の決定ベース）。"
        "一致しないこと自体は異常ではない。",
        "**※ 本試算は、量の比（1件当たり日数）を月報から、量と給付費の水準を年報から取る。**"
        "こうすれば出所の差は見込量に持ち込まれない。",
        "※ 差の理由（過誤調整の反映時期、住所地特例の扱い等）は国保連に照会する。",
    ])
    return ws


def plan_rows():
    """計画素案の見込量表に載せる行を返す。

    [(区分, サービス名, 単位, [R7実績, R9, R10, R11]), …] を介護・予防の別に。
    """
    data = read_geppo()
    _n7, a7, _raw = aggregate(data, 2025)
    act, _sub, gen = T.extract()
    proj = T.project(act, gen, "1号", "1号", 0.0)
    for kind in ("介護", "予防"):
        i = 1 if kind == "介護" else 0
        for row in proj[kind]:
            row["_R7件数"] = act["令和7年度"]["件数"].get(row["サービス"], (0, 0))[i]
    out = {}
    for kind in ("介護", "予防"):
        by = {r["サービス"]: r for r in proj[kind]}
        rows, seen = [], set()
        for cat, svc, _u in T.ORDER_KAIGO:
            if svc in seen or svc not in by:
                continue
            seen.add(svc)
            r = by[svc]
            d7 = a7.get((svc, kind), {})
            k7, h7 = d7.get("件数", 0), d7.get("日数", 0)
            per = (h7 / k7) if k7 else 0
            unit = UNIT.get(svc, "人")
            ken7 = r["_R7件数"]
            nin = [ken7 / 12] + [r["見込"][k][0] / 12
                                 for k in range(1, len(T.EST_YEARS))]
            ryo = ([h7 / 12] + [r["見込"][k][0] * per / 12
                                for k in range(1, len(T.EST_YEARS))]) if per else None
            rows.append((cat, disp(svc, kind), unit, nin, ryo,
                         [r["見込"][k][1] for k in range(1, len(T.EST_YEARS))]))
        out[kind] = rows
    # 給付費の集計（千円）。計画素案の表に使う。
    agg = {"区分": {}, "給付": {}, "計": [0.0] * len(T.PLAN_YEARS)}
    for kind in ("介護", "予防"):
        v = [0.0] * len(T.PLAN_YEARS)
        for r in proj[kind]:
            for k in range(len(T.PLAN_YEARS)):
                v[k] += r["見込"][k + 1][1]
                agg["区分"].setdefault(r["区分"], [0.0] * len(T.PLAN_YEARS))[k] \
                    += r["見込"][k + 1][1]
        agg["給付"][kind] = v
        for k in range(len(T.PLAN_YEARS)):
            agg["計"][k] += v[k]
    ratios = T.kasan_ratios(gen)
    ken = [sum(r["見込"][k + 1][0] for kk in ("介護", "予防") for r in proj[kk])
           for k in range(len(T.PLAN_YEARS))]
    agg["特定入所"] = [agg["計"][k] * ratios["特定入所"] for k in range(len(T.PLAN_YEARS))]
    agg["高額"] = [agg["計"][k] * ratios["高額"] for k in range(len(T.PLAN_YEARS))]
    agg["高額合算"] = [agg["計"][k] * ratios["高額合算"] for k in range(len(T.PLAN_YEARS))]
    agg["手数料"] = [ken[k] * TESURYO_YEN / 1000 for k in range(len(T.PLAN_YEARS))]
    agg["標準給付費"] = [agg["計"][k] + agg["特定入所"][k] + agg["高額"][k]
                    + agg["高額合算"][k] + agg["手数料"][k]
                    for k in range(len(T.PLAN_YEARS))]
    out["集計"] = agg
    return out


# ================================================================ main

def main():
    data = read_geppo()
    n6, a6, raw6 = aggregate(data, 2024)
    n7, a7, raw7 = aggregate(data, 2025)

    act, _sub, gen = T.extract()
    proj = T.project(act, gen, "1号", "1号", 0.0)
    for kind in ("介護", "予防"):
        i = 1 if kind == "介護" else 0
        for row in proj[kind]:
            row["_R7件数"] = act["令和7年度"]["件数"].get(row["サービス"], (0, 0))[i]
            row["_R7給付費"] = act["令和7年度"]["給付費"].get(row["サービス"], (0, 0))[i]

    NENPO = {}
    for fy, lab in ((2024, "令和6年度"), (2025, "令和7年度")):
        NENPO[fy] = sum(act[lab]["給付費"].get(s, (0, 0))[0]
                        + act[lab]["給付費"].get(s, (0, 0))[1]
                        for _c, s, _u in T.ORDER_KAIGO)
    chk = {"act": act, "pop7": gen["令和7年度"]["1号"],
           "proj": {(r["サービス"], k): r["見込"][-1][0]
                    for k in ("介護", "予防") for r in proj[k]}}
    for fy, raw in ((2024, raw6), (2025, raw7)):
        three = sum(raw[k]["給付額"] for k in ("居宅サービス計", "地域密着型サービス計",
                                            "施設サービス計"))
        chk[fy] = (three, NENPO[fy])

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    sheet_intro(wb, a6, a7, chk)
    sheet_jisseki(wb, "01_月報の実績_介護", a6, a7, "介護", "介護給付（要介護1〜5）", n6, n7)
    sheet_jisseki(wb, "02_月報の実績_予防", a6, a7, "予防", "予防給付（要支援1・2）", n6, n7)
    sheet_mikomi(wb, "03_見込量_介護", a7, proj, "介護", "介護給付（要介護1〜5）")
    sheet_mikomi(wb, "04_見込量_予防", a7, proj, "予防", "予防給付（要支援1・2）")
    sheet_keikaku(wb, a7, proj)
    sheet_teiin(wb, a7, proj)
    sheet_check(wb, chk, a7, proj)
    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / f"小野町_第10期_回数日数ベース見込量_{ASOF}.xlsx"
    wb.save(p)

    print(f"生成: {p.name}（{len(wb.sheetnames)}シート）")
    print(f"月報 令和6年度 {n6}か月／令和7年度 {n7}か月")
    print("\n令和7年度の1件当たり日数（主なもの）")
    for svc in ("訪問介護", "訪問看護", "通所介護", "通所リハビリテーション",
                "短期入所生活介護", "認知症対応型通所介護"):
        d = a7.get((svc, "介護"), {})
        if d.get("件数"):
            print(f"  {svc:<22} {d['日数'] / d['件数']:>6.1f} 日／件"
                  f"　単位数/日 {d['単位数'] / d['日数']:>6,.0f}")
    print("\n令和11年度の見込量（月あたり・主なもの）")
    by = {r["サービス"]: r for r in proj["介護"]}
    for svc in ("訪問介護", "通所介護", "短期入所生活介護", "介護老人福祉施設"):
        d = a7.get((svc, "介護"), {})
        ken = by[svc]["見込"][-1][0]
        per = (d["日数"] / d["件数"]) if d.get("件数") else 0
        u = UNIT[svc]
        if u == "人":
            print(f"  {svc:<22} {ken / 12:>7.1f} 人/月")
        else:
            print(f"  {svc:<22} {ken * per / 12:>7.0f} {u}/月　"
                  f"（{ken / 12:.1f} 人/月）")


if __name__ == "__main__":
    main()
