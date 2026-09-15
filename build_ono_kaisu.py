"""小野町 第10期 サービス見込量を算定する。

計画書の見込量表に載せる、サービス別・年度別の利用者数／回（日）数／給付費を確定する。
あわせて地域支援事業費、供給実現性（定員との突合）、据え置いた前提の一覧を作る。

**水準は令和7年度の実績、伸びは自前の利用率×将来推計人口**

  見込量(年度y) ＝ 令和7年度の実績 × 見込件数(y) ÷ 令和7年度の件数

  見込件数は年報から組んだ試算（build_ono_tanka）による。第1号被保険者1人当たりの
  利用率を基準期間の平均で置き、第10期将来推計用推計人口を乗じたもの。
  伸び率は無次元なので、水準の出所が指標ごとに違っても混ぜられる。

**指標ごとの出所**

  人数（1月当たり利用者数）  見える化システムのワークシート（令和7年度）。
                          年報 様式1の6 の受給者数にあたり、計画書の標準の表記である。
  量（1月当たり回・日数）    同じくワークシート。「回数」は給付実績情報の回数であり、
                          国保連月報の「日数」（各サービス種類ごとの実日数）とは
                          別物である。訪問介護で回数が日数を26％、訪問看護で32％
                          上回る。計画書は回数で載せるため、こちらを採る。
  給付費（年間）            年報 様式2。確報であり費用額＝単位数×10円で内的に整合する。

  **施設サービスの額は月報・見える化から採らない。**介護老人福祉施設の令和7年度は
  年報204,850千円に対し月報225,660千円で、月報の費用額は1単位当たり11.95円と
  10円を大きく上回る。差は食費・居住費（特定入所者介護サービス費の対象）であり、
  総給付費に入れると補足給付を二重に見込むことになる。

  量を載せないサービス（人数だけのもの）は、ワークシートに回（日）数の欄がない
  月額・日額の包括報酬のサービスである。

**見える化ワークシートの年度ラベル**

  **ワークシートの「令和6年度」列には令和5年度の実績が入っている。**
  全サービスで年報の令和5年度と0.01％以内で一致する（総給付費1,026,350千円 対
  年報令和5年度1,026,324千円）。令和7年度列は年報の令和7年度とおおむね一致する。
  令和8年度列は審査4か月分の月平均を年換算したもので、サービス別には採れない。
  このため実績として採るのは令和7年度列だけにしている。

**基準期間**

  既定は令和5〜7年度の3年平均。次の2つがそろったサービスだけ令和7年度単年に切り替える。
    ① 令和7年度の件数が3年平均から5％以上離れている
    ② 令和8年度の国保連月報の前年同期比がその向きを裏づける
       （審査202605〜202607と前年の同じ月。増なら1.05以上、減なら0.95以下）
  判定は check_base() が月報から再計算し、07_基準期間の判定 に載せる。

出力: 04_算定・見込量/小野町_第10期_サービス見込量_YYYYMMDD.xlsx
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
import ono_mieruka_ws as W

ROOT = pathlib.Path(__file__).parent / "小野町_引継ぎ_整理済"
ZDIR = ROOT / "18_年報・国保連データ" / "原本_国保連月報"
OUT = ROOT / "04_算定・見込量"
ASOF = "20260915"
ASOF_JP = "令和8年9月15日"

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


def sheet_keikaku(wb, plan):
    """計画書の見込量表の形にそろえたもの。"""
    ws = wb.create_sheet("05_計画書掲載用")
    ws.append(["区分", "サービス", "単位", "令和7年度（実績）"]
              + [f"令和{y[2:]}" for y in T.PLAN_YEARS] + ["基準期間", "備考"])
    ncol = 4 + len(T.PLAN_YEARS)
    for kind, head in (("介護", "■ 介護給付（要介護1〜5）"),
                       ("予防", "■ 予防給付（要支援1・2）")):
        ws.append([head])
        ws.cell(ws.max_row, 1).font = Font(bold=True, size=10)
        tot = [0.0] * (1 + len(T.PLAN_YEARS))
        kyu_tot = [0.0] * len(T.PLAN_YEARS)
        for cat, name, unit, nin, ryo, kyu in plan[kind]:
            svc = _svc_of(name, kind)
            base = "／".join(y[2:] for y in T.base_years(svc, kind))
            for i2, v in enumerate(nin):
                tot[i2] += v
            for i2, v in enumerate(kyu):
                kyu_tot[i2] += v
            if ryo:
                ws.append([cat, name, f"{unit}/月"] + [round(v) for v in ryo]
                          + [base, ""])
                ws.append(["", "", "人/月"] + [round(v, 1) for v in nin] + ["", ""])
            else:
                ws.append([cat, name, "人/月"] + [round(v, 1) for v in nin]
                          + [base, ""])
            ws.append(["", "", "給付費(千円)", None] + [round(v) for v in kyu]
                      + ["", ""])
        ws.append([f"{kind}給付　利用者数の合計", "", "人/月"]
                  + [round(v, 1) for v in tot] + ["", "延べ利用者数。重複を含む"])
        for c in range(1, ncol + 3):
            ws.cell(ws.max_row, c).font = Font(bold=True, size=9)
            ws.cell(ws.max_row, c).fill = KEY
        ws.append([f"{kind}給付　給付費の合計", "", "千円", None]
                  + [round(v) for v in kyu_tot] + ["", ""])
        for c in range(1, ncol + 3):
            ws.cell(ws.max_row, c).font = Font(bold=True, size=9)
            ws.cell(ws.max_row, c).fill = KEY
    style_header(ws)
    for r in range(2, ws.max_row + 1):
        for c in range(4, ncol + 1):
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)):
                ws.cell(r, c).number_format = ("#,##0.0" if isinstance(v, float)
                                               else "#,##0")
                if ws.cell(r, c).fill.fgColor.rgb != KEY.fgColor.rgb:
                    ws.cell(r, c).fill = CALC
    body(ws, wrap=(ncol + 2,))
    widths(ws, [11, 34, 12, 14, 12, 12, 12, 14, 34])
    ws.freeze_panes = "D2"
    notes(ws, [
        "※ 計画書の見込量表にそのまま載せられる形にしたもの。"
        "回・日で表記するサービスは、量・人数・給付費の3行で示している。",
        "**※ 量と人数の令和7年度実績は見える化システムのワークシートによる。**"
        "ワークシートの「回数」は給付実績情報の回数であり、国保連月報の実日数とは別物である。"
        "訪問介護では回数が日数を26％、訪問看護では32％上回る。計画書は回数で載せるため、"
        "こちらを採った。",
        "**※ 給付費の令和7年度実績は年報 様式2 による。**確報であり費用額＝単位数×10円で"
        "内的に整合する。施設サービスは月報・見える化の額に特定入所者介護サービス費が"
        "乗っているため採らない（介護老人福祉施設で年報204,850千円 対 月報225,660千円）。",
        "※ 第10期の見込みは、令和7年度の実績に「自前の見込件数÷令和7年度の件数」を乗じたもの。"
        "見込件数は第1号被保険者1人当たりの利用率（基準期間の平均）×将来推計人口による。",
        "※ 量を載せないサービス（人数だけのもの）は、月額または日額の包括報酬であり、"
        "見える化のワークシートにも回（日）数の欄がない。",
        "※ サービス名は計画書の正式名称による（予防給付は「介護予防」を冠する）。",
    ])
    return ws


_SVC_CACHE = {}


def _svc_of(name, kind):
    """計画書の表示名から正規化名に戻す。"""
    if not _SVC_CACHE:
        for _c, svc, _u in T.ORDER_KAIGO:
            for k in ("介護", "予防"):
                _SVC_CACHE[(disp(svc, k), k)] = svc
    return _SVC_CACHE.get((name, kind), name)


def sheet_teiin(wb, plan):
    """見込量と町内の定員を突き合わせる（手引き手順9 供給実現性）。"""
    ws = wb.create_sheet("06_定員との突合")
    ws.append(["区分", "サービス", "町内定員", "単位",
               "R11見込（人/月）", "R11見込（1日あたり）", "充足率", "判定", "見方"])
    nin11, ryo11 = {}, {}
    for kind in ("介護", "予防"):
        for _cat, name, _u, nin, ryo, _k in plan[kind]:
            svc = _svc_of(name, kind)
            nin11[svc] = nin11.get(svc, 0.0) + nin[-1]
            if ryo:
                ryo11[svc] = ryo11.get(svc, 0.0) + ryo[-1]
    for svc, (cap, unit) in TEIIN.items():
        nin = nin11.get(svc, 0.0)
        if not nin and not cap:
            continue
        if unit == "人":
            per_day, base = None, nin
        elif svc not in ryo11:
            # 小規模多機能型居宅介護のように包括報酬で回（日）数がないものは
            # 登録者数を定員と並べる。1日あたりには直せない。
            per_day, base = None, nin
        else:
            per_day = base = ryo11.get(svc, 0.0) / EIGYO
        rate = (base / cap * 100) if cap else None
        if svc == "小規模多機能型居宅介護":
            judge = "**要確認**"
            how = ("**登録者数と通いの定員を並べたもので、充足率としては読めない。"
                   "登録定員は1事業所29人以下と定められている。"
                   "事業所数と登録定員を町に確認する**")
            ws.append(["通所系", svc, cap or None, unit, nin, per_day, None,
                       judge, how])
            continue
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
    tokuyo = nin11.get("介護老人福祉施設", 0)
    shota = nin11.get("小規模多機能型居宅介護", 0)
    notes(ws, [
        "※ 定員は見える化システムのＤ25〜Ｄ27（令和7年度）による。"
        "出典は介護サービス情報公表システム。",
        f"※ 通所系は、見込延べ回（日）数を月{EIGYO}日の営業日で割って"
        "1日あたりの利用者数にしている。",
        f"**※ 介護老人福祉施設は町内定員54人に対し見込み{tokuyo:.1f}人/月で、"
        f"{tokuyo - 54:.1f}人が町外の施設の利用にあたる。**"
        "第9期の自給率の分析と整合させる。",
        f"**※ 小規模多機能型居宅介護は登録定員が29人以下と定められている。**"
        f"見込み{shota:.1f}人/月は登録者数にあたるため、1事業所であれば上限を超える。"
        "事業所数と登録定員を町に確認する。",
        "※ 認知症対応型共同生活介護の定員は令和5年度71人から令和6年度98人へ増えている。"
        "第9期が「新規整備を原則行わない」方針であったこととの整合を、第9期評価で整理する。",
        "※ 通所リハビリテーションの人数は介護給付と予防給付の合計だが、"
        "1日あたりは介護給付の回数だけによる（介護予防通所リハビリテーションは"
        "月額の包括報酬で回数がない）。町内に事業所がないため定員の判定には影響しない。",
    ])
    return ws


def sheet_kijun(wb, chk_rows):
    """基準期間の置き方を、月報の前年同期比で点検する。"""
    ws = wb.create_sheet("07_基準期間の判定")
    ws.append(["サービス", "給付の別", "令和5年度", "令和6年度", "令和7年度",
               "3年平均", "令和7年度の乖離", "令和8年度 前年同期比",
               "判定", "採用している基準期間", "一致"])
    for name, kind, k5, k6, k7, m3, kai, rat, hantei, genko, ok in chk_rows:
        ws.append([name, kind, k5, k6, k7, m3, kai, rat, hantei, genko,
                   "○" if ok else "**×**"])
    style_header(ws)
    for r in range(2, ws.max_row + 1):
        for c in (3, 4, 5, 6):
            ws.cell(r, c).number_format = "#,##0"
        ws.cell(r, 7).number_format = "+0.0%;-0.0%"
        ws.cell(r, 8).number_format = "0.000"
        h = ws.cell(r, 9).value
        ws.cell(r, 9).fill = WARN if h == "令和7年度単年" else CALC
        if ws.cell(r, 11).value != "○":
            ws.cell(r, 11).fill = PatternFill("solid", fgColor="FFC7CE")
    body(ws)
    widths(ws, [32, 9, 11, 11, 11, 11, 14, 18, 15, 20, 7])
    ws.freeze_panes = "C2"
    notes(ws, [
        "※ 既定の基準期間は令和5〜7年度の3年平均。次の2つがそろったサービスだけ"
        "令和7年度単年に切り替える。"
        "①令和7年度の件数が3年平均から5％以上離れている。"
        "②令和8年度の国保連月報の前年同期比（審査202605〜202607と前年の同じ月）が"
        "その向きを裏づける（増なら1.05以上、減なら0.95以下）。",
        "※ 「一致」が×の行は、この表の判定と実際に採っている基準期間がずれている。"
        "月報を追加で受領したときに変わりうるので、そのつど確かめる。",
        "※ 実績が消滅したもの・制度が廃止されたもの（定期巡回、地域密着型通所介護、"
        "介護療養型医療施設、訪問リハビリテーション、介護医療院）と、"
        "令和6年度から様式2に計上が始まったもの（特定福祉用具販売、住宅改修）は"
        "この判定の対象外とし、別の理由で基準期間を決めている。",
        "**※ 訪問介護は令和5年度958件→令和6年度843件→令和7年度628件と減っているが、"
        "令和8年度の前年同期比は1.132と反転しているため3年平均のまま置いた。**"
        "3年平均だと令和9年度の見込件数は令和7年度実績の1.27倍になる。"
        "基準期間を令和7年度単年にすると給付費は年7,600千円ほど下がる。"
        "協議会にお諮りする事項とする。",
    ])
    return ws


def sheet_sogo(wb):
    """総合事業の事業別実績と、地域支援事業費の第10期見込み。"""
    ws = wb.create_sheet("08_地域支援事業費")
    sg = sogo_jisseki()
    ws.append(["■ 介護予防・日常生活支援総合事業の事業別実績（国保連月報）"])
    ws.cell(ws.max_row, 1).font = Font(bold=True, size=10)
    ws.append(["事業", "令和6年度 件数", "令和6年度 給付額(千円)",
               "令和7年度 件数", "令和7年度 給付額(千円)",
               "令和8年度 件数（3か月）", "令和8年度 給付額(千円)", "1件当たり(円)"])
    hr = ws.max_row
    for name in SOGO_ORDER:
        vals = [sg[y]["事業"][name] for y in ("令和6年度", "令和7年度", "令和8年度")]
        if not any(v[0] or v[1] for v in vals):
            continue
        tan = (vals[1][1] / vals[1][0]) if vals[1][0] else None
        ws.append([name] + [x for v in vals for x in (v[0], round(v[1] / 1000))]
                  + [round(tan) if tan else None])
    ws.append(["月報 計", None, round(sg["令和6年度"]["月報計"] / 1000), None,
               round(sg["令和7年度"]["月報計"] / 1000), None,
               round(sg["令和8年度"]["月報計"] / 1000), None])
    r_geppo = ws.max_row
    ws.append(["年報 様式4 の決算（介護予防・生活支援サービス事業費）", None,
               round(T.CHIIKI_JISSEKI["令和6年度"]["介護予防・生活支援サービス事業"] / 1000),
               None, None, None, None, "令和7年度は年報が未入力"])
    ws.append(["差（決算－月報）", None, round(sg["令和6年度"]["差"] / 1000), None, None,
               None, None, "審査支払手数料・高額介護予防サービス相当事業費等とみる"])
    style_header(ws, hr)
    for r in range(hr + 1, ws.max_row + 1):
        for c in range(2, 9):
            ws.cell(r, c).number_format = "#,##0"
            ws.cell(r, c).fill = CALC
    for c in range(1, 9):
        ws.cell(r_geppo, c).font = Font(bold=True, size=9)

    ws.append([])
    ws.append(["■ 第10期の地域支援事業費（令和6年度決算の据え置き）"])
    ws.cell(ws.max_row, 1).font = Font(bold=True, size=10)
    ws.append(["区分", "事業", "令和6年度実績(千円)", "令和9年度", "令和10年度",
               "令和11年度", "3年計", "置き方・出所"])
    hr2 = ws.max_row
    tot = 0
    for ku, ev, jis, mik, note in chiiki_plan():
        m = round(mik / 1000)
        tot += m * 3
        ws.append([ku, ev, round(jis / 1000), m, m, m, m * 3, note])
    ws.append(["合計", "", None, None, None, None, tot, ""])
    r_tot = ws.max_row
    ws[f"C{r_tot}"] = f"=SUM(C{hr2 + 1}:C{r_tot - 1})"
    for cc in ("D", "E", "F"):
        ws[f"{cc}{r_tot}"] = f"=SUM({cc}{hr2 + 1}:{cc}{r_tot - 1})"
    style_header(ws, hr2)
    for r in range(hr2 + 1, ws.max_row + 1):
        for c in range(3, 8):
            ws.cell(r, c).number_format = "#,##0"
            ws.cell(r, c).fill = KEY if r == r_tot else CALC
    for c in range(1, 9):
        ws.cell(r_tot, c).font = Font(bold=True, size=9)
    body(ws, first=2, wrap=(8,))
    widths(ws, [22, 34, 18, 16, 16, 16, 14, 56])
    notes(ws, [
        "※ 単位は表頭のとおり。給付額は国保連月報（THRK0401）による。",
        "※ 小野町が使っている総合事業のコードは A2（訪問型サービス・独自）、"
        "A6（通所型サービス・独自）、AF（介護予防ケアマネジメント）の3つだけである。"
        "みなし指定（A1・A5）と定率・定額の独自サービス（A3・A4・A7・A8）、"
        "その他の生活支援サービス（A9〜AE）はいずれも実績がない。",
        "**※ A2・A6 を「訪問介護相当サービス」「通所介護相当サービス」として扱った。**"
        "1件当たりは訪問型16,806円・通所型34,414円（令和7年度）で、"
        "介護予防訪問介護相当・介護予防通所介護相当の水準にある。"
        "緩和した基準によるサービスA として整理している場合は、町にご指摘いただきたい。",
        "**※ 第10期は令和6年度決算の据え置き。**令和7年度は年報 様式4 が未入力で"
        "決算額が取れない。月報でみると総合事業は令和7年度が令和6年度の98.5％、"
        "令和8年度（3か月）は年換算でさらに低い。据え置きは実勢をやや上回る安全側になる。",
        "**※ 一般介護予防事業費は令和5年度994千円から令和6年度5,305千円へ5.3倍。**"
        "何の事業かを町に確認したうえで、第10期の通いの場の展開と結びつける。",
        "※ 一般介護予防事業費と包括的支援事業・任意事業費は事業別の内訳が"
        "年報 様式4 にも国保連月報にもない。事業別に置くには町の事業実績が要る。",
    ])
    return ws


def sheet_zentei(wb, plan):
    """据え置いた前提と、まだ確定していないことを一覧にする。"""
    ws = wb.create_sheet("09_据え置いた前提")
    std3 = sum(plan["集計"]["標準給付費"])
    chi3 = sum(v[3] for v in chiiki_plan()) * 3 / 1000
    sogo3 = sum(v[3] for v in chiiki_plan() if v[0] == "総合事業") * 3 / 1000
    pop3 = sum(T.POP_DAI10[y]["1号"] for y in T.PLAN_YEARS)
    base = T.premium(std3, chi3, sogo3, pop3)

    def gap(**kw):
        return T.premium(std3, chi3, sogo3, pop3, **kw)["月額"] - base["月額"]

    ws.append(["#", "項目", "据え置いた値", "据え置いた理由", "確定したら変わる幅",
               "確定の見通し"])
    rows = [
        ("令和9年度の介護報酬改定率", "0.0％（据置）",
         "改定率は令和8年12月頃に示される。それまで置きようがない",
         f"±1.5％で月額±{abs(T.premium(std3 * 1.015, chi3, sogo3, pop3)['月額'] - base['月額']):.0f}円",
         "令和8年12月頃（国）"),
        ("第1号被保険者負担割合", "23％（第9期と同じ）",
         "3年の計画期間ごとに政令で定める。第1期17％から3年ごとに1ポイントずつ"
         "上がってきており、第10期は24％となる可能性がある",
         f"24％なら月額{gap(futan=0.24):+.0f}円", "政令の公布（国）"),
        ("調整交付金の見込交付割合", "5.963％",
         "第9期計画の公表値からの逆算。第10期の全国値は国の基本指針で示される",
         f"1ポイント下がると月額{gap(wari=T.CHOSEI_MIKOMI - 0.01):+.0f}円",
         "国の基本指針"),
        ("地域支援事業費", f"令和6年度決算{chi3 / 3:,.0f}千円を3年度とも据え置き",
         "令和7年度は年報 様式4 が未入力で決算額が取れない。"
         "月報でみると総合事業は令和7年度が令和6年度の98.5％で、据え置きは安全側",
         f"1％で月額{T.premium(std3, chi3 * 1.01, sogo3 * 1.01, pop3)['月額'] - base['月額']:+.1f}円",
         "令和7年度の決算書（町）"),
        ("一般介護予防事業費", "令和6年度決算5,305千円を据え置き",
         "令和5年度の994千円から5.3倍に増えた理由が分からないため、"
         "第10期の展開を織り込めない。**通いの場の拡充は第10期の重点であり、"
         "実際にはこの水準を上回る可能性がある**",
         "1,000千円増で月額＋7.9円",
         "事業実績の受領（町）"),
        ("包括的支援事業・任意事業費", "令和6年度決算29,335千円を据え置き",
         "事業別の内訳が年報 様式4 にも国保連月報にもない",
         "1,000千円増で月額＋1.9円", "事業実績の受領（町）"),
        ("介護給付費準備基金の取崩額", "0千円（取り崩さない）",
         "**年報 様式4 の保有額は令和5・6年度末とも12,000千円だが、"
         "令和4年度末は120,000千円で、繰入金は全年度ゼロなのに残高が10分の1に"
         "なっており系列に矛盾がある。**正しい残高が分からないため取崩を見込まない",
         f"12,000千円で月額{gap(kikin=12000):+.0f}円、"
         f"35,000千円で月額{gap(kikin=35000):+.0f}円",
         "基金運用状況調書（町）"),
        ("審査支払手数料の単価", "60円／件",
         "年報に計上欄がない。見える化の参考係数では国庫負担金の算定対象となる"
         "単価の上限が95円とされている",
         "95円なら月額＋2.8円", "国保連への支払実績（町）"),
        ("施設の整備", "第10期に新規整備なし",
         "町の整備方針が決まっていない。在宅の要介護認定者の24.0％が"
         "入所を検討または申込済みである一方、町内の定員には余力がある",
         "整備があれば段差で増える", "町の方針決定"),
        ("認定者数の計上方法", "令和7年度実績（第1号783人）をそのまま使う",
         "令和7年9月審査分で961人から862人へ99人減っているが、"
         "同じ月の給付件数は1,303件から1,365件へ増えている。"
         "計上方法の変更の有無が確認できていない",
         "分母を認定者数にすると3年計で10.2％低く出る",
         "認定システムの計上方法（町）"),
        ("訪問介護の基準期間", "令和5〜7年度の3年平均",
         "令和5年度958件→令和6年度843件→令和7年度628件と減っているが、"
         "令和8年度の前年同期比は1.132と反転している。"
         "**3年平均だと令和9年度の見込件数は令和7年度実績の1.27倍になる**",
         "令和7年度単年にすると給付費は年7,600千円ほど下がる（月額▲47円）",
         "協議会での判断"),
    ]
    for i, (a, b, c, d, e) in enumerate(rows, start=1):
        ws.append([i, a, b, c, d, e])
    style_header(ws)
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 3).fill = WARN
    body(ws, wrap=(2, 3, 4, 5, 6))
    widths(ws, [4, 26, 30, 54, 30, 22])
    ws.freeze_panes = "B2"
    notes(ws, [
        f"※ 幅は、標準給付費3年計{std3:,.0f}千円・地域支援事業費{chi3:,.0f}千円・"
        f"補正後被保険者数{base['補正後被保険者数']:,.0f}人・予定収納率99.35％のもとでの"
        f"保険料基準額（月額{base['月額']:,.0f}円）からの変化。",
        "**※ この9件が確定しても給付費の水準は動かない。**動くのは保険料であり、"
        "効きが大きいのは第1号被保険者負担割合と調整交付金の見込交付割合という"
        "制度側の2つである。給付費1％は月額62円にとどまる。",
        "※ 町からの受領を待っているのは、令和7年度の決算書、地域支援事業費の事業別内訳、"
        "介護給付費準備基金の残高、認定者数の計上方法の4つ。"
        "いずれも届かない場合は本シートの据え置きのまま計画に載せられる。",
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


# 総合事業の月報コード → 計画書の事業区分。
#   小野町が使っているのは A2・A6・AF の3つだけである。
#   A2/A6 は「独自」（みなし指定の期間が切れたあとの市町村指定）で、
#   1件当たりの額は訪問型16,806円・通所型34,414円（令和7年度）と、
#   介護予防訪問介護相当・介護予防通所介護相当の水準にある。
SOGO_MAP = {
    "A1訪問型サービス（みなし）": "訪問介護相当サービス",
    "A2訪問型サービス（独自）": "訪問介護相当サービス",
    "A3訪問型サービス（独自／定率）": "訪問型サービスA",
    "A4訪問型サービス（独自／定額）": "訪問型サービスA",
    "A5通所型サービス（みなし）": "通所介護相当サービス",
    "A6通所型サービス（独自）": "通所介護相当サービス",
    "A7通所型サービス（独自／定率）": "通所型サービスA",
    "A8通所型サービス（独自／定額）": "通所型サービスA",
    "A9その他の生活支援サービス（配食／定率）": "その他の生活支援サービス",
    "AAその他の生活支援サービス（配食／定額）": "その他の生活支援サービス",
    "ABその他の生活支援サービス（見守り／定率）": "その他の生活支援サービス",
    "ACその他の生活支援サービス（見守り／定額）": "その他の生活支援サービス",
    "ADその他の生活支援サービス（その他／定率）": "その他の生活支援サービス",
    "AEその他の生活支援サービス（その他／定額）": "その他の生活支援サービス",
    "AF介護予防ケアマネジメント": "介護予防ケアマネジメント",
}
SOGO_ORDER = ["訪問介護相当サービス", "訪問型サービスA", "通所介護相当サービス",
              "通所型サービスA", "その他の生活支援サービス", "介護予防ケアマネジメント"]


def sogo_jisseki():
    """介護予防・日常生活支援総合事業の事業別実績を返す。

    国保連月報の総合事業コード（A1〜AF）から年度別の件数と給付額を取り出し、
    年報 様式4 の「介護予防・生活支援サービス事業費」の決算額と突き合わせる。
    決算額と月報の差は審査支払手数料・高額介護予防サービス相当事業費などであり、
    「その他（審査支払手数料等）」として残す。

    返り値 {年度: {"事業": {事業名: (件数, 給付額円)}, "月報計": 円,
                  "決算": 円 or None, "差": 円 or None, "月数": n}}
    """
    data = read_geppo()
    ms = sorted({ym for (ym, _i) in data})
    out = {}
    for fy, jp in ((2024, "令和6年度"), (2025, "令和7年度"), (2026, "令和8年度")):
        yms = [ym for ym in ms if fiscal(ym) == fy]
        ev = {n: [0, 0] for n in SOGO_ORDER}
        for ym in yms:
            for code, name in SOGO_MAP.items():
                ev[name][0] += (data.get((ym, "件数"), {}).get(code, {}) or {}).get("計", 0) or 0
                ev[name][1] += (data.get((ym, "給付額"), {}).get(code, {}) or {}).get("計", 0) or 0
        tot = sum(v[1] for v in ev.values())
        kes = T.CHIIKI_JISSEKI.get(jp, {}).get("介護予防・生活支援サービス事業")
        out[jp] = {"事業": {n: tuple(v) for n, v in ev.items()}, "月報計": tot,
                   "決算": kes, "差": (kes - tot) if kes else None, "月数": len(yms)}
    return out


def chiiki_plan():
    """第10期の地域支援事業費の見込み（円／年）を事業別に返す。

    **置き方は令和6年度決算の据え置き。**年報 様式4 の令和6年度は確報であり、
    令和7年度は年報そのものが未入力で決算額が取れない。国保連月報でみると
    総合事業は令和7年度が令和6年度の98.5％、令和8年度（3か月）はさらに低い。
    高齢者人口も第10期の3年間で3,438人→3,418人とほぼ横ばいであり、
    据え置きは実勢をやや上回る安全側の置き方になる。

    総合事業の事業別内訳は年報 様式4 にないため、国保連月報の令和6年度で割り付ける。
    決算額と月報の差は審査支払手数料・高額介護予防サービス相当事業費とみて
    「その他」に置く。一般介護予防事業費と包括的支援事業・任意事業費は
    月報に載らないため、様式4 の決算額を区分の合計のまま置く。

    返り値 [(区分, 事業, 令和6年度実績円, 第10期見込円, 出所・備考)]
    """
    sg = sogo_jisseki()
    r6 = sg["令和6年度"]
    rows = []
    for name in SOGO_ORDER:
        ken, gak = r6["事業"][name]
        if not (ken or gak):
            continue
        rows.append(("総合事業", name, gak, gak,
                     f"国保連月報の令和6年度（{ken:,}件）。令和7年度は"
                     f"{sg['令和7年度']['事業'][name][1] / 1000:,.0f}千円"))
    if r6["差"]:
        rows.append(("総合事業", "その他（審査支払手数料等）", r6["差"], r6["差"],
                     "年報 様式4 の決算額と国保連月報の差"))
    j = T.CHIIKI_JISSEKI["令和6年度"]
    rows.append(("総合事業", "一般介護予防事業", j["一般介護予防事業"],
                 j["一般介護予防事業"],
                 "年報 様式4。令和5年度994千円から令和6年度5,305千円へ5.3倍。"
                 "**何の事業かは町に確認中。事業別の内訳は様式4にない**"))
    rows.append(("包括的支援事業・任意事業", "包括的支援事業・任意事業",
                 j["包括的支援事業・任意事業"], j["包括的支援事業・任意事業"],
                 "年報 様式4。**地域包括支援センター運営・在宅医療介護連携・"
                 "生活支援体制整備・認知症総合支援・地域ケア会議・任意事業の合計で、"
                 "事業別の内訳は様式4にない**"))
    return rows


def check_base(verbose=True):
    """基準期間の切替判定を月報から再計算し、BASE_OVERRIDE とずれていないか確かめる。

    切替の基準は build_ono_tanka.BASE_OVERRIDE の注記のとおり。
      ① 令和7年度の件数が令和5〜7年度の3年平均から5％以上離れている
      ② 令和8年度の前年同期比（審査202605〜202607と前年の同じ月）がその向きを裏づける
         （増なら1.05以上、減なら0.95以下）

    返り値は [(サービス, 給付の別, R5, R6, R7, 3年平均, 乖離, 前年同期比, 判定, 一致)]。
    """
    data = read_geppo()
    ms = sorted({ym for (ym, _i) in data})
    r8 = [ym for ym in ms if fiscal(ym) == 2026]
    r7s = [f"{int(ym[:4]) - 1:04d}{ym[4:]}" for ym in r8]
    inv = {}
    for code, (name, kind) in MAP.items():
        inv.setdefault((name, kind), []).append(code)
    act, _sub, _gen = T.extract()
    ys = ("令和5年度", "令和6年度", "令和7年度")

    def ken_geppo(codes, yms):
        return sum((data.get((ym, "件数"), {}).get(c, {}) or {}).get("計", 0) or 0
                   for ym in yms for c in codes)

    rows, ng = [], 0
    for _cat, svc, _u in T.ORDER_KAIGO:
        for kind, i in (("介護", 1), ("予防", 0)):
            ken = [act[y]["件数"].get(svc, (0, 0))[i] for y in ys]
            if max(ken) < 12:
                continue
            m3 = sum(ken) / 3
            kai = (ken[2] / m3 - 1) if m3 else 0.0
            codes = inv.get((svc, kind), [])
            a8, a7 = ken_geppo(codes, r8), ken_geppo(codes, r7s)
            rat = (a8 / a7) if a7 else None
            hantei = "3年平均"
            if rat is not None and abs(kai) >= 0.05:
                if kai > 0 and rat >= 1.05:
                    hantei = "令和7年度単年"
                elif kai < 0 and rat <= 0.95:
                    hantei = "令和7年度単年"
            ima = T.base_years(svc, kind)
            genko = ("令和7年度単年" if ima == ["令和7年度"] else
                     "令和6〜7年度2年平均" if len(ima) == 2 else "3年平均")
            # 実績が消滅・制度廃止で0据置にしたものは判定の対象外
            kotei = svc in ("定期巡回・随時対応型訪問介護看護", "地域密着型通所介護",
                            "介護療養型医療施設", "訪問リハビリテーション", "介護医療院",
                            "特定福祉用具販売", "住宅改修")
            ok = kotei or (hantei == genko)
            if not ok:
                ng += 1
            rows.append((disp(svc, kind), kind, ken[0], ken[1], ken[2], m3, kai,
                         rat, hantei, genko, ok))
    if verbose and ng:
        print(f"  ※ 基準期間の判定が {ng} 件、BASE_OVERRIDE とずれています")
    return rows


def plan_rows():
    """計画素案の見込量表に載せる行を返す。

    [(区分, サービス名, 単位, [R7実績, R9, R10, R11] 人数,
      [同] 量 or None, [R9, R10, R11] 給付費), …] を介護・予防の別に。

    **令和7年度の実績を水準とし、自前の伸び率を乗じる。**

      人数（1月当たり利用者数）  見える化ワークシートの令和7年度。年報 様式1の6 の
                              受給者数にあたり、計画書の標準の表記である。
      量（1月当たり回・日数）    同じく見える化の令和7年度。「回数」は給付実績情報の
                              回数であり、国保連月報の実日数とは別物である。
                              訪問介護で2割強、訪問看護で3割、回数が日数を上回る。
      給付費（年間・千円）       年報 様式2 の令和7年度。確報であり、費用額＝単位数×10円
                              で内的に整合する。**施設サービスは月報・見える化の額に
                              特定入所者介護サービス費が乗っているため採らない。**
      伸び率                   自前の試算（第1号被保険者1人当たりの利用率×将来推計人口）
                              による見込件数 ÷ 令和7年度の件数。無次元なので、
                              水準の出所が違っても混ぜられる。

    見える化に令和7年度の値がないサービスは、国保連月報の令和7年度（審査202505〜
    202604）の1件当たり日数で補う。

    **見える化ワークシートの「令和6年度」列は令和5年度の実績である**（全サービスで
    年報の令和5年度と0.01％以内で一致）。年度ラベルが1年ずれているため使わない。
    """
    data = read_geppo()
    _n7, a7, _raw = aggregate(data, 2025)
    act, _sub, gen = T.extract()
    proj = T.project(act, gen, "1号", "1号", 0.0)
    mie_pp = W.per_person("令和7年度")
    mie_vu = W.volume_unit()
    mie, _tot, _ku, _hoken = W.read()
    mie_nin = {}
    for kind, mp in (("介護", W.MAP_KAIGO), ("予防", W.MAP_YOBO)):
        for mname, (_cat, myname) in mp.items():
            v = ((mie[kind].get(mname, {}).get("人数") or {}).get("令和7年度"))
            if v:
                mie_nin[(myname, kind)] = v
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
            name = disp(svc, kind)
            ken7 = r["_R7件数"]
            # 伸び率。令和7年度の件数に対する見込件数の比
            nobi = [(r["見込"][k][0] / ken7) if ken7 else 0.0
                    for k in range(1, len(T.EST_YEARS))]
            # 人数の水準。見える化の受給者数を優先し、なければ件数÷12
            nin7 = mie_nin.get((name, kind), ken7 / 12)
            nin = [nin7] + [nin7 * g for g in nobi]
            # 量の水準。量を載せるサービスかどうかは見える化の行の有無で決める。
            # 月額（日額）の包括報酬のサービス（介護予防通所リハビリテーション・
            # 小規模多機能型居宅介護・認知症対応型共同生活介護・施設・支援）は人数だけ。
            unit, ryo, ryo7 = "人", None, None
            if (name, kind) in mie_vu:
                unit = mie_vu[(name, kind)]
                pp = mie_pp.get((name, kind))
                if pp and nin7:
                    ryo7 = pp[1] * nin7
                else:                       # 見える化に実績がない。月報の日数で補う
                    d7 = a7.get((svc, kind), {})
                    if d7.get("件数", 0) and d7.get("日数", 0):
                        ryo7 = d7["日数"] / 12
                ryo = [ryo7 or 0.0] + [(ryo7 or 0.0) * g for g in nobi]
            rows.append((cat, name, unit, nin, ryo,
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
    plan = plan_rows()
    sheet_keikaku(wb, plan)
    sheet_teiin(wb, plan)
    sheet_kijun(wb, check_base(verbose=False))
    sheet_sogo(wb)
    sheet_zentei(wb, plan)
    sheet_check(wb, chk, a7, proj)
    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / f"小野町_第10期_サービス見込量_{ASOF}.xlsx"
    wb.save(p)

    print(f"生成: {p.name}（{len(wb.sheetnames)}シート）")
    A = plan["集計"]
    std3 = sum(A["標準給付費"])
    chi3 = sum(v[3] for v in chiiki_plan()) * 3 / 1000
    sogo3 = sum(v[3] for v in chiiki_plan() if v[0] == "総合事業") * 3 / 1000
    pop3 = sum(T.POP_DAI10[y]["1号"] for y in T.PLAN_YEARS)
    print(f"総給付費 3年計 {sum(A['計']):,.0f}千円／"
          f"標準給付費 {std3:,.0f}千円／地域支援事業費 {chi3:,.0f}千円")
    for kikin, lab in ((0, "取崩なし"), (12000, "12,000千円取崩"),
                       (35000, "35,000千円取崩")):
        print(f"  保険料基準額（月額）{lab:<16}"
              f"{T.premium(std3, chi3, sogo3, pop3, kikin=kikin)['月額']:>8,.0f}円")
    print(f"  保険料基準額（月額）負担割合24％      "
          f"{T.premium(std3, chi3, sogo3, pop3, futan=0.24)['月額']:>8,.0f}円")
    print("\n令和11年度の見込量（月あたり・主なもの）")
    for kind in ("介護", "予防"):
        for _cat, name, unit, nin, ryo, _k in plan[kind]:
            if name not in ("訪問介護", "通所介護", "短期入所生活介護",
                            "介護老人福祉施設", "認知症対応型共同生活介護"):
                continue
            if ryo:
                print(f"  {name:<22} {ryo[-1]:>7,.0f} {unit}/月"
                      f"　（{nin[-1]:.1f} 人/月）")
            else:
                print(f"  {name:<22} {nin[-1]:>7.1f} 人/月")


if __name__ == "__main__":
    main()
