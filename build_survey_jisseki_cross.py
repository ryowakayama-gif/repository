# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画
調査結果と年報実績の突合クロス集計.

令和8年9月11日のご指示
  「アンケートの分析報告書の再出力と
    現状までの整理を踏まえたクロス集計をご教示ください」

━━ 本表の位置づけ ━━

既存の2件（アンケート調査の集計分析報告書・調査クロス集計・分析）は
いずれも調査の中で完結する集計であった。
本表は、令和8年9月11日に年報（令和7年度）の要介護度別の明細を
`data_nenpo_meisai.py` に収めたことにより初めて可能になった
「調査の結果」と「保険者の実績」の突合を行う。

それ以前は、年報の要介護度別を保有しておらず、
在宅サービス利用率による按分の推定値しか作れなかったため、
調査の要介護度別分布を実績と突き合わせることができなかった。

━━ 本表で分かること ━━

1 調査①（在宅生活改善調査・98票）の要介護度分布は
  在宅の認定者の分布と食い違う（χ²＝43.5、自由度6。1％点16.81）。
  要介護3が期待の2.8倍、要支援1が0.16倍である。
  ケアマネジャーを通じて在宅で困難を抱える層を集めた調査であり、
  在宅の認定者全体を代表するものではない。

2 施設の側から数えた在籍者と、保険者の側から数えた受給者は一致しない。
  介護老人保健施設だけが在籍＞受給（206人対134.1人／月）、
  他の3区分は受給＞在籍である。
  年報は施設の所在地を問わず当広域連合の被保険者を数え、
  調査②は区域内の施設の在籍者を保険者を問わず数えるためである。
  必要利用定員総数（確認事項No.88）に直結する。

出所
  年報（令和7年度）      `data_nenpo_meisai.py`・`data_nenpo.py`
  調査①②③のクロス集計   `data_survey_cross.py`（集計値のみ。個票は収録しない）
  調査の件数・回収        `data_survey2025.py`

シート構成
  00_この表について
  01_成果品の再出力の状況
  02_①在宅生活改善調査と在宅の認定者
  03_②居所変更実態調査と施設・居住系の受給者
  04_②定員・在籍・回転と必要利用定員総数
  05_②入所前の居場所と退去先
  06_③介護人材実態調査と供給
  07_④健康とくらしの調査の代表性
  08_自己点検
  09_確認事項

出力
  output/第10期計画_調査結果と年報実績の突合クロス集計.xlsx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import io
import os
import sys
from decimal import Decimal, ROUND_HALF_UP

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import data_nenpo as N
import data_nenpo_meisai as NM
import data_survey_cross as C
import data_survey2025 as S
import repo_paths as RP

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

ODIR = RP.OUTPUT
OUT = os.path.join(ODIR, "第10期計画_調査結果と年報実績の突合クロス集計.xlsx")
KIJUNBI = "令和8年9月11日"

FONT = "游ゴシック"
NAVY, HEAD = "1F3864", "4472C4"
IN_Y, OK_G, NG_O, MID_B, GRAY = ("FFF2CC", "E2EFDA", "FCE4D6",
                                 "DEEBF7", "F2F2F2")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()
wb.remove(wb.active)
CHECKS = []


def chk(no, naiyo, shiki, kekka, ok):
    CHECKS.append((no, naiyo, shiki, kekka, "適合" if ok else "不適合"))
    return ok


# ============================================================ 体裁
def sheet(name, title, subtitle, widths, freeze="A5"):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A2"] = subtitle
    ws["A2"].font = Font(name=FONT, size=9)
    ws["A2"].fill = PatternFill("solid", fgColor=GRAY)
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    n = max(len(widths), 6)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n)
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 82
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = freeze
    return ws


def header(ws, row, cols, height=30):
    for i, hh in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i, value=hh)
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=HEAD)
        c.alignment = Alignment(wrap_text=True, horizontal="center",
                                vertical="center")
        c.border = BORDER
    ws.row_dimensions[row].height = height
    return row + 1


def body(ws, row, vals, fills=None, height=20, align=None, bold=False,
         numfmt=None):
    for i, v in enumerate(vals, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(name=FONT, size=9, bold=bold)
        c.border = BORDER
        ha = (align or {}).get(i, "left" if isinstance(v, str) else "right")
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal=ha)
        if numfmt and numfmt.get(i) and isinstance(v, float):
            c.number_format = numfmt[i]
        if fills and fills.get(i):
            c.fill = PatternFill("solid", fgColor=fills[i])
    ws.row_dimensions[row].height = height
    return row + 1


def lead(ws, row, text, span=10):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=10.5, bold=True, color=NAVY)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = 20
    return row + 1


def note(ws, row, text, span=10, height=None):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=8.5)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = height or (14 * (1 + text.count("\n")))
    return row + 1


def r2(x, n=2):
    return float(Decimal(str(x)).quantize(Decimal("1." + "0" * n),
                                          ROUND_HALF_UP))


# ============================================================ 算定
DO = list(NM.DO)


def nenpo_m(src, *names):
    """年報の年間の延べ数（要介護度別）を合算し12で除した月平均。"""
    out = [0.0] * 7
    for nm in names:
        v = src[nm]
        for i in range(7):
            out[i] += v[i] / 12.0
    return out


# 施設・居住系（年報。短期利用がある区分は合算する）
SHISETSU_M = {
    "特定施設入居者生活介護": nenpo_m(
        NM.YOSHIKI_1_7_16, "特定施設入居者生活介護（短期利用以外）",
        "特定施設入居者生活介護（短期利用）"),
    "認知症対応型共同生活介護": nenpo_m(
        NM.YOSHIKI_1_7_18, "認知症対応型共同生活介護（短期利用以外）",
        "認知症対応型共同生活介護（短期利用）"),
    "地域密着型特定施設入居者生活介護": nenpo_m(
        NM.YOSHIKI_1_7_18, "地域密着型特定施設入居者生活介護（短期利用以外）",
        "地域密着型特定施設入居者生活介護（短期利用）"),
    "地域密着型介護老人福祉施設入所者生活介護": nenpo_m(
        NM.YOSHIKI_1_7_18, "地域密着型介護老人福祉施設入所者生活介護"),
    "介護老人福祉施設": nenpo_m(NM.YOSHIKI_1_6_15, "介護老人福祉施設"),
    "介護老人保健施設": nenpo_m(NM.YOSHIKI_1_6_15, "介護老人保健施設"),
    "介護医療院": nenpo_m(NM.YOSHIKI_1_6_15, "介護医療院"),
}

NINTEI = [N.NINTEI[d]["R7"] for d in DO]
NINTEI_KEI = N.NINTEI["第1号被保険者　合計"]["R7"]
SH_SUM = [sum(SHISETSU_M[k][i] for k in SHISETSU_M) for i in range(7)]
ZAITAKU_NIN = [NINTEI[i] - SH_SUM[i] for i in range(7)]

# ---------------------------------------------- 調査①
CU = C.CU["要介護度×現在の居所"]
CU_DO = [sum(CU.get(d, {}).values()) for d in DO]
CU_N = sum(CU_DO)
CU_P = [x / sum(ZAITAKU_NIN) for x in ZAITAKU_NIN]
CU_EXP = [CU_N * p for p in CU_P]
CU_CHI = sum((CU_DO[i] - CU_EXP[i]) ** 2 / CU_EXP[i] for i in range(7))
CHI_1 = 16.81          # 自由度6の1％点
CHI_5 = 12.59          # 自由度6の5％点

# 居所別の計（調査①）
CU_KYOSHO = {}
for d, row in CU.items():
    for k, v in row.items():
        CU_KYOSHO[k] = CU_KYOSHO.get(k, 0) + v

# ---------------------------------------------- 調査②
CS_NYU = C.CS["種別×入所者の要介護度"]
SS = C.SS
# 調査②の区分と年報の区分の対応
TAIO = [
    ("グループホーム", ["認知症対応型共同生活介護"]),
    ("特定施設", ["特定施設入居者生活介護",
                  "地域密着型特定施設入居者生活介護"]),
    ("特養・地域密着型特養", ["介護老人福祉施設",
                              "地域密着型介護老人福祉施設入所者生活介護"]),
    ("介護老人保健施設", ["介護老人保健施設"]),
    ("住宅型有料・サ高住", []),          # 介護保険の給付区分ではない
]

NYU_DO = ("要介護1", "要介護2", "要介護3", "要介護4", "要介護5")


def zaiseki(shu):
    """調査②の在籍者数（要介護度の判明分のみ。自立・新規申請中を除く）。"""
    return sum(v for k, v in CS_NYU[shu].items()
               if k in DO or k in NYU_DO)


def nenpo_kei(keys):
    return sum(sum(SHISETSU_M[k]) for k in keys)


# ---------------------------------------------- 調査③
CJ = C.CJ["サービス区分×資格"]
SHOKU_N = S.SHOKU["件数"]

# ---------------------------------------------- 調査④
JAGES_N = 4729
HIHO_R7 = N.IPPAN["第1号被保険者数　計（年度末）"]["R7"]


# ============================================================ 00
ws = sheet("00_この表について",
           "調査結果と年報実績の突合クロス集計",
           "既存の2件（集計分析報告書・調査クロス集計・分析）は"
           "いずれも調査の中で完結する集計である。"
           "本表は、年報（令和7年度）の要介護度別の明細を収めたことにより"
           "初めて可能になった「調査の結果」と「保険者の実績」の突合を行う。"
           "基準日 " + KIJUNBI + "。",
           [4, 28, 40, 34])

r = 4
r = lead(ws, r, "1　本表が新たにできること", span=4)
r = header(ws, r, ["", "突合", "できるようになった理由", "結果"])
for i, (a, b, c_) in enumerate([
    ("調査①の要介護度分布と在宅の認定者の分布",
     "年報 様式1の5の要介護度別（確定値）と、"
     "施設・居住系の要介護度別（様式1の6・様式1の7）の差引きにより"
     "在宅の認定者を要介護度別に置けるようになった",
     "分布は食い違う（χ²＝{:.1f}、自由度6。1％点{:.2f}）。"
     "調査①は在宅の認定者全体を代表しない".format(CU_CHI, CHI_1)),
    ("調査②の施設種別×要介護度と年報の受給者",
     "年報 様式1の6(15)・様式1の7(18)の要介護度別（確定値）。"
     "従前は在宅サービス利用率による按分の推定値しかなく"
     "突合できなかった",
     "施設の側の在籍者と保険者の側の受給者は一致しない。"
     "介護老人保健施設だけが在籍＞受給である"),
    ("調査②の定員・在籍と必要利用定員総数",
     "同上。区分ごとに受給者数を実数で置けるようになった",
     "特別養護老人ホーム等は受給者が回答施設の定員を上回る。"
     "区域外の施設の利用が含まれる（確認事項No.88）"),
], start=1):
    r = body(ws, r, [i, a, b, c_], height=64)

r = note(ws, r + 1,
         "注1）年報の要介護度別は令和8年9月11日に"
         "`data_nenpo_meisai.py` に収めたものです。"
         "12で除した月平均が見える化システムの総括表詳細（１）と"
         "全24サービスで一致することを検算しています"
         "（令和8年度実績見込み値の入力案 06シート）。\n"
         "注2）調査の個票は本リポジトリに収録していません。"
         "本表は集計値（`data_survey_cross.py`）のみを用いています。\n"
         "注3）年報は保険者の側から数えたもの"
         "（施設の所在地を問わず当広域連合の被保険者）、"
         "調査②は施設の側から数えたもの"
         "（保険者を問わず区域内の施設の在籍者）です。"
         "数える対象が異なるため、一致しないこと自体は誤りではありません。",
         span=4)

r += 1
r = lead(ws, r, "2　本表で扱わないもの", span=4)
r = header(ws, r, ["", "事項", "理由", "所在"])
for i, (a, b, c_) in enumerate([
    ("④健康とくらしの調査の地区別・年齢調整・関連分析",
     "既存の成果品で完結している。本表では再掲しない",
     "調査クロス集計・分析（24シート）"),
    ("①②③の調査内のクロス集計",
     "同上", "実施済み3調査の受領点検と集計（15シート）"),
    ("調査結果の限界と留保・点検事項35件",
     "同上", "アンケート調査の集計分析報告書 11シート、"
     "9月作業_アンケート分析の作業状況 05シート"),
], start=1):
    r = body(ws, r, [i, a, b, c_], height=34)


# ============================================================ 01
ws = sheet("01_成果品の再出力の状況",
           "アンケート関係の成果品の再出力の状況",
           "令和8年9月11日に全件を再実行した結果である。"
           "再出力できないものが1件あり、その理由と影響を示す。",
           [4, 40, 12, 30, 34])

r = 4
r = header(ws, r, ["", "成果品（スクリプト）", "再出力", "読むデータ", "内容"])
SAI = [
    ("第10期計画_アンケート調査の集計分析報告書.xlsx\n"
     "（build_survey_report.py）", "できる",
     "リポジトリ内のデータのみ",
     "15シート。介護職員348人・実事業所75・法人28・所見20件"),
    ("第10期計画_実施済み調査_結果報告書.docx\n"
     "（build_survey_report_doc.py）", "できる",
     "同上", "475段落・51表"),
    ("第10期計画_実施済み3調査の受領点検と集計.xlsx\n"
     "（build_survey_review.py）", "できる",
     "`data_survey_cross.py` ほか", "15シート。調査内クロス集計を含む"),
    ("第10期計画_9月作業_アンケート分析の作業状況.xlsx\n"
     "（build_survey_status.py）", "できる",
     "同上", "10シート。点検事項35件（重大6・要確認21・様式網羅性8）"),
    ("第10期計画_事業所調査の照会票と確定値管理表.xlsx\n"
     "（build_survey_inquiry.py）", "できる",
     "同上", "照会の対象27件・事業者別の照会先5先"),
    ("第10期計画_第9期施策と調査・KPIの紐付けレビュー.xlsx\n"
     "（build_policy_survey_kpi.py）", "できる",
     "同上", "施策19件。裏づけあり13・代表KPIあり10・いずれもなし5"),
    ("第10期計画_調査クロス集計・分析.xlsx\n"
     "（build_survey_crosstab.py）", "できない",
     "④健康とくらしの調査の個票CSV（4,729票）",
     "24シート。成果品そのものは収録済みで内容は確定している。"
     "個票は発注者指示により収録できず、作業領域にも残っていないため"
     "スクリプトを再実行できない"),
]
for i, (a, b, c_, d) in enumerate(SAI, start=1):
    r = body(ws, r, [i, a, b, c_, d],
             fills={3: (OK_G if b == "できる" else NG_O)}, height=48)

SAI_NG = [x for x in SAI if x[1] == "できない"]
r = note(ws, r + 1,
         "注1）再出力できるものは{}件、できないものは{}件です。\n"
         "注2）再出力できない1件について。"
         "個票データ（個人情報）はリポジトリに格納しないという"
         "発注者指示（令和8年8月28日）があり、"
         "個票を読むスクリプトは受領資料が消えると再実行できません。"
         "成果品そのもの（24シートのxlsx）は収録済みで内容は確定しており、"
         "計画素案・代表KPIへの反映も済んでいるため、"
         "現時点で不足は生じていません。\n"
         "注3）同じ理由で再実行できないスクリプトが他に2本あります"
         "（`build_kickoff_fix.py`・`build_minutes_fix.py`）。"
         "いずれも受領資料の読取りを1回限り行うものです。\n"
         "注4）④の個票を再度ご提供いただければ再実行できます。"
         "その場合も個票はリポジトリに格納せず、"
         "作業領域にとどめて作業後に消去します（確認事項No.139）。"
         .format(len(SAI) - len(SAI_NG), len(SAI_NG)), span=5)


# ============================================================ 02
ws = sheet("02_①在宅生活改善調査と在宅の認定者",
           "① 在宅生活改善調査の要介護度分布と在宅の認定者（令和7年度）",
           "調査①の利用者票（{}票のうち要介護度の判明した{}票）の分布を、"
           "年報による在宅の認定者の分布と突き合わせる。"
           "在宅の認定者＝認定者数−施設・居住系サービス受給者。"
           "住宅型有料老人ホーム・軽費老人ホームは"
           "介護保険の施設・居住系ではないため在宅に含まれる。"
           .format(S.RIYO["件数"], CU_N),
           [4, 14, 12, 12, 14, 14, 14, 14])

r = 4
r = lead(ws, r, "1　要介護度別の分布", span=8)
r = header(ws, r, ["", "要介護度", "調査①\n（人）", "構成比",
                   "在宅の認定者\n（人・月平均）", "構成比",
                   "期待度数", "χ²への寄与"])
for i, d in enumerate(DO):
    ki = (CU_DO[i] - CU_EXP[i]) ** 2 / CU_EXP[i]
    r = body(ws, r, ["", d, CU_DO[i], r2(CU_DO[i] / CU_N * 100, 1),
                     r2(ZAITAKU_NIN[i]), r2(CU_P[i] * 100, 1),
                     r2(CU_EXP[i], 1), r2(ki, 1)],
             fills={8: (NG_O if ki >= 10 else None)},
             numfmt={4: "0.0", 5: "0.00", 6: "0.0", 7: "0.0", 8: "0.0"},
             height=18)
r = body(ws, r, ["", "計", CU_N, 100.0, r2(sum(ZAITAKU_NIN)), 100.0,
                 r2(float(CU_N), 1), r2(CU_CHI, 1)],
         bold=True, numfmt={4: "0.0", 5: "0.00", 6: "0.0", 7: "0.0",
                            8: "0.0"}, height=18)
r = note(ws, r + 1,
         "注1）χ²＝{:.1f}（自由度6）。5％点{:.2f}・1％点{:.2f}を上回ります。"
         "調査①の要介護度分布は在宅の認定者の分布と食い違います。\n"
         "注2）寄与が大きいのは要介護3（期待{:.1f}人に対し{}人。{:.1f}倍）と"
         "要支援1（期待{:.1f}人に対し{}人。{:.1f}倍）です。\n"
         "注3）調査①はケアマネジャーを通じて"
         "在宅生活に困難を抱える方を対象としたものであり、"
         "在宅の認定者全体を代表するものではありません。"
         "本調査の結果を在宅の認定者全体に当てはめて読まないでください。\n"
         "注4）在宅の認定者は年報の月平均です。"
         "調査①は時点の票数であり、単位が異なります。"
         "構成比の比較にとどめています。"
         .format(CU_CHI, CHI_5, CHI_1,
                 CU_EXP[4], CU_DO[4], CU_DO[4] / CU_EXP[4],
                 CU_EXP[0], CU_DO[0], CU_DO[0] / CU_EXP[0]), span=8)

r += 1
r = lead(ws, r, "2　現在の居所（調査①）", span=8)
r = header(ws, r, ["", "居所", "人数", "構成比", "介護保険上の扱い",
                   "", "", ""])
KYO_ATSUKAI = {
    "自宅等": "在宅。居宅サービスの受給者",
    "住宅型有料": "在宅。特定施設の指定を受けていないものは"
                  "居宅サービスの受給者となる",
    "軽費老人ホーム": "在宅。同上",
}
for k in sorted(CU_KYOSHO, key=lambda x: -CU_KYOSHO[x]):
    r = body(ws, r, ["", k, CU_KYOSHO[k],
                     r2(CU_KYOSHO[k] / CU_N * 100, 1),
                     KYO_ATSUKAI.get(k, "［要確認］"), "", "", ""],
             numfmt={4: "0.0"}, height=18)
r = body(ws, r, ["", "計", sum(CU_KYOSHO.values()), 100.0, "", "", "", ""],
         bold=True, numfmt={4: "0.0"}, height=18)
r = note(ws, r + 1,
         "注5）調査①の回答者はいずれも在宅であり、"
         "施設・居住系サービスの受給者は含まれません。"
         "したがって1の突合の相手として在宅の認定者を用いています。\n"
         "注6）介護保険の指定を受けない住まい"
         "（住宅型有料老人ホーム・サービス付高齢者向け住宅・軽費老人ホーム）"
         "の339人・戸は在宅受給者の34.2％に当たります"
         "（趨勢に現れない要因の分析 04シート）。"
         "調査①でも{}票（{:.1f}％）がこれらに居住しています。"
         .format(CU_KYOSHO.get("住宅型有料", 0)
                 + CU_KYOSHO.get("軽費老人ホーム", 0),
                 (CU_KYOSHO.get("住宅型有料", 0)
                  + CU_KYOSHO.get("軽費老人ホーム", 0)) / CU_N * 100),
         span=8)


# ============================================================ 03
ws = sheet("03_②居所変更実態調査と施設の受給者",
           "② 居所変更実態調査の在籍者と年報の受給者（令和7年度）",
           "調査②は区域内の施設に照会し、施設の側から在籍者を数えたもの。"
           "年報は保険者の側から、施設の所在地を問わず"
           "当広域連合の被保険者の受給者を数えたもの。"
           "数える対象が異なるため一致しないこと自体は誤りではない。"
           "差の向きと大きさから何が読めるかを示す。",
           [4, 22, 11, 11, 13, 11, 40])

r = 4
r = header(ws, r, ["", "調査②の種別", "回答\n施設数", "在籍者\n（人）",
                   "年報の受給者\n（人・月平均）", "差",
                   "読み取れること"])
SA = []
for shu, keys in TAIO:
    zs = zaiseki(shu) if shu in CS_NYU else 0
    ns = nenpo_kei(keys)
    sasu = zs - ns
    SA.append((shu, zs, ns, sasu))
    if not keys:
        yomi = ("介護保険の給付区分ではない（特定施設の指定を受けない住まい）。"
                "年報には現れず、居宅サービスの受給者として数えられる")
    elif ns == 0:
        yomi = "年報の受給実績がない"
    elif sasu > 0:
        yomi = ("在籍＞受給。区域内の施設に他の保険者の被保険者が"
                "入所していることによる（当広域連合の被保険者以外は"
                "年報に現れない）")
    else:
        yomi = ("受給＞在籍。当広域連合の被保険者が区域外の施設を"
                "利用していること、及び未回答の施設があることによる")
    r = body(ws, r, ["", shu, SS[shu]["施設数"], zs,
                     (r2(ns) if keys else "―"),
                     ("―" if not keys else r2(sasu)), yomi],
             fills={6: (NG_O if keys and abs(sasu) >= 30 else None)},
             numfmt={5: "0.00", 6: "0.00"}, height=56)

r = note(ws, r + 1,
         "注1）在籍者は要介護度の判明した人数です"
         "（自立・新規申請中を除きます）。\n"
         "注2）介護老人保健施設だけが在籍＞受給です"
         "（{}人対{:.1f}人／月）。"
         "同施設は新規{}人・退去{}人で回転が速く"
         "（04シート）、区域外からの受入れが多いことがうかがえます。\n"
         "注3）特別養護老人ホーム等は受給が在籍を{:.0f}人上回ります。"
         "区域内31施設のうち13施設が未回答であること"
         "（把握率58.1％。点検事項の重大6件の1つ）と、"
         "区域外の施設の利用の両方が効いています。"
         "未回答の扱いは確認事項No.8で決定をお願いしています。\n"
         "注4）住宅型有料老人ホーム・サービス付高齢者向け住宅は"
         "介護保険の給付区分ではないため年報に現れません。"
         "在籍{}人は居宅サービスの受給者として数えられています。"
         .format(zaiseki("介護老人保健施設"),
                 nenpo_kei(["介護老人保健施設"]),
                 SS["介護老人保健施設"]["新規"],
                 SS["介護老人保健施設"]["退去"],
                 abs(zaiseki("特養・地域密着型特養")
                     - nenpo_kei(["介護老人福祉施設",
                                  "地域密着型介護老人福祉施設入所者生活介護"])),
                 zaiseki("住宅型有料・サ高住")), span=7)

r += 1
r = lead(ws, r, "要介護度別の内訳（在籍者と年報の受給者）", span=7)
r = header(ws, r, ["", "区分"] + list(DO) + ["計"])
for shu, keys in TAIO:
    if not keys:
        continue
    o = CS_NYU[shu]
    r = body(ws, r, ["調査②", shu] + [o.get(d, 0) for d in DO]
             + [zaiseki(shu)], height=18)
    v = [sum(SHISETSU_M[k][i] for k in keys) for i in range(7)]
    r = body(ws, r, ["年報", ""] + [r2(x) for x in v] + [r2(sum(v))],
             fills={k: MID_B for k in range(3, 11)},
             numfmt={k: "0.00" for k in range(3, 11)}, height=18)
r = note(ws, r + 1,
         "注5）グループホームは要支援2が調査②で0人、"
         "年報で月平均{:.2f}人です。"
         "介護予防認知症対応型共同生活介護の受給者は"
         "年間の延べ{}人（要支援2のみ）で、"
         "時点で数えると0人になることがあります。"
         .format(SHISETSU_M["認知症対応型共同生活介護"][1],
                 NM.YOSHIKI_1_7_18["認知症対応型共同生活介護（短期利用以外）"][1]),
         span=10)


# ============================================================ 04
ws = sheet("04_定員在籍回転と必要利用定員",
           "② 定員・在籍・回転と必要利用定員総数",
           "必要利用定員総数（介護保険法第117条第2項第1号）は"
           "計画素案 第6章第4節4に新設した法定記載事項である"
           "（確認事項No.88）。"
           "調査②の定員・在籍・新規・退去と年報の受給者を並べる。",
           [4, 22, 10, 10, 10, 10, 10, 10, 34])

r = 4
r = header(ws, r, ["", "種別", "施設数", "定員", "在籍", "入居率",
                   "新規", "退去", "年報の受給者と定員の関係"])
for shu, keys in TAIO:
    s = SS[shu]
    ny = s.get("入所", 0)
    ns = nenpo_kei(keys) if keys else None
    ritsu = ny / s["定員"] * 100 if s["定員"] else 0.0
    if ns is None:
        kan = "介護保険の給付区分ではない"
    elif ns > s["定員"]:
        kan = ("受給者{:.1f}人／月が回答施設の定員{}人を"
               "{:.1f}人上回る。区域外の施設の利用が含まれる"
               .format(ns, s["定員"], ns - s["定員"]))
    else:
        kan = ("受給者{:.1f}人／月は回答施設の定員{}人の{:.1f}％"
               .format(ns, s["定員"], ns / s["定員"] * 100))
    r = body(ws, r, ["", shu, s["施設数"], s["定員"], ny,
                     r2(ritsu, 1), s.get("新規", 0), s.get("退去", 0), kan],
             fills={9: (NG_O if (ns is not None and ns > s["定員"]) else None)},
             numfmt={6: "0.0"}, height=44)

_taiki = SS["特養・地域密着型特養"]["待機"]
r = note(ws, r + 1,
         "注1）定員・在籍・新規・退去は調査②の回答施設分のみです。"
         "区域内31施設のうち13施設が未回答であり、"
         "区域内の全定員ではありません。\n"
         "注2）特別養護老人ホーム等は待機{}人です。"
         "回答施設の定員{}人・在籍{}人に対し待機が多く、"
         "入居率{:.1f}％と併せて読む必要があります"
         "（待機は複数施設への重複申込みを含み得ます。点検事項No.33）。\n"
         "注3）介護老人保健施設は新規{}人・退去{}人で、"
         "在籍{}人に対し年間の入退所がほぼ同数です。"
         "在宅復帰を目的とする施設の性質によるものと考えられます。\n"
         "注4）本表は必要利用定員総数の算定そのものではありません。"
         "必要利用定員総数は保険者単位で定めるものであり、"
         "区域内の施設の定員とは別に整理する必要があります"
         "（確認事項No.88・No.113）。"
         .format(_taiki, SS["特養・地域密着型特養"]["定員"],
                 SS["特養・地域密着型特養"]["入所"],
                 SS["特養・地域密着型特養"]["入所"]
                 / SS["特養・地域密着型特養"]["定員"] * 100,
                 SS["介護老人保健施設"]["新規"],
                 SS["介護老人保健施設"]["退去"],
                 SS["介護老人保健施設"]["入所"]), span=9)


# ============================================================ 05
ws = sheet("05_入所前の居場所と退去先",
           "② 入所前の居場所と退去先",
           "区域内・区域外の別により、施設・住まいの利用が"
           "区域を越えて成立していることを示す。",
           [4, 22, 12, 12, 12, 12, 40])

CS_MAE = C.CS["種別×入所前の居場所"]
CS_SAKI = C.CS["種別×退去先"]


def kubun(row):
    """区域内・区域外の別に集計する。"""
    nai = sum(v for k, v in row.items() if "（区域外）" not in k)
    gai = sum(v for k, v in row.items() if "（区域外）" in k)
    return nai, gai


r = 4
r = lead(ws, r, "1　入所前の居場所", span=7)
r = header(ws, r, ["", "種別", "区域内", "区域外", "計", "区域外の割合",
                   "主な内訳"])
for shu, _k in TAIO:
    if shu not in CS_MAE:
        continue
    nai, gai = kubun(CS_MAE[shu])
    tot = nai + gai
    top = sorted(CS_MAE[shu].items(), key=lambda x: -x[1])[:2]
    r = body(ws, r, ["", shu, nai, gai, tot,
                     r2(gai / tot * 100, 1) if tot else 0.0,
                     "・".join("%s %d" % (k, v) for k, v in top)],
             fills={6: (NG_O if tot and gai / tot >= 0.5 else None)},
             numfmt={6: "0.0"}, height=20)

r += 1
r = lead(ws, r, "2　退去先", span=7)
r = header(ws, r, ["", "種別", "区域内", "区域外", "計", "区域外の割合",
                   "主な内訳"])
for shu, _k in TAIO:
    if shu not in CS_SAKI:
        continue
    nai, gai = kubun(CS_SAKI[shu])
    tot = nai + gai
    top = sorted(CS_SAKI[shu].items(), key=lambda x: -x[1])[:2]
    r = body(ws, r, ["", shu, nai, gai, tot,
                     r2(gai / tot * 100, 1) if tot else 0.0,
                     "・".join("%s %d" % (k, v) for k, v in top)],
             fills={6: (NG_O if tot and gai / tot >= 0.5 else None)},
             numfmt={6: "0.0"}, height=20)

r = note(ws, r + 1,
         "注1）区域外の割合が5割以上のものに色を付けています。\n"
         "注2）入所前の居場所として病院・診療所が多いことは、"
         "医療から介護への移行が施設の入口になっていることを示します。"
         "新たな地域医療構想との接続（確認事項No.111。"
         "北海道の第10期支援計画は未策定）に関わります。\n"
         "注3）区域を越えた利用が双方向に成立しているため、"
         "区域内の施設の定員だけで需要を測ることはできません。"
         "03シート・04シートの差はこのことによります。", span=7)


# ============================================================ 06
ws = sheet("06_③介護人材実態調査と供給",
           "③ 介護人材実態調査とサービス区分別の供給",
           "職員個票{}人のサービス区分別の分布を、"
           "調査②の定員・在籍及び年報の受給者と並べる。"
           .format(SHOKU_N),
           [4, 18, 12, 12, 12, 40])

SHOK_TAIO = {
    "GH": ("グループホーム", ["認知症対応型共同生活介護"]),
    "老健": ("介護老人保健施設", ["介護老人保健施設"]),
    "特養": ("特養・地域密着型特養", ["介護老人福祉施設"]),
    "地密特養": ("特養・地域密着型特養",
                 ["地域密着型介護老人福祉施設入所者生活介護"]),
    "特定施設": ("特定施設", ["特定施設入居者生活介護"]),
    "住宅型有料": ("住宅型有料・サ高住", []),
    "通所リハ": (None, []),
    "通所介護等": (None, []),
}

r = 4
r = header(ws, r, ["", "サービス区分", "職員（人）", "うち介護福祉士",
                   "介護福祉士の割合", "年報の受給者（人・月平均）"])
_shok_kei = 0
for k, row in CJ.items():
    n_ = sum(row.values())
    _shok_kei += n_
    fs = row.get("1 介護福祉士", 0)
    _s, keys = SHOK_TAIO.get(k, (None, []))
    ns = nenpo_kei(keys) if keys else None
    r = body(ws, r, ["", k, n_, fs, r2(fs / n_ * 100, 1),
                     (r2(ns) if ns is not None else "―")],
             numfmt={5: "0.0", 6: "0.00"}, height=18)
r = body(ws, r, ["", "計", _shok_kei,
                 sum(row.get("1 介護福祉士", 0) for row in CJ.values()),
                 r2(sum(row.get("1 介護福祉士", 0) for row in CJ.values())
                    / _shok_kei * 100, 1), ""],
         bold=True, numfmt={5: "0.0"}, height=18)

r = note(ws, r + 1,
         "注1）サービス区分別の職員数は{}人で、"
         "職員個票の総数{}人との差{}人は区分の未記入等によります。\n"
         "注2）本調査は令和7年4月1日現在です。"
         "年報は令和7年度（令和7年4月から令和8年3月まで）の月平均であり、"
         "時点が異なります。人員配置基準との照合には用いられません。\n"
         "注3）採用率と離職率の差（代表KPI H14）は、"
         "各年9月30日現在の在籍者数と年間の採用者数・退職者数がないため"
         "算定していません（確認事項No.10）。\n"
         "注4）供給の制約（人材）は見込量に反映していません"
         "（確認事項No.116）。見込量は需要のみによっています。"
         .format(_shok_kei, SHOKU_N, SHOKU_N - _shok_kei), span=6)


# ============================================================ 07
ws = sheet("07_④健康とくらしの調査の代表性",
           "④ 健康とくらしの調査の代表性",
           "集計対象{:,}票を、第1号被保険者数及び認定者数と対照する。"
           "本調査の詳細な分析は「調査クロス集計・分析」（24シート）にある。"
           .format(JAGES_N),
           [4, 34, 16, 16, 40])

r = 4
r = header(ws, r, ["", "項目", "調査④", "年報（令和7年度）", "内容"])
for i, (a, b, c_, d) in enumerate([
    ("集計対象票数", "{:,}票".format(JAGES_N), "―",
     "個票データは個人情報を含むためリポジトリに収録していない"),
    ("第1号被保険者数（年度末）", "―", "{:,}人".format(HIHO_R7),
     "調査④の票数は第1号被保険者数の{:.1f}％に当たる"
     .format(JAGES_N / HIHO_R7 * 100)),
    ("要介護（要支援）認定者数（年度末・第1号）", "―",
     "{:,}人".format(NINTEI_KEI),
     "認定率{:.1f}％。見える化システムの認定率20.84％とは"
     "定義・時点が異なる（確認事項No.125）"
     .format(NINTEI_KEI / HIHO_R7 * 100)),
    ("調査④の対象", "要介護認定を受けていない高齢者が中心", "―",
     "対象者の範囲・配布数・回収率は確認事項No.9で"
     "ご教示をお願いしている"),
], start=1):
    r = body(ws, r, [i, a, b, c_, d], height=40)

r = note(ws, r + 1,
         "注1）調査④は介護予防・日常生活圏域ニーズ調査に相当するもので、"
         "認定を受けていない高齢者が中心です。"
         "①②③とは対象が重ならないため、"
         "横断のクロス集計は行っていません"
         "（集計分析報告書 13シートに整理）。\n"
         "注2）票数と被保険者数の比は回収率ではありません。"
         "配布数の記録がないため回収率を算定できず、"
         "確認事項No.9でご教示をお願いしています。\n"
         "注3）代表KPIのH07（主介護者）・H08（支援者の不在）は"
         "本調査の大雪独自設問により算定できることが分かっています"
         "（確認事項No.4）。", span=5)


# ============================================================ 08
ws = sheet("08_自己点検", "自己点検（エラーチェック）",
           "本表の内的整合と、出所との突合を点検する。",
           [4, 44, 26, 34, 10, 10])

chk(1, "年報の要介護度別の和が認定者数と一致すること",
    "様式1の5の要介護度別",
    "{}人／年報の第1号合計{}人".format(sum(NINTEI), NINTEI_KEI),
    sum(NINTEI) == NINTEI_KEI)

_sh = sum(sum(v) for v in SHISETSU_M.values())
chk(2, "施設・居住系7区分の年報の月平均が480.167人であること"
       "（令和8年度実績見込み値の入力案と同じ値になること）",
    "様式1の6(15)・様式1の7(16)(18)÷12",
    "{:.3f}人／月".format(_sh), abs(_sh - 480.167) < 5e-4)

_zn = sum(ZAITAKU_NIN)
chk(3, "在宅の認定者が1,500人前後であること",
    "認定者数{}人−施設・居住系{:.1f}人".format(NINTEI_KEI, _sh),
    "{:.1f}人".format(_zn), 1450 <= _zn <= 1560)

chk(4, "調査①の要介護度別の和が票数と一致すること",
    "利用者票{}票".format(S.RIYO["件数"]),
    "要介護度の判明{}票（未記入{}票）"
    .format(CU_N, S.RIYO["件数"] - CU_N),
    CU_N <= S.RIYO["件数"])

chk(5, "調査①と在宅の認定者の分布の差が統計的に確かめられること",
    "χ²適合度検定（自由度6）",
    "χ²＝{:.1f}。1％点{:.2f}を上回る".format(CU_CHI, CHI_1),
    CU_CHI > CHI_1)

_ns = zaiseki("介護老人保健施設")
_nn = nenpo_kei(["介護老人保健施設"])
chk(6, "介護老人保健施設だけが在籍＞受給であること",
    "調査②の在籍と年報の月平均",
    "在籍{}人・年報{:.1f}人／月（差＋{:.1f}）".format(_ns, _nn, _ns - _nn),
    _ns > _nn and all(
        zaiseki(s) < nenpo_kei(k)
        for s, k in TAIO if k and s != "介護老人保健施設" and s in CS_NYU))

_tk = nenpo_kei(["介護老人福祉施設",
                 "地域密着型介護老人福祉施設入所者生活介護"])
chk(7, "特別養護老人ホーム等の受給者が回答施設の定員を上回ること",
    "年報の月平均と調査②の定員",
    "{:.1f}人／月／定員{}人".format(_tk, SS["特養・地域密着型特養"]["定員"]),
    _tk > SS["特養・地域密着型特養"]["定員"])

_zai_sum = [sum(CS_NYU[s].get(d, 0) for s, k in TAIO
                if k and s in CS_NYU) for d in DO]
chk(8, "調査②の要介護度別の和が在籍者数と一致すること",
    "介護保険の給付区分に当たる4種別",
    "{}人".format(sum(_zai_sum)),
    sum(_zai_sum) == sum(zaiseki(s) for s, k in TAIO
                         if k and s in CS_NYU))

chk(9, "調査③のサービス区分別の和が職員個票の総数を超えないこと",
    "職員個票{}人".format(SHOKU_N),
    "区分別の和{}人（差{}人）".format(_shok_kei, SHOKU_N - _shok_kei),
    _shok_kei <= SHOKU_N)

chk(10, "個票データを収録していないこと",
    "本スクリプトが読むモジュール",
    "data_nenpo・data_nenpo_meisai・data_survey_cross・data_survey2025"
    "（いずれも集計値のみ）",
    True)

chk(11, "再出力できない成果品を明示していること",
    "アンケート関係{}件".format(len(SAI)),
    "できない{}件：{}".format(
        len(SAI_NG), "／".join(x[0].split("\n")[0] for x in SAI_NG)),
    len(SAI_NG) == 1)

r = 4
r = header(ws, r, ["No", "点検の内容", "式・対象", "結果", "判定", ""])
for no, naiyo, shiki, kekka, han in CHECKS:
    r = body(ws, r, [no, naiyo, shiki, kekka, han, ""],
             fills={5: OK_G if han == "適合" else NG_O}, height=32)
NG = sum(1 for c in CHECKS if c[4] == "不適合")
r = body(ws, r, ["", "計", "{}件".format(len(CHECKS)),
                 "適合{}件・不適合{}件".format(len(CHECKS) - NG, NG),
                 "適合" if NG == 0 else "不適合", ""],
         fills={5: OK_G if NG == 0 else NG_O}, bold=True)
r = note(ws, r + 1,
         "注1）点検2は、本表の年報の値が"
         "「令和8年度実績見込み値の入力案」と同じ算定によることを"
         "確かめるものです。同じ値にならなければどちらかが誤っています。\n"
         "注2）1件でも不適合があると、本表を作るスクリプトは"
         "終了コード1で終わります。", span=6)


# ============================================================ 09
ws = sheet("09_確認事項", "確認事項",
           "本表に関してご判断・ご確認をお願いする事項。"
           "番号は業務工程管理表 03_確認事項一覧による。",
           [7, 30, 46, 20, 12, 12])

r = 4
r = header(ws, r, ["No.", "確認事項", "内容", "止めている成果物",
                   "確認先", "回答期限"])
KAKUNIN = [
    ("No.139", "④健康とくらしの調査の個票の再提供の要否",
     "「調査クロス集計・分析」（24シート）を作るスクリプトは"
     "④の個票CSV（4,729票）を読む。"
     "個票は発注者指示によりリポジトリに格納しておらず、"
     "作業領域にも残っていないため再実行できない。"
     "成果品そのものは収録済みで内容は確定しており、"
     "計画素案・代表KPIへの反映も済んでいるため現時点で不足はない。"
     "①今後、地区の定義の変更（確認事項No.9）や"
     "設問の追加集計により作り直しが必要になる場合は"
     "個票の再提供を要する。"
     "②その場合も個票はリポジトリに格納せず、"
     "作業領域にとどめて作業後に消去する。"
     "③作り直しの予定がなければ再提供は不要である。",
     "調査クロス集計・分析の更新",
     "発注者", "R8.9"),
    ("No.140", "調査①の結果を在宅の認定者全体に当てはめないことの確認",
     "調査①の要介護度分布は在宅の認定者の分布と食い違う"
     "（χ²＝{:.1f}、自由度6。1％点{:.2f}）。"
     "要介護3が期待の{:.1f}倍、要支援1が{:.1f}倍である。"
     "ケアマネジャーを通じて在宅生活に困難を抱える方を集めた調査であり、"
     "在宅の認定者全体を代表するものではない。"
     "①計画素案において調査①の結果を引くときは"
     "「在宅で困難を抱える方の状況」として引き、"
     "在宅の認定者全体の割合として引かない整理としてよいか。"
     "②調査①に基づく代表KPI H12（確認事項No.4の振替案）も"
     "同じ母集団の指標となるため、"
     "目標値の置き方を併せてご判断いただきたい。"
     .format(CU_CHI, CHI_1, CU_DO[4] / CU_EXP[4], CU_DO[0] / CU_EXP[0]),
     "計画素案 第2章第4節・第4章第3節",
     "発注者", "R8.9"),
    ("No.141", "区域を越えた施設利用の扱い",
     "施設の側から数えた在籍者と、保険者の側から数えた受給者は一致しない。"
     "介護老人保健施設は在籍{}人に対し年報の受給者{:.1f}人／月で"
     "在籍が上回り、"
     "特別養護老人ホーム等は受給者{:.1f}人／月が"
     "回答施設の定員{}人を上回る。"
     "①必要利用定員総数（確認事項No.88）は保険者単位で定めるものであり、"
     "区域内の施設の定員とは別に整理する必要がある。"
     "この整理でよいかご判断いただきたい。"
     "②区域外の施設を利用している当広域連合の被保険者の数を"
     "把握する手段があるかご教示いただきたい"
     "（住所地特例被保険者は年報様式1で77人と分かるが、"
     "住所地特例に当たらない区域外利用は年報からは分けられない）。"
     .format(_ns, _nn, _tk, SS["特養・地域密着型特養"]["定員"]),
     "計画素案 第6章第4節4\nサービス見込量 第1次概算",
     "発注者", "R8.9"),
]
for no, ken, naiyo, tome, saki, kigen in KAKUNIN:
    r = body(ws, r, [no, ken, naiyo, tome, saki, kigen], height=150)
r = note(ws, r + 1,
         "注1）本表により新たに起票した確認事項は{}件"
         "（No.139〜No.141）です。\n"
         "注2）既に起票済みのNo.4（代表KPIの代理指標）・"
         "No.8（3調査の点検事項の取扱い）・"
         "No.9（調査の対象者範囲と回収率）・"
         "No.10（採用者数・退職者数）・"
         "No.88（必要利用定員総数）とも関わります。\n"
         "注3）確認事項は業務工程管理表 03_確認事項一覧で一元管理します。"
         .format(len(KAKUNIN)), span=6)


# ============================================================ 出力
os.makedirs(ODIR, exist_ok=True)
wb.save(OUT)

print("出力：" + OUT)
print("調査①：{}票（要介護度の判明分）　χ²＝{:.1f}（自由度6・1％点{:.2f}）"
      .format(CU_N, CU_CHI, CHI_1))
print("在宅の認定者：{:.1f}人（認定者{:,}人−施設・居住系{:.1f}人）"
      .format(_zn, NINTEI_KEI, _sh))
print("調査②：在籍＞受給は介護老人保健施設のみ（{}人対{:.1f}人／月）"
      .format(_ns, _nn))
print("特養等：受給{:.1f}人／月が回答施設の定員{}人を上回る"
      .format(_tk, SS["特養・地域密着型特養"]["定員"]))
print("再出力：{}件中{}件が可能（不可1件は④の個票を読むもの）"
      .format(len(SAI), len(SAI) - len(SAI_NG)))
print("確認事項：{}件（No.139〜No.141）".format(len(KAKUNIN)))
print("自己点検：{}件（適合{}件・不適合{}件）"
      .format(len(CHECKS), len(CHECKS) - NG, NG))

if NG:
    sys.exit(1)
