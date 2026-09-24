# -*- coding: utf-8 -*-
"""北塩原村 第8期障がい福祉計画・第4期障がい児福祉計画　サービス見込量の算定（第1次概算）.

令和8年9月24日のご指示
  「大雪広域及び川崎町のブランチ確認の上、見込量算定について作業を進めて下さい」

大雪地区広域連合・川崎町（いずれも介護保険事業計画）で確立した算定の手順を、
障がい福祉計画に移したものである。手順は次の6つで、本表もこれに従う。

  1 基準年度は完結年度をとる
  2 出所を1つに固定し、他の出所との差を明示する
  3 算定パターンを列挙し、採るパターンを理由とともに決める
  4 決着していない事項は過年度を据え置き、据え置いた件を残らず掲げる
  5 自己点検を設計し、通らなければ終了コード1で落とす
  6 給付費は見込量とは別の系列（年間請求件数）から算定する

受託者の分析であり確定値ではない。

━━ 判明したこと ━━

1 数量の基準年度に使えるのは令和7年度だけである。
  令和8年度は年度途中であり、給付実績が完結していない。
  令和6年度は完結しているが、令和6年度に居宅介護が27件から43件へ、
  児童発達支援が23件から13件へ動いており、令和7年度のほうが直近の姿に近い。

2 母数（手帳所持者数）は減るが、サービスの利用は増えてきた。
  手帳所持者数は中位推計で令和7年154人から令和11年145人へ年約2％の減。
  一方、介護給付の請求件数は令和2年度376件から令和7年度624件へ5年で1.66倍
  （年10.7％）である。両者は相殺の方向にあり、
  母数の伸びをそのまま乗じると見込量が実態より低く出る。

3 実績のトレンド延長は、本村の規模では成り立たない。
  就労継続支援B型は5年で1.93倍だが、直近の令和6年度から令和7年度は
  140件から143件（＋2.1％）で頭打ちである。
  5年平均の伸びを3年続けると令和11年度に22人となり、
  村の18歳到達者数（年0〜1人）と整合しない。

4 障がい児は率で動かす意味がない。
  児童発達支援が令和5年度23件から令和7年度1件へ、
  放課後等デイサービスが令和6年度12件から令和7年度27件へ動いており、
  合計では令和5年度40件・令和7年度39件でほぼ変わらない。
  就学に伴う個人単位の移動であり、年少人口の率とは無関係である。

5 したがって採るのはパターン3（基準年度据え置き＋成果目標による個別補正）である。
  パターン1（母数連動）とパターン2（トレンド延長）は比較のために併記する。
  3か年の給付費は、パターン1が約1.58億円、パターン3が約1.67億円、
  パターン2が約2.14億円であり、採用値はこの幅の中にある。

6 月平均利用者数と年間請求件数を人／月×12で結んではならない。
  短期入所は年3件（月平均0.25件）であり、1人／月と計上しても
  年12件にはならない。計画相談支援は請求が計画作成とモニタリングの
  ときだけ発生するため、令和7年度76件は月平均6.3人ではなく
  実利用者22人に対応する。給付費は年間請求件数から算定する。

7 就労継続支援A型は給付実績がない。
  現行計画と計画素案は1人としているが、令和2〜7年度の給付実績に
  A型の計上がない。村に確認するまで見込量は現行計画を据え置き、
  給付費は計上しない（村確認事項 M-4）。

シート構成
  00_この算定について      結論・採用パターン・据え置きの要約
  01_基準年度の選び方      令和7年度を採る理由と、他年度が使えない理由
  02_算定パターンの比較    3パターンの給付費と、採否の理由
  03_基準年度の実績        令和7年度の請求件数・給付費・単価
  04_見込量_障害福祉サービス
  05_見込量_障害児通所支援等
  06_給付費の見込み        年度別・財源区分別
  07_据え置き一覧          据え置いた件と、解除の条件
  08_自己点検              基準年度の再現・成果目標との整合・現行計画との比較
  09_重度障がい者の内数    別表第一 三・四（改正Cで医療的ケア者等に）
  10_村への確認事項

出力
  output/北塩原村_サービス見込量算定.xlsx
"""

import os
import runpy
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from kitashiobara_common import (
    COLORS, FONT, JIDO_KYUFU, JIDO_TOTAL, KAIGO_KYUFU, KAIGO_TOTAL, OUT_DIR,
    POPULATION, TEGATA, TEGATA_YEARS, add_sheet, ensure_out_dir,
    style_header_row, style_note, style_title, write_row,
)

OUT_FILE = f"{OUT_DIR}/北塩原村_サービス見込量算定.xlsx"
KIJUNBI = "基準日：令和8年9月24日"

BASE_FY = 7          # 令和7年度＝基準年度
BASE_CY = 2025       # 令和7年4月1日（母数の基準）
PLAN_YEARS = [2027, 2028, 2029]      # 令和9〜11年度
PLAN_LABELS = ["令和9年度", "令和10年度", "令和11年度"]

# 置き方の語彙
R_BASE = "基準年度据え置き"
R_MIN1 = "実績あり・最低1人"
R_GOAL = "成果目標による"
R_PLAN = "現行計画据え置き"
R_ZERO = "実績なし・0"

RULE_FILL = {
    R_BASE: "2CA02C",
    R_MIN1: "2E75B6",
    R_GOAL: "7030A0",
    R_PLAN: "ED7D31",
    R_ZERO: "808080",
}


# ---------------------------------------------------------------------------
# 見込量の対象サービス
# 表示名, 実績キー（KAIGO_KYUFU/JIDO_KYUFU。合算するものは複数）, 区分,
# 月請求か, 置き方, 見込（人／月）, 年間請求件数の置き方, 根拠・備考
#
# 「年間請求件数の置き方」
#   None        基準年度の請求件数を据え置く
#   ("実績", y) 指定した令和y年度の請求件数・単価を用いる（直近に実績がある場合）
#   ("計上外",) 単価が得られないため給付費に計上しない
# ---------------------------------------------------------------------------
ADULT_PLAN = [
    # --- 訪問系 ---
    ("居宅介護", ("居宅介護",), "訪問系", True, R_BASE, 4, None,
     "令和7年度44件（月平均3.7）。令和6年度に27件から43件へ増えた水準が定着している"),
    ("重度訪問介護", (), "訪問系", True, R_ZERO, 0, None,
     "令和2〜7年度に実績なし。重度の肢体不自由者等。対象者の有無を村に確認する"),
    ("同行援護", (), "訪問系", True, R_ZERO, 0, None,
     "実績なし。視覚障がい者の外出支援。介護保険に相当するサービスがない"),
    ("行動援護", (), "訪問系", True, R_ZERO, 0, None,
     "実績なし。強度行動障害等。介護保険に相当するサービスがない"),
    ("重度障がい者等包括支援", (), "訪問系", True, R_ZERO, 0, None,
     "実績なし。最重度者向け。圏域の供給を確認する"),

    # --- 日中活動系 ---
    ("生活介護", ("生活介護",), "日中活動系", True, R_BASE, 7, None,
     "令和7年度79件（月平均6.6）。令和4年度以降83・84・84・79件で頭打ち。"
     "ほかに基準該当生活介護7件55,760円があり、給付費に加算して計上している",
     ("生活介護（基準該当）",)),
    ("自立訓練（機能訓練）", (), "日中活動系", True, R_ZERO, 0, None,
     "実績なし。リハビリ需要は介護保険側に流れている可能性がある"),
    ("自立訓練（生活訓練）", ("自立訓練（生活訓練）",), "日中活動系", True, R_MIN1, 1, None,
     "令和7年度2件。令和2年度23件から減少したが実績が続いている。0とはしない"),
    ("就労選択支援", (), "日中活動系", True, R_GOAL, 1, ("計上外",),
     "令和7年10月開始。成果目標は年1〜2人。単価の実績がないため給付費は計上しない"),
    ("就労移行支援", ("就労移行支援",), "日中活動系", True, R_GOAL, 1, ("実績", 5),
     "令和7年度は実績なし。成果目標（３）の一般就労移行者数1人に対応して計上する。"
     "単価は直近の実績年度（令和5年度5件744,891円）による"),
    ("就労継続支援Ａ型", (), "日中活動系", True, R_PLAN, 1, ("計上外",),
     "令和2〜7年度の給付実績にＡ型の計上がない。現行計画・計画素案は1人としており"
     "食い違う。村に確認するまで見込量は現行計画を据え置き、給付費は計上しない"),
    ("就労継続支援Ｂ型", ("就労継続支援（B型）",), "日中活動系", True, R_BASE, 12, None,
     "令和7年度143件（月平均11.9）。令和6年度140件から＋2.1％で頭打ち"),
    ("就労定着支援", ("就労定着支援",), "日中活動系", True, R_GOAL, 1, ("実績", 3),
     "令和7年度は実績なし。成果目標（３）の就労定着支援利用者数1人に対応して計上する。"
     "単価は直近の実績年度（令和3年度2件53,230円）による"),
    ("療養介護", (), "日中活動系", True, R_ZERO, 0, None,
     "実績なし。医療機関での長期療養。対象者の有無を確認する"),
    ("短期入所（福祉型）", ("短期入所",), "日中活動系", True, R_MIN1, 1, None,
     "令和7年度3件（令和6年度2件）。現行計画は非計上だが実績があるため計上する"),
    ("短期入所（医療型）", (), "日中活動系", True, R_ZERO, 0, None,
     "実績なし。医療的ケア者の利用が生じた場合に備え、圏域の供給を確認する"),

    # --- 居住系 ---
    ("自立生活援助", (), "居住系", True, R_ZERO, 0, None,
     "実績なし。一人暮らしへの移行支援。地域移行と連動する"),
    ("共同生活援助", ("共同生活援助",),
     "居住系", True, R_BASE, 9, None,
     "令和7年度105件（月平均8.75）。介護保険に相当するサービスがないため"
     "65歳到達後も継続する。特定障害者特別給付費（補足給付）105件1,026,928円は"
     "サービスの利用回数ではないため件数には含めず、給付費にのみ加算している",
     ("共同生活援助（特定障害者特別給付費）",)),
    ("施設入所支援", ("施設入所支援",),
     "居住系", True, R_GOAL, 4, None,
     "令和7年度48件（月平均4.0）。令和4年度から4年間48件で一定。"
     "成果目標（１）の地域移行者数0人・削減0人と整合する。"
     "特定障害者特別給付費12件206,084円は給付費にのみ加算している",
     ("施設入所支援（特定障害者特別給付費）",)),

    # --- 相談支援 ---
    ("計画相談支援", ("計画相談支援",), "相談支援", False, R_PLAN, 22, None,
     "令和7年度76件。請求は計画作成とモニタリングのときだけ発生するため"
     "件数÷12は実利用者数と一致しない。セルフプラン率0％であり、"
     "実利用者数は現行計画の22人を据え置く。給付費は年間請求件数76件から算定する"),
    ("地域移行支援", (), "相談支援", True, R_ZERO, 0, None,
     "実績なし。成果目標（１）（２）と連動する"),
    ("地域定着支援", (), "相談支援", True, R_ZERO, 0, None,
     "実績なし。単身生活者の緊急時対応。地域生活支援拠点と連動する"),
]

CHILD_PLAN = [
    ("児童発達支援", ("児童発達支援",), "障害児通所支援", True, R_MIN1, 1, None,
     "令和7年度1件。令和5年度23件・令和6年度13件からの急減であり、"
     "就学による放課後等デイサービスへの移行と見られる。"
     "新たに利用が生じた場合の影響が大きい（08シートの感度）"),
    ("放課後等デイサービス", ("放課後等デイサービス",), "障害児通所支援", True, R_BASE, 2, None,
     "令和7年度27件（月平均2.25）。令和6年度12件からの急増であり、"
     "就学に伴う児童発達支援からの流入と見られる"),
    ("保育所等訪問支援", ("保育所等訪問支援",), "障害児通所支援", True, R_MIN1, 1, None,
     "令和7年度に初めて実績（3件）。0とはしない"),
    ("居宅訪問型児童発達支援", (), "障害児通所支援", True, R_ZERO, 0, None,
     "実績なし。重症心身障害児等が対象"),
    ("障がい児相談支援", ("障害児相談支援",), "障害児相談支援", False, R_PLAN, 2, None,
     "令和7年度8件。計画相談支援と同じく請求が毎月発生しない。"
     "実利用者数は現行計画の2人を据え置く"),
]

# 障害児入所支援は実施主体が都道府県であり、市町村障害児福祉計画では見込量を定めない。


# ---------------------------------------------------------------------------
# 据え置き一覧
# 番号, 区分, 据え置いた内容, 据え置いた理由, 解除の条件, 影響の向き
# ---------------------------------------------------------------------------
HOLDS = [
    ("H-1", "数量", "全サービスの月平均利用者数を令和7年度の水準で据え置いた",
     "母数（手帳所持者数）は減り、利用実績は増えてきた。両者は相殺の方向にあり、"
     "どちらか一方の率を乗じる根拠がない",
     "村の匿名利用者一覧（年齢・サービス・区分）の受領。"
     "18歳到達者・65歳到達者が分かれば個別に積み上げられる", "中立"),

    ("H-2", "数量", "計画相談支援の利用者数を現行計画の22人／月で据え置いた",
     "請求が計画作成とモニタリングのときだけ発生するため、"
     "令和7年度76件から実利用者数を割り出せない",
     "村の支給決定者数（サービス別・実人数）の受領", "不明"),

    ("H-3", "数量", "障がい児相談支援の利用者数を現行計画の2人／月で据え置いた",
     "H-2と同じ", "同上", "不明"),

    ("H-4", "数量", "就労継続支援Ａ型を現行計画の1人／月で据え置いた",
     "令和2〜7年度の給付実績にＡ型の計上がない。実在するかが確認できていない",
     "村への確認（M-4）。実在しなければ0、実在すれば単価とともに計上する", "低く出る"),

    ("H-5", "単価", "1件あたり給付費を令和7年度の水準で据え置いた",
     "令和9年度の報酬改定率が未告示である",
     "報酬改定率の告示。令和6年度改定は全体＋1.12％だった",
     "低く出る"),

    ("H-6", "数量", "児童発達支援を1人／月・年1件で据え置いた",
     "令和5年度23件から令和7年度1件への急減の理由が確認できていない",
     "村への確認（M-1）。就学による移行であれば妥当、"
     "転出や利用中止であれば新規の発生を見込む必要がある", "低く出る"),

    ("H-7", "数量", "実績のない10サービスを0で据え置いた",
     "令和2〜7年度に給付実績がない。実績ゼロをニーズゼロと判断しないため、"
     "利用が生じた場合の対応方針を計画本文に書く",
     "アンケート調査の利用意向、相談支援事業所からの相談件数", "低く出る"),

    ("H-8", "数量", "重度障がい者（強度行動障害・高次脳機能障害・医療的ケア者）の"
     "内数を0で据え置いた",
     "該当者数が確認できていない。基本指針 別表第一 三・四は"
     "生活介護・短期入所・共同生活援助について内数の設定を求めている",
     "村への確認（M-5・M-6）", "不明"),

    ("H-9", "給付費", "地域生活支援事業の費用を見込量から算定していない",
     "地域生活支援事業は統合補助金であり、給付費とは算定の系統が異なる",
     "村の地域生活支援事業の決算額（令和6・7年度）の受領",
     "対象外"),

    ("H-10", "給付費", "就労選択支援・就労継続支援Ａ型の給付費を計上していない",
     "単価の実績が得られない",
     "就労選択支援は令和8年度の実績、Ａ型は村への確認（M-4）", "低く出る"),
]


# ---------------------------------------------------------------------------
# 村への確認事項（見込量算定関係）
# ---------------------------------------------------------------------------
MURA_CONFIRM = [
    ("M-1", "児童発達支援が令和5年度23件から令和7年度1件へ減った理由",
     "就学による放課後等デイサービスへの移行か、転出・利用中止か",
     "H-6。理由により令和9〜11年度の置き方が変わる", "最優先"),
    ("M-2", "支給決定者数（サービス別・実人数）",
     "請求件数ではなく実人数。計画相談支援・障がい児相談支援の見込量の基礎になる",
     "H-2・H-3", "最優先"),
    ("M-3", "令和9〜11年度の18歳到達者・65歳到達者",
     "到達年度と現在利用しているサービス。個別積上げの流入・流出になる",
     "H-1", "高"),
    ("M-4", "就労継続支援Ａ型の利用者の有無",
     "現行計画・計画素案は1人としているが給付実績に計上がない",
     "H-4・H-10", "高"),
    ("M-5", "強度行動障害の状態にある方・高次脳機能障害のある方の人数",
     "生活介護・短期入所・共同生活援助の内数として個別の見込みを設定する",
     "H-8（別表第一 三・四）", "高"),
    ("M-6", "医療的ケア児1人の年齢と、医療的ケア者（成人）の有無",
     "令和8年8月19日の基本指針改正で医療的ケア者が対象に加わった",
     "H-8。18歳到達後は医療的ケア者として内数に計上する", "高"),
    ("M-7", "地域生活支援事業の決算額（令和6・7年度）",
     "事業別の決算額。補助基準額超過分の村単独負担が分かる",
     "H-9", "中"),
    ("M-8", "重度訪問介護・同行援護・行動援護・療養介護の対象者の有無",
     "実績はないが対象者がいる可能性がある",
     "H-7", "中"),
]


# ============================================================
# 算定
# ============================================================
def _proj():
    """将来推計ブックと同じ計算で、手帳所持者数（中位）と年少人口を求める。"""
    mod = runpy.run_path(
        os.path.join(os.path.dirname(os.path.abspath(__file__)),
                     "build_kitashiobara_projection.py"))
    fl, rh, ip = mod["_forecast_linear"], mod["_round_half_up"], mod["_interpolated_population"]

    base_pop = dict((r[1], r[3]) for r in POPULATION)[2023]
    years = [BASE_CY] + PLAN_YEARS
    tegata = {y: 0 for y in years}
    detail = {}
    for name, actual in TEGATA.items():
        xs = [s for (_, s), v in zip(TEGATA_YEARS, actual) if v is not None]
        ys = [v for v in actual if v is not None]
        rate = ys[xs.index(2023)] / base_pop
        row = {}
        for y in years:
            a = rh(fl(y, ys, xs))
            b = rh(rate * ip(y))
            m = rh((a + b) / 2)
            row[y] = m
            tegata[y] += m
        detail[name] = row

    nen = dict((r[1], r[4]) for r in POPULATION)
    r6, r11 = nen[2024], nen[2029]
    nensho = {y: (nen[y] if y in (2024, 2029) else rh(r6 + (r11 - r6) * (y - 2024) / 5))
              for y in years}
    return detail, tegata, nensho


def base_record(keys, src):
    """基準年度（令和7年度）の請求件数と給付費を、合算キーの分まで足して返す。"""
    ken = gaku = 0
    for k in keys:
        rec = src.get(k, {}).get(BASE_FY)
        if rec:
            ken += rec[0]
            gaku += rec[1]
    return ken, gaku


def unit_price(keys, src, fy):
    """指定年度の1件あたり給付費。"""
    ken = gaku = 0
    for k in keys:
        rec = src.get(k, {}).get(fy)
        if rec:
            ken += rec[0]
            gaku += rec[1]
    return (gaku / ken if ken else 0.0), ken, gaku


def build_rows(plan, src):
    """サービスごとに、基準年度の実績・見込量・年間請求件数・給付費を組み立てる。

    加算キー（特定障害者特別給付費・基準該当）は、サービスの利用回数ではないため
    請求件数には含めず、給付費にのみ足す。
    """
    out = []
    for rec in plan:
        name, keys, kubun, monthly, rule, users, kensu_rule, note = rec[:8]
        addon_keys = rec[8] if len(rec) > 8 else ()

        b_ken, b_gaku = base_record(keys, src)
        a_ken, a_gaku = base_record(addon_keys, src)
        tanka = (b_gaku / b_ken) if b_ken else 0.0

        if kensu_rule is None:
            ken = b_ken
            keijou = True
        elif kensu_rule[0] == "実績":
            tanka, _k, _g = unit_price(keys, src, kensu_rule[1])
            ken = users * 12
            keijou = True
        else:                       # 計上外
            ken = users * 12
            tanka = 0.0
            keijou = False

        gaku = (round(ken * tanka) + a_gaku) if keijou else 0
        out.append({
            "name": name, "kubun": kubun, "monthly": monthly, "rule": rule,
            "base_ken": b_ken, "base_gaku": b_gaku,
            "addon_ken": a_ken, "addon_gaku": a_gaku,
            "users": users, "ken": ken, "tanka": tanka,
            "gaku": gaku, "keijou": keijou, "note": note,
        })
    return out


def patterns(tegata, nensho):
    """3つの算定パターンの3か年給付費を求める（比較用）。"""
    a7 = KAIGO_TOTAL[BASE_FY][1]
    c7 = JIDO_TOTAL[BASE_FY][1]
    ca = (KAIGO_TOTAL[BASE_FY][1] / KAIGO_TOTAL[2][1]) ** 0.2
    cc = (JIDO_TOTAL[BASE_FY][1] / JIDO_TOTAL[2][1]) ** 0.2

    p1, p2 = [], []
    for i, y in enumerate(PLAN_YEARS):
        p1.append(round(a7 * tegata[y] / tegata[BASE_CY])
                  + round(c7 * nensho[y] / nensho[BASE_CY]))
        p2.append(round(a7 * ca ** (i + 2)) + round(c7 * cc ** (i + 2)))
    return p1, p2, ca, cc


# ============================================================
# シート
# ============================================================
def _rule(ws, row, col):
    c = ws.cell(row=row, column=col)
    fill = RULE_FILL.get(str(c.value))
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")


def _yen(ws, row, cols):
    for col in cols:
        ws.cell(row=row, column=col).number_format = "#,##0"


def sheet_about(wb, adult, child, p1, p2, p3):
    ws = add_sheet(
        wb, "00_この算定について",
        "サービス見込量の算定（第1次概算）",
        f"{KIJUNBI}。受託者の分析であり確定値ではありません。"
        "大雪地区広域連合・川崎町で用いた手順を障がい福祉計画に移したものです。",
        [10, 46, 62, 18])
    r = 5
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    style_title(ws.cell(row=r, column=1), "1　結論", fill=COLORS["subhead"], size=11)
    r += 1
    tot3 = sum(p3)
    lines = [
        f"基準年度は令和7年度です。令和8年度は年度途中で給付実績が完結していないため使えません。",
        f"採るのはパターン3（基準年度据え置き＋成果目標による個別補正）です。"
        f"母数（手帳所持者数）は年約2％減りますが、介護給付の請求件数は5年で1.66倍に増えています。"
        f"両者は相殺の方向にあり、どちらか一方の率を乗じる根拠がありません。",
        f"第8期（令和9〜11年度）の給付費は3か年計 {tot3:,}円です"
        f"（令和9年度 {p3[0]:,}／令和10年度 {p3[1]:,}／令和11年度 {p3[2]:,}円）。"
        f"地域生活支援事業と村単独事業は含みません。",
        f"パターン1（母数連動）は3か年 {sum(p1):,}円、"
        f"パターン2（実績トレンド延長）は {sum(p2):,}円です。"
        f"採用値はこの幅の中にあり、パターン1の{sum(p3)/sum(p1):.2f}倍、"
        f"パターン2の{sum(p3)/sum(p2):.2f}倍です。",
        f"据え置いた事項は{len(HOLDS)}件（07シート）、"
        f"村への確認事項は{len(MURA_CONFIRM)}件（10シート）です。"
        f"このうち給付費を低く出す向きに働くものが"
        f"{len([h for h in HOLDS if h[5] == '低く出る'])}件あります。"
        f"単価を令和7年度で固定しており、令和9年度の報酬改定を見込んでいないことが最大のものです。",
        f"月平均利用者数と年間請求件数を人／月×12で結んではいけません。"
        f"短期入所は年3件であり1人／月と計上しても年12件にはならず、"
        f"計画相談支援の76件は月平均6.3人ではなく実利用者22人に対応します。"
        f"給付費は年間請求件数から算定しています。",
    ]
    for i, t in enumerate(lines):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        style_note(ws.cell(row=r, column=1), f"{i + 1}. {t}")
        ws.row_dimensions[r].height = 46
        r += 1

    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    style_title(ws.cell(row=r, column=1), "2　置き方の内訳", fill=COLORS["subhead"], size=11)
    r += 1
    style_header_row(ws, r, ["置き方", "件数", "意味", ""])
    r += 1
    rows = [
        (R_BASE, "令和7年度の実績をそのまま置いた"),
        (R_MIN1, "実績はあるが月平均1人に満たないため1人とした。"
                 "実績があるのに0と置くと、指定の制限（別表第五）の場面で不利になる"),
        (R_GOAL, "第4章の成果目標で人数が定まっているため、それに合わせた"),
        (R_PLAN, "請求が毎月発生しない、または実績が確認できないため現行計画を据え置いた"),
        (R_ZERO, "令和2〜7年度に実績がない。実績ゼロをニーズゼロとは判断しない"),
    ]
    allrows = adult + child
    for i, (k, mean) in enumerate(rows):
        row0 = r
        n = len([x for x in allrows if x["rule"] == k])
        r = write_row(ws, r, [k, n, mean, ""], alt=(i % 2 == 1),
                      aligns=["center", "center", "left", "left"])
        _rule(ws, row0, 1)
        ws.row_dimensions[row0].height = 32
    return ws


def sheet_base_year(wb):
    ws = add_sheet(
        wb, "01_基準年度の選び方",
        "基準年度に令和7年度を採る理由",
        "介護保険事業計画で確立した手順の第1段階です。完結していない年度は基準年度に使えません。",
        [14, 16, 60, 60])
    r = 5
    style_header_row(ws, r, ["年度", "可否", "理由", "本算定での扱い"])
    r += 1
    rows = [
        ("令和8年度", "使えない",
         "年度途中であり給付実績が完結していない。本算定の時点（令和8年9月）で"
         "受領しているのは令和7年度までである",
         "用いない"),
        ("令和7年度", "使える",
         "完結年度であり、12か月分の請求件数・給付費が揃っている。"
         "介護給付624件53,877,076円、障害児39件1,635,703円",
         "数量・単価とも基準年度とする"),
        ("令和6年度", "使えるが採らない",
         "完結しているが、居宅介護が27件から43件へ、児童発達支援が23件から13件へ"
         "動いた年であり、令和7年度のほうが直近の姿に近い",
         "比較のために併記する"),
        ("令和2〜5年度", "使えるが採らない",
         "完結しているが古い。就労選択支援の創設（令和7年10月）、"
         "医療型児童発達支援の一元化（令和6年度）など制度が変わっている",
         "伸びの検証にのみ用いる"),
    ]
    for i, row in enumerate(rows):
        row0 = r
        r = write_row(ws, r, list(row), alt=(i % 2 == 1),
                      aligns=["center", "center", "left", "left"])
        c = ws.cell(row=row0, column=2)
        c.fill = PatternFill("solid", fgColor={"使える": "2CA02C", "使えない": "C00000"}
                             .get(str(c.value), "ED7D31"))
        c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        ws.row_dimensions[row0].height = 46

    r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    style_title(ws.cell(row=r, column=1), "参考　年度別の請求件数と給付費",
                fill=COLORS["subhead"], size=11)
    r += 1
    style_header_row(ws, r, ["年度", "介護給付 件数", "介護給付 給付費", "障害児 件数／給付費"])
    r += 1
    for i, fy in enumerate(range(2, 8)):
        row0 = r
        ak, ag = KAIGO_TOTAL[fy]
        ck, cg = JIDO_TOTAL[fy]
        r = write_row(ws, r, [f"令和{fy}年度", ak, ag, f"{ck}件／{cg:,}円"],
                      alt=(i % 2 == 1),
                      aligns=["center", "right", "right", "right"])
        _yen(ws, row0, [2, 3])
    return ws


def sheet_patterns(wb, tegata, nensho, p1, p2, p3, ca, cc):
    ws = add_sheet(
        wb, "02_算定パターンの比較",
        "算定パターンの列挙と、採るパターンの決定",
        "3つのパターンを並べ、採否の理由を示します。採用値はパターン1とパターン2の幅の中にあります。",
        [22, 20, 20, 20, 20, 64])
    r = 5
    style_header_row(ws, r, ["パターン", "令和9年度", "令和10年度", "令和11年度",
                             "3か年計", "採否の理由"])
    r += 1
    recs = [
        ("1　母数連動", p1,
         f"手帳所持者数（中位推計）の伸びを乗じる。令和7年{tegata[BASE_CY]}人から"
         f"令和11年{tegata[2029]}人へ年約2％減。"
         "採らない。母数は減るが利用実績は5年で1.66倍に増えており、"
         "母数の率をそのまま乗じると見込量が実態より低く出る"),
        ("2　実績トレンド延長", p2,
         f"令和2〜7年度の給付費の年平均伸び率（介護給付{ca - 1:+.1%}／障害児{cc - 1:+.1%}）を延長。"
         "採らない。就労継続支援Ｂ型は5年で1.93倍だが直近は＋2.1％で頭打ちであり、"
         "同じ伸びを3年続けると令和11年度に22人となって18歳到達者数（年0〜1人）と合わない"),
        ("3　基準年度据え置き＋\n　　成果目標による補正", p3,
         "令和7年度の水準を維持し、成果目標で人数が定まるサービスと"
         "実績があるのに0となるサービスだけを個別に動かす。"
         "採用する。小規模自治体では1人の増減が率の効果を上回るため、"
         "率で動かすより実績を置いて個別に補正するほうが説明できる"),
    ]
    for i, (nm, vals, why) in enumerate(recs):
        row0 = r
        r = write_row(ws, r, [nm] + list(vals) + [sum(vals), why], alt=(i % 2 == 1),
                      aligns=["left", "right", "right", "right", "right", "left"])
        _yen(ws, row0, [2, 3, 4, 5])
        ws.row_dimensions[row0].height = 62
        if nm.startswith("3"):
            for col in range(1, 7):
                c = ws.cell(row=row0, column=col)
                c.font = Font(name=FONT, size=10, bold=True, color="1F3864")

    r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    style_title(ws.cell(row=r, column=1), "参考　母数の推移（パターン1の基礎）",
                fill=COLORS["subhead"], size=11)
    r += 1
    style_header_row(ws, r, ["区分", "令和7年\n（基準）", "令和9年", "令和10年", "令和11年",
                             "令和7年＝1とした伸び"])
    r += 1
    seq = [("障がい者手帳所持者数（中位推計）", tegata), ("年少人口（0〜14歳）", nensho)]
    for i, (nm, d) in enumerate(seq):
        r = write_row(ws, r, [nm, d[BASE_CY]] + [d[y] for y in PLAN_YEARS]
                      + ["　".join(f"{d[y] / d[BASE_CY]:.3f}" for y in PLAN_YEARS)],
                      alt=(i % 2 == 1),
                      aligns=["left", "right", "right", "right", "right", "center"])
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    style_note(ws.cell(row=r, column=1),
               "年少人口は令和7年166人から令和11年111人へ33％減りますが、"
               "児童発達支援が令和5年度23件から令和7年度1件へ、"
               "放課後等デイサービスが令和6年度12件から令和7年度27件へ動いた一方で、"
               "障害児給付の合計は令和5年度40件・令和7年度39件でほぼ変わっていません。"
               "就学に伴う個人単位の移動であり、年少人口の率とは無関係です。")
    ws.row_dimensions[r].height = 60
    return ws


def sheet_actual(wb, adult, child):
    ws = add_sheet(
        wb, "03_基準年度の実績",
        "令和7年度（基準年度）の請求件数・給付費・1件あたり給付費",
        "出所は村から受領した「障がいサービス給付実績」です。単価はこの表の値を据え置きます。",
        [28, 14, 12, 16, 15, 11, 18, 56])
    r = 5
    style_header_row(ws, r, ["サービス", "区分", "請求件数", "給付費（円）",
                             "1件あたり\n給付費（円）", "件数÷12",
                             "加算（補足給付等）", "備考"])
    r += 1
    for i, rec in enumerate(adult + child):
        row0 = r
        bk, bg = rec["base_ken"], rec["base_gaku"]
        r = write_row(ws, r,
                      [rec["name"], rec["kubun"], bk if bk else "―",
                       bg if bg else "―",
                       round(bg / bk) if bk else "―",
                       f"{bk / 12:.2f}" if bk else "―",
                       f"{rec['addon_ken']}件 {rec['addon_gaku']:,}円"
                       if rec["addon_ken"] else "―",
                       rec["note"]],
                      alt=(i % 2 == 1),
                      aligns=["left", "center", "right", "right", "right", "right",
                              "right", "left"])
        if bk:
            _yen(ws, row0, [3, 4, 5])
        ws.row_dimensions[row0].height = 32

    r += 1
    ak, ag = KAIGO_TOTAL[BASE_FY]
    ck, cg = JIDO_TOTAL[BASE_FY]
    row0 = r
    r = write_row(ws, r, ["合計（受領データの総計）", "―", ak + ck, ag + cg, "―", "―", "―",
                          f"介護給付 {ak}件 {ag:,}円／障害児 {ck}件 {cg:,}円。"
                          "加算（補足給付・基準該当）を含む"],
                  aligns=["left", "center", "right", "right", "right", "right",
                          "right", "left"])
    _yen(ws, row0, [3, 4])
    for col in range(1, 9):
        ws.cell(row=row0, column=col).font = Font(name=FONT, size=10, bold=True)
    return ws


def _mikomi_sheet(wb, title, heading, note, rows):
    ws = add_sheet(wb, title, heading, note,
                   [28, 14, 12, 20, 13, 13, 13, 14, 18, 52])
    r = 5
    style_header_row(ws, r, ["サービス", "区分", "単位", "置き方",
                             "令和9年度", "令和10年度", "令和11年度",
                             "年間請求件数", "3か年給付費（円）", "根拠・備考"])
    r += 1
    for i, rec in enumerate(rows):
        row0 = r
        unit = "人／月"
        vals = [rec["users"]] * 3
        r = write_row(ws, r,
                      [rec["name"], rec["kubun"], unit, rec["rule"]] + vals
                      + [rec["ken"] if rec["ken"] else "―",
                         rec["gaku"] * 3 if rec["keijou"] else "計上外",
                         rec["note"]],
                      alt=(i % 2 == 1),
                      aligns=["left", "center", "center", "center",
                              "right", "right", "right", "right", "right", "left"])
        _rule(ws, row0, 4)
        if rec["keijou"]:
            _yen(ws, row0, [9])
        ws.row_dimensions[row0].height = 34

    row0 = r
    tot = sum(x["gaku"] for x in rows if x["keijou"]) * 3
    r = write_row(ws, r, ["合計", "―", "―", "―", "", "", "",
                          sum(x["ken"] for x in rows), tot, ""],
                  aligns=["left", "center", "center", "center",
                          "right", "right", "right", "right", "right", "left"])
    _yen(ws, row0, [9])
    for col in range(1, 11):
        ws.cell(row=row0, column=col).font = Font(name=FONT, size=10, bold=True)
    return ws


def sheet_kyufu(wb, adult, child, p3):
    ws = add_sheet(
        wb, "06_給付費の見込み",
        "給付費の見込み（年度別・区分別）",
        "給付費は年間請求件数×1件あたり給付費で算定しています。"
        "月平均利用者数×12では算定していません（短期入所・計画相談支援で成り立たないため）。",
        [30, 20, 20, 20, 20, 50])
    r = 5
    style_header_row(ws, r, ["区分", "令和9年度", "令和10年度", "令和11年度",
                             "3か年計", "備考"])
    r += 1
    a_sum = sum(x["gaku"] for x in adult if x["keijou"])
    c_sum = sum(x["gaku"] for x in child if x["keijou"])
    recs = [
        ("障害福祉サービス等（自立支援給付）", a_sum,
         "居宅介護・生活介護・就労系・居住系・相談支援。"
         "就労選択支援と就労継続支援Ａ型は単価が得られないため計上していない"),
        ("障害児通所支援等", c_sum,
         "児童発達支援・放課後等デイサービス・保育所等訪問支援・障がい児相談支援。"
         "障害児入所支援は都道府県が実施主体のため計上しない"),
    ]
    for i, (nm, v, note) in enumerate(recs):
        row0 = r
        r = write_row(ws, r, [nm, v, v, v, v * 3, note], alt=(i % 2 == 1),
                      aligns=["left", "right", "right", "right", "right", "left"])
        _yen(ws, row0, [2, 3, 4, 5])
        ws.row_dimensions[row0].height = 34
    row0 = r
    r = write_row(ws, r, ["合計"] + p3 + [sum(p3), ""],
                  aligns=["left", "right", "right", "right", "right", "left"])
    _yen(ws, row0, [2, 3, 4, 5])
    for col in range(1, 7):
        ws.cell(row=row0, column=col).font = Font(name=FONT, size=10, bold=True)
    r += 2

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    style_title(ws.cell(row=r, column=1), "法定負担割合による財源の内訳（3か年計）",
                fill=COLORS["subhead"], size=11)
    r += 1
    style_header_row(ws, r, ["負担区分", "割合", "金額（円）", "", "", "根拠"])
    r += 1
    total = sum(p3)
    shares = [("国", 0.50, "障害者総合支援法第95条／児童福祉法第57条の2"),
              ("県", 0.25, "同上"),
              ("村", 0.25, "同上")]
    for i, (nm, sh, kon) in enumerate(shares):
        row0 = r
        r = write_row(ws, r, [nm, f"{sh:.0%}", round(total * sh), "", "", kon],
                      alt=(i % 2 == 1),
                      aligns=["center", "center", "right", "left", "left", "left"])
        _yen(ws, row0, [3])
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    style_note(ws.cell(row=r, column=1),
               "国庫負担基準額を超える分は村の負担になります。"
               "本村は共同生活援助・施設入所支援の利用者の割合が高く、"
               "国庫負担基準額との関係を村の決算で確認する必要があります（確認事項 M-7）。"
               "地域生活支援事業（国1/2以内・県1/4以内の統合補助金）と村単独事業は"
               "この表に含みません。")
    ws.row_dimensions[r].height = 46
    return ws


def sheet_holds(wb):
    ws = add_sheet(
        wb, "07_据え置き一覧",
        f"据え置いた事項（{len(HOLDS)}件）と解除の条件",
        "資料が届かない場合はこのまま確定できます。"
        "「影響の向き」は、据え置いたことで給付費がどちらに偏るかを示します。",
        [8, 12, 44, 48, 44, 14])
    r = 5
    style_header_row(ws, r, ["番号", "区分", "据え置いた内容", "据え置いた理由",
                             "解除の条件", "影響の向き"])
    r += 1
    fill = {"低く出る": "C00000", "高く出る": "ED7D31", "中立": "2CA02C",
            "不明": "808080", "対象外": "808080"}
    for i, rec in enumerate(HOLDS):
        row0 = r
        r = write_row(ws, r, list(rec), alt=(i % 2 == 1),
                      aligns=["center", "center", "left", "left", "left", "center"])
        c = ws.cell(row=row0, column=6)
        if str(c.value) in fill:
            c.fill = PatternFill("solid", fgColor=fill[str(c.value)])
            c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        ws.row_dimensions[row0].height = 46
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    style_note(ws.cell(row=r, column=1),
               "「低く出る」が4件あります。最大のものは単価の据え置き（H-5）です。"
               "令和6年度の報酬改定は全体＋1.12％でした。同程度の改定が令和9年度にあれば"
               "3か年で約180万円の増になります。"
               "資料が届かないことにより、給付費は低く出る側に偏っています。")
    ws.row_dimensions[r].height = 46
    return ws


def sheet_check(wb, adult, child, checks):
    ws = add_sheet(
        wb, "08_自己点検",
        "自己点検",
        "すべて「合」でなければこのブックは生成されません（終了コード1で落とします）。",
        [8, 50, 26, 26, 12, 50])
    r = 5
    style_header_row(ws, r, ["番号", "点検の内容", "期待値", "計算値", "結果", "備考"])
    r += 1
    for i, rec in enumerate(checks):
        row0 = r
        r = write_row(ws, r, list(rec), alt=(i % 2 == 1),
                      aligns=["center", "left", "right", "right", "center", "left"])
        c = ws.cell(row=row0, column=5)
        c.fill = PatternFill("solid", fgColor="2CA02C" if c.value == "合" else "C00000")
        c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        ws.row_dimensions[row0].height = 34

    r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    style_title(ws.cell(row=r, column=1), "感度　児童発達支援が再び利用されたとき",
                fill=COLORS["subhead"], size=11)
    r += 1
    d = JIDO_KYUFU["児童発達支援"]
    style_header_row(ws, r, ["年度", "請求件数", "給付費（円）", "", "", "読み方"])
    r += 1
    for i, fy in enumerate((5, 6, 7)):
        row0 = r
        k, g = d.get(fy, (0, 0))
        r = write_row(ws, r, [f"令和{fy}年度", k, g, "", "",
                              "本算定は令和7年度（1件124,060円）を据え置いている。"
                              "令和5年度の水準に戻れば1年あたり約321万円の増になる"
                              if fy == 7 else ""],
                      alt=(i % 2 == 1),
                      aligns=["center", "right", "right", "left", "left", "left"])
        _yen(ws, row0, [2, 3])
    return ws


def sheet_judo(wb):
    ws = add_sheet(
        wb, "09_重度障がい者の内数",
        "重度障がい者について個別に設定する利用者数の見込み（別表第一 三・四）",
        "令和8年8月19日の基本指針改正により「医療的ケアを必要とする者等」が"
        "「医療的ケア者等」に改められました。内数の設定は努力義務です。",
        [26, 22, 16, 16, 16, 60])
    r = 5
    style_header_row(ws, r, ["サービス", "内数の区分", "令和9年度", "令和10年度",
                             "令和11年度", "備考"])
    r += 1
    services = ["生活介護", "短期入所（福祉型）", "短期入所（医療型）", "共同生活援助"]
    kubun = [
        ("強度行動障害の状態にある者", "村に該当者数を確認する（M-5）"),
        ("高次脳機能障害を有する者", "村に該当者数を確認する（M-5）"),
        ("医療的ケア者等", "医療的ケア児1人が18歳に達した後は医療的ケア者として"
                          "内数に計上する。年齢を村に確認する（M-6）"),
    ]
    i = 0
    for sv in services:
        for kb, note in kubun:
            row0 = r
            r = write_row(ws, r, [sv if kb == kubun[0][0] else "", kb,
                                  "【要確認】", "【要確認】", "【要確認】", note],
                          alt=(i % 2 == 1),
                          aligns=["left", "left", "center", "center", "center", "left"])
            for col in (3, 4, 5):
                c = ws.cell(row=row0, column=col)
                c.fill = PatternFill("solid", fgColor=COLORS["input"])
            ws.row_dimensions[row0].height = 30
            i += 1
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    style_note(ws.cell(row=r, column=1),
               "本村の規模では該当者は0〜1人と見込まれます。0であることも含めて"
               "計画本文に書くことで、努力義務を満たしたことが示せます。"
               "該当者がいる場合は、別表第五による指定の制限の例外として"
               "個別ニーズに対応する事業者の指定を県に求める根拠になります。")
    ws.row_dimensions[r].height = 46
    return ws


def sheet_confirm(wb):
    ws = add_sheet(
        wb, "10_村への確認事項",
        f"村への確認事項（{len(MURA_CONFIRM)}件）",
        "「最優先」の2件が揃わないと、見込量は第1次概算のままになります。",
        [8, 44, 50, 34, 12, 12])
    r = 5
    style_header_row(ws, r, ["番号", "確認事項", "確認したい内容", "対応する据え置き",
                             "優先度", "状態"])
    r += 1
    prio = {"最優先": "C00000", "高": "ED7D31", "中": "808080"}
    for i, rec in enumerate(MURA_CONFIRM):
        row0 = r
        r = write_row(ws, r, list(rec) + ["未受領"], alt=(i % 2 == 1),
                      aligns=["center", "left", "left", "left", "center", "center"])
        c = ws.cell(row=row0, column=5)
        c.fill = PatternFill("solid", fgColor=prio.get(str(c.value), "808080"))
        c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        ws.cell(row=row0, column=6).fill = PatternFill("solid", fgColor="FCE4E4")
        ws.row_dimensions[row0].height = 38
    return ws


# ============================================================
def run_checks(adult, child, p3):
    """自己点検。1件でも落ちたら終了コード1で止める。"""
    checks = []

    def add(no, what, exp, got, note=""):
        ok = (exp == got)
        checks.append((no, what, exp, got, "合" if ok else "否", note))
        return ok

    # 1 基準年度の請求件数が受領データの総計と一致すること
    a_base = sum(x["base_ken"] + x["addon_ken"] for x in adult)
    exp_a = KAIGO_TOTAL[BASE_FY][0]
    add("1", "障害福祉サービスの基準年度請求件数が受領データの総計と一致する",
        exp_a, a_base, "宿泊型自立訓練は令和7年度に実績がないため差は生じない")

    c_base = sum(x["base_ken"] + x["addon_ken"] for x in child)
    add("2", "障害児通所支援等の基準年度請求件数が受領データの総計と一致する",
        JIDO_TOTAL[BASE_FY][0], c_base)

    # 3 基準年度の給付費が一致すること
    add("3", "障害福祉サービスの基準年度給付費が受領データの総計と一致する",
        KAIGO_TOTAL[BASE_FY][1], sum(x["base_gaku"] + x["addon_gaku"] for x in adult))
    add("4", "障害児通所支援等の基準年度給付費が受領データの総計と一致する",
        JIDO_TOTAL[BASE_FY][1], sum(x["base_gaku"] + x["addon_gaku"] for x in child))

    # 5 年度別給付費が3年とも同額であること（据え置きのため）
    add("5", "据え置きのため3年度の給付費が同額である", 1, len(set(p3)),
        "パターン3は基準年度据え置きであり、年度間で動かない")

    # 6 成果目標との整合
    shisetsu = [x for x in adult if x["name"] == "施設入所支援"][0]
    add("6", "施設入所支援の見込量が成果目標（１）の4人と一致する", 4, shisetsu["users"],
        "地域移行者数0人・削減0人と整合する")
    ikou = [x for x in adult if x["name"] == "就労移行支援"][0]
    add("7", "就労移行支援の見込量が成果目標（３）の一般就労移行者数1人と一致する",
        1, ikou["users"])
    teichaku = [x for x in adult if x["name"] == "就労定着支援"][0]
    add("8", "就労定着支援の見込量が成果目標（３）の1人と一致する", 1, teichaku["users"])

    # 9 実績があるサービスの見込量が0でないこと
    zero_with_actual = [x["name"] for x in adult + child
                        if x["base_ken"] > 0 and x["users"] == 0]
    add("9", "基準年度に実績のあるサービスで見込量が0のものがない",
        0, len(zero_with_actual),
        "／".join(zero_with_actual) if zero_with_actual else
        "実績があるのに0と置くと指定の制限の場面で不利になる")

    # 10 給付費を人／月×12で算定していないこと（短期入所で確認）
    tanki = [x for x in adult if x["name"] == "短期入所（福祉型）"][0]
    add("10", "短期入所の年間請求件数が人／月×12になっていない",
        True, tanki["ken"] != tanki["users"] * 12,
        f"年間請求件数{tanki['ken']}件。1人／月×12＝12件ではない")

    # 11 計画相談支援も同じ
    keikaku = [x for x in adult if x["name"] == "計画相談支援"][0]
    add("11", "計画相談支援の年間請求件数が人／月×12になっていない",
        True, keikaku["ken"] != keikaku["users"] * 12,
        f"年間請求件数{keikaku['ken']}件。22人／月×12＝264件ではない")

    # 12 給付費の合計が年度別の値と整合すること
    add("12", "3か年給付費の合計が年度別の和と一致する", sum(p3),
        p3[0] + p3[1] + p3[2])

    ng = [c for c in checks if c[4] == "否"]
    if ng:
        for c in ng:
            print(f"自己点検 否: {c[0]} {c[1]} 期待{c[2]} 計算{c[3]}", file=sys.stderr)
        sys.exit(1)
    return checks


def main():
    ensure_out_dir()
    _detail, tegata, nensho = _proj()

    adult = build_rows(ADULT_PLAN, KAIGO_KYUFU)
    child = build_rows(CHILD_PLAN, JIDO_KYUFU)

    a_sum = sum(x["gaku"] for x in adult if x["keijou"])
    c_sum = sum(x["gaku"] for x in child if x["keijou"])
    p3 = [a_sum + c_sum] * 3
    p1, p2, ca, cc = patterns(tegata, nensho)

    checks = run_checks(adult, child, p3)

    wb = Workbook()
    wb.remove(wb.active)
    sheet_about(wb, adult, child, p1, p2, p3)
    sheet_base_year(wb)
    sheet_patterns(wb, tegata, nensho, p1, p2, p3, ca, cc)
    sheet_actual(wb, adult, child)
    _mikomi_sheet(wb, "04_見込量_障害福祉サービス",
                  "障害福祉サービス等の見込量（令和9〜11年度）",
                  "「置き方」の色は00シートの凡例のとおりです。"
                  "年間請求件数は給付費の算定基礎であり、人／月×12ではありません。",
                  adult)
    _mikomi_sheet(wb, "05_見込量_障害児通所支援等",
                  "障害児通所支援等の見込量（令和9〜11年度）",
                  "障害児入所支援は都道府県が実施主体であり、市町村障害児福祉計画では"
                  "見込量を定めません。",
                  child)
    sheet_kyufu(wb, adult, child, p3)
    sheet_holds(wb)
    sheet_check(wb, adult, child, checks)
    sheet_judo(wb)
    sheet_confirm(wb)
    wb.save(OUT_FILE)

    print(f"作成: {OUT_FILE}")
    print(f"  シート数: {len(wb.sheetnames)}")
    print(f"  基準年度: 令和{BASE_FY}年度"
          f"（介護給付{KAIGO_TOTAL[BASE_FY][0]}件{KAIGO_TOTAL[BASE_FY][1]:,}円／"
          f"障害児{JIDO_TOTAL[BASE_FY][0]}件{JIDO_TOTAL[BASE_FY][1]:,}円）")
    print(f"  採用パターン: 3 基準年度据え置き＋成果目標による補正")
    print(f"  3か年給付費: P1 {sum(p1):,}円／P3 {sum(p3):,}円（採用）／P2 {sum(p2):,}円")
    print(f"  サービス数: 障害福祉{len(adult)}／障害児{len(child)}"
          f"（うち見込量0は{len([x for x in adult + child if x['users'] == 0])}）")
    print(f"  据え置き{len(HOLDS)}件／村への確認事項{len(MURA_CONFIRM)}件")
    print(f"  自己点検: {len(checks)}件すべて合")


if __name__ == "__main__":
    main()
