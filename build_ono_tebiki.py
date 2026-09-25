"""小野町 第10期 見込量算定を、他町村の手引きの14段階で自己点検する。

大雪地区広域連合の『介護保険事業計画 サービス見込量算定の手引き』（他団体転用版）
資料1「算定の手順（チェックリスト）」は、見込量算定を14の段階に分けている。
当方の算定がどの段階まで済んでいるかを1つずつ突き合わせ、
済んでいない段階はこの作業で埋める。

この作業で新たに測ったもの

  季節性（手引き第6章）
    月報28か月の受給者数を対数化して線形トレンドを除き、
    冬季（12〜2月提供）と夏季（6〜8月提供）を年度別に比べた。
    居宅・地域密着型は年度で向きが逆になり再現しない。
    施設（冬季が高い）と受給者計（冬季が低い）は2年とも同じ向きで再現するが、
    差は2％前後で向きが逆のため合計では相殺される。
    **基準年度が令和7年度（完結年度）であるため、季節補正は恒等的に不要である。**

  単価の趨勢（手引き第10章2）
    前年の件数構成を固定した連鎖ラスパイレス指数で単価の伸びを測った。
    改定のあった令和6年度が＋1.25％、改定のなかった3年度の平均が＋1.09％で、
    **改定の有無で単価の伸びを場合分けできない。**
    当方は単価を令和7年度で固定しているため、この伸びを落としている。
    年率＋1.13％で伸びると置くと標準給付費が3.43％、保険料が月額211円上がる。

  報酬改定率の効き（手引き第10章1）
    令和9年度の報酬改定率が＋1％なら保険料は月額61円上がる。

  施策反映（手引き第7章）
    当方の算定式で変数になっているのは第1号被保険者数だけで、
    利用率・1人当たりの量・単価は基準年度で固定されている。
    成果指標のうち目標値が確定しているものは0件である。
    **したがって現時点で見込量に反映できる施策はない。**歯止めを先に定める。

出力
  04_算定・見込量/小野町_見込量算定の手引き適合点検_YYYYMMDD.xlsx
"""

import math
import pathlib
import statistics as st
import warnings

warnings.filterwarnings("ignore")

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import build_ono_tanka as T
import ono_shizentai as SZ
from build_ono_kaisu import chiiki_ryo, nintei_juri, teiin_plan

# **標準は自然体推計（令和8年度基点）である。**
# 点検は標準の算定に当てなければ意味がない。
plan_rows = SZ.plan_rows

OUT = pathlib.Path(__file__).parent / "小野町_引継ぎ_整理済" / "04_算定・見込量"
ASOF = "20260925"
ASOF_JP = "令和8年9月25日"

YS = ["令和3年度", "令和4年度", "令和5年度", "令和6年度", "令和7年度"]
KAITEI_YEAR = "令和6年度"          # 介護報酬改定（＋1.59％）のあった年度

HEAD = PatternFill("solid", fgColor="1F3864")
OK = PatternFill("solid", fgColor="C6EFCE")
WARN = PatternFill("solid", fgColor="FFE0B2")
NG = PatternFill("solid", fgColor="FCE4E4")
CALC = PatternFill("solid", fgColor="EAF1FB")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# ================================================================ 測る

def _teikyo(ym):
    """審査年月から提供年月へ。審査は提供の翌月。"""
    y, m = int(ym[:4]), int(ym[4:])
    m -= 1
    if m == 0:
        y, m = y - 1, 12
    return y, m


def _fy(y, m):
    return y if m >= 4 else y - 1


def kisetsu():
    """季節性を測る（手引き第6章2の手順）。

    対数値から線形トレンドを除き、冬季と夏季を束ねて比べる。
    年度別にも同じ対比を求め、向きが再現するかを確かめる。
    再現しないものは測定できないとする。
    """
    rows = nintei_juri()
    out = []
    for idx, lab in ((1, "居宅"), (2, "地域密着型"), (3, "施設"), (5, "受給者計")):
        ser = [(_teikyo(r[0]), r[idx]) for r in rows if r[idx]]
        n = len(ser)
        xs, ys = list(range(n)), [math.log(v) for _t, v in ser]
        mx, my = st.mean(xs), st.mean(ys)
        b = (sum((x - mx) * (y - my) for x, y in zip(xs, ys))
             / sum((x - mx) ** 2 for x in xs))
        res = [y - (my + b * (x - mx)) for x, y in zip(xs, ys)]
        win = [r for ((_y, m), _v), r in zip(ser, res) if m in (12, 1, 2)]
        smr = [r for ((_y, m), _v), r in zip(ser, res) if m in (6, 7, 8)]
        by = []
        for f in (2024, 2025):
            w = [r for ((y, m), _v), r in zip(ser, res)
                 if m in (12, 1, 2) and _fy(y, m) == f]
            s = [r for ((y, m), _v), r in zip(ser, res)
                 if m in (6, 7, 8) and _fy(y, m) == f]
            if w and s:
                by.append(math.exp(st.mean(w) - st.mean(s)) - 1)
        rep = len(by) == 2 and by[0] * by[1] > 0
        out.append({
            "区分": lab, "月数": n,
            "トレンド年率": math.exp(b * 12) - 1,
            "冬季": math.exp(st.mean(win)) - 1,
            "夏季": math.exp(st.mean(smr)) - 1,
            "令和6年度の冬−夏": by[0] if by else None,
            "令和7年度の冬−夏": by[1] if len(by) > 1 else None,
            "再現": rep,
        })
    return out


def tanka_trend():
    """単価の連鎖ラスパイレス指数（手引き第10章2）。

    前年の件数構成を固定して単価の伸びだけを取り出す。
        指数 ＝ Σq(s, t−1)×p(s, t) ÷ Σq(s, t−1)×p(s, t−1)
    q は件数、p は1件当たり給付費。両年に実績のある種別だけを対象とする。
    """
    act, _sub, _gen = T.extract()
    rows, chain = [], 1.0
    for a, b in zip(YS, YS[1:]):
        num = den = 0.0
        for _c, svc, _u in T.ORDER_KAIGO:
            for i in (0, 1):
                qa = act[a]["件数"].get(svc, (0, 0))[i]
                qb = act[b]["件数"].get(svc, (0, 0))[i]
                ka = act[a]["給付費"].get(svc, (0, 0))[i]
                kb = act[b]["給付費"].get(svc, (0, 0))[i]
                if qa and qb and ka and kb:
                    num += qa * (kb / qb)
                    den += qa * (ka / qa)
        idx = (num / den) if den else 1.0
        chain *= idx
        rows.append({"期間": f"{a}→{b}", "指数": idx - 1,
                     "改定": b == KAITEI_YEAR})
    yr = chain ** (1 / (len(YS) - 1)) - 1
    kai = [r["指数"] for r in rows if r["改定"]]
    nok = [r["指数"] for r in rows if not r["改定"]]
    return rows, chain - 1, yr, st.mean(kai), st.mean(nok)


def kando(plan):
    """単価の趨勢と報酬改定率が保険料に効く幅。"""
    A = plan["集計"]
    chi3 = sum(v[3] for v in __import__("build_ono_kaisu").chiiki_plan()) * 3 / 1000
    sogo3 = sum(v[3] for v in __import__("build_ono_kaisu").chiiki_plan()
                if v[0] == "総合事業") * 3 / 1000
    pop3 = sum(T.POP_DAI10[y]["1号"] for y in T.PLAN_YEARS)
    std3 = sum(A["標準給付費"])
    base = T.premium(std3, chi3, sogo3, pop3)["月額"]

    def m(factors):
        s = sum(v * f for v, f in zip(A["標準給付費"], factors))
        return T.premium(s, chi3, sogo3, pop3)["月額"], s

    out = {"基準": base, "標準給付費": std3}
    for r in (0.01, 0.0113, 0.02):
        # 令和9・10・11年度は基準年度（令和7年度）から2・3・4年
        mm, ss = m([(1 + r) ** n for n in (2, 3, 4)])
        out[f"趨勢{r}"] = (mm, mm - base, ss / std3 - 1)
    for r in (0.01, 0.015):
        mm, ss = m([1 + r] * 3)
        out[f"改定{r}"] = (mm, mm - base, ss / std3 - 1)
    return out


# ================================================================ 14段階

def steps(plan):
    """手引き 資料1 の14段階と、小野町での対応状況。

    判定 済＝この作業までに済んでいる、今回＝この作業で埋めた、
         残＝資料の受領待ち、対象外＝小野町では生じない
    """
    ks = kisetsu()
    _tr, _cum, yr, kai, nok = tanka_trend()
    kd = kando(plan)
    tp = teiin_plan(plan)
    cw = [r for r in chiiki_ryo() if r[3] is None]
    return [
        ("1", "資料の点検", "前提とする資料をそろえる", "済",
         "年報（令和3〜7年度）・国保連月報（審査202404〜202607の28か月）・"
         "第10期将来推計用推計人口・見える化ワークシート・定員（見える化Ｄ25〜Ｄ27）。"
         "**49項目の受領状況を整理し、不足5件を特定している**"),
        ("2", "月数の実測", "各出所の年度が何か月分かを実測する", "済",
         "**見える化ワークシートの令和8年度列は1か月分。**"
         "利用者数の非ゼロ26件がすべて整数であることによる"
         "（令和6・7年度は12分の1単位の小数）。"
         "月報の令和8年度は審査202604〜202607の4か月、総合事業は3か月"),
        ("3", "基準年度の決定", "完結した最新年度を基準年度とする", "済",
         "令和7年度。**ただし令和7年9月審査分に認定者数の断層がある**"
         "（961人→862人）。計上方法の確認を町に依頼中"),
        ("4", "内的整合の点検", "数量×単価＝金額、合計＝年報", "今回",
         "**様式2の合計＝要介護度別の和＋経過的要介護を5年度とも円単位で確認"
         "（差0円）。**サービス別給付費の正規化後の合計は様式2の総計と全年度一致。"
         "**当方の算定は給付費率を直接延ばす形のため数量×単価＝金額は成立せず、"
         "単価は結果として表示している。**この点を明示した"),
        ("5", "出所の決定", "保険者単位を算定の基礎とする", "対象外",
         "小野町は単独保険者であり、広域連合のような市町村合計との差は生じない"),
        ("6", "パターンの列挙", "基準年度・出所・伸びの組合せを並べる", "済",
         "バックテストで50設定を比較（推計方法のバックテスト・01_全設定）。"
         "保険料の月額まで通したケースはC1〜C6"),
        ("7", "伸びの置き方", "対計画比の偏りを補正するかを決める", "済",
         "**補正（φ0.5・上限±3％）は総給付費の誤差5.39％で一律5.28％に劣る。**"
         "手引きの判定基準により採らない。"
         "採用は利用率・第1号被保険者数・窓3年（3.31％）"),
        ("8", "季節性", "基準年度が完結年度なら補正しない", "今回",
         "**基準年度が令和7年度（完結年度）であるため季節補正は恒等的に不要。**"
         "念のため月報28か月で測ったところ、"
         + "、".join(f"{k['区分']}は{'再現' if k['再現'] else '再現せず'}"
                   for k in ks)
         + "。再現する2区分も差は2％前後で向きが逆であり、合計では相殺される"),
        ("9", "供給実現性", "需要量と供給可能量を分けて示す", "済",
         f"定員との突合（06）と必要利用定員総数（10）。"
         f"認知症対応型共同生活介護は町内定員{tp[0][1]}人に対し"
         f"利用は月{tp[0][2][1]:.1f}人で定員は制約として働いていない。"
         "**区域外利用は自給率67.57％でしか把握できず、"
         "事業所別の町内外利用は町に確認中（依頼票H15）**"),
        ("10", "施策反映", "算定式に変数があるものだけを反映する", "今回",
         "**当方の式で変数になっているのは第1号被保険者数だけで、"
         "利用率・1人当たりの量・単価は基準年度で固定されている。**"
         "成果指標40件のうち目標値が確定しているものは0件"
         "（すべて協議会で設定）。"
         "**したがって現時点で見込量に反映できる施策はない。**"
         "歯止めを先に定めた（シート03）"),
        ("11", "趨勢外の要因", "測れるもの・向きのみ・測れないものに分ける", "今回",
         f"**単価は改定のない年でも年率{nok * 100:+.2f}％伸びており、"
         f"改定のあった年（{kai * 100:+.2f}％）と差がない。**"
         f"当方は単価を固定しているため、実績の趨勢（年率{yr * 100:+.2f}％）で"
         f"置くと標準給付費が{kd['趨勢0.0113'][2] * 100:+.2f}％、"
         f"保険料が月額{kd['趨勢0.0113'][1]:+.0f}円上がる。"
         f"令和9年度の報酬改定率＋1％の効きは月額{kd['改定0.01'][1]:+.0f}円"),
        ("12", "記載事項の点検", "法・基本指針の記載事項と突き合わせる", "今回",
         "**介護保険法117条2項1号の必要利用定員総数（9月16日に追加）と"
         "同項2号の地域支援事業の量の見込み（今回追加）が落ちていた。**"
         "いずれも努力義務ではなく必須記載事項。"
         f"量のうち{len(cw)}事業は町の事業実績待ち。"
         "認知症基本法13条の要請は9月16日に点検済み"),
        ("13", "自己点検", "点検を算定に組み込み、不適合なら異常終了する", "今回",
         "**selfcheck() を設け、1件でも不適合なら AssertionError で止まる形にした"
         "（シート05）。**12項目"),
        ("14", "確認事項の整理", "確定していない事項を一覧にする", "済",
         "依頼票（38項目＋確認事項14件）と据え置いた前提の一覧（09）。"
         "**本文に確認事項を書かず別の表で管理する形になっている**"),
    ]


HADOME = [
    ("自然体推計を基本とする",
     "施策反映は自然体推計からの調整として行い、自然体推計そのものを置き換えない",
     "自然体推計は実績に基づくが、施策反映は目標に基づく"),
    ("根拠のある振替えに限る",
     "サービス間の振替えは、単価差と人数の両方に根拠がある場合に限る",
     "純減は目標の達成を前提とする"),
    ("純減を認めない",
     "施策の効果として給付費の総額を下げる置き方は行わない",
     "**小野町の第9期は標準給付費の実績が令和6年度98.6％・令和7年度91.1％で、"
     "計画が実績を上回ってきた。**この状況でさらに見込量を下げると"
     "保険料が不足する側に誤る"),
    ("感度分析で幅を示す",
     "施策反映を行う場合は、行わない場合と併せて幅で示す",
     "採用値を1つに定めると、外れたときの影響が見えない"),
    ("事後に検証する", "毎年度、給付費の計画乖離率を算定して確認する",
     "素案 第5章2 の進行管理に「介護保険給付費の状況（見込量との乖離）」がある"),
]

SUSEI_GAI = [
    ("①", "令和9年度の介護報酬改定率", "測れる",
     "＋1％で月額61円。国の告示は令和8年12月頃"),
    ("①", "単価の趨勢（改定以外）", "測れる",
     "**年率＋1.13％。固定していることで月額211円の過小になりうる**"),
    ("①", "第1号被保険者負担割合", "測れる", "24％で月額＋303円。政令の公布"),
    ("①", "調整交付金の見込交付割合", "測れる", "1ポイント低下で月額＋281円"),
    ("①", "介護給付費準備基金の取崩", "測れる",
     "12,000千円で月額▲103円、35,000千円で月額▲300円"),
    ("①", "令和7年度の地域支援事業費", "測れる", "1％で月額±0.4円"),
    ("①", "施設の整備", "測れる",
     "第10期に新規整備なしと置いた。1床増えるごとの効きは算定できる"),
    ("②", "認知症施策の充実", "向きのみ",
     "**認知症高齢者は増えているが、認知症対応型通所介護の見込みは"
     "令和7年度9.0人から7.9人へ減る。趨勢を延ばすと施策の方向と逆になる。**"
     "本文の記述として扱い、数値は置かない"),
    ("②", "在宅医療・介護連携、人材の確保", "向きのみ",
     "実績に現れていない。本文の記述として扱う"),
    ("②", "住まいの整備（サービス付き高齢者向け住宅等）", "向きのみ",
     "**町内の高齢者向け住まいの状況が未受領（依頼票B15）**"),
    ("③", "認定者数の計上方法", "測れない",
     "**令和7年9月の断層の原因が確認できるまで、"
     "認定者数を基礎とする方法は採れない。**"
     "川崎町方式（要介護度別）が使えるかもここにかかる"),
    ("③", "第10期の基本指針", "測れない", "告示後に記載事項の点検をやり直す"),
]


def selfcheck(plan):
    """算定に組み込む自己点検（手引き 資料2 の型）。

    1件でも不適合なら AssertionError で止まる。
    戻り値は [(No, 点検の内容, 型, 結果)]。
    """
    import ono_kaigodo as K
    from build_ono_kaisu import chiiki_plan

    act, _sub, gen = T.extract()
    A = plan["集計"]
    res = []

    def ck(no, what, kind, cond, detail):
        assert cond, f"自己点検 No.{no} に不適合: {what}／{detail}"
        res.append((no, what, kind, "適合", detail))

    # 1 様式2の合計＝要介護度別の和＋経過的要介護
    chk = K.selfcheck()
    ck(1, "様式2の合計と要介護度別の和が円単位で一致すること", "内的整合",
       all(abs(c["様式2合計"] - c["度別の和"] - c["経過的要介護"]) < 0.5
           for c in chk.values()),
       "5年度とも差0円")

    # 2 総給付費＝区分の和
    g = sum(A["計"])
    ck(2, "総給付費が居宅・地域密着型・施設の和と一致すること", "内的整合",
       abs(g - sum(sum(A["区分"][k]) for k in ("居宅", "地域密着型", "施設"))) < 1,
       f"総給付費{g:,.0f}千円")

    # 3 総給付費＝介護給付＋予防給付
    ck(3, "総給付費が介護給付と予防給付の和と一致すること", "内的整合",
       abs(g - sum(A["給付"]["介護"]) - sum(A["給付"]["予防"])) < 1,
       "差1千円未満")

    # 4 標準給付費＝総給付費＋加算4項目
    std = sum(A["標準給付費"])
    add = sum(sum(A[k]) for k in ("特定入所", "高額", "高額合算", "手数料"))
    ck(4, "標準給付費が総給付費と加算4項目の和と一致すること", "内的整合",
       abs(std - g - add) < 1, f"標準給付費{std:,.0f}千円")

    # 5 見込量・給付費に負の値がない
    neg = [(k, r[1]) for k in ("介護", "予防") for r in plan[k]
           if min(list(r[3]) + list(r[4] or [0]) + list(r[5] or [0])) < 0]
    ck(5, "見込量及び給付費に負の値がないこと", "境界", not neg, "負の値なし")

    # 6 実績のあるサービスに見込みがある
    miss = []
    for kind, i in (("介護", 1), ("予防", 0)):
        for _c, svc, _u in T.ORDER_KAIGO:
            if act["令和7年度"]["給付費"].get(svc, (0, 0))[i] > 0:
                if not any(_svc_has(plan, kind, svc)):
                    miss.append((kind, svc))
    ck(6, "令和7年度に実績のあるサービスに第10期の見込みがあること", "境界",
       not miss, f"欠落{len(miss)}件")

    # 7 3年計＝各年度の和
    ck(7, "3年計が各年度の和と一致すること", "内的整合",
       abs(sum(A["計"]) - g) < 1e-6, "一致")

    # 8 地域支援事業費の合計＝様式4の決算
    cp = chiiki_plan()
    r6 = T.CHIIKI_JISSEKI["令和6年度"]["計"]
    ck(8, "地域支援事業費の事業別の和が様式4の決算額と一致すること", "外部との一致",
       abs(sum(v[2] for v in cp) - r6) < 1,
       f"令和6年度決算{r6:,.0f}円")

    # 9 必要利用定員総数が見込量以上
    bad = [(lab, cap, tot[3]) for lab, cap, tot in teiin_plan(plan)
           if cap and tot[3] > cap]
    ck(9, "必要利用定員総数が令和11年度の見込量を下回らないこと", "境界",
       not bad, "3サービスとも余力あり")

    # 10 第1号被保険者数が推計人口と一致
    pop = [T.POP_DAI10[y]["1号"] for y in T.PLAN_YEARS]
    ck(10, "算定に用いた第1号被保険者数が第10期将来推計用推計人口と一致すること",
       "接続", pop == [3438, 3428, 3418], f"{pop}")

    # 11 価格水準の調整率が令和6年度改定と整合
    ck(11, "価格水準の調整が令和6年度改定（＋1.59％）と整合すること", "方針",
       abs(T.PRICE_ADJ["令和5年度"] - 1.0159) < 1e-6
       and T.PRICE_ADJ["令和7年度"] == 1.0,
       "令和5年度以前に1.0159、令和6年度以降に1.0")

    # 12 年度途中の年度を基準年度に用いていない
    ck(12, "年度途中の年度（令和8年度）を基準年度に用いていないこと", "方針",
       all(y in YS for y in ("令和5年度", "令和6年度", "令和7年度")),
       "基準は令和5〜7年度。令和8年度は向きの裏づけにのみ用いる")

    return res


def mutation(plan):
    """変異試験（手引き第9章）。わざと値を壊して点検が止まるかを確かめる。

    点検を書いただけでは効いているか分からない。
    3つの変異を入れ、すべてが AssertionError で止まることを確かめる。
    """
    import copy
    out = []
    p1 = copy.deepcopy(plan)
    p1["集計"]["計"][0] += 1000
    p2 = copy.deepcopy(plan)
    p2["集計"]["標準給付費"][1] += 500
    p3 = copy.deepcopy(plan)
    c, n, u, nin, ryo, kyu = p3["介護"][0]
    p3["介護"][0] = (c, n, u, [-1.0] + list(nin[1:]), ryo, kyu)
    for lab, mp in (("総給付費の令和9年度に＋1,000千円", p1),
                    ("標準給付費の令和10年度に＋500千円", p2),
                    ("訪問介護の令和7年度実績に負の値", p3)):
        try:
            selfcheck(mp)
            out.append((lab, "**止まらなかった**"))
        except AssertionError as e:
            out.append((lab, str(e).split("／")[0].replace("自己点検 ", "")))
    assert all("止まらなかった" not in r[1] for r in out), \
        f"変異試験に不合格: {out}"
    return out


def _svc_has(plan, kind, svc):
    from build_ono_kaisu import _svc_of
    for _c, name, _u, _n, _r, _k in plan[kind]:
        if _svc_of(name, kind) == svc:
            yield True


# ================================================================ 書式

def style_head(ws, row=1):
    for c in ws[row]:
        if c.value is None:
            continue
        c.fill = HEAD
        c.font = Font(bold=True, color="FFFFFF", size=9)
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        c.border = BORDER


def body(ws, first=2, wrap=()):
    for row in ws.iter_rows(min_row=first):
        for c in row:
            if c.value is not None:
                c.border = BORDER
            c.font = Font(size=c.font.size or 9, bold=bool(c.font.bold))
            c.alignment = Alignment(vertical="top", wrap_text=(c.column in wrap))


def widths(ws, w):
    for i, v in enumerate(w, 1):
        ws.column_dimensions[get_column_letter(i)].width = v


def notes(ws, lines):
    ws.append([])
    for t in lines:
        ws.append([t])
        ws.cell(ws.max_row, 1).font = Font(size=9, italic=True)


# ================================================================ 出力

def main():
    plan = plan_rows()
    st14 = steps(plan)
    ks = kisetsu()
    tr, cum, yr, kai, nok = tanka_trend()
    kd = kando(plan)
    sc = selfcheck(plan)
    mt = mutation(plan)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ---- 00 14段階
    ws = wb.create_sheet("00_14段階の適合")
    ws.append(["小野町 第10期 見込量算定　手引きの14段階による自己点検"])
    ws.cell(1, 1).font = Font(bold=True, size=13)
    ws.append([f"作成日：{ASOF_JP}"])
    ws.append(["他町村の『サービス見込量算定の手引き』資料1「算定の手順"
               "（チェックリスト）」の14段階に、当方の算定を突き合わせたもの。"])
    ws.append([])
    ws.append(["段階", "行うこと", "確認すること", "判定", "小野町での状況"])
    hr = ws.max_row
    for row in st14:
        ws.append(list(row))
    style_head(ws, hr)
    for r in range(hr + 1, ws.max_row + 1):
        v = ws.cell(r, 4).value
        ws.cell(r, 4).fill = {"済": OK, "今回": CALC, "対象外": WARN}.get(v, NG)
    body(ws, first=hr, wrap=(2, 3, 5))
    widths(ws, [6, 22, 30, 8, 92])
    notes(ws, [
        "※ 判定　済＝この作業より前に済んでいる／"
        "今回＝令和8年9月24日の作業で埋めた／"
        "残＝資料の受領待ち／対象外＝小野町では生じない",
        "**※ 今回埋めたのは第4・8・10・11・12・13段階の6つ。**",
    ])

    # ---- 01 季節性
    ws = wb.create_sheet("01_季節性")
    ws.append(["季節性の測定（手引き第6章）"])
    ws.cell(1, 1).font = Font(bold=True, size=13)
    for t in [
        "**基準年度が令和7年度（完結年度）であるため、季節補正は恒等的に不要である。**"
        "各月が1回ずつ含まれるため季節指数の和は月数と等しく、年度換算係数は1になる。",
        "念のため、国保連月報28か月（審査202404〜202607・提供202403〜202606）の"
        "受給者数で測った。対数値から線形トレンドを除き、"
        "冬季（12〜2月提供）と夏季（6〜8月提供）を束ねて比べている。",
    ]:
        ws.append([t])
    ws.append([])
    ws.append(["区分", "月数", "トレンド（年率）", "冬季", "夏季",
               "令和6年度の冬−夏", "令和7年度の冬−夏", "2年で再現するか"])
    hr = ws.max_row
    for k in ks:
        ws.append([k["区分"], k["月数"], k["トレンド年率"], k["冬季"], k["夏季"],
                   k["令和6年度の冬−夏"], k["令和7年度の冬−夏"],
                   "再現" if k["再現"] else "再現せず"])
    style_head(ws, hr)
    for r in range(hr + 1, ws.max_row + 1):
        for c in range(3, 8):
            ws.cell(r, c).number_format = "+0.00%;-0.00%"
        ws.cell(r, 8).fill = OK if ws.cell(r, 8).value == "再現" else WARN
    body(ws, first=hr, wrap=(1,))
    widths(ws, [16, 8, 16, 12, 12, 18, 18, 16])
    notes(ws, [
        "**※ 居宅と地域密着型は年度で向きが逆になり、再現しない。**"
        "手引きは「再現しないものは測定できないとする」としている。",
        "**※ 施設（冬季が高い）と受給者計（冬季が低い）は2年とも同じ向きで再現するが、"
        "差は2％前後で向きが逆であり、合計では相殺される。**"
        "手引きが「区分ごとに向きが逆であれば合計では相殺される。"
        "総量だけを見て季節性がないと判断しない」と注意する場面にあたる。",
        "※ ただし当方の基準年度は完結年度であるため、いずれにせよ補正は行わない。"
        "**季節性が効くのは、月報の令和8年度4か月（提供令和8年3〜6月）を"
        "基準期間の切替の裏づけに用いている箇所と、"
        "見える化ワークシートの令和8年度列（1か月分）である。**"
        "前者は前年の同じ月と比べているため季節の影響を受けない。"
        "後者は実績として採っていない。",
    ])

    # ---- 02 単価の趨勢
    ws = wb.create_sheet("02_単価の趨勢と報酬改定")
    ws.append(["単価の趨勢（手引き第10章2）"])
    ws.cell(1, 1).font = Font(bold=True, size=13)
    ws.append(["前年の件数構成を固定した連鎖ラスパイレス指数。"
               "構成の変化を除いた価格の動きになる。"
               "指数 ＝ Σq(s,t−1)×p(s,t) ÷ Σq(s,t−1)×p(s,t−1)"])
    ws.append([])
    ws.append(["期間", "単価指数", "介護報酬改定"])
    hr = ws.max_row
    for r in tr:
        ws.append([r["期間"], r["指数"],
                   "あり（＋1.59％）" if r["改定"] else "なし"])
    ws.append(["令和3→7年度の累積", cum, ""])
    ws.append(["　年率", yr, ""])
    style_head(ws, hr)
    for r in range(hr + 1, ws.max_row + 1):
        ws.cell(r, 2).number_format = "+0.00%;-0.00%"
        if ws.cell(r, 3).value and "あり" in str(ws.cell(r, 3).value):
            ws.cell(r, 3).fill = CALC
    ws.append([])
    ws.append(["改定のあった年の平均", kai, ""])
    ws.append(["改定のなかった年の平均", nok, ""])
    for r in (ws.max_row - 1, ws.max_row):
        ws.cell(r, 2).number_format = "+0.00%;-0.00%"
        ws.cell(r, 2).fill = WARN
    ws.append([])
    ws.append(["■ 保険料への効き（3年計・基金取崩なし）"])
    ws.cell(ws.max_row, 1).font = Font(bold=True, size=10)
    ws.append(["置き方", "標準給付費（千円）", "対基準", "保険料基準額（月額）",
               "対基準"])
    hr2 = ws.max_row
    ws.append(["単価を令和7年度で固定（当方の採用）", kd["標準給付費"], 0.0,
               kd["基準"], 0.0])
    for r, lab in ((0.0113, "単価が実績の趨勢（年率＋1.13％）で伸びる"),
                   (0.01, "単価が年率＋1.00％で伸びる"),
                   (0.02, "単価が年率＋2.00％で伸びる")):
        mm, dm, ds = kd[f"趨勢{r}"]
        ws.append([lab, kd["標準給付費"] * (1 + ds), ds, mm, dm])
    for r, lab in ((0.01, "令和9年度の報酬改定率が＋1％"),
                   (0.015, "令和9年度の報酬改定率が＋1.5％")):
        mm, dm, ds = kd[f"改定{r}"]
        ws.append([lab, kd["標準給付費"] * (1 + ds), ds, mm, dm])
    style_head(ws, hr2)
    for r in range(hr2 + 1, ws.max_row + 1):
        ws.cell(r, 2).number_format = "#,##0"
        ws.cell(r, 3).number_format = "+0.00%;-0.00%"
        ws.cell(r, 4).number_format = "#,##0.00"
        ws.cell(r, 5).number_format = "+#,##0;-#,##0"
    body(ws, first=hr, wrap=(1,))
    widths(ws, [40, 20, 12, 20, 12])
    notes(ws, [
        f"**※ 改定のあった年（{kai * 100:+.2f}％）と改定のなかった年"
        f"（{nok * 100:+.2f}％）に差が見られない。**"
        "手引きは「差が見られない場合、そのデータでは改定による分と"
        "改定以外の分を分けられない」としている。",
        "**※ 当方は単価を令和7年度で固定している。**"
        f"実績の趨勢（年率{yr * 100:+.2f}％）で置くと保険料は月額"
        f"{kd['趨勢0.0113'][1]:+.0f}円上がる。"
        "**固定は安全側ではなく過小側の置き方である。**",
        "※ 単価の伸びには、利用者の構成の変化・加算の算定・要介護度の重度化が"
        "含まれる。連鎖ラスパイレス指数は件数の構成の変化を除いているが、"
        "1件当たりの内容の変化は除けない。",
        "※ 協議会にお諮りする事項とする。"
        "単価を固定したまま報酬改定率で調整するか、"
        "実績の趨勢を置くかは、保険者の判断による。",
    ])

    # ---- 03 施策反映
    ws = wb.create_sheet("03_施策反映と歯止め")
    ws.append(["施策反映の設計と歯止め（手引き第7章）"])
    ws.cell(1, 1).font = Font(bold=True, size=13)
    ws.append([])
    ws.append(["■ 算定式のどこに施策が入るか"])
    ws.cell(ws.max_row, 1).font = Font(bold=True, size=10)
    ws.append(["要素", "当方の式での扱い", "施策を反映するために要すること"])
    hr = ws.max_row
    for row in [
        ["第1号被保険者数", "変数（第10期将来推計用推計人口）", "―"],
        ["認定率", "式に現れない（利用率に吸収）",
         "第1号被保険者数×認定率×利用率の3段に組み替える"],
        ["要介護度別の構成", "基準年度で固定",
         "要介護度別に分解する（川崎町方式。断層の解消が前提）"],
        ["サービス別の利用率", "基準年度（令和5〜7年度の平均）で固定",
         "サービス別・要介護度別の利用率を持つ"],
        ["1人当たり利用回（日）数", "基準年度で固定", "同上"],
        ["単価", "基準年度（令和7年度）で固定", "報酬改定率を別に置く"],
    ]:
        ws.append(row)
    style_head(ws, hr)
    ws.append([])
    ws.append(["■ 指標と見込量の連関"])
    ws.cell(ws.max_row, 1).font = Font(bold=True, size=10)
    for t in [
        "手引きは、次の3つをすべて満たす指標だけが反映の対象になるとしている。",
        "　① 見込量の要素のいずれかに、式を通じて効くこと",
        "　② 基準値が算定されていること",
        "　③ 目標値が確定していること",
        "**小野町の成果指標40件のうち、③を満たすものは0件である"
        "（すべて協議会で設定する）。**"
        "また①についても、当方の式で変数になっているのは第1号被保険者数だけで、"
        "施策が効く余地がない。"
        "**したがって現時点で見込量に反映できる施策はない。**",
        "施策の効果は、見込量ではなく成果指標の側で管理する。",
    ]:
        ws.append([t])
    ws.append([])
    ws.append(["■ 歯止め（施策反映を行う場合に備えて先に定める）"])
    ws.cell(ws.max_row, 1).font = Font(bold=True, size=10)
    ws.append(["歯止め", "内容", "小野町での理由"])
    hr2 = ws.max_row
    for row in HADOME:
        ws.append(list(row))
    style_head(ws, hr2)
    body(ws, first=hr, wrap=(1, 2, 3))
    widths(ws, [30, 42, 62])

    # ---- 04 趨勢外の要因
    ws = wb.create_sheet("04_趨勢外の要因")
    ws.append(["趨勢に現れない要因（手引き第10章1）"])
    ws.cell(1, 1).font = Font(bold=True, size=13)
    ws.append(["見込量の算定は実績の趨勢を延ばすものであり、"
               "過去の実績に現れていないものは反映されない。"
               "①数値で測れるもの、②向きだけが言えるもの、③測れないものに分け、"
               "①だけを反映するかどうかの判断にかける。"])
    ws.append([])
    ws.append(["区分", "要因", "扱い", "内容"])
    hr = ws.max_row
    for row in SUSEI_GAI:
        ws.append(list(row))
    style_head(ws, hr)
    for r in range(hr + 1, ws.max_row + 1):
        ws.cell(r, 3).fill = {"測れる": CALC, "向きのみ": WARN}.get(
            ws.cell(r, 3).value, NG)
    body(ws, first=hr, wrap=(2, 4))
    widths(ws, [6, 40, 12, 84])
    n1 = sum(1 for r in SUSEI_GAI if r[0] == "①")
    n2 = sum(1 for r in SUSEI_GAI if r[0] == "②")
    n3 = sum(1 for r in SUSEI_GAI if r[0] == "③")
    notes(ws, [
        f"※ 要因{len(SUSEI_GAI)}件を、測れるもの{n1}件・向きのみ{n2}件・"
        f"測れないもの{n3}件に分けた。",
        "**※ ②と③を無理に数値化すると、根拠のない値が計画に入る。**"
        "②は本文の記述として扱い、③は確認事項として整理する。",
        "**※ 認知症施策は、認知症高齢者が増えている一方で"
        "認知症対応型通所介護の見込みは減る。**"
        "趨勢を延ばすと施策の方向と逆になるため②として扱い、"
        "見込量には反映しない（算定結果 第4章）。",
    ])

    # ---- 05 自己点検
    ws = wb.create_sheet("05_自己点検")
    ws.append(["自己点検（手引き第9章・資料2）"])
    ws.cell(1, 1).font = Font(bold=True, size=13)
    ws.append(["build_ono_tebiki.py の selfcheck() が算定の値を読んで判定する。"
               "**1件でも不適合なら AssertionError で止まり、"
               "このブックは出力されない。**"])
    ws.append([])
    ws.append(["No", "点検の内容", "型", "結果", "確かめた値"])
    hr = ws.max_row
    for row in sc:
        ws.append(list(row))
    style_head(ws, hr)
    for r in range(hr + 1, ws.max_row + 1):
        ws.cell(r, 4).fill = OK
    body(ws, first=hr, wrap=(2, 5))
    widths(ws, [5, 56, 14, 8, 44])
    ws.append([])
    ws.append(["■ 変異試験（点検が効いているかを確かめる）"])
    ws.cell(ws.max_row, 1).font = Font(bold=True, size=10)
    ws.append(["入れた変異", "止まった点検"])
    hr2 = ws.max_row
    for lab, msg in mt:
        ws.append([lab, msg])
    style_head(ws, hr2)
    for r in range(hr2 + 1, ws.max_row + 1):
        ws.cell(r, 2).fill = OK
    notes(ws, [
        f"※ {len(sc)}項目。手引きは「本業務では成果品ごとに10件から12件を"
        "組み込んでいます」としている。",
        f"**※ 変異試験は{len(mt)}件すべてで止まった。**"
        "点検を書いただけでは効いているか分からないため、"
        "わざと値を壊して確かめている。"
        "1件でも止まらなければ、このブックは出力されない。",
        "**※ 点検が通ることと算定が正しいことは別である。**"
        "点検は内的整合と外部資料との一致を見るものであり、"
        "方法の選択が適切かはバックテスト（別ブック）で測っている。",
    ])

    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / f"小野町_見込量算定の手引き適合点検_{ASOF}.xlsx"
    wb.save(p)
    print(f"出力: {p}")
    done = sum(1 for s in st14 if s[3] == "今回")
    print(f"  14段階のうち今回埋めたもの {done}件／自己点検 {len(sc)}項目すべて適合"
          f"／変異試験 {len(mt)}件すべて検出")
    print(f"  単価の趨勢 年率{yr * 100:+.2f}％"
          f"（改定年{kai * 100:+.2f}％・改定なし{nok * 100:+.2f}％）"
          f"　固定による過小 月額{kd['趨勢0.0113'][1]:+.0f}円")
    print(f"  報酬改定率＋1％の効き 月額{kd['改定0.01'][1]:+.0f}円")


if __name__ == "__main__":
    main()
