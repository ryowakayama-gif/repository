"""障がい計画：手帳所持者数の将来推計・見込量の点検・利用者別台帳の様式・給付費と財源。

なぜ作るか
  他団体の案件と照らした不足事項の洗い出し（内部検討/小野町_不足事項の洗い出しと許可URL）で、
  障がい計画の優先「高」とした4件（#1手帳の将来推計、#2利用者別の積上げ、
  #3月×12の点検、#4給付費・財源）を進める。

原典
  町から受領した見込量・目標値設定シート（04_算定・見込量、令和8年9月7日受領）
    人口基礎データ：総人口・年齢3区分・手帳所持者数（R5〜R8は町の入力値、R9〜R11は町の推計）
    見込量算定(障害福祉)・(障害児)：R5〜R8の利用者数と量（月あたり）、町の見込
  前回計画（令和6年3月）の手帳所持者数（令和元年〜令和5年。build_ono_soan_shogai.TECHO_SUII）
  障害者総合支援法 第92条・第94条・第95条（e-Gov 法令API、令和9年4月1日時点）

自己点検
  受領シートの手帳所持者数（R5〜R8）が素案（第3版）の値と一致すること、
  3手帳の和が「計」と一致すること、見込量の点検の規則が現に働くこと（否定試験）を確かめる。

出力
  04_算定・見込量/小野町_障がい計画_手帳推計・見込量点検・財源_YYYYMMDD.xlsx / .docx
"""

import pathlib
import sys
import warnings

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

import ono_horei as H
import ono_style as S
from build_ono_soan_shogai import TECHO_NEN, TECHO_SUII

warnings.filterwarnings("ignore")

HERE = pathlib.Path(__file__).parent
ROOT = HERE / "小野町_引継ぎ_整理済"
OUT = ROOT / "04_算定・見込量"
SRC = OUT / "【受領】小野町_第8期障がい福祉計画_第4期障がい児福祉計画_見込量目標値設定シート_20260907.xlsx"
ASOF = "20261020"
ASOF_JP = "令和8年10月20日"
YRS = ["R5", "R6", "R7", "R8", "R9", "R10", "R11"]
TECHO3 = ["身体障害者手帳所持者数", "療育手帳所持者数", "精神障害者保健福祉手帳所持者数"]
KAISHO = 23          # 日中活動系の月の開所日の目安（1人当たり人日の上限の点検に用いる）


def _f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def read_jinko():
    ws = openpyxl.load_workbook(SRC, read_only=True, data_only=True)["人口基礎データ"]
    d = {}
    for r in ws.iter_rows(values_only=True):
        if r[0] and r[2] is not None and _f(r[2]) is not None:
            d[str(r[0])] = [_f(x) for x in r[2:9]]
    return d


def read_mikomi(sheet):
    ws = openpyxl.load_workbook(SRC, read_only=True, data_only=True)[sheet]
    out = []
    for r in ws.iter_rows(min_row=6, values_only=True):
        if not r[1]:
            continue
        out.append({
            "分類": r[0], "サービス": r[1], "単位": r[2],
            "人": [_f(x) for x in r[3:7]], "方法人": r[7],
            "人見込": [_f(x) for x in r[11:14]],
            "量": [_f(x) for x in r[14:18]], "方法量": r[18],
            "量見込": [_f(x) for x in r[22:25]], "備考": r[25] if len(r) > 25 else None})
    return out


# ------------------------------------------------------------------ #1 手帳の将来推計
def suikei(j):
    pop = j["総人口"]
    res = {}
    for nm in TECHO3:
        v = j[nm]
        rate = [a / p for a, p in zip(v[:4], pop[:4])]              # R5〜R8の所持率
        # A：所持率一定（R8の所持率×推計人口）
        a = [rate[3] * p for p in pop[4:]]
        # B：所持率の趨勢（R5〜R8の所持率の線形の傾き）
        n = 4
        xs = list(range(n))
        mx, my = sum(xs) / n, sum(rate) / n
        sl = sum((x - mx) * (y - my) for x, y in zip(xs, rate)) / sum((x - mx) ** 2 for x in xs)
        b = [(rate[3] + sl * (k + 1)) * p for k, p in enumerate(pop[4:])]
        res[nm] = {"実績": v[:4], "所持率": rate, "A": a, "B": b, "町": v[4:],
                   "下": [min(x, y) for x, y in zip(a, b)], "上": [max(x, y) for x, y in zip(a, b)]}
    return res


# ------------------------------------------------------------------ #2・#3 見込量の点検
def tenken(rows):
    """各サービスの1人当たり量と、町の見込の整合を点検する。"""
    out = []
    for x in rows:
        hit = []
        jin, ryo = x["人"], x["量"]
        per = [(q / p) if (p and q is not None and p > 0) else None for p, q in zip(jin, ryo)]
        unit = str(x["単位"] or "")
        # 月×12の取り違え：月あたりの量なのに、1人当たりが月の日数を超えるもの
        if "人日" in unit:
            for y, v in zip(YRS[:4], per):
                if v is not None and v > KAISHO:
                    hit.append("%s 1人当たり%.1f人日（月の開所日%d日を超える。年の値の混入を疑う）"
                               % (y, v, KAISHO))
            for y, p, q in zip(YRS[4:], x["人見込"], x["量見込"]):
                if p and q is not None and q / p > KAISHO:
                    hit.append("%s見込 1人当たり%.1f人日（同上）" % (y, q / p))
        # 実績の急な変化（前年比±50％以上かつ差が2人以上）
        for i in range(1, 4):
            a, b = jin[i - 1], jin[i]
            if a is not None and b is not None and a > 0 and abs(b - a) >= 2 and abs(b - a) / a >= 0.5:
                hit.append("%s→%s 利用者%g→%g人" % (YRS[i - 1], YRS[i], a, b))
        # 量の見込が「人見込×最近の1人当たり量」で説明できるか
        rec = [v for v in per[2:4] if v is not None]
        if rec and all(p is not None for p in x["人見込"]) and all(q is not None for q in x["量見込"]):
            u = sum(rec) / len(rec)
            for y, p, q in zip(YRS[4:], x["人見込"], x["量見込"]):
                if p and abs(q - p * u) > max(2, 0.15 * p * u):
                    hit.append("%s見込 量%g≠人%g×1人当たり%.1f（%.0f）" % (y, q, p, u, p * u))
        out.append((x, per, hit))
    return out


def _hiteishiken():
    """点検の規則が現に働くか：年の値を混ぜた行を作って止まることを確かめる。"""
    bad = [{"分類": "試験", "サービス": "試験", "単位": "人日/月", "人": [2, 2, 2, 2], "方法人": "",
            "人見込": [2, 2, 2], "量": [40, 40, 480, 40], "方法量": "", "量見込": [40, 40, 40],
            "備考": None}]
    return bool(tenken(bad)[0][2])


# ------------------------------------------------------------------ #2 年齢による移行
IKO = [
    ("18歳", "放課後等デイサービス", "生活介護・就労継続支援B型・就労選択支援",
     "高校3年生（特別支援学校高等部を含む）で放課後等デイサービスを利用している人数と進路希望"),
    ("18歳（児者の切替）", "障害児相談支援", "計画相談支援", "同上"),
    ("20歳前後", "障害児入所施設（県）", "施設入所支援・共同生活援助",
     "障害児入所施設に入所している18歳以上の人（継続入所者を含む）"),
    ("65歳", "居宅介護・生活介護など介護保険に相当するもの", "介護保険サービス（介護保険優先）",
     "64歳の利用者の人数とサービス。共生型サービスの利用の可否"),
    ("65歳（相当なし）", "共同生活援助・就労系・同行援護・行動援護", "引き続き障害福祉サービス",
     "65歳以上で利用を続ける人数"),
]

DAICHO = ["整理番号（匿名）", "生年度", "年齢（R8年度末）", "手帳の種類・等級", "障害支援区分",
          "利用中のサービス", "月の量（R8年度）", "移行の予定（18歳・65歳）", "移行先の想定",
          "本人・家族の意向（調査・相談記録）", "R9", "R10", "R11", "備考"]


# ------------------------------------------------------------------ #4 財源
def zaigen():
    r = H.get("障害者総合支援法", "2027-04-01")
    t = {jo: H.tx(H.arts(H.honbun(r))[jo]) for jo in ("92", "94", "95")}
    return [
        ("障害福祉サービス費等（介護給付費・訓練等給付費）、相談支援給付費等",
         "国 1/2・県 1/4・町 1/4", "第95条第1項第1号・第94条第1項第1号",
         "**国・県の負担は「障害福祉サービス費等負担対象額」（区分ごとの人数等で算定する"
         "国庫負担基準）に対するもの。基準を超えた分は町の負担になる**",
         ["障害福祉サービス費等負担対象額」という。）の百分の二十五", "障害福祉サービス費等負担対象額の百分の五十"], t),
        ("自立支援医療費（更生医療）・補装具費", "国 1/2・県 1/4・町 1/4",
         "第95条第1項第2号・第94条第1項第2号", "",
         ["第九十二条第三号及び第四号に掲げる費用のうち、その百分の二十五", "同条第三号及び第四号に掲げる費用の百分の五十"], t),
        ("地域生活支援事業", "国 1/2以内・県 1/4以内（補助）", "第95条第2項第2号・第94条第2項",
         "**予算の範囲内の補助で、負担金ではない。** 交付額が所要額を下回ることがある",
         ["同条第六号に掲げる費用の百分の二十五以内を補助することができる",
          "第九十二条第六号及び第九十三条第二号に掲げる費用の百分の五十以内"], t),
    ]


def jiko(j, mk, zg):
    ng = []
    for nm, (lab, vals) in zip(TECHO3, TECHO_SUII):
        # 素案（第3版）はR5〜R8の4年すべてを持つ
        if list(j[nm][:4]) != list(vals[4:8]):
            ng.append("%s：受領シート（R5〜R8 %s）と素案（%s）が違う" % (lab, j[nm][:4], vals[4:8]))
    for i in range(4):
        if sum(j[nm][i] for nm in TECHO3) != j["障害者手帳所持者数計"][i]:
            ng.append("%s：3手帳の和が計と違う" % YRS[i])
    if not _hiteishiken():
        ng.append("否定試験：年の値の混入を検出できない")
    for row in zg:
        for q in row[4]:
            if q not in "".join(row[5].values()):
                ng.append("財源：条文に「%s」がない" % q)
    return ng


# ------------------------------------------------------------------ 出力
def build_xlsx(j, sk, tk_f, tk_c, zg):
    wb = openpyxl.Workbook()
    hd = PatternFill("solid", fgColor="DDEBF7")
    # 手帳推計
    ws = wb.active
    ws.title = "1_手帳所持者数の推計"
    ws.append(["区分", "年"] + TECHO_NEN[:5] + YRS[1:4] + YRS[4:])
    ws.append(["総人口", "人", None, None, None, None, None] + j["総人口"][1:4] + j["総人口"][4:])
    for nm, (lab, vals) in zip(TECHO3, TECHO_SUII):
        s = sk[nm]
        ws.append([lab, "実績"] + vals[:5] + s["実績"][1:4] + [None] * 3)
        ws.append([lab, "所持率（‰）", None, None, None, None, None] + [round(x * 1000, 1) for x in s["所持率"][1:4]] + [None] * 3)
        ws.append([lab, "推計A 所持率一定"] + [None] * 8 + [round(x) for x in s["A"]])
        ws.append([lab, "推計B 所持率の趨勢"] + [None] * 8 + [round(x) for x in s["B"]])
        ws.append([lab, "町の推計（受領シート）"] + [None] * 8 + s["町"])
    for c in ws[1]:
        c.fill = hd
        c.font = Font(bold=True)
    # 見込量の点検
    for title, tk in (("2_見込量の点検_障害福祉", tk_f), ("2_見込量の点検_障害児", tk_c)):
        w = wb.create_sheet(title)
        w.append(["分類", "サービス", "単位", "R5人", "R6人", "R7人", "R8人", "R5 1人当たり", "R6", "R7", "R8",
                  "町R9人", "町R10人", "町R11人", "町R9量", "町R10量", "町R11量", "点検で止まったこと"])
        for x, per, hit in tk:
            w.append([x["分類"], x["サービス"], x["単位"]] + x["人"]
                     + [round(v, 1) if v is not None else None for v in per]
                     + x["人見込"] + x["量見込"] + ["／".join(hit) or "―"])
        for c in w[1]:
            c.fill = hd
            c.font = Font(bold=True)
    # 年齢による移行
    w = wb.create_sheet("3_年齢による移行")
    w.append(["年齢", "移行前", "移行後", "町に確かめる人数"])
    for row in IKO:
        w.append(list(row))
    # 利用者別台帳（様式）
    w = wb.create_sheet("4_利用者別台帳（様式）")
    w.append(DAICHO)
    for c in w[1]:
        c.fill = hd
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    w.append(["（記入例）A-001", "平成21年度", 17, "療育B", "―", "放課後等デイサービス", "12人日",
              "18歳（R9年度）", "生活介護", "通所を希望（相談記録）", 0, 0, 0, "R9から生活介護に計上"])
    # 財源
    w = wb.create_sheet("5_給付費と財源")
    w.append(["費用", "負担割合", "根拠（障害者総合支援法）", "留意点"])
    for row in zg:
        w.append([row[0], row[1], row[2], row[3].replace("**", "")])
    w.append([])
    w.append(["給付費の推計", "＝見込量（月）×12か月×単価", "単価＝決算額÷実績量（サービスごと、R5〜R7）",
              "【町データ待ち】サービス別の給付費の決算額（R5〜R7）と国庫負担基準の算定額"])
    p = OUT / ("小野町_障がい計画_手帳推計・見込量点検・財源_%s.xlsx" % ASOF)
    wb.save(p)
    return p


def build_docx(j, sk, tk_f, tk_c, zg):
    r = S.Report("小野町　障がい計画　手帳推計・見込量点検・財源")
    r.cover(["おのまち障がい者計画・第8期障がい福祉計画・第4期障がい児福祉計画"],
            ["手帳所持者数の将来推計・見込量の点検・給付費と財源", ASOF_JP,
             "ビズアップ公共コンサルティング株式会社"])
    r.toc_here()
    r.body_here()

    r.h1("1　本書の位置づけ")
    r.p("障がい計画の素案を仕上げるために不足していた4つの作業（手帳所持者数の将来推計、"
        "利用者の積上げによる見込量、月あたりと年あたりの取り違えの点検、給付費と財源）を進めたものです。"
        "**原典は町から受領した見込量・目標値設定シート（令和8年9月7日受領）と、前回計画、"
        "障害者総合支援法の条文です。**")
    r.p("**受領シートには、素案が「手元にない」としていた令和6年の手帳所持者数と、"
        "令和8年の値が入っていました**（身体406人・療育77人・精神47人／身体363人・療育72人・精神60人）。"
        "令和5年と令和7年の値は、前回計画・素案の値と一致することを確かめています。")

    r.h1("2　手帳所持者数の将来推計（#1）")
    r.p("**総人口に対する所持率を用い、2つの方法で推計して幅で示します。** "
        "Aは令和8年の所持率を据え置くもの、Bは令和5〜8年の所持率の趨勢を延ばすものです。"
        "総人口は受領シートの町の推計（令和11年 %s人）によります。" % f"{j['総人口'][6]:,.0f}")
    rows = []
    for nm, (lab, _v) in zip(TECHO3, TECHO_SUII):
        s = sk[nm]
        rows.append([lab, "%g" % s["実績"][3], "%.1f‰" % (s["所持率"][3] * 1000)]
                    + ["%.0f〜%.0f" % (lo, hi) for lo, hi in zip(s["下"], s["上"])]
                    + ["%g" % s["町"][2]])
    r.tbl([["手帳", "R8実績", "R8所持率", "R9（A〜B）", "R10（A〜B）", "R11（A〜B）", "町の推計R11"]] + rows,
          widths=[4.6, 1.8, 2.0, 2.3, 2.3, 2.3, 1.9], right=(1, 2, 3, 4, 5, 6),
          caption="手帳所持者数の将来推計（人）")
    sei = sk["精神障害者保健福祉手帳所持者数"]
    r.p("**精神障害者保健福祉手帳は、所持率の趨勢を延ばすと増え（R11 %.0f人）、"
        "所持率を据え置くと減ります（%.0f人）。** 令和5年の45人から令和8年の60人へ3年で3割余り増えており、"
        "町の推計（%g人）は据え置きに近い置き方です。"
        "**このため、見込量（就労系・相談支援）では"
        "趨勢側（B）を上限として置くことを提案します。**"
        % (sei["B"][2], sei["A"][2], sei["町"][2]))
    r.src("※ 年齢別の所持者数は受領していないため、年齢構成の変化（身体障害者手帳の多くは65歳以上）は"
          "反映していない。年齢別の所持者数を町にお願いする。")

    r.h1("3　見込量の点検（#2・#3）")
    r.p("**受領シートの見込量はすべて「月あたり」です。** 月あたりの利用者数に12を掛けても"
        "年間の実利用者数や請求件数にはなりません（利用の始まりと終わりが年度途中にあるため）。"
        "このため、**1人当たりの量が月の開所日（%d日）を超える行は、年の値が混じっているとみて止めます。**"
        % KAISHO)
    hits = [(x, h) for x, _p, h in tk_f + tk_c if h]
    r.tbl([["サービス", "点検で止まったこと"]] + [[x["サービス"], "／".join(h)] for x, h in hits],
          widths=[4.6, 12.6], caption="見込量の点検で止まったもの")
    over = [x["サービス"] for x, h in hits if any("開所日" in z for z in h)]
    r.p(("**1人当たりの量が月の開所日を超えた行は、%sです。** 月の量に年の値が混じったか、"
         "宿泊を伴う利用を数えた可能性があるため、町に確かめます。" % "・".join(over))
        if over else "**1人当たりの量が月の開所日を超えた行は、ありませんでした。**")
    r.p("ほかに止まったのは、実績の急な変化と、量の見込みが「利用者数×1人当たり量」で"
        "説明できない行です。後者は、町の見込みが量を別の方法（直近の実績の据え置きなど）で"
        "置いているためで、どちらで置くかを町と決める必要があります。")
    r.p("**利用者ごとの積上げ（#2）は、匿名の利用者別台帳がないとできません。** "
        "そのための様式（14列）と、年齢による移行の一覧を作業ファイルに置きました。")
    r.tbl([["年齢", "移行前", "移行後", "町に確かめる人数"]] + [list(x) for x in IKO],
          widths=[2.4, 4.4, 4.4, 6.0], caption="年齢による移行（見込量に反映するもの）")

    r.h1("4　給付費と財源（#4）")
    r.p("**障害者総合支援法の負担の定めは次のとおりです**（e-Gov 法令API、令和9年4月1日時点で確認）。")
    r.tbl([["費用", "負担割合", "根拠（障害者総合支援法）", "留意点"]]
          + [[x[0], x[1], x[2], x[3] or "―"] for x in zg],
          widths=[4.4, 3.2, 4.0, 5.6], caption="障がい福祉の費用の負担")
    r.p("**給付費の推計は「見込量（月）×12か月×単価」で置きます。** 単価はサービスごとに"
        "「決算額÷実績量」で求めます。**サービス別の給付費の決算額（令和5〜7年度）と、"
        "国庫負担基準の算定額を町にお願いします。** 国庫負担基準を超えた分は町の負担になるため、"
        "訪問系の利用が多い年には町の負担が増えます。")

    r.h1("5　町にお願いするもの")
    r.tbl([["№", "資料", "使い道"],
           ["1", "手帳所持者数の年齢別（3手帳、令和8年4月1日）", "推計に年齢構成の変化を入れる"],
           ["2", "匿名の利用者別台帳（様式は作業ファイルの4）", "利用者ごとの積上げと年齢による移行"],
           ["3", "サービス別の給付費の決算額（令和5〜7年度）と国庫負担基準の算定額", "単価と給付費・財源の推計"],
           ["4", "地域生活支援事業の決算額と補助金の交付額（令和5〜7年度）", "補助の実際の割合"]],
          widths=[0.8, 9.6, 6.8], caption="町にお願いするもの")

    p = OUT / ("小野町_障がい計画_手帳推計・見込量点検・財源_%s.docx" % ASOF)
    r.build_toc()
    r.save(p)
    return p


def main():
    j = read_jinko()
    sk = suikei(j)
    tk_f = tenken(read_mikomi("見込量算定(障害福祉)"))
    tk_c = tenken(read_mikomi("見込量算定(障害児)"))
    zg = zaigen()
    ng = jiko(j, None, zg)
    px = build_xlsx(j, sk, tk_f, tk_c, zg)
    pd = build_docx(j, sk, tk_f, tk_c, zg)
    print("■ 障がい計画　手帳推計・見込量点検・財源")
    print("   ", px.relative_to(HERE))
    print("   ", pd.relative_to(HERE))
    print("   点検で止まった行 %d／%d" % (sum(1 for _x, _p, h in tk_f + tk_c if h), len(tk_f + tk_c)))
    print("\n■ 自己点検")
    for x in ng:
        print("   NG", x)
    print("   すべて適合" if not ng else "   **%d件**" % len(ng))
    return 1 if ng else 0


if __name__ == "__main__":
    sys.exit(main())
