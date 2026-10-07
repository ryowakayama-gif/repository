# -*- coding: utf-8 -*-
"""国の推計ワークシート（将来推計総括表）と県の整備量見込みシートを読み取る。

   受領（令和8年10月7日）
     data/国ワークシート/将来推計総括表_R8.10.5.xlsx      見える化システムの出力。保険者番号07402
     data/国ワークシート/施設居住系の整備量見込みシート_R8.10.6.xls  県独自様式
     data/国ワークシート/県通知_8生福第2900号_R8.9.4.pdf
     data/国ワークシート/厚労省事務連絡_第1回推計の集計_R8.9.1.pdf

   村からの申し送り（逐語）
     「県において、第10期介護保険事業計画における介護サービス見込量等の推計（第1回目）等が
       あり、10月5日に提出したデータとなります。なお、福祉用具購入費や住宅改修費は、金額が
       反映されずゼロとなっております。また、新サービスとなる特定地域居宅サービス等事業なども
       検討できていないため、ゼロとなっております。全体的に暫定の数値となりますこと、
       ご承知おきください。」

   このため本モジュールは、読み取った値を**暫定**として扱い、
   ゼロで入っている項目を別に数え上げる。

   python3 scripts/parse_worksheet.py        読み取って CSV に落とす
"""
import csv
import io
import os
import sys
sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths as P

WS = os.path.join(P.DATA, "国ワークシート", "将来推計総括表_R8.10.5.xlsx")
SEIBI = os.path.join(P.DATA, "国ワークシート", "施設居住系の整備量見込みシート_R8.10.6.xls")

# ── 総括表から取り出す値（シート, 行, 列）は版で動くため、見出しで引く ──
def _ws():
    import openpyxl
    return openpyxl.load_workbook(WS, data_only=True)


def _row(ws, lab, col0=1, maxcol=3):
    """左の見出しが lab の行を返す（列1〜3のどれかに一致）"""
    for r in ws.iter_rows(values_only=False):
        for c in r[col0 - 1:maxcol]:
            if c.value is not None and str(c.value).strip() == lab:
                return [x.value for x in r]
    return None


def hokenryo():
    """保険料の推計。(項目 → 第10期の月額) と 3か年の額"""
    wb = _ws()
    s5 = wb["5_保険料推計"]
    out = {}
    for lab in ("保険料基準額（月額）", "準備基金取崩額の影響額", "保険料収納必要額（月額）",
                "総給付費", "その他給付費", "地域支援事業費", "市町村特別給付費等",
                "基準保険料額（月額）"):
        r = _row(s5, lab, 1, 5)
        if r:
            out.setdefault(lab, r[5])
    for lab in ("準備基金の残高（前年度末の見込額）", "準備基金取崩額", "準備基金取崩割合",
                "標準給付費見込額", "第1号被保険者数", "予定保険料収納率",
                "調整交付金見込額", "調整交付金見込交付割合",
                "保険者機能強化推進交付金の交付見込額", "保険料収納必要額"):
        r = _row(s5, lab, 1, 6)
        if r:
            out.setdefault(lab, r[5])
    return out


def nenji():
    """年度別の被保険者数・認定者数・給付費（第10期の3か年）"""
    wb = _ws()
    s1 = wb["1_推計値サマリ"]
    out = {}
    for lab in ("総数", "第1号被保険者数", "第2号被保険者数"):
        r = _row(s1, lab, 1, 3)
        if r:
            out["被保険者_" + lab] = r[6:9]
    # 認定者は2つ目の「総数」。行を順に拾う
    rows = [[c.value for c in r] for r in s1.iter_rows()]
    sousu = [i for i, r in enumerate(rows) if r[0] == "総数"]
    if len(sousu) >= 2:
        out["認定者_総数"] = rows[sousu[1]][6:9]
        for d in ("要支援1", "要支援2", "要介護1", "要介護2",
                  "要介護3", "要介護4", "要介護5"):
            for r in rows[sousu[1]:sousu[1] + 9]:
                if r[2] == d:
                    out["認定者_" + d] = r[6:9]
                    break
    for lab in ("総給付費",):
        for r in rows:
            if r[0] == lab:
                out["給付_" + lab] = r[6:9]
                break
    return out


def seibi():
    """県の整備量見込みシート。サービス → (R6,R7,R8,R9,R10,R11 の利用者数)"""
    import xlrd
    wb = xlrd.open_workbook(SEIBI)
    sh = wb.sheet_by_name("2")
    out = {}
    # 施設の種類だけを鍵にすると、介護専用型と混合型の「有料老人ホーム（特定施設）」が
    # ぶつかって後の行が前の行を消す。サービスの種類と合わせた鍵にする。
    svc = ""
    for i in range(sh.nrows):
        v = sh.row_values(i)
        if len(v) < 9:
            continue
        if str(v[1]).strip():
            svc = " ".join(str(v[1]).split())
        nm = " ".join(str(v[2]).split())
        if not nm or nm == "サービスを提供する施設の種類":
            continue
        try:
            val = [float(x) for x in v[3:9]]
        except (TypeError, ValueError):
            continue
        key = nm if nm in ("特別養護老人ホーム", "介護老人保健施設", "介護医療院") \
            else "%s／%s" % (svc, nm)
        out[key] = val
    # 定員（シート1 の R8末実績見込・R9〜R11）
    sh1 = wb.sheet_by_name("1")
    teiin = {}
    for i in range(sh1.nrows):
        v = sh1.row_values(i)
        nm = " ".join(str(v[2]).split()) if len(v) > 2 else ""
        kb = str(v[3]).strip() if len(v) > 3 else ""
        if nm and kb == "非転換分":
            try:
                teiin[nm] = {"R8末": float(v[10]), "R9": float(v[12]),
                             "R10": float(v[14]), "R11": float(v[16])}
            except (TypeError, ValueError):
                pass
    return out, teiin


def zero_koumoku():
    """ゼロで入っている項目を数え上げる（村の申し送りによる暫定の箇所）"""
    wb = _ws()
    s2 = wb["2_サービス別給付費"]
    nashi = []
    rows = [[c.value for c in r] for r in s2.iter_rows()]
    for i, r in enumerate(rows):
        nm = r[1]
        if not nm or str(nm).strip() in ("", "未使用"):
            continue
        if r[2] != "給付費（千円）":
            continue
        r6, r7, r8 = r[3], r[4], r[5]
        r9 = r[6]
        def f(x):
            try:
                return float(x)
            except (TypeError, ValueError):
                return None
        a, b, c8, d9 = f(r6), f(r7), f(r8), f(r9)
        # 令和6・7年度に実績があるのに令和8年度が0のもの＝月報の月数が足りない恐れ
        if (a or b) and c8 == 0:
            nashi.append((str(nm).strip(), a, b, c8, d9))
    return nashi


def jinko():
    """第1号被保険者数の年齢区分別。前期・後期・後期のうち85歳以上まで入っている。

       国の様式は3区分（前期65〜74歳／後期75〜84歳／後期85歳以上）で持つ。
       調整交付金の後期高齢者加入割合補正係数はこの3区分から算出される
       （「(参考)保険料の推計に要する係数」の全国値も3区分である）。
       素案5-2は前期・後期の2区分で組んでいるため、合計が合っても内訳が食い違う。
    """
    wb = _ws()
    s5 = wb["5_保険料推計"]
    rows = [[c.value for c in r] for r in s5.iter_rows()]
    st = 0
    for i, r in enumerate(rows):
        if r and r[0] and str(r[0]).startswith("６．第１号被保険者数関係"):
            st = i
            break
    out = {}
    for lab in ("第1号被保険者数", "前期(65～74歳)", "後期(75歳～)",
                "後期(75歳～84歳)", "後期(85歳～)"):
        for r in rows[st:st + 12]:
            for c in r[:3]:
                if c is not None and str(c).strip() == lab:
                    out[lab] = [x for x in r[6:9]]   # G・H・I列＝令和9〜11年度
                    break
            if lab in out:
                break
    return out


def keisu():
    """保険料に効く係数。国が与えるものと保険者が入れるものが混ざっている"""
    wb = _ws()
    s5 = wb["5_保険料推計"]
    sk = wb["(参考)保険料の推計に要する係数"]
    rows = [[c.value for c in r] for r in s5.iter_rows()]
    st = 0
    for i, r in enumerate(rows):
        if r and r[0] and str(r[0]).startswith("５．保険料収納必要額関係"):
            st = i
            break

    def pick(lab, n=3):
        for r in rows[st:]:
            for c in r[:6]:
                if c is not None and str(c).strip() == lab:
                    return r[6:6 + n]
        return None
    out = {
        "調整交付金見込交付割合": pick("調整交付金見込交付割合"),
        "後期高齢者加入割合補正係数": pick("後期高齢者加入割合補正係数"),
        "所得段階別加入割合補正係数": pick("所得段階別加入割合補正係数"),
        # 収納率は期ごとに1つ（F列が第10期、J列以降が第11期以降）
        "予定保険料収納率": [pick("予定保険料収納率", 1)[0]] if pick("予定保険料収納率", 1) else None,
    }
    for r in rows[st:]:
        if r and r[0] is not None and str(r[0]).strip() == "予定保険料収納率":
            out["予定保険料収納率"] = [r[5]]
            break
    # 第1号被保険者負担割合は参考シートに期別で入っている（第10期〜第17期）
    for r in sk.iter_rows():
        v = [x.value for x in r]
        for c in v[:4]:
            if c is not None and "第１号被保険者負担割合" in str(c):
                out["第1号被保険者負担割合"] = v[3:9]
                break
    return out


# 素案5-2のB案（前期・後期の2区分。各年の年末）
TOUHOU_JINKO = {"前期": (465, 444, 424), "後期": (538, 551, 563)}


def main():
    h = hokenryo()
    print("■ 国の推計ワークシート（令和8年10月5日提出・暫定）")
    print("  第10期 保険料基準額（月額）      %10.2f 円" % h["保険料基準額（月額）"])
    print("  　保険料収納必要額（月額）        %10.2f 円" % h["保険料収納必要額（月額）"])
    print("  　準備基金取崩額の影響額          %10.2f 円" % h["準備基金取崩額の影響額"])
    print("  準備基金 残高 %s円 → 取崩 %s円（%.1f%%）"
          % ("{:,.0f}".format(h["準備基金の残高（前年度末の見込額）"]),
             "{:,.0f}".format(h["準備基金取崩額"]),
             h["準備基金取崩割合"] * 100))
    print("  標準給付費見込額（3か年計）      %14s 円" % "{:,.0f}".format(h["標準給付費見込額"]))
    print("  予定保険料収納率                 %10.2f%%" % (h["予定保険料収納率"] * 100))
    print("  保険者機能強化推進交付金（3か年） %13s 円"
          % "{:,.0f}".format(h["保険者機能強化推進交付金の交付見込額"]))
    print()
    n = nenji()
    print("■ 第10期の年度別（令和9・10・11年度）")
    for k in ("被保険者_第1号被保険者数", "認定者_総数", "給付_総給付費"):
        print("  %-22s %s" % (k, " ".join("%10s" % ("{:,.0f}".format(x) if x else "-")
                                          for x in n[k])))
    j = jinko()
    kk = keisu()
    print()
    print("■ 第1号被保険者数の年齢区分別（国の様式は3区分）")
    for lab in ("第1号被保険者数", "前期(65～74歳)", "後期(75歳～)",
                "後期(75歳～84歳)", "後期(85歳～)"):
        if j.get(lab):
            print("  %-20s %s" % (lab, " ".join("%6.0f" % (x or 0) for x in j[lab])))
    tj = TOUHOU_JINKO
    print("  ── 素案5-2（B案・2区分）との内訳の差 ──")
    print("  %-20s %s" % ("前期 素案", " ".join("%6d" % x for x in tj["前期"])))
    print("  %-20s %s" % ("後期 素案", " ".join("%6d" % x for x in tj["後期"])))
    kz = [(j["後期(75歳～)"][i] or 0) / (j["第1号被保険者数"][i] or 1) * 100
          for i in range(3)]
    sz = [tj["後期"][i] / (tj["前期"][i] + tj["後期"][i]) * 100 for i in range(3)]
    print("  %-20s %s" % ("後期の割合 国", " ".join("%5.1f%%" % x for x in kz)))
    print("  %-20s %s" % ("後期の割合 素案", " ".join("%5.1f%%" % x for x in sz)))
    print()
    print("■ 保険料に効く係数")
    for lab, v in kk.items():
        if v:
            print("  %-26s %s" % (lab, " ".join(
                ("%8.4f" % x) if isinstance(x, float) else "%8s" % x for x in v)))
    print()
    riyou, teiin = seibi()
    print("■ 県の整備量見込みシート（令和8年10月6日提出）　利用者数")
    print("  %-28s %s" % ("サービス", "  ".join("%5s" % x for x in
                                                ("R6", "R7", "R8", "R9", "R10", "R11"))))
    for k, v in riyou.items():
        if any(v):
            print("  %-28s %s" % (k[:28], "  ".join("%5.0f" % x for x in v)))
    print()
    print("  定員（非転換分）")
    for k, v in teiin.items():
        if any(v.values()):
            print("  %-28s R8末%3.0f → R9%3.0f R10%3.0f R11%3.0f"
                  % (k[:28], v["R8末"], v["R9"], v["R10"], v["R11"]))
    print()
    z = zero_koumoku()
    print("■ 令和6・7年度に実績があるのに令和8年度が0のサービス（%d件）" % len(z))
    print("  厚労省事務連絡 3(2)「令和8年度の月報の月数が少ないため過小になる恐れ」に当たる")
    for nm, a, b, c8, d9 in z:
        print("  %-30s R6 %8.1f  R7 %8.1f  R8 %4.0f  → R9見込 %8s"
              % (nm[:30], a or 0, b or 0, c8, "{:,.0f}".format(d9) if d9 else "0"))
    return 0


if __name__ == "__main__":
    sys.exit(main())


# ══════════════════════════════════════════════════════════════
# 当方の算定との突合（厚労省事務連絡 3(3)(4) が求める確認）
# ══════════════════════════════════════════════════════════════
#   事務連絡は「保険料が第9期と比べて大きく増減している場合には、誤入力等がないか、
#   推計データ全体を確認してください」(3(3))、「実績と推計の経年での変化の状況、
#   伸び率に大きな増減がないか」(3(4)) を求めている。
#   本村は第9期比 ▲18.45% であり、これに当たる。

# 当方の算定（素案5-6・5-7）。素案を直したらここも直す。点検で縛る。
TOUHOU = {
    "第1号被保険者数": (1004, 995, 987),
    "認定者数": (214, 217, 220),
    "サービス諸費": (263470, 263505, 262488),
    "標準給付費見込額": (284987, 285022, 284005),
    "特定入所者介護サービス費等": (14496, 14496, 14496),
    "高額介護サービス費等": (6769, 6769, 6769),
    "高額医療合算介護サービス費等": (0, 0, 0),
    "算定対象審査支払手数料": (252, 252, 252),
    "地域支援事業費": (39886, 39886, 39886),
}
TOUHOU_HOKEN = 6756          # 取崩なしの3か年平均（素案5-7）
TOUHOU_CHOSEI = (4.10, 4.32)  # 調整交付金の見込交付割合（第9期から復元）


def kuni_nenji():
    """国のワークシートから、突合に使う年度別の値を取り出す"""
    wb = _ws()
    s5 = wb["5_保険料推計"]
    rows = [[c.value for c in r] for r in s5.iter_rows()]

    # 「５．保険料収納必要額関係」より後の行から拾う。
    # 同じ見出しが内訳（月額）の節にもあり、そちらは単位が円／月で意味が違う。
    st = 0
    for i, r in enumerate(rows):
        if r and r[0] and str(r[0]).startswith("５．保険料収納必要額関係"):
            st = i
            break

    def pick(lab, col=6):
        for r in rows[st:]:
            for c in r[:6]:
                if c is not None and str(c).strip() == lab:
                    return r[col:col + 3]
        return None
    out = {
        "標準給付費見込額": pick("標準給付費見込額"),
        "特定入所者介護サービス費等": pick("特定入所者介護サービス費等給付額"),
        "高額介護サービス費等": pick("高額介護サービス費等給付額"),
        "高額医療合算介護サービス費等": pick("高額医療合算介護サービス費等給付額"),
        "算定対象審査支払手数料": pick("算定対象審査支払手数料"),
        "地域支援事業費": pick("地域支援事業費"),
        "調整交付金見込交付割合": pick("調整交付金見込交付割合"),
        "第1号被保険者数": pick("第1号被保険者数"),
    }
    n = nenji()
    out["認定者数"] = n["認定者_総数"]
    out["サービス諸費"] = n["給付_総給付費"]
    return out


def totsugou():
    """当方の算定と国のワークシートを並べる"""
    k = kuni_nenji()
    out = []
    for nm, t in TOUHOU.items():
        kk = k.get(nm)
        if not kk:
            out.append((nm, sum(t), None, None))
            continue
        kv = [float(x) if x is not None else 0.0 for x in kk]
        # 国のワークシートは「５．保険料収納必要額関係」の節を円で持つ。
        # 当方の素案は千円で持つため、人数以外は千円に揃える。
        # サービス諸費（総給付費）だけはサマリの節から取るため、もともと千円である。
        EN = ("標準給付費見込額", "特定入所者介護サービス費等", "高額介護サービス費等",
              "高額医療合算介護サービス費等", "算定対象審査支払手数料", "地域支援事業費")
        if nm in EN:
            kv = [x / 1000 for x in kv]
        a, b = sum(t), sum(kv)
        out.append((nm, a, b, ((b - a) / a * 100) if a else 0.0))
    return out
