# -*- coding: utf-8 -*-
"""令和8年度 交付金 該当状況調査票集計表（市町村分・全国一覧表）の検算と県内比較。

   受領（2026/09/25）
     001732614.xlsx  令和8年度 全国集計（市町村）1,779行（該当1,741）×1,563列
     001732616.pdf   令和8年度 評価指標（市町村分）全42頁

   本村は通し番号380（シート上は第398行＝見出し18行＋380）。
   既収録の doc20 と照合したうえで、従前できなかった
   「目標別の県内比較」（福島県59市町村の平均・中央値・本村の県内順位）を算出する。
"""
import csv, os, sys

SRC = "/root/.claude/uploads/134138ca-61f7-57d3-9e9b-5f081a1a345d/78f982ca-001732614.xlsx"
OUT = "/home/user/repository/docs/北塩原村_第10期/data/交付金_目標別の県内比較_令和8年度.csv"
HROWS = 18          # 見出し行数（9〜17行に配点・全国合計・平均点・項目平均・平均得点率・中央値・標準偏差・該当市町村数・該当率）
FUKUSHIMA = 7       # 都道府県番号
MURA = "北塩原村"

# 合計列（0起点）。見出しを縦に連結して同定した
COL = {"推進Ⅰ": 56, "推進Ⅱ": 78, "推進Ⅲ": 107, "推進Ⅳ": 140, "推進合計": 141,
       "支援Ⅰ": 237, "支援Ⅱ": 266, "支援Ⅲ": 300, "支援Ⅳ": 333, "支援合計": 334,
       "総合計": 335, "今年度順位": 336, "前年度得点": 12, "前年度順位": 13}
NAME = {"推進Ⅰ": "持続可能な地域のあるべき姿をかたちにする",
        "推進Ⅱ": "公正・公平な給付を行う体制を構築する",
        "推進Ⅲ": "介護人材の確保その他のサービス提供基盤の整備を推進する",
        "推進Ⅳ": "高齢者が可能な限り自立した日常生活を営む（成果指標群）",
        "支援Ⅰ": "介護予防／日常生活支援を推進する",
        "支援Ⅱ": "認知症総合支援を推進する",
        "支援Ⅲ": "在宅医療・在宅介護連携の体制を構築する",
        "支援Ⅳ": "高齢者が可能な限り自立した日常生活を営む（成果指標群）"}
KEYS = ["推進Ⅰ", "推進Ⅱ", "推進Ⅲ", "推進Ⅳ", "推進合計",
        "支援Ⅰ", "支援Ⅱ", "支援Ⅲ", "支援Ⅳ", "支援合計", "総合計"]


def load():
    from openpyxl import load_workbook
    wb = load_workbook(SRC, read_only=True, data_only=True)
    rows = list(wb["全国集計（市町村）"].iter_rows(values_only=True))
    H, body = rows[:HROWS], rows[HROWS:]
    # 総合計が数値の行のみを集計対象とする
    tc = COL["総合計"]
    data = [r for r in body if r[0] is not None and isinstance(r[tc], (int, float))]
    return H, data


# 見出し行に埋め込まれた全国統計の行位置（0起点）。合計列は項目平均・該当数が空欄になる
SROW = {"配点": 8, "全国合計": 9, "平均点": 10, "項目平均": 11, "平均得点率": 12,
        "中央値": 13, "標準偏差": 14, "該当市町村数": 15, "該当率": 16}

def stats(H, c):
    """見出し列に埋め込まれた全国統計を名前で引ける辞書にする"""
    out = {}
    for k, i in SROW.items():
        v = H[i][c] if c < len(H[i]) else None
        out[k] = float(v) if isinstance(v, (int, float)) else None
    return out


def rank(vals, v):
    """降順の順位（同点は上位の数＋1）"""
    return sum(1 for x in vals if x > v) + 1


def main():
    H, data = load()
    tc = COL["総合計"]
    mura = [r for r in data if r[7] == MURA]
    assert len(mura) == 1, f"北塩原村の行が{len(mura)}件"
    m = mura[0]
    ken = [r for r in data if r[1] == FUKUSHIMA]
    zen = [r[tc] for r in data]

    print("■ 検算（既収録の doc20 との照合）\n")
    st = stats(H, tc)
    n_zen = st["全国合計"] / st["平均点"]   # 該当市町村数の検算
    ck = [
        ("本村の通し番号", int(m[0]), 380),
        ("シート上の行番号", HROWS + int(m[0]), 398),
        ("本村の令和8年度 合計得点", m[tc], 447),
        ("全国の集計対象市町村数", round(n_zen), 1741),
        ("全国平均", round(st["平均点"], 1), 455.1),
        ("本村の全国順位", rank(zen, m[tc]), 1005),
        ("福島県の市町村数", len(ken), 59),
        ("福島県平均", round(sum(r[tc] for r in ken) / len(ken), 1), 456.4),
        ("本村の県内順位", rank([r[tc] for r in ken], m[tc]), 34),
    ]
    ng = 0
    for nm, got, exp in ck:
        ok = str(got) == str(exp)
        print(f"  {'○' if ok else '✕'} {nm:26s} 今回={got}　既収録={exp}")
        ng += 0 if ok else 1
    print(f"\n  {len(ck)}件のうち一致{len(ck)-ng}件／不一致{ng}件")

    # 「著しく得点の低い市町村」の閾値（都道府県分 目標Ⅰ-（ⅰ）-1オ の判定基準）
    rate_avg, sd, hai = st["平均得点率"], st["標準偏差"], st["配点"]
    th = (rate_avg - 2 * (sd / hai)) * hai
    low = [r for r in data if r[tc] < th]
    print(f"\n■ 「著しく得点の低い市町村」の閾値\n")
    print(f"  得点率の平均 {rate_avg:.6f} − 標準偏差（{sd:.2f}点＝得点率 {sd/hai:.6f}）の2倍"
          f" → 得点率 {(rate_avg - 2*sd/hai):.6f} ＝ {th:.1f}点")
    print(f"  全国で下回るのは {len(low)}市町村。福島県では "
          f"{len([r for r in ken if r[tc] < th])}市町村。"
          f"本村は{m[tc]}点で閾値を{m[tc]-th:.1f}点上回る")

    print("\n\n■ 目標別の県内比較（従前は合計のみ。今回の一覧表で目標別に算出できた）\n")
    print(f"{'区分':6s} {'配点':>4s} {'本村':>5s} │ {'県平均':>7s} {'県中位':>6s} {'県内順位':>8s} │ "
          f"{'全国平均':>8s} {'全国順位':>9s}")
    print("─" * 82)
    recs = []
    for k in KEYS:
        c = COL[k]
        s = stats(H, c)
        kv = [r[c] for r in ken]
        zv = [r[c] for r in data]
        kr, zr = rank(kv, m[c]), rank(zv, m[c])
        print(f"{k:6s} {s['配点']:4.0f} {m[c]:5.0f} │ {sum(kv)/len(kv):7.1f} "
              f"{sorted(kv)[len(kv)//2]:6.0f} {kr:4d}位/{len(kv):<3d} │ "
              f"{s['平均点']:8.1f} {zr:5d}位/{len(zv):<4d}")
        recs.append({"区分": k, "目標名": NAME.get(k, ""), "配点": s["配点"], "本村": m[c],
                     "県平均": round(sum(kv) / len(kv), 1),
                     "県中央値": sorted(kv)[len(kv) // 2], "県内順位": kr, "県内市町村数": len(kv),
                     "全国平均": round(s["平均点"], 1), "全国順位": zr, "全国市町村数": len(zv),
                     "県平均との差": round(m[c] - sum(kv) / len(kv), 1),
                     "全国平均との差": round(m[c] - s["平均点"], 1)})
    print("─" * 82)

    print("\n■ 県平均と全国平均のどちらを下回るか\n")
    for r in recs:
        if r["区分"].endswith("合計"):
            continue
        kd, zd = r["県平均との差"], r["全国平均との差"]
        if kd >= 0 and zd >= 0:
            j = "県・全国いずれも上回る"
        elif kd < 0 and zd < 0:
            j = "★県・全国いずれも下回る"
        elif zd < 0:
            j = "全国は下回るが県内では平均以上（県全体が低い領域）"
        else:
            j = "全国は上回るが県内では平均以下（県内の水準が高い領域）"
        print(f"  {r['区分']:6s} {r['目標名'][:30]:32s} 県{kd:+6.1f} 全国{zd:+6.1f}  {j}")

    toreru(H, m, ken, data)

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8-sig", newline="") as fp:
        w = csv.DictWriter(fp, fieldnames=list(recs[0].keys()))
        w.writeheader(); w.writerows(recs)
    print(f"\n保存: {OUT}")
    return 1 if ng else 0


# ── 未得点項目のうち、計画への記載・要綱の整備で取れるもの ──────────
# 評価指標（市町村分）令和8年度 全42頁の留意点を読み、取り方を判定した。
# 列は見出しの文言で引き当てる（直書きの列番号は取り違えやすい）。
# (交付金, 目標, 指標名, 枝, 内容, 取り方)　「取り方」は当方の判断で、村への確認を要するものは★を付す。
TORERU = [
 ("支援Ⅱ", "認知症サポーター等を活用した地域支援体制の構築", "カ",
  "認知症の人及び家族等の意見を踏まえた市町村認知症施策推進計画の策定に着手",
  "素案4-7が既に条件を満たす。意見を反映させる仕組みを整備した上で策定に着手していればよく、策定済みでも対象"),
 ("支援Ⅱ", "認知症サポーター等を活用した地域支援体制の構築", "イ",
  "成年後見制度利用支援事業の対象を市町村長申立・生活保護受給者に限定しない要綱等の整備",
  "要綱の改正のみ。申立費用助成・報酬助成は整備済（doc28§3）"),
 ("支援Ⅱ", "難聴高齢者の早期発見・早期介入", "ア", "普及啓発の取組",
  "★広報・健診案内での周知で足りるか要確認。施策4-4に位置づける"),
 ("支援Ⅱ", "難聴高齢者の早期発見・早期介入", "イ", "早期発見の取組",
  "★健診時の聴力確認等。施策1-6（保健事業と介護予防の一体的実施）と連動させる"),
 ("支援Ⅲ", "在宅医療・介護連携に関する課題・対応策の検討", "ア",
  "①日常の療養支援②入退院支援③急変時の対応④看取りの4場面ごとに提供体制の目指すべき姿を設定",
  "イ（計画への記載・6点）は取得済。4場面ごとの整理がないため0点。素案4-3に4場面の表を置けば取れる"),
 ("支援Ⅲ", "在宅医療・介護連携に関する課題・対応策の検討", "エ",
  "アとウの差から抽出した課題を踏まえた目標の設定・具体的な対応策の立案",
  "アが前提。ウ（定量的把握・5点）は取得済のため、アを整理すれば同時に立案できる"),
 ("支援Ⅲ", "在宅医療・介護連携の具体的取組状況", "イ②",
  "定期的な相談内容等の取りまとめと医療・介護関係者間での共有",
  "①相談窓口・③多職種研修は取得済。取りまとめと共有の手順を定める"),
 ("推進Ⅱ", "給付費適正化事業の取組状況", "ウ",
  "ケアプラン点検に有料老人ホーム・サービス付き高齢者向け住宅の入居者分を含める",
  "★管内に該当施設がない市町村も、被保険者が他市町村の施設に入居して適正に利用しているかの実態把握と"
  "県・他市町村との連携体制があれば対象（留意点）。本村は給付の58.7%が村外であり取り組む意義が大きい"),
 ("推進Ⅱ", "給付費適正化事業の取組状況", "エ",
  "福祉用具の貸与後にリハビリテーション専門職等が適切な利用を点検する仕組み",
  "関係団体・県・近隣市町村の広域団体と連携した仕組みでも対象（留意点）"),
 ("推進Ⅱ", "給付費適正化事業の取組状況", "オ",
  "福祉用具購入費・住宅改修費の申請内容の妥当性をリハビリテーション専門職等が検討する仕組み",
  "購入費・改修費のいずれかでよく、住宅改修費では建築専門職・福祉住環境コーディネーター2級以上も"
  "「リハビリテーション専門職等」に含む（留意点）"),
 ("推進Ⅱ", "給付費適正化事業の取組状況", "ア", "主要3事業（要介護認定の適正化・ケアプラン等の点検・縦覧点検／医療情報との突合）の全てを実施",
  "★方策の策定（32点）は満点、ケアプラン点検・突合の活動指標も満点。要介護認定の適正化の実施状況を要確認"),
]

def levels(H, n):
    """結合セルは左上以外が None になるため、各見出し行の値を横に持ち越す"""
    out = []
    for i in range(8):
        cur, row = None, []
        for c in range(n):
            v = H[i][c] if c < len(H[i]) else None
            if v is not None and str(v).strip() not in ("", "\u3000"):
                cur = str(v).strip()
            row.append(cur)
        out.append(row)
    return out


def find_col(LV, H, n, kf, shihyo, eda):
    """交付金・指標名・枝から列を1つだけ引き当てる"""
    kfn = {"推進": "保険者機能強化推進交付金", "支援": "介護保険保険者努力支援交付金"}[kf[:2]]
    mk = kf[2:]
    hit = []
    for c in range(n):
        if LV[1][c] != kfn or not LV[2][c] or not LV[2][c].startswith(f"目標{mk}"):
            continue
        if LV[5][c] != shihyo or LV[6][c] != eda[0]:
            continue
        if isinstance(H[SROW["配点"]][c] if c < len(H[8]) else None, (int, float)):
            hit.append(c)
    # 同じ枝に小項目（イ①・イ②…）がある場合は8行目で絞る。
    # 8行目は結合セルの持ち越しで枝のない指標にも値が入るため、絞込みは複数該当時のみ行う
    if len(hit) > 1 and len(eda) > 1:
        hit = [c for c in hit if LV[7][c] == eda[1:]]
    assert len(hit) == 1, f"{kf}／{shihyo}／{eda} が{len(hit)}件"
    return hit[0]


def toreru(H, m, ken, data):
    print("\n\n■ 未得点のうち計画への記載・要綱の整備で取れるもの\n")
    print(f"{'交付金':6s} {'配点':>4s} {'全国該当率':>9s} {'県内順位':>9s}  指標")
    print("─" * 100)
    n = len(H[1])
    LV = levels(H, n)
    tot = 0
    for kf, shihyo, eda, naiyo, how in TORERU:
        c = find_col(LV, H, n, kf, shihyo, eda)
        nm = f"{shihyo} {eda}"
        hai = H[SROW["配点"]][c]
        rate = H[SROW["該当率"]][c]
        assert m[c] == 0, f"{nm} は0点ではない（{m[c]}）"
        kv = [r[c] for r in ken if isinstance(r[c], (int, float))]
        kr = sum(1 for x in kv if x > 0) + 1
        tot += hai
        print(f"{kf:6s} {hai:4.0f} {rate:9.1%} {kr:4d}位/{len(kv):<4d}  {nm}")
        print(f"{'':32s}  {naiyo}")
        print(f"{'':32s}  → {how}")
    tc = COL["総合計"]
    now = m[tc]
    zen_avg = H[SROW["平均点"]][tc]
    print("─" * 100)
    print(f"  合計 {tot:.0f}点。現在の{now:.0f}点が {now+tot:.0f}点となり、"
          f"全国平均{zen_avg:.1f}点を{now+tot-zen_avg:.1f}点上回る")
    zv = sorted((r[tc] for r in data), reverse=True)
    r_now = sum(1 for x in zv if x > now) + 1
    r_new = sum(1 for x in zv if x > now + tot) + 1
    kv = [r[tc] for r in ken]
    print(f"  全国順位 {r_now}位 → {r_new}位／{len(zv)}　"
          f"県内順位 {sum(1 for x in kv if x > now)+1}位 → {sum(1 for x in kv if x > now+tot)+1}位／{len(kv)}")
    return tot


if __name__ == "__main__":
    sys.exit(main())
