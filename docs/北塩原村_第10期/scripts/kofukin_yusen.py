# -*- coding: utf-8 -*-
"""保険者機能強化推進交付金等　獲得状況の整理と取得の優先順位（WBS Ⅱ-67・Ⅱ-113）

【なぜ必要か】
  これまでの分析（doc20・doc42・analyze_kofukin.py）は、0点で全国該当率45%以上の
  32項目に絞って対応可否を判定していた。本ファイルは、
  **未取得の点をすべて並べ、取得の見込みで段に分けて優先順位を付ける**。
  あわせて、県内59市町村のうち何団体が取得しているかを項目ごとに示す。

【明細の列と小計の列を分けること】
  集計表の見出しは、明細（指標の枝番）と小計（指標群計・目標計・交付金計）が
  同じ行に並ぶ。**小計の列を明細と数えると配点が2倍になる。**
  現に、目標Ⅱの配点計は明細だけなら100点だが、指標群の小計
  （（ⅰ）体制・取組68点・（ⅱ）活動32点）を含めると200点になる。
  明細の列は「枝番（ア・イ…）か孫番（①②…）が生の値として入っている列」とする。

【項目の配点はそのまま得点になる】
  小計の列を明細に混ぜたまま数えると、明細の合計が目標の合計列のちょうど2倍に
  なるため、「配点は2分の1に換算される」と読めてしまう。そうではない。
  **明細の合計は目標の合計列に等しい**（全国1,741団体で確かめた・check()）。
  ただし名目上の上限はある。支援の目標Ⅰは（ⅱ）活動指標群の明細の配点計が57点で
  小計の配点が48点、目標Ⅳは明細の配点計が160点で満点が100点である。
  いずれも**上限に達している団体は全国に1つもない**（最大42点・100点）。

【県内の他団体の扱い】
  集計表は全国1,741市町村の明細を含むが、**成果品には団体名も個別の数値も掲げない。**
  用いるのは「福島県59市町村のうち何団体が取得しているか」という集計値だけである
  （doc09・verify_mieruka と同じ扱い）。

【取得の見込みの段】
  Ａ　記載で取れる　　計画への記載・要綱の整備・手順の明文化で要件を満たす
  Ｂ　本業務で取れる　本業務の調査・分析の実施that要件を満たす
  Ｃ　事業の実施が前提　事業の実施・実績の水準・他機関との連携が要る
  Ｄ　取れない　　　　国の統計から自動算定される、または他指標の結果として決まる
  Ｅ　未判定　　　　　個別の判定をしていない（全国該当率45%未満の項目が主）

  python3 scripts/kofukin_yusen.py          整理と優先順位を示す
  python3 scripts/kofukin_yusen.py --csv    data/交付金_優先順位.csv に書き出す
  python3 scripts/kofukin_yusen.py --check  換算の規則が全国で成り立つことを確かめる
"""
import collections
import csv
import io
import os
import sys
sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths as P
import analyze_kofukin as AK

SRC = "/root/.claude/uploads/134138ca-61f7-57d3-9e9b-5f081a1a345d/78f982ca-001732614.xlsx"
OUT = os.path.join(P.DATA, "交付金_優先順位.csv")
HROWS = 18
FUKUSHIMA = 7
MURA = "北塩原村"
KIN = {"保険者機能強化推進交付金": "推進", "介護保険保険者努力支援交付金": "支援"}
# 目標ごとの合計列（0起点）
TOT = {("推進", "目標Ⅰ"): 56, ("推進", "目標Ⅱ"): 78,
       ("推進", "目標Ⅲ"): 107, ("推進", "目標Ⅳ"): 140,
       ("支援", "目標Ⅰ"): 237, ("支援", "目標Ⅱ"): 266,
       ("支援", "目標Ⅲ"): 300, ("支援", "目標Ⅳ"): 333}
MANTEN = 100            # 目標ごとの満点
NAME = {"目標Ⅰ": {"推進": "持続可能な地域のあるべき姿をかたちにする",
                   "支援": "介護予防／日常生活支援を推進する"},
        "目標Ⅱ": {"推進": "公正・公平な給付を行う体制を構築する",
                   "支援": "認知症総合支援を推進する"},
        "目標Ⅲ": {"推進": "介護人材の確保その他のサービス提供基盤の整備を推進する",
                   "支援": "在宅医療・在宅介護連携の体制を構築する"},
        "目標Ⅳ": {"推進": "高齢者が可能な限り自立した日常生活を営む（成果指標群）",
                   "支援": "高齢者が可能な限り自立した日常生活を営む（成果指標群）"}}
# 取得の見込みの段。analyze_kofukin.H の判定の文字をこの段に移す
DAN = {"可（記載）": "Ａ", "可（本業務）": "Ｂ", "条件付き可": "Ｃ", "不可": "Ｄ"}
DAN_NA = {"Ａ": "記載で取れる", "Ｂ": "本業務で取れる", "Ｃ": "事業の実施が前提",
          "Ｄ": "取れない", "Ｅ": "未判定"}
# 期待得点の係数。Ｃは事業の実施を伴うため半分に見る
KEISU = {"Ａ": 1.0, "Ｂ": 1.0, "Ｃ": 0.5, "Ｄ": 0.0, "Ｅ": 0.0}

# analyze_kofukin.H は「0点かつ全国該当率45%以上」の32項目に絞った判定表である。
# その外にも、令和8年度の評価指標（市町村分）の原典を読んで判定できたものがある。
# ここに足す。配点・該当率は集計表から引くため書かない。
H2 = {
 ('支援', '目標Ⅱ', '1 / 認知症サポーター等を活用した地域支援体制の構築 / カ'):
   ('可（本業務）',
    '認知症の人および家族等の意見を踏まえた市町村認知症施策推進計画の策定に着手していること。'
    '原典の留意点は「認知症の人と家族等が参画し、意見を反映させる仕組みを整備した上で'
    '計画の策定に着手している場合に評価の対象とする」としている。'
    '本計画が同計画を兼ね（1-2）、意見を聴く方法を4-7(5)に定めている。'
    '仕組みを実際に動かすことが前提となる',
    '4-7・4-7(5)', '―'),
 ('推進', '目標Ⅲ', '2 / 庁内・庁外における連携体制 / エ'):
   ('条件付き可',
    'ア・イの連携体制を活用した「高齢者の住まいの確保と生活の一体的支援」の取組の実施。'
    '本村はア（庁内）・イ（庁外）・ウ①〜⑤をすべて取得済みであり、土台はある。'
    '県の高齢者居住安定確保計画の会津地域の対策（除雪等の支援と集住の促進）に即して'
    '施策2-3・施策2-6を組み立てている。原典の留意点が「取組が着実に進められているか」を'
    '見るとしているため、記載だけでは足りない',
    '施策2-3・施策2-6・1-4', '県計画の改定の予定'),
}


def _rows():
    from openpyxl import load_workbook
    wb = load_workbook(SRC, read_only=True, data_only=True)
    return list(wb["全国集計（市町村）"].iter_rows(values_only=True))


def _ff(rows, ri, n):
    """見出しの横の結合を前方補完する"""
    out, cur = [], None
    for c in range(n):
        v = rows[ri][c]
        if v is not None and str(v).strip():
            cur = str(v).strip()
        out.append(cur)
    return out


def load():
    """(明細の一覧, 本村の行, 県内の行, 全国の行) を返す。

       明細 = (列, 交付金, 目標, 指標群, 指標, 配点)
    """
    rows = _rows()
    n = len(rows[0])
    # 1〜6行は横に結合されているため前方補完する。7行（孫番）は列ごとの生の値で、
    # 補完すると枝番だけの列に前の孫番が付いてしまう（指標名が合わなくなる）。
    L = {i: _ff(rows, i, n) for i in (1, 2, 3, 4, 5, 6)}
    L[7] = [rows[7][c] for c in range(n)]
    det = []
    hai = rows[8]
    for c in range(n):
        if not isinstance(hai[c], (int, float)):
            continue
        k = KIN.get(L[1][c] or "")
        if not k or not (L[2][c] or "").startswith("目標"):
            continue
        # 明細の列は、枝番か孫番が生の値として入っている列。
        # どちらも空の列は小計（指標群計・目標計・交付金計）である。
        nama6 = rows[6][c]
        nama7 = L[7][c]
        if not (nama6 and str(nama6).strip()) and not (nama7 and str(nama7).strip()):
            continue
        si = " / ".join(str(x).strip() for x in (L[4][c], L[5][c], L[6][c], L[7][c])
                        if x is not None and str(x).strip())
        det.append((c, k, (L[2][c] or "").split("　")[0], L[3][c] or "", si,
                    float(hai[c])))
    body = [r for r in rows[HROWS:]
            if r[0] is not None and isinstance(r[335], (int, float))]
    mura = [r for r in body if r[7] == MURA]
    assert len(mura) == 1, "本村の行が%d件" % len(mura)
    ken = [r for r in body if r[1] == FUKUSHIMA]
    assert len(ken) == 59, "福島県の市町村が%d件" % len(ken)
    return det, mura[0], ken, body


def _v(r, c):
    return float(r[c]) if isinstance(r[c], (int, float)) else 0.0


def check(det, body):
    """明細の切り出しが正しいこと。

       ① 全国のすべての団体で、明細の合計が目標の合計列に等しい
          （小計の列を混ぜると2倍になるため、これが切り出しの検算になる）
       ② 明細の素点が目標の満点を超える団体がないこと
    """
    item = collections.defaultdict(list)
    hai = {}
    for c, k, mok, _g, _s, h in det:
        item[(k, mok)].append(c)
        hai[c] = h
    bad, memo = [], []
    for key, cols in sorted(item.items()):
        h = sum(hai[c] for c in cols)
        mx, ng = 0.0, 0
        for r in body:
            raw = sum(_v(r, c) for c in cols)
            mx = max(mx, raw)
            tot = r[TOT[key]]
            if isinstance(tot, (int, float)) and abs(raw - tot) > 1e-9:
                ng += 1
        if ng:
            bad.append("%s%s 明細計と合計列が食い違う団体%d" % (key[0], key[1], ng))
        if mx > MANTEN:
            bad.append("%s%s 素点の最大%.0f が満点%d を超える" % (key[0], key[1], mx, MANTEN))
        memo.append("  %s%s 配点計%5.0f  満点%4d  全国の素点の最大%5.0f"
                    % (key[0], key[1], h, MANTEN, mx))
    return bad, memo


def table():
    """項目ごとに、得点に換算した値と県内の取得状況を並べる"""
    det, mura, ken, body = load()
    out = []
    for c, kin, mok, gun, si, hai in det:
        tok = _v(mura, c)
        man_ken = sum(1 for r in ken if _v(r, c) >= hai)
        ari_ken = sum(1 for r in ken if _v(r, c) > 0)
        hit = AK.H.get((kin, mok, si)) or H2.get((kin, mok, si))
        if hit:
            dan = DAN[hit[0]]
            naze, soan, kakunin = hit[1], hit[2], hit[3]
        elif "成果" in gun:
            dan, naze, soan, kakunin = "Ｄ", "国の統計から自動算定される", "6-3の注記", "―"
        else:
            dan, naze, soan, kakunin = "Ｅ", "個別の判定をしていない", "―", "―"
        out.append({
            "列": c, "交付金": kin, "目標": mok, "目標名": NAME[mok][kin],
            "指標群": gun, "指標": si, "配点": hai, "本村": tok,
            "未取得": hai - tok,
            "県内満点": man_ken, "県内1点以上": ari_ken,
            "県内満点率": man_ken / len(ken),
            "段": dan, "必要なこと": naze, "素案": soan, "確認事項": kakunin,
            "期待得点": (hai - tok) * KEISU[dan],
        })
    return out, mura, ken, body


def mokuhyo(det, mura, ken):
    """目標ごとの得点と、県内・全国の位置"""
    item = collections.defaultdict(list)
    for c, k, mok, _g, _s, _h in det:
        item[(k, mok)].append(c)
    out = []
    for key in sorted(item, key=lambda x: (x[0] != "推進", x[1])):
        cols, tc = item[key], TOT[key]
        mine = mura[tc]
        kv = sorted((r[tc] for r in ken if isinstance(r[tc], (int, float))),
                    reverse=True)
        out.append({
            "交付金": key[0], "目標": key[1], "目標名": NAME[key[1]][key[0]],
            "本村": mine, "満点": MANTEN,
            "県平均": sum(kv) / len(kv), "県中央値": kv[len(kv) // 2],
            "県最高": kv[0], "県最低": kv[-1],
            "県内順位": sum(1 for x in kv if x > mine) + 1, "県内団体数": len(kv),
            "伸ばせる余地": MANTEN - mine,
            "配点計": sum(h for _c, _k, _m, _g, _s, h in det
                          if (_k, _m) == key),
        })
    return out


def nobishiro(rows, mk):
    """見かけの未取得と、実際に伸ばせる点を段ごとに分ける。

    【なぜ分ける必要があるか】
      明細の配点を足し上げると929点になるが、制度の満点は800点である。
      差129点は、択一の枝（成果指標群のア／イは「より上位となった方で得点」）を
      二重に数えたぶんである。目標ごとに見ると
      推進目標Ⅳ 160点／支援目標Ⅳ 160点／支援目標Ⅰ 109点 が満点100点を超えており、
      超過は 60＋60＋9＝129点でちょうど説明がつく。
      したがって「未取得の配点の合計」は見かけの数であり、
      447点の団体が実際に伸ばせるのは 800−447＝353点が上限である。

    【段への割り振りの規則】
      超過が生じている3目標の未取得が、どの段にあるかで決める。
        推進目標Ⅳ・支援目標Ⅳ　未取得はすべてＤ段 → 超過120点はＤ段から引く
        支援目標Ⅰ　　　　　　未取得はＢ・Ｃ・Ｅ段に散るが、Ａ〜Ｃを積み上げても
                              80点で満点100点に届かない（頭打ちの確かめ）。
                              超過9点が効くのはＥ段まで取ったときであるため
                              Ｅ段から引く。
      Ａ・Ｂ・Ｃは見かけと実際が一致する。村にお示しする取得目標はこの3段である。
    """
    man = {(r["交付金"], r["目標"]): r["満点"] for r in mk}
    mine = {(r["交付金"], r["目標"]): r["本村"] for r in mk}
    # 目標ごとの超過（見かけの配点計 − 満点）
    chouka = {(r["交付金"], r["目標"]): max(0, r["配点計"] - r["満点"]) for r in mk}
    # 超過のある目標の未取得が集まっている段
    hiku = collections.Counter()
    for key, ch in chouka.items():
        if not ch:
            continue
        g = collections.Counter()
        for r in rows:
            if (r["交付金"], r["目標"]) == key and r["未取得"] > 0:
                g[r["段"]] += r["未取得"]
        # Ｄ段に全部あるならＤ段、散っているなら最も後ろの段（Ｅ）から引く
        dan = "Ｄ" if set(g) == {"Ｄ"} else max(g, key=lambda d: "ＡＢＣＤＥ".index(d))
        hiku[dan] += ch
    # 目標Ⅳ（成果指標群）は推進・支援で同一の指標であり、集計表に2度現れる。
    # 点は両交付金で別に支払われるため足して正しいが、項目数は二重になる。
    # 支援側を重複として数え、重複を除いた項目数も併せて返す。
    sui_na = {r["指標"] for r in rows if r["交付金"] == "推進"}

    def _dup(r):
        return r["交付金"] == "支援" and r["指標"] in sui_na

    out = []
    for d in "ＡＢＣＤＥ":
        g = [r for r in rows if r["段"] == d and r["未取得"] > 0]
        if not g:
            continue
        mi = sum(r["未取得"] for r in g)
        out.append({"段": d, "内容": DAN_NA[d], "項目": len(g),
                    "重複を除く": len([r for r in g if not _dup(r)]),
                    "見かけ": mi, "実際": mi - hiku[d],
                    "期待得点": sum(r["期待得点"] for r in g)})
    return out, sum(man.values()) - sum(mine.values())


def suii():
    """3か年（令和6〜8年度）の項目別の推移。

       data/交付金評価指標_3か年.csv を、集計表から切り出した明細の鍵で照合する。
       小計の行は鍵が合わないため自然に落ちる。
    """
    det, _m, _k, _b = load()
    keys = {(d[1], d[2], d[4]): d for d in det}
    KN = {"推進交付金": "推進", "支援交付金": "支援"}
    by = {}
    hai_of = {}          # (鍵, 年度) → その年度の配点
    for r in csv.DictReader(io.open(os.path.join(P.DATA, "交付金評価指標_3か年.csv"),
                                    encoding="utf-8-sig")):
        k = (KN.get(r["交付金"], r["交付金"]), r["目標"], r["指標"])
        if k not in keys:
            continue
        try:
            v = float(r["北塩原村"])
        except (TypeError, ValueError):
            continue
        by.setdefault(k, {})[r["年度"]] = v
        try:
            hai_of[(k, r["年度"])] = float(r["配点"])
        except (TypeError, ValueError, KeyError):
            pass
    Y = ("令和6年度", "令和7年度", "令和8年度")
    kubun = collections.Counter()
    ten = collections.Counter()
    nobi = []
    for k, v in by.items():
        if len(v) != 3:
            continue
        a, b, c = (v[y] for y in Y)
        hai = keys[k][5]
        # 素点ではなく得点率（得点÷その年度の配点）で比べる。
        #   国の配点は年度ごとに変わる。素点で比べると、配点が2点から1点に
        #   下がって満点を取り続けている項目が「悪化」に入る。現に支援の
        #   目標Ⅰの3項目（通いの場参加者の健康状態の把握・分析 ウ／
        #   介護予防・生活支援の体制整備 イ・ウ）がそれで、成果品15が
        #   「悪化26項目」、素案が「悪化23項目」と村に別の数を示していた。
        #   analyze_kofukin.suii() と同じ規則に揃える。
        pa, pc = hai_of.get((k, Y[0])), hai_of.get((k, Y[2]), hai)
        ra = a / pa if pa else 0
        rc = c / pc if pc else 0
        if rc > ra + 1e-9:
            g = "改善"
        elif rc < ra - 1e-9:
            g = "悪化"
        elif c == 0:
            g = "3年連続0点"
        elif pc and c >= pc:
            g = "満点で維持"
        else:
            g = "途中で据え置き"
        kubun[g] += 1
        ten[g] += c
        if c != a:
            nobi.append((c - a, k, hai, (a, b, c)))
    nobi.sort()
    return kubun, ten, nobi, len([v for v in by.values() if len(v) == 3])


def main():
    rows, mura, ken, body = table()
    if "--check" in sys.argv:
        d, _m, _k, b = load()
        bad, memo = check(d, b)
        print("■ 明細の切り出しの確かめ（全国%d団体）\n" % len(b))
        print("\n".join(memo))
        if bad:
            print("\n  ✕ 成り立たない\n    " + "\n    ".join(bad[:6]))
            return 1
        print("\n  ○ 明細の合計は目標の合計列に等しい（項目の配点はそのまま得点になる）")
        print("  ○ 素点が満点を超える団体は全国にない")
        return 0

    det0, _m, _k, _b = load()
    mk = mokuhyo(det0, mura, ken)
    sui = sum(r["本村"] for r in mk if r["交付金"] == "推進")
    shi = sum(r["本村"] for r in mk if r["交付金"] == "支援")

    print("■ 令和8年度の獲得状況　合計 %d点／800点（得点率 %.1f%%）\n"
          % (sui + shi, (sui + shi) / 800 * 100))
    print("  推進交付金 %3d点／400点　　支援交付金 %3d点／400点\n" % (sui, shi))
    print("  %-4s %-34s %5s %5s %6s %6s %6s %8s"
          % ("目標", "目標名", "本村", "満点", "県平均", "県中央", "県最高", "県内順位"))
    print("  " + "─" * 94)
    for r in mk:
        print("  %s%s %-34s %5d %5d %6.1f %6d %6d %5d/%d"
              % (r["交付金"][0], r["目標"], r["目標名"][:34], r["本村"], r["満点"],
                 r["県平均"], r["県中央値"], r["県最高"],
                 r["県内順位"], r["県内団体数"]))
    print("\n  ※ 目標Ⅳ（成果指標群）は推進・支援で同一の指標であり、得点も同じである。")
    print("  ※ 県の値は59市町村の集計である。団体名と個別の数値は示さない。")

    koe = [r for r in mk if r["配点計"] > r["満点"]]
    print("\n  項目の配点はそのまま得点になる（明細の合計は目標の合計列に等しい）。")
    print("  （%s）"
          % "／".join("%s%s %.0f点" % (r["交付金"][0], r["目標"], r["配点計"])
                      for r in mk))
    if koe:
        print("  ただし%sは配点計が満点%d点を超える。"
              % ("・".join("%s%s %.0f点" % (r["交付金"][0], r["目標"], r["配点計"])
                           for r in koe), MANTEN))
        print("  成果指標群のア／イなどが択一（より上位となった方で得点）であり、")
        print("  両方を足すと二重に数えることによる。配点の総和%.0f点に対し制度の満点は%d点で、"
              % (sum(r["配点計"] for r in mk), MANTEN * len(mk)))
        print("  差%.0f点は超過（%s）でちょうど説明がつく。"
              % (sum(r["配点計"] for r in mk) - MANTEN * len(mk),
                 "＋".join("%.0f" % (r["配点計"] - r["満点"]) for r in koe)))

    print("\n■ 未取得の点　取得の見込みの段ごと\n")
    nob, nobi_kei = nobishiro(rows, mk)
    print("  %-2s %-18s %4s %8s %9s %11s %9s"
          % ("段", "内容", "項目", "重複を除く", "見かけ", "実際に伸ばせる", "期待得点"))
    print("  " + "─" * 82)
    for x in nob:
        shirushi = "" if x["見かけ"] == x["実際"] else " ←"
        print("  %-2s %-18s %4d %8d %9.0f %11.0f %9.1f%s"
              % (x["段"], x["内容"], x["項目"], x["重複を除く"],
                 x["見かけ"], x["実際"], x["期待得点"], shirushi))
    print("  " + "─" * 82)
    print("  %-2s %-18s %4d %8d %9.0f %11.0f %9.1f"
          % ("計", "", sum(x["項目"] for x in nob),
             sum(x["重複を除く"] for x in nob),
             sum(x["見かけ"] for x in nob), sum(x["実際"] for x in nob),
             sum(x["期待得点"] for x in nob)))
    print("\n  「見かけ」は未取得の配点の合計である。択一の枝を二重に数えているため、")
    print("  これを全部取ることはできない。実際に伸ばせるのは満点%d点との差%.0f点までで、"
          % (MANTEN * len(mk), nobi_kei))
    print("  段ごとの内訳が上の「実際に伸ばせる」である（←を付けた段が減る）。")
    print("  Ａ・Ｂ・Ｃは見かけと実際が一致する。村にお示しする取得目標はこの3段である。")
    print("  「重複を除く」は、目標Ⅳ（成果指標群）が推進・支援で同一の指標であり")
    print("  集計表に2度現れることによる。点は両交付金で別に支払われるため足して正しいが、")
    print("  取り組む対象の数としては1つである。")
    assert abs(sum(x["実際"] for x in nob) - nobi_kei) < 0.5, "段の計が伸びしろと合わない"

    print("\n■ 優先順位　ＡＢＣの項目を期待得点の大きい順に（県内満点率を併記）\n")
    pri = sorted([r for r in rows if r["段"] in ("Ａ", "Ｂ", "Ｃ")
                  and r["未取得"] > 0],
                 key=lambda r: (-r["期待得点"], -r["県内満点率"]))
    print("  %-3s %-2s %-5s %-40s %5s %5s %6s %7s"
          % ("順", "段", "目標", "指標", "配点", "得点", "期待", "県内満点"))
    print("  " + "─" * 100)
    for i, r in enumerate(pri, 1):
        print("  %3d %-2s %s%-4s %-40s %5.0f %5.1f %6.1f %5d/59"
              % (i, r["段"], r["交付金"][0], r["目標"], r["指標"][:40],
                 r["未取得"], r["未取得"], r["期待得点"], r["県内満点"]))

    print("\n■ 未判定（Ｅ段）のうち、県内で多くの団体が取得している項目\n")
    print("  県内で取得している団体が多いのに本村が0点のものは、判定を急ぐ。")
    print("  %-5s %-44s %5s %9s" % ("目標", "指標", "配点", "県内満点"))
    print("  " + "─" * 70)
    ee = sorted([r for r in rows if r["段"] == "Ｅ" and r["未取得"] > 0],
                key=lambda r: (-r["県内満点率"], -r["未取得"]))
    for r in ee[:8]:
        print("  %s%-4s %-44s %5.0f %5d/59" % (r["交付金"][0], r["目標"],
              r["指標"][:44], r["未取得"], r["県内満点"]))
    print("  " + "─" * 70)
    print("  Ｅ段は%d項目%.0f点。うち県内の半数以上が取得しているものは%d項目%.0f点、"
          % (len(ee), sum(r["未取得"] for r in ee),
             len([r for r in ee if r["県内満点率"] >= 0.5]),
             sum(r["未取得"] for r in ee if r["県内満点率"] >= 0.5)))
    print("  4分の1以上は%d項目%.0f点である。"
          % (len([r for r in ee if r["県内満点率"] >= 0.25]),
             sum(r["未取得"] for r in ee if r["県内満点率"] >= 0.25)))

    print("\n■ 見込める得点　目標ごとに積み上げる（満点100点の頭打ちを見る）\n")
    print("  %-5s %5s %7s %7s %7s %7s %6s"
          % ("目標", "いま", "＋Ａ", "＋Ｂ", "＋Ｃ", "満点", "県平均"))
    print("  " + "─" * 62)
    tot = {"いま": 0.0, "Ａ": 0.0, "Ｂ": 0.0, "Ｃ": 0.0}
    for r in mk:
        key = (r["交付金"], r["目標"])
        g = {d: sum(x["未取得"] for x in rows
                    if (x["交付金"], x["目標"]) == key and x["段"] == d)
             for d in "ＡＢＣ"}
        a1 = min(MANTEN, r["本村"] + g["Ａ"])
        a2 = min(MANTEN, r["本村"] + g["Ａ"] + g["Ｂ"])
        a3 = min(MANTEN, r["本村"] + g["Ａ"] + g["Ｂ"] + g["Ｃ"])
        print("  %s%-4s %5d %7.0f %7.0f %7.0f %7d %6.1f"
              % (r["交付金"][0], r["目標"], r["本村"], a1, a2, a3,
                 MANTEN, r["県平均"]))
        tot["いま"] += r["本村"]; tot["Ａ"] += a1; tot["Ｂ"] += a2; tot["Ｃ"] += a3
    print("  " + "─" * 62)
    print("  %-5s %5.0f %7.0f %7.0f %7.0f %7d %6.1f"
          % ("合計", tot["いま"], tot["Ａ"], tot["Ｂ"], tot["Ｃ"], MANTEN * 8, 456.4))
    print("\n  満点100点に頭打ちになる目標はない（上の＋Ｃまででも最大%d点）。"
          % max(min(MANTEN, r["本村"] + sum(x["未取得"] for x in rows
                if (x["交付金"], x["目標"]) == (r["交付金"], r["目標"])
                and x["段"] in "ＡＢＣ")) for r in mk))
    print("  ※ 評価指標は毎年度改定される。第10期（令和9〜11年度）の指標は未公表であり、")
    print("     上の値は令和8年度の指標によるものである。")
    print("  ※ 目標Ⅳ（成果指標群）は推進・支援で同一の指標であるため、")
    print("     片方を動かせばもう片方も同じだけ動く。")

    if "--json" in sys.argv:
        import json
        g = {d: [r for r in rows if r["段"] == d and r["未取得"] > 0] for d in "ＡＢＣＤＥ"}
        pri2 = sorted(g["Ａ"] + g["Ｂ"] + g["Ｃ"],
                      key=lambda r: (-r["期待得点"], -r["県内満点率"]))
        kubun, ten, nobi, n3 = suii()
        nob2, nobi_kei2 = nobishiro(rows, mk)
        zenv = sorted((r[335] for r in body if isinstance(r[335], (int, float))),
                      reverse=True)
        kenv = sorted((r[335] for r in ken if isinstance(r[335], (int, float))),
                      reverse=True)

        def rk(v, a):
            return sum(1 for x in a if x > v) + 1
        tsumi = []
        for r in mk:
            key = (r["交付金"], r["目標"])
            gg = {d: sum(x["未取得"] for x in rows
                         if (x["交付金"], x["目標"]) == key and x["段"] == d)
                  for d in "ＡＢＣ"}
            tsumi.append({
                "交付金": r["交付金"], "目標": r["目標"], "目標名": r["目標名"],
                "いま": r["本村"],
                "A": min(MANTEN, r["本村"] + gg["Ａ"]),
                "B": min(MANTEN, r["本村"] + gg["Ａ"] + gg["Ｂ"]),
                "C": min(MANTEN, r["本村"] + gg["Ａ"] + gg["Ｂ"] + gg["Ｃ"]),
                "県平均": round(r["県平均"], 1), "県中央値": r["県中央値"],
                "県最高": r["県最高"], "県最低": r["県最低"],
                "県内順位": r["県内順位"], "県内団体数": r["県内団体数"],
            })
        out = {
            "日付": "令和8年10月21日",
            "合計": sui + shi, "推進": sui, "支援": shi, "満点": MANTEN * 8,
            "県平均合計": 456.4, "県内順位": 34, "県内団体数": 59,
            "全国平均合計": 455.1, "全国順位": 1005, "全国団体数": 1741,
            "目標別": tsumi,
            "段": [{"段": x["段"], "内容": x["内容"], "項目": x["項目"],
                    "重複を除く": x["重複を除く"],
                    "点": x["見かけ"], "実際": x["実際"],
                    "期待": round(x["期待得点"], 1)}
                   for x in nob2],
            "伸びしろ": nobi_kei2,
            "配点の総和": sum(r["配点計"] for r in mk),
            "超過": [{"交付金": r["交付金"], "目標": r["目標"],
                      "配点計": r["配点計"], "超過": r["配点計"] - r["満点"]}
                     for r in mk if r["配点計"] > r["満点"]],
            "優先順位": [{
                "順": i, "段": r["段"], "交付金": r["交付金"], "目標": r["目標"],
                "指標": r["指標"], "点": r["未取得"],
                "期待": round(r["期待得点"], 1),
                "県内満点": r["県内満点"], "県内団体数": len(ken),
                "必要なこと": r["必要なこと"], "素案": r["素案"],
            } for i, r in enumerate(pri2, 1)],
            "未判定": [{
                "交付金": r["交付金"], "目標": r["目標"], "指標": r["指標"],
                "点": r["未取得"], "県内満点": r["県内満点"],
            } for r in sorted(g["Ｅ"], key=lambda r: (-r["県内満点率"], -r["未取得"]))[:12]],
            "未判定計": {"項目": len(g["Ｅ"]), "点": sum(r["未取得"] for r in g["Ｅ"]),
                         "半数以上": len([r for r in g["Ｅ"] if r["県内満点率"] >= 0.5]),
                         "半数以上の点": sum(r["未取得"] for r in g["Ｅ"]
                                             if r["県内満点率"] >= 0.5),
                         "4分の1以上": len([r for r in g["Ｅ"] if r["県内満点率"] >= 0.25]),
                         "4分の1以上の点": sum(r["未取得"] for r in g["Ｅ"]
                                               if r["県内満点率"] >= 0.25)},
            "順位": [{"点": v, "全国": rk(v, zenv), "県内": rk(v, kenv)}
                     for v in sorted({sui + shi,
                                      sum(t["A"] for t in tsumi),
                                      sum(t["B"] for t in tsumi),
                                      sum(t["C"] for t in tsumi)})],
            "推移": {"そろう項目": n3,
                     "区分": [{"区分": k2, "項目": kubun[k2], "点": ten[k2]}
                              for k2 in ("改善", "悪化", "満点で維持", "3年連続0点",
                                         "途中で据え置き") if kubun[k2]]},
            "悪化": [{"交付金": k2[0], "目標": k2[1], "指標": k2[2], "配点": h,
                      "推移": "%.0f→%.0f→%.0f" % yy, "増減": d2}
                     for d2, k2, h, yy in nobi if d2 < 0],
        }
        fp = P.build("kofukin.json")
        with io.open(fp, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=1)
        print("\n  書き出した： %s" % fp)

    if "--csv" in sys.argv:
        # 見出しは行の辞書の鍵と同じ名でなければならない。
        # extrasaction="ignore" のため、合っていない見出しは黙って空欄になる。
        # 現に「未取得の配点」「得点への換算」「取れる得点」の3列が
        # 299行すべて空で出ていた（行の鍵は「未取得」である）。
        ks = ["交付金", "目標", "目標名", "指標群", "指標", "配点", "本村",
              "未取得", "期待得点",
              "県内満点", "県内1点以上", "段", "必要なこと", "素案", "確認事項"]
        with io.open(OUT, "w", encoding="utf-8", newline="") as f:
            kaketa = [k for k in ks if k not in rows[0]]
            if kaketa:
                raise SystemExit("CSVの見出しに行の鍵がない： %s" % "・".join(kaketa))
            w = csv.DictWriter(f, fieldnames=ks, extrasaction="ignore")
            w.writeheader()
            for r in sorted(rows, key=lambda r: (-r["期待得点"], -r["県内満点率"])):
                w.writerow(r)
        print("\n  書き出した： %s（%d行）" % (OUT, len(rows)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
