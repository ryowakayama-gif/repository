# -*- coding: utf-8 -*-
"""コードブックに従ってダミー回答データを生成する。
   集計プログラムを本番データの到着前に動かして検証するためのもの。
   出力: data/dummy_ニーズ.csv / data/dummy_在宅.csv
"""
import os, json, csv, random

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BOOK = os.path.join(BASE, "data", "集計_コードブック.json")

N_NEEDS, N_ZAITAKU = 430, 96      # 想定回収数（ニーズ850配布×約50%／在宅150×約64%）
SEED = 20260908

CHIKU = ["北山", "大塩", "桧原", "裏磐梯"]
CHIKU_W = [0.55, 0.20, 0.13, 0.12]          # 人口比に近い重み
NENREI = ["65〜74歳", "75〜84歳", "85歳以上"]
NENREI_W = [0.45, 0.36, 0.19]
SEIBETSU = ["男性", "女性"]
SEIBETSU_W = [0.46, 0.54]
NINTEI = ["認定なし", "要支援1・2", "要介護1・2", "要介護3〜5"]
NINTEI_W = [0.80, 0.09, 0.07, 0.04]         # ニーズ調査は要介護1〜5を除くため軽度が中心
YOKAIGO = ["要支援1・2", "要介護1・2", "要介護3〜5"]
YOKAIGO_W = [0.38, 0.37, 0.25]
KINMU = ["フルタイム", "パートタイム", "働いていない"]
KINMU_W = [0.18, 0.14, 0.68]

MUKAITO = 0.06     # 無回答の発生率

# リスク判定に関わる設問は、実際に近い分布になるよう重みを与える。
#   ダミーを一様乱数にすると全員がリスク該当となり、派生変数の検算にならないため。
#   数値は第9期調査および全国の傾向を踏まえたおおよその想定値で、実データとは無関係。
WEIGHTS = {
 # 問2 からだを動かすこと（できる／できるけど／できない）
 "N_問2_1": [72, 20, 8], "N_問2_2": [75, 18, 7], "N_問2_3": [70, 20, 10],
 "N_問2_4": [10, 16, 74],                        # 転倒（何度も／1度／ない）
 "N_問2_5": [14, 30, 32, 24],                    # 転倒への不安
 "N_問2_6": [4, 10, 40, 46],                     # 外出頻度
 "N_問2_7": [8, 24, 34, 34],                     # 外出の減少
 "N_問2_8": [26, 74],                            # 外出を控えている
 # 問3 食べること（はい／いいえ）
 "N_問3_2": [28, 72], "N_問3_3": [22, 78], "N_問3_4": [24, 76],
 "N_問3_5": [92, 8], "N_問3_7": [14, 86],
 "N_問3_8": [62, 20, 12, 6],
 # 問4 認知・IADL（はい／いいえ、できる／できるけど／できない）
 "N_問4_1": [40, 60], "N_問4_2": [90, 10], "N_問4_3": [12, 88],
 "N_問4_4": [78, 12, 10], "N_問4_5": [80, 12, 8], "N_問4_6": [72, 16, 12],
 "N_問4_7": [82, 10, 8], "N_問4_8": [80, 12, 8],
 "N_問4_10": [72, 28], "N_問4_11": [52, 48], "N_問4_12": [78, 22],
 "N_問4_13": [58, 42], "N_問4_14": [66, 34], "N_問4_15": [62, 38], "N_問4_16": [70, 30],
 # 問8 健康・うつ
 "N_問8_1": [12, 58, 24, 6],
 "N_問8_3": [22, 78], "N_問8_4": [18, 82],
}
# 問5(1) 会・グループへの参加頻度は「6．参加していない」が大半
SANKA_W = [1, 1, 2, 4, 8, 84]
# 問8(7) 疾病は「1．ない」を一定割合で発生させる
NASHI_RATE = {"N_問8_7": 0.14, "N_問7_1": 0.05, "N_問7_2": 0.06,
              "N_問7_3": 0.07, "N_問7_4": 0.08}
# 支え合いの有無は4設問で相関するため、一定割合を「いずれも該当者なし」にする。
#   独立に乱数を振ると孤立層が0件になり、S99の検算ができない。
ISOLATION_RATE = 0.045


def w_choice(vals, weights):
    return random.choices(vals, weights=weights)[0]


def skewed(n):
    """選択肢番号を1〜nから、若い番号に寄せて選ぶ（回答分布に山を作るため）"""
    ws = [max(1.0, n - abs(i - n * 0.4)) for i in range(1, n + 1)]
    return random.choices(range(1, n + 1), weights=ws)[0]


def gen(book, hyos, n, extra):
    recs = []
    qs = [q for q in book if q["票"] in hyos]
    for i in range(1, n + 1):
        r = {"管理番号": f"{hyos[0][:1]}{i:04d}"}
        r.update(extra())
        for q in qs:
            if not q["列"]:
                continue
            fmt, opts = q["形式"], q["選択肢"]
            if fmt == "数値":
                if "身長" in q["設問文"]:
                    r[q["列"][0]] = "" if random.random() < 0.12 else random.randint(140, 178)
                    r[q["列"][1]] = "" if r[q["列"][0]] == "" else random.randint(38, 82)
                else:
                    lo, hi = q.get("数値範囲") or [0, 10]
                    r[q["列"][0]] = "" if random.random() < MUKAITO else random.randint(lo, hi)
                continue
            if fmt == "自由":
                r[q["列"][0]] = "" if random.random() < 0.85 else "（自由記述のダミー）"
                continue
            if fmt.startswith("複数"):
                cols = [c for c in q["列"] if not c.endswith("_自由")]
                for c in cols:
                    r[c] = 0
                if random.random() < MUKAITO or not cols:
                    continue
                nashi = NASHI_RATE.get(q["変数"])
                if nashi is not None and random.random() < nashi:
                    r[cols[0]] = 1                       # 「ない」「そのような人はいない」
                    continue
                pool = cols[1:] if nashi is not None else cols
                k = random.randint(1, max(1, min(3, len(pool))))
                for c in random.sample(pool, k):
                    r[c] = 1
            else:
                col = q["列"][0]
                n = max(2, len(opts))
                if random.random() < MUKAITO:
                    r[col] = ""
                elif col in WEIGHTS and len(WEIGHTS[col]) == n:
                    r[col] = random.choices(range(1, n + 1), weights=WEIGHTS[col])[0]
                elif col.startswith("N_問5_1_") and n == 6:
                    r[col] = random.choices(range(1, 7), weights=SANKA_W)[0]
                else:
                    r[col] = skewed(n)
            if q["形式"].endswith("＋自由"):
                r[q["変数"] + "_自由"] = "" if random.random() < 0.7 else "（自由記述のダミー）"
    # 分岐条件を反映（該当しない設問は空欄にする）
        apply_skip(r, hyos)
        if "ニーズ" in hyos and random.random() < ISOLATION_RATE:
            for e in (1, 2, 3, 4):
                for k in list(r):
                    if k.startswith(f"N_問7_{e}_c"):
                        r[k] = 1 if k.endswith("_c8") else 0
        recs.append(r)
    return qs, recs


def apply_skip(r, hyos):
    """調査票の分岐に合わせて、対象外の設問を空欄にする"""
    def clear(prefix):
        for k in list(r):
            if k.startswith(prefix):
                r[k] = 0 if k[len(prefix):].startswith("_c") else ""

    if "ニーズ" in hyos:
        if r.get("N_問1_2") == 1:                      # 介護・介助は必要ない
            clear("N_問1_2_1"); clear("N_問1_2_2")
        elif r.get("N_問1_2") != 3:
            clear("N_問1_2_2")
        if r.get("N_問2_8") != 1:                      # 外出を控えていない
            clear("N_問2_8_1")
        if r.get("N_問3_6") not in (1, 3):             # 入れ歯を利用していない
            clear("N_問3_6_2")
    else:
        if r.get("ZA_問8") != 1:                       # サービスを利用していない
            for c, _n, _o in [("A", 0, 0)]:
                pass
            for k in list(r):
                if k.startswith("ZA_問9_"):
                    r[k] = ""
        else:
            clear("ZA_問10")
        if r.get("ZB_問7") not in (1, 2):              # 働いていない
            clear("ZB_問8"); clear("ZB_問9"); clear("ZB_問10")


def main():
    random.seed(SEED)
    book = json.load(open(BOOK, encoding="utf-8"))

    def extra_needs():
        return {"地区": w_choice(CHIKU, CHIKU_W),
                "年齢階級": w_choice(NENREI, NENREI_W),
                "性別": w_choice(SEIBETSU, SEIBETSU_W),
                "認定状況": w_choice(NINTEI, NINTEI_W)}

    def extra_zai():
        return {"地区": w_choice(CHIKU, CHIKU_W),
                "年齢階級": w_choice(NENREI, [0.20, 0.42, 0.38]),
                "性別": w_choice(SEIBETSU, [0.38, 0.62]),
                "要介護度": w_choice(YOKAIGO, YOKAIGO_W),
                "介護者の勤務形態": w_choice(KINMU, KINMU_W)}

    for hyos, n, extra, name in [(["ニーズ"], N_NEEDS, extra_needs, "ニーズ"),
                                 (["在宅A", "在宅B"], N_ZAITAKU, extra_zai, "在宅")]:
        qs, recs = gen(book, hyos, n, extra)
        cols = list(recs[0].keys())
        path = os.path.join(BASE, "data", f"dummy_{name}.csv")
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
            w.writeheader()
            w.writerows(recs)
        print(f"  -> data/dummy_{name}.csv　{len(recs)}件 / {len(cols)}列")


if __name__ == "__main__":
    main()
