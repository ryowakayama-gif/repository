# -*- coding: utf-8 -*-
"""extract_plan_period.py の抽出ロジックの回帰テスト（python3 test_extract_plan_period.py）"""

import extract_plan_period as X

CASES = [
    # (入力, 期待する(開始西暦, 終期西暦)。Noneは「拾ってはいけない」)
    ("計画期間は、平成30年度から令和9年度までの10年間とします。", (2018, 2027)),
    ("本計画の計画期間は令和3年度～令和12年度までの10年間です。", (2021, 2030)),
    ("計画期間 令和2年度（2020年度）から令和11年度（2029年度）まで", (2020, 2029)),
    ("計画期間：2021年度から2030年度（10年間）", (2021, 2030)),
    ("計画の期間 令和6年度〜令和10年度", (2024, 2028)),
    ("計画期間  平成29年度 － 令和8年度", (2017, 2026)),
    ("計画期間は令和元年度から令和10年度までとする", (2019, 2028)),
    ("計画期間｜令和5年度から令和14年度まで（おおむね10年）", (2023, 2032)),
    # PDF→Word変換で表がセル区切りになった形
    ("計画期間 | 平成30年度 | 令和9年度 |", (2018, 2027)),
    ("計画期間 | 令和2年度 | 令和11年度 |", (2020, 2029)),
    # ノイズ（計画期間ではない年の並び）
    ("2050年カーボンニュートラルを目指し、2030年度までに46%削減", None),
    ("策定 令和6年2月 | 発行 平成30年3月", None),
]


def run():
    ng = 0
    for text, expected in CASES:
        cands = X.find_periods(text)
        best = max(cands, key=X.score) if cands else None
        got = (best[0], best[1]) if best and X.score(best) > 0 else None
        ok = got == expected if expected else got is None
        ng += 0 if ok else 1
        print(f"{'OK ' if ok else 'NG '} {text[:52]:54} 期待={expected} 抽出={got}")
    print(f"\n{len(CASES)}件中 NG {ng}件")
    return ng


if __name__ == "__main__":
    raise SystemExit(1 if run() else 0)
