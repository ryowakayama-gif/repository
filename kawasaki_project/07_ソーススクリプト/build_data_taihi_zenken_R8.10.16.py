# -*- coding: utf-8 -*-
"""対比の洗い出しの全件を data_taihi_zenken_R8.10.16.py に書き出す
（令和8年10月16日）

他メンバー作成の第10期計画素案との対比を、11の持ち場に分けて手分けして
洗い出した結果を、**リポジトリの中のデータとして残す**。

  入力（第一）　07_ソーススクリプト/data_taihi_zenken_R8.10.16.py
                （既にあればそれが正本。作り直さない）
  入力（代替）　対比の作業の記録（journal.jsonl）
                ⚠ セッションごとに作り直される一時の場所にある。
                  **そこにしかない状態にしないために本スクリプトで
                  リポジトリへ書き出す。**

  python3 07_ソーススクリプト/build_data_taihi_zenken_R8.10.16.py [記録のパス]

⚠ **委員の氏名その他の個人情報を書き出さない。** 資料編の持ち場は
  委員名簿を扱うため、書き出す前に機械で確かめ、当たったものは落とす。
"""
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "data_taihi_zenken_R8.10.16.py")

# 既定の記録の場所（⚠ 一時の場所。リポジトリの外）
KIROKU = ("/root/.claude/projects/-home-user-repository/"
          "4be8f82c-e2c9-52ac-b2a9-bad5154d0b13/subagents/workflows/"
          "wf_45522d87-2e6/journal.jsonl")

# 持ち場の名（記録に現れる順）
TANTO = ["総論（第1章・第2章1〜5節）", "調査結果（第2章6節）",
         "基本理念と施策体系（第3章）", "各論1〜3章（健康・安心・地域生活）",
         "地域支援事業（第4章1・2節）", "認知症施策（第4章3節）",
         "給付サービス（第5章）", "給付費・保険料（第6〜8章）",
         "資料編と体裁", "第9期計画の再整理",
         "批判（聞き漏らし）", "批判（取り込みの可否）"]

# ⚠ 個人情報らしい形（当たったものは書き出さない）
PII = [
    (r"(明治|大正|昭和|平成|令和)\d{1,2}年\d{1,2}月\d{1,2}日生", "生年月日"),
    (r"0\d{1,4}-\d{2,4}-\d{3,4}", "電話番号"),
    (r"[\w.+-]+@[\w-]+\.[\w.-]+", "メールアドレス"),
    # 委員名簿の行をそのまま写していないか（氏名の列に値が入っている形）
    (r"委員名[｜|]\s*[^｜|\s]{2,8}[　 ][^｜|\s]{1,8}", "委員名簿の氏名"),
]
# 氏名は値の形では探せないため、名の出てくる語を名指しで落とす
#   ⚠ 町長名は計画書の「はじめに」に載るものであり誤りではないが、
#     当方の内部の整理には要らない。
NAZASHI = ["小山　修作", "小山 修作", "小山修作"]


def yomu(path):
    rs = []
    for line in open(path, encoding="utf-8"):
        if '"result"' not in line:
            continue
        try:
            d = json.loads(line)
        except ValueError:
            continue
        if d.get("type") != "result":
            continue
        r = d.get("result") or {}
        if isinstance(r, dict) and "items" in r:
            rs.append(r)
    return rs


def arau(s):
    """個人情報らしい箇所を落とす。落としたら理由を返す。"""
    riyu = []
    for pat, na in PII:
        if re.search(pat, s):
            s = re.sub(pat, f"［{na}は書き出していません］", s)
            riyu.append(na)
    for w in NAZASHI:
        if w in s:
            s = s.replace(w, "［町長名］")
            riyu.append("町長名")
    return s, riyu


def main():
    if os.path.exists(OUT):
        print("既にあります：", OUT)
        print("  ⚠ リポジトリの中のこのファイルが正本です。作り直しません。")
        print("  作り直す場合は、先にこのファイルを消してください。")
        return
    path = sys.argv[1] if len(sys.argv) > 1 else KIROKU
    if not os.path.exists(path):
        raise SystemExit(
            "記録が見つかりません：" + path
            + "\n  ⚠ 記録は一時の場所にあります。"
              "対比の作業と同じセッションのうちに実行してください。")
    rs = yomu(path)
    rows, aratta = [], []
    for i, r in enumerate(rs):
        tan = TANTO[i] if i < len(TANTO) else f"持ち場{i + 1}"
        for x in r.get("items", []):
            vals = []
            for k in ("topic", "kubun", "tenpu", "tenpu_loc", "touhou",
                      "touhou_loc", "handan", "sagyo", "omomi"):
                s, riyu = arau(str(x.get(k) or ""))
                if riyu:
                    aratta.append((tan, x.get("topic", "")[:24], riyu))
                vals.append(s)
            rows.append([tan] + vals)

    with open(OUT, "w", encoding="utf-8") as f:
        f.write('# -*- coding: utf-8 -*-\n')
        f.write('"""他メンバー素案との対比　洗い出しの全件'
                '（令和8年10月16日）\n\n')
        f.write('11の持ち場に手分けして洗い出した結果である。\n')
        f.write('build_data_taihi_zenken_R8.10.16.py が書き出した。\n\n')
        f.write('⚠ **委員の氏名その他の個人情報は書き出していない。**\n')
        f.write('⚠ 重点の論点（当方が機械で確かめたもの）は\n')
        f.write('  data_taihi_tenpu_R8.10.16.py の RONTEN にある。\n')
        f.write('  本ファイルは洗い出しの全件であり、重複を含む。\n\n')
        f.write('（持ち場, 論点, 区分, 添付の記述, 添付の箇所,\n')
        f.write('  当方の記述, 当方の節, 判断, 作業, 重み）\n"""\n')
        f.write('ZENKEN = [\n')
        for r in rows:
            f.write("    (" + ", ".join(repr(x) for x in r) + "),\n")
        f.write(']\n')

    print("洗い出しの全件を書き出しました")
    print("保存：", OUT)
    print(f"  持ち場 {len(rs)}／論点 {len(rows)}件")
    import collections
    for k, v in collections.Counter(r[0] for r in rows).most_common():
        print(f"   {k:<30}{v:>4}件")
    print(f"  区分：{dict(collections.Counter(r[2] for r in rows))}")
    if aratta:
        print(f"  ── 個人情報らしい箇所を落としました {len(aratta)}件")
        for tan, to, riyu in aratta[:10]:
            print(f"   {tan}／{to}／{'・'.join(riyu)}")
    else:
        print("   ○ 個人情報らしい箇所はありませんでした")


if __name__ == "__main__":
    main()
