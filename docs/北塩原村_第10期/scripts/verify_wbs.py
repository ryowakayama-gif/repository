# -*- coding: utf-8 -*-
"""仕様書（spec_data.S）と WBS（wbs_data.W）と確認事項（wbs_kakunin.K）の整合を機械的に点検する。

   不適合があれば内容を出力して終了コード 1 を返す。

   点検1 WBS の ID が一意であること
   点検2 仕様書の要求事項が参照する WBS ID が実在すること
   点検3 すべての要求事項に対応作業があること（＝仕様書からの漏れがない）
   点検4 仕様書に根拠を持つ WBS 作業が、いずれかの要求事項から参照されていること（＝宙に浮いた作業がない）
   点検5 WBS の「仕様書該当」列と、参照元の要求事項の条項が一致すること
   点検6 仕様書外の作業（手引き・確認）がすべて SHIEN で支援先を示されていること
   点検7 確認事項の「関連WBS No.」が実在する ID を指していること
   点検8 WBS の表示順が 大分類→中分類 の順に整っていること
"""
import os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from spec_data import S, SHIEN, GAIBU
from wbs_data import W
from wbs_kakunin import K, SOLVED

NG = []
def chk(no, name, ok, detail=""):
    print(f"  {'○' if ok else '✕'} 点検{no} {name}" + (f"　… {detail}" if detail else ""))
    if not ok:
        NG.append(f"点検{no} {name}：{detail}")

ids = [w[0] for w in W]
IDX = {w[0]: w for w in W}

# 1
dup = [i for i in set(ids) if ids.count(i) > 1]
chk(1, "WBS IDの一意性", not dup, f"重複 {dup}" if dup else f"{len(ids)}件すべて一意")

# 2
miss = sorted({i for _, _, _, refs, _ in S for i in refs if i not in IDX})
chk(2, "要求事項が参照するWBS IDの実在", not miss, f"不明 {miss}" if miss else f"参照{sum(len(r[3]) for r in S)}件")

# 3
nolink = [(c, t) for c, t, tan, refs, _ in S if not refs and tan != "村"]
chk(3, "仕様書の要求事項の充足", not nolink,
    "／".join(f"{c} {t}" for c, t in nolink) if nolink else f"要求事項{len(S)}件すべてに対応作業あり")

# 4
referred = {i for _, _, _, refs, _ in S for i in refs}
orphan = [(w[0], w[3]) for w in W if w[4].startswith("仕様書") or w[4].startswith("通知")
          if w[0] not in referred]
chk(4, "仕様書に根拠をもつ作業の被参照", not orphan,
    "／".join(f"{a} {b}" for a, b in orphan) if orphan else "")

# 5 条項の一致（「仕様書該当」列の条項から、その作業を参照している要求事項があるか）
bad5 = []
for w in W:
    if not (w[4].startswith("仕様書") or w[4].startswith("通知")):
        continue
    src_clauses = [re.sub(r"-[a-z]$", "", c) for c, _, _, refs, _ in S if w[0] in refs]
    if not any(c == w[4] or c.startswith(w[4]) for c in src_clauses):
        bad5.append(f"{w[0]}（列={w[4]} / 参照元={'・'.join(sorted(set(src_clauses))) or 'なし'}）")
chk(5, "「仕様書該当」列と参照元条項の一致", not bad5, "／".join(bad5) if bad5 else "")

# 6
gaibu = [w[0] for w in W if w[4] in GAIBU]
nosh = [i for i in gaibu if i not in SHIEN]
chk(6, "仕様書外の作業の位置づけの明示", not nosh,
    f"支援先未記載 {nosh}" if nosh else f"仕様書外{len(gaibu)}件すべてに支援先を明示")

# 7
refs_k, bad7 = 0, set()
for row in list(K) + list(SOLVED):
    for r in re.split(r"[,、]\s*", row[3]):
        r = r.strip()
        if not r or r == "―":
            continue
        refs_k += 1
        if r not in IDX:
            bad7.add(r)
chk(7, "確認事項の関連WBS No.の実在", not bad7,
    f"不明 {sorted(bad7)}" if bad7 else f"参照{refs_k}件すべて実在")

# 8
MAJ = ["０ 契約・立上げ", "Ⅰ ニーズ調査", "Ⅱ 計画策定", "Ⅲ 打合せ等", "Ⅳ 成果品・納品", "Ⅴ 管理"]
seq, seen, bad8 = [], set(), []
for w in W:
    key = (w[1], w[2])
    if not seq or seq[-1] != key:
        if key in seen:
            bad8.append(f"{w[1]}／{w[2]}")
        seq.append(key)
        seen.add(key)
if [m for m in (k[0] for k in seq)] != sorted({k[0] for k in seq}, key=MAJ.index) * 0 + \
        [m for m in dict.fromkeys(k[0] for k in seq)] or True:
    pass
order_ok = [MAJ.index(m) for m in dict.fromkeys(k[0] for k in seq)] == \
           sorted([MAJ.index(m) for m in dict.fromkeys(k[0] for k in seq)])
chk(8, "WBSの表示順", not bad8 and order_ok,
    f"分断 {bad8}" if bad8 else ("大分類の順序が不正" if not order_ok else f"{len(seq)}区分が整列"))

print()
if NG:
    print(f"不適合 {len(NG)}件")
    for n in NG:
        print("  -", n)
    sys.exit(1)
print(f"適合：仕様書の要求事項 {len(S)}件／WBS {len(W)}件／確認事項の参照 {refs_k}件")
