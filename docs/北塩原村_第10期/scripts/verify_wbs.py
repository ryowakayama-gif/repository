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
   点検9 未解決の優先度S・Aのすべてに影響度が与えられていること（計画策定側は IMPACT、調査工程側は機械）
   点検10 影響度が LEVEL に定義された値であること、かつ IMPACT に調査工程側の項目が混ざっていないこと
   点検11 影響度の判定に挙げた「止まる対象」のWBS番号が実在すること
   点検12 照会の束がすべてのWBSを覆い、束の間で重複がないこと
   点検13 翌営業日の作業（READY）が実在する未完了のWBSを指し、影響度1（停止）の項目で止まっていないこと
   点検14 未完了のWBSが READY か HOLD_REASON のいずれかで扱われていること
   点検15 進捗の実績日の整合（完了に実績完了日があること、着手に実績開始日があること、
          未着手に実績日がないこと、実績完了日が実績開始日より後であること）
"""
import os, re, sys
sys.dont_write_bytecode = True   # 古いバイトコードで誤った結果が出ることを防ぐ
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from spec_data import S, SHIEN, GAIBU
from wbs_data import W
from wbs_kakunin import K, SOLVED
from wbs_progress import P
from wbs_pending import LEVEL, IMPACT, BUNDLE, READY, HOLD_REASON

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

# ══════════ ペンディング整理の点検（点検9〜14）══════════
def _ids(k):
    return [x.strip() for x in re.split(r"[,、]\s*", k[3])
            if x.strip() and x.strip() != "―"]

PROG = {}
for w in W:
    e = P.get(w[3])
    PROG[w[0]] = (e[1] if e else 0.0, e[0] if e else "未着手", w[3])

SA = [(i, k) for i, k in enumerate(K, 1) if k[5] in ("S", "A")]
SURVEY = {i for i, k in SA if _ids(k) and all(x.startswith("Ⅰ") for x in _ids(k))}
PLAN = {i for i, k in SA} - SURVEY

# 9
miss9 = sorted(PLAN - set(IMPACT))
chk(9, "未解決の優先度S・Aへの影響度の付与", not miss9,
    f"影響度なし {['K-%03d' % i for i in miss9]}" if miss9
    else f"計画策定側{len(PLAN)}件に判定／調査工程側{len(SURVEY)}件は機械で影響度5")

# 10
bad10 = [f"K-{i:03d}" for i, v in IMPACT.items() if v[0] not in LEVEL]
bad10 += [f"K-{i:03d}（調査工程）" for i in IMPACT if i in SURVEY]
chk(10, "影響度の値と対象の正しさ", not bad10, "／".join(bad10) if bad10
    else f"影響度0〜4の{len(IMPACT)}件すべてが定義済みの値")

# 11
bad11 = set()
for i, v in IMPACT.items():
    for m in re.findall(r"[０-９Ⅰ-Ⅴ]+-\d+", v[1]):
        if m not in IDX:
            bad11.add(f"K-{i:03d}:{m}")
chk(11, "「止まる対象」のWBS番号の実在", not bad11,
    f"不明 {sorted(bad11)}" if bad11 else "影響度の判定54件の参照はすべて実在")

# 12
bw, dup = [], []
seen12 = {}
for no, nm, wids, _ in BUNDLE:
    for x in wids:
        if x not in IDX:
            bw.append(f"{no}:{x}")
        if x in seen12:
            dup.append(f"{x}（{seen12[x]}と{no}）")
        seen12[x] = no
uncovered = [w[0] for w in W if w[0] not in seen12]
chk(12, "照会の束の網羅と非重複", not bw and not dup and not uncovered,
    "／".join(bw + dup + [f"束に属さない未完了 {uncovered}" if uncovered else ""]).strip("／")
    if (bw or dup or uncovered) else f"{len(BUNDLE)}束でWBS {len(W)}件をすべて覆い重複なし")

# 13
bad13 = []
STOP = {x for i, v in IMPACT.items() if v[0] == 1
        for x in re.findall(r"[０-９Ⅰ-Ⅴ]+-\d+", v[1])}
for no, wid, doing, judge, hrs in READY:
    if wid not in IDX:
        bad13.append(f"順位{no}:{wid} が実在しない")
        continue
    if PROG[wid][0] >= 1.0:
        bad13.append(f"順位{no}:{wid} は完了済み")
    if wid in STOP:
        bad13.append(f"順位{no}:{wid} は影響度1（停止）の項目で止まっている")
nos = [r[0] for r in READY]
if nos != list(range(1, len(nos) + 1)):
    bad13.append(f"順位が1からの連番でない {nos}")
chk(13, "翌営業日の作業の妥当性", not bad13, "／".join(bad13) if bad13
    else f"{len(READY)}件すべてが未完了・ペンディングなし")

# 14
ready_ids = {r[1] for r in READY}
bad14 = [w[0] for w in W
         if PROG[w[0]][0] < 1.0 and PROG[w[0]][1] != "対象外"
         and not w[0].startswith("Ⅰ")
         and w[0] not in ready_ids and w[0] not in HOLD_REASON]
over14 = [x for x in HOLD_REASON if x not in IDX or PROG.get(x, (0,))[0] >= 1.0
          or x in ready_ids]
chk(14, "未完了のWBSの扱いの明示", not bad14 and not over14,
    "／".join([f"扱い未記載 {bad14}" if bad14 else "", f"不要な理由 {over14}" if over14 else ""]).strip("／")
    if (bad14 or over14) else f"READY {len(ready_ids)}件＋理由 {len(HOLD_REASON)}件で網羅")


# 15
bad15 = []
for nm_, e in P.items():
    st_, pr_, s0, s1, _note = e
    if st_ == "完了" and not s1:
        bad15.append(f"{nm_}：完了に実績完了日がない")
    if st_ == "着手" and not s0:
        bad15.append(f"{nm_}：着手に実績開始日がない")
    if st_ == "未着手" and (s0 or s1):
        bad15.append(f"{nm_}：未着手に実績日がある")
    if s1 and not s0:
        bad15.append(f"{nm_}：実績完了日だけがある")
    if s0 and s1 and s1 < s0:
        bad15.append(f"{nm_}：実績完了日が実績開始日より前")
    if st_ == "完了" and pr_ != 1.00:
        bad15.append(f"{nm_}：完了なのに進捗が{pr_}")
chk(15, "進捗の実績日の整合", not bad15,
    "／".join(bad15) if bad15 else f"進捗反映{len(P)}件すべて整合")


print()
if NG:
    print(f"不適合 {len(NG)}件")
    for n in NG:
        print("  -", n)
    sys.exit(1)
print(f"適合：仕様書の要求事項 {len(S)}件／WBS {len(W)}件／確認事項の参照 {refs_k}件／影響度の判定 {len(IMPACT)}件／照会の束 {len(BUNDLE)}束／翌営業日の作業 {len(READY)}件")
