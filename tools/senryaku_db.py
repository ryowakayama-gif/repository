# -*- coding: utf-8 -*-
"""
経営戦略（水道・下水道）改訂時期調査 データストア

- 調査結果を research/findings.jsonl に追記する（(自治体, 事業, 計画名) で上書き更新）
- 和暦表記（R3, H31 等）から西暦年度への変換を担う
- 判定ロジック（来年で計画期間満了 / 策定から5年経過）を一元管理する
"""

import json
import os
import re

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FINDINGS = os.path.join(BASE_DIR, "research", "findings.jsonl")

# 調査基準日：2026-07-30 ＝ 令和8年度（2026年度）
BASE_FY = 2026
BASE_FY_LABEL = "令和8年度（2026年度）"

# ============================================================
# 和暦 → 西暦年度
# ============================================================
ERA_OFFSET = {"R": 2018, "令和": 2018, "H": 1988, "平成": 1988, "S": 1925, "昭和": 1925}


def to_fy(v):
    """'R3' / '令和3' / 'H31' / 2021 / '2021' → 西暦年度(int)。不明は None。"""
    if v is None or v == "":
        return None
    if isinstance(v, int):
        return v
    s = str(v).strip()
    m = re.match(r"^(R|H|S|令和|平成|昭和)\s*(\d+|元)$", s)
    if m:
        era, num = m.group(1), m.group(2)
        n = 1 if num == "元" else int(num)
        return ERA_OFFSET[era] + n
    m = re.match(r"^(19|20)\d{2}$", s)
    if m:
        return int(s)
    return None


def fy_label(fy):
    """2021 → '令和3年度(2021)'"""
    if fy is None:
        return ""
    if fy >= 2019:
        return f"令和{fy - 2018}年度({fy})"
    if fy >= 1989:
        return f"平成{fy - 1988}年度({fy})"
    return str(fy)


# ============================================================
# 判定
# ============================================================
def judge(rec):
    """計画期間満了・経過年数から改訂該当区分を判定して rec に追記する。"""
    end = rec.get("end_fy")
    made = rec.get("made_fy")
    revised = rec.get("revised_fy")
    eff = revised if revised is not None else made  # 直近の策定・改定年度

    hits = []

    # 判定A：来年で計画期間がおわるもの
    expiry = ""
    if end is not None:
        if end < BASE_FY:
            expiry = f"既に満了（{fy_label(end)}末）"
            hits.append("期間満了済")
        elif end == BASE_FY:
            expiry = "今年度末満了（2027年3月）"
            hits.append("今年度末満了")
        elif end == BASE_FY + 1:
            expiry = "来年度末満了（2028年3月）"
            hits.append("来年度末満了")
        else:
            expiry = f"満了は{fy_label(end)}末"
    rec["expiry_judge"] = expiry

    # 判定B：作成から5年経過する計画
    elapsed = None
    elapsed_judge = ""
    if eff is not None:
        elapsed = BASE_FY - eff
        if elapsed >= 6:
            elapsed_judge = f"{elapsed}年経過（5年超）"
            hits.append("5年超経過")
        elif elapsed == 5:
            elapsed_judge = "5年経過"
            hits.append("5年経過")
        elif elapsed == 4:
            elapsed_judge = "4年経過（翌年度に5年）"
            hits.append("翌年度5年")
        else:
            elapsed_judge = f"{elapsed}年経過"
    rec["elapsed_years"] = elapsed
    rec["elapsed_judge"] = elapsed_judge

    # 優先度
    if "期間満了済" in hits or "5年超経過" in hits or "今年度末満了" in hits:
        pri = "A：改訂着手が必要"
    elif "5年経過" in hits or "来年度末満了" in hits:
        pri = "B：来年度改訂の対象"
    elif "翌年度5年" in hits:
        pri = "C：翌年度以降に検討"
    elif not hits and (end is None and eff is None):
        pri = "－：情報未確認"
    else:
        pri = "C：翌年度以降に検討"
    rec["priority"] = pri
    rec["hit"] = "／".join(hits)
    return rec


# ============================================================
# 追記
# ============================================================
FIELDS = [
    "pref", "area", "muni", "jigyo", "plan_name",
    "made", "made_fy", "start_fy", "end_fy", "revised", "revised_fy",
    "expiry_judge", "elapsed_years", "elapsed_judge", "hit", "priority",
    "source", "note", "confidence",
]


def add(records, pref="北海道", area=None, muni=None):
    """records: dict のリスト。area/muni は共通指定可。"""
    os.makedirs(os.path.dirname(FINDINGS), exist_ok=True)
    existing = load()
    index = {(r["muni"], r["jigyo"], r.get("plan_name", "")): i for i, r in enumerate(existing)}

    added, updated = 0, 0
    for src in records:
        rec = dict(src)
        rec.setdefault("pref", pref)
        if area:
            rec.setdefault("area", area)
        if muni:
            rec.setdefault("muni", muni)
        for k_in, k_out in (("made", "made_fy"), ("revised", "revised_fy"),
                            ("start", "start_fy"), ("end", "end_fy")):
            if k_in in rec and k_out not in rec:
                rec[k_out] = to_fy(rec[k_in])
        for k in ("start_fy", "end_fy", "made_fy", "revised_fy"):
            rec[k] = to_fy(rec.get(k))
        rec.setdefault("confidence", "中")  # 高=一次資料明記 / 中=公式サイト要約 / 低=推定
        rec.setdefault("note", "")
        rec.setdefault("plan_name", "")
        judge(rec)
        key = (rec["muni"], rec["jigyo"], rec.get("plan_name", ""))
        if key in index:
            existing[index[key]] = rec
            updated += 1
        else:
            index[key] = len(existing)
            existing.append(rec)
            added += 1

    with open(FINDINGS, "w", encoding="utf-8") as f:
        for r in existing:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
    print(f"add={added} update={updated} total={len(existing)} "
          f"自治体数={len({r['muni'] for r in existing})}")
    return existing


def load():
    if not os.path.exists(FINDINGS):
        return []
    out = []
    with open(FINDINGS, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                out.append(json.loads(line))
    return out


def done_munis():
    return {r["muni"] for r in load()}


def remaining(csv_path=None):
    """未調査の自治体を振興局順に返す。"""
    import csv
    csv_path = csv_path or os.path.join(BASE_DIR, "research", "hokkaido_municipalities.csv")
    done = done_munis()
    out = []
    with open(csv_path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if row["自治体"] not in done:
                out.append((row["振興局"], row["自治体"]))
    return out
