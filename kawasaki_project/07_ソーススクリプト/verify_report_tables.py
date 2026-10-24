# -*- coding: utf-8 -*-
"""報告書の単純集計表を、個票から作った正解値と機械的に突き合わせる。

使い方
  python3 verify_report_tables.py <報告書のdumpテキスト> <正解値json> <出力txt>

表の直前のキャプション（[Pn]<style> 見出し）と、表内に混ざる設問文の行から
設問を特定し、正解値の同じ設問の件数・割合と比較する。
選択肢ラベルだけでは同じ選択肢を持つ設問（「はい／いいえ」等）を区別できないため、
まず設問文で照合し、補助的に選択肢の一致率を使う。
"""
import json
import re
import sys
import unicodedata


def norm(s):
    s = unicodedata.normalize("NFKC", str(s))
    s = re.sub(r"[\s　]+", "", s)
    s = re.sub(r"[（）()［］\[\]「」・,、。／/？?]", "", s)
    return s.replace("～", "-").replace("〜", "-")


def bigrams(s):
    s = norm(s)
    return {s[i:i + 2] for i in range(len(s) - 1)} or {s}


def sim(a, b):
    A, B = bigrams(a), bigrams(b)
    if not A or not B:
        return 0.0
    return len(A & B) / min(len(A), len(B))


def parse_tables(path):
    """(番号, キャプション, [行のセル列]) を返す"""
    lines = open(path, encoding="utf-8").read().splitlines()
    tables, cur, cap = [], None, ""
    for line in lines:
        m = re.match(r"^--- TABLE (\d+) \(", line)
        if m:
            cur = {"no": int(m.group(1)), "cap": cap, "rows": []}
            continue
        if line.startswith("--- /TABLE"):
            if cur:
                tables.append(cur)
            cur = None
            continue
        if cur is not None:
            cur["rows"].append([c.strip() for c in line.split(" | ")])
            continue
        m = re.match(r"^\[P\d+\](?:\[図[^\]]*\])?<[^>]*>\s*(.+)$", line)
        if m and m.group(1).strip():
            cap = m.group(1).strip()
    return tables


NUM = re.compile(r"^([0-9,]+)\s*件?$")
PCT = re.compile(r"^([0-9]+(?:\.[0-9]+)?)\s*[%％]$")
NHINT = re.compile(r"[nN]\s*[=＝]\s*([0-9,]+)")


def row_values(cells):
    if not cells:
        return None
    cnt = pct = None
    for c in cells[1:]:
        c = c.strip()
        if cnt is None and NUM.match(c):
            cnt = int(NUM.match(c).group(1).replace(",", ""))
            continue
        if pct is None and PCT.match(c):
            pct = float(PCT.match(c).group(1))
    if cnt is None and pct is None:
        return None
    return cells[0], cnt, pct


def match_question(cap, inner, labels, questions):
    """設問文（キャプション＋表内の設問行）で照合し、選択肢一致率で補正する"""
    ls = {norm(x) for x in labels if x}
    best, score = None, 0.0
    for q in questions:
        s_txt = max(sim(cap, q["text"]), *(sim(t, q["text"]) for t in inner)) \
            if inner else sim(cap, q["text"])
        qs = {norm(o["label"]) for o in q["options"]}
        s_opt = len(ls & qs) / max(len(qs), 1) if qs else 0.0
        sc = s_txt * 0.75 + s_opt * 0.25
        if sc > score:
            best, score = q, sc
    return (best, score) if score >= 0.5 else (None, score)


def main():
    src, truth, dst = sys.argv[1], sys.argv[2], sys.argv[3]
    T = json.load(open(truth, encoding="utf-8"))
    qs = T["questions"]
    tables = parse_tables(src)

    out, issues = [], []
    n_ok = n_ng = n_skip = 0
    unmatched = []

    for t in tables:
        vals = [v for v in (row_values(r) for r in t["rows"]) if v]
        inner = [r[0] for r in t["rows"] if len(r) == 1 and r[0]]
        if len(vals) < 2:
            n_skip += 1
            continue
        q, sc = match_question(t["cap"], inner, [v[0] for v in vals], qs)
        if q is None:
            unmatched.append(f"表{t['no']}（{t['cap'][:40]}／一致度{sc:.2f}）")
            n_skip += len(vals)
            continue
        by = {norm(o["label"]): o for o in q["options"]}
        out.append(f"［表{t['no']}］{t['cap'][:44]}")
        out.append(f"　→ 正解値：{q['group']} {q['text'][:44]}（n={q['n']}・"
                   f"一致度{sc:.2f}）")
        # 見出しに書かれた n
        for txt in [t["cap"]] + inner:
            m = NHINT.search(unicodedata.normalize("NFKC", txt))
            if m:
                v = int(m.group(1).replace(",", ""))
                if v != q["n"]:
                    issues.append(f"表{t['no']}【n】報告書 n={v} ≠ 正解 n={q['n']}"
                                  f"（{q['group']} {q['text'][:26]}）")
                    out.append(f"　✗ 見出しの n：{v} ／ 正解 {q['n']}")
                    n_ng += 1
                else:
                    n_ok += 1
                break
        for label, cnt, pct in vals:
            key = norm(label)
            if key in ("有効回答数n", "有効回答数", "合計", "計", "n", "総数"):
                if cnt is not None and cnt != q["n"]:
                    issues.append(f"表{t['no']}【有効回答数】報告書 {cnt} ≠ 正解 {q['n']}"
                                  f"（{q['group']} {q['text'][:26]}）")
                    out.append(f"　✗ 有効回答数：{cnt} ／ 正解 {q['n']}")
                    n_ng += 1
                else:
                    n_ok += 1
                continue
            o = by.get(key)
            if o is None:
                cand = [(sim(label, x["label"]), x) for x in q["options"]]
                cand.sort(key=lambda z: -z[0])
                if cand and cand[0][0] >= 0.7:
                    o = cand[0][1]
                else:
                    out.append(f"　? 選択肢「{label[:26]}」を正解値に見つけられず")
                    n_skip += 1
                    continue
            bad = []
            if cnt is not None and cnt != o["count"]:
                bad.append(f"件数 {cnt}→正解{o['count']}")
            if pct is not None and o["pct"] is not None and abs(pct - o["pct"]) > 0.051:
                bad.append(f"割合 {pct}%→正解{o['pct']}%")
            if bad:
                issues.append(f"表{t['no']}「{label[:22]}」" + " ／ ".join(bad)
                              + f"（{q['group']} {q['text'][:22]}）")
                out.append(f"　✗ {label[:26]}：" + " ／ ".join(bad))
                n_ng += 1
            else:
                n_ok += 1
        out.append("")

    head = [f"照合対象：{src}",
            f"正解値：{truth}（個票{T['records']}件・設問{len(qs)}）",
            f"表{len(tables)}／一致 {n_ok}・不一致 {n_ng}・照合できず {n_skip}", ""]
    if issues:
        head.append(f"■ 不一致 {len(issues)}件")
        head += ["　" + i for i in issues]
        head.append("")
    if unmatched:
        head.append(f"■ 設問を特定できなかった表 {len(unmatched)}件")
        head += ["　" + u for u in unmatched]
        head.append("")
    with open(dst, "w", encoding="utf-8") as f:
        f.write("\n".join(head + out))
    print("\n".join(head))
    print(f"→ {dst}")


if __name__ == "__main__":
    main()
