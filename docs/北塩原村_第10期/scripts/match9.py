# -*- coding: utf-8 -*-
"""第9期報告書から設問文を抽出し、第10期案の各設問と突合して
   「第9期に対応設問があるか（＝時系列比較の可否）」を判定する。"""
import re, csv, unicodedata

RPT = "txt_北塩原村　アンケート調査結果報告書.txt"

def norm(s):
    s = unicodedata.normalize("NFKC", str(s))
    s = re.sub(r"[（）()【】「」、。,\.・\s☑○※]", "", s)
    s = re.sub(r"(あてはまるものすべてに|は一つ|は１つ|いくつでも|複数回答可|複数選択可|１つを選択|ご回答ください|お答えください)", "", s)
    return s

# 第9期の設問文を収集
q9 = []
for ln in open(RPT, encoding="utf-8"):
    s = ln.strip()
    m = re.match(r"^[（(][0-9０-９１-９]{1,2}[）)](.+?)(?:（☑|$)", s)
    if m and len(m.group(1)) > 5:
        q9.append(("ニーズ", m.group(1)))
    m = re.match(r"^問\s?[0-9０-９]+(?:-[0-9]+)?\s*(.+?)(?:（☑|$)", s)
    if m and len(m.group(1)) > 5:
        q9.append(("在宅", m.group(1)))
q9n = [(k, norm(v), v) for k, v in q9]

def find(text, kind):
    t = norm(text)
    if len(t) < 5:
        return ""
    best, bestscore = "", 0.0
    for k, n, orig in q9n:
        if kind and k != kind:
            continue
        if not n:
            continue
        # 包含 or 文字集合の重なり
        if t[:14] and t[:14] in n:
            return orig[:60]
        inter = len(set(t) & set(n))
        score = inter / max(len(set(t)), 1)
        if score > bestscore:
            best, bestscore = orig, score
    return best[:60] if bestscore >= 0.80 else ""

FIELDS = ["資料", "章", "設問番号", "設問文",
          "令和7年8月版の対応項目", "必須/オプション/村独自", "第9期(令和4年8月版)との異同",
          "判定", "対応要否・メモ"]
rows = list(csv.DictReader(open("突合ワークシート.csv", encoding="utf-8-sig")))
hit = 0
for r in rows:
    kind = "ニーズ" if r["資料"].startswith("資料B") else "在宅"
    m = find(r["設問文"], kind)
    if m:
        r["第9期(令和4年8月版)との異同"] = f"第9期に対応設問あり：{m}"
        hit += 1
    else:
        r["第9期(令和4年8月版)との異同"] = "第9期に対応設問を自動検出できず（要目視）"

with open("突合ワークシート.csv", "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=FIELDS)
    w.writeheader(); w.writerows(rows)

print(f"第9期設問の抽出数 = {len(q9)}")
print(f"自動照合ヒット = {hit} / {len(rows)}")
print("\n--- 第9期に対応が見つからなかった設問（＝第10期での新設・改変の候補）---")
for r in rows:
    if "検出できず" in r["第9期(令和4年8月版)との異同"]:
        print(f"  [{r['資料'][:5]}] {r['章'][:20]:22} {r['設問番号']:8} {r['設問文'][:52]}")
