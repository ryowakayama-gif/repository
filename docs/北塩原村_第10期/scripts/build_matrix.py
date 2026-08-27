# -*- coding: utf-8 -*-
"""資料B'（ニーズ調査）・資料C'（在宅介護実態調査）の全設問を抽出し、
   令和7年8月版様式との逐条突合用ワークシート（CSV）を生成する。"""
import re, csv

def load(p):
    return open(p, encoding="utf-8").read().splitlines()

# ---------- 資料B'（ニーズ調査） ----------
rows_b = []
cur_sec = ""
for ln in load("rev_B.txt"):
    s = ln.strip()
    if not s:
        continue
    m = re.match(r"^(問\s?\d+)\s*\|\s*(.+)$", s)
    if m:
        cur_sec = f"{m.group(1)} {m.group(2)}"
        continue
    if s.startswith("はじめに、あなたご自身についてお伺いします"):
        cur_sec = "フェイス（基本属性）"
        continue
    # 設問行：（ア）（１）（10）① など
    m = re.match(r"^([（(][ア-ンａ-ｚ0-9０-９]{1,2}[）)]|①|②)\s*(.+)$", s)
    if m and cur_sec:
        no, txt = m.group(1), m.group(2)
        if len(txt) < 4:
            continue
        rows_b.append({"資料": "資料B'（ニーズ調査）", "章": cur_sec, "設問番号": no,
                       "設問文": re.sub(r"\s+", " ", txt)[:110]})
    # 問5(1)のマトリクス項目
    m2 = re.match(r"^(①|②|③|④|⑤|⑥|⑦|⑧)\s+(.+?)\s*\|", s)
    if m2 and cur_sec.startswith("問５"):
        rows_b.append({"資料": "資料B'（ニーズ調査）", "章": cur_sec,
                       "設問番号": f"(1){m2.group(1)}", "設問文": re.sub(r"\s+", " ", m2.group(2))[:110]})

# ---------- 資料C'（在宅介護実態調査） ----------
rows_c = []
hyou = "Ａ票"
for ln in load("rev_C.txt"):
    s = ln.strip()
    if not s:
        continue
    if s.startswith("Ｂ票"):
        hyou = "Ｂ票"
        continue
    if s.startswith("Ａ票"):
        hyou = "Ａ票"
        continue
    m = re.match(r"^(問\s?[0-9０-９]+(?:-[0-9]+)?)\s*　?\s*(.+)$", s)
    if m:
        txt = re.sub(r"\s+", " ", m.group(2))
        txt = re.sub(r"（[１-９1-9]つを?選択.*?）|（複数(回答|選択)可）|（３つまで選択可）", "", txt)
        rows_c.append({"資料": "資料C'（在宅介護実態調査）", "章": hyou,
                       "設問番号": m.group(1).replace(" ", ""), "設問文": txt.strip()[:110]})
    # 問9のサービス行
    m3 = re.match(r"^([Ａ-ＬA-L])．(.+?)\s*\|", s)
    if m3:
        rows_c.append({"資料": "資料C'（在宅介護実態調査）", "章": f"{hyou}／問９",
                       "設問番号": m3.group(1) + "．",
                       "設問文": re.sub(r"\s+", " ", m3.group(2))[:110]})

FIELDS = ["資料", "章", "設問番号", "設問文",
          "令和7年8月版の対応項目", "必須/オプション/村独自", "第9期(令和4年8月版)との異同",
          "判定", "対応要否・メモ"]
with open("突合ワークシート.csv", "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=FIELDS)
    w.writeheader()
    for r in rows_b + rows_c:
        r.update({k: "" for k in FIELDS if k not in r})
        w.writerow(r)

print(f"資料B' 設問数 = {len(rows_b)}")
print(f"資料C' 設問数 = {len(rows_c)}（問９のサービス行を含む）")
print(f"合計 {len(rows_b)+len(rows_c)} 行を出力")
