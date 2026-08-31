# -*- coding: utf-8 -*-
"""soan_content.py から 素案のMarkdownとdocx用JSONを生成する"""
import os, sys, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import soan_content as S

MD = "/home/user/repository/docs/北塩原村_第10期/18_計画素案.md"
JS = "/tmp/soan.json"

out = [f"# {S.TITLE}・{S.TITLE2}　{S.DRAFT}", "",
       f"**計画期間：{S.SUBTITLE}**　／　{S.ISSUER}　／　{S.DATE}", "",
       "> 本書は素案です。【要確認】は村との協議事項、【要設定】は目標値の設定待ち、",
       "> 編集注記（⚙）は素案段階の申し送りで、計画確定時に削除します。", "", "---", ""]

for ch in S.CH:
    out += [f"# {ch['no']}　{ch['title']}", ""]
    for sec in ch["sections"]:
        out += [f"## {sec['no']}　{sec['title']}", ""]
        for b in sec["blocks"]:
            t = b["t"]
            if t == "p":
                out += [b["v"], ""]
            elif t == "h3":
                out += [f"### {b['v']}", ""]
            elif t == "bullets":
                out += [f"- {x}" for x in b["v"]] + [""]
            elif t == "note":
                out += [f"> ⚙ **編集注記**：{b['v']}", ""]
            elif t in ("table", "kpi"):
                head = b["head"]
                out += ["| " + " | ".join(head) + " |",
                        "|" + "|".join(["---"] * len(head)) + "|"]
                out += ["| " + " | ".join(str(c) for c in r) + " |" for r in b["rows"]]
                out += [""]
    out += ["---", ""]

open(MD, "w", encoding="utf-8").write("\n".join(out))
json.dump({"title": S.TITLE, "title2": S.TITLE2, "subtitle": S.SUBTITLE,
           "draft": S.DRAFT, "issuer": S.ISSUER, "date": S.DATE, "chapters": S.CH},
          open(JS, "w", encoding="utf-8"), ensure_ascii=False)
print(f"Markdown: {MD}  ({len(out)}行)")
print(f"JSON    : {JS}")
