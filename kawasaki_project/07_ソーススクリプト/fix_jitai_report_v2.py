# -*- coding: utf-8 -*-
"""
川崎町 介護実態調査結果報告書 R8.9.2版 仕上げ

  1. 第7章の全国比較の注記が接続の崩れた文になっていたため、文を整える（J-8の仕上げ）
  2. 第7章冒頭の介護者年齢に n を併記
  3. 第11章のKPI注記を章末（紙おむつ事業の注記の後）に移し、達成状況の書き方を明確にする
  4. 第10章「施設・在宅の選択支援」に母集団換算値への参照を追加（J-4との接続）
"""
import re
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

BASE = Path(__file__).resolve().parent.parent
DOC = BASE / "01_第10期_最新版成果品/川崎町_介護実態調査結果報告書_R8.9.2版.docx"

log = []


def find_par(doc, needle):
    for p in doc.paragraphs:
        if needle in p.text:
            return p
    raise KeyError(needle)


def set_text(par, text):
    ts = list(par._p.iter(qn("w:t")))
    ts[0].text = text
    ts[0].set(qn("xml:space"), "preserve")
    for t in ts[1:]:
        t.text = ""


def sub_text(par, old, new):
    ts = list(par._p.iter(qn("w:t")))
    whole = "".join(t.text or "" for t in ts)
    out = re.sub(re.escape(old), lambda m: new, whole)
    if out == whole:
        return False
    ts[0].text = out
    ts[0].set(qn("xml:space"), "preserve")
    for t in ts[1:]:
        t.text = ""
    return True


def main():
    doc = Document(str(DOC))

    # 1. 第7章の注記
    p = find_par(doc, "老老介護」の割合が61.9％に達しており、ただし全国値は")
    set_text(
        p,
        "※厚生労働省「2025（令和7）年国民生活基礎調査」では、同居の主な介護者と要介護者等がいずれも65歳以上である"
        "「老老介護」の割合が61.9％に達している。ただしこれは介護者・要介護者がともに65歳以上である世帯の割合であり、"
        "本調査で算出できる最も近い区分は主な介護者の70歳以上42.1％（n=76・32件）であるため、両者は直接比較できない。"
        "いずれにせよ介護者自身の高齢化は進んでおり、介護者の心身の負担にも配慮した支援が求められる。",
    )
    log.append("[仕上げ1] 第7章の全国比較の注記を、接続の整った文に書き改めた")

    # 2. 第7章冒頭に n を併記
    p = find_par(doc, "最も多い回答は「60代」「70代」で、ともに26.3％である")
    sub_text(
        p,
        "最も多い回答は「60代」「70代」で、ともに26.3％である。60代以上（60代・70代・80歳以上の合計）で68.4％を占めており、",
        "最も多い回答は「60代」「70代」で、ともに26.3％（n=76・各20件）である。"
        "60代以上（60代・70代・80歳以上の合計）で68.4％（52件）、70歳以上で42.1％（32件）を占めており、",
    )
    log.append("[仕上げ2] 第7章冒頭の介護者年齢に n と実件数、70歳以上の値を併記")

    # 3. KPI注記の文言を整え、第11章の末尾へ移す
    p = find_par(doc, "第９期計画で設定した目標値（KPI）の達成状況")
    set_text(
        p,
        "※ 第９期計画で設定した目標値（KPI）の達成状況、及び保険者機能強化推進交付金等の評価結果については、"
        "第１回策定委員会資料において別途整理している。本報告書は、今回の調査結果から読み取れる課題の整理に限定している。"
        "なお、第９期計画のアウトカム指標は「調整済み認定率を令和5年の水準（17.6％）に維持する」ことであり、"
        "令和7年度末の認定率は17.3％（561人／3,240人）で基準値を上回っていないため、現時点では達成の見込みである。"
        "ただし指標は年齢調整済みの認定率と定義されているため、確定値は地域包括ケア「見える化」システムにより確認する。",
    )
    tail = find_par(doc, "高齢者紙おむつ等支給事業の支給額・対象要件は")
    p._p.getparent().remove(p._p)
    tail._p.addnext(p._p)
    log.append("[仕上げ3] 第11章のKPI注記を章末に移し、認定率の達成状況の書き方を明確にした")

    # 4. 第10章③に母集団換算への参照を追加
    p = find_par(doc, "施設検討14.4％、申込済み33.8％と4割超で施設入所の意思があり")
    sub_text(
        p,
        "施設検討14.4％、申込済み33.8％と4割超で施設入所の意思があり、",
        "施設検討14.4％、申込済み33.8％と4割超で施設入所の意思があり（４章のとおり在宅認定者約369人に換算すると"
        "申込済み約125人・検討中約53人）、",
    )
    log.append("[仕上げ4] 第10章③に母集団換算値（約125人／約53人）への参照を追加")

    doc.save(str(DOC))
    print("\n".join(log))


if __name__ == "__main__":
    main()
