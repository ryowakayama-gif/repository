#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""図01を3列×2段に組み替える。

6列（カード幅190px）では文字サイズを印刷8pt相当まで上げると1行10字しか入らず、
役割（受託者／町）を書き切れないため、3列×2段にしてカード幅を約2倍にする。
"""
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src" / "01_アンケート調査の全体フロー.svg"
DST = ROOT / "build" / "01_アンケート調査の全体フロー.svg"

W, H = 1240, 862
COLS, GAP = 3, 24
CARD_W = (W - 70 - GAP * (COLS - 1)) / COLS      # 約384px
CHEV_H, CARD_H = 60, 210
ROW_TOP = [116, 116 + CHEV_H + CARD_H + 34]

STEPS = [
    ("① 調査設計", "brand", [
        ("前回調査の検証", 0), ("施策体系との対応", 0),
        ("標本設計（対象・数）", 0), ("国・道の調査との整合", 0)]),
    ("② 調査票作成", "brand", [
        ("設問案の作成・精査", 0), ("UD・合理的配慮への対応", 0),
        ("庁内協議による修正", 2), ("委員会への報告", 2)]),
    ("③ 印刷・発送", "brand2", [
        ("対象者名簿の提供", 1), ("調査票・封筒の印刷", 0),
        ("宛名作成・封入封緘", 0), ("発送・郵送料の負担", 0)]),
    ("④ 回収・督促", "brand2", [
        ("回収（返送先は受託者）", 0), ("回収状況を町へ日次報告", 0),
        ("督促はがきの作成・発送", 0), ("問合せ窓口の設置", 0)]),
    ("⑤ 入力・点検", "accent", [
        ("入力・ダブル点検", 0), ("論理・整合チェック", 0),
        ("自由記述のテキスト化", 0), ("個票データの匿名化", 0)]),
    ("⑥ 集計・分析", "accent", [
        ("単純集計・クロス集計", 0), ("前回調査との経年比較", 0),
        ("課題の抽出・考察", 0), ("調査結果報告書の作成", 0)]),
]
ROLE = {0: ("f-brand", "受託者"), 1: ("f-surface s-border", "町"), 2: ("f-accent", "双方")}


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def build():
    src = SRC.read_text(encoding="utf-8")
    head = src[:src.index('<rect class="f-surface" width=')]
    head = head.replace('viewBox="0 0 1240 640"', f'viewBox="0 0 {W} {H}"')
    head = head.replace("調査設計から報告書作成までの6ステップと主な作業内容",
                        "調査設計から報告書作成までの6ステップと、町・受託者の役割分担")

    b = [head, f'<rect class="f-surface" width="{W}" height="{H}"/>',
         '<rect class="f-accent" x="34" y="30" width="5" height="24" rx="2.5"/>',
         '<text class="f-ink fs-23" x="50" y="50" font-weight="700" text-anchor="start">'
         'アンケート調査の全体フロー</text>',
         '<text class="f-muted fs-13_5" x="50" y="74" font-weight="400" text-anchor="start">'
         '調査設計から報告書作成までの6ステップと、町・受託者の役割分担</text>']

    # 凡例：塗り・白抜き・ひし形と文字の両方で役割を示す（色だけに依存しない）
    lx = 700
    for i in range(3):
        klass, name = ROLE[i]
        x = lx + i * 180
        if i == 2:  # 双方はひし形で形も変える
            b.append(f'<path class="{klass}" d="M{x + 7},{x * 0 + 40} L{x + 15},{48} '
                     f'L{x + 7},{56} L{x - 1},{48} Z"/>')
        else:
            b.append(f'<rect class="{klass}" x="{x}" y="41" width="14" height="14" rx="2" '
                     f'stroke-width="1.4"/>')
        b.append(f'<text class="f-muted fs-13" x="{x + 24}" y="54" text-anchor="start">'
                 f'{esc(name)}が実施</text>')

    for i, (title, tone, items) in enumerate(STEPS):
        col, row = i % COLS, i // COLS
        x = 35 + col * (CARD_W + GAP)
        y = ROW_TOP[row]
        fill = {"brand": "f-brand", "brand2": "f-brand2", "accent": "f-accent"}[tone]
        tip = 18
        b.append(f'<path class="{fill}" d="M{x},{y} L{x + CARD_W - tip},{y} '
                 f'L{x + CARD_W},{y + CHEV_H / 2} L{x + CARD_W - tip},{y + CHEV_H} '
                 f'L{x},{y + CHEV_H} L{x + tip},{y + CHEV_H / 2} Z"/>')
        b.append(f'<text class="f-onbrand fs-16" x="{x + CARD_W / 2}" y="{y + CHEV_H / 2 + 7}" '
                 f'font-weight="700" text-anchor="middle">{esc(title)}</text>')
        cy = y + CHEV_H + 12
        b.append(f'<rect class="f-surface s-border" x="{x}" y="{cy}" width="{CARD_W}" '
                 f'height="{CARD_H}" rx="9" stroke-width="1.2"/>')
        for j, (txt, role) in enumerate(items):
            ty = cy + 34 + j * 44
            klass, _ = ROLE[role]
            if role == 2:
                b.append(f'<path class="{klass}" d="M{x + 24},{ty - 13} L{x + 31},{ty - 6} '
                         f'L{x + 24},{ty + 1} L{x + 17},{ty - 6} Z"/>')
            else:
                b.append(f'<rect class="{klass}" x="{x + 18}" y="{ty - 12}" width="12" '
                         f'height="12" rx="2" stroke-width="1.4"/>')
            b.append(f'<text class="f-ink fs-13" x="{x + 38}" y="{ty}" text-anchor="start">'
                     f'{esc(txt)}</text>')

    by = ROW_TOP[1] + CHEV_H + CARD_H + 26
    b.append(f'<rect class="f-brandpale" x="35" y="{by}" width="{W - 70}" height="44" rx="8"/>')
    b.append(f'<text class="f-brand fs-14" x="{W / 2}" y="{by + 28}" font-weight="700" '
             f'text-anchor="middle">足寄町の履行期間に対応した工程　約5か月（8月契約→12月28日納品）</text>')
    ny = by + 56
    b.append(f'<rect class="f-surfalt s-border" x="35" y="{ny}" width="{W - 70}" height="66" '
             f'rx="8" stroke-width="1.2"/>')
    b.append(f'<text class="f-brand fs-13" x="55" y="{ny + 26}" font-weight="700" '
             f'text-anchor="start">【個人情報の取扱い】</text>')
    b.append(f'<text class="f-ink fs-12_5" x="55" y="{ny + 50}" text-anchor="start">'
             f'宛名データ・回収票は施錠保管庫で管理し、作業は受託者事業所内に限定。'
             f'完了後は町の指示に従い返却・消去し、消去証明書を提出します。</text>')
    b.append("</svg>")
    return "\n".join(b)


if __name__ == "__main__":
    DST.write_text(build(), encoding="utf-8")
    print("wrote", DST)
