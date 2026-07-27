#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""提案書に差し込む9点の図を生成する。内容はすべて提案書本文に根拠がある。

寸法の考え方：幅1800pxを本文幅15.52cmに割り付けるので 1px≒0.086mm。
figlib.SCALE により、ここで指定する文字サイズ24が印刷時およそ8ptになる。
"""
from pathlib import Path

from figlib import (W, NAVY, SLATE, LIGHT, PALE, INK, MUTED, BORDER, WHITE,
                    S_BLUE, S_ORANGE, canvas, text, block, box,
                    labelled_box, arrow, band, save)

OUT = Path(__file__).resolve().parent.parent / "figures"
OUT.mkdir(exist_ok=True)
M = 20                      # 左右の余白
IW = W - M * 2              # 作図可能幅


# --------------------------------------------------------------- 図1
def fig1():
    im, d = canvas(360)
    heads = ["国の指針\nの整理", "現状分析\nと評価", "課題の\n特定",
             "調査の\n設計と実施", "委員会\nでの討議"]
    subs = ["2章", "3章（１）〜（６）", "3章（７）", "4章", "5章"]
    n = len(heads)
    gap, aw = 12, 26
    bw = (IW - n * (gap + aw)) / (n + 1)
    x, y, h = M, 14, 178
    for i, (hd, sb) in enumerate(zip(heads, subs)):
        box(d, x, y, bw, h, fill=LIGHT, outline=NAVY, radius=12)
        bot = block(d, x + 12, y + 20, bw - 24, hd, size=24, fill=NAVY,
                    align="center", bold=True)
        block(d, x + 12, bot + 14, bw - 24, sb, size=22, fill=MUTED, align="center")
        arrow(d, x + bw + gap / 2, y + h / 2, x + bw + gap / 2 + aw, y + h / 2,
              width=6, head=16)
        x += bw + gap + aw
    box(d, x, y, bw, h, fill=WHITE, outline=BORDER, radius=12, dash=True)
    bot = block(d, x + 12, y + 20, bw - 24, "令和9年度\n計画の策定", size=24,
                fill=NAVY, align="center", bold=True)
    block(d, x + 12, bot + 14, bw - 24, "本業務の範囲外", size=21, fill=MUTED,
          align="center")
    return save(im, OUT / "fig01_overall_flow.png")


# --------------------------------------------------------------- 図2
def fig2():
    """隣の表（障がい分野）と重複しないよう、地域福祉分野の動向だけを図にする。"""
    im, d = canvas(470)
    y, h = 14, 232
    items = [("社会福祉法\n第107条", "福祉各分野に共通する\n事項を定める計画"),
             ("平成29年\n局長連名通知", "包括的な支援体制／\n自殺対策も視野に"),
             ("孤独・孤立\n対策推進法", "令和6年4月施行"),
             ("障害者差別\n解消法の改正", "令和6年4月\n合理的配慮の義務化")]
    n = len(items)
    gap = 24
    bw = (IW - (n - 1) * gap) / n
    for i, (t, b) in enumerate(items):
        x = M + i * (bw + gap)
        box(d, x, y, bw, h, fill=LIGHT, outline=NAVY, radius=12)
        bot = block(d, x + 12, y + 20, bw - 24, t, size=24, fill=NAVY,
                    align="center", bold=True)
        block(d, x + 12, bot + 14, bw - 24, b, size=22, fill=INK, align="center")
    band(d, M, y + h + 20, IW, 96,
         "いずれも第2期計画（令和2〜6年度）の期間中に動いた事項です。"
         "障がい分野の第8期基本指針（下表）とあわせ、4章の調査項目・"
         "庁内調査項目に落とし込みます。", size=23)
    return save(im, OUT / "fig02_national_policy.png")


# --------------------------------------------------------------- 図3
def fig3():
    im, d = canvas(820)
    lw, aw = 760, 70
    rw = IW - lw - aw - 30
    y = 14
    items = [
        ("第2期地域福祉計画の評価",
         "数値目標・評価指標が未設定／進行管理の枠組み／公募委員の応募が0名"),
        ("障がい者福祉計画等の評価",
         "目標設定根拠／児童発達支援センターの記載／活動指標の充実"),
        ("前回アンケート調査の評価",
         "回収率38.3%／無回答率54.9%・27.9%／回答者構成の偏り"),
    ]
    bh, bg = 158, 20
    for i, (t, b) in enumerate(items):
        labelled_box(d, M, y + i * (bh + bg), lw, bh, t, b,
                     fill=PALE, outline=NAVY, title_size=25, body_size=23)
    total = len(items) * bh + (len(items) - 1) * bg
    arrow(d, M + lw + 12, y + total / 2, M + lw + aw - 6, y + total / 2)
    rx = M + lw + aw + 30
    rh = 400
    labelled_box(d, rx, y + (total - rh) / 2, rw, rh, "評価から導かれる6つの課題領域",
                 "① 評価体制の充実\n② 目標設定の精緻化\n③ 調査手法の精度向上\n"
                 "④ 第8期基本指針への対応\n⑤ 社協・関連計画との役割整理\n"
                 "⑥ 新規法制度・社会動向への対応",
                 fill=WHITE, outline=NAVY, title_bg=NAVY, title_size=24, body_size=24)
    band(d, M, y + total + 22, IW, 66,
         "→ 6つの課題領域それぞれを、どの調査で把握するかを4章で設計します。", size=23)
    return save(im, OUT / "fig03_evaluation_to_issues.png")


# --------------------------------------------------------------- 図4
def fig4():
    im, d = canvas(900)
    y = 16
    d.ellipse([M + 4, y + 8, M + 30, y + 34], fill=S_BLUE)
    text(d, (M + 42, y + 2), "70歳未満", size=24, fill=INK)
    d.ellipse([M + 230, y + 8, M + 256, y + 34], fill=S_ORANGE)
    text(d, (M + 268, y + 2), "70歳以上", size=24, fill=INK)
    text(d, (W - M, y + 2), "単位：%", size=22, fill=MUTED, anchor="ra")

    y0 = y + 66
    text(d, (M, y0), "回答者構成と人口構成の比較", size=25, fill=NAVY, bold=True)
    lx, bx, bw, bh = M, M + 500, IW - 500, 62
    rows = [("前回調査の回答者（n=766）", 44.1, "44.1%", "55.9%"),
            ("18歳以上人口（概算）", 35.0, "約35%", "約65%")]
    for i, (lab, share, l1, l2) in enumerate(rows):
        yy = y0 + 62 + i * (bh + 24)
        text(d, (lx, yy + bh / 2), lab, size=24, fill=INK, anchor="lm")
        wo = bw * share / 100
        d.rectangle([bx, yy, bx + wo - 2, yy + bh], fill=S_ORANGE)
        d.rectangle([bx + wo + 2, yy, bx + bw, yy + bh], fill=S_BLUE)
        text(d, (bx + wo / 2, yy + bh / 2), l1, size=24, fill=WHITE, anchor="mm")
        text(d, (bx + wo + (bw - wo) / 2, yy + bh / 2), l2, size=24, fill=WHITE,
             anchor="mm")
    gy = y0 + 62
    d.line([bx + bw * 0.35, gy - 12, bx + bw * 0.441, gy - 12], fill=MUTED, width=5)
    text(d, (bx + bw * 0.395, gy - 52), "9ポイントの偏り", size=22, fill=MUTED,
         anchor="ma")

    y1 = y0 + 310
    text(d, (M, y1), "年代差が大きい設問ほど、回答者構成の偏りが全体値に効く",
         size=25, fill=NAVY, bold=True)
    ax = M + 620
    aw2 = IW - 620 - 350
    scale = lambda v: ax + aw2 * v / 80.0
    rows = [("民生委員・児童委員を\n知っている", 44.8, 77.1, "32.3"),
            ("暮らしやすいと思う", 59.8, 69.2, "9.4"),
            ("避難場所を知らない", 11.2, 7.4, "3.8")]
    top, step = y1 + 106, 100
    axis_bottom = top + (len(rows) - 1) * step + 36
    for g in (0, 20, 40, 60, 80):
        d.line([scale(g), y1 + 58, scale(g), axis_bottom], fill=(228, 228, 228), width=2)
        text(d, (scale(g), axis_bottom + 8), f"{g}", size=22, fill=MUTED, anchor="ma")
    for i, (lab, under, over, gapv) in enumerate(rows):
        yy = top + i * step
        nl = lab.count("\n") + 1
        block(d, M, yy - 16 - (nl - 1) * 21, 600, lab, size=24, fill=INK, lh=1.3)
        x1, x2 = scale(under), scale(over)
        d.line([min(x1, x2), yy, max(x1, x2), yy], fill=(210, 210, 210), width=7)
        d.ellipse([x1 - 13, yy - 13, x1 + 13, yy + 13], fill=S_BLUE)
        d.ellipse([x2 - 13, yy - 13, x2 + 13, yy + 13], fill=S_ORANGE)
        text(d, (x1, yy - 48), f"{under}", size=22, fill=S_BLUE, anchor="ma")
        text(d, (x2, yy - 48), f"{over}", size=22, fill=S_ORANGE, anchor="ma")
        text(d, (W - M, yy), f"年代差 {gapv}ポイント", size=23, fill=INK, anchor="rm")
    band(d, M, axis_bottom + 56, IW, 96,
         "年代差32ポイントの設問では人口構成比での補正が3.6ポイント効き、"
         "年代差が小さい設問では1ポイント未満にとどまります（補正後の値は本文の表）。",
         fill=PALE, color=INK, size=23)
    return save(im, OUT / "fig04_response_bias.png")


# --------------------------------------------------------------- 図5
def fig5():
    im, d = canvas(760)
    cw, gap = (IW - 72) / 3, 36
    y, h = 14, 480
    cols = [
        ("町民アンケート\n（2,000票）",
         "・暮らしの困りごと\n・相談先の認知度\n・孤独を感じる頻度\n"
         "・支え合いの実感／幸福感\n・策定委員会への参加意向\n・K6（こころの状態）",
         "本人にしか答えられない主観", NAVY),
        ("障がい者アンケート\n（500票）",
         "・手帳種別ごとの生活実態\n・使えていないサービス\n・就労意向（分岐設計）\n"
         "・差別を受けた経験\n・記入方法の区分",
         "当事者ご本人の声", SLATE),
        ("庁内調査・関係団体調査\n社協ヒアリング",
         "・のぞまないセルフプラン率\n・強度行動障害の実数\n・4つの中核機能の充足\n"
         "・人材の確保・定着状況\n・社協の実践計画と外部評価\n・重層事業／CSW配置",
         "制度運用と組織の実態", SLATE),
    ]
    for i, (t, b, foot, col) in enumerate(cols):
        x = M + i * (cw + gap)
        labelled_box(d, x, y, cw, h, t, b, fill=WHITE, outline=col,
                     title_bg=col, title_size=24, body_size=23)
        band(d, x + 14, y + h - 80, cw - 28, 62, foot, fill=LIGHT, color=NAVY, size=23)
    band(d, M, y + h + 22, IW, 96,
         "調査票に載せるのは本人の主観に限定し、制度運用や組織の実態は庁内調査・"
         "関係団体調査で補完します。設問数を増やさないための設計です。", size=23)
    return save(im, OUT / "fig05_survey_roles.png")


# --------------------------------------------------------------- 図6
def fig6():
    im, d = canvas(680)
    cw, gap = (IW - 56) / 2, 56
    y, h = 14, 360
    labelled_box(d, M, y, cw, h, "① 層化抽出（配布の設計）",
                 "足寄小学校区 約1,400票／大誉地 約200票／\n"
                 "芽登 約200票／螺湾 約200票\n\n"
                 "小さな地区でも回答数そのものを確保することで、"
                 "地区別分析の標本誤差を縮小します。",
                 fill=WHITE, outline=NAVY, title_bg=NAVY, title_size=25, body_size=24)
    labelled_box(d, M + cw + gap, y, cw, h, "② 二段階ウェイトバック集計",
                 "（a）抽出ウェイト＝地区ごとの抽出率の逆数\n"
                 "（b）無回答調整ウェイト＝性・年齢階級別の回収率の差\n\n"
                 "地区間・属性間の代表性の偏りを補正します。"
                 "標本誤差そのものは縮小しません。",
                 fill=WHITE, outline=SLATE, title_bg=SLATE, title_size=25, body_size=24)
    band(d, M, y + h + 22, IW, 140,
         "この2つは役割が異なります。誤差の縮小は①配布の設計が、代表性の補正は"
         "②集計の設計が担います。報告書には補正前後の値・信頼区間・実効標本数を併記し、"
         "回答数が少ないセルは有意差の有無を明示します。", size=23)
    return save(im, OUT / "fig06_sampling_weighting.png")


# --------------------------------------------------------------- 図7
def fig7():
    im, d = canvas(660)
    stages = [("契約〜8月", "体制整備\n調査設計の確定"), ("8〜9月", "第1回\n委員会・協議会"),
              ("9〜10月", "調査票の発送\n庁内・関係団体調査"),
              ("10〜11月", "入力・集計・分析\nウェイトバック集計"),
              ("11〜12月上旬", "第2回\n委員会・協議会"),
              ("12月中旬〜28日", "とりまとめ\n納品")]
    n = len(stages)
    gap, aw = 10, 22
    bw = (IW - (n - 1) * (gap + aw)) / n
    y, h = 14, 226
    xs = []
    for i, (t, b) in enumerate(stages):
        x = M + i * (bw + gap + aw)
        xs.append(x)
        hl = i in (1, 4)
        box(d, x, y, bw, h, fill=LIGHT if hl else WHITE,
            outline=NAVY if hl else BORDER, radius=12)
        bot = block(d, x + 8, y + 20, bw - 16, t, size=23,
                    fill=NAVY if hl else MUTED, align="center", bold=hl)
        block(d, x + 8, bot + 12, bw - 16, b, size=22, fill=INK, align="center")
        if i < n - 1:
            arrow(d, x + bw + gap / 2, y + h / 2, x + bw + gap / 2 + aw, y + h / 2,
                  width=6, head=14)
    for i in (1, 4):
        arrow(d, xs[i] + bw / 2, y + h + 6, xs[i] + bw / 2, y + h + 42,
              color=NAVY, width=6, head=16)
    band(d, M, y + h + 54, IW, 170,
         "委員会資料の基本構成：現状の整理（基礎データ・アンケート分析・現行計画の評価）／"
         "課題の整理（横断テーマ別）／論点と選択肢の提示／第8期基本指針との整合確認／"
         "次期計画策定作業への申し送り事項。開催回数は地域福祉・障がいそれぞれ1〜2回を"
         "見積り上の前提とし、書面・オンライン開催の併用もご提案します。", size=23)
    return save(im, OUT / "fig07_committee_flow.png")


# --------------------------------------------------------------- 図8
def fig8():
    """役職表と重複しないよう、成果品が仕上がるまでの流れだけを図にする。"""
    im, d = canvas(460)
    steps = [("作成", "調査・分析担当者"), ("照査", "照査担当者\n（作成者と別担当）"),
             ("内部品質確認", "管理責任者"), ("納品", "町")]
    n = len(steps)
    gap, aw = 18, 36
    bw = (IW - (n - 1) * (gap + aw)) / n
    y, h = 14, 176
    for i, (t, b) in enumerate(steps):
        x = M + i * (bw + gap + aw)
        box(d, x, y, bw, h, fill=LIGHT if i < n - 1 else WHITE,
            outline=NAVY if i < n - 1 else BORDER, radius=12)
        bot = block(d, x + 12, y + 28, bw - 24, t, size=25, fill=NAVY,
                    align="center", bold=True)
        block(d, x + 12, bot + 14, bw - 24, b, size=23, fill=INK, align="center")
        if i < n - 1:
            arrow(d, x + bw + gap / 2, y + h / 2, x + bw + gap / 2 + aw, y + h / 2)
    band(d, M, y + h + 20, IW, 120,
         "週次の工程表・課題一覧・版管理表で進捗を共有します。個人情報は宛名情報と"
         "回答データを分離し、暗号化・アクセス権限の限定・データ消去記録の作成を"
         "徹底します。", size=23)
    return save(im, OUT / "fig08_team.png")


# --------------------------------------------------------------- 図9
def fig9():
    im, d = canvas(740)
    stages = ["国の指針\n（2章）", "評価・課題整理\n（3章）", "調査の設計・実施\n（4章）",
              "委員会での討議\n（5章）"]
    n = len(stages)
    gap, aw = 22, 30
    bw = (IW - (n - 1) * (gap + aw)) / n
    y, h = 14, 150
    xs = []
    for i, s in enumerate(stages):
        x = M + i * (bw + gap + aw)
        xs.append(x)
        box(d, x, y, bw, h, fill=LIGHT, outline=NAVY, radius=12)
        block(d, x + 12, y + 34, bw - 24, s, size=24, fill=NAVY, align="center",
              bold=True)
        if i < n - 1:
            arrow(d, x + bw + gap / 2, y + h / 2, x + bw + gap / 2 + aw, y + h / 2)
    props = {
        0: ["④ 第8期基本指針への即応"],
        1: ["① 評価可能な計画への転換", "⑤ 町・社協・関連計画の役割整理"],
        2: ["② 調査設計の刷新", "③ 障がい児の実態把握の補完",
            "⑥ 記入方法の区分（本人の声）"],
    }
    hh, vg = 112, 14
    maxrows = max(len(v) for v in props.values())
    for i, items in props.items():
        x = xs[i]
        yy = y + h + 28
        arrow(d, x + bw / 2, y + h + 4, x + bw / 2, yy - 4, width=5, head=13)
        for j, s in enumerate(items):
            box(d, x, yy + j * (hh + vg), bw, hh, fill=WHITE, outline=SLATE, radius=10)
            block(d, x + 12, yy + j * (hh + vg) + 20, bw - 24, s, size=23,
                  fill=INK, align="center", lh=1.3)
    by = y + h + 28 + maxrows * (hh + vg) + 8
    band(d, M, by, IW, 66,
         "いずれの独自提案も、最終的には5章の委員会資料に反映されます。", size=23)
    return save(im, OUT / "fig09_original_proposals.png")


FIGURES = [fig1, fig2, fig3, fig4, fig5, fig6, fig7, fig8, fig9]

if __name__ == "__main__":
    for f in FIGURES:
        print("wrote", f())
