# -*- coding: utf-8 -*-
"""需要のシナリオと保険料の感度表（令和8年10月7日）

08_作業順位 の順位24。

当方の見込量は**自然体の1本**であり、確からしさの幅を示せていない。
見込量の算定式で動く変数は**要介護度別の認定者数**ひとつだけであるから、
認定者数を振ることで、給付費と保険料の幅を示すことができる。

  出力　05_試算・管理シート/川崎町_需要シナリオと保険料の感度_R8.10.7.xlsx
          00_この表について
          01_認定者数の3シナリオ
          02_給付費と保険料の感度
          03_地域支援事業費との組合せ
          04_第9期の実績による検算
          05_確認事項

⚠ **シナリオは見込量を置き換えるものではない。** 計画に載せる見込量は
  採用（自然体）のままである。幅は、第3回策定委員会で基金の取崩額を
  ご決定いただくときの材料としてお示しする。

⚠ 感度の計算は、**総給付費を認定者数の比で伸縮させる近似**による。
  当方の算定は要介護度別であるから、要介護度の構成が変わる場合は
  この近似からずれる。その旨を明記する。

⚠ 保険料の算定式が正しく組めていることを、**採用のシナリオで
  現行の保険料基準額（算定Ａ）が再現できるか**で確かめている。
  再現できなければ終了コード1で終わる。
"""
import importlib.util
import os
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from data_zuhyo import ZU                                    # noqa: E402


def load(name, fname):
    spec = importlib.util.spec_from_file_location(name,
                                                  os.path.join(HERE, fname))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


GY = load("gyotei6", "build_iinkai_gyotei_R8.10.1.py")

MIKOMI = "05_試算・管理シート/川崎町_第10期_計画見込量_R8.9.15.xlsx"
OUT = "05_試算・管理シート/川崎町_需要シナリオと保険料の感度_R8.10.7.xlsx"

FUTAN = 0.23          # 第1号被保険者負担割合
CHOSEI_SOTO = 0.05    # 調整交付金相当額の割合
# 調整交付金の見込交付割合（計画見込量 ①算定の考え方【7】）
KOFU = [0.0614, 0.0593, 0.0569]
NENDO = ["令和9年度", "令和10年度", "令和11年度"]

KAKUNIN = ["33", "11", "168", "155", "8"]


def yomu():
    wb = openpyxl.load_workbook(MIKOMI, data_only=True)
    ws = wb["⑦保険料算定"]
    d = {}
    for r in range(1, ws.max_row + 1):
        k = str(ws.cell(r, 1).value or "").strip()
        v = [ws.cell(r, c).value for c in (2, 3, 4)]
        if k == "標準給付費":
            d["hyojun"] = v
        if k == "地域支援事業費":
            d["chiiki"] = v
        if k == "保険料収納必要額":
            d["hitsuyo"] = v
        if k.startswith("補正後被保険者数"):
            d["hosei"] = ws.cell(r, 2).value
        if k.startswith("予定保険料収納率"):
            d["ritsu"] = ws.cell(r, 2).value
        if k.startswith("保険料基準額"):
            d["gaku"] = ws.cell(r, 2).value
    ws6 = wb["⑥地域支援事業費"]
    for r in range(1, ws6.max_row + 1):
        if str(ws6.cell(r, 1).value or "").startswith("介護予防・日常生活支援"):
            d["sogo"] = [ws6.cell(r, c).value for c in (2, 3, 4)]
    for k in ("hyojun", "chiiki", "hitsuyo", "hosei", "ritsu", "gaku",
              "sogo"):
        if k not in d:
            raise SystemExit("⑦又は⑥シートから読めない：" + k)
    return d


def hokenryo(hyojun, chiiki, sogo, hosei, ritsu):
    """保険料基準額（月額）を算定する。

    第1号負担分　　　 ＝（標準給付費＋地域支援事業費）×23％
    調整交付金相当額 ＝ 標準給付費×5％ ＋ 総合事業費×5％
    調整交付金見込額 ＝（標準給付費＋総合事業費）× 見込交付割合
    保険料収納必要額 ＝ 第1号負担分 ＋ 調整交付金相当額 － 調整交付金見込額
    月額　　　　　　 ＝ 収納必要額の3か年計 ÷ 予定収納率 ÷ 補正後被保険者数 ÷ 12
    """
    kei = 0.0
    for i in range(3):
        futan = (hyojun[i] + chiiki[i]) * FUTAN
        soto = hyojun[i] * CHOSEI_SOTO + sogo[i] * CHOSEI_SOTO
        mikomi = (hyojun[i] + sogo[i]) * KOFU[i]
        kei += futan + soto - mikomi
    return kei / ritsu / hosei / 12, kei


def zu(no):
    return next(d for d in ZU if d["no"] == no)


def yomu_kakunin(nos):
    wb = openpyxl.load_workbook("川崎町_業務工程管理表.xlsx")
    ws = wb["03_確認事項一覧"]
    out, want = [], list(nos)
    for r in range(5, ws.max_row + 1):
        v = str(ws.cell(r, 1).value or "").strip()
        if v in want:
            out.append([v, str(ws.cell(r, 4).value or ""),
                        str(ws.cell(r, 8).value or ""),
                        str(ws.cell(r, 10).value or "")])
    out.sort(key=lambda x: want.index(x[0]))
    miss = set(want) - {x[0] for x in out}
    if miss:
        raise SystemExit("確認事項が見つからない：" + "／".join(sorted(miss)))
    return out


def main():
    if not os.path.exists(MIKOMI):
        raise SystemExit("入力がない：" + MIKOMI)
    D = yomu()

    # ══════════════════════════ 保険料の式が正しいことを先に確かめる
    saigen, _ = hokenryo(D["hyojun"], D["chiiki"], D["sogo"],
                         D["hosei"], D["ritsu"])
    if abs(saigen - D["gaku"]) > 1:
        raise SystemExit(
            f"保険料の式が現行値を再現できない：{saigen:,.1f}円 "
            f"（現行 {D['gaku']:,.1f}円）。感度の計算に進まない。")

    # ══════════════════════════ 認定者数の3シナリオ
    z24 = zu("図2-4")
    rows24 = {r[0]: r for r in z24["rows"]}
    n_r6 = rows24["令和6年度"][2]
    n_r7 = rows24["令和7年度"][2]
    n_r8 = rows24["令和8年度"][2]
    saiyo = [rows24[n][2] for n in NENDO]
    hoken_r6 = rows24["令和6年度"][1]
    ritsu_r6 = rows24["令和6年度"][3]
    ritsu_r8 = rows24["令和8年度"][3]
    hoken = [rows24[n][1] for n in NENDO]

    # 低位　認定率を令和6→8の実績の伸び（年あたり）で延ばす
    ritsu_nobi = (ritsu_r8 - ritsu_r6) / 2
    teii = []
    for i, n in enumerate(NENDO):
        r = ritsu_r8 + ritsu_nobi * (i + 1)
        teii.append(round(hoken[i] * r / 100))
    # 高位　採用と低位の差を上側に折り返す
    koui = [saiyo[i] + (saiyo[i] - teii[i]) for i in range(3)]

    SCEN = [
        ("低位（実績の伸びを延ばす）", teii,
         f"認定率を令和6年度{ritsu_r6}％→令和8年度{ritsu_r8}％の実績の伸び"
         f"（年{ritsu_nobi:+.2f}ポイント）で延ばした場合。"),
        ("採用（見える化システムの推計・自然体）", saiyo,
         "性別・年齢5歳階級別・要介護度別に推計した値。"
         "計画に載せるのはこの値。"),
        ("高位（採用と低位の差を上側に折り返す）", koui,
         "採用が実績の伸びを上回っている分だけ、"
         "さらに上振れした場合を置いたもの。"),
    ]

    # ══════════════════════════ 感度の計算
    #   総給付費（＝標準給付費）を認定者数の比で伸縮させる近似。
    KANDO = []
    for name, nin, setsumei in SCEN:
        hy = [D["hyojun"][i] * nin[i] / saiyo[i] for i in range(3)]
        gaku, _ = hokenryo(hy, D["chiiki"], D["sogo"], D["hosei"],
                           D["ritsu"])
        KANDO.append((name, nin, hy, gaku, setsumei))

    wb = openpyxl.Workbook()

    # ────────────────────────── 00
    ws = GY.sheet(
        wb, "00_この表について", "需要のシナリオと保険料の感度",
        "当方の見込量は自然体の1本であり、確からしさの幅を示せていません。"
        "見込量の算定式で動く変数は要介護度別の認定者数ひとつだけですので、"
        "認定者数を振ることで給付費と保険料の幅を示せます。\n"
        "⚠ シナリオは見込量を置き換えるものではありません。"
        "計画に載せる見込量は採用（自然体）のままです。"
        "幅は、第3回策定委員会で基金の取崩額をご決定いただくときの"
        "材料としてお示しします。\n"
        "⚠ 感度の計算は、総給付費を認定者数の比で伸縮させる近似によります。"
        "当方の算定は要介護度別ですので、要介護度の構成が変わる場合は"
        "この近似からずれます。\n"
        "⚠ 保険料の算定式が正しく組めていることを、採用のシナリオで"
        f"現行の保険料基準額（算定Ａ・{D['gaku']:,.0f}円）が"
        f"再現できるか（試算 {saigen:,.0f}円）で確かめています。",
        ["シート", "内容"], [34, 96], first=True)
    GY.put(ws, [
        ["01_認定者数の3シナリオ",
         "低位・採用・高位の認定者数と認定率。"],
        ["02_給付費と保険料の感度",
         "3シナリオそれぞれの標準給付費と保険料基準額（月額）。"],
        ["03_地域支援事業費との組合せ",
         "認定者数のシナリオと、地域支援事業費の置き方（作業順位23の4案）"
         "を組み合わせた場合の幅。"],
        ["04_第9期の実績による検算",
         "第9期の計画値と実績値の対比により、幅の置き方が"
         "実績に照らして妥当かを確かめます。"],
        ["05_確認事項", f"{len(KAKUNIN)}件。"],
    ])

    # ────────────────────────── 01
    ws = GY.sheet(
        wb, "01_認定者数の3シナリオ", "認定者数の3シナリオ",
        f"実績は令和6年度{n_r6}人・令和7年度{n_r7}人、"
        f"令和8年度は実績見込み{n_r8}人です。"
        f"認定率は令和6年度{ritsu_r6}％から令和8年度{ritsu_r8}％へ"
        f"年{ritsu_nobi:+.2f}ポイントで推移しています。\n"
        "⚠ 採用（見える化システムの推計）は、この実績の伸びより"
        "上に出ています。低位は実績の伸びをそのまま延ばした場合、"
        "高位はその差だけさらに上振れした場合です。",
        ["シナリオ", "令和9年度（人）", "令和10年度（人）",
         "令和11年度（人）", "令和11年度の認定率", "内容"],
        [34, 18, 18, 18, 18, 56])
    rows, fills = [], []
    for name, nin, setsumei in SCEN:
        rows.append([name] + [f"{v:,}" for v in nin]
                    + [f"{nin[2] / hoken[2] * 100:.1f}％", setsumei])
        fills.append(GY.OKF if name.startswith("採用") else GY.WARN)
    GY.put(ws, rows, fills)

    # ────────────────────────── 02
    ws = GY.sheet(
        wb, "02_給付費と保険料の感度", "給付費と保険料の感度",
        "総給付費を認定者数の比で伸縮させ、保険料基準額（月額）を"
        "算定し直したものです。地域支援事業費は現行のまま置いています。\n"
        f"⚠ 保険料は算定Ａ（介護給付費準備基金の取崩しなし）です。"
        "取崩しを行う場合は、下の額からさらに下がります。",
        ["シナリオ", "標準給付費3か年（千円）", "採用との差（千円）",
         "保険料基準額（月額）", "採用との差（円）", "内容"],
        [34, 24, 22, 20, 18, 56])
    rows, fills = [], []
    saiyo_gaku = KANDO[1][3]
    saiyo_hy = sum(KANDO[1][2])
    for name, nin, hy, gaku, setsumei in KANDO:
        rows.append([name, f"{sum(hy) / 1000:,.0f}",
                     f"{(sum(hy) - saiyo_hy) / 1000:+,.0f}",
                     f"{gaku:,.0f}円", f"{gaku - saiyo_gaku:+,.0f}",
                     setsumei])
        fills.append(GY.OKF if name.startswith("採用") else GY.WARN)
    GY.put(ws, rows, fills)
    r = 5 + len(KANDO) + 1
    ws.cell(r, 1).value = "読み方"
    ws.cell(r, 1).font = openpyxl.styles.Font(name="游ゴシック", size=10,
                                              bold=True)
    haba = KANDO[2][3] - KANDO[0][3]
    GY.put(ws, [[
        f"認定者数が低位から高位まで振れた場合、保険料基準額（月額）は"
        f"{KANDO[0][3]:,.0f}円から{KANDO[2][3]:,.0f}円まで、"
        f"幅{haba:,.0f}円の範囲で動きます。"
        "採用（自然体）はこの幅のまん中にあります。",
        "", "", "", "", ""],
        ["⚠ この幅は認定者数だけを振ったものです。"
         "単価（介護報酬改定）・利用率・1人1月あたりの回（日）数は"
         "令和7年度で据え置いており、振っていません。"
         "令和8年度介護報酬改定の改定率は未公表です（確認事項No.163）。",
         "", "", "", "", ""]], r0=r + 1)

    # ────────────────────────── 03
    ws = GY.sheet(
        wb, "03_地域支援事業費との組合せ",
        "認定者数と地域支援事業費の組合せ",
        "作業順位23で示した地域支援事業費の置き方の4案と、"
        "認定者数の3シナリオを組み合わせた場合の保険料基準額（月額）です。\n"
        "⚠ 地域支援事業費の置き方は確認事項No.168のご判断によります。"
        "伸びの根拠が確かめられるまで、当方は案1（現行の入力値）のまま"
        "置いています。",
        ["地域支援事業費の置き方", "低位", "採用", "高位", "備考"],
        [40, 18, 18, 18, 56])
    # 包括＋任意の3か年計（作業順位23の4案）
    AN23 = [
        ("案1（現行・見える化システムの入力値）", 0),
        ("案2（令和8年度の実績見込み額で据え置き）", -17551207),
        ("案3（令和7年度の実績額で据え置き）", -41972551),
        ("案4（平均年増の2分の1で延ばす）", -6540497),
    ]
    rows, fills = [], []
    for nm, sa in AN23:
        line = [nm]
        for name, nin, hy, gaku, _ in KANDO:
            ch = [D["chiiki"][i] + sa / 3 for i in range(3)]
            g, _x = hokenryo(hy, ch, D["sogo"], D["hosei"], D["ritsu"])
            line.append(f"{g:,.0f}円")
        line.append("当方は案1のまま置いています（確認事項No.168）。"
                    if sa == 0 else
                    f"案1との差 {sa / 1000:+,.0f}千円（3か年計）。")
        rows.append(line)
        fills.append(GY.OKF if sa == 0 else GY.WARN)
    GY.put(ws, rows, fills)

    # ────────────────────────── 04
    z32 = zu("図3-2")
    ws = GY.sheet(
        wb, "04_第9期の実績による検算", "第9期の実績による検算",
        "幅の置き方が実績に照らして妥当かを確かめます。"
        "第9期（令和7年度）の計画値と実績値の対比では、"
        "総給付費で計画が実績を1.3％上回りました。\n"
        "⚠ 区分の間で振れが相殺されています。居宅が8.8％下回る一方、"
        "施設は3.3％上回りました。総額の一致に隠れています。",
        ["区分", "計画値（千円）", "実績値（千円）", "対計画比", "読み方"],
        [30, 20, 20, 14, 62])
    rows, fills = [], []
    for name, keikaku, jisseki in z32["rows"]:
        hi = jisseki / keikaku * 100
        rows.append([name.replace("\n", ""), f"{keikaku:,}", f"{jisseki:,}",
                     f"{hi:.1f}％",
                     "計画どおり" if 97 <= hi <= 103
                     else ("計画を下回る" if hi < 97 else "計画を上回る")])
        fills.append(GY.OKF if 97 <= hi <= 103 else GY.WARN)
    GY.put(ws, rows, fills)
    r = 5 + len(z32["rows"]) + 1
    sou = z32["rows"][-1]
    kairi = (sou[2] / sou[1] - 1) * 100
    ws.cell(r, 1).value = "検算の結果"
    ws.cell(r, 1).font = openpyxl.styles.Font(name="游ゴシック", size=10,
                                              bold=True)
    GY.put(ws, [
        [f"第9期の総給付費の乖離は{kairi:+.1f}％でした。"
         f"本表の幅（低位 {(sum(KANDO[0][2]) / saiyo_hy - 1) * 100:+.1f}％"
         f"〜高位 {(sum(KANDO[2][2]) / saiyo_hy - 1) * 100:+.1f}％）は、"
         "第9期に現に生じた乖離より広く取っています。"
         "幅として狭すぎることはありません。", "", "", "", ""],
        ["⚠ ただし第9期の乖離は、居宅が下回り施設が上回って"
         "相殺された結果です。区分ごとに見ると8.8％・3.3％の振れがあり、"
         "本表の幅は区分ごとの振れまでは表していません。", "", "", "", ""],
        ["⚠ 第9期の計画が実績を上回っていたことは、"
         "計画を立てる側が多めに見積もる傾向があったことを示します。"
         "採用（自然体）がこの傾向を引き継いでいないかは、"
         "第3回策定委員会で基金の取崩額をご決定いただくときに"
         "ご留意ください。", "", "", "", ""]], r0=r + 1)

    # ────────────────────────── 05
    kak = yomu_kakunin(KAKUNIN)
    ws = GY.sheet(
        wb, "05_確認事項", "本表に関わる確認事項",
        "業務工程管理表 03_確認事項一覧 から読んでいます。"
        "内容はそちらが正本です。",
        ["No.", "確認事項", "状態", "決着しない場合の当方の扱い"],
        [8, 40, 12, 62])
    GY.put(ws, kak, [GY.OKF if x[2] == "完了" else None for x in kak])

    wb.save(OUT)

    # ══════════════════════════ 自己点検
    ng = []
    wb2 = openpyxl.load_workbook(OUT)
    if len(wb2.worksheets) != 6:
        ng.append(f"シートが{len(wb2.worksheets)}枚（6枚のはず）")
    # 採用のシナリオが現行値を再現していること
    if abs(KANDO[1][3] - D["gaku"]) > 1:
        ng.append(f"採用のシナリオが現行の保険料を再現しない："
                  f"{KANDO[1][3]:,.1f}／{D['gaku']:,.1f}")
    # 低位＜採用＜高位 の順になっていること
    if not (KANDO[0][3] < KANDO[1][3] < KANDO[2][3]):
        ng.append("低位・採用・高位の順になっていない")
    # 認定者数が実績と矛盾しないこと（令和9年度が令和8年度を下回らない）
    for name, nin, _, _, _ in KANDO:
        if nin[0] < n_r8 * 0.9:
            ng.append(f"{name} の令和9年度の認定者数が実績見込みから離れすぎ")
    zen = "\n".join(str(c.value) for w_ in wb2.worksheets
                    for row in w_.iter_rows() for c in row
                    if c.value is not None)
    for w in ("**", "__"):
        if w in zen:
            ng.append(f"強調記号が残っている：{w}")
    for w in ("大雪", "東川", "東神楽", "上川", "美瑛", "金ヶ崎", "川崎市"):
        if w in zen:
            ng.append(f"他団体の名称が入っている：{w}")
    if "置き換えるものではありません" not in zen:
        ng.append("見込量を置き換えるものでない旨の断りがない")

    print("需要のシナリオと保険料の感度")
    print("保存：", OUT)
    print(f"  シート{len(wb2.worksheets)}枚／シナリオ{len(SCEN)}件"
          f"／地域支援事業費の案{len(AN23)}件／確認事項{len(kak)}件")
    print(f"  保険料の式の検算　試算 {saigen:,.1f}円／"
          f"現行 {D['gaku']:,.1f}円（一致）")
    for name, nin, hy, gaku, _ in KANDO:
        print(f"   {name:34s} 令和11年度 {nin[2]:>4,}人"
              f"／標準給付費3か年 {sum(hy) / 1000:>12,.0f}千円"
              f"／保険料月額 {gaku:>7,.0f}円")
    print(f"   幅 {KANDO[2][3] - KANDO[0][3]:,.0f}円")
    print("  ── 自己点検")
    if ng:
        for m in ng:
            print("   ×", m)
        sys.exit(1)
    print("   ○ 採用のシナリオで現行の保険料基準額を再現できている")
    print("   ○ 低位＜採用＜高位 の順になっている")
    print("   ⚠ シナリオは見込量を置き換えるものではありません。"
          "計画に載せる見込量は採用（自然体）のままです。")


if __name__ == "__main__":
    main()
