# -*- coding: utf-8 -*-
"""概要版（A3両面）が収まるかを面積で積む（WBS Ⅳ-92）

   頁数ではなく面積で見る。概要版は1面のなかで段を組むため、
   行送りを追う推定（estimate_layout）より面積の積み上げが実際に近い。

     python3 scripts/estimate_gaiyou.py

   【積み方】
   本文・表は字数から面積を出す。1文字あたりの面積は
   （字の幅 × 行送り）であり、見出しは字の大きさの倍率をかける。
   図は PNG の縦横比から、与えた幅での高さを出す。
   要素の間の空きは、面の使える面積に占める割合（KUKAN）で見込む。

   【この推定でできないこと】
   段の割り方・図の置き方・折りの位置は紙面を見ないと決まらない。
   本環境では docx を PDF に変換できないため、**紙面の確認はつねに未実施**である。
   面積が収まることは、収まる見込みが立つことを意味するにとどまる。
"""
import os
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gaiyou_data as G                              # noqa: E402
import estimate_pages as EP                          # noqa: E402

KUKAN = 0.22        # 要素の間の空き・段間・図の表題が占める割合
ZU_HABA = 125.0     # 図の幅（mm）。A3縦の使える幅267mmの半分弱

#    表と枠は、ます目の余白・罫線・埋まらないます目があるため、
#    字数から出した面積のとおりには収まらない。1.6倍で見込む。
#    （字数だけで積むと、6行の表が6行分より狭く出てしまう）
HYO_BAI = 1.6


def moji_mm2(kazu, kubun):
    """字数から面積（mm2）。見出しは字が大きい"""
    bai = G.MIDASHI_BAI.get(kubun, 1.0)
    w = G.MOJI * bai
    h = G.MOJI * bai * G.GYOKAN
    return kazu * w * h


def zu_mm2(fn, haba=ZU_HABA):
    """図の面積（mm2）。表題の1行分を足す"""
    w_in, h_in = EP.fig_in(os.path.join(EP.FIGDIR, fn))
    h_mm = haba * h_in / w_in
    return haba * h_mm, haba, h_mm


def _haba_of(bikou):
    """備考に「幅Nmm」と書かれていればその幅を使う。なければ既定の幅"""
    import re
    m = re.search(r"幅(\d+)mm", bikou or "")
    return float(m.group(1)) if m else ZU_HABA


def run():
    out = []
    for men, nazuke, w, h, yohaku in G.MEN:
        tsukaeru = (w - yohaku * 2) * (h - yohaku * 2)
        tsumi = 0.0
        gyo = []
        for y in G.YOSO:
            if y[0] != men:
                continue
            _men, no, kubun, midashi, moto, kazu, fn, _bikou = y
            if kubun == "図":
                a, zw, zh = zu_mm2(fn, _haba_of(_bikou))
                a += moji_mm2(30, "本文")        # 図の表題と出典
                gyo.append((no, kubun, midashi, moto,
                            "%.0f×%.0fmm" % (zw, zh), a))
            else:
                a = moji_mm2(kazu, kubun)
                if kubun in ("表", "枠"):
                    a *= HYO_BAI
                gyo.append((no, kubun, midashi, moto, "%d字" % kazu, a))
            tsumi += a
        aki = tsukaeru * KUKAN
        out.append((men, nazuke, tsukaeru, tsumi, aki, gyo))
    return out


def main():
    print("■ 概要版（A3両面）の収まりの見込み")
    print("   A3縦 297×420mm・余白15mm・本文9pt・行送り1.5倍")
    print("   要素の間の空きは使える面積の%d%%として見込む\n" % (KUKAN * 100))
    ng = []
    for men, nazuke, tsukaeru, tsumi, aki, gyo in run():
        nokori = tsukaeru - tsumi - aki
        wari = 100.0 * (tsumi + aki) / tsukaeru
        print("── %s　%s ──────────────────────────" % (men, nazuke))
        print("   %-3s %-6s %-44s %-14s %10s %8s"
              % ("#", "区分", "要素", "出どころ", "大きさ", "面積mm2"))
        for no, kubun, midashi, moto, ookisa, a in gyo:
            print("   %-3d %-6s %-44s %-14s %10s %8.0f"
                  % (no, kubun, midashi[:44], moto, ookisa, a))
        print("   ────────────────────────────────────────")
        print("   使える面積 %.0f ／ 要素 %.0f ＋ 空き %.0f ＝ %.0f（%.0f%%）"
              % (tsukaeru, tsumi, aki, tsumi + aki, wari))
        if nokori < 0:
            print("   ✗ %.0fmm2 あふれる（字数を%.0f字ほど減らすか、図を小さくする）"
                  % (-nokori, -nokori / (G.MOJI ** 2 * G.GYOKAN)))
            ng.append(men)
        else:
            print("   ○ %.0fmm2 の余り（本文で約%.0f字分）"
                  % (nokori, nokori / (G.MOJI ** 2 * G.GYOKAN)))
        print()

    print("■ 骨子の段階で決まっていないもの（%d件）" % len(G.MIKETTEI))
    for koumoku, joken, basho, bikou in G.MIKETTEI:
        print("   ・%-22s %s" % (koumoku, joken))
        print("     %-22s %s" % ("（%s）" % basho, bikou))
    print()
    print("※ 段の割り方・図の置き方・折りの位置は紙面を見ないと決まらない。")
    print("   本環境では docx を PDF に変換できないため、紙面の確認は未実施である。")
    print("   面積が収まることは、収まる見込みが立つことを意味するにとどまる。")
    return 1 if ng else 0


if __name__ == "__main__":
    sys.exit(main())
