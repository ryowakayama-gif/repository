# -*- coding: utf-8 -*-
"""事業所照会票のご回答の受入検査（令和8年10月16日）

08_作業順位 の順位27。

事業所からご回答（電子記入様式の xlsx）が届いたときに、
**そのまま素案に載せてよいか**を機械で確かめる。
いただく前から用意しておくことで、届いた日に次へ進める。

  python3 07_ソーススクリプト/check_shokai_kaito_R8.10.16.py [回答のxlsx ...]

  引数を省くと、お配りする電子記入様式そのものを見る
  （まだ1件もご記入がない状態を確かめる）。

検査
  1　様式が揃っているか（9様式。シートを増減していないか）
  2　設問が揃っているか（照会票の正本と突き合わせる）
  3　必ず要る欄（事業所の名称・ご記入者・ご記入の年月日）が埋まっているか
  4　様式1・様式2のご回答があるか（この2つで素案 2-7 の【町確認】が埋まる）
  5　数として読めるか。合計が内訳の和と合っているか
  6　人数の大小が逆になっていないか（実人数 ≦ 定員 ほか）
  7　「0」と「記録なし」と空欄が区別されているか
  8　個人が特定される事項が書かれていないか（氏名・生年月日の形）
  9　他団体のものが混じっていないか（福岡県田川郡川崎町・神奈川県川崎市）
 10　複数の事業所のご回答で、同じ職員が重複して数えられていないか

⚠ 本検査は**様式と整合**を見るものであり、値が正しいかどうかは確かめられない。
  値の確かさは、町の指定台帳との突合・事業所への再照会による。

⚠ **「未記入」は不適合としない。** 届いた日の段階ではほとんどが未記入である。
  未記入の件数と、様式1・様式2が返っているかどうかを示す。
"""
import importlib.util
import os
import re
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))


def load(name, fname):
    spec = importlib.util.spec_from_file_location(name,
                                                  os.path.join(HERE, fname))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


# 設問の正本は照会票を作るスクリプトの側に置く（2か所に持たない）
B = load("shokai27", "build_jigyosho_shokai_R8.10.16.py")

KINYU = "03_委員会・説明資料/川崎町_事業所照会票_電子記入様式_R8.10.16.xlsx"

# 他団体のもの（取り違えやすい）
BETSU = ["田川郡", "福岡県", "神奈川県", "川崎市高津", "川崎市中原",
         "川崎市川崎区", "川崎市幸区", "川崎市多摩", "川崎市宮前",
         "川崎市麻生"]
# 個人が特定される形
#   ⚠ **氏名は値の形では探せない。** 機械では見ず、未実施として件数で示す
#     （点検スキル3「氏名・住所・自由記述・画像の中の文字は機械では拾えない」）。
PII = [(r"(明治|大正|昭和|平成|令和)\d{1,2}年\d{1,2}月\d{1,2}日生",
        "生年月日の形"),
       (r"^\D*\d{10}\D*$", "10桁の番号（被保険者番号の疑い）")]
# ⚠ 事業所番号は10桁である（正しい記載）。この欄は10桁の走査から外す。
PII_NOZOKU = ["事業所番号"]

# 合計と内訳の対応（設問のid, 合計の欄, 内訳の欄の並び）
GOKEI = [
    ("zen_kyosho", "合計",
     ["ご自宅から（合計）", "病院・診療所から", "他の介護保険施設から",
      "高齢者向け住まいから"]),
    ("taisho_saki", "合計",
     ["ご自宅へ", "病院・診療所へ", "他の介護保険施設へ",
      "高齢者向け住まいへ", "お亡くなりになった方"]),
    ("jugyoin", "合計", list(B.SHOKU)),
]
# 大小の向き（設問のid, 小さいはずの欄, 大きいはずの欄）
DAISHO = [
    ("teiin", "うち川崎町の被保険者", "現に受け入れている利用者の実人数"),
    ("teiin", "うち他の市町村の被保険者", "現に受け入れている利用者の実人数"),
    ("machi", "うち川崎町の被保険者", "お申込みの人数（合計）"),
    ("machi", "うち現在ご自宅にお住まいの方", "お申込みの人数（合計）"),
    ("touroku", "基準日の属する月に1回以上提供した人数（実働者数）",
     "登録している訪問介護員等の人数"),
    ("mitori", "うち再入所に至った件数", "入所者が医療機関に入院された件数"),
    ("cm", "うち主任介護支援専門員", "介護支援専門員の人数"),
]


def suji(v):
    """最初に現れる数を返す。読めなければ None。

    ⚠ 職種別の欄は「常勤換算 10.0人／実人数 12人」のように**2つの数が
      1つのセルに入る。** 常勤換算を先に書く様式としているため、
      最初の数を取れば常勤換算になる。
    ⚠ 「記録なし」「お答えを控えます」は数ではない（None を返す）。
      0件と区別するため、ここで0に丸めない。
    """
    if isinstance(v, (int, float)):
        return float(v)
    if not isinstance(v, str):
        return None
    if re.search(r"(記録なし|お答えを控え|分からない|未定)", v):
        return None
    m = re.search(r"\d+(?:\.\d+)?", v.replace(",", "").replace("，", ""))
    return float(m.group(0)) if m else None


def yomu(path):
    """回答の xlsx を {様式: {項目: 記入}} に読む。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    out = {}
    for ws in wb.worksheets:
        code = ws.title.split("_")[0]
        d = {}
        for r in range(5, ws.max_row + 1):
            k = ws.cell(r, 2).value
            v = ws.cell(r, 4).value       # ⚠ 4列目がご記入欄
            if k is None:
                continue
            k = str(k).strip()
            # 同じ項目名が様式の中で複数回出ることがある（「合計」など）。
            # 先に出たものを残す（設問の順に読むため）。
            if k not in d:
                d[k] = v
        out[code] = d
    return out, [ws.title for ws in wb.worksheets]


def mireru(path, ng, note, kazu):
    kaito, sheets = yomu(path)
    name = path.split("/")[-1]
    bg = B.bangou()

    # ── 検査1　様式が揃っているか
    kitai = [f"{y[0]}_{y[1]}"[:31] for y in B.YOSHIKI]
    if sheets != kitai:
        chigai = [f"{a}→{b}" for a, b in zip(kitai, sheets) if a != b]
        if len(sheets) != len(kitai):
            ng.append(f"検査1　{name}　シートが{len(sheets)}枚"
                      f"（{len(kitai)}枚のはず）")
        elif chigai:
            ng.append(f"検査1　{name}　様式の名が違う："
                      + "／".join(chigai[:3]))

    # ── 検査2　設問が揃っているか
    for yo, sid, _toi, kata, *_ in B.SETSUMON:
        if kata not in ("kv", "sel"):
            continue
        atama = B.RAN[sid][0]
        atama = atama[0] if isinstance(atama, tuple) else "当てはまるものに○"
        if yo in kaito and atama not in kaito[yo]:
            ng.append(f"検査2　{name}　{B.toi_label(sid)} の欄がない："
                      f"{atama[:16]}")

    # ── 検査3・4　必ず要る欄とご回答の有無
    kinyu_ari = []
    for code in kitai:
        yo = code.split("_")[0]
        d = kaito.get(yo, {})
        # 記入されている欄の数（冒頭の3欄を除く）
        atai = [k for k, v in d.items()
                if v not in (None, "") and k not in
                ("事業所の名称", "ご記入者のお名前", "ご記入の年月日")]
        if atai:
            kinyu_ari.append(yo)
            for k in ("事業所の名称", "ご記入者のお名前", "ご記入の年月日"):
                if not d.get(k):
                    ng.append(f"検査3　{name}　{yo}　"
                              f"ご回答があるのに「{k}」が未記入")
    kazu["記入のある様式"] = kinyu_ari
    if "様式9" in kinyu_ari:
        if len(kinyu_ari) > 1:
            note.append(f"{name}　様式9（町外の事業所用）のほかにも"
                        "ご記入があります。町内の事業所であれば"
                        "様式1から様式8でご記入ください")
    else:
        for y in ("様式1", "様式2"):
            if y not in kinyu_ari:
                note.append(f"{name}　{y}のご回答がありません"
                            "（この2つで素案 2-7 の把握すべきデータが"
                            "埋まりますので、督促の対象です）")

    # ── 検査5　合計と内訳
    #   ⚠ 内訳が一部しか埋まっていない段階でも効く形にする。
    #     ①埋まっている内訳の和が合計を**超えている**ことは、
    #       内訳が未完成であっても誤りである。
    #     ②内訳がすべて埋まっているときは、和と合計が一致すること。
    for sid, go, uchi in GOKEI:
        yo = bg[sid][0]
        d = kaito.get(yo, {})
        g = suji(d.get(go))
        if g is None:
            continue
        vs = [suji(d.get(u)) for u in uchi]
        sonota = [suji(v) for k, v in d.items() if k.startswith("その他")]
        aru = [v for v in vs if v is not None] \
            + [v for v in sonota if v is not None]
        if not aru:
            continue
        if sum(aru) > g + 0.05:
            ng.append(f"検査5　{name}　{B.toi_label(sid)}　"
                      f"内訳の和{sum(aru):g}が合計{g:g}を超えています")
        elif all(v is not None for v in vs) and abs(sum(aru) - g) > 0.05:
            ng.append(f"検査5　{name}　{B.toi_label(sid)}　"
                      f"内訳がすべて埋まっていますが、"
                      f"和{sum(aru):g}が合計{g:g}と合っていません")

    # ── 検査6　大小の向き
    for sid, chi, dai in DAISHO:
        yo = bg[sid][0]
        d = kaito.get(yo, {})
        a, b = suji(d.get(chi)), suji(d.get(dai))
        if a is not None and b is not None and a > b + 0.05:
            ng.append(f"検査6　{name}　{B.toi_label(sid)}　"
                      f"「{chi[:18]}」{a:g} が「{dai[:18]}」{b:g} を"
                      "超えています")

    # ── 検査7　0と記録なしと空欄の区別
    #   ⚠ ご記入のある様式に限って見る。まだ届いていない様式の空欄は
    #     「未記入」であって「0と記録なしの区別がない」のではない。
    for yo, sid, _t, kata, *_ in B.SETSUMON:
        if kata != "kv" or yo not in kinyu_ari:
            continue
        for k, v in B.RAN[sid]:
            if "記録なし" not in str(v):
                continue
            x = kaito.get(yo, {}).get(k)
            if x in (None, ""):
                note.append(f"{name}　{B.toi_label(sid)}　「{k[:20]}」が"
                            "空欄です。0件か記録なしかをご記入いただけますと"
                            "集計できます")

    # ── 検査8・9　個人情報と他団体
    zen = "\n".join(f"{k}　{v}" for d in kaito.values()
                    for k, v in d.items() if v not in (None, ""))
    for d in kaito.values():
        for k, v in d.items():
            if v in (None, ""):
                continue
            if any(x in k for x in PII_NOZOKU):
                continue
            for pat, nm in PII:
                m = re.search(pat, str(v))
                if m:
                    ng.append(f"検査8　{name}　{nm}が書かれています："
                              f"{k[:16]}／{m.group(0)[:20]}")
    for b in BETSU:
        if b in zen:
            ng.append(f"検査9　{name}　他団体のものが混じっています：{b}")
    return kaito


def main():
    paths = sys.argv[1:]
    hairu = False
    if not paths:
        paths = [KINYU]
        hairu = True
        if not os.path.exists(KINYU):
            print("電子記入様式がありません：", KINYU)
            print("  先に build_jigyosho_shokai_R8.10.16.py を"
                  "実行してください。")
            sys.exit(1)
    ng, note, kazu = [], [], {}
    kaitos = []
    for p in paths:
        if not os.path.exists(p):
            ng.append(f"ファイルがありません：{p}")
            continue
        kaitos.append((p, mireru(p, ng, note, kazu)))

    # ── 検査10　複数の事業所のご回答で職員が重複していないか
    if len(kaitos) > 1:
        namae = {}
        for p, k in kaitos:
            nm = (k.get("様式1", {}).get("事業所の名称")
                  or k.get("様式9", {}).get("事業所の名称"))
            if nm:
                namae.setdefault(str(nm).strip(), []).append(p)
        for nm, ps in namae.items():
            if len(ps) > 1:
                ng.append(f"検査10　同じ事業所の名称「{nm}」のご回答が"
                          f"{len(ps)}件あります。様式1（事業所票）は"
                          "1枚だけですので、職員の数が重複して"
                          "数えられるおそれがあります")

    print("事業所照会票のご回答の受入検査")
    print(f"  見たファイル {len(paths)}件")
    for p in paths:
        print("   ", p)
    if hairu:
        print("  ⚠ お配りする電子記入様式そのものを見ています"
              "（まだご回答が届いていない状態）。")
    print(f"  様式{len(B.YOSHIKI)}種／設問{len(B.SETSUMON)}問"
          f"／合計と内訳の突合{len(GOKEI)}組／大小の向き{len(DAISHO)}組")
    print("  ── 走査できなかった範囲（未実施）")
    print("   氏名・住所・自由記述は値の形では探せないため走査していません。"
          "自由記入の欄は目で見てください。")
    if kazu.get("記入のある様式"):
        print("  ご記入のある様式：", "／".join(kazu["記入のある様式"]))
    for m in note[:20]:
        print("   △", m)
    if len(note) > 20:
        print(f"   △ ほか{len(note) - 20}件")
    if ng:
        for m in ng[:30]:
            print("   ×", m)
        print(f"   不適合 {len(ng)}件")
        print("  ⚠ 不適合がある間は、素案に載せないでください。")
        sys.exit(1)
    if not kazu.get("記入のある様式"):
        print("   △ まだ1問もご記入がありません（様式をお渡しした段階です）")
        print("  ⚠ 事業所からご回答が届いたら、ファイルを引数で渡して"
              "もう一度実行してください。")
        return
    print("   ○ 様式・設問・必ず要る欄・合計と内訳・大小の向き・"
          "個人情報・他団体の混入はいずれも適合")
    print("  ⚠ 本検査は様式と整合を見るものです。"
          "値が正しいかどうかは、町の指定台帳との突合と再照会によります。")


if __name__ == "__main__":
    main()
