# -*- coding: utf-8 -*-
"""介護サービス情報公表システム（本体）から町内事業所の詳細を採録する
（令和8年10月20日）

  python3 07_ソーススクリプト/fetch_kouhyou_detail_R8.10.20.py

なぜ要るか
  同日に取得したオープンデータ（24列）には、**定員の一部・入所者数・
  待機者数・職員数・指定の年月日が含まれない。** これらを
  「事業所ヒアリングでしか取れないもの」として整理していたが、
  ⚠ **公表システム本体の詳細ページには現に載っている。**

  入口（事業所ごと・サービスごとに1ページ）
    https://www.kaigokensaku.mhlw.go.jp/04/index.php
      ?action_kouhyou_detail_022_kihon=true
      &JigyosyoCd=<事業所番号10桁>-00&ServiceCd=<サービスコード>

  ⚠ **対象はオープンデータの抽出結果（リポジトリの中）から起こす。**
  事業所番号をスクリプトに書き写さない（正本を2か所に持たない）。

取り出すもの
  ①指定の年月日・指定の更新年月日　②定員（入所・入居・利用）
  ③入所者・入居者・利用者の数（要介護別×年齢別）　④待機者数
  ⑤職種別の従業者数（実人数・常勤換算）　⑥退所先　⑦記入日・公表日

出すもの
  09_元資料/その他データ/公表システム_詳細_R8.10.20/*.html　（原典）
  09_元資料/その他データ/公表システム_詳細_川崎町_R8.10.20.csv　（項目・値）

  ⚠ **一度取ったページは取り直さない。** 既にあれば通信しない
  （環境が作り直されても、原典がリポジトリにあれば同じ結果が出る）。

終了コード　0＝検査1〜8すべて適合／1＝1件でも不適合
"""
import csv
import html
import os
import re
import sys
import time
import urllib.request

MOTO = "09_元資料/その他データ/公表システム_オープンデータ_川崎町_R8.6末.csv"
HOZON = "09_元資料/その他データ/公表システム_詳細_R8.10.20"
SAKI = "09_元資料/その他データ/公表システム_詳細_川崎町_R8.10.20.csv"

URL = ("https://www.kaigokensaku.mhlw.go.jp/04/index.php"
       "?action_kouhyou_detail_022_kihon=true"
       "&JigyosyoCd={cd}-00&ServiceCd={sv}")

# 町（宮城県柴田郡川崎町）であることの確かめ
MACHI = "川崎町"
KEN = "宮城県"
CODE = "043249"
# ⚠ 川崎市（神奈川県）・福岡県田川郡川崎町と取り違えないための見張り
TORICHIGAE = ("川崎市", "福岡県", "田川郡", "神奈川県")

CELL = re.compile(r"<(th|td)\b([^>]*)>(.*?)</\1>", re.S)
ABBR = re.compile(r"abbr='([^']*)'")
TR = re.compile(r"<tr[^>]*>(.*?)</tr>", re.S)


def moji(s):
    """セルの中の文字。画像の「あり」「なし」は alt から採る。"""
    s = re.sub(r"<img[^>]*?alt=\"([^\"]*)\"[^>]*?>", r"\1", s)
    s = re.sub(r"<br\s*/?>", " ", s)
    s = re.sub(r"<[^>]+>", "", s)
    s = html.unescape(s).replace("　", " ")
    return re.sub(r"\s+", " ", s).strip()


def hiraku(url, saki):
    """ページを取る。既にあれば通信しない。"""
    if os.path.exists(saki) and os.path.getsize(saki) > 50000:
        return open(saki, encoding="utf-8", errors="replace").read(), False
    req = urllib.request.Request(
        url, headers={"User-Agent": "Mozilla/5.0 (kawasaki-town-plan)"})
    with urllib.request.urlopen(req, timeout=60) as r:
        t = r.read().decode("utf-8", "replace")
    os.makedirs(os.path.dirname(saki), exist_ok=True)
    open(saki, "w", encoding="utf-8").write(t)
    time.sleep(1.5)                      # 相手方に負荷をかけない
    return t, True


def yomu(t):
    """詳細ページを（項目, 値…）の並びにする。

    ⚠ **文字だけで探さない。** 各 <th> の abbr 属性が項目の名であり、
      行の先頭が空の <th> のときは直前の行の項目の下位である。

    ⚠ **<th> を持たない行がある。** 見出しの行（要介護１〜５・合計）と
      値の行が分かれている表（介護老人保健施設の入所者の人数）では、
      値の行に <th> がない。**直前の項目の値として扱う**（捨てると
      入所者の数が落ちる）。
    """
    oya = ""
    mae = ""
    gyo = []
    for r in TR.finditer(t):
        body = r.group(1)
        ths, tds, saki_kara = [], [], None
        for m in CELL.finditer(body):
            if m.group(1) == "th":
                a = ABBR.search(m.group(2))
                a = a.group(1) if a else ""
                if saki_kara is None:
                    saki_kara = (a == "" and moji(m.group(3)) == "")
                if a:
                    ths.append(moji(a))
            else:
                tds.append(moji(m.group(3)))
        if not ths and tds:
            komoku = mae                 # ⚠ 直前の項目の値
        elif saki_kara and oya:
            komoku = oya + " ▸ " + " / ".join(ths)
        else:
            komoku = " / ".join(ths)
            if ths:
                oya = ths[0]
        if komoku:
            mae = komoku
        if komoku or tds:
            gyo.append((komoku, tds))
    return gyo


def hitotsu(gyo, *na):
    """項目の名が na のいずれかで始まる行の、最初の値を返す。"""
    for k, v in gyo:
        for n in na:
            if k == n or k.startswith(n):
                if v and v[0]:
                    return v[0]
    return ""


def main():
    futeki = []
    mi = []      # 埋まらなかったもの（「適合」に丸めない）

    def judge(ok, midashi, setsu=""):
        print(("   ○ " if ok else "   × ") + midashi
              + (("　" + setsu) if setsu else ""))
        if not ok:
            futeki.append(midashi)

    print("介護サービス情報公表システム（本体）からの採録"
          "（令和8年10月20日）")
    print()

    # ── 対象はリポジトリの中のオープンデータ抽出結果から起こす
    if not os.path.exists(MOTO):
        print(f"   × 入力がありません：{MOTO}")
        print("      先に fetch_kouhyou_opendata_R8.10.20.py を"
              "実行してください。")
        sys.exit(1)
    taisho = []
    with open(MOTO, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            taisho.append((row["事業所番号"], row["サービスコード"],
                           row["サービスの種類"], row["事業所名"],
                           row["定員"], row["市区町村名"],
                           row["都道府県コード又は市町村コード"]))
    print(f"  ── 対象　{MOTO}　{len(taisho)}件")
    print()

    # ══════ 採録
    dedata = []          # (サービスコード, 種類, 事業所名, 番号, 項目, 値)
    page = {}
    tsushin = 0
    print("  ── ① ページを取る")
    for cd, sv, shu, na, teiin, shi, kc in taisho:
        saki = os.path.join(HOZON, f"詳細_{cd}_{sv}.html")
        try:
            t, atarashii = hiraku(URL.format(cd=cd, sv=sv), saki)
        except Exception as e:
            judge(False, f"{na}（{shu}）が取れない", str(e)[:70])
            mi.append(f"{na}（{shu}）：ページが取れなかった（{e}）")
            continue
        tsushin += 1 if atarashii else 0
        gyo = yomu(t)
        page[(cd, sv)] = (t, gyo, shu, na, teiin)
        for k, v in gyo:
            # ⚠ **値のない行（表の見出し）も残す。** 捨てると
            #   「退所者の人数（前年度の状況）」の期間が分からなくなり、
            #   続く「自宅等」「医療機関」がどの表の行か決められない。
            if k:
                dedata.append((sv, shu, na, cd, k, " | ".join(v)))
        print(f"     {cd} {sv} {shu}　{'取得' if atarashii else '既にある'}"
              f"　行 {len(gyo)}")
    print()

    # ══════ 検査1　14件すべて取れたか
    print("  ── 検査1　対象のすべてが取れたか")
    judge(len(page) == len(taisho),
          f"{len(taisho)}件すべて取得（実際 {len(page)}件）")

    # ══════ 検査2　事業所番号がページと一致するか
    print("  ── 検査2　ページの事業所番号が一致するか")
    chigai = []
    for (cd, sv), (t, gyo, shu, na, _) in page.items():
        ban = hitotsu(gyo, "介護保険事業所番号", "事業所番号")
        ban = re.sub(r"\D", "", ban)
        if ban != cd:
            chigai.append(f"{na}（{shu}）ページ={ban} 一覧={cd}")
    judge(not chigai, "ページの事業所番号が一覧と一致",
          "／".join(chigai[:3]))

    # ══════ 検査3　川崎町（宮城県柴田郡）であるか
    print("  ── 検査3　宮城県柴田郡川崎町のものであるか")
    chigai = []
    for (cd, sv), (t, gyo, shu, na, _) in page.items():
        if MACHI not in t:
            chigai.append(f"{na}（{shu}）に「{MACHI}」がない")
    judge(not chigai, f"すべてのページに「{MACHI}」がある",
          "／".join(chigai[:3]))

    # ══════ 検査4　⚠ 川崎市・福岡県川崎町が混じらないか
    print("  ── 検査4　⚠ 川崎市（神奈川県）・福岡県田川郡川崎町が"
          "混じらないか")
    chigai = []
    for (cd, sv), (t, gyo, shu, na, _) in page.items():
        # ⚠ 所在地の行だけを見る（別の都道府県の病院名等で誤らないため）
        jusho = " ".join(v for k, vs in gyo for v in vs
                         if "所在地" in k or "住所" in k)
        for w in TORICHIGAE:
            if w in jusho:
                chigai.append(f"{na}（{shu}）の所在地に「{w}」")
    judge(not chigai, "所在地に他団体（川崎市・福岡県川崎町）が混じらない",
          "／".join(chigai[:3]))

    # ══════ 検査5　定員がオープンデータと食い違わないか
    print("  ── 検査5　定員がオープンデータと食い違わないか")
    #   ⚠ オープンデータの定員は0（未報告）のものがある。
    #     **0を「定員0人」と読まない。** 本体に数があればそれを採る。
    teiin_de = {}
    chigai, oginau = [], []
    for (cd, sv), (t, gyo, shu, na, od) in page.items():
        v = hitotsu(gyo, "入所定員", "入居定員", "利用定員")
        kazu = re.findall(r"(\d+)人", v)
        hon = int(kazu[-1]) if kazu else None
        teiin_de[(cd, sv)] = (hon, v)
        try:
            odv = int(od)
        except ValueError:
            odv = None
        if hon is None:
            continue
        if odv == 0:
            oginau.append(f"{na}（{shu}）オープンデータ0 → 本体 {hon}人")
        elif odv is not None and odv != hon:
            chigai.append(f"{na}（{shu}）本体={hon} オープンデータ={odv}")
    judge(not chigai, "両者に数があるものは一致",
          "／".join(chigai[:3]))
    if oginau:
        print(f"     △ オープンデータが0（未報告）で本体に数があったもの"
              f"{len(oginau)}件")
        for s in oginau:
            print(f"        {s}")

    # ══════ 検査6　記入日・公表日が取れたか
    print("  ── 検査6　記入日・公表日（いつ時点の値か）")
    hiduke = {}
    nashi = []
    for (cd, sv), (t, gyo, shu, na, _) in page.items():
        k = re.search(r"記入日[：:]\s*(\d{4})年(\d{2})月(\d{2})日", t)
        h = re.search(r"(\d{4})年(\d{2})月(\d{2})日[\d:]*(?:&nbsp;|\s)*公表",
                      t)
        hiduke[(cd, sv)] = ("/".join(k.groups()) if k else "",
                            "/".join(h.groups()) if h else "")
        if not k:
            nashi.append(f"{na}（{shu}）")
    judge(not nashi, "すべてのページから記入日が取れた",
          "／".join(nashi[:3]))

    # ══════ 検査7　⚠ 埋まらなかったものを数で残す
    print("  ── 検査7　⚠ 埋まらなかったもの（「適合」に丸めない）")
    for (cd, sv), (t, gyo, shu, na, _) in page.items():
        hon, v = teiin_de[(cd, sv)]
        if hon is None and sv in ("510", "520", "540", "320", "210", "150",
                                  "160"):
            mi.append(f"{na}（{shu}）：定員が載っていない")
    print(f"     {'○' if not mi else '△'} 埋まらなかったもの {len(mi)}件")
    for s in mi:
        print(f"        {s}")

    # ══════ 書き出し
    os.makedirs(os.path.dirname(SAKI), exist_ok=True)
    with open(SAKI, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["サービスコード", "サービスの種類", "事業所名",
                    "事業所番号", "記入日", "公表日", "項目", "値"])
        for sv, shu, na, cd, k, v in dedata:
            ki, ko = hiduke.get((cd, sv), ("", ""))
            w.writerow([sv, shu, na, cd, ki, ko, k, v])
    print()
    print("  ── ② 書き出し")
    print(f"     {SAKI}　{len(dedata):,}行")
    print(f"     {HOZON}/　{len(page)}ページ（原典）")

    # ══════ 検査8　書き出したものが読み返せるか
    print("  ── 検査8　書き出したものが読み返せるか")
    with open(SAKI, encoding="utf-8-sig", newline="") as f:
        yomi = list(csv.DictReader(f))
    judge(len(yomi) == len(dedata),
          f"{len(dedata):,}行が読み返せる（実際 {len(yomi):,}行）")
    judge(all(r["事業所番号"] and len(r["事業所番号"]) == 10 for r in yomi),
          "事業所番号がすべて10桁")

    print()
    print(f"  通信したページ {tsushin}件"
          f"（既にあったもの {len(page) - tsushin}件）")
    print()
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        for m in futeki:
            print(f"      {m}")
        sys.exit(1)
    print(f"   ○ 検査1〜8　不適合0件"
          f"（⚠ 埋まらなかったもの {len(mi)}件は別に数えています）")


if __name__ == "__main__":
    main()
