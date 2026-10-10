# -*- coding: utf-8 -*-
"""重複回答の排除（WBS Ⅰ-36）

   仕様書4Ⅰ(2)は「紙・ウェブ双方で回答された場合に重複集計されないよう
   管理番号等で管理」することを求めている。紙とウェブの併用は第9期にはなく、
   第10期からの新規要件である。

     python3 scripts/shukei_dedupe.py                        ダミーデータで通す
     python3 scripts/shukei_dedupe.py --needs <csv> --zaitaku <csv> --out <csv>
     python3 scripts/shukei_dedupe.py --自己試験              排除そのものを確かめる

   【要点は「一意に定まること」である】
   同じ管理番号で2件あったとき、どちらを採るかが人や実行のたびに変われば、
   回収数も構成比も変わる。本スクリプトは採否の順序を5段で定め、
   最後の段を回答の中身で決めることで、
   **読み込んだ順序を入れ替えても結果が変わらない**ようにしている（--自己試験 で確認）。

   【ウェブの管理番号は回答者が手で入れる】
   依頼状に管理番号の欄を設けている（doc04 B-1）。手で入れるため、
   全角・小文字・ハイフン・前後の空白・先頭の0の有無が揺れる。
   突合の前に正規化する。空欄や名簿にない番号も起こりうる。

   【名簿に当たらない回答は捨てない】
   管理番号が空、または名簿にない回答も、回答そのものは有効である。
   捨てれば村の回答者の声を落とすことになる。単純集計には入れ、
   属性のクロス（軸B 地区・軸C 年齢・軸D 性別）には入れられない旨を記録する。
   分母は設問ごとの有効回答であるため（基本方針9）、これで成立する。
   ただし**管理番号が空の回答どうしの重複は見つけられない**。
   件数を報告書に明記する。
"""
import argparse
import collections
import csv
import io
import os
import re
import sys
import unicodedata

sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths as P                                   # noqa: E402
import shukei_check as SC                           # noqa: E402

# ── 重複の型 ──────────────────────────────────
#    (記号, 名, 何が起きているか, どう扱うか)
KATA = [
 ("D1", "紙が2件",
  "同じ管理番号の紙の回答が2件以上ある",
  "同一人が2枚返した（予備票）か、入力の重複。採否の順序で1件に絞る"),
 ("D2", "紙とウェブ",
  "同じ管理番号で紙とウェブの回答がある",
  "仕様書4Ⅰ(2)が想定している重複。採否の順序で1件に絞る"),
 ("D3", "ウェブが2件",
  "同じ管理番号のウェブの回答が2件以上ある",
  "フォームの二度押しまたは再送信。採否の順序で1件に絞る"),
 ("D4", "管理番号が空",
  "管理番号が入っていない（ウェブで入れ忘れた等）",
  "回答は残す。名簿と突合できないため属性のクロスには入れない。"
  "空どうしの重複は見つけられないため件数を記録する"),
 ("D5", "管理番号の形が違う",
  "正規化しても票の形式（ニーズ＝数字のみ／在宅＝英字＋数字）に合わない",
  "原票に当たって直す。直らなければ D4 と同じ扱いとする"),
 ("D6", "名簿にない管理番号",
  "正規化後の管理番号が対象者名簿にない",
  "原票に当たって直す。直らなければ回答は残し、属性のクロスには入れない"),
 ("D7", "名簿の番号が正規化で重なる",
  "名簿の2つの管理番号が、正規化すると同じ値になる",
  "先頭の0を落とす規則が使えない。正規化の規則を見直すまで突合を行わない"),
]

# ── 管理番号の正規化 ──────────────────────────────
#    L09（管理番号の不備）と同じ規則を使う。規則が2か所に分かれると、
#    点検は「重複なし」と言い、排除は1件落とすという食い違いが起きる。
KEISHIKI = SC.KEISHIKI
seiki = SC.seiki
kata_ga_au = SC.kata_ga_au

# ── 採否の順序（5段）───────────────────────────────
#    上の段で決まればそこで決まる。最後の段まで中身で決めるため、
#    読み込んだ順序に依存しない。
SAIHI = [
 ("①", "有効票を採る（白票は採らない）",
  "白票（必須設問がすべて無回答。shukei_check L01）は集計に用いない"),
 ("②", "受け付けが後のものを採る",
  "回答者が出し直した場合、後のものが訂正の意図によるとみられる"),
 ("③", "同じ日ならウェブを採る",
  "ウェブは入力を経ないため、入力の誤りが入らない"),
 ("④", "回答した設問が多いものを採る",
  "情報量の多いほうを残す"),
 ("⑤", "それでも並ばなければ回答の中身の並びで決める",
  "読み込んだ順序で決めると、実行のたびに結果が変わりうる"),
]

HOUHOU_COL = "回答方法"
UKETSUKE_COL = "受付日"
WEB = "Web"


def _uketsuke(r):
    """受付日を並べ替えに使える値にする。入っていなければ最も古い扱い"""
    s = unicodedata.normalize("NFKC", str(r.get(UKETSUKE_COL, "") or "")).strip()
    s = s.replace("/", "-").replace(".", "-")
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if not m:
        return ""
    return "%04d-%02d-%02d" % tuple(int(x) for x in m.groups())


def _web(r):
    return unicodedata.normalize("NFKC",
                                 str(r.get(HOUHOU_COL, "") or "")).strip().lower() \
        in ("web", "ウェブ", "ウエブ")


def _kaitou_su(r, ent):
    return sum(1 for e in ent if SC._answered(r, e))


def _naiyou(r):
    """回答の中身を並べ替えに使える形にする（読み込んだ順に依存させないため）"""
    return tuple(sorted((k, str(v)) for k, v in r.items() if not k.startswith("_")))


def _rank(r, ent, hakuhyou):
    """小さいほうを採る"""
    return (0 if r["_id"] not in hakuhyou else 1,        # ① 有効票
            _uketsuke(r) == "", _minus(_uketsuke(r)),    # ② 受け付けが後
            0 if _web(r) else 1,                         # ③ ウェブ
            -_kaitou_su(r, ent),                         # ④ 回答数が多い
            _naiyou(r))                                  # ⑤ 中身の並び


def _minus(s):
    """文字列を降順に並べるための鍵（後の日付が先に来るように）"""
    return tuple(-ord(c) for c in s)


def dedupe(rows, hyo, kb, meibo=None):
    """1つの票の回答データから重複を排除する。

       戻り値 (採用した行, 記録, 件数)
         記録 ＝ (管理番号, 記号, 内容) の並び
         件数 ＝ 読み込み・採用・除外・型別などの数え
    """
    ent = [e for e in kb
           if (e["票"] == hyo or (hyo == "在宅" and e["票"].startswith("在宅")))
           and e["列"]]
    hissu = [e for e in ent if e["区分"] == "必須"]
    kiroku = []
    kazu = collections.Counter()
    kazu["読み込み"] = len(rows)

    # 白票の判定は shukei_check と同じ規則で行う（二重に定義しない）
    hakuhyou = set()
    for i, r in enumerate(rows):
        r["_id"] = i
        if hissu and not any(SC._answered(r, e) for e in hissu):
            hakuhyou.add(i)
    kazu["白票"] = len(hakuhyou)

    # 名簿も同じ規則で正規化する（規則が違うと突合が崩れる）
    meibo_s = None
    if meibo is not None:
        meibo_s = {seiki(b)[0] for b in meibo}
        # 先頭の0を落とす規則は、名簿の番号が1対1で写るときにしか使えない。
        # 名簿で「0001」と「1」が別人として並んでいれば、この規則は使えない。
        if len(meibo_s) != len(set(meibo)):
            kazu["D7"] += 1
            kiroku.append(("―", "D7",
                           "名簿%d件が正規化で%d件に重なる。突合を行わない"
                           % (len(set(meibo)), len(meibo_s))))
            meibo_s = None

    # ── 正規化と、突合できない件の仕分け ──────────────
    group = collections.defaultdict(list)
    tsukigo_fuka = []                     # 名簿と突合できない行
    for r in rows:
        ban, kara = seiki(r.get("管理番号"))
        r["_番号"] = ban
        if kara:
            kazu["D4"] += 1
            kiroku.append(("（空）", "D4", "管理番号が入っていない"))
            tsukigo_fuka.append(r)
            continue
        if not kata_ga_au(ban, hyo):
            kazu["D5"] += 1
            kiroku.append((ban, "D5", "%s の形式（%s）に合わない"
                           % (hyo, KEISHIKI["ニーズ" if hyo == "ニーズ" else "在宅"][1])))
            tsukigo_fuka.append(r)
            continue
        if meibo_s is not None and ban not in meibo_s:
            kazu["D6"] += 1
            kiroku.append((ban, "D6", "対象者名簿にない"))
            tsukigo_fuka.append(r)
            continue
        group[ban].append(r)

    # ── 同じ管理番号の束から1件を採る ────────────────
    saiyou = []
    for ban in sorted(group):
        g = group[ban]
        if len(g) > 1:
            n_web = sum(1 for r in g if _web(r))
            kigou = "D3" if n_web == len(g) else ("D2" if n_web else "D1")
            kazu[kigou] += 1
            kazu["除外"] += len(g) - 1
        g_sorted = sorted(g, key=lambda r: _rank(r, ent, hakuhyou))
        saiyou.append(g_sorted[0])
        for r in g_sorted[1:]:
            kiroku.append((ban, kigou,
                           "%s（受付 %s・回答%d問）を除外し、%s（受付 %s・回答%d問）を採った"
                           % (WEB if _web(r) else "紙", _uketsuke(r) or "―",
                              _kaitou_su(r, ent),
                              WEB if _web(g_sorted[0]) else "紙",
                              _uketsuke(g_sorted[0]) or "―",
                              _kaitou_su(g_sorted[0], ent))))

    saiyou += tsukigo_fuka
    kazu["採用"] = len(saiyou)
    kazu["突合できない"] = len(tsukigo_fuka)
    kazu["重複した番号"] = sum(1 for b in group if len(group[b]) > 1)
    return saiyou, kiroku, kazu


def meibo_kensa(saiyou, hyo):
    """採用した行のうち、属性のクロスに入れられる件数を数える"""
    ok = sum(1 for r in saiyou if r.get("_番号") and kata_ga_au(r["_番号"], hyo))
    return ok, len(saiyou) - ok


# ── 自己試験 ────────────────────────────────────
def _hyo_ent(kb, hyo):
    return [e for e in kb
            if (e["票"] == hyo or (hyo == "在宅" and e["票"].startswith("在宅")))
            and e["列"]]


def selftest():
    import random
    kb = SC.codebook()
    shippai = []

    def miru(namae, ok, kuwashiku=""):
        print("   %-46s %s%s" % (namae, "通った" if ok else "✗ 通らない",
                                 "　" + kuwashiku if kuwashiku else ""))
        if not ok:
            shippai.append("%s：%s" % (namae, kuwashiku))

    print("■ 自己試験　管理番号の正規化")
    for moto, machi in (("０００１", "1"), ("0001", "1"), ("1", "1"),
                        ("n0001", "N1"), ("N-0001", "N1"), ("Ｎ０００１", "N1"),
                        (" N0001 ", "N1"), ("N 0001", "N1")):
        deta = seiki(moto)[0]
        miru("「%s」→「%s」" % (moto, machi), deta == machi,
             "" if deta == machi else "得たのは「%s」" % deta)
    miru("空欄は空と判定される", seiki("　")[1] is True)
    miru("ニーズの形式：数字のみが通る", kata_ga_au("1", "ニーズ"))
    miru("ニーズの形式：英字つきは通らない", not kata_ga_au("N1", "ニーズ"))
    miru("在宅の形式：英字＋数字が通る", kata_ga_au("N1", "在宅"))
    miru("在宅の形式：数字のみは通らない", not kata_ga_au("1", "在宅"))

    print("\n■ 自己試験　重複の型（D1〜D6）")
    base, ent = SC._kirei(kb, "ニーズ"), _hyo_ent(kb, "ニーズ")
    base = base[0]

    def tsukuru(*shitei):
        """(管理番号, 回答方法, 受付日) の並びから行を作る"""
        out = []
        for ban, h, d in shitei:
            r = dict(base)
            r["管理番号"] = ban
            r[HOUHOU_COL] = h
            r[UKETSUKE_COL] = d
            out.append(r)
        return out

    for kigou, shitei, machi_saiyou in (
            ("D1", [("1", "紙", "2026/10/12"), ("1", "紙", "2026/10/13")], 1),
            ("D2", [("1", "紙", "2026/10/12"), ("1", WEB, "2026/10/13")], 1),
            ("D3", [("1", WEB, "2026/10/12"), ("1", WEB, "2026/10/13")], 1),
            ("D4", [("", WEB, "2026/10/13")], 1),
            ("D5", [("N1", "紙", "2026/10/13")], 1),
    ):
        rows = tsukuru(*shitei)
        sai, kir, kaz = dedupe(rows, "ニーズ", kb)
        deta = {k[1] for k in kir}
        miru("%s %s件 → 採用%d件" % (kigou, len(rows), len(sai)),
             kigou in deta and len(sai) == machi_saiyou,
             "" if kigou in deta else "鳴ったのは %s" % sorted(deta))

    rows = tsukuru(("1", "紙", "2026/10/13"))
    sai, kir, kaz = dedupe(rows, "ニーズ", kb, meibo={"0002"})
    miru("D6 名簿にない番号", "D6" in {k[1] for k in kir})

    rows = tsukuru(("1", "紙", "2026/10/13"))
    sai, kir, kaz = dedupe(rows, "ニーズ", kb, meibo={"0001"})
    miru("名簿にある番号（先頭の0が違う）は鳴らない", not kir,
         "" if not kir else "鳴った %s" % [k[1] for k in kir])

    rows = tsukuru(("1", "紙", "2026/10/13"))
    sai, kir, kaz = dedupe(rows, "ニーズ", kb, meibo=["0001", "1", "2"])
    deta = {k[1] for k in kir}
    miru("D7 名簿の番号が正規化で重なると突合しない",
         "D7" in deta and "D6" not in deta,
         "" if "D7" in deta else "鳴ったのは %s" % sorted(deta))

    print("\n■ 自己試験　採否の順序（①〜⑤）")
    #  ① 白票は採らない
    r_haku = dict(base); r_haku["管理番号"] = "1"
    r_haku[HOUHOU_COL] = WEB; r_haku[UKETSUKE_COL] = "2026/10/14"
    for e in [x for x in ent if x["区分"] == "必須"]:
        for c in e["列"]:
            r_haku.pop(c, None)
    r_yuu = dict(base); r_yuu["管理番号"] = "1"
    r_yuu[HOUHOU_COL] = "紙"; r_yuu[UKETSUKE_COL] = "2026/10/12"
    sai, _k, _z = dedupe([r_haku, r_yuu], "ニーズ", kb)
    miru("① 受け付けが後でも白票は採らない", len(sai) == 1 and not _web(sai[0]))

    #  ② 受け付けが後
    rows = tsukuru(("1", "紙", "2026/10/12"), ("1", "紙", "2026/10/14"))
    sai, _k, _z = dedupe(rows, "ニーズ", kb)
    miru("② 受け付けが後のものを採る",
         len(sai) == 1 and _uketsuke(sai[0]) == "2026-10-14")

    #  ③ 同じ日ならウェブ
    rows = tsukuru(("1", "紙", "2026/10/13"), ("1", WEB, "2026/10/13"))
    sai, _k, _z = dedupe(rows, "ニーズ", kb)
    miru("③ 同じ日ならウェブを採る", len(sai) == 1 and _web(sai[0]))

    #  ④ 回答した設問が多いもの
    r_oo = dict(base); r_oo["管理番号"] = "1"
    r_oo[HOUHOU_COL] = "紙"; r_oo[UKETSUKE_COL] = "2026/10/13"
    r_suku = dict(r_oo)
    nuku = [x for x in ent if x["区分"] != "必須"][:5]
    for e in nuku:
        for c in e["列"]:
            r_suku.pop(c, None)
    sai, _k, _z = dedupe([r_suku, r_oo], "ニーズ", kb)
    miru("④ 回答した設問が多いものを採る",
         len(sai) == 1 and _kaitou_su(sai[0], ent) == _kaitou_su(r_oo, ent))

    #  ⑤ ①〜④が並ぶ2件を入れ替えても同じ1件を採る
    #     ここが読み込んだ順で決まると、実行のたびに結果が変わりうる。
    tan = [x for x in ent if x["形式"] == "単一" and len(x["選択肢"]) >= 3]
    r_a = dict(base); r_a["管理番号"] = "1"
    r_a[HOUHOU_COL] = "紙"; r_a[UKETSUKE_COL] = "2026/10/13"
    r_b = dict(r_a)
    #  回答数は変えずに、中身だけを変える（①〜④はすべて並ぶ）
    for e in tan[:3]:
        ima = str(r_a.get(e["変数"], ""))
        hoka = [str(n) for n, _l in e["選択肢"] if str(n) != ima]
        if hoka:
            r_b[e["変数"]] = hoka[0]
    r_a["_id"], r_b["_id"] = 0, 1
    narabi = (_rank(r_a, ent, set())[:5] == _rank(r_b, ent, set())[:5])
    sai1, _k, _z = dedupe([dict(r_a), dict(r_b)], "ニーズ", kb)
    sai2, _k, _z = dedupe([dict(r_b), dict(r_a)], "ニーズ", kb)
    miru("⑤ ①〜④が並ぶ2件は、入れ替えても同じ1件を採る",
         narabi and len(sai1) == 1 and len(sai2) == 1
         and _naiyou(sai1[0]) == _naiyou(sai2[0]),
         "①〜④が並んでいない（試験が成り立たない）" if not narabi else "")

    #  読み込んだ順序を入れ替えても結果が変わらない
    print("\n■ 自己試験　読み込んだ順序に依存しないこと")
    needs = os.path.join(P.DATA, "dummy_ニーズ.csv")
    if os.path.exists(needs):
        moto = list(csv.DictReader(io.open(needs, encoding="utf-8-sig")))
        kijun = None
        onaji = True
        for kai in range(5):
            tame = [dict(r) for r in moto]
            random.Random(100 + kai).shuffle(tame)
            sai, _k, kaz = dedupe(tame, "ニーズ", kb)
            ima = (kaz["採用"], kaz["除外"],
                   sorted(_naiyou(r) for r in sai))
            if kijun is None:
                kijun = ima
            elif ima != kijun:
                onaji = False
        miru("ダミーデータを5回入れ替えても採用が一致する", onaji,
             "採用%d件・除外%d件" % (kijun[0], kijun[1]) if onaji else "食い違う")
    else:
        shippai.append("ダミーデータがないため順序の試験ができない")

    #  ⑥ 件数の帳尻
    print("\n■ 自己試験　件数の帳尻")
    rows = tsukuru(("1", "紙", "2026/10/12"), ("1", WEB, "2026/10/13"),
                   ("2", "紙", "2026/10/12"), ("", WEB, "2026/10/13"),
                   ("N9", "紙", "2026/10/13"))
    sai, kir, kaz = dedupe(rows, "ニーズ", kb)
    miru("読み込み − 除外 ＝ 採用",
         kaz["読み込み"] - kaz["除外"] == kaz["採用"],
         "読み込み%d − 除外%d ＝ 採用%d"
         % (kaz["読み込み"], kaz["除外"], kaz["採用"]))

    print()
    if shippai:
        print("■ 自己試験は通りませんでした（%d件）" % len(shippai))
        for x in shippai:
            print("   ・%s" % x)
        return 1
    print("■ 自己試験　すべて通りました")
    return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--needs", default=os.path.join(P.DATA, "dummy_ニーズ.csv"))
    ap.add_argument("--zaitaku", default=os.path.join(P.DATA, "dummy_在宅.csv"))
    ap.add_argument("--out", default=None)
    ap.add_argument("--自己試験", action="store_true", dest="jiko")
    a = ap.parse_args()

    if a.jiko:
        return selftest()

    chigai = SC.ban_chigai()
    if chigai:
        print("※ コードブックの版が合っていません（%s）。" % chigai)
        print("   白票の判定と回答数の数えに影響するため、結果は暫定です。\n")

    kb = SC.codebook()
    zen = []
    for hyo, path in (("ニーズ", a.needs), ("在宅", a.zaitaku)):
        if not os.path.exists(path):
            print("  %s：%s がない" % (hyo, path))
            continue
        rows = list(csv.DictReader(io.open(path, encoding="utf-8-sig")))
        sai, kir, kaz = dedupe(rows, hyo, kb)
        zen += [(hyo,) + k for k in kir]
        print("■ %s" % hyo)
        print("   読み込み %d件 − 除外 %d件 ＝ 採用 %d件（白票 %d件を含む）"
              % (kaz["読み込み"], kaz["除外"], kaz["採用"], kaz["白票"]))
        for kigou, nm, _itsu, _dou in KATA:
            if kaz.get(kigou):
                print("   %s %-20s %4d件" % (kigou, nm, kaz[kigou]))
        ok, ng = meibo_kensa(sai, hyo)
        print("   うち属性のクロスに入れられる %d件／入れられない %d件" % (ok, ng))
        if not kir:
            print("   重複・不備はありません")

    if a.out:
        with io.open(a.out, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["票", "管理番号", "記号", "内容"])
            w.writerows(zen)
        print("\n書き出し：%s（%d件）" % (a.out, len(zen)))
    if chigai:
        print("\n※ 上の結果は暫定です。コードブックの版が合っていません。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
