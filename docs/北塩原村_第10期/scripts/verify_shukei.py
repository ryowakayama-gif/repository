# -*- coding: utf-8 -*-
"""集計仕様書の自己点検。

   分母と指標の定義が、集計が始まってから崩れることを防ぐ。
   不適合があれば終了コード1を返す。
"""
import os as _os_p
import sys as _sys_p
_sys_p.path.insert(0, _os_p.path.dirname(_os_p.path.abspath(__file__)))
import paths as _P   # 置き場所はここで決める（じか書きしない）
import io, os, re, sys
sys.dont_write_bytecode = True   # 古いバイトコードで誤った結果が出ることを防ぐ
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from shukei_data import N, Z, DERIVED, CROSS, AXES, SHU_HO
import json
import shukei_check as SC   # 精度管理は実際に走る側から引く
import shukei_dedupe as DD  # 重複回答の排除も同じく

MD = os.path.join(_P.BASE, "25_集計仕様書.md")
MD24 = os.path.join(_P.BASE, "24_アンケート調査報告書_骨子案.md")
R = []
def chk(no, name, ok, detail=""):
    R.append((no, name, ok, detail))

md = open(MD, encoding="utf-8").read()
md24 = open(MD24, encoding="utf-8").read()   # 報告書骨子案（doc24）

# 1 クロス表すべてに分母が書かれていること
nob = [c[0] for c in CROSS if not str(c[5]).strip() or str(c[5]).strip() == "―"]
chk(1, "クロス表に分母が書かれていること", len(CROSS) >= 20 and not nob,
    f"分母が空 {nob}" if nob else f"{len(CROSS)}表すべてに分母を記載")

# 2 分母に「回収票の総数」を使っていないこと
bad2 = [c[0] for c in CROSS if "回収票" in str(c[5]) and "総数" in str(c[5])]
chk(2, "分母に回収票の総数を用いていないこと", not bad2,
    f"総数で割っている {bad2}" if bad2 else f"{len(CROSS)}表とも設問または両設問の有効回答")

# 3 分母の書き方が「設問」「両設問」「名簿由来」のいずれかに揃っていること
KEY = ("両設問", "設問", "有効回答", "名簿由来")
bad3 = [c[0] for c in CROSS if not any(k in str(c[5]) for k in KEY)]
chk(3, "分母の書き方が揃っていること", not bad3,
    f"書き方が不揃い {bad3}" if bad3 else f"{len(CROSS)}表とも定型の語で記載")

# 4 主指標・補足指標の整理があり、主に分母が書かれていること
nom = [x[0] for x in SHU_HO if "分母" not in x[1] and "データ" not in x[1] and "判定" not in x[1]]
chk(4, "主指標に出所と分母が書かれていること", len(SHU_HO) >= 5 and not nom,
    f"分母がない {nom}" if nom else f"{len(SHU_HO)}件すべてに出所を記載")

# 5 基本方針が11項目で、9〜11が md にあること
need5 = ["分母は設問ごとの有効回答に統一", "同じ指標を複数の設問から作らない", "表ごとに分母とn数を書く"]
miss5 = [x for x in need5 if x not in md]
chk(5, "基本方針9〜11が集計仕様書にあること", "集計の基本方針（11項目）" in md and not miss5,
    f"欠落 {miss5}" if miss5 else "11項目・方針9〜11を記載")

# 6 主指標・補足指標の表が md にもあること（xlsx と md の両方に置く）
miss6 = [x[0] for x in SHU_HO if x[0] not in md]
chk(6, "主指標・補足指標が集計仕様書にあること", not miss6,
    f"mdにない {miss6}" if miss6 else f"{len(SHU_HO)}件すべて記載")

# 7 在宅調査の要介護度が2区分に揃っていること（名簿に要支援がいないため）
bad7 = [c[0] for c in CROSS if c[4] == "在宅" and "I（3区分）" in (c[2], c[3])]
chk(7, "在宅調査の要介護度が2区分であること", not bad7,
    f"3区分のまま {bad7}" if bad7 else "在宅のクロスは2区分に統一")

# 8 n<30 の断りが残っていること
chk(8, "n<30を参考値とする断りがあること", "n<30" in md and "参考値" in md)

# ── ここから doc24（報告書骨子案）との照合 ────────────────
#    集計仕様書を正とし、報告書の書き方が食い違っていないことを確かめる。

# 9 分母の方針が報告書骨子にも書かれていること（方針9・11に対応）
need9 = [
    "構成比の分母は、その設問に回答した票の数（設問ごとの有効回答）です",
    "回収票の総数で割った値は用いません",
    "行・列の双方に回答した票の数です",
    "時点の異なる値を足して合計としません",
]
miss9 = [x[:16] for x in need9 if x not in md24]
chk(9, "分母の方針が報告書骨子にもあること", not miss9,
    f"欠落 {miss9}" if miss9 else "分母の定義・総数で割らない・両設問・時点を記載")

# 10 主指標・補足指標の5件が報告書骨子にも書かれていること（方針10）
miss10 = [x[0] for x in SHU_HO if x[0] not in md24]
chk(10, "主指標・補足指標が報告書骨子にもあること", not miss10,
    f"doc24にない {miss10}" if miss10 else f"{len(SHU_HO)}件すべて記載")

# 11 クロス表の分母が、集計仕様書の定義のまま報告書骨子に書かれていること
miss11 = [c[0] for c in CROSS if c[0] not in md24 or str(c[5]) not in md24]
chk(11, "クロス表の分母が報告書骨子と一致すること", not miss11,
    f"不一致 {miss11}" if miss11 else "X-01〜X-20の分母が doc24 §13 と一致")

# 12 在宅調査の要介護度が、報告書骨子でも2区分であること
bad12 = "要介護度（要支援1・2" in md24
chk(12, "報告書骨子の在宅の要介護度が2区分であること",
    not bad12 and "要介護1・2／要介護3〜5の2区分" in md24,
    "3区分の記載が残っている" if bad12 else "2区分に統一（名簿に要支援なし）")

# 13 問7の枝番が集計仕様書と一致していること（村独自3問は(9)(10)(11)）
old13 = [x for x in ("問7(1)〜(10)", "問7(8)(9)", "問7(8)(9)(10)") if x in md24]
q7 = [q for (_, q, e, *_r) in N if q == "問7"]
chk(13, "問7の枝番が集計仕様書と一致すること",
    not old13 and "問7(11)" in md24 and len(q7) == 11,
    f"旧番号が残っている {old13}" if old13 else "村独自3問は問7(9)(10)(11)・全11枝")

# 14 リスク判定の領域数が一致していること（R01〜R10＋R99）
rcode = [d[0] for d in DERIVED if d[0].startswith("R") and d[0] != "R99"]
miss14 = [c for c in rcode if c not in md24]
chk(14, "リスク判定が10領域＋R99で一致すること",
    len(rcode) == 10 and not miss14 and "R99" in md24 and "R01〜R11" not in md24,
    f"doc24にない {miss14}" if miss14 else "R01〜R10・R99が両文書で一致")

# 15 集計軸Iの区分が2区分に揃っていること（凡例・md・doc24）
axI = [a for a in AXES if a[0] == "I"][0]
chk(15, "集計軸Iの区分が2区分であること",
    "要支援" not in axI[2] and "要介護1・2／要介護3〜5（2区分）" in md,
    f"軸Iの区分 {axI[2]}")

# 16 中核の★表の数が、CROSS と両文書の記述で一致すること
star = [c[0] for c in CROSS if str(c[6]).startswith("★")]
hokyo = [c[0] for c in CROSS if not str(c[6]).startswith("★")]
say = f"★を付した{len(star)}表が分析編"
chk(16, "中核の★表の数が一致すること",
    say in md and f"**★を付した{len(star)}表が分析編（第Ⅱ部）の中核**" in md24
    and f"残る{len(hokyo)}表" in md,
    f"★{len(star)}表・補強{len(hokyo)}表")

# ── ここから精度管理（doc25 §7）との照合 ───────────────
#    論理チェックは shukei_check.py が実際に走る側。仕様書がそれと
#    食い違えば、村にお示しした内容と現に走る点検が別物になる。

# 17 論理チェックの12項目が、記号も項目名も集計仕様書にあること
miss17 = [k for k, nm, _i, _n in SC.RULES if k not in md or nm not in md]
chk(17, "論理チェックの項目が集計仕様書にあること",
    len(SC.RULES) == 12 and not miss17 and "論理チェック（12項目）" in md,
    f"mdにない {miss17}" if miss17 else f"L01〜L12の{len(SC.RULES)}項目を記載")

# 18 コードブックの点検5項目が、記号も項目名も集計仕様書にあること
miss18 = [k for k, nm, _i, _n in SC.CB_RULES if k not in md or nm not in md]
chk(18, "コードブックの点検が集計仕様書にあること",
    len(SC.CB_RULES) == 5 and not miss18 and "コードブックの点検（5項目）" in md,
    f"mdにない {miss18}" if miss18 else f"C01〜C05の{len(SC.CB_RULES)}項目を記載")

# 19 ベリファイの3つの扱いが集計仕様書に書かれていること
need19 = ["全件を二重入力して照合", "10パーセントを抽出して照合", "対象外",
          "入力の誤りの件数 ÷ 照合した項目数"]
miss19 = [x for x in need19 if x not in md]
chk(19, "ベリファイの対象と方法が集計仕様書にあること", not miss19,
    f"欠落 {miss19}" if miss19 else "全件・抽出・対象外・記録の仕方を記載")

# 20 分岐の規則が、親も子も集計仕様書にあること
miss20 = [f"{oya}→{ko}" for _h, oya, _j, ko, _k in SC.BRANCH
          if oya not in md or ko not in md]
chk(20, "分岐の規則が集計仕様書にあること", len(SC.BRANCH) >= 5 and not miss20,
    f"mdにない {miss20}" if miss20 else f"{len(SC.BRANCH)}件の親子を記載")

# 21 コードブックの版が合っていないことの断りがあること
#    版が合えば ban_chigai() が None を返すので、断りは不要になる
chigai21 = SC.ban_chigai()
need21 = ["令和8年9月4日版", "校了版", "暫定", "--暫定"]
miss21 = [x for x in need21 if x not in md]
chk(21, "コードブックの版の断りが集計仕様書にあること",
    (not chigai21) or (not miss21),
    f"欠落 {miss21}" if miss21 else
    ("版が合っているため断りは不要" if not chigai21 else "版違い・暫定の断りを記載"))

# 22 §1 のシートの件数が shukei_data と一致すること
need22 = [f"**{len(DERIVED)}件**", f"**{len(CROSS)}表**",
          f"**{len(N)}設問**", f"**{len(Z)}設問**"]
miss22 = [x for x in need22 if x not in md]
chk(22, "シートの件数が集計仕様書と一致すること", not miss22,
    f"§1と不一致 {miss22}" if miss22 else
    f"ニーズ{len(N)}・在宅{len(Z)}・派生{len(DERIVED)}・クロス{len(CROSS)}")

# 23 派生変数がすべて集計仕様書に載っていること
miss23 = [d[0] for d in DERIVED if d[0] not in md]
chk(23, "派生変数がすべて集計仕様書にあること", not miss23,
    f"mdにない {miss23}" if miss23 else f"{len(DERIVED)}件すべて記載")

# 24 上限・範囲の参照先がコードブックに実在すること（C05 を点検側でも押さえる）
kb24 = json.load(open(os.path.join(_P.DATA, "集計_コードブック.json"), encoding="utf-8"))
v24 = {e["変数"] for e in kb24}
c24 = {c for e in kb24 for c in e["列"]}
bad24 = ([v for v in SC.LIMIT if v not in v24]
         + [c for c in SC.RANGE if c not in c24])
chk(24, "上限・範囲の参照先がコードブックにあること", not bad24,
    f"参照先がない {bad24}" if bad24 else
    f"LIMIT {len(SC.LIMIT)}件・RANGE {len(SC.RANGE)}件とも実在")

# 25 点検そのものが自己試験を通ること（欠陥を入れて鳴ることを確かめた記録）
#    selftest() は標準出力に書くため、ここでは結果だけを見る
import contextlib
with contextlib.redirect_stdout(io.StringIO()):
    jiko = SC.selftest()
chk(25, "論理チェックが自己試験を通ること", jiko == 0,
    "自己試験が通らない" if jiko else "欠陥を入れて19件すべてが鳴ることを確認")

# ── ここから重複回答の排除（doc25 §8）との照合 ──────────────

# 26 重複の型7件が、記号も型名も集計仕様書にあること
miss26 = [k for k, nm, _n, _d in DD.KATA if k not in md or nm not in md]
chk(26, "重複の型が集計仕様書にあること",
    len(DD.KATA) == 7 and not miss26 and "重複と不備の型（D1〜D7）" in md,
    f"mdにない {miss26}" if miss26 else f"D1〜D7の{len(DD.KATA)}件を記載")

# 27 採否の順序5段が、段も採る方も集計仕様書にあること
miss27 = [d for d, toru, _r in DD.SAIHI if d not in md or toru not in md]
chk(27, "採否の順序が集計仕様書にあること",
    len(DD.SAIHI) == 5 and not miss27,
    f"mdにない {miss27}" if miss27 else f"①〜⑤の{len(DD.SAIHI)}段を記載")

# 28 管理番号の正規化が、点検側と排除側で同じ規則であること
#    規則が分かれると、点検は「重複なし」と言い、排除は1件落とす
chk(28, "管理番号の正規化が点検と排除で同じ規則であること",
    DD.seiki is SC.seiki and DD.kata_ga_au is SC.kata_ga_au
    and DD.KEISHIKI is SC.KEISHIKI,
    "規則が分かれている" if DD.seiki is not SC.seiki else
    "shukei_dedupe は shukei_check.seiki を使っている")

# 29 管理番号の形式が doc43 のとおり集計仕様書にあること
need29 = ["数字のみ", "英字＋数字", "文字列として扱"]
miss29 = [x for x in need29 if x not in md]
chk(29, "管理番号の形式が集計仕様書にあること", not miss29,
    f"欠落 {miss29}" if miss29 else "ニーズ＝数字のみ・在宅＝英字＋数字・文字列扱いを記載")

# 30 名簿に当たらない回答を捨てない方針が書かれていること
need30 = ["捨てません", "単純集計", "属性のクロス", "重複は見つけられません"]
miss30 = [x for x in need30 if x not in md]
chk(30, "名簿に当たらない回答の扱いが集計仕様書にあること", not miss30,
    f"欠落 {miss30}" if miss30 else "単純集計には入れ、属性のクロスには入れない旨を記載")

# 31 2つの名簿に重複する2人の扱いが書かれ、村の確認事項になっていること
from wbs_kakunin import K as KAKUNIN
k31 = [k for k in KAKUNIN if "2つの名簿" in k[1] or "在宅介護実態調査の名簿にも" in k[2]]
need31 = ["2つの名簿に重複する2人", "重複回答ではありません"]
miss31 = [x for x in need31 if x not in md] + ([] if k31 else ["確認事項への起票"])
chk(31, "2つの名簿に重複する2人の扱いが集計仕様書と確認事項にあること", not miss31,
    f"欠落 {miss31}" if miss31 else "doc25 §8-5 に記載・確認事項に起票済み")

# 32 重複回答の排除が自己試験を通ること
with contextlib.redirect_stdout(io.StringIO()):
    jiko32 = DD.selftest()
chk(32, "重複回答の排除が自己試験を通ること", jiko32 == 0,
    "自己試験が通らない" if jiko32 else
    "正規化・型D1〜D7・採否①〜⑤・順序に依存しないこと・件数の帳尻を確認")

# 33 ダミーデータにわざと重複が入っていること
#    0件のダミーでは「鳴らなかった」ことしか示せない
import shukei_dummy as DM
chk(33, "ダミーデータにわざと重複・不備が入っていること",
    sum(c for _k, c, _t in DM.JUUFUKU) >= 5 and len(DM.JUUFUKU) >= 5,
    f"{len(DM.JUUFUKU)}型・計{sum(c for _k, c, _t in DM.JUUFUKU)}件を仕込んでいる")

# 34 集計本体が重複を落としたデータで動くこと
#    shukei_run.py が排除を通さないと、紙とウェブの両方で答えた方が2件になる
run_py = io.open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "shukei_run.py"), encoding="utf-8").read()
#    文字列を1か所見るだけでは、片方の票だけ素通しにしても通ってしまう。
#    2票それぞれが排除を通り、派生変数が排除後の行から作られることを見る。
bad34 = []
if "import shukei_dedupe" not in run_py:
    bad34.append("shukei_dedupe を取り込んでいない")
if run_py.count("DD.dedupe(") != 2:
    bad34.append("DD.dedupe の呼び出しが%d回（2票で2回であるべき）"
                 % run_py.count("DD.dedupe("))
for yobi in ('DD.dedupe(needs_raw, "ニーズ"', 'DD.dedupe(zai_raw, "在宅"'):
    if yobi not in run_py:
        bad34.append("%s がない" % yobi)
for tsukai in ("derive_needs(r) for r in n_sai", "derive_zaitaku(r) for r in z_sai"):
    if tsukai not in run_py:
        bad34.append("%s がない（排除後の行を使っていない）" % tsukai)
if "for r in read_csv(" in run_py:
    bad34.append("read_csv の結果を直に派生変数へ渡している")
chk(34, "集計本体が重複を落としたデータで動くこと", not bad34,
    "／".join(bad34[:2]) if bad34 else
    "2票それぞれが DD.dedupe を通り、派生変数は排除後の行から作られている")

# 35 ダミーデータの件数が、集計プログラムの文書（doc31）と合っていること
md31 = io.open(os.path.join(_P.BASE, "31_集計プログラムの構成と実行手順.md"),
               encoding="utf-8").read()
tasu = sum(c for _k, c, _t in DM.JUUFUKU)
need35 = ["ニーズ%d件" % (DM.N_NEEDS + tasu), "在宅%d件" % (DM.N_ZAITAKU + tasu),
          "ニーズ%d件" % (DM.N_NEEDS + tasu - 6), "在宅%d件" % (DM.N_ZAITAKU + tasu - 6)]
miss35 = [x for x in need35 if x not in md31]
chk(35, "ダミーデータの件数が集計プログラムの文書と合うこと", not miss35,
    f"doc31にない {miss35}" if miss35 else
    f"読み込み {DM.N_NEEDS + tasu}／{DM.N_ZAITAKU + tasu}・"
    f"集計 {DM.N_NEEDS + tasu - 6}／{DM.N_ZAITAKU + tasu - 6}")

w = max(len(n) for _, n, _, _ in R)
print("■ 集計仕様書の自己点検")
ng = 0
for no, name, ok, detail in R:
    print(f"  {no:>3}  {name:<{w}}  {'適合' if ok else '不適合'}" + (f"  {detail}" if detail else ""))
    ng += 0 if ok else 1
print(f"\n  {len(R)}件のうち適合{len(R)-ng}件／不適合{ng}件")
sys.exit(1 if ng else 0)
