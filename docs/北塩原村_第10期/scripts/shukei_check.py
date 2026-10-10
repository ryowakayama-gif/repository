# -*- coding: utf-8 -*-
"""入力データの精度管理 ― 論理チェックとベリファイ（WBS Ⅰ-104）

   回収した調査票を入力したデータに、入力の誤りや回答の矛盾がないかを機械で見る。
   集計を回す前に通し、鳴った件を1件ずつ原票に当たって直す。

     python3 scripts/shukei_check.py                         ダミーデータで通す
     python3 scripts/shukei_check.py --needs <csv> --zaitaku <csv> --out <csv>

   【コードブックの版に依存する】
   本スクリプトの判定は、設問の番号・選択肢・分岐をコードブックから引く。
   コードブックが実際に配られた調査票と違う版であれば、
   **鳴るべきものが鳴らず、鳴らなくてよいものが鳴る。**
   このため、走らせる前に版を確かめ、食い違っていれば止める。
   承知のうえで動かすときは --暫定 を付ける（結果は暫定である旨を明記して出す）。

   【ベリファイの方針】
   ベリファイ（再入力による照合）は、入力の誤りを見つける唯一の確かな方法である。
   全件を二重入力する余力はないため、次のように分ける。

     全件を二重入力して照合するもの
       ① 区分が「必須」の設問
       ② 派生変数に使う設問（リスク判定 R01〜R10、生活機能4層 L01、支え合い S01・S02）
       理由：リスク判定と生活機能4層は、資料3の3-3（元気な軽度者）・
             3-4（生活機能からみた4つの層）・3-6（担い手の可能性）の基礎である。
             1件の入力の誤りが層の人数を動かし、基本理念・基本目標の判断材料が狂う。

     10%を抽出して照合するもの
       上記以外の設問（オプション・村独自）
       理由：単純集計の割合に効くが、層の判定には使わない。
             抽出で誤りが見つかった場合は、その設問を全件に広げる。

     ベリファイの対象外
       ウェブ回答（入力を経ないため）。ただし取込の件数が
       フォームの回答数と一致することは確かめる。

   数えた結果は「入力の誤りの件数 ÷ 照合した項目数」として記録する。
   doc25（集計仕様書）に載せ、村へお示しする。
"""
import argparse
import collections
import csv
import io
import json
import os
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths as P                                   # noqa: E402
import shukei_codebook as CB                        # noqa: E402

KB_JSON = os.path.join(P.DATA, "集計_コードブック.json")

# 実際に配られた調査票の版。回答データはこの版に従う。
HITSUYOU_BAN = "令和8年9月9日 校了版"

# ── 論理チェックの項目 ────────────────────────────────
#    (記号, 名, いつ鳴るか, 直し方)
RULES = [
 ("L01", "白票",
  "区分が「必須」の設問がすべて無回答",
  "有効回答から外す。外した件数を回収数とは別に記録する"),
 ("L02", "必須設問の無回答",
  "区分が「必須」の設問が無回答（白票を除く）",
  "原票に当たる。原票でも空欄なら無回答のまま集計する（基本方針7）"),
 ("L03", "選択肢の範囲外",
  "単一回答の値が、コードブックの選択肢にない"
  "（選択肢が1つ以下の設問は判定を省く。C02 を先に直す）",
  "入力の誤り。原票に当たって直す"),
 ("L04", "単一回答に複数の印",
  "単一回答の列に区切り文字を含む値が入っている",
  "原票に当たる。原票で複数に印があれば「無回答」とする（基本方針8）"),
 ("L05", "複数回答の上限超え",
  "「3つまで」と定めた設問で4つ以上が選ばれている",
  "原票に当たる。原票でも上限を超えていれば、そのまま集計し件数を記録する"),
 ("L06", "答えるべきなのに無回答",
  "分岐の親の条件を満たすのに、子の設問がすべて無回答",
  "原票に当たる。回答者が飛ばしたのであれば無回答のまま"),
 ("L07", "答えるべきでないのに回答",
  "分岐の親の条件を満たさないのに、子の設問に回答がある",
  "原票に当たる。入力の誤りでなければ、親の回答の誤りを疑う"),
 ("L08", "数値の範囲外",
  "数値の設問が、定めた範囲の外にある",
  "原票に当たって直す"),
 ("L09", "管理番号の不備",
  "管理番号が空、または同じ番号が2件以上ある",
  "重複は L09 で挙げ、重複回答の排除（Ⅰ-36）で扱う"),
 ("L10", "名簿にない管理番号",
  "対象者名簿にない管理番号",
  "名簿は個人情報のため、照合は件数と番号の集合のみで行う"),
 ("L11", "認定データとの矛盾",
  "認定状況が「認定なし」なのに要介護度が入っている等",
  "要介護認定データの受領後に実施する（Ⅰ-37）"),
 ("L12", "直線回答（参考）",
  "4択以上の単一回答で、同じ選択肢番号が10問以上続いている"
  "（2択が連なる箇所は偶然でも続くため数えない）",
  "誤りとは限らない。自由記述とあわせて見て、無効とするかを判断する"),
]

# ── 分岐の規則 ──────────────────────────────────
#    (票, 親の変数, 親の条件, 子の変数, 根拠)
#    親の条件は「この値のときに子を答える」。複数回答の親は、
#    その選択肢が選ばれていれば条件を満たすものとする。
BRANCH = [
 ("ニーズ", "N_問1_2", (2, 3), "N_問1_2_1", "問1(2)で「1．必要ない」以外の方のみ"),
 ("ニーズ", "N_問1_2", (3,), "N_問1_2_2", "問1(2)で「3．現在受けている」の方のみ"),
 ("ニーズ", "N_問3_6", (1, 3), "N_問3_6_2", "問3(6)で入れ歯を利用の方のみ"),
 ("在宅A", "ZA_問8", (2,), "ZA_問10", "問8で「2．利用していない」の方のみ"),
 ("在宅B", "ZB_問7", (1, 2), "ZB_問8", "問7で「1．」「2．」の方のみ"),
]

# ── 複数回答の上限（鍵は変数名）─────────────────────────
LIMIT = {"ZB_問6": 3, "ZB_問9": 3}

# ── 数値の設問の範囲（鍵はデータの列名）──────────────────────
#    数値の設問は1問が複数の列に分かれることがある（身長・体重）。
#    変数名ではなく列名で持つ。C05 が、ここに書いた列が実在するかを見る。
RANGE = {
 "N_問3_1_身長": (100, 200),      # cm
 "N_問3_1_体重": (25, 150),       # kg
 "N_問8_2": (0, 10),              # 主観的幸福感（0〜10点）
}

# ── L12 直線回答 ────────────────────────────────
#    2択の設問（問4の「はい・いいえ」など）が連なる箇所では、
#    同じ値が10問続くのはごく普通に起きる（問4(4)〜(16)は13問が2択）。
#    偶然で起きにくい4択以上の単一回答だけを数える。
CHOKUSEN = 10          # 直線回答とみなす連続の数
CHOKUSEN_MIN = 4       # 数える対象とする選択肢の数の下限
KUGIRI = ",，、 /・"    # L04 単一回答に複数の印とみなす区切り

# ── コードブック自体の点検 ───────────────────────────
#    論理チェックは設問の番号・選択肢・分岐をコードブックから引く。
#    コードブックが壊れていれば、鳴った件はデータの誤りではなく
#    コードブックの不備である。データを見る前にここで分けておく。
CB_RULES = [
 ("C01", "変数の重複",
  "同じ変数が2件以上定義されている",
  "調査票に当たってどちらが正しいかを決め、片方を消す"),
 ("C02", "単一回答の選択肢が足りない",
  "単一回答なのに選択肢が1つ以下（調査票から拾えていない疑い）",
  "調査票の表を見て OVERRIDES に書く。直すまで L03 の判定は省く"),
 ("C03", "複数回答の列と選択肢が合わない",
  "複数回答の列の数が選択肢の数と違う",
  "選択肢を数え直す。列は選択肢ごとに1つ"),
 ("C04", "分岐の参照先がない",
  "BRANCH に書いた親または子の変数がコードブックにない",
  "設問番号の振り直しを疑う。BRANCH を直す"),
 ("C05", "上限・範囲の参照先がない",
  "LIMIT の変数、または RANGE の列がコードブックにない",
  "LIMIT・RANGE を直す。参照先がないと L05・L08 は何も見ない"),
]


def codebook():
    return json.load(io.open(KB_JSON, encoding="utf-8"))


def ban_chigai():
    """コードブックの版が、実際に配られた調査票と食い違っていれば理由を返す"""
    moto = getattr(CB, "BUILT_FROM", None)
    if moto is None:
        return "コードブックがどの版から作られたかが記録されていない"
    if moto != HITSUYOU_BAN:
        return ("コードブックは%s から作られている。"
                "実際に配られたのは%s である" % (moto, HITSUYOU_BAN))
    return None


def codebook_check(kb):
    """コードブック自体を見る。戻り値は (記号, 対象, 内容) の並び"""
    out = []
    # C01 変数の重複
    tally = collections.Counter(e["変数"] for e in kb)
    for v, n in sorted(tally.items()):
        if n > 1:
            ban = "・".join(sorted({"%s %s" % (e["票"], e["番号"])
                                    for e in kb if e["変数"] == v}))
            out.append(("C01", v, "%d件が同じ変数を使っている（%s）" % (n, ban)))
    for e in kb:
        tgt = "%s %s" % (e["票"], e["番号"])
        # C02 単一回答の選択肢が足りない
        if e["形式"].startswith("単一") and len(e["選択肢"]) <= 1 and e["列"]:
            out.append(("C02", tgt,
                        "単一回答だが選択肢が%d個しかない" % len(e["選択肢"])))
        # C03 複数回答の列と選択肢が合わない
        if e["形式"].startswith("複数") and len(e["列"]) != len(e["選択肢"]):
            out.append(("C03", tgt, "列が%d、選択肢が%d"
                        % (len(e["列"]), len(e["選択肢"]))))
    # C04 分岐の参照先
    byvar = {e["変数"] for e in kb}
    for h, oya, _j, ko, _k in BRANCH:
        for yaku, v in (("親", oya), ("子", ko)):
            if v not in byvar:
                out.append(("C04", "%s %s" % (h, v),
                            "分岐の%sの変数がコードブックにない" % yaku))
    # C05 上限・範囲の参照先
    for v in sorted(LIMIT):
        if v not in byvar:
            out.append(("C05", v, "LIMIT の変数がコードブックにない"))
    cols = {c for e in kb for c in e["列"]}
    for c in sorted(RANGE):
        if c not in cols:
            out.append(("C05", c, "RANGE の列がコードブックにない"))
    return out


def _branch_children():
    """子の変数 → [(親の変数, 条件)]。L02 は条件を満たさない子に鳴らせない"""
    d = collections.defaultdict(list)
    for _h, oya, jouken, ko, _k in BRANCH:
        d[ko].append((oya, jouken))
    return d


def _val(r, col):
    return str(r.get(col, "")).strip()


def _multi_selected(r, e):
    """複数回答で選ばれている選択肢の番号"""
    return [n for n, _lab in e["選択肢"]
            if _val(r, "%s_c%s" % (e["変数"], n)) == "1"]


def _answered(r, e):
    if e["形式"].startswith("複数"):
        return bool(_multi_selected(r, e))
    return _val(r, e["変数"]) != ""


def _mitasu(r, eo, jouken):
    """分岐の親の条件を満たすか。親が無回答なら None（どちらとも言えない）"""
    if eo["形式"].startswith("複数"):
        if not _answered(r, eo):
            return None
        return bool(set(_multi_selected(r, eo)) & set(jouken))
    vo = _val(r, eo["変数"])
    if vo == "":
        return None
    return vo in {str(x) for x in jouken}


def check(rows, hyo, kb, meibo=None):
    """1つの票の回答データを見る。戻り値は (管理番号, 記号, 設問, 内容) の並び

       マトリクスの親（列を持たない項目）は、データを持たないので対象外。
       分岐の子は、親の条件を満たすときだけ必須とみなす（満たさない件は L07）。
       選択肢がコードブックに1つ以下しかない設問は、L03 の判定を省く（C02）。
    """
    out = []
    ent = [e for e in kb
           if (e["票"] == hyo or (hyo == "在宅" and e["票"].startswith("在宅")))
           and e["列"]]                   # 列を持たない＝マトリクスの親は対象外
    byvar = {e["変数"]: e for e in ent}
    ko = _branch_children()
    tan = [e for e in ent
           if e["形式"] == "単一" and len(e["選択肢"]) >= CHOKUSEN_MIN]
    hissu = [e for e in ent if e["区分"] == "必須" and e["変数"] not in ko]

    # L09 管理番号
    ban = collections.Counter(_val(r, "管理番号") for r in rows)
    for r in rows:
        b = _val(r, "管理番号")
        if not b:
            out.append(("（空）", "L09", "管理番号", "管理番号が空"))
        elif ban[b] > 1:
            out.append((b, "L09", "管理番号", "同じ管理番号が%d件ある" % ban[b]))
        if meibo is not None and b and b not in meibo:
            out.append((b, "L10", "管理番号", "対象者名簿にない"))

    for r in rows:
        b = _val(r, "管理番号") or "（空）"

        # L01 白票
        if hissu and not any(_answered(r, e) for e in hissu):
            out.append((b, "L01", "―", "必須設問がすべて無回答"))
            continue                      # 白票は以降の判定にかけない

        for e in ent:
            v = _val(r, e["変数"])
            # L02 必須の無回答
            #     分岐の子は、親の条件を満たすときだけ必須とみなす。
            #     （満たさないのに答えている件は L07 で挙げる）
            if e["区分"] == "必須" and not _answered(r, e):
                iru = True
                for oya, jouken in ko.get(e["変数"], []):
                    eo = byvar.get(oya)
                    if eo is None or _mitasu(r, eo, jouken) is not True:
                        iru = False
                if iru:
                    out.append((b, "L02", e["番号"], "無回答"))
            # L04 単一回答に複数の印
            if e["形式"] == "単一" and v and any(k in v for k in KUGIRI):
                out.append((b, "L04", e["番号"], "値が「%s」" % v))
            # L03 選択肢の範囲外（選択肢が1つ以下なら判定を省く＝C02）
            elif e["形式"] == "単一" and v and len(e["選択肢"]) >= 2:
                ok = {str(n) for n, _l in e["選択肢"]}
                if v not in ok:
                    out.append((b, "L03", e["番号"],
                                "値「%s」は選択肢（%s）にない"
                                % (v, "・".join(sorted(ok, key=len)[:8]))))
            # L05 複数回答の上限超え
            if e["形式"].startswith("複数") and e["変数"] in LIMIT:
                sel = _multi_selected(r, e)
                if len(sel) > LIMIT[e["変数"]]:
                    out.append((b, "L05", e["番号"],
                                "%d個選ばれている（上限%d）"
                                % (len(sel), LIMIT[e["変数"]])))
            # L08 数値の範囲外（1問が複数の列に分かれることがある）
            if e["形式"] == "数値":
                for col in e["列"]:
                    if col not in RANGE:
                        continue
                    cv = _val(r, col)
                    if not cv:
                        continue
                    lo, hi = RANGE[col]
                    try:
                        x = float(cv)
                    except ValueError:
                        out.append((b, "L08", e["番号"],
                                    "%s が数値でない「%s」" % (col, cv)))
                        continue
                    if not (lo <= x <= hi):
                        out.append((b, "L08", e["番号"],
                                    "%s の値%sが範囲（%d〜%d）の外"
                                    % (col, cv, lo, hi)))

        # L06・L07 分岐
        for h, oya, jouken, kov, _konkyo in BRANCH:
            if not (h == hyo or (hyo == "在宅" and h.startswith("在宅"))):
                continue
            eo, ek = byvar.get(oya), byvar.get(kov)
            if eo is None or ek is None:
                continue
            mitasu = _mitasu(r, eo, jouken)
            if mitasu is None:              # 親が無回答。どちらとも言えない
                continue
            kotae = _answered(r, ek)
            if mitasu and not kotae:
                out.append((b, "L06", ek["番号"],
                            "%s に当たるのに無回答" % eo["番号"]))
            if (not mitasu) and kotae:
                out.append((b, "L07", ek["番号"],
                            "%s に当たらないのに回答がある" % eo["番号"]))

        # L12 直線回答
        rn = [_val(r, e["変数"]) for e in tan]
        run = best = 1
        for i in range(1, len(rn)):
            run = run + 1 if (rn[i] and rn[i] == rn[i - 1]) else 1
            best = max(best, run)
        if best >= CHOKUSEN:
            out.append((b, "L12", "―", "同じ選択肢が%d問続いている" % best))
    return out


# ── 自己試験 ────────────────────────────────────
#    点検は、欠陥を入れたときに必ず鳴り、きれいなデータでは鳴らないこと
#    を確かめてから使う。--自己試験 で走る。
def _kirei(kb, hyo):
    """コードブックどおりに正しく答えた1件を作る（鳴る項目がないこと）"""
    ent = [e for e in kb
           if (e["票"] == hyo or (hyo == "在宅" and e["票"].startswith("在宅")))
           and e["列"]]
    # 分岐の親は、その親につく条件すべてを満たす値を置く
    oyane = collections.defaultdict(list)
    for _h, oya, jouken, _ko, _k in BRANCH:
        oyane[oya].append(set(jouken))
    oyaval = {}
    for oya, jlist in oyane.items():
        kyou = set.intersection(*jlist) if jlist else set()
        oyaval[oya] = sorted(kyou)[0] if kyou else sorted(jlist[0])[0]

    r = {"管理番号": "試0001"}
    for i, e in enumerate(ent):
        opts = [n for n, _l in e["選択肢"]]
        if e["変数"] in oyaval:
            r[e["変数"]] = str(oyaval[e["変数"]])
            continue
        if e["形式"].startswith("複数"):
            # 上限のある設問は1つだけ選ぶ
            for n, _l in e["選択肢"][:1]:
                r["%s_c%s" % (e["変数"], n)] = "1"
        elif e["形式"] == "数値":
            for col in e["列"]:
                lo, hi = RANGE.get(col, (1, 2))
                r[col] = str(int((lo + hi) / 2))
        elif opts:
            # 同じ値が続くと L12 が鳴るので、順に値をずらす
            r[e["変数"]] = str(opts[i % len(opts)])
        else:
            r[e["変数"]] = "1"             # 選択肢を拾えていない設問（C02）
    return r, ent


def selftest():
    """欠陥を入れて鳴ること・きれいな件では鳴らないことを確かめる"""
    kb = codebook()
    shippai = []

    def miru(namae, rows, hyo, matsu, meibo=None):
        hit = check(rows, hyo, kb, meibo)
        deta = {h[1] for h in hit}
        if matsu:
            ok = matsu in deta
            kekka = "鳴った" if ok else "✗ 鳴らない"
        else:
            ok = not deta
            kekka = "鳴らない（正しい）" if ok else "✗ 鳴ってしまう %s" % sorted(deta)
        print("   %-4s %-34s %s" % (matsu or "（なし）", namae, kekka))
        if not ok:
            shippai.append("%s（%s）：%s" % (namae, matsu or "なし", sorted(deta)))
        return hit

    print("■ 自己試験　論理チェック")
    for hyo in ("ニーズ", "在宅"):
        base, ent = _kirei(kb, hyo)
        miru("%s　正しく答えた1件" % hyo, [dict(base)], hyo, None)

    base, ent = _kirei(kb, "ニーズ")
    hissu = [e for e in ent if e["区分"] == "必須"]
    tan4 = [e for e in ent
            if e["形式"] == "単一" and len(e["選択肢"]) >= CHOKUSEN_MIN]

    r = dict(base)                                    # L01 白票
    for e in hissu:
        for c in e["列"]:
            r.pop(c, None)
    miru("必須をすべて空にする", [r], "ニーズ", "L01")

    r = dict(base)                                    # L02 必須の無回答
    e = [x for x in hissu if x["形式"] == "単一"][0]
    r[e["変数"]] = ""
    miru("必須1問（%s）を空にする" % e["番号"], [r], "ニーズ", "L02")

    r = dict(base)                                    # L03 選択肢の範囲外
    e = tan4[0]
    r[e["変数"]] = "99"
    miru("%s に選択肢外の99を入れる" % e["番号"], [r], "ニーズ", "L03")

    r = dict(base)                                    # L04 複数の印
    r[tan4[1]["変数"]] = "1,2"
    miru("%s に「1,2」を入れる" % tan4[1]["番号"], [r], "ニーズ", "L04")

    r = dict(base)                                    # L05 上限超え
    kekka = None
    for v, lim in LIMIT.items():
        ez = [x for x in kb if x["変数"] == v and x["列"]]
        if not ez:
            continue
        ez = ez[-1]
        zb, _ = _kirei(kb, "在宅")
        for n, _l in ez["選択肢"][:lim + 1]:
            zb["%s_c%s" % (v, n)] = "1"
        kekka = miru("%s を上限+1個選ぶ" % ez["番号"], [zb], "在宅", "L05")
        break
    if kekka is None:
        shippai.append("L05 を試せる設問がない")

    r = dict(base)                                    # L06 答えるべきなのに無回答
    h, oya, jouken, kov, _k = BRANCH[0]
    ek = [x for x in kb if x["変数"] == kov][0]
    for c in ek["列"]:
        r.pop(c, None)
    miru("%s に当たるのに子を空にする" % oya, [r], "ニーズ", "L06")

    r = dict(base)                                    # L07 答えるべきでないのに回答
    sotogai = [x for x in range(1, 9) if x not in set(jouken)]
    r[oya] = str(sotogai[0])
    miru("%s を条件外にして子に回答を残す" % oya, [r], "ニーズ", "L07")

    r = dict(base)                                    # L08 数値の範囲外
    col = sorted(RANGE)[0]
    r[col] = str(RANGE[col][1] + 50)
    miru("%s を範囲の上を超える値にする" % col, [r], "ニーズ", "L08")

    r = dict(base)                                    # L09 管理番号
    miru("同じ管理番号を2件にする", [dict(base), r], "ニーズ", "L09")
    r = dict(base)
    r["管理番号"] = ""
    miru("管理番号を空にする", [r], "ニーズ", "L09")

    r = dict(base)                                    # L10 名簿にない
    miru("名簿にない管理番号", [r], "ニーズ", "L10", meibo={"試9999"})

    r = dict(base)                                    # L12 直線回答
    for e in tan4[:CHOKUSEN]:
        r[e["変数"]] = "1"
    miru("4択以上の%d問を同じ値にする" % CHOKUSEN, [r], "ニーズ", "L12")

    print("\n■ 自己試験　コードブックの点検")

    def cbmiru(namae, kb2, matsu):
        deta = {h[0] for h in codebook_check(kb2)}
        ok = matsu in deta
        print("   %-4s %-34s %s" % (matsu, namae, "鳴った" if ok else "✗ 鳴らない"))
        if not ok:
            shippai.append("%s（%s）：%s" % (namae, matsu, sorted(deta)))

    def copy_kb():
        return json.loads(json.dumps(kb, ensure_ascii=False))

    k2 = copy_kb()                                    # C01 変数の重複
    k2.append(json.loads(json.dumps(k2[0], ensure_ascii=False)))
    cbmiru("同じ変数の項目を足す", k2, "C01")

    k2 = copy_kb()                                    # C02 選択肢が足りない
    for e in k2:
        if e["形式"] == "単一" and e["列"]:
            e["選択肢"] = e["選択肢"][:1]
            break
    cbmiru("単一回答の選択肢を1個に削る", k2, "C02")

    k2 = copy_kb()                                    # C03 列と選択肢が合わない
    for e in k2:
        if e["形式"].startswith("複数"):
            e["列"] = e["列"][:-1]
            break
    cbmiru("複数回答の列を1つ減らす", k2, "C03")

    k2 = [e for e in copy_kb() if e["変数"] != BRANCH[0][1]]   # C04 分岐の親を消す
    cbmiru("分岐の親の項目を消す", k2, "C04")

    k2 = copy_kb()                                    # C05 範囲の列を消す
    col = sorted(RANGE)[0]
    for e in k2:
        if col in e["列"]:
            e["列"] = [c for c in e["列"] if c != col]
    cbmiru("RANGE に書いた列を消す", k2, "C05")

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
    ap.add_argument("--暫定", action="store_true", dest="zantei",
                    help="コードブックの版が合わなくても動かす")
    ap.add_argument("--自己試験", action="store_true", dest="jiko",
                    help="欠陥を入れて点検が鳴ることを確かめる")
    a = ap.parse_args()

    if a.jiko:
        return selftest()

    chigai = ban_chigai()
    if chigai:
        print("■ コードブックの版が合いません")
        print("   %s。" % chigai)
        print("   このまま走らせると、鳴るべきものが鳴らず、")
        print("   鳴らなくてよいものが鳴ります。")
        print("   校了版の調査票を受領してコードブックを作り直してください"
              "（doc34 §4）。")
        if not a.zantei:
            print("\n   承知のうえで動かすときは --暫定 を付けてください。")
            return 2
        print("   【--暫定 が付いているため、暫定の結果として続けます】\n")

    kb = codebook()

    # ── コードブック自体の点検 ───────────────────────
    cbhit = codebook_check(kb)
    if cbhit:
        print("■ コードブックの点検　%d件" % len(cbhit))
        cc = collections.Counter(h[0] for h in cbhit)
        for kigou, nm, _itsu, naoshi in CB_RULES:
            if cc.get(kigou):
                print("   %s %-24s %3d件　→ %s" % (kigou, nm, cc[kigou], naoshi))
        for kigou, tgt, naiyou in cbhit:
            print("      %s %s：%s" % (kigou, tgt, naiyou))
        print("   ※ ここが鳴っている間、下の結果はデータの誤りとは言えません。\n")
    else:
        print("■ コードブックの点検　鳴った項目はありません\n")

    zen = []
    for hyo, path in (("ニーズ", a.needs), ("在宅", a.zaitaku)):
        if not os.path.exists(path):
            print("  %s：%s がない" % (hyo, path))
            continue
        rows = list(csv.DictReader(io.open(path, encoding="utf-8-sig")))
        hit = check(rows, hyo, kb)
        zen += [(hyo,) + h for h in hit]
        print("■ %s　%d件を見て、%d件が鳴りました" % (hyo, len(rows), len(hit)))
        cnt = collections.Counter(h[1] for h in hit)
        for kigou, nm, _itsu, _naoshi in RULES:
            if cnt.get(kigou):
                print("   %s %-22s %4d件" % (kigou, nm, cnt[kigou]))
        if not hit:
            print("   鳴った項目はありません")

    if a.out:
        with io.open(a.out, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["票", "管理番号", "記号", "設問", "内容"])
            w.writerows(zen)
        print("\n書き出し：%s（%d件）" % (a.out, len(zen)))
    if chigai or cbhit:
        riyuu = []
        if chigai:
            riyuu.append("コードブックの版が合っていません")
        if cbhit:
            riyuu.append("コードブックの点検が%d件鳴っています" % len(cbhit))
        print("\n※ 上の結果は暫定です。%s。" % "、".join(riyuu))
    return 0


if __name__ == "__main__":
    sys.exit(main())
