# -*- coding: utf-8 -*-
"""町内事業所の実態（公表システム本体からの採録）の**数値の正本**
（令和8年10月20日）

  ⚠ **この値を別の場所に書き写さない。** 試算シート・計画素案・
    ご報告のメモはいずれもここから引く（CLAUDE.md 第3節）。

入力　09_元資料/その他データ/公表システム_詳細_川崎町_R8.10.20.csv
      （fetch_kouhyou_detail_R8.10.20.py が原典のページから起こしたもの）

  単体で実行すると、読み取った値を表にして示し、検査1〜7を行う。
    python3 07_ソーススクリプト/data_jigyosho_jittai_R8.10.20.py
"""
import collections
import csv
import os
import re
import sys

CSV = "09_元資料/その他データ/公表システム_詳細_川崎町_R8.10.20.csv"

# 指定権者　⚠ 居宅介護支援は平成30年4月1日から市町村が指定権者である
MACHI_SHITEI = {"320", "430", "540"}        # 町が指定
KEN_SHITEI = {"110", "140", "150", "160", "210", "220", "510", "520"}

# 入所・入居の定員と在所者数を比べてよいサービス
#   ⚠ 通所（150・160）の利用定員は1日当たりであり、
#     短期入所生活介護（210）は空床利用型であるため、入所率を算定しない。
NYUSHO = {"510", "520", "540", "320"}

# 職種別の従業者の表の値は6つ（専従常勤・兼務常勤・専従非常勤・兼務非常勤・
# 合計・常勤換算）。⚠ 資格の表（延べ人数・常勤・非常勤）と形で区別する。
SHOKUSHU = ("医師", "薬剤師", "生活相談員", "支援相談員", "相談援助員",
            "看護職員", "介護職員", "管理者", "計画作成担当者",
            "管理栄養士", "栄養士", "機能訓練指導員", "介護支援専門員",
            "訪問介護員等", "理学療法士", "作業療法士", "言語聴覚士",
            "歯科衛生士", "調理員", "事務員", "その他の従業者")


def kazu(s):
    """「30人」「18.3人」から数を取る。複数あれば並びで返す。"""
    return [float(x) for x in re.findall(r"(-?[\d.]+)\s*人", s)]


def hitotsume(s):
    v = kazu(s)
    return v[0] if v else None


def yomu(csv_path=CSV):
    """(事業所番号, サービスコード) ごとの実態を返す。"""
    if not os.path.exists(csv_path):
        raise SystemExit(f"入力がありません：{csv_path}\n"
                         "先に fetch_kouhyou_detail_R8.10.20.py を"
                         "実行してください。")
    nama = collections.defaultdict(list)
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            nama[(r["事業所番号"], r["サービスコード"])].append(r)

    out = {}
    for key, rs in nama.items():
        ban, sv = key
        d = {"事業所番号": ban, "サービスコード": sv,
             "サービスの種類": rs[0]["サービスの種類"],
             "事業所名": re.sub(r"\s+", "", rs[0]["事業所名"]),
             "記入日": rs[0]["記入日"], "公表日": rs[0]["公表日"],
             "指定権者": ("町" if sv in MACHI_SHITEI else
                      "県" if sv in KEN_SHITEI else "［要確認］")}
        hiku = {}
        for r in rs:
            hiku.setdefault(r["項目"], []).append(r["値"])

        def v1(*na):
            """項目の名が na のいずれかで始まる行の、最初の**値のある**もの。

            ⚠ **値のない行（表の見出し）を飛ばす。** 通所介護の
              「利用定員」は見出しの行に値がなく、その下の
              「利用定員 ▸ 指定通所介護事業所」に30人が入る。
              見出しを拾うと定員が落ちる（現にそれで落とした）。
            """
            for n in na:
                for k, vs in hiku.items():
                    if k == n or k.startswith(n):
                        for v in vs:
                            if v.strip():
                                return v
            return ""

        d["法人の名称"] = re.sub(r"\s+", "", v1("法人等の名称", "法人の名称"))
        d["事業所の形態"] = v1("事業所の形態")
        d["指定の年月日"] = v1("指定の年月日")
        d["指定の更新年月日"] = v1("指定の更新年月日")

        # ── 定員
        teiin_moto = v1("入所定員", "入居定員", "利用定員")
        v = kazu(teiin_moto)
        d["定員の原文"] = teiin_moto
        d["定員"] = v[-1] if v else None
        yu = re.search(r"(\d+)ユニット", teiin_moto)
        d["ユニット数"] = int(yu.group(1)) if yu else None

        # ── 在所者（入所者・入居者・利用者）
        danjo = v1("入所者の男女別人数", "入居者の男女別人数",
                   "利用者の男女別人数")
        dj = kazu(danjo)
        d["男性"] = dj[0] if len(dj) >= 2 else None
        d["女性"] = dj[1] if len(dj) >= 2 else None
        if len(dj) >= 2:
            d["在所者数"] = dj[0] + dj[1]
        else:
            # 介護老人保健施設は男女別を置かず、要介護別の合計のみ
            v = kazu(v1("介護サービスの入所者への提供実績"))
            d["在所者数"] = v[-1] if v else None
        m = re.search(r"([\d.]+)歳", v1("入所者の平均年齢",
                                       "入居者の平均年齢",
                                       "利用者の平均年齢"))
        d["平均年齢"] = float(m.group(1)) if m else None
        m = re.search(r"([\d,.]+)日", v1("入所者の平均的な入所日数",
                                        "利用者の平均的な利用日数"))
        d["平均在所日数"] = (float(m.group(1).replace(",", ""))
                        if m else None)

        # ── 入所率　⚠ 入所・入居のサービスに限る
        if (sv in NYUSHO and d["定員"] and d["在所者数"] is not None):
            d["入所率"] = d["在所者数"] / d["定員"] * 100
        else:
            d["入所率"] = None

        # ── 待機者数　⚠ 定義は「入所申込者の数－（入所定員－入所者）」
        machi = ""
        for k, vs in hiku.items():
            if k.startswith("待機者数") and "その人数" in k:
                machi = vs[0]
        d["待機者数"] = hitotsume(machi)
        d["待機者数の有無"] = ("あり" if any(
            k.startswith("待機者数") and vs[0] == "あり"
            for k, vs in hiku.items()) else
            ("なし" if any(k.startswith("待機者数") for k in hiku)
             else "項目なし"))

        # ── 職種別の従業者（実人数の合計と常勤換算）
        shoku = []
        for na in SHOKUSHU:
            for k, vs in hiku.items():
                if k != na:
                    continue
                for s in vs:
                    v = kazu(s)
                    if len(v) == 6:
                        shoku.append((na, *v))
                        break
                break
        d["職種別"] = shoku
        d["実人数の合計"] = sum(x[5] for x in shoku) if shoku else None
        d["常勤換算の合計"] = sum(x[6] for x in shoku) if shoku else None

        # ── 前年度の採用者数・退職者数
        #   ⚠ 職種を3つずつ区切った表が複数あり、項目の名が同じである。
        #     すべてを合算する（職種別の内訳は本表では持たない）。
        d["採用者数"] = sum(sum(kazu(s)) for s in hiku.get("前年度の採用者数",
                                                      []))
        d["退職者数"] = sum(sum(kazu(s)) for s in hiku.get("前年度の退職者数",
                                                      []))
        d["採用退職の区分数"] = len(hiku.get("前年度の採用者数", []))

        # ── 退所先　⚠ 「入所前の居所」は公表の項目にない
        #   ⚠ **項目の名だけで探すと誤る。**「その他」は職種の
        #     「その他の従業者」と紛れ、「医療機関」は協力医療機関と紛れる。
        #     **退所者の表の見出しを見つけ、その後に続く行だけを読む。**
        #   ⚠ **期間が施設の種類で違う。** 特別養護老人ホームは前年度、
        #     介護老人保健施設は記入日前月から直近3か月である。
        TAISHO_NA = ("自宅等", "介護保険施設",
                     "特別養護老人ホーム以外の社会福祉施設",
                     "医療機関", "死亡者", "その他")
        #   ⚠ **一致は完全一致で見る。** `startswith` で見ると
        #     「その他」が「その他の従業者」（常勤換算の値）に当たる。
        #     現にそれで誤った表を作った（令和8年10月20日）。
        taisho, kikan = {}, ""
        for r in rs:
            k, v = r["項目"], r["値"]
            if ("退所者の人数" in k or "退居した者の人数" in k) and not kikan:
                m = re.search(r"（([^）]*)）", k)
                kikan = m.group(1) if m else k
            if k in TAISHO_NA:
                n = kazu(v)
                if len(n) >= 2:
                    taisho[k] = n[-1]
        # ⚠ 6つの行がそろっていないものは退所先の表ではない
        #   （居宅介護支援の別の表の「その他」に当たる）
        if len(taisho) < 5:
            taisho, kikan = {}, ""
        d["退所先"] = taisho
        d["退所先の期間"] = kikan

        # ── 夜勤・1人当たり
        m = re.search(r"([\d.]+)人",
                      v1("看護職員及び介護職員１人当たりの入所者数",
                         "看護職員及び介護職員１人当たりの利用者数"))
        d["職員1人当たりの入所者数"] = float(m.group(1)) if m else None

        out[key] = d
    return out


# ════════════════════════════════════════ 単体で動かしたときの表示と検査
def main():
    j = yomu()
    futeki = []

    def judge(ok, mi, setsu=""):
        print(("   ○ " if ok else "   × ") + mi + (("　" + setsu)
                                                   if setsu else ""))
        if not ok:
            futeki.append(mi)

    print("町内事業所の実態（公表システム本体）　令和8年10月20日")
    print()
    print("  ── 入所・入居系の定員と在所の状況")
    print(f"     {'サービス':<22}{'事業所':<20}{'定員':>6}{'在所':>6}"
          f"{'入所率':>8}{'待機':>6}  記入日")
    nyu = [d for d in j.values() if d["サービスコード"] in NYUSHO]
    for d in sorted(nyu, key=lambda x: x["サービスコード"]):
        ritsu = "%.1f%%" % d["入所率"] if d["入所率"] else "―"
        machi = "%.0f" % d["待機者数"] if d["待機者数"] else "―"
        print(f"     {d['サービスの種類'][:20]:<22}{d['事業所名'][:18]:<20}"
              f"{d['定員'] or 0:>6.0f}{d['在所者数'] or 0:>6.0f}"
              f"{ritsu:>8}{machi:>6}  {d['記入日']}")
    print()
    print("  ── 職員の状況（常勤換算・前年度の採用と退職）")
    for d in sorted(j.values(), key=lambda x: x["サービスコード"]):
        print(f"     {d['サービスの種類'][:20]:<22}{d['事業所名'][:16]:<18}"
              f"実人数 {d['実人数の合計'] or 0:>5.1f}　"
              f"常勤換算 {d['常勤換算の合計'] or 0:>5.1f}　"
              f"採用 {d['採用者数']:>3.0f}　退職 {d['退職者数']:>3.0f}")
    print()

    # ══════ 検査1　14件そろっているか
    print("  ── 検査1　14の指定がそろっているか")
    judge(len(j) == 14, f"14件（実際 {len(j)}件）")

    # ══════ 検査2　入所率が100％を超えないか
    print("  ── 検査2　入所率が100％を超えないか")
    koe = [f"{d['事業所名']}（{d['サービスの種類']}）{d['入所率']:.1f}％"
           for d in j.values() if d["入所率"] and d["入所率"] > 100.5]
    judge(not koe, "在所者数が定員を超えていない", "／".join(koe))

    # ══════ 検査3　男女別の和と要介護別の合計が一致するか
    print("  ── 検査3　男女別の和と要介護別の合計が一致するか")
    chigai = []
    for d in j.values():
        if d["男性"] is None:
            continue
        # 要介護別の合計（最終列の和）を数え直す
        pass
    judge(True, "⚠ この検査は未実施（要介護別の列の対応付けが要る）")
    print("        未実施の理由：要介護別×年齢別の表の列の対応付けを"
          "本表では持っていないため。")

    # ══════ 検査4　認知症対応型共同生活介護の定員
    print("  ── 検査4　認知症対応型共同生活介護の定員")
    gh = [d for d in j.values() if d["サービスコード"] == "320"]
    teiin = sum(d["定員"] or 0 for d in gh)
    judge(len(gh) == 2 and teiin == 27,
          f"2事業所・定員計27人（実際 {len(gh)}事業所・{teiin:.0f}人）",
          "／".join(f"{d['事業所名']} {d['定員の原文']}" for d in gh))

    # ══════ 検査5　指定権者の別
    print("  ── 検査5　指定権者の別（町か県か）")
    machi = [d for d in j.values() if d["指定権者"] == "町"]
    judge(all(d["指定権者"] in ("町", "県") for d in j.values()),
          f"すべてに指定権者がある（町 {len(machi)}件・"
          f"県 {len(j) - len(machi)}件）")

    # ══════ 検査6　記入日が取れているか
    print("  ── 検査6　いつ時点の値か（記入日）")
    nashi = [d["事業所名"] for d in j.values() if not d["記入日"]]
    judge(not nashi, "すべてに記入日がある", "／".join(nashi))
    hi = sorted(d["記入日"] for d in j.values() if d["記入日"])
    print(f"        記入日の幅　{hi[0]}〜{hi[-1]}")

    # ══════ 検査7　⚠ 一度やった取り違えが戻っていないか
    print("  ── 検査7　⚠ 一度やった取り違えが戻っていないか")
    tsu = [d for d in j.values() if d["サービスコード"] == "150"]
    judge(bool(tsu) and tsu[0]["定員"] == 30,
          "通所介護の利用定員が30人である（見出しの空行を拾っていない）",
          f"実際 {tsu[0]['定員'] if tsu else None}")
    hasu = [f"{d['事業所名']} {k}={v}"
            for d in j.values() for k, v in d["退所先"].items()
            if v != int(v)]
    judge(not hasu,
          "退所先の値が整数である（「その他の従業者」の常勤換算を"
          "拾っていない）", "／".join(hasu[:3]))

    # ══════ 検査8　⚠ 埋まらないものを数で残す
    print("  ── 検査8　⚠ 公表システムに項目がないもの（未実施）")
    mi = []
    for d in j.values():
        if d["サービスコード"] in ("540",) and d["待機者数の有無"] == "項目なし":
            mi.append(f"{d['事業所名']}（{d['サービスの種類']}）："
                      "待機者数の項目がない")
    mi.append("入所前の居所（在宅か医療機関か）：公表の項目にない")
    mi.append("稼働率（延べ利用日数÷延べ定員日数）：公表の項目にない")
    mi.append("休止床の数：公表の項目にない")
    print(f"     △ {len(mi)}件（「適合」に丸めていません）")
    for s in mi:
        print(f"        {s}")

    print()
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        for m in futeki:
            print(f"      {m}")
        sys.exit(1)
    print(f"   ○ 検査1〜8　不適合0件"
          f"（⚠ 未実施 {len(mi) + 1}件は別に数えています）")


if __name__ == "__main__":
    main()
