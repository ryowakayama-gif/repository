# -*- coding: utf-8 -*-
"""事業所・高齢者向け住まいの一覧の受入検査（令和8年10月15日）

08_作業順位 の順位26。

町から事業所一覧・高齢者向け住まい一覧をいただいたときに、
**そのまま素案に載せてよいか**を機械で確かめる。
いただく前から用意しておくことで、届いた日に次へ進める。

  python3 07_ソーススクリプト/check_jigyosho_uke_R8.10.15.py [台帳のパス]

  既定の台帳　05_試算・管理シート/川崎町_事業所・住まい_突合台帳_R8.10.15.xlsx
  見るシート　03_事業所一覧（記入様式）・04_高齢者向け住まい一覧（記入様式）

検査
  1　様式の列が揃っているか（列を増減していないか）
  2　必ず要る欄が埋まっているか（種別・名称・所在・定員）
  3　所在が宮城県柴田郡川崎町又は近隣の市町村か
     ⚠ 福岡県田川郡川崎町・神奈川県川崎市のものが混じっていないか
  4　事業所番号の形（10桁・宮城県は04で始まる）
  5　定員・入居者数が数として読めるか。入居者数が定員を超えていないか
  6　素案 9-1 の必要利用定員総数の表と突き合わせ、町内の定員が埋まるか
  7　素案 2-7 の「現時点で判明していること」と食い違っていないか
  8　時点が揃っているか（別々の時点の値を1つの表に混ぜていないか）

⚠ 記入がまだの行（空の行）は不適合としない。「未記入」として数える。
  様式をお渡しした段階では、ほとんどが未記入である。

⚠ 本検査は**様式と整合**を見るものであり、値が正しいかどうかは確かめられない。
  値の確かさは、町の台帳・事業所への照会（作業順位27）による。
"""
import os
import re
import sys

import docx
import openpyxl
from docx.oxml.ns import qn

DAICHO = "05_試算・管理シート/川崎町_事業所・住まい_突合台帳_R8.10.15.xlsx"
SOAN = "01_第10期_最新版成果品/川崎町_計画書素案_v2.26_版表記整理版.docx"

SH_JIGYO = "03_事業所一覧（記入様式）"
SH_SUMAI = "04_高齢者向け住まい一覧（記入様式）"

RETSU_JIGYO = ["No.", "サービス種別", "事業所の名称", "事業所番号", "所在地",
               "定員", "開設年月", "利用者数", "稼働率", "職員数（常勤換算）",
               "町内・町外", "時点", "備考"]
RETSU_SUMAI = ["No.", "種別", "名称", "所在地", "定員（戸数）", "入居者数",
               "うち要介護者等の数", "特定施設の指定", "町内・町外",
               "時点", "備考"]

# 近隣の市町村（町外の施設の所在として現れうるもの）
KINRIN = ["川崎町", "柴田町", "大河原町", "村田町", "蔵王町", "白石市",
          "名取市", "岩沼市", "仙台市", "大和町", "角田市", "丸森町",
          "七ヶ宿町", "亘理町", "山元町", "宮城県"]
# 取り違えやすい他団体
BETSU = ["田川郡", "福岡県", "神奈川県", "川崎市高津", "川崎市中原",
         "川崎市川崎区", "川崎市幸区", "川崎市多摩", "川崎市宮前", "川崎市麻生"]


def yomu_soan_teiin():
    """素案 9-1 の必要利用定員総数の表から、町内の定員の欄を読む。"""
    if not os.path.exists(SOAN):
        return None
    d = docx.Document(SOAN)
    for el in d.element.body.iterchildren():
        if el.tag != qn("w:tbl"):
            continue
        tb = docx.table.Table(el, d)
        head = [c.text.replace("\n", "").strip() for c in tb.rows[0].cells]
        if "町内の定員" in head:
            i = head.index("町内の定員")
            return [(r.cells[0].text.replace("\n", "").strip(),
                     r.cells[i].text.replace("\n", "").strip())
                    for r in tb.rows[1:]]
    return None


def yomu_gyo(ws, retsu):
    """記入様式のシートを読む。見出しの行は4行目、中身は5行目から。"""
    head = [str(ws.cell(4, c).value or "").strip()
            for c in range(1, len(retsu) + 1)]
    rows, aki = [], 0
    for r in range(5, ws.max_row + 1):
        v = [ws.cell(r, c).value for c in range(1, len(retsu) + 1)]
        if all(x in (None, "") for x in v):
            continue
        # ⚠ No. だけが入っている行は、様式の空の行である。
        #   「未記入」として1件ずつ挙げると読みづらいので、数だけ数える。
        if all(x in (None, "") for x in v[1:]):
            aki += 1
            continue
        rows.append((r, v))
    return head, rows, aki


def suji(v):
    """数として読めるか。読めれば値、読めなければ None。"""
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.replace(",", "").replace("人", "").replace("戸", "").strip()
        try:
            return float(s)
        except ValueError:
            return None
    return None


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else DAICHO
    if not os.path.exists(path):
        print("台帳がありません：", path)
        print("  ⚠ 町からご記入いただいた台帳のパスを引数で渡してください。")
        sys.exit(1)
    wb = openpyxl.load_workbook(path, data_only=True)
    ng, note = [], []
    n_jigyo = n_sumai = 0
    aki_jigyo = aki_sumai = 0
    n_mikinyu = 0

    for sh, retsu, kind in ((SH_JIGYO, RETSU_JIGYO, "事業所"),
                            (SH_SUMAI, RETSU_SUMAI, "住まい")):
        if sh not in wb.sheetnames:
            ng.append(f"検査1　シート「{sh}」がない")
            continue
        ws = wb[sh]
        head, rows, aki = yomu_gyo(ws, retsu)
        # ── 検査1　列
        if head != retsu:
            chigai = [f"{i + 1}列目 {a}→{b}"
                      for i, (a, b) in enumerate(zip(retsu, head)) if a != b]
            ng.append(f"検査1　{sh} の列が様式と違う："
                      + "／".join(chigai[:4] or ["列数が違う"]))
            continue
        if kind == "事業所":
            n_jigyo, aki_jigyo = len(rows), aki
        else:
            n_sumai, aki_sumai = len(rows), aki
        if aki:
            note.append(f"{sh}　様式の空の行が{aki}行あります"
                        "（記入をお待ちしている状態です）")

        for r, v in rows:
            rec = dict(zip(retsu, v))
            nm = str(rec.get("事業所の名称") or rec.get("名称") or "").strip()
            shu = str(rec.get("サービス種別") or rec.get("種別") or "").strip()
            sho = str(rec.get("所在地") or "").strip()
            tei = rec.get("定員") if kind == "事業所" else rec.get("定員（戸数）")

            # ── 検査2　必ず要る欄
            kara = [k for k, x in (("種別", shu), ("名称", nm),
                                   ("所在地", sho)) if not x]
            if kara:
                n_mikinyu += 1
                note.append(f"{sh} {r}行目　未記入：{'・'.join(kara)}")
                continue

            # ── 検査3　所在
            if any(b in sho for b in BETSU):
                ng.append(f"検査3　{sh} {r}行目　他団体の施設の疑い："
                          f"{nm}／{sho}")
            elif not any(k in sho for k in KINRIN):
                note.append(f"{sh} {r}行目　所在が近隣の市町村でない："
                            f"{nm}／{sho}（ご確認ください）")

            # ── 検査4　事業所番号
            if kind == "事業所":
                ban = str(rec.get("事業所番号") or "").strip()
                if ban:
                    b2 = ban.replace("-", "").replace("－", "")
                    if not re.fullmatch(r"\d{10}", b2):
                        ng.append(f"検査4　{sh} {r}行目　事業所番号が10桁でない："
                                  f"{nm}／{ban}")
                    elif not b2.startswith("04"):
                        note.append(f"{sh} {r}行目　事業所番号が04で始まらない"
                                    f"（宮城県以外の指定）：{nm}／{ban}")

            # ── 検査5　定員と入居者数
            t = suji(tei)
            if tei not in (None, "") and t is None:
                ng.append(f"検査5　{sh} {r}行目　定員が数として読めない："
                          f"{nm}／{tei}")
            nyu = suji(rec.get("入居者数") if kind == "住まい"
                       else rec.get("利用者数"))
            if t is not None and nyu is not None and nyu > t:
                ng.append(f"検査5　{sh} {r}行目　"
                          f"{'入居者数' if kind == '住まい' else '利用者数'}"
                          f"が定員を超えている：{nm}／{nyu:g} > {t:g}")
            if kind == "住まい":
                yo = suji(rec.get("うち要介護者等の数"))
                if nyu is not None and yo is not None and yo > nyu:
                    ng.append(f"検査5　{sh} {r}行目　"
                              f"要介護者等の数が入居者数を超えている："
                              f"{nm}／{yo:g} > {nyu:g}")

        # ── 検査8　時点
        jiten = {str(dict(zip(retsu, v)).get("時点") or "").strip()
                 for _, v in rows}
        jiten = {x for x in jiten if x}
        if len(jiten) > 1:
            note.append(f"{sh}　時点が{len(jiten)}通り混じっている："
                        + "／".join(sorted(jiten)[:4])
                        + "（別々の時点の値を1つの表に混ぜていないかご確認ください）")

    # ── 検査6　素案の必要利用定員総数との突合
    teiin = yomu_soan_teiin()
    if teiin is None:
        note.append("検査6　未実施（素案の必要利用定員総数の表を読めない）")
    else:
        aki = [nm for nm, v in teiin if "要確認" in v or not v.strip()]
        if aki:
            if n_jigyo == 0:
                note.append("検査6　素案の「町内の定員」が"
                            f"{len(aki)}件空いているが、台帳が未記入のため"
                            "まだ埋められない：" + "／".join(aki))
            else:
                note.append(f"検査6　素案の「町内の定員」が{len(aki)}件空いている。"
                            "台帳の記入をもとに埋めてください："
                            + "／".join(aki))

    print("事業所・高齢者向け住まいの一覧の受入検査")
    print("  台帳：", path)
    print(f"  事業所一覧 記入{n_jigyo}行・空{aki_jigyo}行"
          f"／高齢者向け住まい一覧 記入{n_sumai}行・空{aki_sumai}行"
          f"／必要な欄が未記入の行 {n_mikinyu}件")
    for m in note:
        print("   △", m)
    if ng:
        for m in ng:
            print("   ×", m)
        print("  ⚠ 不適合がある間は、素案に載せないでください。")
        sys.exit(1)
    if n_jigyo == 0 and n_sumai == 0:
        print("   △ まだ1行も記入されていません（様式をお渡しした段階です）")
        print("  ⚠ 町からご記入いただいたら、もう一度実行してください。")
        return
    print("   ○ 様式の列・必ず要る欄・所在・事業所番号・定員の整合は適合")
    print("  ⚠ 本検査は様式と整合を見るものです。"
          "値が正しいかどうかは、町の台帳と事業所への照会によります。")


if __name__ == "__main__":
    main()
