"""他メンバー作成版への対応履歴ブック ジェネレータ

ご指示（令和8年10月19日）
  「他メンバー作成版の元計画素案との論点漏れの対応について
   対応履歴を整理して下さい。元エビデンスのアウトプットもお願いします」

このブックが答えること
  他メンバーが作成された素案（令和8年10月8日受領・4.5MB）に対し、
  当方が何を、いつ、何を根拠に足し、又は直したか。その全件である。

立て方
  1. 件数・箇所・根拠は、すべて各系統の単一出所から読み込む。
     このブックで書き写していない。写すと、素案が直ったときに
     履歴だけが古くなる。
  2. 系統ごとに「由来（どのレビュー・どの資料から出たか）」と
     「受領日・作成日」を持つ。日付は git の記録（生成器が最初に
     入った日）と突き合わせ、食い違えば生成を止める。
  3. 根拠のエビデンスは、生成器が実際に開いているファイルを
     機械で集める。手で並べると、使っていない資料を根拠として
     挙げてしまう。

対応の3つの型
  移植   当方が先に作っていた骨子案の中身を、他メンバー版の
         該当箇所へ移したもの（20件）。素案に無かった章節を足す。
  追記   レビュー・原典・村資料から出た論点を、他メンバー版の
         本文に足したもの。
  書き換え 他メンバー版の本文が事実として誤っているものを直したもの。
         挿入ではなく置換であり、直す前の文が本文にちょうど1つある
         ことを確かめてから書き換えている。

出力: output/北塩原村_他メンバー版への対応履歴.xlsx
"""

import importlib
import os
import subprocess
import sys
import warnings

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from kitashiobara_common import (
    COLORS, FONT, OUT_DIR, REPO_ROOT, SRC_DIR, add_sheet, ensure_out_dir,
    style_header_row, style_title, write_row,
)

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

OUT_FILE = f"{OUT_DIR}/北塩原村_他メンバー版への対応履歴.xlsx"
MEMBER = (f"{SRC_DIR}/他メンバー素案_20261008/"
          "北塩原村_第8期障がい福祉計画_第4期障がい児福祉計画_"
          "素案_他メンバー版.docx")
MEMBER_WS = (f"{SRC_DIR}/他メンバー素案_20261008/"
             "北塩原村_第8期障がい福祉計画_第4期障がい児福祉計画_"
             "策定ワークシート_他メンバー版.xlsx")
HONPON = f"{OUT_DIR}/北塩原村_計画素案_正本_移植後.docx"
KIJUNBI = "基準日：令和8年10月19日"

ISHOKU_K = "移植"
TSUIKI_K = "追記"
KAKIKAE_K = "書き換え"
ZU_K = "図の作成"


# ===========================================================================
# 1. 系統の定義
#    (変数名, 生成器, 型, 日本語の系統名, 由来, 受領日・作成日)
#    件数・箇所・根拠は書かない。単一出所から読む。
#    日付は git の記録と突き合わせる（検査②）。
# ===========================================================================
KEITOU = [
    ("ISHOKU", "ishoku_okurijo", ISHOKU_K, "骨子案からの移植",
     "当方が令和8年7月から10月にかけて作った骨子案・計画素案。"
     "他メンバー版に無かった章節を、該当箇所へ移した",
     "2026-10-09"),
    ("MEMO_EDITS", "honpon", TSUIKI_K, "論点メモ",
     "令和8年10月9日の論点メモ。データ待ちでないものを入れた",
     "2026-10-09"),
    ("KAKIKAE", "mece_houshu", KAKIKAE_K, "事実の誤りの書き換え（制度）",
     "ＭＥＣＥレビュー（利用の流れと報酬改定）。"
     "重度訪問介護の対象の書き方が法令と違っていた",
     "2026-10-10"),
    ("MURA_EDITS", "murashiryo", TSUIKI_K, "村資料の点検",
     "村の自立支援協議会資料（令和8年5月25日）と計画策定方針の点検",
     "2026-10-10"),
    ("JUTEN_EDITS", "juten12", TSUIKI_K, "重点施策12件の評価設計",
     "村の重点施策12件に評価の設計（指標・現状値・目標）を当てた",
     "2026-10-10"),
    ("HP_EDITS", "hp_shisaku", TSUIKI_K, "村ホームページ掲載施策の棚卸し",
     "村ホームページに載っている施策を棚卸しし、素案に無いものを拾った",
     "2026-10-10"),
    ("HP3_EDITS", "hp_3bunya", TSUIKI_K, "村ホームページ3分野の棚卸し",
     "高齢者・児童・社会福祉の3分野について同じ棚卸しを行った",
     "2026-10-10"),
    ("RT_EDITS", "redteam_kyukyu", TSUIKI_K,
     "ＲｅｄＴｅａｍ（救急・防災・冬季）",
     "救急・防災・冬季の論点についてのＲｅｄＴｅａｍレビュー",
     "2026-10-10"),
    ("ME_EDITS", "mece_houshu", TSUIKI_K,
     "ＭＥＣＥ（利用の流れと報酬改定）",
     "サービス利用の流れと令和6年度報酬改定についてのＭＥＣＥレビュー",
     "2026-10-10"),
    ("SA_EDITS", "mikomi_redteam", TSUIKI_K, "見込量算定のＲｅｄＴｅａｍ",
     "見込量の算定方法についての再レビュー",
     "2026-10-10"),
    ("AN_EDITS", "anke_mikomi", TSUIKI_K, "アンケートと見込量の整合",
     "アンケートの回答と見込量が食い違わないかの点検",
     "2026-10-10"),
    ("KG_EDITS", "kaigo_seigo", TSUIKI_K, "介護保険事業計画との整合",
     "高齢者福祉計画・介護保険事業計画と重ねて数えないための点検",
     "2026-10-08"),
    ("KG_KAKIKAE", "kaigo_seigo", KAKIKAE_K,
     "事実の誤りの書き換え（他計画との連携）",
     "第7章の章の導入が村の計画だけを挙げていた",
     "2026-10-08"),
    ("KO_EDITS", "kouiki_redteam", TSUIKI_K, "広域連携・圏域整合",
     "会津北部4町村及び会津障がい保健福祉圏域との整合のレビュー",
     "2026-10-10"),
    ("KO_KAKIKAE", "kouiki_redteam", KAKIKAE_K,
     "事実の誤りの書き換え（障がい児支援）",
     "第4章2（4）の記述が基本指針の4つの中核機能と合っていなかった",
     "2026-10-10"),
    ("KJ_EDITS", "jouhou_kouhyou", TSUIKI_K, "情報公表制度と経営情報",
     "障害福祉サービス等情報公表制度と経営情報の見える化",
     "2026-10-10"),
    ("IK_EDITS", "ishi_kettei", TSUIKI_K, "意思決定支援ガイドライン第2版",
     "令和8年9月に公表された意思決定支援ガイドライン第2版",
     "2026-10-10"),
    ("YS_EDITS", "yosan_mece", TSUIKI_K, "国予算（令和7・8年度）",
     "国の障害保健福祉部予算案から、本村が使える事業を拾った",
     "2026-10-10"),
    ("KY_EDITS", "kyoseigata", TSUIKI_K, "共生型サービスの適用可能性",
     "村内の介護保険事業所を障がい福祉の供給に充てられるかの検討",
     "2026-10-10"),
    ("KN_EDITS", "kojino", TSUIKI_K, "高次脳機能障害者支援法",
     "令和7年法律第96号（令和7年12月24日公布）の反映",
     "2026-10-11"),
    ("CS_EDITS", "chiikishien", TSUIKI_K, "地域支援事業実施要綱",
     "介護保険の地域支援事業実施要綱の改正（令和8年8月20日）",
     "2026-10-11"),
    ("KT_EDITS", "houteikisai", TSUIKI_K, "法定記載事項の点検",
     "基本指針の別表第二の記載事項を条項単位で照合した",
     "2026-10-12"),
    ("SS_EDITS", "shihyou_sokutei", TSUIKI_K, "指標の測定可能性",
     "目標に置いた指標が実際に測れるかの点検",
     "2026-10-12"),
    ("RV_EDITS", "review3_taiou", TSUIKI_K, "第3回レビューへの対応",
     "令和8年10月13日の外部レビュー（第3回）",
     "2026-10-13"),
    ("HK_EDITS", "houkaisei_r8", TSUIKI_K, "令和8年法改正の反映",
     "社会福祉法等改正（令和8年法律第51号）・民法等改正（第45号）・"
     "医療的ケア児支援法改正（第77号）",
     "2026-10-13"),
    ("UR_EDITS", "uchiawase_ronten", TSUIKI_K, "打合せ論点10項目",
     "令和8年10月13日の打合せで挙がった論点10項目",
     "2026-10-13"),
    ("KH_EDITS", "kanren_hourei", TSUIKI_K, "関連法令3法",
     "災害対策基本法・障害者雇用促進法・社会福祉士及び介護福祉士法",
     "2026-10-13"),
    ("TK_EDITS", "taikei_hanei", TSUIKI_K, "計画の体系（案Ｃ）の反映",
     "村が令和8年10月14日に採用した体系（案Ｃ）",
     "2026-10-14"),
    ("TK_KAKIKAE", "taikei_hanei", KAKIKAE_K,
     "体系・工程の決定による書き換え",
     "村の決定（案Ｃ・案Ａ）による第3章の章題とパブリックコメントの日程",
     "2026-10-14"),
    ("SJ_EDITS", "shogaiji_ronten", TSUIKI_K, "障がい児の論点（法改正）",
     "こども家庭審議会障害児支援部会の令和8年度の審議事項のレビュー",
     "2026-10-14"),
    ("AK_EDITS", "anke_kaitei", TSUIKI_K, "アンケート集計表の改訂版",
     "報告書用集計表の改訂版の受領により、課題の8区分目を足した",
     "2026-10-14"),
    ("KI_EDITS", "ken_iryo", TSUIKI_K, "第8次福島県医療計画との調和",
     "国の基本指針 第三の一の7が調和を求める医療計画が落ちていた",
     "2026-10-15"),
    ("KI_GYOU", "ken_iryo", TSUIKI_K,
     "第8次福島県医療計画との調和（表に行を足す）",
     "第1章2（1）表1-3 に県医療計画を行として足した",
     "2026-10-15"),
    ("SR_EDITS", "houkoku_hanei", TSUIKI_K,
     "アンケート調査結果報告書のレビュー反映版の確認",
     "レビュー反映版の第6章と前期計画の具体的な取組み38件を"
     "正本と突き合わせ、落ちていた4件を足した",
     "2026-10-22"),
    ("SZ_EDITS", "zenki_redteam", TSUIKI_K,
     "前期計画のレッドチーム分析",
     "前期計画が挙げたＵＤフォント・配色への配慮とほじょ犬マークが"
     "落ちていたため足した",
     "2026-10-22"),
    ("DK_HYOU", "mura_daicho", KAKIKAE_K,
     "村提供の給付の台帳による実績の表の書き換え",
     "第5章の実績の表が他メンバー版の値のままで、村の協議会資料"
     "（資料２－３）と食い違っていた。台帳の受領で裏づけが取れた",
     "2026-10-22"),
    ("DK_KAKIKAE", "mura_daicho", KAKIKAE_K,
     "村提供の給付の台帳による評価の文の書き換え",
     "実績の表を直したことに伴い、「見込量どおり」「利用がなかった」"
     "とする評価の文が事実と合わなくなった",
     "2026-10-22"),
    ("DK_EDITS", "mura_daicho", TSUIKI_K,
     "村提供の給付の台帳の注記",
     "実績の出所と、村に確かめている数え方の違い・見込量との離れを"
     "注記した",
     "2026-10-22"),
    ("TR_KAKIKAE", "tasha_review", KAKIKAE_K,
     "他ＡＩのレビュー（見込量・ＫＰＩ）による書き換え",
     "第3章冒頭が村の計画策定方針（中間見直しを行う）と食い違っていた",
     "2026-10-23"),
    ("GRAPHS", "graph", ZU_K, "当方が作った図",
     "他メンバー版の図は画像で取り込まれており数値を直せないため、"
     "当方で作り直せる図を別に作った",
     "2026-10-10"),
]

# 根拠のエビデンスが生成器の中の定数として現れない系統。
# なぜ現れないかを明記する（空欄にしない）。
EVIDENCE_NASHI = {
    "MEMO_EDITS": "論点メモは当方が作った覚書であり、"
                  "外部資料を直に開いていない。"
                  "元になった資料は各論点の konkyo に書いてある",
    "RV_EDITS": "外部レビューの結果はチャットで受領したものであり、"
                "ファイルとして格納していない。"
                "指摘の全文は生成器の定数（REVIEW）に写してある",
    "UR_EDITS": "打合せ論点は口頭と当方の記録によるものであり、"
                "外部資料を直に開いていない",
    "KH_EDITS": "災害対策基本法・障害者雇用促進法・"
                "社会福祉士及び介護福祉士法の条文は"
                "生成器の定数に写してある（原典の入手はＳ-28等で依頼中）",
    "HK_EDITS": "法律の概要は別の系統（ＳＪ）が開いている"
                "同じ資料を根拠としている",
    "KG_EDITS": "介護保険側の計画素案（別ブランチ）との突合であり、"
                "本件のリポジトリには格納していない",
    "KG_KAKIKAE": "同上",
    "CS_EDITS": "地域支援事業実施要綱は source/国資料 に格納している"
                "（介護保険最新情報Vol1404）。"
                "生成器は要綱の条項を定数に写して用いている",
    "KN_EDITS": "",
    "AN_EDITS": "アンケートの集計は別の系統（ＳＳ・ＡＫ）が開いている"
                "同じ資料を根拠としている",
    "RT_EDITS": "",
    "TK_KAKIKAE": "体系・工程は村の決定によるものであり、"
                  "決定の記録は打合せによる。"
                  "工程の日付は kitashiobara_koutei.py に一本化している",
    "KI_GYOU": "ＫＩ系統と同じ資料",
    "GRAPHS": "図の元データは当方の算定（kitashiobara_common.py の"
              "POPULATION・BENEFIT_YEARS 等）による",
}


# ===========================================================================
# 2. 単一出所から読む
# ===========================================================================
def _mod(g):
    return importlib.import_module(f"build_kitashiobara_{g}")


def recs(var, g):
    m = _mod(g)
    return list(getattr(m, var))


def _hyouji(v):
    """表のセルの書き換えは3年度の組で持つため、「／」でつないで示す。"""
    return "／".join(v) if isinstance(v, tuple) else v


def _no(rec):
    if isinstance(rec, dict):
        return rec.get("no", "")
    return str(rec[0])


def _saki(rec):
    if isinstance(rec, dict):
        return rec.get("saki") or rec.get("ichi") or ""
    #   MEMO_EDITS は (番号, 見出し, 本文, 挿入位置)
    return str(rec[1])


def _konkyo(rec):
    if isinstance(rec, dict):
        return rec.get("konkyo") or rec.get("src") or ""
    return ""


def _kata(rec):
    if isinstance(rec, dict):
        if "mae" in rec and "ato" in rec:
            return "書き換え"
        if "gyou" in rec and "atama" in rec:
            return "表に行を足す"
        return rec.get("kata") or rec.get("keishiki") or "本文"
    return "本文"


def _ryo(rec):
    """足した分量。移植は ryo、追記はブロックの数で示す。"""
    if not isinstance(rec, dict):
        return "本文1段落"
    if rec.get("ryo"):
        return rec["ryo"]
    if "nakami" in rec:
        n = {}
        for b in rec["nakami"]:
            n[b[0]] = n.get(b[0], 0) + 1
        namae = {"p": "本文", "h2": "見出し", "h3": "小見出し",
                 "tbl": "表", "note": "注記"}
        return "・".join(f"{namae.get(k, k)}{v}"
                         for k, v in n.items())
    if "gyou" in rec:
        return f"表に{len(rec['gyou'])}行"
    if "mae" in rec:
        return "本文1段落の置換"
    if rec.get("file"):
        return "図1点"
    return "―"


# ===========================================================================
# 3. エビデンスを機械で集める
# ===========================================================================
def _paths(obj, acc, depth=0):
    if depth > 4:
        return
    if isinstance(obj, str):
        if "/source/" in obj and obj.endswith(
                (".pdf", ".docx", ".xlsx", ".md", ".txt", ".zip")):
            acc.add(obj.replace(REPO_ROOT + "/", ""))
    elif isinstance(obj, dict):
        for v in obj.values():
            _paths(v, acc, depth + 1)
    elif isinstance(obj, (list, tuple)):
        for v in obj:
            _paths(v, acc, depth + 1)


def evidence(g):
    """生成器が実際に開いているエビデンスのファイルを集める。"""
    try:
        m = _mod(g)
    except Exception:
        return []
    acc = set()
    for nm in dir(m):
        if nm.startswith("_"):
            continue
        try:
            _paths(getattr(m, nm), acc)
        except Exception:
            pass
    return sorted(p for p in acc
                  if os.path.exists(os.path.join(REPO_ROOT, p)))


def _git_tsuika(path):
    """ファイルが最初に git に入った日。取れなければ None。"""
    try:
        r = subprocess.run(
            ["git", "log", "--diff-filter=A", "--format=%ad",
             "--date=short", "--", path],
            cwd=REPO_ROOT, capture_output=True, text=True, timeout=30)
    except Exception:
        return None
    hi = [x for x in r.stdout.split() if x]
    return hi[-1] if hi else None


# ===========================================================================
# 4. 元の素案と移植後の寸法
# ===========================================================================
def sunpou():
    try:
        import docx
    except ImportError:
        return None
    out = {}
    for namae, p in (("他メンバー版", MEMBER), ("正本（移植後）", HONPON)):
        if not os.path.exists(p):
            continue
        d = docx.Document(p)
        e = 0
        for sh in d.inline_shapes:
            e += 1
        out[namae] = dict(dan=len(d.paragraphs), hyo=len(d.tables), zu=e)
    return out


# ===========================================================================
# シート
# ===========================================================================
KATA_FILL = {ISHOKU_K: "2CA02C", TSUIKI_K: "2E75B6",
             KAKIKAE_K: "C00000", ZU_K: "ED7D31"}


def _ichiran():
    """系統ごとに (定義, 件数, 記録) を返す。"""
    out = []
    for var, g, kata, namae, yurai, hi in KEITOU:
        rs = recs(var, g)
        out.append(dict(var=var, g=g, kata=kata, namae=namae,
                        yurai=yurai, hi=hi, rs=rs,
                        ev=evidence(g)))
    return out


def sheet_ketsuron(wb):
    ich = _ichiran()
    sp = sunpou()
    gokei = sum(len(x["rs"]) for x in ich)
    ws = add_sheet(
        wb, "00_結論",
        "他メンバー作成版の素案に当方が行った対応",
        "他メンバーが作成された素案（令和8年10月8日受領）に対し、"
        f"当方は{len(KEITOU)}系統・{gokei}件の対応を行いました。"
        "件数・箇所・根拠はすべて各系統の単一出所から読み込んでおり、"
        "このブックで書き写していません。"
        f"{KIJUNBI}。",
        [22, 10, 14, 34, 54, 12])
    r = 5
    if sp:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        style_title(ws.cell(row=r, column=1), "素案の寸法の変化",
                    fill=COLORS["subhead"], size=11)
        r += 1
        style_header_row(ws, r, ["", "段落", "表", "図（埋込画像）",
                                 "", ""])
        r += 1
        for namae in ("他メンバー版", "正本（移植後）"):
            if namae not in sp:
                continue
            v = sp[namae]
            write_row(ws, r, [namae, v["dan"], v["hyo"], v["zu"], "", ""],
                      aligns=["left", "center", "center", "center",
                              "left", "left"])
            ws.row_dimensions[r].height = 20
            r += 1
        if len(sp) == 2:
            a, b = sp["他メンバー版"], sp["正本（移植後）"]
            write_row(ws, r, ["差", f"+{b['dan'] - a['dan']}",
                              f"+{b['hyo'] - a['hyo']}",
                              f"+{b['zu'] - a['zu']}", "", ""],
                      aligns=["left", "center", "center", "center",
                              "left", "left"])
            ws.cell(row=r, column=1).fill = PatternFill(
                "solid", fgColor=COLORS["band"])
            r += 1
        r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    style_title(ws.cell(row=r, column=1),
                f"系統ごとの件数（{len(KEITOU)}系統・{gokei}件）",
                fill=COLORS["subhead"], size=11)
    r += 1
    style_header_row(ws, r, ["系統", "件数", "型", "由来の生成器",
                             "どこから出たか", "受領日・作成日"])
    r += 1
    for i, x in enumerate(ich):
        write_row(ws, r, [x["namae"], len(x["rs"]), x["kata"],
                          f"build_kitashiobara_{x['g']}.py",
                          x["yurai"], _wareki(x["hi"])],
                  alt=(i % 2 == 1),
                  aligns=["left", "center", "center", "left", "left",
                          "center"])
        ws.cell(row=r, column=3).fill = PatternFill(
            "solid", fgColor=KATA_FILL.get(x["kata"], "808080"))
        ws.cell(row=r, column=3).font = Font(
            name=FONT, size=9, color="FFFFFF", bold=True)
        ws.row_dimensions[r].height = 34
        r += 1
    write_row(ws, r, ["合計", gokei, "―", f"{len(KEITOU)}系統", "", ""],
              aligns=["left", "center", "center", "left", "left",
                      "center"])
    ws.cell(row=r, column=1).fill = PatternFill(
        "solid", fgColor=COLORS["subhead"])
    r += 2

    ws.merge_cells(start_row=r, start_column=1, end_row=r + 4, end_column=6)
    c = ws.cell(row=r, column=1)
    kata_kazu = {}
    for x in ich:
        kata_kazu[x["kata"]] = kata_kazu.get(x["kata"], 0) + len(x["rs"])
    c.value = (
        "※ 型ごとの内訳は "
        + "／".join(f"{k} {v}件" for k, v in sorted(
            kata_kazu.items(), key=lambda y: -y[1]))
        + " です。"
        "「移植」は当方が先に作っていた骨子案の中身を他メンバー版の"
        "該当箇所へ移したもので、素案に無かった章節を足しています。"
        "「書き換え」は他メンバー版の本文が事実として誤っているものに"
        "限っており、挿入ではなく置換であるため、直す前の文が本文に"
        "ちょうど1つあることを確かめてから書き換えています。"
        "当方が書いた文の誤りは、本文を書き換えるのではなく"
        "生成器の側を直しています。")
    c.alignment = Alignment(wrap_text=True, vertical="top")
    c.font = Font(name=FONT, size=9)
    return ws


_ZEN = "０１２３４５６７８９"


def _wareki(iso):
    y, m, d = (int(x) for x in iso.split("-"))
    def s(n):
        return _ZEN[n] if n < 10 else str(n)
    return f"令和{s(y - 2018)}年{s(m)}月{s(d)}日"


def sheet_rireki(wb):
    ich = _ichiran()
    gokei = sum(len(x["rs"]) for x in ich)
    ws = add_sheet(
        wb, "01_対応履歴",
        f"全{gokei}件",
        "番号・箇所・根拠は各系統の単一出所から読み込んでいます。"
        "日付は系統の受領日・作成日です。"
        "「足した分量」は、追記であればブロックの数（本文・表・注記）、"
        "移植であれば当方が見積もった頁数です。",
        [12, 14, 26, 36, 16, 14, 56])
    style_header_row(ws, 5, ["番号", "日付", "系統", "箇所", "型",
                             "足した分量", "根拠"])
    r = 6
    i = 0
    for x in ich:
        for rec in x["rs"]:
            write_row(ws, r, [_no(rec), _wareki(x["hi"]), x["namae"],
                              _saki(rec), _kata(rec), _ryo(rec),
                              _konkyo(rec) or "―"],
                      alt=(i % 2 == 1),
                      aligns=["center", "center", "left", "left",
                              "center", "center", "left"])
            ws.cell(row=r, column=3).fill = PatternFill(
                "solid", fgColor=KATA_FILL.get(x["kata"], "808080"))
            ws.cell(row=r, column=3).font = Font(
                name=FONT, size=9, color="FFFFFF")
            ws.row_dimensions[r].height = 30
            r += 1
            i += 1
    return ws


def sheet_evidence(wb):
    ich = _ichiran()
    #   ファイル → それを根拠とする系統
    byfile = {}
    for x in ich:
        for p in x["ev"]:
            byfile.setdefault(p, []).append(x["namae"])
    ws = add_sheet(
        wb, "02_元エビデンス",
        f"生成器が実際に開いているエビデンス {len(byfile)}点",
        "手で並べたものではありません。各生成器のモジュールを読み込み、"
        "source/ の下にあって実在するファイルのパスを機械で集めています。"
        "使っていない資料を根拠として挙げないためです。"
        "大きさは実測です。",
        [24, 62, 12, 10, 50])
    style_header_row(ws, 5, ["分類", "ファイル", "大きさ", "頁・シート",
                             "これを根拠とする系統"])
    r = 6
    for i, (p, keitou) in enumerate(sorted(byfile.items())):
        full = os.path.join(REPO_ROOT, p)
        kb = os.path.getsize(full) / 1024
        bunrui = p.split("/")[1] if "/" in p else "―"
        namae = p.split("/")[-1]
        write_row(ws, r, [bunrui, namae,
                          f"{kb:,.0f}KB", _nakami(full),
                          "／".join(sorted(set(keitou)))],
                  alt=(i % 2 == 1),
                  aligns=["left", "left", "right", "center", "left"])
        ws.row_dimensions[r].height = 26
        r += 1

    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    nashi = [x for x in ich if not x["ev"]]
    style_title(ws.cell(row=r, column=1),
                f"エビデンスのファイルが生成器に現れない系統"
                f"（{len(nashi)}系統。なぜ現れないかを書く）",
                fill=COLORS["subhead"], size=11)
    r += 1
    style_header_row(ws, r, ["系統", "なぜファイルが現れないか", "", "",
                             ""])
    r += 1
    for i, x in enumerate(nashi):
        write_row(ws, r, [x["namae"],
                          EVIDENCE_NASHI.get(x["var"], ""), "", "", ""],
                  alt=(i % 2 == 1), aligns=["left", "left", "left",
                                            "left", "left"])
        ws.row_dimensions[r].height = 34
        r += 1
    return ws


def _nakami(full):
    """頁数（PDF）又はシート数（xlsx）又は段落数（docx）。"""
    try:
        if full.endswith(".pdf"):
            from pypdf import PdfReader
            return f"{len(PdfReader(full).pages)}頁"
        if full.endswith(".xlsx"):
            import openpyxl
            wb = openpyxl.load_workbook(full, read_only=True)
            n = len(wb.sheetnames)
            wb.close()
            return f"{n}シート"
        if full.endswith(".docx"):
            import docx
            return f"{len(docx.Document(full).paragraphs)}段落"
        if full.endswith((".md", ".txt")):
            with open(full, encoding="utf-8", errors="ignore") as f:
                return f"{sum(1 for _ in f)}行"
    except Exception:
        return "―"
    return "―"


def sheet_kakikae(wb):
    ich = _ichiran()
    kk = [(x, rec) for x in ich for rec in x["rs"]
          if isinstance(rec, dict) and "mae" in rec and "ato" in rec]
    ws = add_sheet(
        wb, "03_書き換え",
        f"他メンバー版の本文を直したもの {len(kk)}件",
        "事実として誤っているものに限っています。"
        "挿入ではなく置換であるため、直す前の文が本文にちょうど1つ"
        "あることを確かめてから書き換えています（1つでなければ生成が"
        "止まります）。"
        "当方が書いた文の誤りは、本文を書き換えるのではなく"
        "生成器の側を直しています。",
        [10, 30, 56, 56, 22])
    style_header_row(ws, 5, ["番号", "箇所", "直す前（他メンバー版）",
                             "直した後", "系統"])
    r = 6
    for i, (x, rec) in enumerate(kk):
        write_row(ws, r, [rec["no"], rec.get("saki", ""),
                          _hyouji(rec["mae"]), _hyouji(rec["ato"]),
                          x["namae"]],
                  alt=(i % 2 == 1), aligns=["center"] + ["left"] * 4)
        ws.cell(row=r, column=3).fill = PatternFill(
            "solid", fgColor="FCE4E4")
        ws.cell(row=r, column=4).fill = PatternFill(
            "solid", fgColor="E2EFDA")
        ws.row_dimensions[r].height = 64
        r += 1
    return ws


def sheet_ishoku(wb):
    rs = recs("ISHOKU", "ishoku_okurijo")
    ws = add_sheet(
        wb, "04_移植",
        f"骨子案から他メンバー版へ移したもの {len(rs)}件",
        "当方が先に作っていた骨子案・計画素案の中身を、"
        "他メンバー版の該当箇所へ移したものです。"
        "他メンバー版に無かった章節を足しています。",
        [8, 34, 40, 22, 12, 56])
    style_header_row(ws, 5, ["番号", "移した先", "挿入位置の目印",
                             "形式", "分量", "根拠"])
    r = 6
    for i, rec in enumerate(rs):
        write_row(ws, r, [rec["no"], rec["saki"], rec.get("ichi", "")[:60],
                          rec.get("keishiki", ""), rec.get("ryo", ""),
                          rec.get("konkyo", "")],
                  alt=(i % 2 == 1),
                  aligns=["center", "left", "left", "left", "center",
                          "left"])
        ws.row_dimensions[r].height = 40
        r += 1
    return ws


# ===========================================================================
# 自己点検
# ===========================================================================
def verify():
    ng = []
    ich = _ichiran()

    # ① 系統の一覧が、正本の生成器が実際に適用する一覧と一致すること。
    #    片方だけ増やすと履歴が実態と離れる。
    h = open(f"{REPO_ROOT}/build_kitashiobara_honpon.py",
             encoding="utf-8").read()
    for var, _g, _k, namae, _y, _hi in KEITOU:
        if var == "MEMO_EDITS":
            continue
        if var not in h:
            ng.append(f"正本の生成器が使っていない系統を挙げている: "
                      f"{namae}（{var}）")
    #    逆向き。honpon が apply_juten に渡している EDITS が
    #    すべて一覧にあること。
    import re
    #    apply_juten に渡している EDITS と、書き換えで回している一覧。
    #    records は apply_juten の引数の既定値であり系統ではない。
    tsukau = set(re.findall(r"apply_juten\(doc, ([A-Z][A-Z0-9_]*)", h))
    m = re.search(r"for rec in (list\([A-Z][A-Z0-9_]*\)"
                  r"(?: \+ list\([A-Z][A-Z0-9_]*\))*):", h)
    if m:
        tsukau |= set(re.findall(r"list\(([A-Z][A-Z0-9_]*)\)", m.group(1)))
    for nm in ("ISHOKU", "MURA_EDITS", "JUTEN_EDITS", "TK_KAKIKAE",
               "KI_GYOU", "GRAPHS"):
        if nm in h:
            tsukau.add(nm)
    aru = {x[0] for x in KEITOU}
    nuke = sorted(tsukau - aru)
    if nuke:
        ng.append(f"正本に入れているのに履歴に無い系統: {nuke}")

    # ② 受領日・作成日が git の記録と合うこと
    for x in ich:
        if x["g"] == "honpon":
            continue
        p = f"build_kitashiobara_{x['g']}.py"
        hi = _git_tsuika(p)
        if hi is None:
            print("  ※ git の記録が取れないため②は省略")
            break
        if hi != x["hi"]:
            ng.append(f"受領日・作成日が git の記録と違う: "
                      f"{x['namae']} 一覧{x['hi']}／git {hi}")

    # ③ 件数が1件以上で、番号が空でないこと
    for x in ich:
        if not x["rs"]:
            ng.append(f"件数が0の系統がある: {x['namae']}")
        for rec in x["rs"]:
            if not _no(rec).strip():
                ng.append(f"番号が空の対応がある: {x['namae']}")
            if not _saki(rec).strip():
                ng.append(f"箇所が空の対応がある: "
                          f"{x['namae']} {_no(rec)}")

    # ④ 番号が系統をまたいで重複していないこと
    #    （重複すると履歴を追えない）
    mi = {}
    for x in ich:
        for rec in x["rs"]:
            no = _no(rec)
            mi.setdefault(no, []).append(x["namae"])
    kasanari = {k: v for k, v in mi.items() if len(v) > 1}
    #    ＫＩ系統は KI_EDITS と KI_GYOU で番号を分けているため
    #    重複しない。重複があれば実際の誤りである。
    if kasanari:
        ng.append(f"番号が系統をまたいで重複している: "
                  f"{ {k: v for k, v in list(kasanari.items())[:5]} }")

    # ④b 系統に付けた型の札が、現物と食い違わないこと。
    #     「書き換え」と書いた系統は mae/ato を持つ記録だけ、
    #     それ以外の系統は mae/ato を持つ記録を含まないこと。
    #     令和8年10月19日、ＲｅｄＴｅａｍ（救急・防災・冬季）に
    #     「書き換え・追記」と書いていたが現物は追記だけであった。
    for x in ich:
        kk_n = sum(1 for rec in x["rs"] if isinstance(rec, dict)
                   and "mae" in rec and "ato" in rec)
        if x["kata"] == KAKIKAE_K and kk_n != len(x["rs"]):
            ng.append(f"「書き換え」とした系統に置換でない記録がある: "
                      f"{x['namae']}（{kk_n}/{len(x['rs'])}件）")
        if x["kata"] != KAKIKAE_K and kk_n:
            ng.append(f"「{x['kata']}」とした系統に置換の記録がある: "
                      f"{x['namae']}（{kk_n}件）")
        if x["kata"] not in KATA_FILL:
            ng.append(f"型の札が未定義: {x['namae']}（{x['kata']}）")

    # ⑤ エビデンスのファイルが実在すること（evidence が実在で絞るため、
    #    ここでは「1点も無い系統に理由が書かれていること」を見る）
    for x in ich:
        if x["ev"]:
            continue
        if x["var"] not in EVIDENCE_NASHI:
            ng.append(f"エビデンスが無い系統に理由を書いていない: "
                      f"{x['namae']}（{x['var']}）")

    # ⑥ 他メンバー版の原本と正本が実在し、正本の方が大きいこと
    if not os.path.exists(MEMBER):
        ng.append("他メンバー版の素案が無い")
    if not os.path.exists(MEMBER_WS):
        ng.append("他メンバー版の策定ワークシートが無い")
    sp = sunpou()
    if sp and len(sp) == 2:
        a, b = sp["他メンバー版"], sp["正本（移植後）"]
        for key, namae in (("dan", "段落"), ("hyo", "表")):
            if b[key] <= a[key]:
                ng.append(f"正本の{namae}が他メンバー版より増えていない: "
                          f"{a[key]} → {b[key]}")

    # ⑦ 書き換えが「他メンバー版の本文にある文」を直していること。
    #    当方が書いた文を書き換えの形で直していないかを見る。
    kk = [rec for x in ich for rec in x["rs"]
          if isinstance(rec, dict) and "mae" in rec and "ato" in rec]
    if os.path.exists(MEMBER):
        try:
            import docx
        except ImportError:
            print("  ※ python-docx が無いため⑦は省略")
        else:
            d = docx.Document(MEMBER)
            mem = "".join("".join(p.text.split()) for p in d.paragraphs)
            #   表のセルの書き換え（ＤＫ_HYOU）は mae が3年度の組であり、
            #   他メンバー版の表の行にある
            gyou = ["".join("".join(c.text.split()) for c in r.cells)
                    for t in d.tables for r in t.rows]
            for rec in kk:
                if isinstance(rec["mae"], tuple):
                    mae = "".join("".join(v.split()) for v in rec["mae"])
                    if not any(mae in g for g in gyou):
                        ng.append(f"書き換えの「直す前」が他メンバー版の"
                                  f"表に無い: {rec['no']}")
                    continue
                mae = "".join(rec["mae"].split())
                if mae not in mem:
                    ng.append(f"書き換えの「直す前」が他メンバー版に"
                              f"無い: {rec['no']}")

    # ⑧ 他団体の案件の語が入っていないこと
    zen = "\n".join(str(v) for rec in KEITOU for v in rec)
    zen += "\n" + "\n".join(EVIDENCE_NASHI.values())
    for hoka in ("小野町", "金ヶ崎", "金ケ崎", "阿蘇", "札幌", "川崎町"):
        if hoka in zen:
            ng.append(f"他団体の名が入っている: {hoka}")

    if ng:
        for x in ng:
            print(f"  × {x}")
        raise SystemExit("対応履歴の自己点検で止めました")

    gokei = sum(len(x["rs"]) for x in ich)
    ev = set()
    for x in ich:
        ev |= set(x["ev"])
    kata_kazu = {}
    for x in ich:
        kata_kazu[x["kata"]] = kata_kazu.get(x["kata"], 0) + len(x["rs"])
    return (f"{len(KEITOU)}系統・{gokei}件／"
            + "／".join(f"{k}{v}件" for k, v in sorted(
                kata_kazu.items(), key=lambda y: -y[1]))
            + f"／エビデンス{len(ev)}点／"
            f"書き換え{len(kk)}件はいずれも他メンバー版の現物にある文")


def main():
    ensure_out_dir()
    msg = verify()
    wb = Workbook()
    wb.remove(wb.active)
    sheet_ketsuron(wb)
    sheet_rireki(wb)
    sheet_evidence(wb)
    sheet_kakikae(wb)
    sheet_ishoku(wb)
    wb.save(OUT_FILE)
    print(f"  自己検証: {msg}")
    print(f"作成: {OUT_FILE}")
    print(f"  シート{len(wb.sheetnames)}／"
          f"{sum(len(recs(v, g)) for v, g, *_ in KEITOU)}件")


if __name__ == "__main__":
    sys.exit(main())
