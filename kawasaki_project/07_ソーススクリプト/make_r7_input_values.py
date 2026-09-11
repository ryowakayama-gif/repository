# -*- coding: utf-8 -*-
"""見える化システムの「令和8年度の実績（見込み）値」に貼り付ける令和7年度の値を作る。

令和8年度は令和8年5月の1か月分であるため、これを完結年度（令和7年度）の
月平均に差し替える。出所はすべて介護保険事業状況報告（年報・令和7年度）で、
総括表の令和7年度の人数・回数と小数第4位まで一致することを確かめてある。

  認定者数　　　　　　　　　… 様式1の5（12）要介護(要支援)認定者数
  施設・居住系サービス利用者数… 様式1の7(16)・(18)、様式1の6(15) ÷12
  在宅サービス利用者数　　　… 様式1の7(16)・(18)、様式2（件数）÷12
  在宅サービス利用回（日）数 … 様式1の7(17)・(19) ÷12

使い方
  python3 make_r7_input_values.py [出力txt]
"""
import sys
import warnings
from decimal import Decimal, ROUND_HALF_UP

import openpyxl

warnings.filterwarnings("ignore")

NEN = ("/home/user/repository/kawasaki_project/09_元資料/R7実績データ/"
       "年報データ_2025_川崎町.xlsx")
SOK = ("/home/user/repository/kawasaki_project/09_元資料/R8実績データ/R8.9.11受領版/"
       "【川崎町】第10期_将来推計総括表_R8.9.11出力.xlsx")

CARE = ["要支援1", "要支援2", "要介護1", "要介護2", "要介護3", "要介護4", "要介護5"]
wb = openpyxl.load_workbook(NEN, data_only=True)


def row7(ws, r, c_s1, c_s2, c_k1):
    """要支援1・2 と 要介護1〜5 の7つを取り出す"""
    out = [ws.cell(r, c_s1).value or 0, ws.cell(r, c_s2).value or 0]
    out += [ws.cell(r, c_k1 + i).value or 0 for i in range(5)]
    return [float(v) for v in out]


def r1(x):
    return float(Decimal(str(x)).quantize(Decimal("0.001"), ROUND_HALF_UP))


def to_int(vals, group=((0, 1), (2, 3, 4, 5, 6))):
    """小数の月平均を整数に丸める。要支援群・要介護群それぞれで
    合計を四捨五入し、最大剰余法で配分して合計を保つ。"""
    out = [0] * len(vals)
    for g in group:
        s = sum(vals[i] for i in g)
        tgt = int(Decimal(str(s)).quantize(Decimal("1"), ROUND_HALF_UP))
        base = [int(vals[i]) for i in g]
        rem = sorted(g, key=lambda i: -(vals[i] - int(vals[i])))
        need = tgt - sum(base)
        for j, i in enumerate(g):
            out[i] = base[j]
        k = 0
        while need > 0 and k < len(rem) * 4:
            out[rem[k % len(rem)]] += 1
            need -= 1
            k += 1
        while need < 0:
            for i in sorted(g, key=lambda i: (vals[i] - int(vals[i]))):
                if out[i] > 0 and need < 0:
                    out[i] -= 1
                    need += 1
    return out


# ================================================ 1. 認定者数（様式1の5）
ws = wb["様式１の５ 総数"]
NIN_ROWS = ["65〜69歳", "70〜74歳", "75〜79歳", "80〜84歳", "85〜89歳",
            "90歳以上", "第2号被保険者"]
nin = {}
for sex, r0 in (("男", 12), ("女", 22)):
    rows = []
    for i in range(6):                     # 年齢階級6区分
        rows.append(row7(ws, r0 + i, 5, 6, 9))
    rows.append(row7(ws, r0 + 6, 5, 6, 9))  # 第2号被保険者
    nin[sex] = rows

# ================================================ 2〜3. 利用者数
w16 = wb["様式１の７（１６）居宅介護"]
w18 = wb["様式１の７（１８）地域密着型"]
w17 = wb["様式１の７（１７）居宅介護"]
w19 = wb["様式１の７（１９）地域密着型（２０）施設介護"]
w16r = {str(w16.cell(r, 3).value or "").strip(): r for r in range(12, 27)}
w18r = {str(w18.cell(r, 3).value or "").strip().replace("\n", ""): r
        for r in range(12, 25)}
w18r = {k.replace("　", ""): v for k, v in w18r.items()}
w17r = {str(w17.cell(r, 3).value or "").strip(): r for r in range(12, 23)}
w19r = {str(w19.cell(r, 3).value or "").strip(): r for r in range(12, 18)}
w2 = wb["様式２（件数）"]
w2r = {}
for r in range(11, 53):
    nm = None
    for c in (3, 4, 5):
        v = w2.cell(r, c).value
        if v and str(v).strip():
            nm = str(v).strip()
    if nm:
        w2r.setdefault(nm, r)
w16f = wb["様式１の６"]


def g16(name):
    return row7(w16, w16r[name], 4, 5, 8)


def g18(name):
    return row7(w18, w18r[name], 4, 5, 8)


def g2(name):
    return row7(w2, w2r[name], 6, 7, 10)


def add(*a):
    return [sum(x) for x in zip(*a)]


def div12(v):
    return [x / 12 for x in v]


# 施設サービス（様式1の6(15)：要介護1〜5のみ）
def gfac(r):
    return [0.0, 0.0] + [float(w16f.cell(r, 7 + i).value or 0) for i in range(5)]


# 画面上の入力欄マスク（1＝白の入力欄、0＝グレーの「―」）
MASK_ALL = (1, 1, 1, 1, 1, 1, 1)
MASK_K = (0, 0, 1, 1, 1, 1, 1)      # 要支援は「―」
MASK_S2 = (0, 1, 1, 1, 1, 1, 1)     # 要支援1のみ「―」

SHISETSU = [
    ("特定施設入居者生活介護",
     div12(add(g16("特定施設入居者生活介護（短期利用以外）"),
               g16("特定施設入居者生活介護（短期利用）"))), MASK_ALL),
    ("認知症対応型共同生活介護",
     div12(add(g18("認知症対応型共同生活介護（短期利用以外）"),
               g18("認知症対応型共同生活介護（短期利用）"))), MASK_S2),
    ("地域密着型特定施設入居者生活介護",
     div12(add(g18("地域密着型特定施設入居者生活介護（短期利用以外）"),
               g18("地域密着型特定施設入居者生活介護（短期利用）"))), MASK_K),
    ("地域密着型介護老人福祉施設入所者生活介護",
     div12(g18("地域密着型介護老人福祉施設入所者生活介護")), MASK_K),
    ("介護老人福祉施設", div12(gfac(25)), MASK_K),
    ("介護老人保健施設", div12(gfac(28)), MASK_K),
    ("介護医療院", div12(gfac(34)), MASK_K),
]

ZAITAKU = [
    ("訪問介護", div12(g16("訪問介護")), MASK_K),
    ("訪問入浴介護", div12(g16("訪問入浴介護")), MASK_ALL),
    ("訪問看護", div12(g16("訪問看護")), MASK_ALL),
    ("訪問リハビリテーション", div12(g16("訪問リハビリテーション")), MASK_ALL),
    ("居宅療養管理指導", div12(g16("居宅療養管理指導")), MASK_ALL),
    ("通所介護", div12(g16("通所介護")), MASK_K),
    ("通所リハビリテーション", div12(g16("通所リハビリテーション")), MASK_ALL),
    ("短期入所生活介護", div12(g16("短期入所生活介護")), MASK_ALL),
    ("短期入所療養介護（老健）",
     div12(g16("短期入所療養介護（介護老人保健施設）")), MASK_ALL),
    ("短期入所療養介護（病院等）", div12(g16("短期入所療養介護（病院等）")), MASK_ALL),
    ("短期入所療養介護（介護医療院）",
     div12(g16("短期入所療養介護（介護医療院）")), MASK_ALL),
    ("福祉用具貸与", div12(g16("福祉用具貸与")), MASK_ALL),
    ("特定福祉用具購入費", div12(g2("特定福祉用具販売")), MASK_ALL),
    ("住宅改修費", div12(g2("住宅改修")), MASK_ALL),
    ("介護予防支援・居宅介護支援", div12(g16("介護予防支援・居宅介護支援")), MASK_ALL),
    ("定期巡回・随時対応型訪問介護看護",
     div12(g18("定期巡回・随時対応型訪問介護看護")), MASK_K),
    ("夜間対応型訪問介護", div12(g18("夜間対応型訪問介護")), MASK_K),
    ("地域密着型通所介護", div12(g18("地域密着型通所介護")), MASK_K),
    ("認知症対応型通所介護", div12(g18("認知症対応型通所介護")), MASK_ALL),
    ("小規模多機能型居宅介護",
     div12(add(g18("小規模多機能型居宅介護（短期利用以外）"),
               g18("小規模多機能型居宅介護（短期利用）"))), MASK_ALL),
    ("看護小規模多機能型居宅介護",
     div12(add(g18("複合型サービス(看護小規模多機能型居宅介護)（短期利用以外）"),
               g18("複合型サービス(看護小規模多機能型居宅介護)（短期利用）"))), MASK_K),
]


def g17(name):
    return row7(w17, w17r[name], 4, 5, 8)


def g19(name):
    return row7(w19, w19r[name], 4, 5, 8)


KAISU = [
    ("訪問介護（回）", div12(g17("訪問介護（回）")), MASK_K),
    ("訪問入浴介護（回）", div12(g17("訪問入浴介護（回）")), MASK_ALL),
    ("訪問看護（回）", div12(g17("訪問看護（回）")), MASK_ALL),
    ("訪問リハビリテーション（回）", div12(g17("訪問リハビリテーション（回）")), MASK_ALL),
    ("通所介護（回）", div12(g17("通所介護（回）")), MASK_K),
    ("通所リハビリテーション（回）", div12(g17("通所リハビリテーション（回）")), MASK_K),
    ("短期入所生活介護（日）", div12(g17("短期入所生活介護（日）")), MASK_ALL),
    ("短期入所療養介護（老健）（日）",
     div12(g17("短期入所療養介護（介護老人保健施設）（日）")), MASK_ALL),
    ("短期入所療養介護（病院等）（日）", div12(g17("短期入所療養介護（病院等）（日）")), MASK_ALL),
    ("短期入所療養介護（介護医療院）（日）",
     div12(g17("短期入所療養介護（介護医療院）（日）")), MASK_ALL),
    ("地域密着型通所介護（回）", div12(g19("地域密着型通所介護（回）")), MASK_K),
    ("認知症対応型通所介護（回）", div12(g19("認知症対応型通所介護（回）")), MASK_ALL),
]

L = []


def P(s=""):
    L.append(s)


def emit(title, rows, note=""):
    P()
    P("=" * 86)
    P(f"■ {title}")
    if note:
        P(f"　{note}")
    P("=" * 86)
    P()
    P("【貼り付け用A（小数・推奨）】　入力欄が小数を受け付ける場合はこちら")
    P("\t".join(CARE))
    for nm, v, kind in rows:
        cells = [(f"{v[i]:.3f}".rstrip("0").rstrip(".") or "0") if kind[i] else ""
                 for i in range(7)]
        P("\t".join(cells) + f"\t← {nm}")
    P()
    P("【貼り付け用B（整数）】　入力欄が整数しか受け付けない場合はこちら")
    P("\t".join(CARE))
    lost = []
    for nm, v, kind in rows:
        iv = to_int(v)
        cells = [(str(iv[i]) if kind[i] else "") for i in range(7)]
        P("\t".join(cells) + f"\t← {nm}")
        if sum(v) > 0 and sum(iv) == 0:
            lost.append((nm, sum(v)))
    if lost:
        P()
        P("　★ 整数にすると0になり、令和9〜11年度も0のまま残るサービス")
        for nm, x in lost:
            P(f"　　・{nm}（令和7年度の月平均 {x:.3f}人／年間 {x*12:.0f}件）")
        P("　　→ この{}件は小数での入力（貼り付け用A）を強くお勧めします。".format(len(lost)))
        P("　　　 整数しか入らない場合は、当該サービスだけ 1 と置くことをご検討ください")
        P("　　　 （0 にすると計画に載らなくなります）。")
    P()
    P("【検算用（年報の年間値÷12。小数第3位まで）】")
    P(f"{'サービス':<34}" + "".join(f"{c:>9}" for c in CARE) + f"{'要支援計':>9}{'要介護計':>9}")
    for nm, v, kind in rows:
        P(f"{nm[:34]:<34}" + "".join(f"{x:>9.3f}" for x in v)
          + f"{sum(v[:2]):>9.3f}{sum(v[2:]):>9.3f}")


P("川崎町 第10期　見える化システム「令和8年度の実績（見込み）値」の差し替え値")
P("作成：ビズアップ公共コンサルティング 札幌事業所／令和8年9月11日")
P("保険者番号 04324")
P()
P("令和8年度は令和8年5月の1か月分であるため、完結年度である令和7年度の")
P("月平均（年報の年間値÷12）に差し替える。出所はすべて")
P("介護保険事業状況報告（年報・令和7年度）＝ 年報データ_2025_川崎町.xlsx。")
P()
P("※ 整数は、要支援群・要介護群それぞれで合計を四捨五入し、")
P("　 最大剰余法で配分している（丸めても計が合う）。")

# ---------------- 認定者数
P()
P("=" * 86)
P("■ ① 認定者数の実績値の設定（令和8年度）")
P("　出所：年報 様式1の5（12）要介護(要支援)認定者数（令和8年3月末現在）")
P("=" * 86)
for sex in ("男", "女"):
    P()
    P(f"【（{1 if sex=='男' else 2}）{sex}　貼り付け用（タブ区切り）】")
    P("\t".join(CARE))
    for i, lab in enumerate(NIN_ROWS):
        P("\t".join(str(int(x)) for x in nin[sex][i]) + f"\t← {lab}")
    tot = [sum(nin[sex][i][j] for i in range(6)) for j in range(7)]
    P(f"　（自動計算される第1号被保険者の行：{'／'.join(str(int(x)) for x in tot)}"
      f"　計{int(sum(tot))}人）")
ga = [sum(nin[s][i][j] for s in ("男", "女") for i in range(7)) for j in range(7)]
P()
P(f"　総数（男＋女、第2号を含む）：{'／'.join(str(int(x)) for x in ga)}"
  f"　計 {int(sum(ga))}人")

emit("② 施設・居住系サービス利用者数の実績見込み値の設定（令和8年度）",
     SHISETSU,
     "出所：年報 様式1の7(16)・(18)、様式1の6(15) の受給者数 ÷12")
emit("③ 在宅サービス利用者数の実績見込み値の編集（令和8年度）",
     ZAITAKU,
     "出所：年報 様式1の7(16)・(18)、様式2（件数）の受給者数 ÷12")
emit("④ 在宅サービス利用回（日）数の実績見込み値の編集（令和8年度）",
     KAISU,
     "出所：年報 様式1の7(17)・(19) の利用回（日）数 ÷12")

# ---------------- 検算
P()
P("=" * 86)
P("■ 検算：総括表の令和7年度の人数・回数と一致するか")
P("=" * 86)
wb2 = openpyxl.load_workbook(SOK, data_only=True)
ws2 = wb2["2_サービス別給付費"]
sok = {}
section = ""
cur = None
for r in range(8, ws2.max_row + 1):
    a, b, c = (ws2.cell(r, i).value for i in (1, 2, 3))
    if a and str(a).strip():
        section = str(a).strip()
    if not (c and str(c).strip()):
        continue
    if "給付費" in str(c):
        cur = str(b).strip() if b and str(b).strip() else section
    elif cur and not section.startswith("合計"):
        v = ws2.cell(r, 5).value
        if isinstance(v, (int, float)):
            sok[(cur, "人数" if "人数" in str(c) else "量")] = v

CHK = [
    ("訪問介護", "人数", ZAITAKU, "訪問介護", MASK_K),
    ("通所介護", "人数", ZAITAKU, "通所介護", MASK_K),
    ("通所リハビリテーション", "人数", ZAITAKU, "通所リハビリテーション", MASK_K),
    ("介護予防通所リハビリテーション", "人数", ZAITAKU, "通所リハビリテーション", "y"),
    ("福祉用具貸与", "人数", ZAITAKU, "福祉用具貸与", MASK_K),
    ("介護予防福祉用具貸与", "人数", ZAITAKU, "福祉用具貸与", "y"),
    ("（４）居宅介護支援", "人数", ZAITAKU, "介護予防支援・居宅介護支援", MASK_K),
    ("（３）介護予防支援", "人数", ZAITAKU, "介護予防支援・居宅介護支援", "y"),
    ("介護老人福祉施設", "人数", SHISETSU, "介護老人福祉施設", MASK_K),
    ("介護老人保健施設", "人数", SHISETSU, "介護老人保健施設", MASK_K),
    ("認知症対応型共同生活介護", "人数", SHISETSU, "認知症対応型共同生活介護", MASK_K),
    ("介護予防認知症対応型共同生活介護", "人数", SHISETSU,
     "認知症対応型共同生活介護", "y"),
    ("地域密着型介護老人福祉施設入所者生活介護", "人数", SHISETSU,
     "地域密着型介護老人福祉施設入所者生活介護", MASK_K),
    ("訪問介護", "量", KAISU, "訪問介護（回）", MASK_K),
    ("通所介護", "量", KAISU, "通所介護（回）", MASK_K),
    ("短期入所生活介護", "量", KAISU, "短期入所生活介護（日）", MASK_K),
]
P(f"{'総括表の項目':<40}{'総括表R7':>11}{'本書の値':>11}{'差':>9}")
ng = 0
for nm, kind, src, key, side in CHK:
    v = next((x for a, x, _ in src if a == key), None)
    if v is None or (nm, kind) not in sok:
        continue
    mine = sum(v[:2]) if side == "y" else sum(v[2:])
    d = mine - sok[(nm, kind)]
    if abs(d) > 0.0005:
        ng += 1
    P(f"{nm[:40]:<40}{sok[(nm,kind)]:>11.4f}{mine:>11.4f}{d:>9.4f}")
P()
P(f"→ 不一致 {ng} 件")

out = sys.argv[1] if len(sys.argv) > 1 else "R7差し替え値.txt"
with open(out, "w", encoding="utf-8") as f:
    f.write("\n".join(L) + "\n")
print("\n".join(L))
print(f"\n→ {out}")


# ================================================ 貼り付け用 Excel
def build_xlsx(path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wbx = Workbook()
    hdr_f = PatternFill("solid", fgColor="DCE6F1")
    grey = PatternFill("solid", fgColor="D9D9D9")
    warn = PatternFill("solid", fgColor="FFF2CC")
    bd = Border(*[Side(style="thin", color="808080")] * 4)
    b = Font(bold=True)

    def sheet(title, note, rows, first=False, labels=None):
        ws = wbx.active if first else wbx.create_sheet()
        ws.title = title
        ws["A1"] = note
        ws["A1"].font = Font(bold=True, size=12)
        ws["A3"] = "※ 白いセルの値をシステムの入力欄にそのまま貼り付けてください。"
        ws["A4"] = "※ グレーのセルは画面上「―」で入力できない欄です。"
        r0 = 6
        ws.cell(r0, 1, "サービス／区分").font = b
        ws.cell(r0, 1).fill = hdr_f
        ws.cell(r0, 1).border = bd
        for j, c in enumerate(CARE):
            x = ws.cell(r0, 2 + j, c)
            x.font = b
            x.fill = hdr_f
            x.border = bd
            x.alignment = Alignment(horizontal="center")
        ws.cell(r0, 9, "要支援計").font = b
        ws.cell(r0, 10, "要介護計").font = b
        for i, (nm, v, kind) in enumerate(rows):
            r = r0 + 1 + i
            ws.cell(r, 1, nm).border = bd
            for j in range(7):
                x = ws.cell(r, 2 + j)
                x.border = bd
                if kind[j]:
                    x.value = round(v[j], 3)
                    x.number_format = "0.###"
                    if 0 < v[j] < 0.5:
                        x.fill = warn
                else:
                    x.value = "―"
                    x.fill = grey
                    x.alignment = Alignment(horizontal="center")
            ws.cell(r, 9, round(sum(v[:2]), 3)).number_format = "0.###"
            ws.cell(r, 10, round(sum(v[2:]), 3)).number_format = "0.###"
        ws.column_dimensions["A"].width = 38
        for col in "BCDEFGH":
            ws.column_dimensions[col].width = 10
        ws.column_dimensions["I"].width = 10
        ws.column_dimensions["J"].width = 10
        ws.cell(r0 + len(rows) + 2, 1,
                "黄色のセルは月平均が0.5人未満のため、整数で入力すると0になります。"
                "小数で入力できない場合はご相談ください。")
        return ws

    # ① 認定者数
    ws = wbx.active
    ws.title = "①認定者数"
    ws["A1"] = "① 認定者数の実績値の設定（令和8年度）に貼り付ける令和7年度の値"
    ws["A1"].font = Font(bold=True, size=12)
    ws["A3"] = "出所：介護保険事業状況報告（年報・令和7年度）様式1の5（12）／令和8年3月末現在"
    r = 5
    for sex in ("男", "女"):
        ws.cell(r, 1, f"（{1 if sex=='男' else 2}）{sex}").font = b
        r += 1
        ws.cell(r, 1, "区分").font = b
        ws.cell(r, 1).fill = hdr_f
        ws.cell(r, 1).border = bd
        for j, c in enumerate(CARE):
            x = ws.cell(r, 2 + j, c)
            x.font = b
            x.fill = hdr_f
            x.border = bd
            x.alignment = Alignment(horizontal="center")
        ws.cell(r, 9, "合計").font = b
        r += 1
        for i, lab in enumerate(NIN_ROWS):
            ws.cell(r, 1, lab).border = bd
            for j in range(7):
                x = ws.cell(r, 2 + j, int(nin[sex][i][j]))
                x.border = bd
            ws.cell(r, 9, int(sum(nin[sex][i])))
            r += 1
        tot = [sum(nin[sex][i][j] for i in range(6)) for j in range(7)]
        ws.cell(r, 1, "（自動計算）第1号被保険者").font = Font(italic=True)
        for j in range(7):
            x = ws.cell(r, 2 + j, int(tot[j]))
            x.fill = grey
        ws.cell(r, 9, int(sum(tot))).fill = grey
        r += 3
    ws.column_dimensions["A"].width = 28
    for col in "BCDEFGHI":
        ws.column_dimensions[col].width = 10

    sheet("②施設・居住系", "② 施設・居住系サービス利用者数の実績見込み値の設定（令和8年度）"
          "に貼り付ける令和7年度の月平均", SHISETSU)
    sheet("③在宅利用者数", "③ 在宅サービス利用者数の実績見込み値の編集（令和8年度）"
          "に貼り付ける令和7年度の月平均", ZAITAKU)
    sheet("④在宅利用回日数", "④ 在宅サービス利用回（日）数の実績見込み値の編集（令和8年度）"
          "に貼り付ける令和7年度の月平均", KAISU)
    wbx.save(path)
    return path


if len(sys.argv) > 2:
    print("Excel:", build_xlsx(sys.argv[2]))
