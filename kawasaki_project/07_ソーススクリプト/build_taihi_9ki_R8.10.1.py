# -*- coding: utf-8 -*-
"""第9期の計画値と実績値の対比表（令和8年10月1日）

08_作業順位 の順位3。ペンディング束3（第9期の実績・12件）を前に進める。

**町のご記入を待っていた部分のうち、介護給付サービスの実績は
当方の手元で埋められる。** 令和7年度の年報（様式2 給付費）による
独立算定の結果が `川崎町_第10期_計画見込量_R8.9.15.xlsx` にあり、
計画値は第9期計画書（令和6年3月）の「サービス別の人数・給付額見込み」に
印字されている。両者を機械で突き合わせる。

  入力　05_試算・管理シート/川崎町_サービス別_計画値実績対比表_R8.9.xlsx
        05_試算・管理シート/川崎町_第10期_計画見込量_R8.9.15.xlsx
  出力　05_試算・管理シート/川崎町_第9期_計画値実績対比表_R8.10.1.xlsx

**残るのは地域支援事業・高齢者福祉サービスの実績であり、
これは町のご記入を要する（確認事項No.8・No.133）。**
どれが埋まり、どれが残るかを 07_記入状況 で数える。

⚠ `川崎町_第10期_計画見込量_R8.9.15.xlsx` は
   `project_mikomi_R8.9.15.py` を import すると書き換わるため、
   本スクリプトは openpyxl で読むだけにする。
"""
import re
import shutil
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SRC = "05_試算・管理シート/川崎町_サービス別_計画値実績対比表_R8.9.xlsx"
MIKOMI = "05_試算・管理シート/川崎町_第10期_計画見込量_R8.9.15.xlsx"
OUT = "05_試算・管理シート/川崎町_第9期_計画値実績対比表_R8.10.1.xlsx"

OKF = PatternFill("solid", fgColor="E2EFDA")     # 当方で埋めた
NGF = PatternFill("solid", fgColor="FFF2CC")     # 町のご記入待ち
HEAD = PatternFill("solid", fgColor="5B9BD5")
TITLE = PatternFill("solid", fgColor="1F4E78")
LEAD = PatternFill("solid", fgColor="F2F2F2")

# 第9期計画書（令和6年3月）p.92 の区分別の計画値（千円／年）
#   令和5年度・令和6年度・令和7年度・令和8年度の4か年が印字されている。
#   第9期は令和6〜8年度。
KEIKAKU_KUBUN = {
    "居宅サービス（介護予防を含む）": {"R6": 352330, "R7": 353651,
                                        "R8": 354822},
    "地域密着型サービス（介護予防を含む）": {"R6": 179657, "R7": 179884,
                                              "R8": 179884},
    "施設サービス": {"R6": 476882, "R7": 477485, "R8": 481428},
}
# 令和7年度の実績（年報 様式2 給付費・単位 千円）
#   素案 2-4 の表及び図2-6 と同じ値（9.98億円）
JISSEKI_KUBUN_R7 = {
    "居宅サービス（介護予防を含む）": 322680,
    "地域密着型サービス（介護予防を含む）": 182190,
    "施設サービス": 493350,
}

# 見込量ワークブックのサービス名 → 対比表のサービス名（表記の揺れを吸収）
ALIAS = {
    "短期入所療養介護（老健）": "短期入所療養介護",
    "介護予防短期入所療養介護（老健）": "介護予防短期入所療養介護",
    "地域密着型介護老人福祉施設入所者生活介護":
        "地域密着型介護老人福祉施設入居者生活介護",
    "介護老人福祉施設": "介護老人福祉施設（特別養護老人施設）",
    "介護老人保健施設": "介護老人保健施設（老人保健施設）",
    "訪問介護": "訪問介護（ホームヘルプサービス）",
    "特定介護予防福祉用具購入費": "介護予防特定福祉用具購入費",
    "住宅改修費": "住宅改修",
    # 居宅介護支援と介護予防支援は見込量ワークブックでは1行に束ねてあり、
    # 対比表の2行に割り振れないため、当方では埋めない。
}


def norm(s):
    """突き合わせ用に、括弧書き・記号・空白を取り除く。"""
    s = str(s or "")
    s = re.sub(r"[（(].*?[）)]", "", s)
    s = re.sub(r"[\s　・（）()／/]", "", s)
    return s


def read_mikomi():
    """見込量ワークブックから令和7年度実績（給付費・利用者数）を読む。"""
    wb = openpyxl.load_workbook(MIKOMI, data_only=True)
    out = {}
    for sh in ("③介護予防サービス", "④介護サービス"):
        ws = wb[sh]
        for r in range(2, ws.max_row + 1):
            name = ws.cell(r, 1).value
            item = str(ws.cell(r, 2).value or "")
            v = ws.cell(r, 3).value
            if not name or not isinstance(v, (int, float)):
                continue
            if "合計" in str(name):
                continue
            key = norm(ALIAS.get(name, name))
            d = out.setdefault(key, {"名": name})
            if item.startswith("給付費"):
                d["給付費"] = v
            elif item.startswith("利用者数"):
                d["利用者数"] = v
    return out


def main():
    mik = read_mikomi()
    shutil.copy(SRC, OUT)
    wb = openpyxl.load_workbook(OUT)

    filled, left, unmatched = 0, 0, []
    SHEETS = ["01_居宅サービス", "02_地域密着型サービス", "03_施設サービス",
              "04_総合事業・一般介護予防", "05_包括的支援事業・任意事業",
              "06_高齢者福祉サービス"]
    tally = {}
    for sh in SHEETS:
        ws = wb[sh]
        head = [str(ws.cell(4, c).value or "") for c in range(1, 16)]
        try:
            c_name = head.index("サービス・事業名") + 1
            c_ind = head.index("指標") + 1
            c_r7k = head.index("R7計画") + 1
            c_r7j = head.index("R7実績") + 1
        except ValueError:
            continue
        n_fill, n_left = 0, 0
        for r in range(5, ws.max_row + 1):
            name = ws.cell(r, c_name).value
            if not name:
                continue
            ind = str(ws.cell(r, c_ind).value or "")
            keikaku = ws.cell(r, c_r7k).value
            if ws.cell(r, c_r7j).value not in (None, ""):
                continue
            # 介護給付サービスのみ当方で埋められる
            if sh.startswith(("04_", "05_", "06_")):
                if keikaku not in (None, "", "―"):
                    n_left += 1
                    ws.cell(r, c_r7j).fill = NGF
                continue
            key = norm(name)
            d = mik.get(key)
            if d is None:
                if keikaku not in (None, "", "―") and keikaku != 0:
                    unmatched.append(f"{sh}／{name}")
                    n_left += 1
                    ws.cell(r, c_r7j).fill = NGF
                continue
            # 給付費（千円／年）だけを埋める。
            # 延利用人数は年報の様式1の6によるものであり、
            # 見込量ワークブックの「人／月」から12倍して作ると
            # 年報の延べ人数と一致しないため、当方では埋めない。
            v = None
            if "給付費" in ind:
                v = d.get("給付費")
                if v is not None:
                    v = round(v)
            if v is None:
                n_left += 1
                ws.cell(r, c_r7j).fill = NGF
                continue
            ws.cell(r, c_r7j).value = v
            ws.cell(r, c_r7j).fill = OKF
            n_fill += 1
            # 達成率の欄には既に数式が入っているため書き換えない
        tally[sh] = (n_fill, n_left)
        filled += n_fill
        left += n_left

    # ══════════════════════════ 区分別のまとめ（新しいシート）
    if "08_区分別のまとめ" in wb.sheetnames:
        del wb["08_区分別のまとめ"]
    ws = wb.create_sheet("08_区分別のまとめ")
    ws.cell(1, 1).value = "第9期の計画値と実績値の対比（区分別・給付費）"
    ws.cell(1, 1).font = Font(name="游ゴシック", size=14, bold=True,
                              color="FFFFFF")
    ws.cell(1, 1).fill = TITLE
    ws.cell(2, 1).value = (
        "計画値は第9期計画書（令和6年3月）の「介護保険給付額サービス別の"
        "人数・給付額見込み目標」による。"
        "実績値は介護保険事業状況報告（年報・令和7年度／様式2 給付費）による。"
        "令和6年度の実績は区分別に把握できていないため、"
        "令和7年度（第9期の完結年度であり第10期算定の基準年度）で対比する。"
        "令和8年度は月報4か月分であり年度の実績として扱わない。"
        "⚠ 交付金 推進 目標Ⅰ（ⅱ）2「事業計画の進捗状況」（12点・本町0点）が"
        "求める分析はこの対比である。")
    ws.cell(2, 1).font = Font(name="游ゴシック", size=9)
    ws.cell(2, 1).fill = LEAD
    ws.cell(2, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)
    ws.row_dimensions[2].height = 72
    head = ["区分", "令和7年度 計画値（千円）", "令和7年度 実績値（千円）",
            "差（千円）", "達成率（％）", "乖離の要因（当方の見立て）"]
    for c, (h, w) in enumerate(zip(head, [34, 20, 20, 16, 13, 78]), start=1):
        cell = ws.cell(4, c)
        cell.value = h
        cell.font = Font(name="游ゴシック", size=9, bold=True, color="FFFFFF")
        cell.fill = HEAD
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = w
    YOIN = {
        "居宅サービス（介護予防を含む）":
            "計画を30,971千円（8.8％）下回る。訪問介護・通所介護の"
            "利用が計画を下回っていることによる。"
            "町内事業所の縮小（訪問介護の登録ヘルパーの減少）と、"
            "施設・居住系への移行が同時に進んでいると見られる"
            "（確認事項No.50・No.51）。",
        "地域密着型サービス（介護予防を含む）":
            "計画を2,306千円（1.3％）上回る。認知症対応型共同生活介護"
            "（グループホーム）がほぼ計画どおりに推移している。"
            "定期巡回・認知症対応型通所介護は区域内に事業所がないが"
            "給付実績がある（町外事業所の利用）。",
        "施設サービス":
            "計画を15,865千円（3.3％）上回る。"
            "特別養護老人ホーム・介護老人保健施設とも計画を上回る。"
            "住所地特例24人（令和8年3月末）に表れるとおり、"
            "町外施設の利用が常態化していることによる。",
    }
    r = 5
    tot_k = tot_j = 0
    for k, keikaku in KEIKAKU_KUBUN.items():
        kk = keikaku["R7"]
        jj = JISSEKI_KUBUN_R7[k]
        tot_k += kk
        tot_j += jj
        for c, v in enumerate((k, kk, jj, jj - kk,
                               round(jj / kk * 100, 1), YOIN[k]), start=1):
            cell = ws.cell(r, c)
            cell.value = v
            cell.font = Font(name="游ゴシック", size=9)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    for c, v in enumerate(("総給付費", tot_k, tot_j, tot_j - tot_k,
                           round(tot_j / tot_k * 100, 1),
                           "計画を12,800千円（1.3％）下回る。"
                           "区分の間で振れが相殺されており、総額では計画に"
                           "ほぼ沿っている。**在宅から施設・居住系への"
                           "移行が進んでいることが、総額の一致に隠れている。**"
                           "第10期は、この移行が続くことを前提に見込む。"),
                          start=1):
        cell = ws.cell(r, c)
        cell.value = v
        cell.font = Font(name="游ゴシック", size=9, bold=True)
        cell.fill = OKF
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A5"

    # ══════════════════════════ 07_記入状況
    ws = wb["07_記入状況"]
    ws.cell(2, 1).value = (
        f"令和8年10月1日更新　令和7年度の実績欄のうち、介護給付サービスの"
        f"{filled}件を当方が年報から埋めた（緑）。"
        f"残る{left}件は町のご記入を要する（黄）。")
    r = 4
    for sh in SHEETS:
        if sh in tally:
            ws.cell(r, 4).value = tally[sh][0]
            ws.cell(r, 5).value = tally[sh][1]
            r += 1
    ws.cell(3, 4).value = "当方が埋めたR7実績"
    ws.cell(3, 5).value = "町のご記入待ち"

    wb.save(OUT)
    print("保存：", OUT)
    print(f"  当方が埋めた令和7年度の実績 {filled}件")
    print(f"  町のご記入を要するもの　　　{left}件")
    for sh, (a, b) in tally.items():
        print(f"    {sh}　埋めた {a}／残り {b}")
    if unmatched:
        print("  ⚠ 見込量ワークブックに対応するサービスがない行"
              f"（{len(unmatched)}件）")
        for u in unmatched[:12]:
            print("     ", u)
    print("  ── 区分別のまとめ")
    print(f"    総給付費　計画 {tot_k:,}千円／実績 {tot_j:,}千円"
          f"／達成率 {tot_j / tot_k * 100:.1f}％")
    # 自己点検　区分の合計が素案の総給付費（9.98億円）と一致するか
    if abs(tot_j - 998220) > 1:
        print("  × 区分別の実績の合計が素案 2-4 の総額と一致しない")
        sys.exit(1)
    print("    ○ 区分別の実績の合計は素案 2-4 の総額（998,220千円）と一致")


if __name__ == "__main__":
    main()
