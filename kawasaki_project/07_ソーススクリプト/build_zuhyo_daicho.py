# -*- coding: utf-8 -*-
"""図表データ管理台帳

大雪地区広域連合の `build_zuhyo_daicho.py` と同じ考え方による。
計画本文に差し込む図の数値をエクセルの上で管理し、
数値を直せば計画本文の図が追随するようにする。

  00_この表について
  01_図表台帳　　　　　計画本文に差し込む図の一覧
  02_素案との突合　　　図の数値と素案の表・本文との突合
  03_交付金との紐付け　交付金の評価が求める分析と図表の対応
  04_不足している図　　新たに作る図の候補
  05_資料7との突合　　素案巻末の「図表番号一覧」と実際の差し込みの突合
  D_fig…　　　　　　　図ごとのデータシート（ここで数値を直す）
  99_自己点検　　　　　1件でも不適合があると終了コード1で終わる

使い方：
  python3 07_ソーススクリプト/build_zuhyo_daicho.py
  （数値を直すときは data_zuhyo.py を直し、
   build_plan_figures.py → fix_soan_v260_figures.py の順に実行する）
"""
import glob
import os
import re
import sys
import zipfile

import docx
import openpyxl
from docx.oxml.ns import qn
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, "07_ソーススクリプト")
from data_zuhyo import KOFU, KOUHO, WAKU, ZU  # noqa: E402

OUT = "05_試算・管理シート/川崎町_図表データ管理台帳_R8.9.30.xlsx"
SOAN = "01_第10期_最新版成果品/川崎町_計画書素案_v2.8_第9期対比版.docx"
SOAN_ALT = "01_第10期_最新版成果品/川崎町_計画書素案_v2.7_図表追加版.docx"
FIGDIR = "08_図表"

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD = PatternFill("solid", fgColor="D9E1F2")
OKF = PatternFill("solid", fgColor="E2EFDA")
NGF = PatternFill("solid", fgColor="FCE4D6")
NOTE = PatternFill("solid", fgColor="FFF2CC")
INPUT = PatternFill("solid", fgColor="FFFFCC")
WRAP = Alignment(vertical="top", wrap_text=True)


def sheet(wb, name, title, lead, head, widths):
    ws = wb.create_sheet(name)
    ws.cell(1, 1, title).font = Font(bold=True, size=12)
    ws.cell(2, 1, lead).alignment = WRAP
    ws.merge_cells(start_row=2, start_column=1, end_row=2,
                   end_column=max(len(head), 2))
    ws.row_dimensions[2].height = 30
    for c, (h, w) in enumerate(zip(head, widths), start=1):
        cell = ws.cell(4, c, h)
        cell.font = Font(bold=True)
        cell.fill = HEAD
        cell.border = BOX
        cell.alignment = WRAP
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A5"
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    return ws


def put(ws, rows, judge=None, start=5):
    for i, r in enumerate(rows, start=start):
        for j, v in enumerate(r, start=1):
            cell = ws.cell(i, j, v)
            cell.border = BOX
            cell.alignment = WRAP
        if judge:
            f = judge(r)
            if f:
                for j in range(1, len(r) + 1):
                    ws.cell(i, j).fill = f
    return start + len(rows)


def soan_path():
    return SOAN if os.path.exists(SOAN) else SOAN_ALT


def cap_of(d):
    """素案に置くキャプションの文字列。

    画像に焼き込む表題（title）に、素案のキャプションにだけ添える
    断り書き（cap_add）を足したもの。図9-1 の「令和8年9月時点の算定・
    確定値は第3回策定委員会で決定」がこれに当たる。
    """
    return d["title"] + (d["cap_add"] or "")


def read_soan():
    """素案から、図のキャプション・資料7の一覧・表の数値を読む。"""
    p = soan_path()
    d = docx.Document(p)
    caps, order, refs = {}, [], []
    kids = list(d.element.body.iterchildren())
    sec = None
    for i, el in enumerate(kids):
        if el.tag != qn("w:p"):
            continue
        t = "".join(n.text or "" for n in el.iter(qn("w:t"))).strip()
        ms = re.match(r"^(\d+-\d+)[　 ]", t)
        if ms:
            sec = ms.group(1)
        m = re.match(r"^(図\d+-\d+)　(.+)$", t)
        if m:
            caps[m.group(1)] = m.group(2)
            order.append(m.group(1))
            continue
        # 本文からの参照（キャプションそのものは除く）
        for mr in re.finditer(r"図\d+-\d+", t):
            refs.append((mr.group(0), sec, t[:60]))
    shiryo7 = []
    for t in d.tables:
        h = [c.text.strip() for c in t.rows[0].cells]
        if h[:2] == ["図番号", "図表名"]:
            for r in t.rows[1:]:
                shiryo7.append([c.text.strip() for c in r.cells])
    # 画像の数
    with zipfile.ZipFile(p) as z:
        n_img = sum(1 for n in z.namelist() if n.startswith("word/media/"))
    tables = []
    for t in d.tables:
        tables.append([[c.text.strip() for c in r.cells] for r in t.rows])
    return p, caps, order, shiryo7, n_img, tables, d, refs


def find_in_tables(tables, *needles):
    """素案の表の中に、すべての語を含む行があるかを見る。"""
    for tb in tables:
        for row in tb:
            s = "".join(row)
            if all(n in s for n in needles):
                return "／".join(x for x in row if x)[:80]
    return None


def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    p_soan, caps, order, shiryo7, n_img, tables, doc, refs = read_soan()
    pngs = {os.path.basename(x) for x in glob.glob(f"{FIGDIR}/*.png")}
    used = {d["png"] for d in ZU}

    # ══════════════════════════ 00
    ws = sheet(wb, "00_この表について", "川崎町 第10期計画　図表データ管理台帳",
               "計画本文に差し込む図の数値をエクセルの上で管理し、"
               "数値を直せば計画本文の図が追随するようにしたものです。"
               "大雪地区広域連合の図表データ管理台帳と同じ考え方によります。",
               ["No.", "区分", "内容", "点数・箇所"], [6, 26, 92, 20])
    n = put(ws, [
        ("1", "計画本文に差し込む図",
         "図の表題・種類・数値・体裁を1件ずつ持つ。各図に1枚のデータシートを設けている",
         f"{len(ZU)}点"),
        ("2", "素案に差し込んでいるもの",
         f"素案（{os.path.basename(p_soan)}）の本文に画像として置いているもの",
         f"{n_img}点"),
        ("3", "08_図表にある画像",
         "作図の出力先。使っていない古い版の画像が残っていないかを99シートで確かめる",
         f"{len(pngs)}点（うち使用{len(used & pngs)}点）"),
        ("4", "数値の正本", "07_ソーススクリプト/data_zuhyo.py。"
         "本台帳のデータシート（D_fig…）はここから作っている", "―"),
        ("5", "作図", "07_ソーススクリプト/build_plan_figures.py", "―"),
        ("6", "素案への差し込み",
         "07_ソーススクリプト/fix_soan_v260_figures.py", "―"),
    ])
    n += 1
    for t in (
        "【数値を直す手順】①本台帳のデータシート（D_fig…）で直したい値を確かめる　"
        "②07_ソーススクリプト/data_zuhyo.py の同じ値を直す　"
        "③build_plan_figures.py を実行して図を作り直す　"
        "④fix_soan_v260_figures.py を実行して素案の画像を差し替える　"
        "⑤本台帳を作り直して99シートの自己点検が適合であることを確かめる。",
        "【なぜ台帳を置くか】本台帳を置く前、図の数値は create_charts.py・"
        "make_charts_v114.py・make_charts_v115.py・make_charts_v116.py の"
        "4つに散らばっていました。"
        "そのため素案Ver.2.3で保険料の3パターンを是正した"
        "（6,170→6,143円・5,518→5,464円）のに図9-1が追随せず、"
        "本文と図で値が食い違っていました。",
        "【図番号の焼き込み】図の表題は画像の中に描き込んでいます。"
        "素案Ver.2.5で図番号を本文の出現順に振り直したとき、"
        "画像の中の番号が追随していませんでした"
        "（画像内「図2-5」＝素案の図2-2 など4点）。"
        "本台帳と作図は同じ data_zuhyo.py の no を用いるため、"
        "この食い違いは起きません。",
    ):
        ws.cell(n, 1, t).alignment = WRAP
        ws.cell(n, 1).fill = NOTE
        ws.merge_cells(start_row=n, start_column=1, end_row=n, end_column=4)
        ws.row_dimensions[n].height = 46
        n += 1

    # ══════════════════════════ 01_図表台帳
    ws = sheet(wb, "01_図表台帳", "図表台帳　計画本文に差し込む図の一覧",
               "図表番号・表題・差し込み先・出典は data_zuhyo.py から、"
               "系列と区分の数は図の数値から読んでいます。",
               ["図表番号", "図の表題", "種類", "差し込み先", "資料（出典）",
                "系列数", "区分数", "画像", "データシート", "素案との突合"],
               [8, 44, 18, 26, 40, 7, 7, 22, 22, 14])
    rows01 = []
    for d in ZU:
        series = len(d["head"]) - 1
        rows01.append((d["no"], d["title"], d["kind"], d["sec"], d["src"],
                       series, len(d["rows"]), d["png"],
                       "D_" + d["png"].replace(".png", ""),
                       "一致" if caps.get(d["no"]) == cap_of(d) else "要確認"))
    put(ws, rows01,
        judge=lambda r: OKF if r[9] == "一致" else NGF)

    # ══════════════════════════ 02_素案との突合
    ws = sheet(wb, "02_素案との突合", "素案の表・本文との突合",
               "図の数値を素案の表・本文と機械で突き合わせました。"
               "同じ数値が2か所にある場合は、どちらかを直したときに"
               "もう一方が取り残されます。",
               ["図表番号", "図の表題", "突合の手掛かり", "素案に見つかったもの",
                "判定"], [8, 40, 52, 56, 12])
    rows02 = []
    # 突合の手掛かりは data_zuhyo.py の needle が正本（2か所に持たない）
    for d in ZU:
        hits = []
        for nd in d["needle"]:
            f = find_in_tables(tables, nd)
            if f is None:
                body = "\n".join(p.text for p in doc.paragraphs)
                f = nd if nd in body else None
            hits.append((nd, f))
        miss = [nd for nd, f in hits if f is None]
        rows02.append((d["no"], d["title"], d["check"],
                       "／".join(f"{nd}：{'あり' if f else 'なし'}"
                                for nd, f in hits),
                       "一致" if not miss else f"食い違い（{'・'.join(miss)}）"))
    put(ws, rows02, judge=lambda r: OKF if r[4] == "一致" else NGF)

    # ══════════════════════════ 03_交付金との紐付け
    ws = sheet(wb, "03_交付金との紐付け", "保険者機能強化推進交付金等の評価と図表の対応",
               "交付金の評価指標のうち、地域の状況の分析・把握が評価の対象となるものを"
               "取り上げ、それを裏づける図表が本文にあるかを見ています。",
               ["交付金", "評価指標の大項目", "本町の得点", "図表",
                "対応する図表番号", "備考"], [8, 34, 14, 10, 24, 76])
    put(ws, KOFU, judge=lambda r: NGF if r[3] == "ない" else
        (NOTE if r[3] == "一部" else OKF))

    # ══════════════════════════ 04_不足している図
    ws = sheet(wb, "04_不足している図", "新たに作る図の候補",
               "交付金の評価が求める分析（03シート）と、"
               "大雪地区広域連合の協議用素案（図36点）を比較材料として読んで"
               "気づいた点によります。",
               ["No.", "作る図の案", "何のために要るか", "数値の出所",
                "作れる時期", "状態"], [6, 40, 54, 34, 26, 10])
    put(ws, KOUHO)

    # ══════════════════════════ 05_資料7との突合
    ws = sheet(wb, "05_資料7との突合", "素案巻末「図表番号一覧」との突合",
               "巻末の一覧に掲げた図番号・図表名・掲載箇所・出典を、"
               "実際の差し込みと突き合わせました。"
               "一覧は計画の読み手に示すものであり、実際と食い違ってはなりません。",
               ["図番号", "図表名（資料7）", "実際の表題", "掲載箇所（資料7）",
                "実際の差し込み先", "出典（資料7）", "判定"],
               [8, 40, 40, 26, 26, 36, 14])
    rows05 = []
    s7 = {r[0]: r for r in shiryo7}
    for d in ZU:
        r = s7.get(d["no"])
        if r is None:
            rows05.append((d["no"], "―", d["title"], "―", d["sec"], "―",
                           "一覧にない"))
            continue
        same = (r[1] == d["title"] and d["sec"].split()[0] in r[2]
                and r[3] == d["src"])
        rows05.append((d["no"], r[1], d["title"], r[2], d["sec"], r[3],
                       "一致" if same else "要是正"))
    for no in set(s7) - {d["no"] for d in ZU}:
        rows05.append((no, s7[no][1], "―", s7[no][2], "差し込みなし",
                       s7[no][3], "実際にない"))
    put(ws, rows05, judge=lambda r: OKF if r[6] == "一致" else NGF)

    # ══════════════════════════ データシート
    for d in ZU:
        name = ("D_" + d["png"].replace(".png", ""))[:31]
        ws = wb.create_sheet(name)
        ws.cell(1, 1, f"{d['no']}　{d['title']}").font = Font(bold=True, size=12)
        ws.cell(2, 1, f"種類：{d['kind']}／差し込み先：{d['sec']}　　"
                      f"資料：{d['src']}").alignment = WRAP
        ws.cell(3, 1, "数値のセルのみを直してください。"
                      "行・列の増減、見出しの書き換えは行わないでください。"
                      "直したあとは data_zuhyo.py の同じ値を直し、"
                      "build_plan_figures.py を実行し直します。").alignment = WRAP
        for r in (1, 2, 3):
            ws.merge_cells(start_row=r, start_column=1, end_row=r,
                           end_column=max(len(d["head"]), 3))
            ws.row_dimensions[r].height = 22 if r == 1 else 28
        for c, h in enumerate(d["head"], start=1):
            cell = ws.cell(5, c, h)
            cell.font = Font(bold=True)
            cell.fill = HEAD
            cell.border = BOX
            cell.alignment = WRAP
            ws.column_dimensions[get_column_letter(c)].width = \
                34 if c == 1 else 18
        for i, row in enumerate(d["rows"], start=6):
            for j, v in enumerate(row, start=1):
                cell = ws.cell(i, j, "" if v is None else v)
                cell.border = BOX
                cell.alignment = WRAP
                if j > 1:
                    cell.fill = INPUT
        r = 6 + len(d["rows"]) + 1
        if d["note"]:
            ws.cell(r, 1, "注記：" + d["note"].replace("\n", " ")).alignment = WRAP
            ws.cell(r, 1).fill = NOTE
            ws.merge_cells(start_row=r, start_column=1, end_row=r,
                           end_column=max(len(d["head"]), 3))
            ws.row_dimensions[r].height = 40
            r += 1
        ws.cell(r, 1, "突合：" + d["check"]).alignment = WRAP
        ws.cell(r, 1).fill = NOTE
        ws.merge_cells(start_row=r, start_column=1, end_row=r,
                       end_column=max(len(d["head"]), 3))
        ws.row_dimensions[r].height = 34

    # ══════════════════════════ 99_自己点検
    ws = sheet(wb, "99_自己点検", "自己点検",
               "1件でも不適合があると終了コード1で終わります。"
               "数値を直したときは本シートの適合を確かめてください。",
               ["No.", "確かめたこと", "式・方法", "結果", "判定"],
               [6, 54, 44, 44, 10])
    checks = []

    def add(name, how, res, ok):
        checks.append((str(len(checks) + 1), name, how, res,
                       "適合" if ok else "不適合"))

    add("図の定義の数と、作図された画像の数が一致すること",
        "len(ZU) == 08_図表 にある定義の画像",
        f"定義 {len(ZU)} ／ 画像 {len(used & pngs)}",
        len(used & pngs) == len(ZU))
    add("素案に差し込んだ画像の数と、図の定義の数が一致すること",
        "word/media の画像数 == len(ZU)",
        f"素案 {n_img} ／ 定義 {len(ZU)}", n_img == len(ZU))
    miss_cap = [d["no"] for d in ZU if caps.get(d["no"]) != cap_of(d)]
    add("素案のキャプションが図の表題と一致すること",
        "素案の「図N-M　表題」== ZU の title＋cap_add",
        "食い違い なし" if not miss_cap else "／".join(miss_cap),
        not miss_cap)
    add("素案の図番号が本文の出現順であること",
        "出現順 == sorted(no)",
        "／".join(order) if order else "―",
        order == [d["no"] for d in ZU])
    bad02 = [r[0] for r in rows02 if r[4] != "一致"]
    add("図の数値が素案の表・本文と一致すること",
        "02シートの突合",
        "食い違い なし" if not bad02 else "／".join(bad02), not bad02)
    bad05 = [r[0] for r in rows05 if r[6] != "一致"]
    add("巻末の図表番号一覧が実際の差し込みと一致すること",
        "05シートの突合",
        "食い違い なし" if not bad05 else "／".join(bad05), not bad05)
    # 本文の参照が、その図のある節と食い違っていないか
    #   図番号を振り直したときに本文の参照が追随しないことが実際に起きた
    #   （2-2 の本文が高齢化率の図を「図2-5」と参照していた）。
    secs = {d["no"]: d["sec"].split()[0] for d in ZU}
    bad_ref = []
    for no, sec, txt in refs:
        if no not in secs:
            bad_ref.append(f"{no}（素案にない図・{sec}）")
        elif sec is not None and sec != secs[no] and not sec.startswith("1-"):
            # 第1章は計画の全体を述べる章であり、他章の図を指すことがある
            bad_ref.append(f"{no}を{sec}の本文が参照（図は{secs[no]}）")
    add("本文の図の参照が、その図のある節と合っていること",
        "本文の段落の「図N-M」の節 == data_zuhyo.py の sec"
        "（表の中の参照は巻末の一覧と重なるため見ていない）",
        f"参照{len(refs)}件・食い違い なし" if not bad_ref
        else "／".join(bad_ref),
        not bad_ref)
    # 未作成の候補には、数値が届いたらすぐ作れるよう枠を置いてあるか
    #   （08_作業順位 の順位9。到着後の手戻りを防ぐため）
    midone = [k[0] for k in KOUHO if str(k[5]).startswith("未作成")]
    waku_no = {w["kou"] for w in WAKU}
    nowaku = [n for n in midone if n not in waku_no]
    add("未作成の図に、作図の枠が用意されていること",
        "KOUHO の未作成 ⊆ WAKU の候補番号",
        f"未作成 {len(midone)}件・枠 {len(WAKU)}件・枠なし なし"
        if not nowaku else "枠がない候補：" + "／".join(nowaku),
        not nowaku)
    stale = sorted(pngs - used)
    add("08_図表 に使っていない画像が残っていないこと",
        "08_図表/*.png - ZU の png",
        "なし" if not stale else f"{len(stale)}点：" + "／".join(stale[:5]),
        not stale)
    # data_zuhyo.py を置く前に図の数値を持っていた旧版の作図スクリプト。
    # 07_ソーススクリプト に残っている限り、数値の正本が2か所になる。
    OLD_SCRIPTS = ("create_charts", "make_chart_premium_v110",
                   "make_charts_v114", "make_charts_v115", "make_charts_v116")
    left = [f"{x}.py" for x in OLD_SCRIPTS
            if os.path.exists(f"07_ソーススクリプト/{x}.py")]
    add("図の数値の正本が1か所であること",
        "data_zuhyo.py のほかに図の数値を持つスクリプトが"
        "07_ソーススクリプト にないこと",
        "data_zuhyo.py のみ" if not left
        else f"{len(left)}点が残っている：" + "／".join(left)
             + "。09_元資料へ移すか削除する（04シートの状態欄を参照）",
        not left)
    put(ws, checks, judge=lambda r: OKF if r[4] == "適合" else NGF)

    wb.save(OUT)
    print("保存：", OUT)
    print(f"  素案：{os.path.basename(p_soan)}")
    print(f"  シート {len(wb.sheetnames)}（うちデータシート {len(ZU)}）")
    ngs = [c for c in checks if c[4] == "不適合"]
    for c in checks:
        mark = "○" if c[4] == "適合" else "×"
        print(f"  {mark} {c[1]}：{c[3]}")
    print(f"  自己点検 適合 {len(checks) - len(ngs)}件／不適合 {len(ngs)}件")
    if ngs:
        sys.exit(1)


if __name__ == "__main__":
    main()
