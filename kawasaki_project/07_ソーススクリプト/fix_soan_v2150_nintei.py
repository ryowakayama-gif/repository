# -*- coding: utf-8 -*-
"""計画素案 Ver.2.15（認定者数・第1号被保険者数の実績の是正）

**他メンバー作成の第10期計画素案（令和8年10月16日受領）と当方の正本を
突き合わせ、年報（令和6・7年度の2か年）で確かめたところ、
当方の素案 2-3 の実績値が年報と20か所食い違っていた。**

  入力　01_第10期_最新版成果品/川崎町_計画書素案_v2.14_見込量補完版.docx
  出力　01_第10期_最新版成果品/川崎町_計画書素案_v2.15_実績是正版.docx

是正1　第1号被保険者数（2-3 の認定率の表・1-6 の本文）
  当方　令和6年度3,270人／令和7年度3,285人／令和8年度3,257人
  年報　令和6年度末3,261人／令和7年度末3,240人（令和8年度は3,229人）
  ⚠ **3,270は令和5年度末、3,285は令和4年度末の値であり、年度が2年ずれていた。**
  calc_hokenryo_R8.9.2.py の HIHOKEN は正しい値を持っており、
  **保険料の算定には影響しない。** 2-1 の表（令和7年度末3,240人）とも
  食い違っていた。

是正2　要介護度別認定者数（2-3 の表）
  当方　令和6年度 27/73/69/126/126/89/52（計562人）
        令和7年度 33/85/70/116/131/84/49（計568人）
  年報　令和6年度 28/75/77/119/128/85/53（計565人）
        令和7年度 36/90/61/103/130/96/50（計566人）
  ⚠ 当方の値は年報とも月報とも別系統（総括表）のものであった。
  ⚠ **他メンバー作成の素案の値は年報と完全に一致していた。**
  見込量は令和8年度以降の見える化システムの推計値の比で伸ばしているため、
  **見込量には影響しない。**

是正3　認定率の再計算と本文の書き直し
  正しい値で計算すると、認定率は令和6年度17.1％→令和11年度18.1％の
  1.0ポイント上昇（当方は16.9％→18.1％の1.2ポイントと書いていた）。
  ⚠ **内訳の動きが変わる。** 令和6→7年度は、要支援1・2が23人増え、
  要介護1・2が32人減り、要介護4が11人増える。
  **総数はほぼ横ばいだが、軽度と重度が増え、中間が減る二極化である。**
  当方は「要支援2と要介護4が増加し、要介護2が減少」と書いており
  向きは合っていたが、要介護1の減（▲16人）を捉えていなかった。

是正4　第1号被保険者のいる世帯数（2-1）
  当方　令和3年度末 2,181世帯
  年報　令和7年度末 2,193世帯（令和6年度末 2,198世帯）
  ⚠ 年報に新しい実数があるのに、令和3年度末の値を使っていた。

⚠ **保険料・給付費・見込量の数値は一切変えていない。**
"""
import copy
import os
import re
import shutil
import sys
import zipfile

import docx
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from fix_soan_v111 import set_el, set_tc                     # noqa: E402

SRC = "01_第10期_最新版成果品/川崎町_計画書素案_v2.14_見込量補完版.docx"
DST = "01_第10期_最新版成果品/川崎町_計画書素案_v2.15_実績是正版.docx"

# 是正1・2　表の行の差し替え（古い行の並び → 新しい行の並び）
#   ⚠ セルの文字をそのまま突き合わせるため、表記を1字も変えずに書く。
HYO_NINTEI = [
    (["令和6年度（実績）", "27", "73", "69", "126", "126", "89", "52",
      "562人"],
     ["令和6年度（実績）", "28", "75", "77", "119", "128", "85", "53",
      "565人"]),
    (["令和7年度（実績）", "33", "85", "70", "116", "131", "84", "49",
      "568人"],
     ["令和7年度（実績）", "36", "90", "61", "103", "130", "96", "50",
      "566人"]),
]
HYO_RITSU = [
    (["令和6年度", "3,270人", "554人", "16.9%", "実績"],
     ["令和6年度", "3,261人", "557人", "17.1%", "実績"]),
    (["令和7年度", "3,285人", "560人", "17.0%", "実績"],
     ["令和7年度", "3,240人", "561人", "17.3%", "実績"]),
    (["令和8年度", "3,257人", "562人", "17.3%", "実績見込み"],
     ["令和8年度", "3,229人", "562人", "17.4%", "実績見込み"]),
]

# 是正3　本文の差し替え
HONBUN = [
    # 1-6
    ("第1号被保険者数が3,270人（令和6年度）と小規模であることから",
     "第1号被保険者数が3,240人（令和7年度末）と小規模であることから"),
    # 2-3 認定者数の推移
    ("要介護（支援）認定者数は、令和6年度562人から令和7年度568人へ6人増加し、"
     "令和8年度は567人とほぼ横ばいで推移しています。"
     "第10期（令和9〜11年度）は576〜581人と見込まれ、"
     "令和7年度からの3か年でおよそ2％の増加となります。"
     "要介護度別に見ると、要支援2と要介護4が増加し、要介護2が減少する傾向に"
     "あり、認定者数の総数以上に重度化が進んでいます。",
     "要介護（支援）認定者数は、令和6年度565人から令和7年度566人へ1人増と"
     "ほぼ横ばいで、令和8年度は567人を見込んでいます。"
     "第10期（令和9〜11年度）は576〜581人と見込まれ、"
     "令和7年度からの3か年でおよそ2％の増加となります。"
     "⚠ 要介護度別に見ると、総数がほぼ横ばいである一方で、"
     "令和6年度から令和7年度にかけて要支援1・2が23人増え、"
     "要介護1・2が32人減り、要介護4が11人増えています。"
     "軽度と重度が増え、中間が減る二極化が起きており、"
     "総数の動きだけでは需要の変化を読めません。"),
    # 2-3 中長期
    ("要支援2と要介護4が増え、要介護2が減るという重度化の傾向がみられます。",
     "要支援1・2と要介護4が増え、要介護1・2が減るという二極化の傾向が"
     "みられます。"),
    # 2-3 認定率
    ("第1号被保険者の認定率は、令和6年度16.9％から令和11年度18.1％へ"
     "1.2ポイント上昇する見込みです。これは認定者数が増える一方で"
     "第1号被保険者数が3,270人から3,177人へ減少するためで、",
     "第1号被保険者の認定率は、令和6年度17.1％から令和11年度18.1％へ"
     "1.0ポイント上昇する見込みです。これは認定者数が増える一方で"
     "第1号被保険者数が3,261人から3,177人へ減少するためで、"),
    # 2-3 ●
    ("● 認定者数は3か年で約2％の増加にとどまるが、第1号被保険者数の減少に"
     "より認定率は1.2ポイント上昇する。",
     "● 認定者数は3か年で約2％の増加にとどまるが、第1号被保険者数の減少に"
     "より認定率は1.0ポイント上昇する。"),
    # 2-3 出典
    ("出典：令和6〜8年度：介護保険事業状況報告／令和9〜11年度：本計画の算定値。"
     "認定率＝第1号被保険者の認定者数÷第1号被保険者数",
     "出典：令和6・7年度は介護保険事業状況報告（年報・各年度末現在）、"
     "令和8年度は町提供実績データによる見込み、令和9〜11年度は本計画の算定値。"
     "認定率＝第1号被保険者の認定者数÷第1号被保険者数"),
    # 2-1 世帯数
    ("令和3年度末時点で、第1号被保険者のいる世帯数は2,181世帯です。",
     "第1号被保険者のいる世帯数は、令和7年度末で2,193世帯です"
     "（令和6年度末2,198世帯。介護保険事業状況報告の年報による）。"),
]

# 2-3 の表の出典・注に加える1文
CHU_23 = (
    "⚠ 令和6年度及び令和7年度の実績は介護保険事業状況報告（年報）の"
    "各年度末現在の値です。令和8年度以降は地域包括ケア「見える化」システムの"
    "推計（令和8年度は町提供実績データによる見込み）によるため、"
    "令和7年度と令和8年度の間で出所が替わります。")


def main():
    if not os.path.exists(SRC):
        raise SystemExit("入力がない：" + SRC)
    shutil.copy(SRC, DST)
    doc = docx.Document(DST)
    body = doc.element.body

    # ── 雛形
    tpl = None
    for el in body.iterchildren():
        if el.tag == qn("w:p"):
            s = "".join(x.text or "" for x in el.iter(qn("w:t"))).strip()
            if s.startswith("⚠ 宮城県自身の評価"):
                tpl = copy.deepcopy(el)
                break
    if tpl is None:
        raise SystemExit("囲みの雛形が見つからない")

    # ══════ 是正1・2　表の行を差し替える
    n_nintei = n_ritsu = 0
    hyo_nintei_el = None
    for el in body.iterchildren():
        if el.tag != qn("w:tbl"):
            continue
        tb = docx.table.Table(el, doc)
        for r in tb.rows:
            cs = [c.text.replace("\n", "").strip() for c in r.cells]
            for furui, atarashii in HYO_NINTEI:
                if cs[:len(furui)] == furui:
                    for i, v in enumerate(atarashii):
                        set_tc(r.cells[i]._tc, v)
                    n_nintei += 1
                    hyo_nintei_el = el
            for furui, atarashii in HYO_RITSU:
                if cs[:len(furui)] == furui:
                    for i, v in enumerate(atarashii):
                        set_tc(r.cells[i]._tc, v)
                    n_ritsu += 1
    if n_nintei != len(HYO_NINTEI):
        raise SystemExit(f"要介護度別の行の差し替えが{n_nintei}件"
                         f"（{len(HYO_NINTEI)}件のはず）")
    if n_ritsu != len(HYO_RITSU):
        raise SystemExit(f"認定率の行の差し替えが{n_ritsu}件"
                         f"（{len(HYO_RITSU)}件のはず）")

    # ══════ 是正3・4　本文の差し替え
    n_hon = 0
    nokori = []
    for furui, atarashii in HONBUN:
        atta = False
        for el in body.iterchildren():
            if el.tag != qn("w:p"):
                continue
            s = "".join(x.text or "" for x in el.iter(qn("w:t")))
            if furui in s:
                set_el(el, s.replace(furui, atarashii))
                n_hon += 1
                atta = True
                break
        if not atta:
            nokori.append(furui[:34])
    if nokori:
        raise SystemExit("本文が見つからない：" + "／".join(nokori))

    # ══════ 2-3 の表の後ろに出所の注を置く
    if hyo_nintei_el is not None:
        e = copy.deepcopy(tpl)
        set_el(e, CHU_23)
        hyo_nintei_el.addnext(e)

    # ══════ 版表記
    n_ver = 0
    for el in body.iter(qn("w:t")):
        if el.text and "Ver.2.14" in el.text:
            el.text = el.text.replace("Ver.2.14", "Ver.2.15")
            n_ver += 1
    if n_ver == 0:
        raise SystemExit("本文に Ver.2.14 の表記が見つからない")
    doc.save(DST)
    naoshi_header(DST)

    # ══════════════════════════ 自己点検
    d2 = docx.Document(DST)
    d0 = docx.Document(SRC)
    zen = "".join(n.text or "" for n in d2.element.body.iter(qn("w:t")))
    ng = []
    for v in ("3,270人", "3,285人", "3,257人", "568人", "562人へ", "16.9％",
              "1.2ポイント", "2,181世帯"):
        if v in zen:
            ng.append(f"古い数値が残っている：{v}")
    for v in ("3,261人", "3,240人", "3,229人", "566人", "565人", "17.1％",
              "1.0ポイント", "2,193世帯"):
        if v not in zen:
            ng.append(f"新しい数値が入っていない：{v}")
    # 保険料・給付費は変えていないこと
    for v in ("6,822", "998,222", "3,313,925", "152,500", "9,743.1",
              "3,231", "3,203", "3,177"):
        if v not in zen:
            ng.append(f"主な数値が消えている：{v}")
    if len(d2.tables) != len(d0.tables):
        ng.append(f"表の数が変わっている（{len(d0.tables)}→{len(d2.tables)}）")
    if len(d2.paragraphs) != len(d0.paragraphs) + 1:
        ng.append(f"段落が{len(d0.paragraphs)}→{len(d2.paragraphs)}"
                  "（1つだけ増えるはず）")
    for w in ("**", "__"):
        if w in zen:
            ng.append(f"強調記号が残っている：{w}")
    z = zipfile.ZipFile(DST)
    hdr = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>",
                             z.read("word/header1.xml").decode("utf-8")))
    if "Ver.2.15" not in hdr:
        ng.append(f"ヘッダーの版表記が更新されていない：{hdr}")

    print("計画素案 Ver.2.15（実績是正版）")
    print("保存：", DST)
    print(f"  段落 {len(d2.paragraphs)}（Ver.2.14 は {len(d0.paragraphs)}）"
          f"／表 {len(d2.tables)}")
    print(f"  是正1・2　表の行の差し替え　要介護度別 {n_nintei}行"
          f"／認定率 {n_ritsu}行")
    print(f"  是正3・4　本文の差し替え　{n_hon}か所（出所の注を1か所）")
    print("  ── 自己点検")
    if ng:
        for m in ng:
            print("   ×", m)
        sys.exit(1)
    print("   ○ 認定者数・第1号被保険者数が年報の値になった")
    print("   ○ 認定率が 認定者数÷第1号被保険者数 で再現する")
    print("   ○ 保険料・給付費・見込量の数値は変えていない")
    print("  ⚠ 目次の頁番号は fill_toc_pages_v200.py で作り直してください。")
    print("  ⚠ 図2-4（要介護認定者数と認定率の推移）は data_zuhyo.py の値を"
          "直して作り直す必要があります。")


def naoshi_header(path):
    z = zipfile.ZipFile(path)
    names = z.namelist()
    data = {n: z.read(n) for n in names}
    z.close()
    n_fix = 0
    for n in names:
        if not re.match(r"word/(header|footer)\d*\.xml", n):
            continue
        s = data[n].decode("utf-8")
        s2 = re.sub(r"Ver\.2\.14", "Ver.2.15", s)
        if s2 != s:
            data[n] = s2.encode("utf-8")
            n_fix += 1
    if n_fix == 0:
        raise SystemExit("ヘッダーの版表記（Ver.2.14）が見つからない")
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, data[n])
    shutil.move(tmp, path)


if __name__ == "__main__":
    main()
