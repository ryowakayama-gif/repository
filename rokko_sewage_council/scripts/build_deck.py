# -*- coding: utf-8 -*-
"""第2回審議会資料：全パターン（①②④）シミュレーション版の生成."""
import copy
import json
import pathlib
import re
import shutil
import tempfile
import zipfile
from pptx import Presentation
from pptx.util import Inches

HERE = pathlib.Path(__file__).resolve().parent.parent
BASE_PPTX = HERE / "source" / "六戸町_第2回審議会資料_案.pptx"     # 第1回時点の原稿（22ページ）
OUT_PPTX = HERE / "output" / "六戸町_第2回審議会資料.pptx"
M = json.loads((HERE / "data" / "metrics.json").read_text(encoding="utf-8"))
YI = {"R7": 0, "R8": 1, "R9": 2, "R10": 3, "R11": 4, "R12": 5}

# ---------------------------------------------------------------- 料金テーブル
BLOCKS = [("基本(0〜10㎥)", None), ("11〜20㎥", 10), ("21〜30㎥", 10), ("31〜40㎥", 10),
          ("41〜50㎥", 10), ("51〜70㎥", 20), ("71〜100㎥", 30), ("101〜150㎥", 50), ("151㎥〜", None)]
CAPS = [20, 30, 40, 50, 70, 100, 150]   # 従量区分の上限（基本10㎥の次から）

RATES = {
    "現行":   [1000, 120, 120, 130, 130, 140, 140, 140, 160],
    "①最終": [1500, 150, 155, 160, 165, 175, 185, 205, 225],
    "②最終": [1200, 180, 185, 190, 195, 205, 215, 235, 255],
    "④最終": [1400, 160, 165, 170, 175, 185, 195, 215, 235],
    # 2か年改定 R8中間（= ROUND((現行+最終)/2, 0)：モデル算定値）
    "①中間": [1250, 135, 138, 145, 148, 158, 163, 173, 193],
    "②中間": [1100, 150, 153, 160, 163, 173, 178, 188, 208],
    "④中間": [1200, 140, 143, 150, 153, 163, 168, 178, 198],
    # 3か年改定
    "①第1段": [1150, 130, 130, 140, 140, 150, 155, 160, 180],
    "①第2段": [1300, 140, 140, 150, 150, 160, 170, 180, 200],
    "②第1段": [1050, 140, 140, 150, 150, 160, 165, 170, 190],
    "②第2段": [1100, 160, 160, 170, 170, 180, 190, 200, 220],
    "④第1段": [1100, 130, 135, 140, 145, 155, 155, 165, 185],
    "④第2段": [1200, 140, 150, 150, 160, 170, 170, 190, 210],
}

PAT = {"①": "標準型", "②": "家庭軽減型", "④": "段階累進型"}


def fee(rate, vol):
    """水量 vol ㎥ の月額使用料（税抜・円）を段階累進で算定."""
    total = rate[0]
    prev = 10
    for i, cap in enumerate(CAPS, start=1):
        if vol <= prev:
            break
        total += (min(vol, cap) - prev) * rate[i]
        prev = cap
    if vol > 150:
        total += (vol - 150) * rate[8]
    return total


def tax(v):
    return int(round(v * 1.1))


def yen(v):
    return f"{v:,}円"


# ---------------------------------------------------------------- テキスト設定
def set_text(shape, text):
    """段落構造を保ったままテキストを差し替える（書式は既存ランを継承）."""
    tf = shape.text_frame
    lines = text.split("\n")
    paras = tf.paragraphs
    while len(paras) < len(lines):                       # 段落が足りなければ複製
        new = copy.deepcopy(paras[-1]._p)
        paras[-1]._p.addnext(new)
        paras = tf.paragraphs
    for i, line in enumerate(lines):
        p = paras[i]
        if not p.runs:                                   # ラン不在なら直前段落から借用
            src = next((q for q in paras if q.runs), None)
            if src is None:
                p.text = line
                continue
            p._p.append(copy.deepcopy(src.runs[0]._r))
        p.runs[0].text = line
        for r in p.runs[1:]:
            r._r.getparent().remove(r._r)
    for p in list(paras[len(lines):]):                   # 余剰段落を削除
        p._p.getparent().remove(p._p)


def sh(slide):
    """shape名 -> shape（同名は最初のもの）と、同名リストの両方を返す."""
    one, many = {}, {}
    for s in slide.shapes:
        one.setdefault(s.name, s)
        many.setdefault(s.name, []).append(s)
    return one, many


def put(slide, mapping):
    one, _ = sh(slide)
    for name, val in mapping.items():
        if name not in one:
            raise KeyError(f"shape {name!r} not found")
        set_text(one[name], val)


def put_seq(slide, name, values):
    """同名 shape が複数ある場合に順番で流し込む."""
    _, many = sh(slide)
    tgts = many[name]
    assert len(tgts) == len(values), f"{name}: {len(tgts)} shapes vs {len(values)} values"
    for t, v in zip(tgts, values):
        set_text(t, v)


# ---------------------------------------------------------------- 指標ヘルパ
def pct(biz, metric, key, year):
    return f"{M[biz][metric][key][YI[year]] * 100:.1f}%"



# ---------------------------------------------------------------- 構成の組み替え
# 原稿はパターン②のみを扱う22ページ。パターン①/④用にスライドを複製し、章立てを組み直す。
DUPLICATE = [6, 6, 11, 11, 15, 15, 16, 19, 19]      # 原稿のページ番号（複製元）
ORDER = [1, 2, 3, 4, 5, 23, 6, 24, 7, 8, 9, 10, 25, 11, 26, 12, 13, 14,
         27, 15, 28, 16, 29, 17, 18, 30, 19, 31, 20, 21, 22]
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
SLIDE_CT = ("application/vnd.openxmlformats-officedocument."
            "presentationml.slide+xml")


def duplicate_slide(root, src_no):
    """ppt/slides/slideN.xml を複製し、パッケージへの登録まで行って新しい番号を返す."""
    slides = sorted(int(m.group(1)) for m in
                    (re.fullmatch(r"slide(\d+)\.xml", f.name)
                     for f in (root / "ppt/slides").glob("slide*.xml")) if m)
    new_no = max(slides) + 1
    shutil.copy(root / f"ppt/slides/slide{src_no}.xml",
                root / f"ppt/slides/slide{new_no}.xml")
    src_rels = root / f"ppt/slides/_rels/slide{src_no}.xml.rels"
    if src_rels.exists():
        rels = src_rels.read_text(encoding="utf-8")
        rels = re.sub(r'<Relationship[^>]*notesSlide[^>]*/>', "", rels)   # ノートは引き継がない
        (root / f"ppt/slides/_rels/slide{new_no}.xml.rels").write_text(rels, encoding="utf-8")

    ct = root / "[Content_Types].xml"
    txt = ct.read_text(encoding="utf-8")
    entry = (f'<Override PartName="/ppt/slides/slide{new_no}.xml" '
             f'ContentType="{SLIDE_CT}"/>')
    ct.write_text(txt.replace("</Types>", entry + "</Types>"), encoding="utf-8")

    rp = root / "ppt/_rels/presentation.xml.rels"
    txt = rp.read_text(encoding="utf-8")
    next_id = max(int(n) for n in re.findall(r'Id="rId(\d+)"', txt)) + 1
    rid = "rId%d" % next_id
    entry = ('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/'
             'officeDocument/2006/relationships/slide" '
             'Target="slides/slide%d.xml"/>' % (rid, new_no))
    rp.write_text(txt.replace("</Relationships>", entry + "</Relationships>"), encoding="utf-8")
    return new_no


def restructure(base_pptx, work_dir):
    """複製と並べ替えを済ませた中間 pptx のパスを返す."""
    root = work_dir / "unpacked"
    with zipfile.ZipFile(base_pptx) as z:
        z.extractall(root)
    for src_no in DUPLICATE:
        duplicate_slide(root, src_no)

    rels = (root / "ppt/_rels/presentation.xml.rels").read_text(encoding="utf-8")
    rid_of = {int(no): rid for rid, no in
              re.findall(r'Id="(rId\d+)"[^>]*Target="slides/slide(\d+)\.xml"', rels)}
    pres_path = root / "ppt/presentation.xml"
    pres = pres_path.read_text(encoding="utf-8")
    lst = re.search(r"<p:sldIdLst>.*?</p:sldIdLst>", pres, re.S).group(0)
    new = "<p:sldIdLst>" + "".join(
        f'<p:sldId id="{256 + i}" r:id="{rid_of[n]}"/>' for i, n in enumerate(ORDER)
    ) + "</p:sldIdLst>"
    pres_path.write_text(pres.replace(lst, new), encoding="utf-8")

    staged = work_dir / "staged.pptx"
    with zipfile.ZipFile(staged, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(root.rglob("*")):
            if f.is_file():
                z.write(f, f.relative_to(root).as_posix())
    return staged


_tmp = tempfile.TemporaryDirectory()
prs = Presentation(restructure(BASE_PPTX, pathlib.Path(_tmp.name)))
sl = list(prs.slides)




def S_(n):
    return sl[n - 1]


SRC_NOTE = "出典：六戸町_公共／農集_使用料改定.xlsx「パターン別比較」（社人研人口推計ベース）"

# ================================================================ 1. 表紙
put(S_(1), {
    "org_name": "六戸町 下水道使用料審議委員会",
    "title1": "第２回",
    "title2": "審議会 資料",
    "date_txt": "令和８年８月19日（水）　18:30〜",
    "loc_txt": "六戸町役場 ２階 小会議室",
})
# 表紙は元テキストより長いため、装飾（右下の四角 x=9.68"）に掛からない範囲で幅を広げる
_one, _ = sh(S_(1))
for _n, _w in (("org_name", 7.0), ("date_txt", 7.3), ("loc_txt", 7.3)):
    _one[_n].width = Inches(_w)

# ================================================================ 2. 目次
put_seq(S_(2), "item_desc", [
    "第1回審議内容のまとめ・委員意見への対応方針",
    "使用料改定案の検討（パターン①②④の収支シミュレーション・妥当性評価）",
    "段階的値上げの検討（2か年vs3か年・パターン別段階単価一覧）",
    "財政健全化目標との整合性（経常収支比率・経費回収率・目標値設定）",
    "住民への影響・説明対応（パターン別影響額・説明会計画）",
    "料金体系見直し・答申案策定・意見集約",
])

# ================================================================ 3. 前回振り返り
put_seq(S_(3), "pb", [
    "公共の経費回収率47%・農集35%と低水準。使用料改定が必要。経営持続のための抜本的対策を確認した。",
    "目標①経常収支比率100%以上、目標②経費回収率47%以上（段階的向上）を目標として設定した。",
    "パターン①②④を提示。20㎥で税抜3,000円・税込3,300円を達成。今回は3パターンすべての収支シミュレーションを提示する。",
    "5年に1回の使用料改定検証が交付要件。未達成の場合は社会資本整備総合交付金の対象外となる。",
    "2か年（R8中間→R9最終）および3か年（R8第1段→R9第2段→R10最終）の段階単価を提示・確認した。",
])

# ================================================================ 5. 01-1 評価軸
put(S_(5), {
    "r41": "○ 中程度", "r42": "◎ 高い", "r43": "○ 中程度",
    "concl": "▶ 推奨案：パターン②（家庭軽減型）　少量使用者に配慮しつつ、3案中で最も使用料収入が大きく財政健全化への効果が高い。",
})

# ---------------------------------------------------------------- 章扉の説明文
put(S_(4), {"desc": "本章では、パターン①②④それぞれの収支シミュレーション結果を示し、"
                    "推奨案の選定理由と改定率の妥当性について整理します。"})
put(S_(18), {"desc": "本章では、パターン①②④それぞれの経常収支比率・経費回収率の見込みを確認し、"
                     "財政健全化目標値の設定について審議いただきます。"})

# ================================================================ 6/7/8. 01-2〜4 パターン別 収支シミュレーション
SIM = {"①": 6, "②": 7, "④": 8}
SUB = {"①": "01-2", "②": "01-3", "④": "01-4"}
for p, idx in SIM.items():
    k2, k3 = p + "2", p + "3"
    s = S_(idx)
    d = {
        "hdr_txt": f"使用料改定案の検討　{SUB[p]}　収支シミュレーション（パターン{p}採用時）",
        "ttl": f"パターン{p}（{PAT[p]}）採用時の財政収支見込み",
        "lh0": "公共下水道", "lh1": "農業集落排水",
        "g1h0": "目標① 経常収支比率（100%以上）　※（ ）内は3か年改定",
        "g1h1": "目標① 経常収支比率（100%以上）　※（ ）内は3か年改定",
        "g2h0": "目標② 経費回収率（47%以上）　※（ ）内は3か年改定",
        "g2h1": "目標② 経費回収率（47%以上）　※（ ）内は3か年改定",
    }
    for i, biz in enumerate(["公共", "農集"]):
        d[f"cr{i}0"] = f"  R7（現行）：{pct(biz,'経常収支比率','現行','R7')}"
        d[f"cr{i}1"] = f"  R8：{pct(biz,'経常収支比率',k2,'R8')}　（{pct(biz,'経常収支比率',k3,'R8')}）"
        d[f"cr{i}2"] = f"  R9：{pct(biz,'経常収支比率',k2,'R9')}　（{pct(biz,'経常収支比率',k3,'R9')}）"
        d[f"er{i}0"] = f"  R7（現行）：{pct(biz,'経費回収率','現行','R7')}"
        d[f"er{i}1"] = f"  R8：{pct(biz,'経費回収率',k2,'R8')}　（{pct(biz,'経費回収率',k3,'R8')}）"
        d[f"er{i}2"] = f"  R9：{pct(biz,'経費回収率',k2,'R9')}　（{pct(biz,'経費回収率',k3,'R9')}）"
    ko_cr10, ko_er10 = pct("公共", "経常収支比率", k2, "R10"), pct("公共", "経費回収率", k2, "R10")
    no_cr10, no_er10 = pct("農集", "経常収支比率", k2, "R10"), pct("農集", "経費回収率", k2, "R10")
    d["concl6"] = (
        f"パターン{p}（{PAT[p]}）を2か年改定で採用した場合、公共下水道は経常収支比率がR9に{pct('公共','経常収支比率',k2,'R9')}"
        f"（R10以降 {ko_cr10}）となり、目標①の100%以上を達成する見込みです。経費回収率は現行34.4%から"
        f"R9 {pct('公共','経費回収率',k2,'R9')}・R10 {ko_er10}へ改善します。\n"
        f"農業集落排水は経費回収率がR9 {pct('農集','経費回収率',k2,'R9')}・R10 {no_er10}へ改善する一方、"
        f"経常収支比率はR10で{no_cr10}にとどまり、引き続き他会計繰入等による補填が必要となる見込みです。\n"
        f"3か年改定とした場合は最終単価の適用がR10となるため、目標達成が1年遅れます（上表（ ）内）。\n"
        f"{SRC_NOTE}")
    put(s, d)

# ================================================================ 9. 01-5 収支見込み比較
s = S_(9)
d = {"hb": "  使用料改定案の検討　01-5　収支見込み比較（パターン①②④・2か年改定）",
     "ttl": "パターン別　財政収支見込み比較（公共下水道・農業集落排水）　2か年改定ベース"}
for i, p in enumerate(["①", "②", "④"]):
    k = p + "2"
    d[f"ph{i}"] = f"パターン{p}（{PAT[p]}）"
    d[f"ch{i}"] = "経常収支比率（目標100%以上）"
    d[f"eh{i}"] = "経費回収率（目標47%以上）"
    d[f"ach{i}"] = "経費回収率／経常収支比率"
    for j, y in enumerate(["R7", "R8", "R9"]):
        key = "現行" if y == "R7" else k
        d[f"cpr{i}{j}"] = f"  {y}：{pct('公共','経常収支比率',key,y)}"
        d[f"epr{i}{j}"] = f"  {y}：{pct('公共','経費回収率',key,y)}"
        d[f"ar{i}{j}"] = f"  {y}：{pct('農集','経費回収率',key,y)} / {pct('農集','経常収支比率',key,y)}"
d["note"] = ("▶ 使用料収入は②＞④＞①の順に大きく、公共下水道の経費回収率はR9で①43.1%・④44.1%・②46.1%。"
             "経常収支比率は3案ともR9に100%以上を達成する見込みです。農業集落排水は経費回収率が大きく改善する一方、"
             "経常収支比率は3案とも100%未満（R9で84〜86%）にとどまります。")
put(s, d)

# ================================================================ 10. 01-6 改定率の妥当性
put(S_(10), {
    "hdr_txt": "使用料改定案の検討　01-6　改定率の妥当性評価",
    "lh7": "青森県内自治体との比較（20㎥・税込）",
    "cn0": "六戸町（改定前）", "cp0": "2,420円", "cr0": "—",
    "cn1": "六戸町（改定後・共通）", "cp1": "3,300円", "cr1": "＋36.4%",
    "cn2": "八戸市", "cp2": "3,383円", "cr2": "—",
    "cn3": "東北町", "cp3": "3,300円", "cr3": "—",
    "cn4": "県内平均（公共）", "cp4": "3,063円", "cr4": "—",
    "bi3": "現行比の改定率", "bv3": "＋36.4%",
    "bi4": "県内平均との差", "bv4": "＋237円",
    "eval7": "▶ 改定後の水準（税込3,300円）は青森県内平均3,063円と概ね同水準。1日あたりの負担増は約29円で、住民への影響は許容範囲内と評価されます（県内比較は令和6年3月31日現在・第1回資料より）。",
})

# ================================================================ 12. 02-1 論点整理
put(S_(12), {
    "m00": "  • 1回の値上げ幅が大きい（基本料+100〜250円）",
    "m10": "  • 目標達成がR10まで遅延（交付金要件の充足も遅延）",
})

# ================================================================ 13/14/15. 02-2〜4 パターン別 段階単価
STG = {"①": 13, "②": 14, "④": 15}
SUB2 = {"①": "02-2", "②": "02-3", "④": "02-4"}
for p, idx in STG.items():
    s = S_(idx)
    lead = "推奨案（パターン②）" if p == "②" else f"パターン{p}（{PAT[p]}）"
    d = {"hdr_txt": f"段階的値上げの検討　{SUB2[p]}　段階単価一覧（パターン{p}・税抜き・円）",
         "ttl": f"{lead}の段階単価一覧"}
    cols = ["現行", p + "中間", p + "最終", p + "第1段", p + "第2段", p + "最終"]
    for r, (lab, _) in enumerate(BLOCKS):
        d[f"d{r}0"] = lab
        for c, key in enumerate(cols, start=1):
            d[f"d{r}{c}"] = f"{RATES[key][r]:,}"
    put(s, d)

# ================================================================ 16. 02-5 段階単価比較
s = S_(16)
d = {"hb": "  段階的値上げの検討　02-5　段階単価比較（パターン①②④・2か年）"}
cols = ["現行", "①中間", "①最終", "②中間", "②最終", "④中間", "④最終"]
for r, (lab, _) in enumerate(BLOCKS):
    d[f"d{r}0"] = lab
    for c, key in enumerate(cols, start=1):
        d[f"d{r}{c}"] = f"{RATES[key][r]:,}"
d["note"] = "※ 20㎥での最終月額（税抜）は3案とも3,000円で共通。R8中間単価は（現行＋最終）÷2の四捨五入。"
put(s, d)

# ================================================================ 17. 02-6 ロードマップ
put(S_(17), {
    "hdr_txt": "段階的値上げの検討　02-6　改定ロードマップ案",
    "ttl": "改定ロードマップ案（パターン①②④共通）",
})

# ================================================================ 19/20/21. 03-1〜3 パターン別 経常収支比率
CR = {"①": 19, "②": 20, "④": 21}
SUB3 = {"①": "03-1", "②": "03-2", "④": "03-3"}
for p, idx in CR.items():
    k = p + "2"
    s = S_(idx)
    d = {"hdr_txt": f"財政健全化目標との整合性　{SUB3[p]}　経常収支比率の見込み（パターン{p}）",
         "ttl": f"経常収支比率の見込み（パターン{p}・{PAT[p]}／2か年改定）"}
    for i, biz in enumerate(["公共", "農集"]):
        d[f"rv{i}0"] = f"R7（現行）\n経常収支比率：{pct(biz,'経常収支比率','現行','R7')}\n（改定なし）"
        d[f"rv{i}1"] = f"R8（中間単価）\n経常収支比率：{pct(biz,'経常収支比率',k,'R8')}\n（改定開始）"
        d[f"rv{i}2"] = f"R9（最終単価）\n経常収支比率：{pct(biz,'経常収支比率',k,'R9')}\n（改定完了）"
        d[f"rv{i}3"] = f"R10〜\n経常収支比率：{pct(biz,'経常収支比率',k,'R10')}\n（安定期）"
    ok = M["公共"]["経常収支比率"][k][YI["R9"]] >= 1.0
    d["nt13"] = (f"✔ パターン{p}採用により、公共下水道はR9に{pct('公共','経常収支比率',k,'R9')}"
                 f"{'と目標の100%以上を達成' if ok else 'となり目標100%に届かず'}する見込みです。"
                 f"農業集落排水はR10で{pct('農集','経常収支比率',k,'R10')}にとどまり、"
                 f"100%達成には他会計繰入等による補填が引き続き必要です。　{SRC_NOTE}")
    put(s, d)

# ================================================================ 22/23. 03-4/5 経費回収率
TGT = {"R7": "—", "R8": "—", "R9": "47%以上", "R10": "50%以上", "R12": "55%以上"}
for idx, biz, sub in [(22, "公共", "03-4"), (23, "農集", "03-5")]:
    s = S_(idx)
    label = "公共下水道" if biz == "公共" else "農業集落排水"
    d = {"hdr_txt": f"財政健全化目標との整合性　{sub}　経費回収率の見込み（{label}）",
         "ttl": f"経費回収率の見込みと目標値の設定（案）　{label}"}
    put_seq(s, "th14", ["年度", "改定なし", "パターン①", "パターン②", "パターン④", "目標水準（案）"])
    for r, y in enumerate(["R7", "R8", "R9", "R10", "R12"]):
        d[f"d14{r}0"] = y
        d[f"d14{r}1"] = pct(biz, "経費回収率", "現行", y)
        for c, p in enumerate(["①", "②", "④"], start=2):
            key = "現行" if y == "R7" else p + "2"
            d[f"d14{r}{c}"] = pct(biz, "経費回収率", key, y)
        d[f"d14{r}5"] = TGT[y]
    if biz == "公共":
        d["nt14"] = ("▶ 審議事項：公共下水道は今回改定のみではR9で43〜46%にとどまり、目標水準47%に到達しません。"
                     "R10以降の段階的適正化（次期改定）とあわせた目標設定について審議をお願いします。")
    else:
        d["nt14"] = ("▶ 審議事項：農業集落排水はパターン②でR9に54.2%となり目標水準47%を達成しますが、"
                     "R11以降は有収水量の減少により低下します。次期改定（R12）での再設定を前提とした目標設定を提案します。")
    put(s, d)

# ================================================================ 24. 03-6 全パターン指標比較
s = S_(24)
d = {"hb": "  財政健全化目標との整合性　03-6　全パターン指標比較",
     "ttl": "パターン別　経常収支比率・経費回収率の見込み比較（2か年改定）"}
for r, y in enumerate(["R7", "R8", "R9"]):
    d[f"cr{r}0"] = y
    d[f"er{r}0"] = y
    for c, (biz, p) in enumerate([(b, p) for b in ("公共", "農集") for p in ("①", "②", "④")], start=1):
        key = "現行" if y == "R7" else p + "2"
        d[f"cr{r}{c}"] = pct(biz, "経常収支比率", key, y)
        d[f"er{r}{c}"] = pct(biz, "経費回収率", key, y)
d["note"] = ("▶ 使用料収入が最も大きいパターン②が両指標とも最良（R9：公共 経常101.1%・回収46.1%／農集 回収54.2%）。"
             "④・①がこれに続きます。公共下水道は3案ともR9に経常収支比率100%以上を達成する一方、"
             "農業集落排水は3案とも100%未満にとどまります。")
put(s, d)

# ================================================================ 26/27/28. 04-1〜3 パターン別 影響額
IMP = {"①": 26, "②": 27, "④": 28}
SUB4 = {"①": "04-1", "②": "04-2", "④": "04-3"}
HOME = [(10, "10㎥（1〜2人世帯）"), (20, "20㎥（3人世帯）　★"), (30, "30㎥（3〜4人世帯）"), (50, "50㎥（4人以上）")]
BIZC = [(50, "50㎥（小規模）"), (100, "100㎥（中規模）"), (200, "200㎥（大規模）"), (500, "500㎥（大型施設）")]
for p, idx in IMP.items():
    s = S_(idx)
    rate = RATES[p + "最終"]
    d = {"hdr_txt": f"住民への影響・説明対応　{SUB4[p]}　家庭・事業所別影響額（パターン{p}・税込）",
         "ttl": f"家庭・事業所別　月額影響額（パターン{p}・{PAT[p]}・税込）"}
    for pre, rows in (("hr", HOME), ("br", BIZC)):
        for r, (vol, lab) in enumerate(rows):
            cur, new = tax(fee(RATES["現行"], vol)), tax(fee(rate, vol))
            d[f"{pre}{r}0"] = lab
            d[f"{pre}{r}1"] = yen(cur)
            d[f"{pre}{r}2"] = yen(new)
            d[f"{pre}{r}3"] = f"＋{new - cur:,}円"
    up = tax(fee(rate, 20)) - tax(fee(RATES["現行"], 20))
    d["nt16"] = (f"★ 20㎥（一般家庭の標準的な使用量）での月額は税込{tax(fee(rate,20)):,}円"
                 f"（税抜{fee(rate,20):,}円）。増加額は月額＋{up:,}円、年間＋{up*12:,}円、"
                 f"1日あたり約{round(up/30.4):,}円です。※各区分の単価を段階累進で積み上げて算定。")
    put(s, d)

# ================================================================ 29. 04-4 使用水量別月額
s = S_(29)
d = {"hb": "  住民への影響・説明対応　04-4　使用水量別月額使用料比較"}
for r, vol in enumerate([10, 20, 30, 40]):
    cur = tax(fee(RATES["現行"], vol))
    d[f"d{r}0"] = f"{vol}㎥"
    d[f"d{r}1"] = yen(cur)
    for c, p in enumerate(["①", "②", "④"]):
        v = fee(RATES[p + "最終"], vol)
        base = 2 + c * 3
        d[f"d{r}{base}"] = yen(v)
        d[f"d{r}{base+1}"] = yen(tax(v))
        d[f"d{r}{base+2}"] = f"＋{tax(v)-cur:,}円"
put(s, d)

# ================================================================ 30. 04-5 住民説明会
put(S_(30), {"hdr_txt": "住民への影響・説明対応　04-5　住民説明会の実施計画（案）"})

# ================================================================ レイアウト調整
# 差し替えたテキストが元より長い箇所があるため、見出し・注記の器を実寸に合わせて調整する
SLIDE_W = 13.333
NOTE_NAMES = ("nt13", "nt14", "nt16", "note", "eval7", "concl")
BG = ("bg1", "bg2", "hdr_bg", "hdr_line", "hdr_accent", "hb", "hl", "ha")


def _vlen(s):
    """全角換算の文字数（半角は0.5）."""
    return sum(0.5 if ord(c) < 0x2000 else 1.0 for c in s)


def _fontpt(shape, default=11.0):
    for para in shape.text_frame.paragraphs:
        for r in para.runs:
            if r.font.size:
                return r.font.size.pt
    return default


def fit_note(slide, shape):
    """注記の高さを行数に合わせ、上の要素とページ番号の間に収める."""
    pt = _fontpt(shape)
    cap = max(1.0, (shape.width.inches - 0.35) * 72 / pt)
    lines = max(1, -(-int(_vlen(shape.text_frame.text) * 100) // int(cap * 100)))
    need = lines * pt * 1.45 / 72 + 0.18
    h = max(shape.height.inches, need)
    above = [sp.top.inches + sp.height.inches for sp in slide.shapes
             if sp is not shape and sp.name not in BG and sp.top is not None
             and sp.top.inches + sp.height.inches <= shape.top.inches + 0.05]
    top = min(shape.top.inches, 7.40 - h)
    if above:
        top = max(top, max(above) + 0.08)
    shape.height, shape.top = Inches(h), Inches(top)


for _s in sl:
    _one, _many = sh(_s)
    for _n in ("ttl", "hdr_txt"):                       # 見出しは右方向に余白があるので広げる
        for _sp in _many.get(_n, []):
            _sp.width = Inches(max(_sp.width.inches, SLIDE_W - _sp.left.inches - 0.35))
    for _n in NOTE_NAMES:
        for _sp in _many.get(_n, []):
            fit_note(_s, _sp)

# 02-5 は元から最終行と注記が重なるため、データ行の送りを詰める
_one, _ = sh(S_(16))
for _r in range(9):
    for _c in range(8):
        _sp = _one[f"d{_r}{_c}"]
        _sp.top, _sp.height = Inches(2.13 + _r * 0.50), Inches(0.46)

# 段階単価表の見出しは2行（例「R9 最終／（2か年）」）で元の行高に収まらないため、
# 見出し行を高くし、区分バーを少し上へ逃がす（02-2〜4）
for _idx in (13, 14, 15):
    _one, _many = sh(S_(_idx))
    for _n in ("g2yr", "g3yr"):
        _one[_n].top = Inches(1.68)
    for _sp in _many["th"]:
        _sp.top, _sp.height = Inches(1.97), Inches(0.45)

# 住民負担指標テーブルがスライド右端をはみ出すため内側へ寄せる（01-6）
_one, _ = sh(S_(10))
for _r in range(5):
    _one[f"bi{_r}"].width = Inches(3.95)
    _one[f"bv{_r}"].left, _one[f"bv{_r}"].width = Inches(11.13), Inches(2.13)

# ================================================================ ページ番号
for i, s in enumerate(sl, start=1):
    one, _ = sh(s)
    if "pgnum" in one:
        set_text(one["pgnum"], str(i))

OUT_PPTX.parent.mkdir(parents=True, exist_ok=True)
prs.save(OUT_PPTX)
print(f"saved {OUT_PPTX} ({len(sl)} slides)")
