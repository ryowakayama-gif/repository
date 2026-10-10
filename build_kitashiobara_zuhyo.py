# -*- coding: utf-8 -*-
"""
北塩原村　図表の棚卸しと整理（令和8年10月10日）

出力: output/北塩原村_図表一覧.xlsx

何をするものか
  正本（他メンバー版への移植後）の図と表を機械で数え上げ、
  成果品として印刷できる状態かを確かめる。
  ・図は18か所（画像20点）。いずれも画像であり、数値を直せない
  ・表は109点（データ表75・注記ボックス34）
  ・図表番号は0件、図表目次もない
  ・出所（資料：）が付いている表は4点のみ

色の扱い（令和8年10月10日の指示）
  図は他メンバー版に合わせてカラーで進める。
  印刷については確認事項として村・他メンバーに残す。
  仕様書5の成果品①は「A4判・両面 約60頁・モノクロ・コピー・くるみ製本」で
  あり、仕様と図の作りが食い違っているためである。

  なお、モノクロにしたときに何が起きるかは測ってある。
  画像の画素からグレー値を計算すると、
  緑 RGB(46,158,107) と 青 RGB(61,134,198) はいずれもグレー値119で、
  完全に同じ濃さになる。
  白黒でコピーされると「身体障害者手帳」と「精神障害者保健福祉手帳」の棒を
  見分けられない。
  当方が作る図は、カラーのままでも同じ図の中の系列のグレー値を
  25以上離す（緑と青を並べて使わない）ことで、この形の事故を避ける。

機械で確かめていること
  ・図の位置（章・直前の文・直後の出所）と画像の大きさ
  ・図ごとの主な色とグレー値、同じ濃さになる組
  ・表の位置（章・節）・行列数・注記ボックスかどうか・出所の有無
  これらは正本の現物から毎回読み直すため、正本が変われば結果も変わる。
"""

import collections
import io
import os
import sys
import zipfile

import docx
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from kitashiobara_common import (  # noqa: E402
    COLORS, FONT, add_sheet, ensure_out_dir, style_header_row, write_row,
)

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_FILE = f"{REPO_ROOT}/output/北塩原村_図表一覧.xlsx"
HONPON = f"{REPO_ROOT}/output/北塩原村_計画素案_正本_移植後.docx"

NG = "FFC7CE"
OKC = "E2EFDA"
HI = "FFF2CC"

# モノクロにしたときに見分けられるとみなすグレー値の差
#   印刷の網点とコピーの劣化を見込み、25（256階調）を下限とする。
GRAY_MIN = 25


# ===========================================================================
# 1. 正本から図表を読み取る
# ===========================================================================
def read_honpon():
    doc = docx.Document(HONPON)
    z = zipfile.ZipFile(HONPON)
    items = []
    for ch in doc.element.body.iterchildren():
        tag = ch.tag.split("}")[-1]
        if tag == "p":
            p = Paragraph(ch, doc)
            embeds = [e.get(qn("r:embed"))
                      for e in ch.findall(".//" + qn("a:blip"))]
            items.append({"kind": "p", "text": p.text.strip(),
                          "style": p.style.name, "embeds": embeds})
        elif tag == "tbl":
            t = Table(ch, doc)
            items.append({"kind": "tbl", "obj": t})
    # 画像の関係 id → ファイル名
    rels = {}
    for r in doc.part.rels.values():
        if "image" in r.reltype:
            rels[r.rId] = r.target_part.partname.lstrip("/")
    return doc, z, items, rels


def gray(rgb):
    r, g, b = rgb
    return round(0.299 * r + 0.587 * g + 0.114 * b)


CHROMA_MIN = 14        # これ以上色みがあれば「色で分けている」とみなす


def hue(rgb):
    """色合い（0〜360度）。白と混ざっても色合いは変わらない。"""
    import colorsys
    r, g2, b = (v / 255 for v in rgb)
    return colorsys.rgb_to_hsv(r, g2, b)[0] * 360


def image_colors(z, path, top=8):
    """画像の主な色（背景の白と文字の黒を除く）とグレー値を返す。

    色み（彩度）も併せて返す。グレーだけで描かれた図は、
    もともとモノクロであり、印刷しても見え方が変わらない。
    """
    im = Image.open(io.BytesIO(z.read(path))).convert("RGB")
    size = im.size
    small = im.resize((min(im.width, 400), min(im.height, 400)))
    try:
        data = list(small.get_flattened_data())
    except AttributeError:
        data = list(small.getdata())
    c = collections.Counter(data)
    # 同じ系列でも、縁のぼかしで白と混ざった画素が別の色として数えられる。
    # 色合い（色相）が同じものは1つの系列とみなし、代表として
    # いちばん色が濃い（彩度の高い）ものを採る。
    groups = {}
    for rgb, n in c.most_common(200):
        if min(rgb) > 235 or max(rgb) < 70:
            continue          # 白地と黒文字は除く
        if n < 120:
            continue
        ch = max(rgb) - min(rgb)
        if ch < CHROMA_MIN:
            key = ("灰", round(gray(rgb) / 40))     # 目盛線・軸などの灰
        else:
            key = ("色", round(hue(rgb) / 15))      # 15度ごとにまとめる
        cur = groups.get(key)
        if cur is None or ch > cur[3]:
            groups[key] = (rgb, gray(rgb), n + (cur[2] if cur else 0), ch)
        else:
            groups[key] = (cur[0], cur[1], cur[2] + n, cur[3])
    out = sorted(groups.values(), key=lambda t: -t[2])[:top]
    return size, out


def collide(colors):
    """モノクロにすると見分けられなくなる色の組を返す。

    比べるのは色みのある色どうしだけ。
    目盛線や軸のグレーは系列ではないため、判定に混ぜない。
    """
    hue = [c for c in colors if c[3] >= CHROMA_MIN]
    bad = []
    for i in range(len(hue)):
        for j in range(i + 1, len(hue)):
            d = abs(hue[i][1] - hue[j][1])
            if d < GRAY_MIN:
                bad.append((hue[i][0], hue[j][0], d))
    return bad


def is_mono(colors):
    """もともとグレーだけで描かれているか。"""
    return not any(c[3] >= CHROMA_MIN for c in colors)


def survey():
    doc, z, items, rels = read_honpon()
    zu, hyo = [], []
    chap = sec = ""
    for i, it in enumerate(items):
        if it["kind"] == "p":
            if it["style"] == "Heading 1":
                chap = it["text"]
            elif it["style"] == "Heading 2":
                sec = it["text"]
            if not it["embeds"]:
                continue
            mae = ""
            for j in range(i - 1, max(-1, i - 4), -1):
                if items[j]["kind"] == "p" and items[j]["text"]:
                    mae = items[j]["text"]
                    break
            ato = ""
            for j in range(i + 1, min(len(items), i + 4)):
                if items[j]["kind"] == "p" and items[j]["text"]:
                    ato = items[j]["text"]
                    break
            for rid in it["embeds"]:
                path = rels.get(rid, "")
                if not path:
                    continue
                size, cols = image_colors(z, path)
                zu.append({
                    "chap": chap, "sec": sec, "path": path, "size": size,
                    "mae": mae, "ato": ato, "colors": cols,
                    "bad": collide(cols), "mono": is_mono(cols),
                    "bytes": z.getinfo(path).file_size,
                })
        else:
            t = it["obj"]
            r, c = len(t.rows), len(t.columns)
            ato = ""
            for j in range(i + 1, min(len(items), i + 3)):
                if items[j]["kind"] == "p" and items[j]["text"]:
                    ato = items[j]["text"]
                    break
            mae = ""
            for j in range(i - 1, max(-1, i - 4), -1):
                if items[j]["kind"] == "p" and items[j]["text"]:
                    mae = items[j]["text"]
                    break
            hyo.append({
                "chap": chap, "sec": sec, "rows": r, "cols": c,
                "head": " / ".join(x.text.strip().replace("\n", " ")
                                   for x in t.rows[0].cells[:4])[:60],
                "note": (r, c) == (1, 1),
                "src": any(k in ato for k in ("資料", "（各年", "（令和")),
                "ato": ato[:40], "mae": mae[:40],
            })
    return zu, hyo


# ===========================================================================
# 2. 図表番号と出所の付け方（規約）
# ===========================================================================
KIYAKU = [
    ("番号の形", "図は「図１-１」、表は「表１-１」。"
     "章ごとに通し番号を振り、章が変われば1に戻す",
     "章をまたぐ差し替えがあっても、ほかの章の番号が動かない"),
    ("番号を振る対象", "データを示す図と表のすべて。"
     "注記ボックス（【村へ確認】等の1×1の枠）には振らない",
     "注記は計画の確定時に削るため、番号を振ると欠番が出る"),
    ("置き場所", "表題は表の上、図の表題は図の下。"
     "いずれも本文と同じ幅に収める",
     "自治体の計画書で広く使われている形に合わせる"),
    ("出所の書き方", "図表の下に「資料：○○（基準日）」。"
     "村が作成したものは「資料：北塩原村」、"
     "アンケートは「資料：障がい福祉に関するアンケート調査（令和８年）」",
     "いま出所が付いているのは109表のうち4点しかない。"
     "数値の根拠を後から追えるようにする"),
    ("単位の書き方", "表の右上に「（人）」「（人日）」「（千円）」。"
     "図は軸のそばに置く", "表頭に単位を混ぜると列が読みにくくなる"),
    ("図表目次", "目次の後に「図表目次」を置き、"
     "図と表を分けて章順に並べる",
     "約60頁の計画書で図表が120点を超えるため、"
     "本文の目次だけでは図表にたどり着けない"),
    ("色の使い方", "他メンバー版と同じ3色（緑・橙・青）から使う。"
     "同じ図の中で系列が2つ以上あるときは、"
     "グレー値が25以上離れる組を選ぶ（緑と橙は52離れる。"
     "緑と青はいずれも119で同じ濃さになるため並べて使わない）。"
     "系列が接する積み上げでは網かけ（45度・135度）も足す",
     "図はカラーで進めるが、計画書は窓口で白黒コピーされることがある。"
     "色のほかにもう一つ手がかりを残しておけば、"
     "コピーされても系列を追える"),
    ("値のラベル", "表が併記されている図は、"
     "最新年度と最大・最小だけにラベルを付ける。"
     "表がない図は全ての値にラベルを付ける",
     "同じ数値を表と図の両方に全部書くと、"
     "直すときに片方が取り残される"),
]


# ===========================================================================
# 3. グラフにすべき表／表に戻すべき図
# ===========================================================================
GRAPH_KOUHO = [
    ("第5章7（1）給付費の推移", "令和2〜7年度の介護給付費等・障害児給付費・計",
     "積み上げ棒（6年×2系列）",
     "6年の推移と内訳を同時に見せられる。"
     "介護給付費等が約49.6％増えたことが一目で分かる",
     "当方のデータで今すぐ作れる（移植20）"),
    ("第5章7（3）サービス別の構成", "令和7年度のサービス別給付費8区分",
     "横棒（降順）",
     "就労継続支援（Ｂ型）が35.5％を占めることを示す。"
     "円グラフは8区分では読み取れないため使わない",
     "当方のデータで今すぐ作れる（移植20）"),
    ("第6章1（1）高齢化率", "令和3〜8年度の高齢化率",
     "折れ線（1系列）",
     "38.3％から43.1％への上昇。1系列のため凡例は要らない",
     "当方のデータで今すぐ作れる（正本 第6章）"),
    ("第4章1 成果目標の達成状況", "14項目の達成・一部達成・未達成",
     "表のまま（図にしない）",
     "区分が3つで項目が14。"
     "図にすると項目名が読めなくなるため表が適する",
     "現状のままでよい"),
    ("第5章1 サービス見込量", "11表・サービス種別ごとの実績と見込量",
     "主要サービスのみ折れ線（小さな図を並べる）",
     "全サービスを1枚にすると系列が多すぎる。"
     "利用の多い4サービス（生活介護・就労継続支援Ｂ型・共同生活援助・"
     "計画相談支援）に絞る",
     "村の実績（Ｍ-19）の確定後"),
    ("第2章6 障がい者手帳所持者数の将来推計", "令和9〜11年度の推計",
     "折れ線（実績と推計を分ける）",
     "実績は実線、推計は破線にして、どこから推計かを示す",
     "村の実績（Ｍ-28）の確定後"),
]


# ===========================================================================
# 4. 足りない図
# ===========================================================================
TARINAI = [
    ("計画の位置づけ図", "第1章2", "既にある（画像）",
     "国・県・村の関係と、本計画に内包する成年後見の計画（V-1で追加）を"
     "描き足す必要がある",
     "作り直し", "未着手（元データがないため。Ｓ-19）"),
    ("施策体系図", "第3章3", "ない",
     "7つの基本施策と本計画の成果目標8区分の対応が本文と表だけで、"
     "一覧できる図がない",
     "新しく作る",
     "作成済み（図Z-4）。基本目標と基本施策の対応は第4次障がい者計画の"
     "体系図にしかなく当方は未入手のため、線で結ばず図中に注記した"),
    ("関係機関との連携図", "第5章6", "ない",
     "移植19で11機関を表にしたが、"
     "本人を中心にどの機関がどの場面で関わるかの図がない"
     "（論点Ｒ-3で求められているもの）",
     "新しく作る",
     "作成済み（図Z-5）。移植19の表の11機関を7分野に分け、"
     "相談の入口である村保健福祉課を中心に置いた"),
    ("地域生活支援拠点の5機能", "第4章2（5）", "ない",
     "W-3で機能ごとの記録の表を入れた。"
     "5つの機能と会津北部4町村の関係を図にすると、"
     "村単独でなく圏域で確保していることが伝わる",
     "新しく作る",
     "作成済み（図Z-6）。W-3の表から5機能を起こし、"
     "4町村の拠点と本村の関わりを上下に置いた"),
    ("緊急時の4場面", "第5章2（11）", "ない",
     "W-8で平時・急病時・介護者不在時・災害時の表を入れた。"
     "時間の流れと仕組みの対応を図にできる",
     "新しく作る（表で足りるなら不要）",
     "作成済み（図Z-7）。平時・急病時は高齢者施策であり"
     "障がいのある方が対象かを確認していないこと（Ｍ-50・Ｍ-52）を"
     "図中に明示した"),
    ("65歳到達時の判定の流れ", "第5章5（2）", "ない",
     "W-1で確認6項目、移植18で4区分の表を入れた。"
     "判定の順序はフロー図の方が分かりやすい",
     "新しく作る",
     "作成済み（図Z-8）。4段の判定と3つの結論、"
     "W-1の記録6項目を1枚にした"),
    ("ＰＤＣＡサイクル図", "第1章6", "既にある（画像）",
     "W-10で年次の点検評価の7項目を入れたため、"
     "図と本文の対応を確かめる",
     "確認のみ", "未着手（他メンバーの図との突合が要るため）"),
]


# ===========================================================================
# 5. 確認事項
# ===========================================================================
KAKUNIN = [
    ("Ｓ-19", "他メンバー", "図の元データ",
     "既存の18図は画像で取り込まれており、数値を直せない。"
     "元のデータ（Excel等）と作図に使った道具を共有してほしい。"
     "村の実績が入れば、ほとんどの図を作り直すことになる",
     "令和8年10月16日"),
    ("Ｓ-20", "他メンバー", "緑と青の組合せ",
     "図はカラーで進めることとしたが、"
     "緑 RGB(46,158,107) と 青 RGB(61,134,198) はグレー値がいずれも119で、"
     "白黒コピーすると見分けられない。"
     "手帳の種類別のように3系列並べる図では、"
     "青を橙（グレー値171）に替えるか、網かけを足すことを提案したい",
     "令和8年10月16日"),
    ("Ｍ-69", "村", "印刷の色",
     "仕様書5の成果品①は「モノクロ・コピー・くるみ製本」であるが、"
     "図はカラーで作成している。"
     "①カラー印刷とする場合の費用と部数の扱い"
     "②モノクロ印刷とする場合に図を作り直す必要があること"
     "③窓口での白黒コピーを想定するかどうか、を確認したい",
     "令和8年10月20日"),
    ("Ｓ-21", "他メンバー", "図表番号と図表目次",
     "図表番号を振り、目次の後に図表目次を置くことを提案したい。"
     "番号の形（図１-１・表１-１）と、"
     "注記ボックスには番号を振らないことを合わせたい",
     "令和8年10月16日"),
]


# ===========================================================================
# シートの組み立て
# ===========================================================================
def sheet_matome(wb, zu, hyo):
    n_bad = sum(1 for z in zu if z["bad"])
    n_note = sum(1 for h in hyo if h["note"])
    n_src = sum(1 for h in hyo if h["src"])
    ws = add_sheet(
        wb, "00_結論",
        "正本の図表の棚卸し（令和8年10月10日）",
        "正本の現物から図と表を数え上げ、成果品として印刷できる状態かを"
        "確かめたもの。数値はすべて現物から読み直している。",
        [30, 100, 22])
    style_header_row(ws, 5, ["項目", "内容", "備考"])
    rows = [
        ("数えたもの",
         f"図{len(zu)}点／表{len(hyo)}点"
         f"（データ表{len(hyo) - n_note}・注記ボックス{n_note}）。"
         "図表番号は0件、図表目次もない",
         "01・03シート"),
        ("色はカラーで進める（令和8年10月10日の指示）",
         "図は他メンバー版に合わせてカラーで作る。"
         "印刷については確認事項として残す"
         "（仕様書5の成果品①はモノクロ・コピー・くるみ製本であり、"
         "仕様と図の作りが食い違っている）",
         "07シート。確認Ｍ-69・Ｓ-20"),
        ("ただし白黒コピーで潰れる組がある",
         f"{n_bad}点の図で、系列の色をグレーにすると同じ濃さになる。"
         "緑 RGB(46,158,107) と 青 RGB(61,134,198) はどちらもグレー値119で"
         "完全に一致する。"
         "カラーで印刷しても、計画書は窓口で白黒コピーされることがある。"
         "当方が作る図は、カラーのままでも同じ図の中の系列の"
         "グレー値を25以上離す（緑と青を並べて使わない）",
         "02シート"),
        ("出所がほとんど付いていない",
         f"表{len(hyo)}点のうち、直後に出所らしき行があるのは{n_src}点のみ。"
         "数値の根拠を後から追えない。"
         "庁議・議会・パブリックコメントで必ず問われる",
         "03シート。規約は06シート"),
        ("図は画像であり数値を直せない",
         "他メンバー版の図19点はすべて画像として貼られている。"
         "村の実績（Ｍ-19ほか）が入ると数値が変わるため、"
         "ほとんどの図を作り直すことになる。"
         "元データと作図の道具を他メンバーと共有しておく必要がある",
         "確認Ｓ-19"),
        ("当方が作った図3点",
         "給付費の推移・サービス別の給付費・高齢化率の推移を、"
         "他メンバー版と同じ色（緑）で入れた"
         "（build_kitashiobara_graph.py）。"
         "数値は申し送りの表と正本の表から読み込むため、"
         "表が直れば図も作り直せる。"
         "いずれも1系列であり、色で見分ける必要がない。"
         "2系列以上になるときは、グレー値が25以上離れる組"
         "（緑と橙。差52）を使う",
         "04シート。図の作り方の手本"),
        ("図表目次がない",
         "約60頁の計画書に図表が120点以上ある。"
         "本文の目次だけでは図表にたどり着けないため、"
         "目次の後に図表目次を置く",
         "06シート"),
        ("グラフにすべき表",
         "給付費の推移、サービス別の構成、高齢化率は"
         "当方のデータで今すぐグラフにできる。"
         "見込量と将来推計は村の実績が入ってから",
         "04シート"),
        ("足りない図",
         "施策体系図、関係機関との連携図（論点Ｒ-3）、"
         "地域生活支援拠点の5機能、緊急時の4場面、"
         "65歳到達時の判定の流れ。"
         "いずれも本文と表はあるが図がない",
         "05シート"),
        ("件数",
         f"図{len(zu)}点／表{len(hyo)}点／規約{len(KIYAKU)}件／"
         f"グラフ候補{len(GRAPH_KOUHO)}件／足りない図{len(TARINAI)}件／"
         f"確認事項{len(KAKUNIN)}件", "―"),
    ]
    r = 6
    for i, rec in enumerate(rows):
        r = write_row(ws, r, list(rec), alt=(i % 2 == 1))
    return ws


def sheet_zu(wb, zu):
    ws = add_sheet(
        wb, "01_図の一覧",
        "正本に入っている図（すべて画像）",
        "位置は正本の現物から読み取った。"
        "「モノクロ」欄は、画像の画素から計算したグレー値が"
        f"{GRAY_MIN}（256階調）以上離れているかで判定している。",
        [6, 16, 20, 40, 26, 14, 12, 26])
    style_header_row(ws, 5, ["No.", "章", "節", "直前の文", "直後の行",
                             "画像", "大きさ", "モノクロ"])
    r = 6
    for i, z in enumerate(zu, 1):
        han = ("潰れる" if z["bad"]
               else ("はじめからモノクロ" if z["mono"] else "見分けられる"))
        r = write_row(ws, r, [
            f"図{i}", z["chap"][:10], z["sec"][:14], z["mae"][:38],
            z["ato"][:24], os.path.basename(z["path"]),
            f"{z['size'][0]}×{z['size'][1]}", han],
            alt=(i % 2 == 1),
            aligns=["center", "left", "left", "left", "left", "left",
                    "center", "center"])
        c = ws.cell(row=r - 1, column=8)
        c.fill = PatternFill("solid", fgColor=(NG if z["bad"] else OKC))
    return ws


def sheet_mono(wb, zu):
    ws = add_sheet(
        wb, "02_モノクロ判定",
        "図の色をグレーにしたときの濃さ",
        "画像の画素を数え、背景の白と文字の黒を除いた主な色について"
        "グレー値（0が黒・255が白）を計算した。"
        f"差が{GRAY_MIN}未満の組は、モノクロ印刷で見分けられない。",
        [8, 24, 14, 10, 8, 52])
    style_header_row(ws, 5, ["No.", "画像", "色（RGB）", "グレー値",
                             "色み", "判定"])
    r = 6
    for i, z in enumerate(zu, 1):
        if not z["colors"]:
            continue
        for rgb, g, n, ch in z["colors"]:
            bad = any(rgb in (a, b) for a, b, _d in z["bad"])
            han = ("ほかの系列と同じ濃さになる" if bad
                   else ("グレー（目盛線・軸・網かけ）" if ch < CHROMA_MIN
                         else ""))
            r = write_row(ws, r, [
                f"図{i}", os.path.basename(z["path"]),
                f"({rgb[0]},{rgb[1]},{rgb[2]})", g, ch, han],
                alt=(i % 2 == 1),
                aligns=["center", "left", "center", "center", "center",
                        "left"])
            if bad:
                ws.cell(row=r - 1, column=4).fill = PatternFill(
                    "solid", fgColor=NG)
    r += 1
    style_header_row(ws, r, ["潰れる組", "グレー値の差", "どこに出るか"])
    r += 1
    seen = set()
    for i, z in enumerate(zu, 1):
        for a, b, d in z["bad"]:
            key = (a, b)
            if key in seen:
                continue
            seen.add(key)
            r = write_row(ws, r, [
                f"RGB{a} と RGB{b}", d,
                f"図{i} ほか（{os.path.basename(z['path'])}）"],
                aligns=["left", "center", "left"],
                fills=[NG, None, None])
    return ws


def sheet_hyo(wb, hyo):
    ws = add_sheet(
        wb, "03_表の一覧",
        "正本に入っている表",
        "注記ボックス（【村へ確認】等の1×1の枠）は計画の確定時に削るため、"
        "番号を振らない。「出所」は表の直後に資料名らしき行があるかどうか。",
        [6, 16, 22, 8, 8, 44, 10, 10])
    style_header_row(ws, 5, ["No.", "章", "節", "行", "列", "表頭",
                             "注記", "出所"])
    r = 6
    no = {}
    for h in hyo:
        if h["note"]:
            ban = "―"
        else:
            key = h["chap"][:3]
            no[key] = no.get(key, 0) + 1
            ban = f"表{key[1]}-{no[key]}"
        r = write_row(ws, r, [
            ban, h["chap"][:10], h["sec"][:16], h["rows"], h["cols"],
            h["head"], "注記" if h["note"] else "", "あり" if h["src"] else ""],
            alt=(r % 2 == 0),
            aligns=["center", "left", "left", "center", "center", "left",
                    "center", "center"])
        if not h["note"] and not h["src"]:
            ws.cell(row=r - 1, column=8).fill = PatternFill(
                "solid", fgColor=NG)
    return ws


def sheet_graph(wb):
    ws = add_sheet(
        wb, "04_グラフにすべき表",
        "表のままにするか、グラフにするか",
        "グラフにするのは、推移と構成比のように「形」で読ませたいものだけ。"
        "区分が多いものや、数値そのものを引用されるものは表のままが適する。",
        [26, 32, 22, 48, 28])
    style_header_row(ws, 5, ["箇所", "データ", "グラフの型", "なぜそうするか",
                             "作れる時期"])
    r = 6
    for i, rec in enumerate(GRAPH_KOUHO):
        r = write_row(ws, r, list(rec), alt=(i % 2 == 1))
        if "今すぐ" in rec[4]:
            ws.cell(row=r - 1, column=5).fill = PatternFill(
                "solid", fgColor=OKC)
    return ws


def sheet_tarinai(wb):
    ws = add_sheet(
        wb, "05_足りない図",
        "本文と表はあるが図がないもの",
        "図を増やすこと自体が目的ではない。"
        "関係が二次元（だれが・どの場面で）になっているものだけ図にする。",
        [26, 16, 18, 50, 20, 46])
    style_header_row(ws, 5,
                     ["図", "入れる先", "現状", "なぜ要るか", "扱い", "状態"])
    r = 6
    for i, rec in enumerate(TARINAI):
        r = write_row(ws, r, list(rec), alt=(i % 2 == 1))
    return ws


def sheet_kiyaku(wb):
    ws = add_sheet(
        wb, "06_番号と出所の規約",
        "図表番号・出所・図表目次の付け方",
        "正本と当方の素案・概要版のすべてに同じ規約を当てる。",
        [20, 60, 50])
    style_header_row(ws, 5, ["項目", "決め方", "理由"])
    r = 6
    for i, rec in enumerate(KIYAKU):
        r = write_row(ws, r, list(rec), alt=(i % 2 == 1))
    return ws


def sheet_kakunin(wb):
    ws = add_sheet(
        wb, "07_確認事項",
        "他メンバー・村への確認",
        "図は画像で入っているため、元データがないと作り直せない。"
        "モノクロ印刷の可否は成果品の仕様に関わる。",
        [8, 14, 22, 62, 14])
    style_header_row(ws, 5, ["番号", "相手", "事項", "内容", "期限の目安"])
    r = 6
    for i, rec in enumerate(KAKUNIN):
        r = write_row(ws, r, list(rec), alt=(i % 2 == 1),
                      aligns=["center", "center", "left", "left", "center"])
    return ws


# ===========================================================================
# 自己点検
# ===========================================================================
def verify(zu, hyo):
    ng = []
    if not zu:
        ng.append("図が1点も見つからない（読み取りに失敗している）")
    if len(hyo) < 100:
        ng.append(f"表が{len(hyo)}点しかない（正本は100点を超える）")
    # グレー値の計算が正しいこと（白は255、黒は0）
    if gray((255, 255, 255)) != 255 or gray((0, 0, 0)) != 0:
        ng.append("グレー値の計算が合っていない")
    # 緑と青が同じ濃さになることを確かめる（本書の主張の根拠）
    g1, g2 = gray((46, 158, 107)), gray((61, 134, 198))
    if abs(g1 - g2) >= GRAY_MIN:
        ng.append(f"緑と青が潰れるとした根拠が崩れている（{g1} と {g2}）")
    # 潰れる図が1点以上あること（0なら判定が働いていない）
    if not any(z["bad"] for z in zu):
        ng.append("潰れる図が0点（判定が働いていない可能性）")
    # 当方が作った図が正本に入っていること（表題で確かめる）
    from build_kitashiobara_graph import GRAPHS
    body = "\n".join(p.text for p in docx.Document(HONPON).paragraphs)
    for g in GRAPHS:
        if g["title"] not in body:
            ng.append(f"当方の図が正本にない: {g['title']}")
    # 当方の図は1系列であり、潰れる組を持たないこと
    titles = {g["title"] for g in GRAPHS}
    # 当方の図は、直後の行が表題になっている（画像の名前では見分けられない）
    ours = [z for z in zu if z["ato"] in titles]
    if len(ours) != len(GRAPHS):
        ng.append(f"当方の図が{len(ours)}点しか見つからない"
                  f"（{len(GRAPHS)}点のはず）")
    for z in ours:
        if z["bad"]:
            ng.append(f"当方の図に潰れる組がある: {z['path']}")
    # 緑と青の組が実際に検出されていること（判定が働いている証拠）
    found = any((46, 158, 107) in (a, b) and (61, 134, 198) in (a, b)
                for z in zu for a, b, _d in z["bad"])
    if not found:
        ng.append("緑と青の組が検出されていない（判定が働いていない）")
    # 章の割り当てができていること
    if any(not z["chap"] and z is not zu[0] for z in zu[1:]):
        ng.append("章が割り当てられていない図がある")
    for rec in KAKUNIN:
        if rec[1] not in ("村", "県", "他メンバー"):
            ng.append(f"相手が未定義: {rec[0]}")
    if ng:
        print("自己点検 不合格:")
        for e in ng:
            print("   -", e)
        raise SystemExit(1)
    print(f"  自己点検: 図{len(zu)}点・表{len(hyo)}点を現物から読み取り／"
          f"グレー値の計算を白黒で検算／"
          f"緑{g1}と青{g2}の差{abs(g1 - g2)}（{GRAY_MIN}未満）を確認")


def main():
    ensure_out_dir()
    zu, hyo = survey()
    verify(zu, hyo)

    wb = Workbook()
    wb.remove(wb.active)
    sheet_matome(wb, zu, hyo)
    sheet_zu(wb, zu)
    sheet_mono(wb, zu)
    sheet_hyo(wb, hyo)
    sheet_graph(wb)
    sheet_tarinai(wb)
    sheet_kiyaku(wb)
    sheet_kakunin(wb)
    wb.save(OUT_FILE)

    n_bad = sum(1 for z in zu if z["bad"])
    n_note = sum(1 for h in hyo if h["note"])
    n_src = sum(1 for h in hyo if h["src"])
    print(f"作成: {OUT_FILE}")
    print(f"  シート数: {len(wb.sheetnames)}　{wb.sheetnames}")
    print(f"  図{len(zu)}点（モノクロで潰れる{n_bad}点）／"
          f"表{len(hyo)}点（注記{n_note}・出所あり{n_src}）")
    print(f"  グラフ候補{len(GRAPH_KOUHO)}件／足りない図{len(TARINAI)}件／"
          f"規約{len(KIYAKU)}件／確認事項{len(KAKUNIN)}件")


if __name__ == "__main__":
    main()
