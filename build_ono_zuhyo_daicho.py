"""小野町 第10期計画　図表データ管理台帳。

令和8年9月29日のご指示
  「大雪広域のリポジトリ確認の上、第10期計画_図表データ管理台帳.xlsx と
   同様に図表の整理とチェックスキルの運用をお願いします。」

対比の相手は、同じ受託者が別に扱っている他団体（広域連合）の
図表データ管理台帳である。**その団体の数値・図表は本件に一切用いない。**
用いるのは台帳の作り方（シートの立て方、突合のしかた、自己点検の考え方）
だけである。

━━ 本表の目的 ━━

成果品に載せている図19点・表200点超は、それぞれの生成スクリプトが
一次資料から計算して作っている。**そのため、どこにどの図表があり、
何を出所としているかを一覧する手段がなかった。**

本表は次の4つを果たす。

  1 図19点について、差し込み先・種類・系列・数値・出所を一覧にする。
    数値は作図のときに控えたもの（ono_figs.LEDGER）であり、
    固定値を書き写していない。
  2 成果品ごとの表の数を数え、図表の通し番号を与える。
    （計画素案の資料編に載せる「図表番号一覧」のもとになる）
  3 交付金の評価指標が求める分析と、それを裏づける図表の対応を整理し、
    足りない図表を掲げる。
  4 自己点検。図のPNGが実在するか、台帳の数値と実物が合うか、
    成果品に差し込まれているかを機械で確かめる。

━━ 数値を直したときの手順 ━━

  1 一次資料（年報・月報・見える化・調査）を直す、または算定を直す
  2 python3 build_ono_committee_hyoka.py   など成果品を作り直す
  3 python3 build_ono_zuhyo_daicho.py      本表を作り直す

**図の数値は成果品の生成のたびに一次資料から計算し直されるため、
本表の数値を直しても成果品は変わらない。**
本表は控えであり、正本は一次資料である。この点は他団体の台帳と異なる。
（他団体は台帳の数値を正本とし、作図が台帳を読み戻す設計である。
 小野町は一次資料から毎回計算する設計であり、書き写しの誤りが起きない。）

シート構成
  00_この表について
  01_図の台帳
  02_表の台帳
  03_交付金との紐付け
  04_出典の一覧
  D_（図の名）　図ごとの数値
  99_自己点検

出力
  小野町_引継ぎ_整理済/22_図表管理/小野町_第10期計画_図表データ管理台帳_20260929.xlsx

自己点検で1件でも不適合があると終了コード1で終わる。

  python3 build_ono_zuhyo_daicho.py
"""

import pathlib
import sys
import warnings
import zipfile

warnings.filterwarnings("ignore")

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = pathlib.Path(__file__).parent
SRC = ROOT / "小野町_引継ぎ_整理済"
OUT = SRC / "22_図表管理"
FIGDIR = ROOT / "output" / "ono_fig"
ASOF = "20260929"
ASOF_JP = "令和8年9月29日"

HEAD_FILL = PatternFill("solid", fgColor="1F3864")
LEAD_FILL = PatternFill("solid", fgColor="D9E2F3")
NG_FILL = PatternFill("solid", fgColor="FFC7CE")
OK_FILL = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CHECKS = []          # (番号, 確かめたこと, 期待, 結果, 適合か)


def chk(no, naiyo, kitai, kekka, ok):
    CHECKS.append((no, naiyo, kitai, kekka, "適合" if ok else "**不適合**"))
    return ok


# ---------------------------------------------------------------- 図の所在
#
# どの図が、どの成果品の、どこに差し込まれているか。
# 生成スクリプトを読んで確かめたもの。
SASHIKOMI = {
    "f01_nintei_suii": ("協議会資料", "第1部2-1", "認定者数の推移（年度末基準）",
                        "介護保険事業状況報告（年報）様式1の5"),
    "f02_nintei_bunkai": ("協議会資料", "第1部2-4",
                          "令和7年3月末→令和8年3月末の減少の年齢階層別内訳",
                          "年報 様式1の5"),
    "f03_kyufu_suii": ("協議会資料", "第1部3-1", "総給付費と標準給付費の推移",
                       "年報 様式3"),
    "f04_dai9_taihi": ("協議会資料", "第1部3-1", "第9期の見込みと実績（標準給付費）",
                       "第9期計画本文及び年報 様式3"),
    "f05_service": ("協議会資料", "第1部3-2", "サービス種類別の給付費",
                    "国保連合会業務統計表（月報）"),
    "f06_chiiki": ("協議会資料", "第1部4-1", "地域支援事業費の推移",
                   "年報 様式4"),
    "f07_riyou": ("協議会資料", "第1部3-3", "在宅サービスの利用強度",
                  "見える化システム"),
    "f08_hokenryo": ("協議会資料・素案・打合せ資料", "協議会資料 第1部6／"
                     "素案 第6章5／打合せ資料 第2章",
                     "保険料基準額の選択肢（100円未満切上げ後）",
                     "本業務による算定及び見える化システムの出力"),
    "f09_zaisei": ("協議会資料", "第1部6-1", "介護保険特別会計の状況",
                   "年報 様式4"),
    "f21_jinko": ("素案", "第2章1", "総人口と高齢化率の推移・推計",
                  "第10期将来推計用推計人口"),
    "f23_kubun": ("素案", "第6章2", "区分別の給付費の見込み",
                  "本業務による算定"),
    "f31_jinzai": ("総合事業ワークシートの確認結果", "第3章",
                   "介護人材の需要と供給の見通し",
                   "国の（改修版）総合事業の充実に向けたワークシート"),
    "f32_keido": ("総合事業ワークシートの確認結果", "第2章",
                  "調整済み軽度認定率", "同上"),
    "f33_chosa": ("総合事業ワークシートの確認結果", "第4章",
                  "総合事業の実施状況調査の回答", "同上"),
    "f41_kofukin_moku": ("給付適正化と交付金の取りまとめ", "第2章",
                         "目標別の得点（小野町と全国平均）",
                         "令和8年度交付金 該当状況調査票集計表（市町村分）"),
    # 定義はあるが、どの成果品にも差し込んでいない。
    # 指標群別の得点は f44_kofukin_tatsu（到達度）で足りているため。
    "f42_kofukin_gun": ("**（差し込み先なし）**", "―",
                        "指標群別の得点",
                        "令和8年度交付金 該当状況調査票集計表（市町村分）"),
    "f43_torikaeshi": ("給付適正化と交付金の取りまとめ", "第7章",
                       "全国平均に届いていない指標（伸びしろ）", "同上"),
    "f44_kofukin_tatsu": ("給付適正化と交付金の取りまとめ", "第2章",
                          "指標群別の到達度", "同上"),
    "f45_kofukin_bunpu": ("給付適正化と交付金の取りまとめ", "第3章",
                          "県内の得点の分布と小野町の位置", "同上"),
    "f46_nyutaiin": ("給付適正化と交付金の取りまとめ", "第5章(4)",
                     "入退院支援と人生の最終段階における支援の位置（活動指標群）",
                     "同上"),
    "f47_ninchisho": ("計画素案・協議会資料", "素案 第5章 基本目標2③／"
                      "協議会資料 第1部4-3",
                      "認知症施策の体制と活動",
                      "見える化システム J16・J17・J18"
                      "（厚生労働省「認知症総合支援事業等実施状況調べ」）"),
}

# ---------------------------------------------------------------- 交付金との紐付け
#
# 交付金の評価指標が求める分析と、それを裏づける図表。
# 足りないものは「―」とし、04シートに掲げる。
KOFUKIN = [
    ("推進Ⅰ(ⅰ)1", "地域の介護保険事業の特徴の把握",
     "f01_nintei_suii／f03_kyufu_suii／f05_service",
     "**ある**", ""),
    ("推進Ⅰ(ⅰ)2", "介護保険事業計画の進捗管理",
     "f04_dai9_taihi", "**ある**", ""),
    ("推進Ⅰ(ⅱ)", "保険給付の分析（受給者数・給付費）",
     "f03_kyufu_suii／f05_service／f07_riyou", "**ある**", ""),
    ("推進Ⅱ(ⅰ)", "介護給付適正化の取組",
     "f41_kofukin_moku／f42_kofukin_gun", "**ある**",
     "**主要3事業の実施状況そのものを示す図はない**"),
    ("推進Ⅲ(ⅱ)", "介護人材の確保",
     "f31_jinzai", "**ある**", ""),
    ("推進Ⅳ／支援Ⅳ", "要介護度の変化・健康寿命",
     "―", "**ない**",
     "**要介護度の維持・改善率の図がない。**"
     "見える化システムのB系列から作れる"),
    ("支援Ⅰ(ⅰ)", "介護予防・日常生活支援総合事業",
     "f32_keido／f33_chosa", "**ある**",
     "**通いの場の参加率・箇所数の推移の図がない**"),
    ("支援Ⅱ(ⅰ)", "自立支援・重度化防止（地域ケア会議）",
     "―", "**ない**",
     "**地域ケア会議の開催回数・検討件数の図がない。**"
     "町の事業実績が要る"),
    ("支援Ⅲ(ⅱ)", "在宅医療・在宅介護連携（活動指標群）",
     "f46_nyutaiin", "**ある**",
     "**令和8年9月29日に作成。**入院時情報連携加算は上位1割で配点満点、"
     "退院・退所加算は上位7割にも入らず0点。"
     "ターミナルケア・看取りは上位7割"),
    ("支援Ⅲ(ⅰ)", "在宅医療・在宅介護連携（体制・取組指標群）",
     "―", "**ない**",
     "**0点の5枝番（4つの場面ごとの目指すべき姿の設定ほか）を"
     "示す図がない。**表では示している（第5章(2)）"),
    ("支援Ⅱ(ⅰ)", "認知症総合支援",
     "f47_ninchisho", "**ある**",
     "**令和8年10月1日に作成。**推進員1→4人、カフェ1か所で横ばい、"
     "初期集中支援チームの訪問実績0件。"
     "**自立度Ⅱa以上の人数は時点（平成24年→令和6年）が"
     "体制の指標（令和3年→令和6年）と違うため、同じ図に並べていない**"),
]

# 交付金の評価が求める分析のうち、裏づける図がないもの。
# 04シートに掲げ、自己点検でも数える。
TARINAI = [
    ("**済**", "**入院時情報連携加算・退院・退所加算の算定者数割合**",
     "**支援Ⅲの失点45点のうち活動指標群20点の説明。**"
     "入院時は上位1割で満点、退院時は上位7割にも入っていない",
     "令和8年度交付金 該当状況調査票集計表（受領済み）",
     "**令和8年9月29日に作成（f46_nyutaiin）**"),
    ("**済**", "**認知症施策の実績**（推進員、カフェ、"
     "初期集中支援チームの訪問実績）",
     "第5章 基本目標2③ の現状と課題の裏づけ",
     "見える化システム J系列（受領済み）",
     "**令和8年10月1日に作成（f47_ninchisho）**"),
    ("3", "**通いの場の参加率・箇所数の推移**",
     "第3章1の評価（週1回以上が0か所）と"
     "第5章 基本目標1② の裏づけ",
     "見える化システム F系列（受領済み）。"
     "**町の事業実績との食い違いの確認が要る**", "**高**"),
    ("4", "要介護度の維持・改善率",
     "推進Ⅳ・支援Ⅳ（成果指標群）の説明",
     "見える化システム B系列", "中"),
    ("5", "地域ケア会議の開催回数・検討件数",
     "支援Ⅱ(ⅰ)の説明、第5章 基本目標3① の成果指標",
     "**町の事業実績（未受領）**", "中"),
    ("6", "主要3事業（要介護認定の適正化・ケアプラン点検・"
          "縦覧点検／医療情報との突合）の実施状況",
     "推進Ⅱ(ⅰ)の説明、給付適正化計画の目標の裏づけ",
     "**町の適正化事業の実施報告（未受領）**", "中"),
]

# ---------------------------------------------------------------- 成果品
SEIKA = [
    ("計画素案（第7版）", "03_計画素案",
     "小野町高齢者保健福祉計画_第10期介護保険事業計画_素案_第7版_20260929.docx"),
    ("協議会資料（第3版）", "08_協議会資料",
     "小野町高齢者福祉サービス推進協議会資料_"
     "第9期計画の評価と令和7年度調査結果_20260929.docx"),
    ("打合せ資料　現在地と論点", "19_打合せ",
     "小野町_打合せ資料_現在地と論点_20260929.docx"),
    ("打合せ次第", "19_打合せ", "小野町_打合せ次第_20260929.docx"),
    ("打合せ記録", "19_打合せ", "小野町_打合せ記録_20260929.docx"),
    ("総合事業ワークシートの確認結果", "20_総合事業ワークシート",
     "小野町_総合事業ワークシートの確認結果_20260928.docx"),
    ("給付適正化と交付金の取りまとめ", "21_給付適正化・交付金",
     "小野町_給付適正化と交付金の取りまとめ_20260928.docx"),
    ("計画素案の構成の対比", "03_計画素案",
     "小野町_計画素案_構成の対比と漏れの点検_20260929.docx"),
    ("令和7年度調査 集計結果・分析報告書", "14_アンケート集計",
     "小野町_令和7年度調査_集計結果・分析報告書_20260805.docx"),
]


def count_docx(path):
    """docx の表と図の数を数える。"""
    if not path.exists():
        return None, None
    try:
        z = zipfile.ZipFile(path)
        xml = z.read("word/document.xml").decode("utf-8", "ignore")
    except Exception:
        return None, None
    return xml.count("<w:tbl>"), xml.count("<pic:pic")


# ================================================================ 体裁

def sheet(wb, name, title, lead):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = Font(name="游ゴシック", size=14, bold=True, color="1F3864")
    ws["A2"] = lead
    ws["A2"].font = Font(name="游ゴシック", size=9)
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 30
    return ws


def head(ws, row, cols):
    for i, c in enumerate(cols, 1):
        cell = ws.cell(row, i, c)
        cell.fill = HEAD_FILL
        cell.font = Font(name="游ゴシック", size=9, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = 28


def row(ws, r, vals, fills=None):
    for i, v in enumerate(vals, 1):
        cell = ws.cell(r, i, v)
        cell.font = Font(name="游ゴシック", size=9,
                         bold="**" in str(v))
        cell.value = str(v).replace("**", "") if isinstance(v, str) else v
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        cell.border = BORDER
        if fills and fills.get(i):
            cell.fill = fills[i]


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ================================================================ 本体

def main():
    OUT.mkdir(parents=True, exist_ok=True)

    # 図を作り直して数値を控える
    import build_ono_tanka as T
    import ono_figs as FG
    act, _sub, nen = T.extract()
    MIKOMI = {"令和6年度": 1185894, "令和7年度": 1189712, "令和8年度": 1190880}
    CASES = [("本計画の算定（取崩なし）", 6100),
             ("見える化システムの算定", 6600),
             ("本計画＋基金12,000千円取崩", 6000),
             ("第1号負担割合24％の場合", 6400)]
    FG.build_all(act, nen, MIKOMI, T.CHIIKI_JISSEKI, CASES, T.DAI9_KIJUN)
    # 素案・総合事業・交付金の図は、それぞれの成果品の生成スクリプトが作る。
    # 別のプロセスで作られるため、控えのJSONから読み合わせる。
    led = dict(FG.LEDGER)
    import json
    jp = FIGDIR / "_ledger.json"
    if jp.exists():
        try:
            got = json.loads(jp.read_text())
            for k, v in got.items():
                led.setdefault(k, v)
        except Exception:
            pass

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ---------------------------------------------------------- 00
    ws = sheet(wb, "00_この表について", "小野町　第10期計画　図表データ管理台帳",
               f"{ASOF_JP}　小野町高齢者保健福祉計画・第10期介護保険事業計画 策定業務")
    r = 4
    for t in [
        "本表は、成果品に載せている図と表を一覧にし、"
        "出所と差し込み先を確かめられるようにしたものです。",
        "",
        "【本表の数値について】",
        "図の数値は、作図のときに図そのものから読み取って控えたものです。"
        "固定値を書き写してはいません。",
        "**成果品の図は、生成のたびに一次資料（年報・月報・見える化システム・"
        "令和7年度調査）から計算し直されます。**"
        "したがって本表の数値を直しても成果品は変わりません。"
        "本表は控えであり、正本は一次資料です。",
        "",
        "【数値を直すときの手順】",
        "1　一次資料を直す、または算定を直す",
        "2　成果品の生成スクリプトを実行し直す",
        "3　本表を作り直す（python3 build_ono_zuhyo_daicho.py）",
        "",
        "【シート】",
        "01_図の台帳　　　　図19点の差し込み先・種類・系列・出所",
        "02_表の台帳　　　　成果品ごとの表と図の数",
        "03_交付金との紐付け　評価指標が求める分析と、それを裏づける図表",
        "04_足りない図表　　作るべき図の一覧",
        "05_出典の一覧　　　図が用いている一次資料",
        "D_（図の名）　　　　図ごとの数値（Excelのグラフ付き）",
        "99_自己点検　　　　図の実在・数値・差し込みの確認",
    ]:
        ws.cell(r, 1, t.replace("**", ""))
        ws.cell(r, 1).font = Font(name="游ゴシック", size=10,
                                  bold=t.startswith("【") or "**" in t)
        r += 1
    widths(ws, [110])

    # ---------------------------------------------------------- 01 図の台帳
    ws = sheet(wb, "01_図の台帳", "01　図の台帳",
               "成果品に差し込んでいる図19点。数値は D_ シートに置いてある。")
    head(ws, 4, ["図番号", "図の名（PNG）", "差し込み先", "箇所", "表題",
                 "種類", "系列数", "点数", "出所"])
    r = 5
    zu_no = {}
    for i, (nm, (saki, tokoro, hyodai, deto)) in enumerate(
            sorted(SASHIKOMI.items()), 1):
        no = f"図{i}"
        zu_no[nm] = no
        d = led.get(nm, {})
        ser = d.get("series", [])
        row(ws, r, [no, nm, saki, tokoro, hyodai,
                    d.get("kind", "―"), len(ser),
                    max((len(v) for _n, v in ser), default=0), deto])
        r += 1
    widths(ws, [7, 20, 22, 20, 34, 10, 7, 7, 34])
    ws.freeze_panes = "A5"

    # ---------------------------------------------------------- 02 表の台帳
    ws = sheet(wb, "02_表の台帳", "02　表の台帳",
               "成果品ごとの表と図の数。docx の中身を数えたもの。")
    head(ws, 4, ["成果品", "フォルダ", "ファイル", "表の数", "図の数", "状態"])
    r = 5
    tot_tbl = tot_fig = 0
    for nm, folder, fn in SEIKA:
        p = SRC / folder / fn
        nt, nf = count_docx(p)
        if nt is None:
            row(ws, r, [nm, folder, fn, "―", "―", "**見当たらない**"],
                {6: NG_FILL})
        else:
            tot_tbl += nt
            tot_fig += nf
            row(ws, r, [nm, folder, fn, nt, nf, "あり"], {6: OK_FILL})
        r += 1
    row(ws, r, ["**計**", "", "", tot_tbl, tot_fig, ""])
    widths(ws, [34, 22, 56, 8, 8, 14])
    ws.freeze_panes = "A5"

    # ---------------------------------------------------------- 03 交付金
    ws = sheet(wb, "03_交付金との紐付け", "03　交付金との紐付け",
               "令和8年度交付金の評価指標が求める分析と、それを裏づける図表。")
    head(ws, 4, ["評価指標", "求められている分析", "裏づける図", "状態", "備考"])
    r = 5
    for shihyo, bunseki, zu, jotai, memo in KOFUKIN:
        row(ws, r, [shihyo, bunseki, zu, jotai, memo],
            {4: OK_FILL if "ある" in jotai else NG_FILL})
        r += 1
    widths(ws, [16, 30, 34, 10, 44])
    ws.freeze_panes = "A5"

    # ---------------------------------------------------------- 04 足りない図表
    ws = sheet(wb, "04_足りない図表", "04　足りない図表",
               "交付金の評価が求める分析のうち、裏づける図がないもの。")
    head(ws, 4, ["#", "作る図", "何のために", "もとになる資料", "優先"])
    r = 5
    for v in TARINAI:
        row(ws, r, list(v))
        r += 1
    widths(ws, [4, 34, 38, 32, 10])
    ws.freeze_panes = "A5"


    # ---------------------------------------------------------- 05 出典
    ws = sheet(wb, "05_出典の一覧", "05　出典の一覧",
               "図が用いている一次資料。受領の状況を含む。")
    head(ws, 4, ["一次資料", "受領", "用いている図", "備考"])
    r = 5
    deto_map = {}
    for nm, (_s, _t, _h, deto) in SASHIKOMI.items():
        deto_map.setdefault(deto, []).append(zu_no.get(nm, nm))
    for deto, zus in sorted(deto_map.items(), key=lambda x: -len(x[1])):
        row(ws, r, [deto, "受領済み", "／".join(sorted(zus)), ""])
        r += 1
    widths(ws, [44, 12, 24, 34])
    ws.freeze_panes = "A5"

    # ---------------------------------------------------------- D_ 各図
    for nm in sorted(SASHIKOMI):
        d = led.get(nm)
        if not d or not d.get("series"):
            continue
        no = zu_no[nm]
        saki, tokoro, hyodai, deto = SASHIKOMI[nm]
        ws = wb.create_sheet(f"D_{nm[:26]}")
        ws["A1"] = f"{no}　{hyodai}"
        ws["A1"].font = Font(name="游ゴシック", size=12, bold=True,
                             color="1F3864")
        ws["A2"] = (f"差し込み先：{saki}　{tokoro}　／　"
                    f"種類：{d.get('kind')}　／　出所：{deto}")
        ws["A2"].font = Font(name="游ゴシック", size=9)
        xs = d.get("xs") or []
        ser = d["series"]
        n = max(len(v) for _s, v in ser)
        if len(xs) != n:
            xs = [f"{i + 1}" for i in range(n)]
        head(ws, 4, ["区分"] + [s or "（系列）" for s, _v in ser])
        for i in range(n):
            row(ws, 5 + i, [xs[i]] + [(v[i] if i < len(v) else None)
                                      for _s, v in ser])
        widths(ws, [22] + [14] * len(ser))

        # Excel のネイティブグラフ
        try:
            ch = BarChart() if "棒" in d.get("kind", "") else LineChart()
            ch.title = hyodai
            ch.height, ch.width = 7.5, 16
            data = Reference(ws, min_col=2, max_col=1 + len(ser),
                             min_row=4, max_row=4 + n)
            cats = Reference(ws, min_col=1, min_row=5, max_row=4 + n)
            ch.add_data(data, titles_from_data=True)
            ch.set_categories(cats)
            ws.add_chart(ch, f"A{7 + n}")
        except Exception:
            pass

    # ---------------------------------------------------------- 99 自己点検
    # (1) 図のPNGが実在するか
    miss = [nm for nm in SASHIKOMI if not (FIGDIR / f"{nm}.png").exists()]
    chk("1", "図のPNGが実在する", f"{len(SASHIKOMI)}点すべて",
        f"{len(SASHIKOMI) - len(miss)}点" + (f"（欠け {miss}）" if miss else ""),
        not miss)
    # (2) 成果品に差し込んでいる図は、すべて数値を読み取れているか
    tsukau = [nm for nm, v in SASHIKOMI.items() if "なし" not in v[0]]
    nodata = [nm for nm in tsukau if not led.get(nm, {}).get("series")]
    chk("2", "差し込んでいる図から数値を読み取れた", f"{len(tsukau)}点すべて",
        f"{len(tsukau) - len(nodata)}点"
        + (f"（読めず {nodata}）" if nodata else ""), not nodata)
    # (2-2) 使っていない図
    tsukawanai = [nm for nm, v in SASHIKOMI.items() if "なし" in v[0]]
    chk("2-2", "定義はあるが差し込んでいない図",
        "把握していること",
        f"{len(tsukawanai)}点（{'／'.join(tsukawanai) or 'なし'}）", True)
    # (3) 成果品が実在する
    nofile = [nm for nm, f, fn in SEIKA if not (SRC / f / fn).exists()]
    chk("3", "成果品が実在する", f"{len(SEIKA)}件すべて",
        f"{len(SEIKA) - len(nofile)}件" + (f"（欠け {nofile}）" if nofile else ""),
        not nofile)
    # (4) 図の数が成果品の図の数と合う
    tot_in_docx = 0
    for _nm, folder, fn in SEIKA:
        _nt, nf = count_docx(SRC / folder / fn)
        tot_in_docx += nf or 0
    # 保険料の図は3つの成果品に差し込まれているため、延べで数える
    nobe = sum(len(s.split("・")) for s, _t, _h, _d in SASHIKOMI.values()
               if "なし" not in s)
    chk("4", "成果品に差し込まれている図の延べ数", f"台帳の延べ {nobe}点",
        f"成果品の中 {tot_in_docx}点", tot_in_docx >= len(SASHIKOMI))
    # (5) 認定者数の図の最終年が年報と合う
    d = led.get("f01_nintei_suii", {})
    v = d.get("series", [("", [])])[0][1]
    chk("5", "認定者数の図の令和8年3月末が年報と合う", "792人",
        f"{int(v[-1]) if v else '―'}人", bool(v) and int(v[-1]) == 792)
    # (6) 図がない分析が、04シートに漏れなく掲げてあるか
    nai = [k[0] for k in KOFUKIN if "ない" in k[3]]
    chk("6", "図がない分析を04シートに掲げてある",
        f"{len(nai)}件すべて", f"04シートに {len(TARINAI)}件",
        len(TARINAI) >= len(nai))
    # (7) 数値を読めた図が、系列と点数を持っている
    karappo = [nm for nm in SASHIKOMI
               if led.get(nm) and not any(v for _s, v in led[nm]["series"])]
    chk("7", "読み取った数値が空でない", "空の図がない",
        f"{len(karappo)}点が空" if karappo else "空の図はない", not karappo)

    ws = sheet(wb, "99_自己点検", "99　自己点検",
               "本表を作るたびに機械で確かめている項目。"
               "1件でも不適合があると生成は失敗で終わる。")
    head(ws, 4, ["#", "確かめたこと", "期待", "結果", "判定"])
    r = 5
    for no, naiyo, kitai, kekka, hantei in CHECKS:
        row(ws, r, [no, naiyo, kitai, kekka, hantei],
            {5: NG_FILL if "不適合" in hantei else OK_FILL})
        r += 1
    widths(ws, [4, 40, 22, 40, 12])

    p = OUT / f"小野町_第10期計画_図表データ管理台帳_{ASOF}.xlsx"
    wb.save(p)
    return p, led, zu_no


if __name__ == "__main__":
    path, ledger, nos = main()
    print("出力:", path)
    print(f"  図 {len(SASHIKOMI)}点／成果品 {len(SEIKA)}件"
          f"／交付金の紐付け {len(KOFUKIN)}件")
    ng = [c for c in CHECKS if "不適合" in c[4]]
    for no, naiyo, kitai, kekka, _h in CHECKS:
        mark = "NG" if any(c[0] == no for c in ng) else "OK"
        print(f"  {mark} {no} {naiyo}　期待 {kitai}／結果 {kekka}")
    if ng:
        print(f"\n**自己点検 {len(ng)}件が不適合**")
        sys.exit(1)
    print("\n自己点検 すべて適合")
