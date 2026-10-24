# -*- coding: utf-8 -*-
"""仙南圏域の在宅系の地域密着型サービス事業所（介護サービス情報公表システム オープンデータ）
（令和8年10月23日）

  python3 07_ソーススクリプト/fetch_kinrin_chiiki_R8.10.23.py

第2回策定委員会資料 選択３の「近隣市町村の指定事業所の有無は確認できておりません【町確認】」を
公表資料で確かめる。対象は宮城県の高齢者福祉圏域「仙南圏域」（2市7町。素案 1-2・宮城県
推進委員会資料2）の、町内に事業所がない4つのサービス。

  710 夜間対応型訪問介護／730 小規模多機能型居宅介護／
  760 定期巡回・随時対応型訪問介護看護／770 看護小規模多機能型居宅介護

⚠ 地域密着型サービスは、原則として事業所の所在市町村の被保険者が利用するもの。町外の事業所を
  本町の被保険者が使うには、本町による指定（区域外の指定）が要る。本表は事業所の所在だけを
  示し、本町が指定しているかどうかは公表資料では分からない（町のご確認が残る）。
⚠ 電話番号は収録しない。事業所名は内部の確認用であり、成果品には件数だけを書く。

入口と抽出の手順は fetch_kouhyou_opendata_R8.10.20.py（CSV の場所はページから読む）。
出力　09_元資料/その他データ/公表システム_仙南圏域_在宅系地域密着型_R8.6末.csv
終了コード　0＝取得・点検とも適合／1＝不適合
"""
import csv
import importlib.util
import io
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("kouhyou", os.path.join(HERE, "fetch_kouhyou_opendata_R8.10.20.py"))
K = importlib.util.module_from_spec(spec)
spec.loader.exec_module(K)

SENNAN = ["白石市", "角田市", "蔵王町", "七ヶ宿町", "大河原町", "村田町", "柴田町", "川崎町", "丸森町"]
SHU = {"710": "夜間対応型訪問介護", "730": "小規模多機能型居宅介護",
       "760": "定期巡回・随時対応型訪問介護看護", "770": "看護小規模多機能型居宅介護"}
OUT = "09_元資料/その他データ/公表システム_仙南圏域_在宅系地域密着型_R8.6末.csv"


def main():
    saki, kijun_ok = K.nyugushi_yomu()
    print(f"介護サービス情報公表システム オープンデータ（基準日の見出し {'一致' if kijun_ok else '⚠ 変わっている'}）")
    gyo = []
    ng = 0
    for k, na in SHU.items():
        if k not in saki:
            print(f"   × {k} {na} の CSV が入口のページにない")
            ng += 1
            continue
        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as f:
            tmp = f.name
        code, size = K.toru(saki[k][0], tmp)
        raw = open(tmp, "rb").read()
        os.unlink(tmp)
        if code != "200" or not raw:
            print(f"   × {k} {na} が取れない（HTTP {code}）")
            ng += 1
            continue
        for enc in ("utf-8-sig", "cp932"):
            try:
                t = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        rows = list(csv.DictReader(io.StringIO(t)))
        h = rows[0].keys() if rows else []
        ken = next(c for c in h if c.startswith("都道府県名"))
        shi = next(c for c in h if c.startswith("市区町村名"))
        jn = next(c for c in h if c.startswith("事業所名"))
        bn = next(c for c in h if c.startswith("事業所番号"))
        n = 0
        for r in rows:
            if r[ken] == "宮城県" and r[shi] in SENNAN:
                gyo.append([k, na, r[shi], r[jn], r[bn]])
                n += 1
        print(f"   ○ {k} {na}：宮城県 {sum(1 for r in rows if r[ken] == '宮城県')}件のうち仙南圏域 {n}件")
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["サービスコード", "サービスの種類", "市町村", "事業所名", "事業所番号"])
        w.writerows(sorted(gyo))
    print(f"  {OUT}（{len(gyo)}行）")
    for k, na in SHU.items():
        m = sorted({g[2] for g in gyo if g[0] == k})
        print(f"    {na}：{len([g for g in gyo if g[0] == k])}件（{'・'.join(m) or 'なし'}）")
    if any(g[2] == "川崎町" for g in gyo):
        print("   × 町内に事業所が現れた（素案 2-7 の「いずれも0件」と食い違う）")
        ng += 1
    sys.exit(1 if ng else 0)


if __name__ == "__main__":
    main()
