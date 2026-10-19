# -*- coding: utf-8 -*-
"""公表名簿の取り出し先への疎通の確認（令和8年10月19日）

  python3 07_ソーススクリプト/check_setsuzoku_R8.10.19.py

なぜ要るか
  作業順位26（県の公表名簿との突合）は、当方の作業環境から
  公表システム・県・町・国・e-Stat のいずれにも接続できないため
  （プロキシが CONNECT に対して403を返す）、組み替えて実施した。
  **環境の設定で許可が加わりしだい、突合そのものを行える。**
  その判定を一手で、記録の残る形で行う。

  終了コード　0＝必須の先がすべて通る／1＝1つでも通らない

⚠ **ここに挙げる先は、町・県・国が公表しているものに限る。**
  地域包括ケア「見える化」システムは**挙げない。**
  IDパスワードは預かっているが当方で直接データを触る想定ではなく、
  見える化対応は町が進めるためである（CLAUDE.md 3）。

⚠ **川崎町のドメインは当方で確かめられていない。**
  遮断されているため解決するかを試せない。
  ⚠ **川崎町は福岡県田川郡にも同名の町がある。**
  通るようになったら、まず町の名称と所在地が宮城県柴田郡であることを
  確かめる（検査2がこれを行う）。
"""
import re
import subprocess
import sys

# （用途, ホスト, 必須か, 何に効くか）
SAKI = [
    ("介護サービス情報公表システム", "www.kaigokensaku.mhlw.go.jp", True,
     "事業所番号・名称・所在地・定員・職員数（確認事項No.29）。"
     "宮城県の入口は /04/"),
    ("宮城県", "www.pref.miyagi.jp", True,
     "指定介護サービス事業所の一覧／有料老人ホームの届出状況の公表一覧"
     "（No.29・No.154）"),
    ("川崎町", "www.town.kawasaki.miyagi.jp", True,
     "⚠ **ドメインは未確認。** 地域密着型サービスの指定・町の計画・"
     "例規（No.181 の条例・規則）"),
    ("厚生労働省", "www.mhlw.go.jp", True,
     "第10期基本指針・全国課長会議資料・交付金の評価指標（No.163）"),
    ("政府統計の総合窓口（e-Stat）", "www.e-stat.go.jp", True,
     "介護保険事業状況報告・人口（高齢化率の出典。No.175）"),
    ("サービス付き高齢者向け住宅情報提供システム", "www.satsuki-jutaku.jp",
     False,
     "⚠ **No.154（最優先）に直に効く。** サ高住の登録状況。"
     "突合台帳の経路⑧に挙げたがURLを書いていなかった"),
    ("総務省", "www.soumu.go.jp", False,
     "過疎地域の指定・特定地域づくり事業（素案の比較で引用済み）"),
    ("総務省統計局", "www.stat.go.jp", False,
     "国勢調査・人口推計"),
    ("福祉医療機構（WAM NET）", "www.wam.go.jp", False,
     "介護事業所情報の補助的な確認"),
    ("法令検索（e-Gov）", "laws.e-gov.go.jp", False,
     "条文の原文（No.181 の資料編の条文抜粋。"
     "⚠ 添付の引用をそのまま使わないため要る）"),
]

MIERUKA = ("地域包括ケア「見える化」システム", "mieruka.mhlw.go.jp",
           "⚠ **挙げない。** 当方で直接データを触る想定ではない"
           "（CLAUDE.md 3）。見える化対応は町が進める")


def tameshi(host, timeout=15):
    """HTTPS の疎通を試す。戻り　(コード, 説明)"""
    try:
        p = subprocess.run(
            ["curl", "-sS", "-o", "/dev/null", "-w", "%{http_code}",
             "--max-time", str(timeout), f"https://{host}/"],
            capture_output=True, text=True, timeout=timeout + 10)
        code = (p.stdout or "").strip().splitlines()
        code = code[-1] if code else ""
        err = (p.stderr or "").strip()
        if code == "000" or not code.isdigit():
            if "403" in err:
                return ("403", "プロキシが接続を拒む（環境の設定で未許可）")
            if "Could not resolve" in err or "resolve host" in err:
                return ("名前解決できない", "⚠ ドメインが違うおそれ")
            if "timed out" in err.lower() or "timeout" in err.lower():
                return ("応答なし", "待ち時間を超えた")
            return ("通らない", err.splitlines()[-1][:60] if err else "理由不明")
        return (code, "通る" if code.startswith(("2", "3")) else
                "接続はできたが応答がこの状態")
    except subprocess.TimeoutExpired:
        return ("応答なし", "待ち時間を超えた")
    except FileNotFoundError:
        return ("未実施", "curl がない")


def main():
    print("公表名簿の取り出し先への疎通の確認（令和8年10月19日）")
    print()
    print(f"  {'用途':<34}{'区分':<6}{'結果':<16}説明")
    print("  " + "─" * 104)
    kekka = []
    for yoto, host, hissu, nani in SAKI:
        code, setsu = tameshi(host)
        tsuru = code.startswith(("2", "3"))
        kekka.append((yoto, host, hissu, tsuru, code, setsu, nani))
        mark = "○" if tsuru else "×"
        print(f"  {mark} {yoto[:32]:<32}{'必須' if hissu else '任意':<6}"
              f"{code:<16}{setsu}")
        print(f"      https://{host}/")
    print()
    print(f"  ⚠ {MIERUKA[0]}（{MIERUKA[1]}）")
    print(f"      {MIERUKA[2]}")
    print()

    hissu_ng = [x for x in kekka if x[2] and not x[3]]
    nin_ng = [x for x in kekka if not x[2] and not x[3]]
    tsuru = [x for x in kekka if x[3]]

    print("  ── まとめ")
    print(f"   通る {len(tsuru)}件／通らない {len(kekka) - len(tsuru)}件"
          f"（うち必須 {len(hissu_ng)}件）")
    if hissu_ng:
        print()
        print("   ⚠ **必須の先が通らないため、作業順位26の突合は行えません。**")
        print("      環境のネットワークの設定に次のホストを加えていただく"
              "必要があります。")
        for x in hissu_ng:
            print(f"        {x[1]}")
        if nin_ng:
            print("      あわせて加えていただくと効くもの：")
            for x in nin_ng:
                print(f"        {x[1]}")
        print("      ⚠ **「Allow package managers」のチェックは"
              "外さないでください。**")
        print("         外すと pypi 等が通らなくなり、"
              "openpyxl・python-docx の導入ができません。")
        sys.exit(1)

    print()
    print("   ○ **必須の先はすべて通ります。作業順位26の突合を行えます。**")
    if nin_ng:
        print(f"   △ 任意の先が {len(nin_ng)}件通りません："
              + "／".join(x[1] for x in nin_ng))

    # ══════ 検査2　川崎町が宮城県柴田郡であることを確かめる
    #   ⚠ 福岡県田川郡にも同名の町がある。通ったら必ず確かめる。
    machi = [x for x in kekka if x[1] == "www.town.kawasaki.miyagi.jp"]
    if machi and machi[0][3]:
        print()
        print("  ── 検査2　川崎町が宮城県柴田郡であることの確認")
        p = subprocess.run(
            ["curl", "-sS", "--max-time", "20", "-L",
             "https://www.town.kawasaki.miyagi.jp/"],
            capture_output=True, text=True)
        hon = re.sub(r"<[^>]+>", " ", p.stdout or "")
        shibata = "柴田郡" in hon or "989-15" in hon
        tagawa = "田川郡" in hon or "福岡県" in hon
        if shibata and not tagawa:
            print("   ○ 宮城県柴田郡の川崎町である（柴田郡又は郵便番号を確認）")
        elif tagawa:
            print("   × ⚠ **福岡県田川郡の川崎町である。ドメインが違う。**")
            sys.exit(1)
        else:
            print("   △ 柴田郡の表記を拾えなかった。"
                  "⚠ **目で確かめてから突合に使う。**")


if __name__ == "__main__":
    main()
