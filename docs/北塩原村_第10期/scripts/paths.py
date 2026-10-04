# -*- coding: utf-8 -*-
"""ファイルの置き場所を1か所で決める.

【なぜ必要か】
  これまで各スクリプトが "/home/user/repository/..." をじか書きしていた。
  この書き方だと、作業環境が作り直されてリポジトリが別の場所に置かれたとき、
  **別の場所にあるリポジトリへ書き込む**か、見つからずに止まる。
  金ケ崎町の案件では本文そのものが版管理の外にあり、環境の作り直しで消えた。
  本村は本文を版管理に置いているため消えないが、置き場所がじか書きのままだと
  復元したリポジトリで組み直せない。

  自分（このファイル）の位置から数えることで、リポジトリがどこに置かれても働く。
    scripts/paths.py → scripts/ → 北塩原村_第10期/ → docs/ → リポジトリの根
"""
import os

SCRIPTS = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(SCRIPTS)                 # docs/北塩原村_第10期
DOCS = os.path.dirname(BASE)                    # docs
ROOT = os.path.dirname(DOCS)                    # リポジトリの根

OUT = os.path.join(ROOT, "output")              # 成果品
FIGURES = os.path.join(OUT, "figures")          # 図のPNG
DATA = os.path.join(BASE, "data")               # 算定の途中のデータ・受領データ
# 組立ての途中のファイル（JSON）。/tmp に置くと、複製した別のリポジトリと
# 同じファイルを奪い合い、片方の中身でもう片方が組まれる。
# リポジトリの中に置けば、どこに復元しても取り違えない。版管理には入れない。
BUILD = os.path.join(OUT, "_build")


def build(*a):
    os.makedirs(BUILD, exist_ok=True)
    return os.path.join(BUILD, *a)


def out(*a):
    return os.path.join(OUT, *a)


def data(*a):
    return os.path.join(DATA, *a)


def base(*a):
    return os.path.join(BASE, *a)


if __name__ == "__main__":
    for k in ("ROOT", "DOCS", "BASE", "SCRIPTS", "OUT", "FIGURES", "DATA"):
        print("%-9s %s" % (k, globals()[k]))
