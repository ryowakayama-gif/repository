#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""配布用ZIP（追試可能なエビデンス一式）を組み立てる。

リポジトリ上の配置と配布時の配置は異なるため、ここで対応づける。
配布物はスクリプト相対でパスを解決するので、展開すればそのまま再実行できる。

    $ python3 mkbundle.py [出力先.zip]
"""
import io
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = (sys.argv[1] if len(sys.argv) > 1
       else os.path.join(HERE, '階上町_6-10㎥単価差縮小案_エビデンス.zip'))

# (配布先パス, リポジトリ上のパス)。ディレクトリを指すと中身を再帰的に入れる。
LAYOUT = [
    ('00_入力データ',                                  '00_入力データ'),
    ('01_成果物/下水道使用料改定説明資料.docx',        '03_説明資料/下水道使用料改定説明資料.docx'),
    ('01_成果物/階上町下水道使用料改定_試算エビデンス.xlsx',
     '階上町下水道使用料改定_試算エビデンス.xlsx'),
    ('01_成果物/チェック結果.md',                      'チェック結果.md'),
    ('01_成果物/R8予算と経営戦略の差異分析.md',        '04_R8予算と経営戦略の差異分析.md'),
    ('02_図',                                          '03_説明資料/figures'),
    ('03_根拠データ',                                  'evidence/data'),
    ('04_スクリプト/recompute.py',                     'recompute.py'),
    ('04_スクリプト/check.py',                         'check.py'),
    ('04_スクリプト/build_xlsx.py',                    'build_xlsx.py'),
    ('04_スクリプト/mkfigs.py',                        'mkfigs.py'),
    ('04_スクリプト/build.js',                         '03_説明資料/build.js'),
    ('05_ログ/check_results.json',                     '03_説明資料/check_results.json'),
    ('05_ログ/analyze_choutei_実行ログ.txt',           'evidence/analyze_choutei_実行ログ.txt'),
    ('05_ログ/compare_plans_実行ログ.txt',             'evidence/compare_plans_実行ログ.txt'),
    ('06_旧版（参考）/01_現状確認メモ.md',             '01_現状確認メモ.md'),
    ('06_旧版（参考）/02_打ち合わせ用サマリー.html',   '02_打ち合わせ用サマリー.html'),
    ('README.md',                                      'evidence/README.md'),
    ('再現手順.md',                                    '再現手順.md'),
]


def pairs():
    for dest, src in LAYOUT:
        full = os.path.join(HERE, src)
        if os.path.isdir(full):
            for root, dirs, files in os.walk(full):
                dirs[:] = sorted(d for d in dirs if d != '__pycache__')
                for fn in sorted(files):
                    if fn.startswith('.'):
                        continue
                    p = os.path.join(root, fn)
                    yield os.path.join(dest, os.path.relpath(p, full)).replace(os.sep, '/'), p
        elif os.path.exists(full):
            yield dest, full
        else:
            raise SystemExit('見つかりません: %s' % src)


def main():
    items = list(pairs())
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
        for dest, src in items:
            z.write(src, dest)
    print('bundle:', OUT)
    print('files :', len(items))
    for dest, _ in items:
        print('  ', dest)


if __name__ == '__main__':
    main()
