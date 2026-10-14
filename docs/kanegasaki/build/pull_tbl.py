# -*- coding: utf-8 -*-
"""ch_soan.py（協議用）から TBL ブロックを行番号で抜き出す。

　公表用（ch_kohyo.py）へ表を引き継ぐときに、手で打ち直さないための道具である。
"""
import io
import sys


def pull(path, start):
    """start 行（1起点・'TBL([' の行）から対応する '], widths=...)' までを返す。"""
    L = io.open(path, encoding='utf-8').read().split('\n')
    i = start - 1
    assert L[i].strip().startswith('TBL(['), (start, L[i][:40])
    out = []
    depth = 0
    while i < len(L):
        out.append(L[i])
        depth += L[i].count('(') - L[i].count(')')
        if depth == 0 and len(out) > 1:
            break
        i += 1
    return '\n'.join(out)


if __name__ == '__main__':
    print(pull(sys.argv[1], int(sys.argv[2])))
