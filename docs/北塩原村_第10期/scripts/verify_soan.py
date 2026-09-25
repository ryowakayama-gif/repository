# -*- coding: utf-8 -*-
"""計画素案の数値の自己点検

【背景】
  素案（scripts/soan_content.py）の数値は、算定スクリプトの出力を書き写したもので
  ある。算定を改めたときに素案が追随しないと、両者が静かにずれる。
  他案件（大雪地区広域連合）では素案の生成が算定を直接読む形に改められたが、
  本村は当面この点検により乖離を検出する。

【使い方】
  python3 scripts/verify_soan.py
  すべて適合すれば終了コード0、1件でも不適合があれば1を返す。
"""
import csv, json, os, re, sys, subprocess

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, 'data')
MD = os.path.join(BASE, '18_計画素案.md')

RESULTS = []


def chk(no, name, ok, detail=''):
    RESULTS.append((no, name, ok, detail))


def num(v):
    try:
        return float(str(v).replace(',', '').replace('円', '').replace('人', '')
                     .replace('千', '').replace('%', '').replace('／月', '')
                     .replace('▲', '-').replace('＋', '').strip())
    except (TypeError, ValueError):
        return None


def has_num(t, v):
    """素案の本文に数値vが（3桁区切りの有無を問わず）現れるか"""
    n = num(v)
    if n is None:
        return False
    i = int(round(n))
    return f'{i:,}' in t or str(i) in t


def load_csv(name):
    path = os.path.join(DATA, name)
    if not os.path.exists(path):
        return None
    with open(path, encoding='utf-8-sig') as f:
        return list(csv.DictReader(f))


def md_text():
    with open(MD, encoding='utf-8') as f:
        return f.read()


def in_md(s, *needles):
    """素案の本文に needles がすべて含まれるか"""
    return all(n in s for n in needles)


def main():
    if not os.path.exists(MD):
        print('18_計画素案.md がありません。build_soan.py を先に実行してください。')
        return 1
    t = md_text()

    # ── 1〜3　人口と認定者数（案Bと素案5-2・5-3の一致）────────────
    cert = load_csv('第10期_要介護度別認定者数推計.csv') or []
    b = {r['年']: r for r in cert if r['案'] == 'B' and r['認定率の基準'] == '令和8年'}
    for i, y in enumerate(['令和9年', '令和10年', '令和11年'], start=1):
        r = b.get(y)
        if not r:
            chk(i, f'5-2 第1号被保険者数（{y}）', False, '算定CSVに行がない')
            continue
        ins, ninntei = r['第1号被保険者計'], r['認定者計']
        ok = has_num(t, ins) and has_num(t, ninntei)
        chk(i, f'5-2・5-3 {y}（第1号{int(num(ins)):,}人・認定者{ninntei}人）', ok)

    # ── 4　区分別受給者数（令和11年度）────────────────────
    kubun = load_csv('第10期_サービス区分別受給者数推計.csv') or []
    tot = {r['区分']: r['令和11年度'] for r in kubun if r['要介護度'] == '計'}
    ok = all(has_num(t, v) for v in tot.values()) and len(tot) == 3
    chk(4, '5-4 区分別受給者数の令和11年度（在宅%s・居住系%s・施設%s）'
        % (tot.get('在宅サービス'), tot.get('居住系サービス'), tot.get('施設サービス')), ok)

    # ── 5　サービス諸費と標準給付費の内部整合 ──────────────────
    svc = {'令和9年度': 263470, '令和10年度': 263505, '令和11年度': 262488}
    hojo = 21517
    ok = True
    detail = []
    for y, v in svc.items():
        std = v + hojo
        if not has_num(t, std):
            ok = False
            detail.append(f'{y}の標準給付費{std:,}が素案にない')
    if not has_num(t, sum(svc.values()) + hojo*3):
        ok = False
        detail.append('3か年計854,014が素案にない')
    chk(5, '5-6 標準給付費＝サービス諸費＋補助給付等（年度別と3か年計）', ok, '／'.join(detail))

    # ── 6　介護給付費＋予防給付費＝サービス諸費 ─────────────────
    split = load_csv('第10期_介護給付費_予防給付費の分離.csv') or []
    ok, detail = True, []
    for r in split:
        y = r['年度']
        if y not in svc:
            continue
        kaigo, yobo, total = num(r['介護給付費(千円)']), num(r['予防給付費(千円)']), num(r['サービス諸費(千円)'])
        if kaigo is None or yobo is None:
            continue
        if abs(kaigo + yobo - total) > 1:
            ok = False
            detail.append(f'{y}の内訳の合計がサービス諸費と合わない')
        if not has_num(t, kaigo) or not has_num(t, yobo):
            ok = False
            detail.append(f'{y}の内訳が素案にない')
    chk(6, '5-6 介護給付費＋予防給付費＝サービス諸費（素案と算定の一致）', ok, '／'.join(detail))

    # ── 7　地域支援事業費 ─────────────────────────
    ch = load_csv('第10期_地域支援事業費見込み.csv') or []
    row = next((r for r in ch if r['区分'] == '合計'), None)
    ok, detail = False, ''
    if row:
        vals = [row.get(k) for k in row if k.startswith('令和')]
        ok = all(has_num(t, v) for v in vals if v)
        three = sum(num(v) or 0 for v in vals)
        if not has_num(t, three):
            ok = False
            detail = f'3か年計{int(three):,}千円が素案にない'
    chk(7, '5-5(2) 地域支援事業費（年度別と3か年計）', ok, detail)

    # ── 8　年度別の保険料収納必要額 ──────────────────────
    yr = load_csv('第10期_保険料の年度別必要額.csv') or []
    ok, detail = bool(yr), []
    for r in yr:
        for key, lab in [('保険料収納必要額(千円)', '必要額'), ('月額(円)', '月額')]:
            v = num(r[key])
            if v is None:
                continue
            if not has_num(t, v):
                ok = False
                detail.append(f"{r['年度']}の{lab}{int(v):,}が素案にない")
    chk(8, '5-7 年度別の保険料収納必要額と月額', ok, '／'.join(detail))

    # ── 9　中長期の保険料 ─────────────────────────
    lt = load_csv('第10期_中長期保険料見通し.csv') or []
    ok, detail = bool(lt), []
    for r in lt:
        if not r['区分'].startswith('保険料基準額'):
            continue
        for k, v in r.items():
            if k == '区分' or not v:
                continue
            if not has_num(t, v):
                ok = False
                detail.append(f'{r["区分"]}の{k}={v}が素案にない')
    chk(9, '5-11 中長期の保険料基準額（αの2案）', ok, '／'.join(detail))

    # ── 10　村内定員の到達率 ───────────────────────
    sup = load_csv('第10期_村内定員の到達率.csv') or []
    ok = bool(sup)
    for r in sup:
        for k, v in r.items():
            if not k.endswith('の到達率(%)') or not v:
                continue
            if f'{float(v):.1f}%' not in t:
                ok = False
    chk(10, '5-4(5) 村内定員に対する到達率', ok)

    # ── 11　利用率の分母の検算が素案に反映されているか ───────────
    chk(11, '5-4 利用率の分母（要介護度別認定者数であること）',
        in_md(t, '要介護度別認定者数', '第1号被保険者数', '0件'))

    # ── 12　単価の趨勢と報酬改定率の二重計上の断り ────────────
    chk(12, '5-4(10)① 趨勢と改定率を重ねない旨の注記', in_md(t, '二重に数える'))

    # ── 13　推計の誤差と6-4の見直しの目安の整合 ─────────────
    chk(13, '6-4 見直しの目安（±12%が推計の誤差11.9%に対応）',
        in_md(t, '11.9%', '±12%超'))

    # ── 14　認知症基本法の条項 ──────────────────────
    chk(14, '4-7 認知症基本法の条項（第14条〜第21条の8項目）',
        in_md(t, '第14条から第21条までの8項目'))

    # ── 15　施策番号と節番号の衝突がないこと ─────────────────
    heads = re.findall(r'^## (\S+?)[　 ]', t, re.M)
    dup = {h for h in heads if heads.count(h) > 1}
    chk(15, '節番号の重複がないこと', not dup, '重複: ' + '・'.join(sorted(dup)) if dup else '')

    # ── 16　据え置き項目の一覧があること ───────────────────
    chk(16, '5-12 据え置き項目と確定の条件の一覧', in_md(t, '現時点で据え置いた項目と確定の条件'))

    # ── 17　3-5 施策体系の対応表が第4章の施策と一致すること ────────
    import soan_content as SC
    ch = {c["no"]: c for c in SC.CH}
    sis = []
    for sec in ch["第4章"]["sections"]:
        if sec["no"].startswith("施策"):
            ti = sec["title"]
            sis.append((f'{sec["no"].replace("施策", "")} {ti.split("　【")[0]}',
                        ti.split("【")[1].rstrip("】")))
        elif sec["no"] == "4-5":
            for b in sec["blocks"]:
                if b["t"] == "table":
                    sis += [(r[0], "新設") for r in b["rows"]]
    t35 = [x for x in ch["第3章"]["sections"] if x["no"] == "3-5"]
    if not t35:
        chk(17, '3-5 施策体系の対応表と第4章の施策の一致', False, '3-5 がない')
    else:
        rows = t35[0]["blocks"][1]["rows"]
        bad = [f'{a}≠{r[0]}' for (a, k), r in zip(sis, rows) if r[0] != a or r[1] != k]
        chk(17, '3-5 施策体系の対応表と第4章の施策の一致',
            len(rows) == len(sis) and not bad,
            f'件数 {len(rows)}／{len(sis)}・' + '・'.join(bad[:3]) if (bad or len(rows) != len(sis))
            else f'施策{len(sis)}本すべて一致')

    # ── 18　2-8 目標別の内訳が交付金の集計表の算定と一致すること ──────
    import csv as _csv
    f18 = '/home/user/repository/docs/北塩原村_第10期/data/交付金_目標別の県内比較_令和8年度.csv'
    if not os.path.exists(f18):
        chk(18, '2-8 目標別の内訳と交付金の算定の一致', False, 'CSVがない（parse_kofukin_shichoson.py を実行）')
    else:
        D = {r['区分']: r for r in _csv.DictReader(open(f18, encoding='utf-8-sig'))}
        t28 = None
        for sec in ch['第2章']['sections']:
            if sec['no'] == '2-8':
                for b in sec['blocks']:
                    if b['t'] == 'table' and b['head'] == ['交付金・目標', '本村', '県平均', '県内順位', '全国平均', '全国順位']:
                        t28 = b
        bad = []
        if t28 is None:
            bad.append('2-8の目標別表がない')
        else:
            for row in t28['rows']:
                kf = ('推進' if row[0].startswith('推進') else '支援') + row[0].split('目標')[1][0]
                d = D.get(kf)
                if not d:
                    bad.append(f'{kf}が算定にない'); continue
                for i, k in ((1, '本村'), (2, '県平均'), (3, '県内順位'), (4, '全国平均'), (5, '全国順位')):
                    got = row[i].replace('点', '').replace('位', '').replace(',', '')
                    exp = str(d[k])
                    if float(got) != float(exp):
                        bad.append(f'{kf} {k} 素案{row[i]}≠算定{d[k]}')
        chk(18, '2-8 目標別の内訳と交付金の算定の一致', not bad,
            '・'.join(bad[:4]) if bad else f'8目標×5項目すべて一致')

    # ── 出力 ─────────────────────────────
    w = max(len(n) for _, n, _, _ in RESULTS)
    print('■ 計画素案の自己点検')
    ng = 0
    for no, name, ok, detail in RESULTS:
        mark = '適合' if ok else '不適合'
        print(f'  {no:>3}  {name:<{w}}  {mark}' + (f'  {detail}' if detail else ''))
        if not ok:
            ng += 1
    print(f'\n  {len(RESULTS)}件のうち適合{len(RESULTS)-ng}件／不適合{ng}件')
    if ng:
        print('  不適合があります。算定を改めたら素案（scripts/soan_content.py）にも反映してください。')
    return 1 if ng else 0


if __name__ == '__main__':
    sys.exit(main())
