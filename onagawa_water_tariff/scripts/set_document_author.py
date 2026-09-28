# -*- coding: utf-8 -*-
"""成果品（Word/Excel）の作成者・最終保存者を統一する。

zip の docProps/core.xml のみを書き換え、他のエントリはそのまま複製するため、
数式・書式・結合セルには一切触れない。
受領資料（01_source_evidence および町から受領した調査票）は対象外。
"""
import os, re, shutil, sys, zipfile, datetime

AUTHOR  = '若山諒太'
COMPANY = 'ビズアップ公共コンサルティング株式会社'

TARGETS = [
    '02_deliverables/01_rate_reform_simulation_MAIN.xlsx',
    '02_deliverables/02_total_cost_simulation.xlsx',
    '02_deliverables/03_tariff_period_cost_calc.xlsx',
    '02_deliverables/04_tariff_1m3_unit_by_diameter.xlsx',
    '02_deliverables/05_volume_zone_by_wateramount.xlsx',
    '02_deliverables/06_volume_zone_analysis_R1toR7.xlsx',
    '02_deliverables/07_zone_analysis_R1toR6.xlsx',
    '02_deliverables/08_billing_summary_R1toR7_v2.xlsx',
    '04_prefecture_unification/01_final_deliverables/女川町水道料金_比較検討資料.docx',
    '04_prefecture_unification/01_final_deliverables/宮城県水道料金体系統一化_検討報告書.docx',
    '04_prefecture_unification/01_final_deliverables/宮城県水道料金体系統一化_検討資料.docx',
    '04_prefecture_unification/01_final_deliverables/料金詳細確認_確定版.docx',
    '04_prefecture_unification/01_final_deliverables/水産加工業影響_定量化資料.docx',
    '04_prefecture_unification/01_final_deliverables/県内臨海団体比較_影響試算.docx',
    '06_capital_plan_update/09_assumption_switch_cost_calc.xlsx',
    '08_word_revisions/mgmt_strategy_report_draft_revised.docx',
    '09_interim_report/女川町水道料金改定_中間報告書.docx',
    '09_interim_report/女川町水道料金改定_中間報告書_指摘事項対応表.docx',
    '09_interim_report/女川町水道料金改定_最終報告に向けた工程表.docx',
    '09_interim_report/onagawa_adopted_calc.xlsx',
    '10_wbs/onagawa_wbs_R8.xlsx',
    '10_wbs/女川町上下水道経営指標評価_WBSと進捗状況.docx',
]

CORE_TMPL = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>{title}</dc:title><dc:subject>令和8年度上下水道経営指標評価書作成等業務委託</dc:subject><dc:creator>{author}</dc:creator><cp:lastModifiedBy>{author}</cp:lastModifiedBy><cp:revision>1</cp:revision><dcterms:created xsi:type="dcterms:W3CDTF">{ts}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{ts}</dcterms:modified><cp:category>{company}</cp:category></cp:coreProperties>'''

def esc(s):
    return (s.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;'))

def existing_title(data):
    m = re.search(r'<dc:title>(.*?)</dc:title>', data.decode('utf-8', 'ignore'), re.S)
    return m.group(1) if m and m.group(1).strip() else ''

def rewrite(path, ts):
    title = ''
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if 'docProps/core.xml' in names:
            title = existing_title(z.read('docProps/core.xml'))
    if not title:
        title = esc(os.path.splitext(os.path.basename(path))[0])
    core = CORE_TMPL.format(title=title, author=esc(AUTHOR), company=esc(COMPANY), ts=ts)
    tmp = path + '.tmp'
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        wrote = False
        for item in zin.infolist():
            if item.filename == 'docProps/core.xml':
                zout.writestr(item, core.encode('utf-8')); wrote = True
            else:
                zout.writestr(item, zin.read(item.filename))
        if not wrote:
            zout.writestr('docProps/core.xml', core.encode('utf-8'))
    shutil.move(tmp, path)
    return title

def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    ts = datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ')
    for rel in TARGETS:
        p = os.path.join(root, rel)
        if not os.path.exists(p):
            print('見つかりません:', rel); continue
        t = rewrite(p, ts)
        print(f'{AUTHOR} に設定: {rel}')

if __name__ == '__main__':
    main()
