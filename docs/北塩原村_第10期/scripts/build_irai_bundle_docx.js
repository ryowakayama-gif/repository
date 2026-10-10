// 照会票（束ごと）Word生成。build_irai_bundle.py が出した output/_build/irai_bundle.json を読む。
//   node scripts/build_irai_bundle_docx.js 1    → 束1
//   node scripts/build_irai_bundle_docx.js      → 束1〜束6 をまとめて
const fs = require('fs');
const path = require('path');
// 置き場所は自分の位置から数える（じか書きしない）。
const ROOT = path.resolve(__dirname, '..', '..', '..');
const OUTDIR = path.join(ROOT, 'output');
const d = require(path.join(__dirname, 'docxlib.js'));
const {Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, PageBreak,
       Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, Footer,
       PageNumber, TableLayoutType, LineRuleType} = d;

const C = JSON.parse(fs.readFileSync(path.join(OUTDIR, '_build', 'irai_bundle.json'), 'utf8'));
const FONT = '游明朝', FONTG = '游ゴシック';
const NAVY = '1F3864', BLUE = '2E75B6', NOTE = 'FFF3F3', KEYB = 'FFF2CC', GREY = '595959';
const TBLW = 9360;
// 部ごとの網掛け：第1部（停止）だけを目立たせる
const PARTFILL = {1: KEYB};

const p = (text, o = {}) => new Paragraph({
  spacing: {after: o.after ?? 120, line: o.line ?? 300, lineRule: LineRuleType.EXACT},
  alignment: o.align, indent: o.indent, keepNext: o.keepNext,
  children: [new TextRun({text, font: o.font ?? FONT, size: o.size ?? 21,
                          bold: o.bold, color: o.color})],
});

// ご回答の欄は空の罫線。村が書き入れるため、行の高さを確保する。
function table(head, rows, widths, fill) {
  const cols = widths.map(w => Math.round(TBLW * w / 100));
  cols[cols.length - 1] += TBLW - cols.reduce((a, b) => a + b, 0);
  const cell = (txt, i, isHead, isAns) => new TableCell({
    width: {size: cols[i], type: WidthType.DXA},
    shading: {type: ShadingType.CLEAR,
              fill: isHead ? BLUE : (isAns ? 'FFFFFF' : (fill || 'FFFFFF'))},
    margins: {top: 70, bottom: 70, left: 90, right: 90},
    children: [new Paragraph({
      spacing: {after: 0, line: 250, lineRule: LineRuleType.EXACT},
      alignment: isHead ? AlignmentType.CENTER : undefined,
      children: [new TextRun({text: String(txt), font: FONTG, size: 17,
                              bold: isHead, color: isHead ? 'FFFFFF' : '000000'})],
    })],
  });
  return new Table({
    layout: TableLayoutType.FIXED,
    columnWidths: cols, width: {size: TBLW, type: WidthType.DXA},
    rows: [new TableRow({tableHeader: true, cantSplit: true,
                         children: head.map((h, i) => cell(h, i, true, false))}),
           ...rows.map(r => new TableRow({
             cantSplit: true,
             height: {value: 700, rule: 'atLeast'},   // ご回答を書き込める高さ
             children: r.map((v, i) => cell(v, i, false, i === r.length - 1)),
           }))],
  });
}

function build(b) {
  const kids = [];
  kids.push(p(C.meta.date, {align: AlignmentType.RIGHT, size: 21, font: FONTG, after: 200}));
  kids.push(p(C.meta.atesaki, {size: 24, bold: true, font: FONTG, after: 120}));
  kids.push(p(C.meta.issuer, {align: AlignmentType.RIGHT, size: 21, font: FONTG, after: 400}));
  kids.push(p(C.meta.title, {align: AlignmentType.CENTER, size: 26, bold: true, font: FONTG, color: NAVY, after: 60}));
  kids.push(p(C.meta.title2, {align: AlignmentType.CENTER, size: 26, bold: true, font: FONTG, color: NAVY, after: 120}));
  kids.push(p(`${C.meta.subtitle}　${b.no}　${b.title}`,
              {align: AlignmentType.CENTER, size: 30, bold: true, font: FONTG, color: NAVY, after: 320}));
  kids.push(p(`ご確認・ご提供をお願いしたい事項　${b.n}件`,
              {align: AlignmentType.CENTER, size: 22, font: FONTG, color: GREY, after: 280}));
  kids.push(p('このまとまりのねらい', {size: 22, bold: true, font: FONTG, color: NAVY, after: 100, keepNext: true}));
  kids.push(p(b.nerai, {after: 240}));
  kids.push(p('ご回答の方法', {size: 22, bold: true, font: FONTG, color: NAVY, after: 100, keepNext: true}));
  C.goannai.forEach((t, i) => kids.push(p(`${i + 1}　${t}`, {after: 100, indent: {left: 120}})));
  kids.push(p('', {after: 160}));
  kids.push(table(['部', '内容', '件数'],
    b.parts.map(s => [s.title.split('　')[0],
                      s.title.split('　').slice(1).join('　'),
                      `${s.rows.length}件`]),
    [12, 70, 18], null));

  b.parts.forEach(s => {
    kids.push(new Paragraph({children: [new PageBreak()]}));
    kids.push(new Paragraph({
      heading: HeadingLevel.HEADING_1, spacing: {before: 0, after: 240},
      shading: {type: ShadingType.CLEAR, fill: NAVY}, keepNext: true,
      children: [new TextRun({text: s.title, font: FONTG, size: 26, bold: true, color: 'FFFFFF'})],
    }));
    if (s.lead) kids.push(new Paragraph({
      spacing: {before: 60, after: 220, line: 280, lineRule: LineRuleType.EXACT}, indent: {left: 200, right: 200},
      shading: {type: ShadingType.CLEAR, fill: NOTE}, keepNext: true,
      border: {left: {style: BorderStyle.SINGLE, size: 18, color: 'C00000', space: 8}},
      children: [new TextRun({text: s.lead, font: FONTG, size: 18, color: GREY})],
    }));
    const rows = s.rows.map((r, i) => [
      String(i + 1), r[0], r[1], r[2], r[4] === '一部回答' ? '（一部ご回答済）' : '',
    ]);
    kids.push(table(['#', 'ご確認・ご提供をお願いしたいこと', '当方の現在の扱い', 'ご回答が必要な時期', 'ご回答'],
                    rows, [4, 34, 29, 13, 20], PARTFILL[s.lv]));
  });

  kids.push(new Paragraph({children: [new PageBreak()]}));
  kids.push(p('お問い合わせ・ご返送先', {size: 24, bold: true, font: FONTG, color: NAVY, after: 200}));
  kids.push(p('ビズアップ公共コンサルティング株式会社', {size: 21, font: FONTG, after: 60}));
  kids.push(p('第10期北塩原村高齢者福祉計画・介護保険事業計画策定業務　担当', {size: 21, font: FONTG, after: 240}));
  kids.push(p('本票の「#」は、各部の中での通し番号です。ご回答の際は「部」と「#」をお示しいただけますと助かります。', {after: 120}));
  kids.push(p('様式や記入の仕方についてのご相談も承ります。お手元にない資料は「なし」とご記入いただければ、当方で代わりの手だてを検討します。', {after: 120}));

  return new Document({
    // 既定の書体と大きさを宣言する（宣言しないと Word の既定が段落記号に効く）
    styles: {default: {document: {run: {font: FONT, size: 21}},
                       heading1: {run: {font: FONTG, size: 26, bold: true, color: 'FFFFFF'}}}},
    sections: [{
      properties: {page: {margin: {top: 1134, bottom: 1134, left: 1134, right: 1134}}},
      footers: {default: new Footer({children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({children: [`${b.no}　`, PageNumber.CURRENT],
                                font: FONTG, size: 18, color: GREY})],
      })]})},
      children: kids,
    }],
  });
}

const want = process.argv[2];
const targets = want ? C.bundles.filter(b => b.no === '束' + want || b.no === want) : C.bundles;
if (!targets.length) {
  console.error('束が見つかりません： ' + want + '（あるのは ' + C.bundles.map(b => b.no).join('・') + '）');
  process.exit(1);
}
(async () => {
  for (const b of targets) {
    const n = b.no.replace('束', '');
    const out = path.join(OUTDIR, `13-${n}_北塩原村第10期_照会票_${b.no}_${b.title}.docx`);
    const buf = await Packer.toBuffer(build(b));
    fs.writeFileSync(out, buf);
    console.log('保存: ' + path.basename(out) + ' ' + Math.round(buf.length / 1024) + ' KB');
  }
})();
