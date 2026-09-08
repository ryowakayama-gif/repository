// 資料提供・確認依頼書 Word生成（irai_content.py が出力した /tmp/irai.json を読む）
const fs = require('fs');
const d = require('/tmp/node_modules/docx');
const {Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, PageBreak,
       Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, Footer, PageNumber} = d;

const C = JSON.parse(fs.readFileSync('/tmp/irai.json', 'utf8'));
const FONT = '游明朝', FONTG = '游ゴシック';
const NAVY = '1F3864', BLUE = '2E75B6', NOTE = 'FFF3F3', KEYB = 'FFF2CC', GREY = '595959';
const TBLW = 9360;

const p = (text, o = {}) => new Paragraph({
  spacing: {after: o.after ?? 120, line: o.line ?? 300},
  alignment: o.align, indent: o.indent,
  children: [new TextRun({text, font: o.font ?? FONT, size: o.size ?? 21,
                          bold: o.bold, color: o.color})],
});

function table(head, rows, widths) {
  const cols = widths.map(w => Math.round(TBLW * w / 100));
  cols[cols.length - 1] += TBLW - cols.reduce((a, b) => a + b, 0);
  const cell = (txt, i, isHead, star) => new TableCell({
    width: {size: cols[i], type: WidthType.DXA},
    shading: {type: ShadingType.CLEAR, fill: isHead ? BLUE : (star ? KEYB : 'FFFFFF')},
    margins: {top: 70, bottom: 70, left: 90, right: 90},
    children: [new Paragraph({
      spacing: {after: 0, line: 250},
      alignment: isHead ? AlignmentType.CENTER : undefined,
      children: [new TextRun({text: String(txt), font: FONTG, size: 17,
                              bold: isHead, color: isHead ? 'FFFFFF' : '000000'})],
    })],
  });
  return new Table({
    columnWidths: cols, width: {size: TBLW, type: WidthType.DXA},
    rows: [new TableRow({tableHeader: true, children: head.map((h, i) => cell(h, i, true, false))}),
           ...rows.map(r => {
             const star = String(r[1] || '').startsWith('★');
             return new TableRow({children: r.map((v, i) => cell(v, i, false, star))});
           })],
  });
}

const kids = [];
kids.push(p(C.date, {align: AlignmentType.RIGHT, size: 21, font: FONTG, after: 200}));
kids.push(p(C.atesaki, {size: 24, bold: true, font: FONTG, after: 120}));
kids.push(p(C.issuer, {align: AlignmentType.RIGHT, size: 21, font: FONTG, after: 400}));
kids.push(p(C.title, {align: AlignmentType.CENTER, size: 26, bold: true, font: FONTG, color: NAVY, after: 60}));
kids.push(p(C.title2, {align: AlignmentType.CENTER, size: 26, bold: true, font: FONTG, color: NAVY, after: 120}));
kids.push(p(C.subtitle, {align: AlignmentType.CENTER, size: 30, bold: true, font: FONTG, color: NAVY, after: 320}));
C.lead.forEach(t => kids.push(p(t, {after: 140})));

// 目次代わりの一覧
kids.push(p('', {after: 100}));
kids.push(table(['区分', '件数', '期限'],
  C.sections.map(s => {
    const m = s.title.match(/【(.+?)】/);
    return [s.no + '　' + s.title.replace(/【.+?】/, ''), String(s.rows.length) + '件', m ? m[1] : ''];
  }), [58, 12, 30]));
kids.push(p('', {after: 200}));
kids.push(p('※ 黄色の網掛けと★印は、決まらないと後続の作業が止まる項目です。', {size: 19, font: FONTG, color: GREY, after: 120}));
kids.push(p('※ 番号は当方の管理番号です。ご回答の際に番号をお示しいただけますと助かります。', {size: 19, font: FONTG, color: GREY, after: 200}));

C.sections.forEach((s, i) => {
  kids.push(new Paragraph({children: [new PageBreak()]}));
  kids.push(new Paragraph({
    heading: HeadingLevel.HEADING_1, spacing: {before: 0, after: 240},
    shading: {type: ShadingType.CLEAR, fill: NAVY},
    children: [new TextRun({text: `${s.no}　${s.title}`, font: FONTG, size: 26, bold: true, color: 'FFFFFF'})],
  }));
  if (s.lead) kids.push(p(s.lead, {after: 160}));
  if (s.note) kids.push(new Paragraph({
    spacing: {before: 60, after: 200, line: 280}, indent: {left: 200, right: 200},
    shading: {type: ShadingType.CLEAR, fill: NOTE},
    border: {left: {style: BorderStyle.SINGLE, size: 18, color: 'C00000', space: 8}},
    children: [new TextRun({text: '※ ' + s.note, font: FONTG, size: 18, color: GREY})],
  }));
  kids.push(table(s.head, s.rows, s.widths));
});

kids.push(new Paragraph({children: [new PageBreak()]}));
kids.push(p('お問い合わせ先', {size: 24, bold: true, font: FONTG, color: NAVY, after: 200}));
kids.push(p('ビズアップ公共コンサルティング株式会社', {size: 21, font: FONTG, after: 60}));
kids.push(p('第10期北塩原村高齢者福祉計画・介護保険事業計画策定業務　担当', {size: 21, font: FONTG, after: 240}));
kids.push(p('本依頼書の各項目には管理番号を付しています。ご回答は項目ごとでも、まとめてでも結構です。ご不明な点や、様式・形式についてのご相談も承ります。', {after: 120}));

const doc = new Document({
  styles: {default: {heading1: {run: {font: FONTG, size: 26, bold: true, color: 'FFFFFF'}}}},
  sections: [{
    properties: {page: {margin: {top: 1134, bottom: 1134, left: 1134, right: 1134}}},
    footers: {default: new Footer({children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({children: [PageNumber.CURRENT], font: FONTG, size: 18, color: GREY})],
    })]})},
    children: kids,
  }],
});

const OUT = '/home/user/repository/output/09_北塩原村第10期_資料提供確認依頼書.docx';
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log('保存: ' + OUT + ' ' + Math.round(buf.length / 1024) + ' KB');
});
