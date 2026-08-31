// 第10期計画 素案 Word生成（soan_content.py が出力した /tmp/soan.json を読む）
const fs = require('fs');
const d = require('/tmp/node_modules/docx');
const {Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, PageBreak,
       Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, LevelFormat,
       TableOfContents, Header, Footer, PageNumber} = d;

const C = JSON.parse(fs.readFileSync('/tmp/soan.json', 'utf8'));
const FONT = '游明朝';
const FONTG = '游ゴシック';
const NAVY = '1F3864', BLUE = '2E75B6', BAND = 'DDEBF7', NOTE = 'FFF3F3', GREY = '595959';
const TBLW = 9360;   // A4 縦 本文幅（DXA）

const p = (text, o = {}) => new Paragraph({
  spacing: {after: o.after ?? 120, line: o.line ?? 300},
  alignment: o.align, indent: o.indent,
  border: o.border,
  shading: o.shade ? {type: ShadingType.CLEAR, fill: o.shade} : undefined,
  children: [new TextRun({text, font: o.font ?? FONT, size: o.size ?? 21,
                          bold: o.bold, color: o.color})],
});

function table(head, rows, widths) {
  const cols = widths.map(w => Math.round(TBLW * w / 100));
  const diff = TBLW - cols.reduce((a, b) => a + b, 0);
  cols[cols.length - 1] += diff;
  const cell = (txt, i, isHead) => new TableCell({
    width: {size: cols[i], type: WidthType.DXA},
    shading: {type: ShadingType.CLEAR, fill: isHead ? BLUE : 'FFFFFF'},
    margins: {top: 60, bottom: 60, left: 90, right: 90},
    children: [new Paragraph({
      spacing: {after: 0, line: 240},
      alignment: isHead ? AlignmentType.CENTER : undefined,
      children: [new TextRun({text: String(txt), font: FONTG, size: 17,
                              bold: isHead, color: isHead ? 'FFFFFF' : '000000'})],
    })],
  });
  return new Table({
    columnWidths: cols,
    width: {size: TBLW, type: WidthType.DXA},
    rows: [new TableRow({tableHeader: true, children: head.map((h, i) => cell(h, i, true))}),
           ...rows.map(r => new TableRow({children: r.map((v, i) => cell(v, i, false))}))],
  });
}

const kids = [];
// ── 表紙 ──
kids.push(p('', {after: 2400}));
kids.push(p(C.title,  {align: AlignmentType.CENTER, size: 36, bold: true, font: FONTG, color: NAVY, after: 180}));
kids.push(p(C.title2, {align: AlignmentType.CENTER, size: 36, bold: true, font: FONTG, color: NAVY, after: 600}));
kids.push(p(`計画期間　${C.subtitle}`, {align: AlignmentType.CENTER, size: 24, font: FONTG, after: 1800}));
kids.push(p(C.draft,  {align: AlignmentType.CENTER, size: 44, bold: true, font: FONTG, after: 2400}));
kids.push(p(C.date,   {align: AlignmentType.CENTER, size: 24, font: FONTG, after: 120}));
kids.push(p(C.issuer, {align: AlignmentType.CENTER, size: 28, bold: true, font: FONTG, after: 0}));
kids.push(new Paragraph({children: [new PageBreak()]}));

// ── 凡例 ──
kids.push(p('本書の見方', {size: 28, bold: true, font: FONTG, color: NAVY, after: 240}));
kids.push(p('本書は素案です。計画の確定にあたり、次の表記は解消または削除します。', {after: 240}));
kids.push(table(['表記', '意味'], [
  ['【要確認】', '北塩原村との協議により確定する事項です。'],
  ['【要設定】', '目標値等を今後設定する箇所です。'],
  ['⚙ 編集注記', '素案段階の申し送りです。計画確定時に削除します。'],
], [22, 78]));
kids.push(p('', {after: 240}));
kids.push(new Paragraph({children: [new PageBreak()]}));

// ── 目次 ──
kids.push(p('目　次', {align: AlignmentType.CENTER, size: 32, bold: true, font: FONTG, color: NAVY, after: 360}));
kids.push(new TableOfContents('目次', {hyperlink: true, headingStyleRange: '1-2'}));
kids.push(new Paragraph({children: [new PageBreak()]}));

// ── 本文 ──
C.chapters.forEach((ch, ci) => {
  if (ci > 0) kids.push(new Paragraph({children: [new PageBreak()]}));
  kids.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: {before: 0, after: 300},
    shading: {type: ShadingType.CLEAR, fill: NAVY},
    children: [new TextRun({text: `${ch.no}　${ch.title}`, font: FONTG, size: 30,
                            bold: true, color: 'FFFFFF'})],
  }));
  ch.sections.forEach(sec => {
    kids.push(new Paragraph({
      heading: HeadingLevel.HEADING_2,
      spacing: {before: 320, after: 180},
      border: {bottom: {style: BorderStyle.SINGLE, size: 12, color: BLUE, space: 4}},
      children: [new TextRun({text: `${sec.no}　${sec.title}`, font: FONTG, size: 24,
                              bold: true, color: NAVY})],
    }));
    sec.blocks.forEach(b => {
      if (b.t === 'p') kids.push(p(b.v));
      else if (b.t === 'h3') kids.push(p(b.v, {size: 21, bold: true, font: FONTG,
                                               color: BLUE, after: 100}));
      else if (b.t === 'bullets') b.v.forEach(x => kids.push(new Paragraph({
        numbering: {reference: 'bul', level: 0}, spacing: {after: 60, line: 300},
        children: [new TextRun({text: x, font: FONT, size: 21})],
      })));
      else if (b.t === 'note') {
        kids.push(new Paragraph({
          spacing: {before: 100, after: 160, line: 280},
          indent: {left: 200, right: 200},
          shading: {type: ShadingType.CLEAR, fill: NOTE},
          border: {left: {style: BorderStyle.SINGLE, size: 18, color: 'C00000', space: 8}},
          children: [new TextRun({text: `⚙ 編集注記：${b.v}`, font: FONTG, size: 18, color: GREY})],
        }));
      } else if (b.t === 'table' || b.t === 'kpi') {
        kids.push(table(b.head, b.rows, b.widths));
        kids.push(p('', {after: 160}));
      }
    });
  });
});

const doc = new Document({
  creator: 'ビズアップ公共コンサルティング株式会社',
  title: `${C.title}・${C.title2}　${C.draft}`,
  styles: {default: {document: {run: {font: FONT, size: 21}}}},
  numbering: {config: [{reference: 'bul', levels: [{
    level: 0, format: LevelFormat.BULLET, text: '●', alignment: AlignmentType.LEFT,
    style: {paragraph: {indent: {left: 420, hanging: 210}},
            run: {font: FONTG, size: 16, color: BLUE}},
  }]}]},
  sections: [{
    properties: {page: {margin: {top: 1418, right: 1134, bottom: 1418, left: 1134}}},
    headers: {default: new Header({children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      border: {bottom: {style: BorderStyle.SINGLE, size: 6, color: 'BFBFBF', space: 2}},
      children: [new TextRun({text: `${C.title}・${C.title2}　${C.draft}`,
                              font: FONTG, size: 16, color: GREY})],
    })]})},
    footers: {default: new Footer({children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({children: [PageNumber.CURRENT], font: FONTG, size: 18, color: GREY})],
    })]})},
    children: kids,
  }],
});

Packer.toBuffer(doc).then(buf => {
  const out = '/home/user/repository/output/06_北塩原村第10期_計画素案.docx';
  fs.writeFileSync(out, buf);
  console.log('保存:', out, (buf.length / 1024).toFixed(0) + ' KB');
});
