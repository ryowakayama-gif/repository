// 第2回策定委員会 資料 Word生成（shiryo_content.py が出力した /tmp/shiryo.json を読む）
const fs = require('fs');
const d = require('/tmp/node_modules/docx');
const {Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, PageBreak,
       Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, LevelFormat,
       TableOfContents, Footer, PageNumber} = d;

const C = JSON.parse(fs.readFileSync('/tmp/shiryo.json', 'utf8'));
const FONT = '游明朝', FONTG = '游ゴシック';
const NAVY = '1F3864', BLUE = '2E75B6', BAND = 'DDEBF7', NOTE = 'FFF3F3', KEYB = 'FFF2CC', GREY = '595959';
const TBLW = 9360;

const p = (text, o = {}) => new Paragraph({
  spacing: {after: o.after ?? 120, line: o.line ?? 300},
  alignment: o.align, indent: o.indent, border: o.border,
  shading: o.shade ? {type: ShadingType.CLEAR, fill: o.shade} : undefined,
  children: [new TextRun({text, font: o.font ?? FONT, size: o.size ?? 21,
                          bold: o.bold, color: o.color})],
});

function table(head, rows, widths) {
  const cols = widths.map(w => Math.round(TBLW * w / 100));
  cols[cols.length - 1] += TBLW - cols.reduce((a, b) => a + b, 0);
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
    columnWidths: cols, width: {size: TBLW, type: WidthType.DXA},
    rows: [new TableRow({tableHeader: true, children: head.map((h, i) => cell(h, i, true))}),
           ...rows.map(r => new TableRow({children: r.map((v, i) => cell(v, i, false))}))],
  });
}

const kids = [];
// 表紙
kids.push(p('', {after: 2600}));
kids.push(p(C.title,  {align: AlignmentType.CENTER, size: 32, bold: true, font: FONTG, color: NAVY, after: 160}));
kids.push(p(C.title2, {align: AlignmentType.CENTER, size: 32, bold: true, font: FONTG, color: NAVY, after: 700}));
kids.push(p(C.subtitle, {align: AlignmentType.CENTER, size: 24, font: FONTG, after: 1600}));
kids.push(p('資料2　第9期計画の進捗と評価', {align: AlignmentType.CENTER, size: 24, font: FONTG, after: 100}));
kids.push(p('資料5　サービス見込み量の考え方（案）', {align: AlignmentType.CENTER, size: 24, font: FONTG, after: 100}));
kids.push(p('資料6　介護保険料の概算シミュレーション', {align: AlignmentType.CENTER, size: 24, font: FONTG, after: 100}));
kids.push(p('資料7　成果目標・活動指標（案）', {align: AlignmentType.CENTER, size: 24, font: FONTG, after: 1600}));
kids.push(p(C.date,   {align: AlignmentType.CENTER, size: 24, font: FONTG, after: 120}));
kids.push(p(C.issuer, {align: AlignmentType.CENTER, size: 28, bold: true, font: FONTG, after: 0}));
kids.push(new Paragraph({children: [new PageBreak()]}));

// 本書の見方
kids.push(p('本資料の見方', {size: 28, bold: true, font: FONTG, color: NAVY, after: 240}));
kids.push(p('本資料は、令和8年11月に開催する第2回策定委員会の資料の一部です。資料3（アンケート調査結果）と資料4（基本理念・基本目標・施策体系）は、調査の集計後に作成します。', {after: 180}));
kids.push(table(['表記', '意味'], [
  ['【要確認】', '北塩原村への確認により確定する事項です。'],
  ['【要設定】', '目標値等を今後設定する箇所です。'],
  ['【試算中】／【推計中】', '国の推計ワークシートの提供後に算出する数値です。'],
  ['網掛けの囲み', '委員の皆様にご留意いただきたい点です。'],
  ['注記の囲み', '未確定の事項と、その確定の見通しです。'],
], [26, 74]));
kids.push(p('', {after: 200}));
kids.push(p('資料5・6は骨格です。数値は、国の推計ワークシートの提供（第10期基本指針の告示後）と、村の実績データの受領をもって確定します。本日は「考え方」についてご意見をいただきたく存じます。', {after: 240}));
kids.push(new Paragraph({children: [new PageBreak()]}));

// 目次
kids.push(p('目　次', {align: AlignmentType.CENTER, size: 32, bold: true, font: FONTG, color: NAVY, after: 360}));
kids.push(new TableOfContents('目次', {hyperlink: true, headingStyleRange: '1-2'}));
kids.push(new Paragraph({children: [new PageBreak()]}));

// 本文
C.chapters.forEach((ch, ci) => {
  if (ci > 0) kids.push(new Paragraph({children: [new PageBreak()]}));
  kids.push(new Paragraph({
    heading: HeadingLevel.HEADING_1, spacing: {before: 0, after: 300},
    shading: {type: ShadingType.CLEAR, fill: NAVY},
    children: [new TextRun({text: `${ch.no}　${ch.title}`, font: FONTG, size: 30, bold: true, color: 'FFFFFF'})],
  }));
  ch.sections.forEach(sec => {
    if (!sec.no.endsWith('-0')) {
      kids.push(new Paragraph({
        heading: HeadingLevel.HEADING_2, spacing: {before: 320, after: 180},
        border: {bottom: {style: BorderStyle.SINGLE, size: 12, color: BLUE, space: 4}},
        children: [new TextRun({text: `${sec.no}　${sec.title}`, font: FONTG, size: 24, bold: true, color: NAVY})],
      }));
    }
    sec.blocks.forEach(b => {
      if (b.t === 'p') kids.push(p(b.v));
      else if (b.t === 'h3') kids.push(p(b.v, {size: 21, bold: true, font: FONTG, color: BLUE, after: 100}));
      else if (b.t === 'bullets') b.v.forEach(x => kids.push(new Paragraph({
        numbering: {reference: 'bul', level: 0}, spacing: {after: 60, line: 300},
        children: [new TextRun({text: x, font: FONT, size: 21})],
      })));
      else if (b.t === 'key') kids.push(new Paragraph({
        spacing: {before: 140, after: 200, line: 300}, indent: {left: 200, right: 200},
        shading: {type: ShadingType.CLEAR, fill: KEYB},
        border: {left: {style: BorderStyle.SINGLE, size: 18, color: 'BF8F00', space: 8}},
        children: [new TextRun({text: b.v, font: FONTG, size: 20, bold: true})],
      }));
      else if (b.t === 'note') kids.push(new Paragraph({
        spacing: {before: 120, after: 180, line: 280}, indent: {left: 200, right: 200},
        shading: {type: ShadingType.CLEAR, fill: NOTE},
        border: {left: {style: BorderStyle.SINGLE, size: 18, color: 'C00000', space: 8}},
        children: [new TextRun({text: '※ ' + b.v, font: FONTG, size: 18, color: GREY})],
      }));
      else if (b.t === 'table') { kids.push(table(b.head, b.rows, b.widths)); kids.push(p('', {after: 160})); }
    });
  });
});

const doc = new Document({
  styles: {default: {
    heading1: {run: {font: FONTG, size: 30, bold: true, color: 'FFFFFF'}},
    heading2: {run: {font: FONTG, size: 24, bold: true, color: NAVY}},
  }},
  numbering: {config: [{reference: 'bul', levels: [{level: 0, format: LevelFormat.BULLET, text: '・',
    alignment: AlignmentType.LEFT, style: {paragraph: {indent: {left: 400, hanging: 200}}}}]}]},
  sections: [{
    properties: {page: {margin: {top: 1134, bottom: 1134, left: 1134, right: 1134}}},
    footers: {default: new Footer({children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({children: [PageNumber.CURRENT], font: FONTG, size: 18, color: GREY})],
    })]})},
    children: kids,
  }],
});

const OUT = '/home/user/repository/output/08_北塩原村第10期_第2回策定委員会資料.docx';
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log('保存: ' + OUT + ' ' + Math.round(buf.length / 1024) + ' KB');
});
