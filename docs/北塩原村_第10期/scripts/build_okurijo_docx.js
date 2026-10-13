// 障がい福祉計画担当への申し送り Word生成
//   build_okurijo.py が出した output/_build/okurijo.json を読む。
//
// 相手（障がい福祉計画担当）の申し送りの様式に合わせている。
//   ・節は ■ 全角数字
//   ・当方から渡す事項は、優先度・区分・事項・確かめた事実・お願い・
//     回答がない場合の扱い・由来の7列。相手の照会票に行として貼れる形にする。
//   ・表の最後の行が次の頁にひとりで残るのを防ぐ（寡婦行の処理）。
//     申し送りは頁数の枠がないため、行数によらずすべての表に掛ける。
const fs = require('fs');
const path = require('path');
// 置き場所は自分の位置から数える（じか書きしない）。
//   scripts/ → 北塩原村_第10期/ → docs/ → リポジトリの根
const ROOT = path.resolve(__dirname, '..', '..', '..');
const OUTDIR = path.join(ROOT, 'output');
const d = require(path.join(__dirname, 'docxlib.js'));
const {Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, PageBreak,
       Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, Footer, PageNumber,
       LineRuleType, TableLayoutType} = d;

const C = JSON.parse(fs.readFileSync(path.join(OUTDIR, '_build', 'okurijo.json'), 'utf8'));
const FONT = '游明朝', FONTG = '游ゴシック';
const NAVY = '1F3864', BLUE = '2E75B6', NOTE = 'FFF3F3', GREY = '595959';
const TBLW = 9360;   // 表の幅（本文幅ではない）

const p = (text, o = {}) => new Paragraph({
  spacing: {after: o.after ?? 120, line: o.line ?? 300, lineRule: LineRuleType.EXACT},
  alignment: o.align, indent: o.indent, keepNext: o.keepNext,
  children: [new TextRun({text, font: o.font ?? FONT, size: o.size ?? 21,
                          bold: o.bold, color: o.color})],
});

function table(head, rows, widths) {
  const cols = widths.map(w => Math.round(TBLW * w / 100));
  cols[cols.length - 1] += TBLW - cols.reduce((a, b) => a + b, 0);
  const cell = (txt, i, isHead, keep) => new TableCell({
    width: {size: cols[i], type: WidthType.DXA},
    shading: {type: ShadingType.CLEAR, fill: isHead ? BLUE : 'FFFFFF'},
    margins: {top: 60, bottom: 60, left: 90, right: 90},
    children: [new Paragraph({
      spacing: {after: 0, line: 240, lineRule: LineRuleType.EXACT},
      alignment: isHead ? AlignmentType.CENTER : undefined,
      keepNext: keep,
      children: [new TextRun({text: String(txt), font: FONTG, size: 16,
                              bold: isHead, color: isHead ? 'FFFFFF' : '000000'})],
    })],
  });
  return new Table({
    columnWidths: cols, width: {size: TBLW, type: WidthType.DXA},
    // 列幅を宣言どおりに固定する（既定の自動調整だと Word が中身に合わせて変える）
    layout: TableLayoutType.FIXED,
    rows: [new TableRow({tableHeader: true, cantSplit: true,
                         children: head.map((h, i) => cell(h, i, true, undefined))}),
           ...rows.map((r, ri) => new TableRow({cantSplit: true,
             children: r.map((v, i) => cell(v, i, false,
               (ri === rows.length - 2) || undefined))}))],
  });
}

// 事項ごとの縦組み。7列の表にすると1列が数文字になって読めない。
// 見出しの帯を1行で渡し、下に「札｜中身」の2列で並べる。
function jikou(midashi, rows) {
  const cols = [Math.round(TBLW * 0.22), 0];
  cols[1] = TBLW - cols[0];
  const head = new TableRow({
    tableHeader: true, cantSplit: true,
    children: [new TableCell({
      columnSpan: 2, width: {size: TBLW, type: WidthType.DXA},
      shading: {type: ShadingType.CLEAR, fill: NAVY},
      margins: {top: 70, bottom: 70, left: 110, right: 110},
      children: [new Paragraph({
        spacing: {after: 0, line: 260, lineRule: LineRuleType.EXACT}, keepNext: true,
        children: [new TextRun({text: midashi, font: FONTG, size: 19,
                                bold: true, color: 'FFFFFF'})],
      })],
    })],
  });
  const cell = (txt, i, keep) => new TableCell({
    width: {size: cols[i], type: WidthType.DXA},
    shading: {type: ShadingType.CLEAR, fill: i === 0 ? 'EAF0F8' : 'FFFFFF'},
    margins: {top: 60, bottom: 60, left: 110, right: 110},
    children: [new Paragraph({
      spacing: {after: 0, line: 260, lineRule: LineRuleType.EXACT}, keepNext: keep,
      children: [new TextRun({text: String(txt), font: FONTG, size: 17,
                              bold: i === 0, color: i === 0 ? NAVY : '000000'})],
    })],
  });
  return new Table({
    columnWidths: cols, width: {size: TBLW, type: WidthType.DXA},
    layout: TableLayoutType.FIXED,
    rows: [head, ...rows.map((r, ri) => new TableRow({
      cantSplit: true,
      children: r.map((v, i) => cell(v, i, (ri === rows.length - 2) || undefined)),
    }))],
  });
}

const kids = [];
// ── 送り状 ──────────────────────────────────────────────
kids.push(p(C.meta.date, {align: AlignmentType.RIGHT, size: 21, font: FONTG, after: 200}));
kids.push(p(C.meta.atesaki, {size: 24, bold: true, font: FONTG, after: 120}));
kids.push(p(C.meta.issuer, {align: AlignmentType.RIGHT, size: 21, font: FONTG, after: 400}));
kids.push(p(C.meta.title, {align: AlignmentType.CENTER, size: 24, bold: true, font: FONTG, color: NAVY, after: 60}));
kids.push(p(C.meta.title2, {align: AlignmentType.CENTER, size: 24, bold: true, font: FONTG, color: NAVY, after: 120}));
kids.push(p(C.meta.subtitle, {align: AlignmentType.CENTER, size: 30, bold: true, font: FONTG, color: NAVY, after: 320}));
C.maegaki.forEach(t => kids.push(p(t, {after: 140})));

// 目次代わりの一覧
kids.push(p('', {after: 100}));
kids.push(table(['節', '件数'],
  C.sections.map(s => [s.no + '　' + s.title,
                       String((s.items || s.rows).length) + '件']),
  [80, 20]));
kids.push(p('', {after: 160}));
kids.push(p('※ 優先度は貴側の様式に合わせております（Ａ・Ｂ）。'
            + '「―」は当方からのご回答・お知らせで、ご回答を要しないものです。',
            {size: 19, font: FONTG, color: GREY, after: 120}));

// ── 各節 ───────────────────────────────────────────────
C.sections.forEach(s => {
  kids.push(new Paragraph({children: [new PageBreak()]}));
  kids.push(new Paragraph({
    heading: HeadingLevel.HEADING_1, spacing: {before: 0, after: 240},
    shading: {type: ShadingType.CLEAR, fill: NAVY}, keepNext: true,
    children: [new TextRun({text: `■　${s.no}　${s.title}`,
                            font: FONTG, size: 26, bold: true, color: 'FFFFFF'})],
  }));
  if (s.lead) kids.push(p(s.lead, {after: 160, keepNext: true}));
  if (s.kind === 'jikou') {
    s.items.forEach((it, i) => {
      if (i) kids.push(p('', {after: 140}));
      kids.push(jikou(it.midashi, it.rows));
    });
  } else {
    kids.push(table(s.head, s.rows, s.widths));
  }
  if (s.note) {
    kids.push(p('', {after: 80}));
    kids.push(new Paragraph({
      spacing: {before: 60, after: 200, line: 280, lineRule: LineRuleType.EXACT},
      indent: {left: 200, right: 200},
      shading: {type: ShadingType.CLEAR, fill: NOTE},
      border: {left: {style: BorderStyle.SINGLE, size: 18, color: 'C00000', space: 8}},
      children: [new TextRun({text: '※ ' + s.note, font: FONTG, size: 18, color: GREY})],
    }));
  }
});

// ── 連絡先 ──────────────────────────────────────────────
kids.push(new Paragraph({children: [new PageBreak()]}));
kids.push(p('お問い合わせ先', {size: 24, bold: true, font: FONTG, color: NAVY, after: 200}));
kids.push(p(C.meta.footer, {size: 21, font: FONTG, after: 60}));
kids.push(p('第10期北塩原村高齢者福祉計画・介護保険事業計画策定業務　担当',
            {size: 21, font: FONTG, after: 240}));
kids.push(p('本書の各事項には由来（貴側の照会番号・申し送りの節）を付しております。'
            + '貴側の照会票に行として立てていただく際は、由来をそのままお使いいただけます。'
            + 'ご不明な点、様式についてのご相談も承ります。', {after: 120}));

const doc = new Document({
  // 既定の書体と大きさを宣言する（宣言しないと Word の既定が段落記号に効く）
  styles: {default: {document: {run: {font: FONT, size: 21}},
                     heading1: {run: {font: FONTG, size: 26, bold: true, color: 'FFFFFF'}}}},
  sections: [{
    properties: {page: {margin: {top: 1134, bottom: 1134, left: 1134, right: 1134}}},
    footers: {default: new Footer({children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({children: [PageNumber.CURRENT], font: FONTG, size: 18, color: GREY})],
    })]})},
    children: kids,
  }],
});

const OUT = path.join(OUTDIR, '14_北塩原村第10期_障がい福祉計画担当への申し送り.docx');
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log('保存: ' + OUT + ' ' + Math.round(buf.length / 1024) + ' KB');
});
