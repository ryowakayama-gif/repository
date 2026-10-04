// 策定委員会 資料 Word生成（*_content.py が出力した JSON を読む）
//   node scripts/build_shiryo_docx.js            第2回（既定）
//   node scripts/build_shiryo_docx.js 3          第3回
// 回ごとに別の組立てを書くと体裁が食い違うため、同じ組立てを使い回す。
const fs = require('fs');
const d = require('/tmp/node_modules/docx');
const {Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, PageBreak,
       Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, LevelFormat,
       TableOfContents, Footer, PageNumber, ImageRun, TableLayoutType} = d;
const path = require('path');
const KAI = (process.argv[2] === '3') ? 3 : 2;
const SRC_PY = KAI === 3 ? 'shiryo3_content.py' : 'shiryo_content.py';
const JSON_PATH = KAI === 3 ? '/tmp/shiryo3.json' : '/tmp/shiryo.json';
const OUT_PATH = KAI === 3
  ? '/home/user/repository/output/12_北塩原村第10期_第3回策定委員会資料.docx'
  : '/home/user/repository/output/08_北塩原村第10期_第2回策定委員会資料.docx';
// 入力のJSONが元データより古いときは止める（古い成果品を作らないため）
(function () {
  // __dirname 基準にする。cwd に依らず走らせるため
  const src = path.join(__dirname, SRC_PY);
  const js  = JSON_PATH;
  if (!fs.existsSync(js)) {
    console.error('入力の ' + js + ' がありません。先に ' + src + ' を実行してください。');
    process.exit(1);
  }
  if (fs.statSync(js).mtimeMs < fs.statSync(src).mtimeMs) {
    console.error('入力の ' + js + ' が ' + src + ' より古いです。先に ' + src + ' を実行してください。');
    process.exit(1);
  }
})();


const FIGDIR = '/home/user/repository/output/figures';
function pngSize(buf) {              // PNG IHDR: 幅=16..19 / 高さ=20..23（ビッグエンディアン）
  return {w: buf.readUInt32BE(16), h: buf.readUInt32BE(20)};
}
// 本文の高さ（10.12in）のこの割合を超える図は、頁の頭から置く。
// build_soan_docx.js の ZENMEN_H_IN と同じ割合を用いる。
const ZENMEN_H_IN = 10.12 * 0.60;

function figTall(b) {
  const buf = fs.readFileSync(path.join(FIGDIR, b.file));
  const dim = pngSize(buf);
  const hIn = (b.width || 6.3) * dim.h / dim.w;
  return Math.min(hIn, 9.3) >= ZENMEN_H_IN;
}

function figure(b) {
  const buf = fs.readFileSync(path.join(FIGDIR, b.file));
  // 本文高さは 14570 twip ＝ 10.12in。表題・出典・前後の空きを除くと図に使えるのは
  // 9.3in までである。縦長の図は幅を縮めて1頁に収める（Word は画像を縮めない）。
  const MAX_H_IN = 9.3;
  const dim = pngSize(buf);
  let wIn = b.width || 6.3, hIn = wIn * dim.h / dim.w;
  if (hIn > MAX_H_IN) { hIn = MAX_H_IN; wIn = hIn * dim.w / dim.h; }
  const out = [new Paragraph({
    alignment: AlignmentType.CENTER, spacing: {before: 160, after: 60},
    keepNext: true, keepLines: true,   // 図・表題・出典を頁の境目で引き離さない
    children: [new ImageRun({type: 'png', data: buf,
               transformation: {width: Math.round(wIn * 96), height: Math.round(hIn * 96)}})],
  }), new Paragraph({
    alignment: AlignmentType.CENTER, spacing: {after: b.source ? 40 : 200},
    keepNext: true, keepLines: true,
    children: [new TextRun({text: b.caption, font: FONTG, size: 18, bold: true, color: NAVY})],
  })];
  if (b.source) out.push(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: {after: 200},
    children: [new TextRun({text: '出典：' + b.source, font: FONTG, size: 16, color: GREY})],
  }));
  return out;
}

const C = JSON.parse(fs.readFileSync(JSON_PATH, 'utf8'));
// 書式は計画素案と同じ体裁（本文 游ゴシック 10.5pt）に合わせる。図表の色味は現行のまま
const FONT = '游ゴシック', FONTG = '游ゴシック';
const NAVY = '1F3864', BLUE = '2E75B6', BAND = 'DDEBF7', NOTE = 'FFF3F3', KEYB = 'FFF2CC', GREY = '595959';
const TBLW = 9360;

const p = (text, o = {}) => new Paragraph({
  spacing: {after: o.after ?? 120, line: o.line ?? 300},
  alignment: o.align, indent: o.indent, border: o.border,
  keepNext: o.keep, keepLines: o.keep,   // 見出しを次の要素から切り離さない
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
    // 列幅を宣言どおりに固定する。既定の自動調整だと Word が中身に合わせて
    // 列幅を変えてしまい、widths で決めた割付も頁数の推定も当てにならなくなる。
    layout: TableLayoutType.FIXED,
    rows: [new TableRow({tableHeader: true, cantSplit: true, children: head.map((h, i) => cell(h, i, true))}),
           ...rows.map(r => new TableRow({cantSplit: true, children: r.map((v, i) => cell(v, i, false))}))],
  });
}

const kids = [];
// 表紙
kids.push(p('', {after: 2600}));
kids.push(p(C.title,  {align: AlignmentType.CENTER, size: 32, bold: true, font: FONTG, color: NAVY, after: 160}));
kids.push(p(C.title2, {align: AlignmentType.CENTER, size: 32, bold: true, font: FONTG, color: NAVY, after: 700}));
kids.push(p(C.subtitle, {align: AlignmentType.CENTER, size: 24, font: FONTG, after: 1600}));
kids.push(p('資料2　第9期計画の進捗と評価', {align: AlignmentType.CENTER, size: 24, font: FONTG, after: 100}));
kids.push(p('資料4　第10期計画の基本理念・基本目標・施策体系（案）', {align: AlignmentType.CENTER, size: 24, font: FONTG, after: 100}));
kids.push(p('資料5　サービス見込み量の考え方（案）', {align: AlignmentType.CENTER, size: 24, font: FONTG, after: 100}));
kids.push(p('資料6　介護保険料の概算シミュレーション', {align: AlignmentType.CENTER, size: 24, font: FONTG, after: 100}));
kids.push(p('資料7　成果目標・活動指標（案）', {align: AlignmentType.CENTER, size: 24, font: FONTG, after: 1600}));
kids.push(p(C.date,   {align: AlignmentType.CENTER, size: 24, font: FONTG, after: 120}));
kids.push(p(C.issuer, {align: AlignmentType.CENTER, size: 28, bold: true, font: FONTG, after: 0}));
kids.push(new Paragraph({children: [new PageBreak()]}));

// 本書の見方
kids.push(p('本資料の見方', {size: 28, bold: true, font: FONTG, color: NAVY, after: 240}));
kids.push(p('本資料は、令和8年11月に開催する第2回策定委員会の資料の一部です。資料3（アンケート調査結果）は、調査の集計後に作成します。', {after: 180}));
kids.push(table(['表記', '意味'], [
  ['【要確認】', '北塩原村への確認により確定する事項です。'],
  ['【要設定】', '目標値等を今後設定する箇所です。'],
  ['【試算中】／【推計中】', '国の推計ワークシートの提供後に算出する数値です。'],
  ['網掛けの囲み', '委員の皆様にご留意いただきたい点です。'],
  ['注記の囲み', '未確定の事項と、その確定の見通しです。'],
], [26, 74]));
kids.push(p('', {after: 200}));
kids.push(p('資料5・6は骨格です。数値は、国の推計ワークシートの提供（第10期基本指針の告示後）と、村の実績データの受領をもって確定します。本日は「考え方」についてご意見をいただきたく存じます。', {after: 180}));
kids.push(p('資料4は、本委員会でご議論いただきたい事項の中心です。各施策の根拠となる住民ニーズの数値は、資料3で補足します。', {after: 240}));
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
    keepNext: true, keepLines: true,   // 見出しだけが頁の最後に取り残されるのを防ぐ
    shading: {type: ShadingType.CLEAR, fill: NAVY},
    children: [new TextRun({text: `${ch.no}　${ch.title}`, font: FONTG, size: 30, bold: true, color: 'FFFFFF'})],
  }));
  ch.sections.forEach(sec => {
    if (!sec.no.endsWith('-0')) {
      kids.push(new Paragraph({
        heading: HeadingLevel.HEADING_2, spacing: {before: 320, after: 180},
        keepNext: true, keepLines: true,
        border: {bottom: {style: BorderStyle.SINGLE, size: 12, color: BLUE, space: 4}},
        children: [new TextRun({text: `${sec.no}　${sec.title}`, font: FONTG, size: 25, bold: true, color: NAVY})],
      }));
    }
    sec.blocks.forEach(b => {
      if (b.t === 'p') kids.push(p(b.v));
      else if (b.t === 'h3') kids.push(p(b.v, {size: 22, bold: true, font: FONTG, color: BLUE, after: 100, keep: true}));
      else if (b.t === 'bullets') b.v.forEach(x => kids.push(new Paragraph({
        numbering: {reference: 'bul', level: 0}, spacing: {after: 60, line: 300},
        children: [new TextRun({text: x, font: FONT, size: 20})],
      })));
      else if (b.t === 'key') {
        const lines = String(b.v).split('\n');
        const runs = [];
        lines.forEach((ln, i) => runs.push(new TextRun({
          text: ln, font: FONTG, size: 20, bold: true, break: i > 0 ? 1 : 0,
        })));
        kids.push(new Paragraph({
          spacing: {before: 140, after: 200, line: 300}, indent: {left: 200, right: 200},
          shading: {type: ShadingType.CLEAR, fill: KEYB},
          border: {left: {style: BorderStyle.SINGLE, size: 18, color: 'BF8F00', space: 8}},
          children: runs,
        }));
      }
      else if (b.t === 'note') kids.push(new Paragraph({
        spacing: {before: 120, after: 180, line: 280}, indent: {left: 200, right: 200},
        shading: {type: ShadingType.CLEAR, fill: NOTE},
        border: {left: {style: BorderStyle.SINGLE, size: 18, color: 'C00000', space: 8}},
        children: [new TextRun({text: '※ ' + b.v, font: FONTG, size: 17, color: GREY})],
      }));
      else if (b.t === 'table') { kids.push(table(b.head, b.rows, b.widths)); kids.push(p('', {after: 160})); }
      else if (b.t === 'fig') {
        if (figTall(b)) kids.push(new Paragraph({children: [new PageBreak()]}));
        figure(b).forEach(x => kids.push(x));
      }
    });
  });
});

const doc = new Document({
  styles: {default: {
    heading1: {run: {font: FONTG, size: 30, bold: true, color: 'FFFFFF'}},
    heading2: {run: {font: FONTG, size: 25, bold: true, color: NAVY}},
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

const OUT = OUT_PATH;
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log('保存: ' + OUT + ' ' + Math.round(buf.length / 1024) + ' KB');
});
