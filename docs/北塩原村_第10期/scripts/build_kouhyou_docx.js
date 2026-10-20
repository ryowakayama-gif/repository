// 公表用素案の構成案　Word生成（kouhyou_map.py --json が出力した kouhyou.json を読む）
const fs = require('fs');
const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..', '..');
const OUTDIR = path.join(ROOT, 'output');
const d = require(path.join(__dirname, 'docxlib.js'));
const {Document, Packer, Paragraph, TextRun, AlignmentType, PageBreak,
       Table, TableRow, TableCell, WidthType, ShadingType, Footer, PageNumber,
       LineRuleType} = d;

const C = JSON.parse(fs.readFileSync(path.join(OUTDIR, '_build', 'kouhyou.json'), 'utf8'));
const FONT = '游明朝', FONTG = '游ゴシック';
const NAVY = '1F3864', BLUE = '2E75B6', GREY = '595959';
const KEYB = 'FFF2CC', OK = 'E8F2E8', GRAY = 'F2F2F2';
const TBLW = 9360;

const p = (t, o = {}) => new Paragraph({
  keepNext: o.keepNext,
  spacing: {after: o.after ?? 120, line: o.line ?? 300, lineRule: LineRuleType.EXACT},
  alignment: o.align,
  children: [new TextRun({text: t, font: o.font ?? FONT, size: o.size ?? 21,
                          bold: o.bold, color: o.color})],
});
const h1 = (t) => new Paragraph({
  keepNext: true,
  spacing: {before: 280, after: 150, line: 320, lineRule: LineRuleType.EXACT},
  children: [new TextRun({text: t, font: FONTG, size: 26, bold: true, color: NAVY})],
});
const h2 = (t) => new Paragraph({
  keepNext: true,
  spacing: {before: 190, after: 110, line: 300, lineRule: LineRuleType.EXACT},
  children: [new TextRun({text: t, font: FONTG, size: 22, bold: true, color: BLUE})],
});
const note = (t) => new Paragraph({
  spacing: {after: 110, line: 275, lineRule: LineRuleType.EXACT},
  children: [new TextRun({text: t, font: FONTG, size: 19, color: GREY})],
});

function table(head, rows, widths, o = {}) {
  const cols = widths.map(w => Math.round(TBLW * w / 100));
  cols[cols.length - 1] += TBLW - cols.reduce((a, b) => a + b, 0);
  const right = new Set(o.right || []);
  const cell = (txt, i, isHead, fill) => new TableCell({
    width: {size: cols[i], type: WidthType.DXA},
    shading: {type: ShadingType.CLEAR, fill: isHead ? BLUE : (fill || 'FFFFFF')},
    margins: {top: 55, bottom: 55, left: 80, right: 80},
    children: [new Paragraph({
      spacing: {after: 0, line: 235, lineRule: LineRuleType.EXACT},
      alignment: isHead ? AlignmentType.CENTER
                        : (right.has(i) ? AlignmentType.RIGHT : undefined),
      children: [new TextRun({text: String(txt), font: FONTG, size: o.size ?? 16,
                              bold: isHead, color: isHead ? 'FFFFFF' : '000000'})],
    })],
  });
  return new Table({
    columnWidths: cols, width: {size: TBLW, type: WidthType.DXA},
    rows: [new TableRow({tableHeader: true,
                         children: head.map((x, i) => cell(x, i, true, null))}),
           ...rows.map((r, j) => new TableRow({
             children: r.map((v, i) => cell(v, i, false, (o.fills || [])[j])),
           }))],
  });
}

const K = [];
K.push(p(C['日付'], {align: AlignmentType.RIGHT, font: FONTG, after: 180}));
K.push(p('北塩原村　第10期高齢者福祉計画・第10期介護保険事業計画',
         {align: AlignmentType.CENTER, font: FONTG, size: 21, color: GREY, after: 70}));
K.push(p('公表用素案の構成案', {align: AlignmentType.CENTER, font: FONTG,
         size: 30, bold: true, color: NAVY, after: 60}));
K.push(p('― 第9期計画の目次体系への当て方 ―', {align: AlignmentType.CENTER,
         font: FONTG, size: 22, color: NAVY, after: 280}));

K.push(h1('1　2つの素案'));
K.push(p('村のご指示により、素案を協議用と公表用に分けます。', {after: 120}));
K.push(table(['', '中身', '分量'],
  [['協議用素案', 'これまでの論点をすべて含みます。策定委員会での協議と村との確認に用います',
    `${C['分量']['協議用素案']}頁（制約なし）`],
   ['公表用素案', '第9期計画の目次体系を踏襲します。新たな論点は指針・通達・法令に基づくものだけを加えます',
    `推定${C['分量']['公表用の推定']}頁（目安${C['分量']['目安']}頁・最大${C['分量']['許容']}頁）`]],
  [16, 62, 22], {fills: [GRAY, KEYB]}));
K.push(p('', {after: 100}));
K.push(p(`協議用素案は${C['節の数']['協議用']}節あります。第9期計画の目次は39項目です。`
       + `公表用はこの39項目に当てたうえで、指針・通達・法令による${C['追加'].length}項目を加え、`
       + `${C['節の数']['協議用のみ']}節を協議用にのみ置きます。`, {after: 140}));
K.push(note('※ 推定は送りによる空白を小さく見るため、実測は0〜3頁多く出ます。'
          + '組んで変換して数えたうえで改めてお示しします。'));

K.push(h1('2　第9期の目次に、協議用素案のどの節を当てるか'));
K.push(p('網掛けの行は、指針・通達・法令により新たに加える項目です。', {after: 130}));
C['章'].forEach(sho => {
  K.push(h2(`${sho['章']}　${sho['章名']}`));
  if (sho['章'] === '第4章') {
    K.push(p('基本目標は第9期と同じ4つです。'
      + Object.entries(C['基本目標']).map(([k, v]) => `基本目標${k} ${v}`).join('／')
      + '。番号の左側が基本目標、右側が施策です。', {after: 110, size: 20}));
  }
  const rows = sho['項目'].map(r => [
    r['番号'],
    r['項目'],
    (r['協議用素案'] || []).join('・') || '（新設）',
    r['追加'] ? '追加' : (r['第9期の頁'] ? `p${r['第9期の頁']}` : ''),
  ]);
  K.push(table(['番号', '第9期の項目／追加する項目', '当てる節', '区分'], rows,
    [9, 42, 39, 10],
    {fills: sho['項目'].map(r => r['追加'] ? KEYB : null)}));
  const memos = sho['項目'].filter(r => r['申し送り'] && !r['追加']);
  if (memos.length) {
    K.push(p('', {after: 60}));
    memos.forEach(r => K.push(note(`${r['番号']}　${r['申し送り']}`)));
  }
  K.push(p('', {after: 80}));
});

K.push(new Paragraph({children: [new PageBreak()]}));
K.push(h1('3　新たに加える項目と、その根拠'));
K.push(p('いずれも第9期計画に対応する記載がなく、法令または基本指針が求めるものです。', {after: 130}));
C['追加'].forEach(a => {
  K.push(new Paragraph({
    keepNext: true,
    spacing: {before: 160, after: 50, line: 280, lineRule: LineRuleType.EXACT},
    children: [new TextRun({text: `${a['章']} ${a['番号']}　${a['項目']}`,
                            font: FONTG, size: 20, bold: true, color: NAVY}),
               new TextRun({text: `　［${a['別']}］　← ${a['節'].join('・')}`,
                            font: FONTG, size: 18, color: GREY})],
  }));
  K.push(p(a['根拠'], {after: 60, size: 20}));
});

K.push(h1('4　協議用にのみ置く節'));
K.push(p(`${C['協議用のみ'].length}節です。いずれも第9期計画に対応する記載がないか、`
       + '他の節に織り込めるものです。', {after: 130}));
K.push(table(['節', '公表用に載せない理由'],
  C['協議用のみ'].map(x => [x['節'], x['理由']]), [10, 90]));

K.push(new Paragraph({children: [new PageBreak()]}));
K.push(h1('5　ご確認いただきたいこと'));
[['第9期の39項目への当て方でよいか',
  '2の表のとおりです。1つの項目に複数の節を当てている箇所（たとえば第4章4-1に'
  + '施策4-1・4-2・4-3・4-5を当てる）が、第9期の記述の範囲に収まるかをご覧ください。'],
 ['新たに加える6項目はこれで足りるか',
  '3のとおり、法令または基本指針が求めるものに限っています。'
  + 'ほかに指針・通達で加えるべきものがあればご指示ください。'],
 ['協議用にのみ置く12節の扱いでよいか',
  '4のとおりです。公表用に残すものがあればご指示ください。'],
 ['第4章の施策を24から15に束ねてよいか',
  '協議用素案は24の施策を置いています。第9期は14施策です。'
  + 'これに指針による「介護人材の確保と生産性向上」を加えた15に束ねる案です。'
  + '束ね方が変わると本文と相互参照の付け替えをやり直すことになるため、'
  + 'この点のご判断を先にいただきたくお願いします。'],
].forEach(([t, b], i) => {
  K.push(new Paragraph({
    keepNext: true,
    spacing: {before: 170, after: 50, line: 280, lineRule: LineRuleType.EXACT},
    children: [new TextRun({text: `(${i + 1})　${t}`, font: FONTG, size: 21,
                            bold: true, color: BLUE})],
  }));
  K.push(p(b, {after: 60}));
});
K.push(p('', {after: 100}));
K.push(note('※ ご確認をいただくまで公表用の組み立てには着手しません。'
          + '協議用素案はいまのままお示しできますので、第2回策定委員会の日程には影響しません。'));

const doc = new Document({
  styles: {default: {document: {run: {font: FONT, size: 21}}}},
  sections: [{
    properties: {page: {margin: {top: 1134, right: 1134, bottom: 1134, left: 1134}}},
    footers: {default: new Footer({children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({children: [PageNumber.CURRENT], font: FONTG, size: 18,
                              color: GREY})],
    })]})},
    children: K,
  }],
});
const OUT = path.join(OUTDIR, '16_北塩原村第10期_公表用素案の構成案.docx');
Packer.toBuffer(doc).then(b => {
  fs.writeFileSync(OUT, b);
  console.log('保存: ' + OUT + ' ' + Math.round(b.length / 1024) + ' KB');
});
