// 交付金の獲得状況と取得の優先順位　Word生成
//   kofukin_yusen.py --json が出力した output/_build/kofukin.json を読む。
//
// 県内の他団体の扱い：59市町村の集計値（平均・中央値・最高・最低・順位・取得団体数）
// だけを載せる。**団体名と個別の数値は載せない。**
const fs = require('fs');
const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..', '..');
const OUTDIR = path.join(ROOT, 'output');
const d = require(path.join(__dirname, 'docxlib.js'));
const {Document, Packer, Paragraph, TextRun, AlignmentType, PageBreak,
       Table, TableRow, TableCell, WidthType, ShadingType, Footer, PageNumber,
       LineRuleType, HeadingLevel} = d;

const C = JSON.parse(fs.readFileSync(path.join(OUTDIR, '_build', 'kofukin.json'), 'utf8'));
const FONT = '游明朝', FONTG = '游ゴシック';
const NAVY = '1F3864', BLUE = '2E75B6', GREY = '595959';
const KEYB = 'FFF2CC', WARN = 'FDE9E9', OK = 'E8F2E8';
const TBLW = 9360;

// 目標の見出し。「推目標Ⅰ」は列の中で折り返すため「推進Ⅰ」とする
const lbl = (r) => r['交付金'] + r['目標'].slice(2);

const p = (text, o = {}) => new Paragraph({
  spacing: {after: o.after ?? 120, line: o.line ?? 300, lineRule: LineRuleType.EXACT},
  alignment: o.align, indent: o.indent,
  children: [new TextRun({text, font: o.font ?? FONT, size: o.size ?? 21,
                          bold: o.bold, color: o.color})],
});
const h1 = (t) => new Paragraph({
  keepNext: true,
  spacing: {before: 300, after: 160, line: 320, lineRule: LineRuleType.EXACT},
  children: [new TextRun({text: t, font: FONTG, size: 26, bold: true, color: NAVY})],
});
const h2 = (t) => new Paragraph({
  keepNext: true,
  spacing: {before: 200, after: 120, line: 300, lineRule: LineRuleType.EXACT},
  children: [new TextRun({text: t, font: FONTG, size: 22, bold: true, color: BLUE})],
});
const note = (t) => new Paragraph({
  spacing: {after: 120, line: 280, lineRule: LineRuleType.EXACT},
  children: [new TextRun({text: t, font: FONTG, size: 19, color: GREY})],
});

// fills: 行ごとの網掛け（null で白）。align: 右寄せにする列の番号
function table(head, rows, widths, o = {}) {
  const cols = widths.map(w => Math.round(TBLW * w / 100));
  cols[cols.length - 1] += TBLW - cols.reduce((a, b) => a + b, 0);
  const right = new Set(o.right || []);
  const cell = (txt, i, isHead, fill) => new TableCell({
    width: {size: cols[i], type: WidthType.DXA},
    shading: {type: ShadingType.CLEAR, fill: isHead ? BLUE : (fill || 'FFFFFF')},
    margins: {top: 60, bottom: 60, left: 80, right: 80},
    children: [new Paragraph({
      spacing: {after: 0, line: 240, lineRule: LineRuleType.EXACT},
      alignment: isHead ? AlignmentType.CENTER
                        : (right.has(i) ? AlignmentType.RIGHT : undefined),
      children: [new TextRun({text: String(txt), font: FONTG, size: o.size ?? 17,
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
// ───────── 表紙 ─────────
K.push(p(C['日付'], {align: AlignmentType.RIGHT, font: FONTG, after: 200}));
K.push(p('北塩原村　第10期高齢者福祉計画・第10期介護保険事業計画',
         {align: AlignmentType.CENTER, font: FONTG, size: 21, color: GREY, after: 80}));
K.push(p('保険者機能強化推進交付金等', {align: AlignmentType.CENTER, font: FONTG,
         size: 30, bold: true, color: NAVY, after: 60}));
K.push(p('獲得状況の整理と取得の優先順位', {align: AlignmentType.CENTER, font: FONTG,
         size: 30, bold: true, color: NAVY, after: 320}));

K.push(h1('1　要点'));
const yoten = [
 ['いまの得点', `${C['合計']}点／${C['満点']}点（得点率 ${(C['合計'] / C['満点'] * 100).toFixed(1)}％）`,
  `県内${C['県内順位']}位／${C['県内団体数']}市町村　全国${C['全国順位']}位／${C['全国団体数']}市町村`],
 ['県・全国との差', `県平均 ${C['県平均合計']}点　全国平均 ${C['全国平均合計']}点`,
  `県平均を ${(C['県平均合計'] - C['合計']).toFixed(1)}点、全国平均を ${(C['全国平均合計'] - C['合計']).toFixed(1)}点下回る`],
 ['計画への記載で取れる', `${C['段'][0]['項目']}項目 ${C['段'][0]['点']}点`,
  `取得すると ${C['合計'] + C['段'][0]['点']}点。県平均・全国平均を上回る`],
 ['本業務の実施で取れる', `${C['段'][1]['項目']}項目 ${C['段'][1]['点']}点`,
  `あわせて ${C['合計'] + C['段'][0]['点'] + C['段'][1]['点']}点`],
 ['事業の実施が前提', `${C['段'][2]['項目']}項目 ${C['段'][2]['点']}点`,
  `すべて行えば ${C['合計'] + C['段'][0]['点'] + C['段'][1]['点'] + C['段'][2]['点']}点。村のご判断を要する`],
 ['本村では動かせない', `${C['段'][3]['項目']}項目 ${C['段'][3]['点']}点`,
  '国の統計から自動算定される成果指標群など'],
 ['判定が済んでいない', `${C['段'][4]['項目']}項目 ${C['段'][4]['点']}点`,
  `うち県内の4分の1以上が取得しているものが ${C['未判定計']['4分の1以上']}項目 ${C['未判定計']['4分の1以上の点']}点`],
];
K.push(table(['', '規模', '効き方'], yoten, [22, 26, 52],
             {fills: [KEYB, null, OK, OK, null, null, null]}));
K.push(p('', {after: 100}));
K.push(p('最大の弱点は支援交付金の目標Ⅱ（認知症総合支援）です。32点で県内57位／59市町村にあり、県平均を22.5点下回ります。'
       + '一方、推進交付金の目標Ⅰ（持続可能な地域のあるべき姿）は76点で県内7位にあり、県平均を12.7点上回ります。', {after: 140}));
K.push(note('※ 本書の点はすべて令和8年度の評価指標によるものです。評価指標は毎年度改定されるため、'
          + '第10期（令和9年度から令和11年度まで）に適用される指標は未公表です。'));
K.push(note('※ 県内の値は福島県59市町村の集計です。本書は団体名と個別の団体の数値を載せていません。'));
K.push(note('※ 出どころ：令和8年度保険者機能強化推進交付金及び介護保険保険者努力支援交付金に係る'
          + '該当状況調査票集計表（市町村分・厚生労働省）および同評価指標（市町村分）。本村は通し番号380です。'));

// ───────── 2 獲得状況 ─────────
K.push(new Paragraph({children: [new PageBreak()]}));
K.push(h1('2　令和8年度の獲得状況'));
K.push(p(`保険者機能強化推進交付金 ${C['推進']}点／400点、介護保険保険者努力支援交付金 ${C['支援']}点／400点、`
       + `あわせて ${C['合計']}点／${C['満点']}点です。`, {after: 140}));
K.push(h2('(1) 目標別の得点と県内・全国の位置'));
K.push(table(['目標', '目標名', '本村', '満点', '県平均', '県中央', '県最高', '県最低', '県内順位'],
  C['目標別'].map(r => [lbl(r), r['目標名'],
    String(r['いま']), '100', r['県平均'].toFixed(1), String(r['県中央値']),
    String(r['県最高']), String(r['県最低']), `${r['県内順位']}／${r['県内団体数']}`]),
  [9, 31, 7, 7, 8, 8, 8, 8, 14],
  {right: [2, 3, 4, 5, 6, 7], size: 16,
   fills: C['目標別'].map(r => r['いま'] < r['県平均'] ? WARN : OK)}));
K.push(p('', {after: 100}));
K.push(note('※ 薄い赤は県平均を下回る目標、薄い緑は上回る目標です。'));
K.push(note('※ 目標Ⅳ（成果指標群）は推進交付金・支援交付金に同一の指標が置かれており、得点も同じです。'
          + '片方が動けばもう片方も同じだけ動きます。'));

K.push(h2('(2) 3か年の推移'));
K.push(p(`3か年（令和6年度から令和8年度まで）そろって比べられる項目は ${C['推移']['そろう項目']}項目です。`, {after: 120}));
K.push(table(['区分', '項目数', '令和8年度の得点計'],
  C['推移']['区分'].map(r => [r['区分'], `${r['項目']}項目`, `${r['点']}点`]),
  [40, 25, 35], {right: [1, 2]}));
K.push(p('', {after: 100}));
K.push(p(`合計得点は令和6年度295点、令和7年度339点、令和8年度${C['合計']}点と伸びています。`
       + `ただし ${C['推移']['区分'].find(x => x['区分'] === '悪化')['項目']}項目が下がり、`
       + `${C['推移']['区分'].find(x => x['区分'] === '3年連続0点')['項目']}項目は3年続けて0点です。`, {after: 140}));

K.push(h2('(3) 下がった項目'));
K.push(table(['目標', '指標', '配点', '推移'],
  C['悪化'].slice(0, 14).map(r => [lbl(r), r['指標'],
    `${r['配点']}点`, r['推移']]),
  [10, 60, 10, 20], {right: [2], size: 16}));
K.push(p('', {after: 100}));
K.push(note('推進目標Ⅳの11項目は、要介護度の変化（成果指標群）が国の統計から自動算定されるものです。'
          + '推進目標Ⅱのオ（福祉用具購入費・住宅改修費の妥当性の検討）と支援目標Ⅲのア（在宅医療・介護連携の'
          + '課題・対応策の検討）は、いずれも取組を続けることで取り戻せる項目です。'));

// ───────── 3 未取得の内訳 ─────────
K.push(h1('3　未取得の点の内訳'));
K.push(p(`未取得は ${C['段'].reduce((a, b) => a + b['項目'], 0)}項目 ${C['段'].reduce((a, b) => a + b['点'], 0)}点です。`
       + '取得の見込みで5つの段に分けました。', {after: 140}));
K.push(table(['段', '内容', '項目数', '点', '何が要るか'],
  [['Ａ', C['段'][0]['内容'], `${C['段'][0]['項目']}項目`, `${C['段'][0]['点']}点`,
    '計画への記載・要綱の整備・手順の明文化で足ります'],
   ['Ｂ', C['段'][1]['内容'], `${C['段'][1]['項目']}項目`, `${C['段'][1]['点']}点`,
    'ニーズ調査・在宅介護実態調査・見える化分析・認知症施策推進計画の策定を行うこと自体が要件を満たします'],
   ['Ｃ', C['段'][2]['内容'], `${C['段'][2]['項目']}項目`, `${C['段'][2]['点']}点`,
    '事業の実施・実績の水準・他機関との連携が前提です。村のご判断を要します'],
   ['Ｄ', C['段'][3]['内容'], `${C['段'][3]['項目']}項目`, `${C['段'][3]['点']}点`,
    '国の統計から自動算定される成果指標群と、他の指標の結果として決まるものです'],
   ['Ｅ', C['段'][4]['内容'], `${C['段'][4]['項目']}項目`, `${C['段'][4]['点']}点`,
    '個別の判定をしていません。全国の45％未満の市町村しか取得していない項目が主です']],
  [6, 18, 10, 8, 58], {right: [2, 3], fills: [KEYB, OK, null, null, null]}));

// ───────── 4 優先順位 ─────────
K.push(new Paragraph({children: [new PageBreak()]}));
K.push(h1('4　取得の優先順位'));
K.push(p('段Ａ・Ｂ・Ｃの29項目を、期待得点（配点に取得の確からしさを掛けたもの。段Ａ・Ｂは1.0、段Ｃは0.5）の'
       + '大きい順に並べました。県内で取得している団体の数を併記しています。'
       + '県内で多くの団体が取得している項目は、取りやすく、取れていないことの説明もつきにくいものです。', {after: 140}));
K.push(table(['順', '段', '目標', '指標', '点', '期待', '県内で取得'],
  C['優先順位'].map(r => [String(r['順']), r['段'], lbl(r),
    r['指標'], `${r['点']}点`, `${r['期待']}`,
    `${r['県内満点']}／${r['県内団体数']}`]),
  [5, 5, 10, 50, 8, 8, 14],
  {right: [4, 5, 6], size: 16,
   fills: C['優先順位'].map(r => r['段'] === 'Ａ' ? KEYB : (r['段'] === 'Ｂ' ? OK : null))}));
K.push(p('', {after: 100}));
K.push(note('※ 黄色は段Ａ（記載で取れる）、緑は段Ｂ（本業務で取れる）です。'));

K.push(new Paragraph({children: [new PageBreak()]}));
K.push(h1('5　第1順位群　記載と本業務で取れる13項目65点'));
K.push(p('段Ａ・Ｂの13項目は、新たな施設や大きな予算を要しません。'
       + '何をすれば取れるかと、本計画のどこで扱っているかを示します。', {after: 140}));
C['優先順位'].filter(r => r['段'] === 'Ａ' || r['段'] === 'Ｂ').forEach(r => {
  K.push(new Paragraph({
    keepNext: true,
    spacing: {before: 180, after: 60, line: 280, lineRule: LineRuleType.EXACT},
    children: [new TextRun({text: `${r['順']}　${lbl(r)}　${r['指標']}`,
                            font: FONTG, size: 20, bold: true, color: NAVY}),
               new TextRun({text: `　（${r['点']}点・県内${r['県内満点']}／${r['県内団体数']}団体が取得）`,
                            font: FONTG, size: 18, color: GREY})],
  }));
  K.push(p(`要ること：${r['必要なこと']}`, {after: 40, size: 20}));
  K.push(p(`本計画での扱い：${r['素案']}`, {after: 60, size: 20, color: GREY, font: FONTG}));
});

// ───────── 6 未判定 ─────────
K.push(h1('6　判定が済んでいない項目のうち、判定を急ぐもの'));
K.push(p(`段Ｅは ${C['未判定計']['項目']}項目 ${C['未判定計']['点']}点です。`
       + `うち県内の半数以上が取得しているものが ${C['未判定計']['半数以上']}項目 ${C['未判定計']['半数以上の点']}点、`
       + `4分の1以上が ${C['未判定計']['4分の1以上']}項目 ${C['未判定計']['4分の1以上の点']}点あります。`
       + '県内で取得している団体が多いのに本村が0点のものから判定します。', {after: 140}));
K.push(table(['目標', '指標', '点', '県内で取得'],
  C['未判定'].map(r => [lbl(r), r['指標'], `${r['点']}点`,
    `${r['県内満点']}／${C['県内団体数']}`]),
  [10, 60, 10, 20], {right: [2, 3], size: 16}));

// ───────── 7 見込める得点 ─────────
K.push(new Paragraph({children: [new PageBreak()]}));
K.push(h1('7　見込める得点'));
K.push(p('目標ごとに積み上げました。満点100点を超えると切り捨てられるため、目標ごとに確かめています。', {after: 140}));
K.push(table(['目標', '目標名', 'いま', '＋Ａ', '＋Ｂ', '＋Ｃ', '満点', '県平均'],
  C['目標別'].map(r => [lbl(r), r['目標名'],
    String(r['いま']), String(r['A']), String(r['B']), String(r['C']), '100',
    r['県平均'].toFixed(1)]),
  [9, 33, 8, 8, 8, 8, 8, 10], {right: [2, 3, 4, 5, 6, 7], size: 16}));
const tn = C['目標別'].reduce((a, r) => a + r['いま'], 0);
const ta = C['目標別'].reduce((a, r) => a + r['A'], 0);
const tb = C['目標別'].reduce((a, r) => a + r['B'], 0);
const tc = C['目標別'].reduce((a, r) => a + r['C'], 0);
K.push(p('', {after: 100}));
const ran = (v) => C['順位'].find(x => Math.abs(x['点'] - v) < 0.5) || {全国: '', 県内: ''};
const sa = (v) => (v >= 0 ? '＋' : '') + v.toFixed(1) + '点';
K.push(table(['', '合計', '全国平均との差', '県平均との差', '全国順位', '県内順位'],
  [['いま', `${tn}点`, sa(tn - C['全国平均合計']), sa(tn - C['県平均合計']),
    `${ran(tn)['全国']}／${C['全国団体数']}`, `${ran(tn)['県内']}／${C['県内団体数']}`],
   ['＋Ａ 記載で取れる', `${ta}点`, sa(ta - C['全国平均合計']), sa(ta - C['県平均合計']),
    `${ran(ta)['全国']}／${C['全国団体数']}`, `${ran(ta)['県内']}／${C['県内団体数']}`],
   ['＋Ｂ 本業務で取れる', `${tb}点`, sa(tb - C['全国平均合計']), sa(tb - C['県平均合計']),
    `${ran(tb)['全国']}／${C['全国団体数']}`, `${ran(tb)['県内']}／${C['県内団体数']}`],
   ['＋Ｃ 事業の実施が前提', `${tc}点`, sa(tc - C['全国平均合計']), sa(tc - C['県平均合計']),
    `${ran(tc)['全国']}／${C['全国団体数']}`, `${ran(tc)['県内']}／${C['県内団体数']}`]],
  [26, 12, 17, 17, 14, 14], {right: [1, 2, 3, 4, 5], size: 16,
                             fills: [null, KEYB, OK, null]}));
K.push(p('', {after: 100}));
K.push(p(`満点100点に頭打ちになる目標はありません（＋Ｃまで行っても最大 ${Math.max(...C['目標別'].map(r => r['C']))}点）。`, {after: 140}));
K.push(note('※ 交付額は得点に第1号被保険者数等を乗じて国が配分します。'
          + '配分の総額や他団体の得点によっても変わるため、本書は得点までを示しています。'));

K.push(h1('8　留意すること'));
[['評価指標は毎年度改定される',
  '本書の点は令和8年度の評価指標によるものです。第10期に適用される指標は未公表であり、'
  + '項目の入れ替えや配点の変更があり得ます（令和4年度にあった「文書負担軽減に係る取組」は'
  + '令和6年度以降の指標から外れています）。'],
 ['成果指標群は本村では動かせない',
  '目標Ⅳ（要介護度の変化・健康寿命の延伸）は国が自動算定します。本村は対象者が少なく、'
  + '1人か2人の状態の変化で5点単位で動きます。下がった26項目のうち11項目がこれにあたります。'],
 ['県内の値は集計である',
  '県平均・県中央値・県最高・県最低・県内順位・取得団体数は福島県59市町村の集計です。'
  + '個別の団体の名称と数値は載せていません。'],
].forEach(([t, b]) => {
  K.push(new Paragraph({
    keepNext: true,
    spacing: {before: 160, after: 50, line: 280, lineRule: LineRuleType.EXACT},
    children: [new TextRun({text: t, font: FONTG, size: 20, bold: true, color: BLUE})],
  }));
  K.push(p(b, {after: 60, size: 20}));
});

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

const OUT = path.join(OUTDIR, '15_北塩原村第10期_交付金の獲得状況と優先順位.docx');
Packer.toBuffer(doc).then(b => {
  fs.writeFileSync(OUT, b);
  console.log('保存: ' + OUT + ' ' + Math.round(b.length / 1024) + ' KB');
});
