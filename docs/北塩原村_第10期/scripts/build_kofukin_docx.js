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
// 書体は計画書・委員会資料と同じ BIZ UDPゴシック（令和8年10月24日に游明朝・游ゴシックから改めた）
const FONT = 'BIZ UDPゴシック', FONTG = 'BIZ UDPゴシック';
const NAVY = '1F3864', BLUE = '2E75B6', GREY = '595959';
const KEYB = 'FFF2CC', WARN = 'FDE9E9', OK = 'E8F2E8';
const TBLW = 9360;

// 目標の見出し。「推目標Ⅰ」は列の中で折り返すため「推進Ⅰ」とする
const lbl = (r) => r['交付金'] + r['目標'].slice(2);

const p = (text, o = {}) => new Paragraph({
  // 11ptを超える字は行送りを最小値指定にする（固定だと Word で字の上が切れる。doc114）
  spacing: (o.size ?? 21) > 22
    ? {after: o.after ?? 120, line: Math.max(o.line ?? 300, Math.ceil((o.size ?? 21) * 13.5)), lineRule: LineRuleType.AT_LEAST}
    : {after: o.after ?? 120, line: o.line ?? 300, lineRule: LineRuleType.EXACT},
  alignment: o.align, indent: o.indent,
  children: [new TextRun({text, font: o.font ?? FONT, size: o.size ?? 21,
                          bold: o.bold, color: o.color})],
});
const h1 = (t) => new Paragraph({
  keepNext: true,
  spacing: {before: 300, after: 160, line: 360, lineRule: LineRuleType.AT_LEAST},
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
 ['計画への記載で取得をめざせる', `${C['段'][0]['項目']}項目 ${C['段'][0]['点']}点`,
  `取得すると ${C['合計'] + C['段'][0]['点']}点。県平均・全国平均を上回る`],
 ['本業務の実施で取得をめざせる', `${C['段'][1]['項目']}項目 ${C['段'][1]['点']}点`,
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
// 素案6-3・参考資料19・第2回資料7-3c の区分との対照（外部レビュー R14）
K.push(p('素案6-3・参考資料（交付金の評価と第10期の目標）・第2回策定委員会資料7-3cは、令和8年度に0点で全国の45％以上の市町村が取得している32項目を4つに分けています。'
       + '本書の段は未取得の全項目を分けたものであり、母集団が違います。同じ項目は同じ扱いであり、差は32項目の外の項目だけです。', {after: 100}));
K.push(table(['本書の段', '本書の項目・点', 'うち32項目の内（素案6-3等の区分）', '32項目の外の項目'],
  C['対照'].map(x => [
    {'Ａ': '段Ａ 記載で取得をめざせる', 'Ｂ': '段Ｂ 本業務で取得をめざせる', 'Ｃ': '段Ｃ 事業の実施が前提'}[x['段']],
    `${x['項目']}項目 ${x['点']}点`,
    `${x['32項目内_項目']}項目 ${x['32項目内_点']}点（${{'Ａ': '計画への記載と実施で取得をめざせる', 'Ｂ': '本業務の実施で取得をめざせる', 'Ｃ': '条件付きで取得をめざせる'}[x['段']]}）`,
    x['32項目外'].length ? x['32項目外'].map(y => `${y['交付金']} ${y['指標']}（${y['点']}点）`).join('・') : '―']),
  [22, 16, 34, 28], {size: 16}));
K.push(note('※ 素案6-3の取得目標67点は、段Ａの62点に、段Ｂのうち認知症施策推進計画の策定（5点。全国該当率42.8％のため32項目の外）を加えたものです。'));
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

K.push(h2('(2) 条件の近い団体との比較'));
K.push(p('得点は保険者の規模と強く相関します。本村は第1号被保険者'
       + `${C['仲間']['第1号被保険者数'].toLocaleString('en-US')}人で`
       + `「${C['仲間']['区分名']}」の区分にあたり、過疎地域にも該当します。`
       + '全国平均との比較だけでは本村の位置を低く見せることになるため、'
       + '条件の近い団体の平均も示します。', {after: 140}));
K.push(table(['比較の相手', '団体数', '平均', `本村${C['合計']}点との差`],
  C['仲間']['比較'].map(x => [x['相手'], `${x['団体数']}団体`, `${x['平均']}点`,
    (x['差'] >= 0 ? '＋' : '▲') + Math.abs(x['差']).toFixed(1) + '点']),
  [40, 16, 18, 26], {right: [1, 2, 3]}));
K.push(table(['第1号被保険者数の区分', '団体数', '平均'],
  C['仲間']['規模別'].map(r => [r['区分名'], `${r['団体数']}団体`, `${r['平均']}点`]),
  [44, 24, 32], {right: [1, 2]}));
K.push(p('全国平均と比べると下回りますが、同規模の団体の平均とは'
       + `${Math.abs(C['仲間']['比較'].find(x => x['相手'].startsWith('同規模（'))['差']).toFixed(1)}点、`
       + '同規模かつ過疎地域の団体の平均とは'
       + `${Math.abs(C['仲間']['比較'].find(x => x['相手'] === '同規模かつ過疎地域')['差']).toFixed(1)}点`
       + '上回ります。本村の取組の水準は、条件の近い団体の中では高い位置にあります。'
       + 'ただし未取得の項目があることは変わりません。', {after: 140}));
K.push(p('示しているのは団体数と平均という集計値のみで、個別の団体名と得点は掲げていません。'
       + '区分は令和7年3月31日現在の第1号被保険者数、過疎地域該当は一部過疎を含む'
       + '令和5年4月1日時点によります。', {after: 180}));
K.push(h2('(3) 3か年の推移'));
K.push(p(`3か年（令和6年度から令和8年度まで）そろって比べられる項目は ${C['推移']['そろう項目']}項目です。`, {after: 120}));
K.push(table(['区分', '項目数', '令和8年度の得点計'],
  C['推移']['区分'].map(r => [r['区分'], `${r['項目']}項目`, `${r['点']}点`]),
  [40, 25, 35], {right: [1, 2]}));
K.push(p('', {after: 100}));
K.push(p(`合計得点は令和6年度295点、令和7年度339点、令和8年度${C['合計']}点と伸びています。`
       + `ただし ${C['推移']['区分'].find(x => x['区分'] === '悪化')['項目']}項目が下がり、`
       + `${C['推移']['区分'].find(x => x['区分'] === '3年連続0点')['項目']}項目は3年続けて0点です。`, {after: 140}));

K.push(h2('(4) 下がった項目'));
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
const DANIRU = ['計画への記載・要綱の整備・手順の明文化と、それに沿った実施・記録を要します。評価は国の該当状況調査への回答によります',
  'ニーズ調査・在宅介護実態調査・見える化分析・認知症施策推進計画の策定を行うこと自体が評価の対象となる取組です',
  '事業の実施・実績の水準・他機関との連携が前提です。村のご判断を要します',
  '国の統計から自動算定される成果指標群と、他の指標の結果として決まるものです',
  '個別の判定をしていません。全国の45％未満の市町村しか取得していない項目が主です'];
K.push(p(`未取得の配点を合計すると ${C['段'].reduce((a, b) => a + b['項目'], 0)}項目 ${C['段'].reduce((a, b) => a + b['点'], 0)}点ですが、`
       + `これは見かけの数です。実際に伸ばせるのは満点${C['満点']}点との差 ${C['伸びしろ']}点までです。`
       + `評価指標の枝の配点を足し上げると ${C['配点の総和']}点になりますが、`
       + `成果指標群のア／イなどは択一（より上位となった方で得点）であり、両方を足すと二重に数えます。`
       + `差 ${C['配点の総和'] - C['満点']}点は`
       + C['超過'].map(x => `${x['交付金'][0]}${x['目標']} ${x['超過']}点`).join('・')
       + `の超過でちょうど説明がつきます。取得の見込みで5つの段に分けました。`, {after: 140}));
K.push(table(['段', '内容', '項目数', '重複を除く', '見かけ', '実際に伸ばせる', '何が要るか'],
  C['段'].map((x, i) => [x['段'], x['内容'], `${x['項目']}項目`,
    `${x['重複を除く']}項目`, `${x['点']}点`, `${x['実際']}点`, DANIRU[i]]),
  [5, 15, 8, 9, 8, 10, 45], {right: [2, 3, 4, 5], fills: [KEYB, OK, null, null, null]}));
K.push(p('段Ａ・Ｂ・Ｃは見かけと実際が一致します（目標ごとに積み上げても満点100点に届かないことを'
       + '「6　見込める得点」で確かめています）。村にお示しする取得目標はこの3段です。', {after: 140}));
K.push(p('「重複を除く」は、目標Ⅳ（成果指標群）が推進交付金と介護保険保険者努力支援交付金で'
       + '同一の指標であり、集計表に2度現れることによるものです。点は両交付金で別に支払われるため'
       + '足して正しいのですが、取り組む対象の数としては1つです。'
       + `未取得の項目は重複を含めて ${C['段'].reduce((a, b) => a + b['項目'], 0)}項目、`
       + `重複を除くと ${C['段'].reduce((a, b) => a + b['重複を除く'], 0)}項目です。`, {after: 140}));

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
K.push(note('※ 黄色は段Ａ（記載で取得をめざせる）、緑は段Ｂ（本業務で取得をめざせる）です。'));

K.push(new Paragraph({children: [new PageBreak()]}));
K.push(h1('5　第1順位群　記載と本業務で取得をめざす13項目65点'));
K.push(p('段Ａ・Ｂの13項目は、新たな施設や大きな予算を要しません。'
       + '何をすれば評価の対象となるかと、本計画のどこで扱っているかを示します。', {after: 140}));
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
   ['＋Ａ 記載で取得をめざせる', `${ta}点`, sa(ta - C['全国平均合計']), sa(ta - C['県平均合計']),
    `${ran(ta)['全国']}／${C['全国団体数']}`, `${ran(ta)['県内']}／${C['県内団体数']}`],
   ['＋Ｂ 本業務で取得をめざせる', `${tb}点`, sa(tb - C['全国平均合計']), sa(tb - C['県平均合計']),
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
