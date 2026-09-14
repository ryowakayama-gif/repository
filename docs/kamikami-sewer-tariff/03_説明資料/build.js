const fs = require('fs');
const d = require('docx');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, ImageRun,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, PageBreak,
  Header, Footer, PageNumber, convertMillimetersToTwip,
} = d;

const DIR = '/tmp/claude-0/-home-user-repository/670c168c-8281-57ba-9df0-b54358bb5879/scratchpad/';
const FONT = 'MS Pゴシック';
const TEAL = '12939E', RUST = 'CF5B1A', INK = '1A1A1A', MUTED = '5F6F74';
const HEAD_BG = 'DCEBED', ACCENT_BG = 'FBEAE0', ZEBRA = 'F4F7F7';
const TW = 9600;   // 表の総幅（DXA）

const P = (text, o = {}) => new Paragraph({
  alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 120, line: 300 },
  indent: o.indent, border: o.border,
  children: [new TextRun({ text, bold: o.bold, size: o.size ?? 21, color: o.color ?? INK, font: FONT })],
});
const RUNS = (runs, o = {}) => new Paragraph({
  alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 120, line: 300 },
  children: runs.map(r => new TextRun({ text: r.t, bold: r.b, color: r.c ?? INK, size: r.s ?? 21, font: FONT })),
});
const H1 = t => new Paragraph({
  heading: HeadingLevel.HEADING_1, spacing: { before: 320, after: 160 }, keepNext: true,
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: TEAL, space: 4 } },
  children: [new TextRun({ text: t, bold: true, size: 26, color: INK, font: FONT })],
});
const H2 = t => new Paragraph({
  heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 100 }, keepNext: true,
  children: [new TextRun({ text: t, bold: true, size: 22, color: TEAL, font: FONT })],
});
const CAPTION = t => new Paragraph({
  spacing: { before: 60, after: 200 },
  children: [new TextRun({ text: t, size: 17, color: MUTED, font: FONT })],
});
const FIGTITLE = t => new Paragraph({
  spacing: { before: 160, after: 60 }, keepLines: true,
  children: [new TextRun({ text: t, bold: true, size: 19, color: INK, font: FONT })],
});
const IMG = (file, w, h) => new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { after: 40 },
  children: [new ImageRun({ type: 'png', data: fs.readFileSync(DIR + file), transformation: { width: w, height: h } })],
});

function cell(text, o = {}) {
  return new TableCell({
    width: { size: o.w, type: WidthType.DXA },
    shading: o.bg ? { type: ShadingType.CLEAR, fill: o.bg, color: 'auto' } : undefined,
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    verticalAlign: 'center',
    children: [new Paragraph({
      alignment: o.align ?? (o.num ? AlignmentType.RIGHT : AlignmentType.LEFT),
      spacing: { after: 0, line: 260 },
      children: [new TextRun({ text: String(text), bold: o.bold, size: o.size ?? 18,
                               color: o.color ?? INK, font: FONT })],
    })],
  });
}
function table(widths, head, body, opts = {}) {
  const rows = [new TableRow({
    tableHeader: true,
    children: head.map((t, i) => cell(t, { w: widths[i], bg: HEAD_BG, bold: true,
                                           align: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER })),
  })];
  body.forEach((r, ri) => {
    const hi = (opts.highlight || []).includes(ri);
    rows.push(new TableRow({
      children: r.map((t, i) => cell(t, {
        w: widths[i], num: i > 0, bold: hi,
        bg: hi ? ACCENT_BG : (ri % 2 === 1 ? ZEBRA : undefined),
      })),
    }));
  });
  return new Table({
    columnWidths: widths, width: { size: TW, type: WidthType.DXA }, rows,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: 'A9BEBE' },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: 'A9BEBE' },
      left: { style: BorderStyle.SINGLE, size: 6, color: 'A9BEBE' },
      right: { style: BorderStyle.SINGLE, size: 6, color: 'A9BEBE' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: 'D3DEDE' },
      insideVertical: { style: BorderStyle.SINGLE, size: 4, color: 'D3DEDE' },
    },
  });
}
const NOTE = (label, lines) => new Table({
  columnWidths: [TW], width: { size: TW, type: WidthType.DXA },
  borders: {
    top: { style: BorderStyle.SINGLE, size: 4, color: 'E0CBBC' },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: 'E0CBBC' },
    left: { style: BorderStyle.SINGLE, size: 18, color: RUST },
    right: { style: BorderStyle.SINGLE, size: 4, color: 'E0CBBC' },
    insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE },
  },
  rows: [new TableRow({ children: [new TableCell({
    width: { size: TW, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: 'FDF6F1', color: 'auto' },
    margins: { top: 140, bottom: 140, left: 180, right: 160 },
    children: [
      new Paragraph({ spacing: { after: 70 },
        children: [new TextRun({ text: label, bold: true, size: 18, color: RUST, font: FONT })] }),
      ...lines.map(l => new Paragraph({ spacing: { after: 60, line: 290 },
        children: [new TextRun({ text: l, size: 20, color: INK, font: FONT })] })),
    ],
  })] })],
});

const kids = [];

// ---------------- 表題 ----------------
kids.push(new Paragraph({
  spacing: { before: 200, after: 60 },
  children: [new TextRun({ text: '階上町下水道事業', size: 22, color: TEAL, bold: true, font: FONT })],
}));
kids.push(new Paragraph({
  spacing: { after: 100 },
  children: [new TextRun({ text: '下水道使用料の改定について', size: 40, bold: true, color: INK, font: FONT })],
}));
kids.push(new Paragraph({
  spacing: { after: 240 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: INK, space: 6 } },
  children: [new TextRun({ text: '― 6〜10㎥区分の是正を含む料金体系の見直し案 ―', size: 24, color: MUTED, font: FONT })],
}));
kids.push(RUNS([{ t: '令和8年9月14日　　建設課', s: 19, c: MUTED }], { after: 300 }));

kids.push(NOTE('この資料の要点', [
  '① 現行体系は6〜10㎥の単価が40.7円と極端に低く、この区分だけで従量水量の約4割を流しながら収入は5%にとどまっている。',
  '② この区分を11〜50㎥と同じ単価に揃えるだけで、年間約849万円（現行比+19.8%）の増収余地がある。',
  '③ ただし一度に揃えると月10㎥の世帯が+57.4%となるため、段階的に是正する。',
  '④ 推奨案は「基本使用料は据置・6〜10㎥を40.7円→80円」。全体+10%で、月5㎥以下の世帯は据置となる。',
]));

// ---------------- 1 ----------------
kids.push(H1('1　本資料の趣旨'));
kids.push(P('令和7年度の調定データ9,070件と令和6年度決算統計を分析した結果、現行の使用料体系に構造的な偏りがあることが判明した。本資料は、その内容と是正案を整理したものである。'));
kids.push(P('なお、経費回収率（決算統計32表ベース）は公共下水道49.9%・漁業集落排水36.9%であり、経営戦略が令和16年度の目標として掲げる水準（公共40%・漁集26%）は既に達成している。したがって本件は「目標未達を埋める改定」ではなく、「一般会計繰入をどこまで減らすか」と「体系の偏りをどう正すか」の判断である。'));

// ---------------- 2 ----------------
kids.push(H1('2　現行使用料体系の構造'));
kids.push(H2('2-1　現行の料金表'));
kids.push(table([2200, 2200, 1700, 1750, 1750],
  ['区分（1か月あたり・税込）', '基本使用料（5㎥まで）', '6〜10㎥', '11〜50㎥', '51㎥〜'],
  [['一般汚水', '1,108.8円', '40.7円', '191.4円', '220.0円'],
   ['（税抜換算）', '1,008円', '37円', '174円', '200円']]));
kids.push(CAPTION('表1　現行の下水道使用料（公衆浴場は基本使用料1,108.8円＋6㎥〜62.7円/㎥）。調定は隔月のため、請求額はこの2倍となる。'));

kids.push(H2('2-2　単価は月10㎥で底を打つ'));
kids.push(P('1㎥あたりの負担額を使用量別に計算すると、月10㎥で131.2円/㎥まで下がり、そこから緩やかに戻るU字カーブになる。基本使用料があるため少量ほど平均単価が高くなるのは体系上自然だが、6〜10㎥の単価が低すぎるため底が深い。'));
kids.push(IMG('fig1.png', 620, 268));
kids.push(FIGTITLE('図1　月使用量別の1㎥あたり負担額（現行・一般汚水・税込）'));
kids.push(P('51㎥超の単価は「220円 −  2,031.7円÷使用量」で表され、220円/㎥に漸近する。つまりどれだけ大量に使っても、月5㎥の世帯が負担している221.8円/㎥を上回る層は存在しない。公共下水道の使用量の中央値は月10.5㎥であり、ちょうど単価の底に中央値の世帯がいる。', { after: 200 }));

kids.push(H2('2-3　6〜10㎥は「水量の4割・収入の5%」'));
kids.push(IMG('fig2.png', 620, 240));
kids.push(FIGTITLE('図2　区分別の収入シェアと従量水量シェア（公共下水道・令和7年度）'));
kids.push(table([2700, 1750, 1550, 1850, 1750],
  ['区分', '収入（円）', '収入シェア', '従量水量（㎥）', '水量シェア'],
  [['基本使用料（5㎥まで）', '15,864,710', '44.5%', '—', '—'],
   ['6〜10㎥', '1,888,236', '5.3%', '46,394', '37.9%'],
   ['11〜50㎥', '12,699,581', '35.7%', '66,351', '54.1%'],
   ['51㎥〜（大口）', '2,156,440', '6.1%', '9,802', '8.0%'],
   ['特殊算定（日割・異動等）', '3,007,563', '8.4%', '—', '—']], { highlight: [1] }));
kids.push(CAPTION('表2　公共下水道の区分別収入。従量水量は2か月分の合計。漁業集落排水もほぼ同じ構造（6〜10㎥は水量33.9%・収入5.5%）。'));

// ---------------- 3 ----------------
kids.push(H1('3　6〜10㎥区分の是正余地'));
kids.push(P('6〜10㎥の単価40.7円は、11〜50㎥の191.4円の4.7分の1である。仮にこの区分を11〜50㎥と同じ単価に揃えると、2事業合計で次の増収となる。'));
kids.push(table([4200, 2700, 2700],
  ['算定', '数値', '備考'],
  [['6〜10㎥の従量水量（2事業計）', '56,362㎥', '2か月分の合計'],
   ['現行単価との差', '150.7円/㎥', '191.4円 − 40.7円'],
   ['増収額', '8,493,753円', '56,362㎥ × 150.7円'],
   ['現行調定額に対する比率', '+19.8%', '現行 42,987,132円']], { highlight: [2, 3] }));
kids.push(CAPTION('表3　6〜10㎥を11〜50㎥と同単価に揃えた場合の増収余地。'));
kids.push(NOTE('つまり', [
  '6〜10㎥区分の是正だけで、改定率19.8%に相当する増収余地がある。',
  '言い換えれば、この区分を放置したまま全体を一律に引き上げると、本来この区分が負担すべき分を他の区分が肩代わりし続けることになる。',
]));

// ---------------- 4 ----------------
kids.push(H1('4　是正水準の選択肢'));
kids.push(P('全体の改定率を+10%に固定し、基本使用料を据え置いたうえで、6〜10㎥をどこまで引き上げるかを4段階で比較した。6〜10㎥を上げるほど11〜50㎥は抑えられるため、負担の配分が変わる。'));
kids.push(table([2150, 1200, 1100, 1400, 1250, 1250, 1250],
  ['是正水準', '6〜10㎥', '是正率', '11〜50㎥', '月10㎥', '月20㎥', '月50㎥'],
  [['是正なし', '40.7円', '0%', '235.9円', '±0%', '+13.8%', '+19.8%'],
   ['弱い是正', '60円', '12.8%', '223.0円', '+7.4%', '+12.8%', '+15.2%'],
   ['中程度の是正（推奨）', '80円', '26.1%', '209.5円', '+15.0%', '+11.7%', '+10.3%'],
   ['強い是正', '110円', '46.0%', '189.4円', '+26.4%', '+10.1%', '+3.0%'],
   ['（参考）完全是正', '191.4円', '100%', '134.7円', '+57.4%', '+5.8%', '−16.9%']], { highlight: [2] }));
kids.push(CAPTION('表4　是正水準別の料金と世帯影響（いずれも全体+10%・基本使用料は据置・51㎥〜は242.0円）。是正率は40.7円から191.4円までの到達度。'));
kids.push(IMG('fig3.png', 620, 268));
kids.push(FIGTITLE('図3　是正水準別の負担増減率（月使用量別）'));
kids.push(NOTE('完全是正は一度にはできない', [
  '6〜10㎥を一気に191.4円へ揃えると、月10㎥の世帯は+57.4%となる一方、月30㎥以上はむしろ減額（月50㎥で−16.9%）となり、負担の逆転が生じる。',
  'したがって是正は段階的に行う必要がある（第7章）。',
]));

// ---------------- 5 ----------------
kids.push(H1('5　推奨案と各家庭への影響'));
kids.push(H2('5-1　推奨案の料金表'));
kids.push(P('「基本使用料は据置、6〜10㎥を40.7円→80円」を推奨する。少量利用者（月5㎥以下・件数の約2割）を据え置きながら、最も割安な区分を是正できる。'));
kids.push(table([2000, 2300, 1750, 1750, 1800],
  ['料金案（1か月・税込）', '基本使用料（5㎥まで）', '6〜10㎥', '11〜50㎥', '51㎥〜'],
  [['現行', '1,108.8円', '40.7円', '191.4円', '220.0円'],
   ['推奨案 +5%', '1,108.8円（据置）', '60円', '200.7円', '231.0円'],
   ['推奨案 +10%', '1,108.8円（据置）', '80円', '209.5円', '242.0円'],
   ['推奨案 +15%', '1,108.8円（据置）', '90円', '225.1円', '253.0円']], { highlight: [2] }));
kids.push(CAPTION('表5　推奨案の料金表。区分の数は現行のままであり、条例改正は単価の差替えのみで足りる。'));

kids.push(H2('5-2　各家庭への影響（推奨案+10%）'));
kids.push(P('調定は隔月のため、実際の請求額は月額の2倍となる。次表は2か月分の請求額で示した。'));
kids.push(table([1100, 1800, 1700, 1700, 1650, 1650],
  ['月使用量', '世帯像の目安', '現行（2か月）', '推奨案（2か月）', '差額', '増減率'],
  [['5㎥', '単身・高齢世帯', '2,217円', '2,217円', '±0円', '±0%'],
   ['8㎥', '2人世帯', '2,461円', '2,697円', '+236円', '+9.6%'],
   ['10㎥', '2人世帯', '2,624円', '3,017円', '+393円', '+15.0%'],
   ['15㎥', '3〜4人世帯', '4,538円', '5,112円', '+574円', '+12.6%'],
   ['20㎥', '4人世帯', '6,452円', '7,207円', '+755円', '+11.7%'],
   ['25㎥', '5人世帯', '8,366円', '9,302円', '+936円', '+11.2%'],
   ['30㎥', '5人以上世帯', '10,280円', '11,397円', '+1,117円', '+10.9%'],
   ['40㎥', '小規模事業所', '14,108円', '15,587円', '+1,479円', '+10.5%'],
   ['50㎥', '小規模事業所', '17,936円', '19,777円', '+1,841円', '+10.3%'],
   ['100㎥', '大口事業所', '39,936円', '43,977円', '+4,041円', '+10.1%']], { highlight: [0, 2] }));
kids.push(CAPTION('表6　推奨案+10%における2か月分請求額の変化（一般汚水・税込）。'));

kids.push(H2('5-3　どの世帯がどれだけいるか'));
kids.push(table([1650, 1950, 1150, 1250, 1250, 1150, 1200],
  ['月使用量帯', '世帯像の目安', '公共 件数', '公共 件数%', '公共 金額%', '漁集 件数', '漁集 件数%'],
  [['月5㎥以下', '単身・高齢世帯など', '1,657', '21.7%', '10.3%', '285', '19.9%'],
   ['月6〜10㎥', '2人世帯など', '1,870', '24.5%', '12.8%', '296', '20.7%'],
   ['月11〜20㎥', '3〜4人世帯など', '2,384', '31.2%', '29.1%', '505', '35.3%'],
   ['月21〜30㎥', '5人以上世帯など', '935', '12.2%', '21.5%', '247', '17.3%'],
   ['月31〜50㎥', '小規模事業所など', '244', '3.2%', '8.6%', '69', '4.8%'],
   ['月51㎥〜', '大口（事業所・公共施設）', '64', '0.8%', '9.3%', '22', '1.5%'],
   ['特殊算定', '日割・異動等', '487', '6.4%', '8.4%', '5', '0.3%']], { highlight: [0, 1] }));
kids.push(CAPTION('表7　月使用量帯別の構成（令和7年度調定）。据置となる月5㎥以下は件数の約2割、金額の約1割を占める。'));

// ---------------- 6 ----------------
kids.push(H1('6　増収額と経費回収率'));
kids.push(table([1700, 1900, 1600, 1500, 1500, 1400],
  ['料金案', '増収額（年額）', '増収率', '経費回収率 公共', '経費回収率 漁集', '是正率'],
  [['現行', '—', '—', '49.9%', '36.9%', '0%'],
   ['推奨案 +5%', '2,148,290円', '+5.00%', '52.4%', '38.7%', '12.8%'],
   ['推奨案 +10%', '4,295,086円', '+9.99%', '54.8%', '40.6%', '26.1%'],
   ['推奨案 +15%', '6,447,884円', '+15.00%', '57.3%', '42.6%', '32.7%']], { highlight: [2] }));
kids.push(CAPTION('表8　増収額はR7年度奇数月6回調定に各案を当てはめた再計算（税込）。経費回収率はR6決算統計（32表）に増収率を乗じたもので、汚水処理費はR6水準で据置。'));
kids.push(NOTE('繰入金への影響', [
  '基準内繰入金は改定しても変わらない。分流式下水道等に要する経費は「基準額のうち使用料収入で賄いきれない分」として満額繰入となっており、+15%程度の改定ではこの状態が続くためである。',
  '基準内繰入金が減り始めるのは、使用料収入が汚水維持管理費を超えたとき（公共で現行の2.00倍、漁集で2.71倍）であり、今回の改定率では届かない。',
  'したがって増収分は、基準外繰入金の削減（一般会計負担の軽減）または事業収支の改善に回る。',
]));

// ---------------- 7 ----------------
kids.push(H1('7　段階的な是正の考え方'));
kids.push(P('経営戦略は令和8年度に+10%、令和13年度にさらに+10%（現行比+20%）の改定を織り込んでいる。この2段階に合わせて6〜10㎥を是正すると、次のようになる。'));
kids.push(table([2400, 1400, 1300, 1400, 1550, 1550],
  ['段階', '6〜10㎥', '是正率', '11〜50㎥', '51㎥〜', '現行比'],
  [['現行', '40.7円', '0%', '191.4円', '220.0円', '—'],
   ['第1段階（令和8年度）', '80円', '26.1%', '209.5円', '242.0円', '+10%'],
   ['第2段階（令和13年度）', '120円', '52.6%', '227.2円', '264.0円', '+20%']], { highlight: [1] }));
kids.push(CAPTION('表9　段階的是正の料金表（1か月・税込）。基本使用料はいずれの段階も1,108.8円で据置。'));
kids.push(IMG('fig4.png', 620, 254));
kids.push(FIGTITLE('図4　段階的是正による単価の段差の縮小'));
kids.push(P('現行4.70倍の段差は、第1段階で2.62倍、第2段階で1.89倍まで縮まる。2段階を経ても完全是正（1.00倍）には至らないが、10年かけて偏りを大きく改善できる。'));
kids.push(table([1300, 1650, 1650, 1450, 1750, 1800],
  ['月使用量', '現行（2か月）', '第1段階（2か月）', '増減率', '第2段階（2か月）', '現行比増減率'],
  [['5㎥', '2,217円', '2,217円', '±0%', '2,217円', '±0%'],
   ['10㎥', '2,624円', '3,017円', '+15.0%', '3,417円', '+30.2%'],
   ['20㎥', '6,452円', '7,207円', '+11.7%', '7,961円', '+23.4%'],
   ['30㎥', '10,280円', '11,397円', '+10.9%', '12,505円', '+21.6%'],
   ['50㎥', '17,936円', '19,777円', '+10.3%', '21,593円', '+20.4%'],
   ['100㎥', '39,936円', '43,977円', '+10.1%', '47,993円', '+20.2%']], { highlight: [0] }));
kids.push(CAPTION('表10　段階別の2か月請求額（一般汚水・税込）。'));

// ---------------- 8 ----------------
kids.push(H1('8　留意事項'));
[['調定データの範囲',
  '本資料の増収試算は、令和7年度の奇数月6回調定（9,070件・42,987,132円）に基づく。通常の隔月検針者については概ね12か月相当だが、偶数月調定および毎月検針者の一部を含まない。毎月検針の大口は年12回のうち6回しか捕捉できていないため、増収額はやや過小に出ている可能性がある。'],
 ['使用水量の推計',
  '調定明細に使用水量の記載がないため、請求金額から算定式を逆引きして推計している。基本料金帯（2,217円）は1〜10㎥のいずれでも同額のため水量を特定できず、この分は推計の対象外としている。ただし改定後の金額には影響しない。'],
 ['経費回収率の定義',
  '決算統計32表の汚水処理費（維持管理費の汚水分のみ）に対する使用料収入の比率である。経営戦略が用いている経費回収率（公共35.0%・漁集25.8%、令和4年度）は資本費を含む分母であり、本資料の数値とは連続しない。'],
 ['汚水処理費の前提',
  '経費回収率の試算は汚水処理費を令和6年度水準で据え置いている。維持管理費の変動は織り込んでいない。'],
 ['端数処理',
  '2か月分の請求額は「税抜額×1.1」の1円未満を切り捨てて算定している。条例改正にあたっては端数処理の方法を明記する必要がある。']
].forEach(([t, b]) => {
  kids.push(RUNS([{ t: '■ ' + t, b: true, s: 20 }], { before: 120, after: 40 }));
  kids.push(P(b, { indent: { left: 240 }, after: 120 }));
});

kids.push(H1('参考　数値の出典'));
[['令和7年度 調定簿明細（漁業集落排水・公共下水道）', '調定件数・調定額・金額別度数分布'],
 ['【R7】使用料集計ブック', '金額別度数分布、水量区分別構成'],
 ['令和6年度 地方公営企業決算状況調査 32表・40表', '汚水処理費、使用料収入、有収水量、繰入金'],
 ['階上町下水道事業経営戦略（改定）令和7年2月', '改定の織り込み、経営目標、近隣比較（表2.9）'],
 ['階上町ホームページ', '現行の使用料体系']
].forEach(([a, b]) => kids.push(RUNS([{ t: '・' + a + '　', s: 19 }, { t: b, s: 19, c: MUTED }], { after: 60 })));
kids.push(P('本資料の数値は、別添の試算エビデンス（Excel）で全て追跡できる。', { before: 160, size: 19, color: MUTED }));

// ---------------- 出力 ----------------
const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 21, color: INK } } } },
  sections: [{
    properties: {
      page: {
        size: { width: convertMillimetersToTwip(210), height: convertMillimetersToTwip(297) },
        margin: { top: convertMillimetersToTwip(22), bottom: convertMillimetersToTwip(20),
                  left: convertMillimetersToTwip(20), right: convertMillimetersToTwip(20) },
      },
    },
    headers: { default: new Header({ children: [new Paragraph({
      alignment: AlignmentType.RIGHT, spacing: { after: 0 },
      children: [new TextRun({ text: '階上町下水道使用料の改定について', size: 16, color: MUTED, font: FONT })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], size: 17, color: MUTED, font: FONT })] })] }) },
    children: kids,
  }],
});
Packer.toBuffer(doc).then(b => {
  fs.writeFileSync(DIR + 'doc/下水道使用料改定説明資料.docx', b);
  console.log('docx written:', b.length, 'bytes');
});
