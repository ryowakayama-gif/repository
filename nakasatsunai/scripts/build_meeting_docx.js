// 打合せ確認事項・影響度一覧（Word出力）
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ShadingType,
  Header, Footer, PageNumber, VerticalAlign,
} = require('docx');

const D = JSON.parse(fs.readFileSync(process.argv[3] || 'items.json', 'utf8'));

const JP  = { ascii: 'Yu Gothic', eastAsia: 'Yu Gothic', hAnsi: 'Yu Gothic' };
const MIN = { ascii: 'Yu Mincho', eastAsia: 'Yu Mincho', hAnsi: 'Yu Mincho' };

const INK = '1A2422', TEAL = '0E5E62', GREY = '5E6B67';
const RED = 'A33520', AMBER = '8A6512', GREEN = '2F6B45';
const RULE = 'C8D2CE', HEADBG = 'E6EDEB', ZEBRA = 'F5F8F7', CARDBG = 'EFF4F2';
const W = 9638;                                   // A4・余白20mm の本文幅

const IMPACT_COLOR = { '大': RED, '中': AMBER, '小': GREEN };
const STATE_COLOR = { '仮反映済': GREEN, '照会中': AMBER, '方針待ち': AMBER, '未確認': GREY };

/* ---------- 小物 ---------- */
const run = (text, o = {}) => new TextRun({ text, font: o.font || JP, size: o.size || 20, bold: o.bold, color: o.color || INK });

const para = (text, o = {}) => new Paragraph({
  spacing: { line: 300, before: o.before || 0, after: o.after == null ? 120 : o.after },
  indent: o.indent,
  children: [run(text, { size: o.size || 21, color: o.color, bold: o.bold })],
});

const h1 = (n, text) => new Paragraph({
  spacing: { before: 380, after: 160 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: TEAL, space: 4 } },
  children: [run(n + '　', { size: 26, bold: true, color: TEAL }), run(text, { size: 26, bold: true })],
});

const h2 = (text) => new Paragraph({
  spacing: { before: 280, after: 120 },
  children: [run(text, { size: 22, bold: true, color: TEAL })],
});

const note = (text) => new Paragraph({
  spacing: { before: 60, after: 160 }, indent: { left: 170 },
  children: [run(text, { size: 18, color: GREY })],
});

function cell(children, w, o = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: o.bg ? { type: ShadingType.CLEAR, fill: o.bg, color: 'auto' } : undefined,
    verticalAlign: o.top ? VerticalAlign.TOP : VerticalAlign.CENTER,
    margins: { top: 70, bottom: 70, left: 110, right: 110 },
    columnSpan: o.span,
    children: Array.isArray(children) ? children : [children],
  });
}

const BORDERS = {
  top: { style: BorderStyle.SINGLE, size: 6, color: RULE },
  bottom: { style: BorderStyle.SINGLE, size: 6, color: RULE },
  left: { style: BorderStyle.SINGLE, size: 6, color: RULE },
  right: { style: BorderStyle.SINGLE, size: 6, color: RULE },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 6, color: RULE },
  insideVertical: { style: BorderStyle.SINGLE, size: 6, color: RULE },
};

/** 単純な表：見出し＋データ行 */
function table(widths, header, rows, opt = {}) {
  const mk = (cells, o) => new TableRow({
    tableHeader: o.header,
    children: cells.map((c, i) => {
      const [text, so] = Array.isArray(c) ? c : [c, {}];
      const segs = String(text).split('\n');
      return cell(new Paragraph({
        spacing: { line: 270, after: 0 }, alignment: so.align,
        children: segs.map((s, k) => new TextRun({
          text: s, font: JP, size: o.header ? 18 : 18,
          bold: o.header || so.bold, color: so.color || INK, break: k > 0 ? 1 : undefined,
        })),
      }), widths[i], { bg: o.header ? HEADBG : o.bg, top: opt.top });
    }),
  });
  return new Table({
    columnWidths: widths,
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    borders: BORDERS,
    rows: [
      ...(header ? [mk(header, { header: true })] : []),
      ...rows.map((r, i) => mk(r, { bg: i % 2 ? ZEBRA : undefined })),
    ],
  });
}

/** 確認事項1件のカード */
function card(it) {
  const LW = 1750, RW = W - LW;
  const label = (s) => new Paragraph({
    spacing: { line: 270, after: 0 },
    children: [run(s, { size: 18, bold: true, color: TEAL })],
  });
  const val = (s) => new Paragraph({
    spacing: { line: 270, after: 0 },
    children: [run(s, { size: 18 })],
  });

  const rows = [];

  // 見出し行：ID・影響度・タイトル
  rows.push(new TableRow({
    children: [
      cell(new Paragraph({
        spacing: { line: 270, after: 0 },
        children: [
          run(it.id, { size: 20, bold: true, color: INK }),
          new TextRun({ text: '', break: 1 }),
          run('影響度 ' + it.impact, { size: 17, bold: true, color: IMPACT_COLOR[it.impact] }),
          new TextRun({ text: '', break: 1 }),
          run(it.state || '未確認', { size: 16, bold: true, color: STATE_COLOR[it.state] || GREY }),
        ],
      }), LW, { bg: CARDBG }),
      cell(new Paragraph({
        spacing: { line: 270, after: 0 },
        children: [run(it.title, { size: 21, bold: true })],
      }), RW, { bg: CARDBG }),
    ],
  }));

  const add = (l, v) => rows.push(new TableRow({
    children: [cell(label(l), LW, { top: true }), cell(val(v), RW, { top: true })],
  }));

  add('背景・現状', it.background);
  add('確認したいこと', it.ask);
  if (it.options && it.options.length) {
    rows.push(new TableRow({
      children: [
        cell(label('選択肢'), LW, { top: true }),
        cell(new Paragraph({
          spacing: { line: 270, after: 0 },
          children: it.options.flatMap((o, k) => [
            ...(k ? [new TextRun({ text: '', break: 1 })] : []),
            run('□ ' + o, { size: 18 }),
          ]),
        }), RW, { top: true }),
      ],
    }));
  } else {
    add('回答形式', '資料のご提供（自由記述）');
  }
  add('影響範囲', it.scope);
  add('未決の場合', it.risk);
  if (it.note) add('現在の状況', it.note);

  return new Table({
    columnWidths: [LW, RW],
    width: { size: W, type: WidthType.DXA },
    borders: BORDERS,
    rows,
  });
}

/* ================= 本文 ================= */
const body = [];
const items = D.items;
const cats = Object.keys(D.categories);
const count = (c, im) => items.filter((i) => (!c || i.cat === c) && (!im || i.impact === im)).length;

/* 表題 */
body.push(new Paragraph({ spacing: { after: 60 }, children: [run(D.meta.subtitle, { font: MIN, size: 24, color: GREY })] }));
body.push(new Paragraph({
  spacing: { after: 200 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: INK, space: 8 } },
  children: [run(D.meta.title, { font: MIN, size: 40, bold: true })],
}));
body.push(new Table({
  columnWidths: [1750, W - 1750],
  width: { size: W, type: WidthType.DXA },
  borders: {
    top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
    left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: RULE },
    insideVertical: { style: BorderStyle.NONE },
  },
  rows: [
    ['作成日', D.meta.date],
    ['確認事項', `全${items.length}件（影響度 大${count(null, '大')}件・中${count(null, '中')}件・小${count(null, '小')}件）`],
    ['根拠', D.meta.basis],
  ].map(([a, b]) => new TableRow({
    children: [
      cell(new Paragraph({ spacing: { line: 270, after: 0 }, children: [run(a, { size: 18, bold: true, color: TEAL })] }), 1750),
      cell(new Paragraph({ spacing: { line: 270, after: 0 }, children: [run(b, { size: 18 })] }), W - 1750),
    ],
  })),
}));

/* 1 本書の使い方 */
body.push(h1('１', '本書の使い方'));
body.push(para('本書は、中札内村下水道事業経営戦略（案）の修正を進めるにあたり、打合せの場でご確認いただきたい事項を整理したものである。各事項には影響度を付し、未決のまま残した場合に何が止まるのかを明記した。'));
body.push(para('回答は別添の「確認結果入力シート（Excel）」にご記入いただきたい。本書の確認事項IDと入力シートの行は1対1で対応している。選択肢を示した事項はいずれかを選び、必要に応じて補足をご記入いただければ、そのまま本文およびシミュレーションへ反映する。'));
body.push(para('なお、朱書き校正紙でご指示いただいた9件と、公表版との照合により根拠が確定した事項は、すでに本文およびシミュレーションへ反映済みである。本書はそれ以外の未決事項のみを扱う。'));

/* 2 影響度の区分 */
body.push(h1('２', '影響度の区分'));
body.push(table([900, W - 900], ['区分', '定義'],
  D.impactDef.map(([k, v]) => [[k, { bold: true, color: IMPACT_COLOR[k], align: AlignmentType.CENTER }], v]), { top: true }));

/* 3 サマリ */
body.push(h1('３', '確認事項サマリ'));
body.push(table(
  [3600, 1200, 1200, 1200, 1200, 1238],
  ['区分', '大', '中', '小', '計', '主な影響先'],
  [
    ...cats.map((c) => [
      `${c}．${D.categories[c]}`,
      [String(count(c, '大') || '−'), { align: AlignmentType.CENTER, bold: count(c, '大') > 0, color: count(c, '大') ? RED : GREY }],
      [String(count(c, '中') || '−'), { align: AlignmentType.CENTER, color: count(c, '中') ? AMBER : GREY }],
      [String(count(c, '小') || '−'), { align: AlignmentType.CENTER, color: count(c, '小') ? GREEN : GREY }],
      [String(count(c)), { align: AlignmentType.CENTER, bold: true }],
      { A: '本文p.2・p.19・p.20', B: '本文p.18・p.19・p.31', C: '収支計画全体', D: 'Excel・図表' }[c],
    ]),
    [
      ['合計', { bold: true }],
      [String(count(null, '大')), { align: AlignmentType.CENTER, bold: true, color: RED }],
      [String(count(null, '中')), { align: AlignmentType.CENTER, bold: true, color: AMBER }],
      [String(count(null, '小')), { align: AlignmentType.CENTER, bold: true, color: GREEN }],
      [String(items.length), { align: AlignmentType.CENTER, bold: true }],
      '',
    ],
  ]
));

body.push(h2('■　優先してご確認いただきたい事項（影響度 大）'));
body.push(para('次の5件は、決まらないと本文または収支計画を確定できない。打合せではこの5件を先にお願いしたい。'));
body.push(table([760, 2700, 1100, W - 4560], ['ID', '確認事項', '状態', '決まらないと止まること'],
  items.filter((i) => i.impact === '大').map((i) => [
    [i.id, { align: AlignmentType.CENTER, bold: true }], i.title,
    [i.state || '未確認', { align: AlignmentType.CENTER, bold: true, color: STATE_COLOR[i.state] || GREY }], i.risk,
  ]), { top: true }));

/* 4 確認事項 */
body.push(h1('４', '確認事項'));
for (const c of cats) {
  const list = items.filter((i) => i.cat === c);
  body.push(h2(`${c}．${D.categories[c]}　（${list.length}件）`));
  for (const it of list) {
    body.push(card(it));
    body.push(new Paragraph({ spacing: { after: 160 }, children: [] }));
  }
}

/* 5 未決時の作業への波及 */
body.push(h1('５', '未決のまま進めた場合の作業への波及'));
body.push(para('本文の図表はすべてExcelから貼り付けた画像であり、リンクされた表ではない。そのためExcelを修正するたびにWordへの貼り直しが発生する。C区分（経営方針）が未決のままExcelを再計算すると、確定後にもう一度すべての図表を貼り直すことになるため、C-1からC-3を先に決めていただきたい。'));
body.push(table([2200, W - 2200], ['未決の事項', '止まる作業'],
  [
    ['C-1 使用料改定の時期と率', 'シミュレーション4パターンの採用案が決まらず、p.27〜32の本文と図表、ロードマップが着手できない。交付金の支給要件に関わるため後戻りのコストが大きい。'],
    ['C-2 繰入金の前提', '経常収支比率が変わるため、目標達成の可否判定（p.32）が確定しない。'],
    ['C-3 令和8年度以降の投資配分', '減価償却費・支払利息・企業債残高が動き、収支計画とシミュレーションの全数値が変わる。p.25・p.26の表も差し替えとなる。'],
    ['B-2 将来人口の推計根拠', '有収水量・使用料収入の予測の説明が成り立たず、p.31とp.23の図表が確定しない。'],
    ['A-1 主要施設一覧', 'p.20が他団体の記載のまま残り、本文を完成できない。'],
  ], { top: true }));

body.push(new Paragraph({
  spacing: { before: 300, after: 0 },
  border: { top: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 6 } },
  children: [run('回答は別添「中札内村下水道事業経営戦略_確認結果入力シート」にご記入ください。確認事項IDが対応しています。', { size: 17, color: GREY })],
}));

/* ================= 出力 ================= */
const doc = new Document({
  styles: { default: { document: { run: { font: JP, size: 21, color: INK }, paragraph: { spacing: { line: 300 } } } } },
  sections: [{
    properties: { page: { margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } } },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT, spacing: { after: 0 },
          children: [run(`${D.meta.subtitle}　${D.meta.title}`, { size: 16, color: GREY })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER, spacing: { before: 0 },
          children: [new TextRun({ children: [PageNumber.CURRENT, ' / ', PageNumber.TOTAL_PAGES], font: JP, size: 16, color: GREY })],
        })],
      }),
    },
    children: body,
  }],
});

Packer.toBuffer(doc).then((b) => {
  fs.writeFileSync(process.argv[2], b);
  console.log('wrote', process.argv[2], b.length, 'bytes /', items.length, 'items');
});
