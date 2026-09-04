// 中札内村下水道事業経営戦略　修正箇所一覧（Word出力）
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, HeadingLevel, AlignmentType, BorderStyle, ShadingType,
  Header, Footer, PageNumber, VerticalAlign,
} = require('docx');

const JP = { ascii: 'Yu Gothic', eastAsia: 'Yu Gothic', hAnsi: 'Yu Gothic' };
const MIN = { ascii: 'Yu Mincho', eastAsia: 'Yu Mincho', hAnsi: 'Yu Mincho' };

const INK = '1A2422';
const TEAL = '0E5E62';
const GREY = '5E6B67';
const RED = 'A33520';
const GREEN = '2F6B45';
const AMBER = '8A6512';
const RULE = 'C8D2CE';
const HEADBG = 'E6EDEB';
const ZEBRA = 'F5F8F7';

const W = 9638;                     // A4（余白20mm）の本文幅

/* ---------- 小物 ---------- */
const t = (text, o = {}) => new TextRun({ text, font: JP, ...o });

const p = (text, o = {}) => new Paragraph({
  spacing: { line: 300, before: o.before ?? 0, after: o.after ?? 120 },
  alignment: o.align,
  indent: o.indent,
  children: Array.isArray(text) ? text : [t(text, { size: o.size ?? 21, color: o.color ?? INK, bold: o.bold })],
});

const h1 = (n, text) => new Paragraph({
  spacing: { before: 360, after: 160 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: TEAL, space: 4 } },
  children: [
    new TextRun({ text: n + '　', font: JP, size: 26, bold: true, color: TEAL }),
    new TextRun({ text, font: JP, size: 26, bold: true, color: INK }),
  ],
});

const h2 = (text) => new Paragraph({
  spacing: { before: 260, after: 120 },
  children: [new TextRun({ text, font: JP, size: 22, bold: true, color: TEAL })],
});

const note = (text) => new Paragraph({
  spacing: { before: 60, after: 160 },
  indent: { left: 170 },
  children: [new TextRun({ text, font: JP, size: 18, color: GREY })],
});

/** セル */
function cell(children, w, o = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: o.bg ? { type: ShadingType.CLEAR, fill: o.bg, color: 'auto' } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: Array.isArray(children) ? children : [children],
  });
}

/** 文字列（または [文字列, オプション] ）配列から1行つくる */
function row(cells, widths, o = {}) {
  return new TableRow({
    tableHeader: o.header,
    children: cells.map((c, i) => {
      const spec = Array.isArray(c) ? c : [c, {}];
      const [text, so] = spec;
      const runs = String(text).split('\n').map((seg, k) => new TextRun({
        text: seg, font: JP, size: o.header ? 17 : 17,
        bold: o.header || so.bold, color: so.color || (o.header ? INK : INK),
        break: k > 0 ? 1 : undefined,
      }));
      return cell(
        new Paragraph({ spacing: { line: 260, before: 0, after: 0 }, alignment: so.align, children: runs }),
        widths[i],
        { bg: o.header ? HEADBG : o.bg }
      );
    }),
  });
}

function table(widths, header, rows) {
  return new Table({
    columnWidths: widths,
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: RULE },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: RULE },
      left: { style: BorderStyle.SINGLE, size: 6, color: RULE },
      right: { style: BorderStyle.SINGLE, size: 6, color: RULE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 6, color: RULE },
      insideVertical: { style: BorderStyle.SINGLE, size: 6, color: RULE },
    },
    rows: [
      row(header, widths, { header: true }),
      ...rows.map((r, i) => row(r, widths, { bg: i % 2 ? ZEBRA : undefined })),
    ],
  });
}

const DONE = { text: '修正済', color: GREEN, bold: true };
const TODO = { text: '未着手', color: RED, bold: true };
const ASK = { text: '照会中', color: AMBER, bold: true };
const HOLD = { text: '方針待ち', color: AMBER, bold: true };
const st = (s) => [s.text, { color: s.color, bold: true, align: AlignmentType.CENTER }];

/* ================= 本文 ================= */
const body = [];

/* --- 表題 --- */
body.push(new Paragraph({
  spacing: { after: 60 },
  children: [new TextRun({ text: '中札内村下水道事業経営戦略（案）', font: MIN, size: 24, color: GREY })],
}));
body.push(new Paragraph({
  spacing: { after: 200 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: INK, space: 8 } },
  children: [new TextRun({ text: '修正箇所一覧', font: MIN, size: 40, bold: true, color: INK })],
}));

body.push(new Table({
  columnWidths: [1900, 7738],
  width: { size: W, type: WidthType.DXA },
  borders: {
    top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
    left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: RULE },
    insideVertical: { style: BorderStyle.NONE },
  },
  rows: [
    ['対象', '中札内村下水道事業経営戦略（案）　令和7（2025）年度〜令和16（2034）年度'],
    ['修正依頼', '朱書き校正紙（令和8年1月4日付コメント）＋提供資料3点（区域図・管渠延長・施設の概要）'],
    ['照合資料', '中札内村下水道事業経営戦略（R7〜R16）令和7年2月策定・公表版（経営比較分析表を含む）'],
    ['作成日', '令和8（2026）年9月4日'],
  ].map(([a, b]) => new TableRow({
    children: [
      cell(new Paragraph({ spacing: { line: 260, after: 0 }, children: [new TextRun({ text: a, font: JP, size: 18, bold: true, color: TEAL })] }), 1900),
      cell(new Paragraph({ spacing: { line: 260, after: 0 }, children: [new TextRun({ text: b, font: JP, size: 18, color: INK })] }), 7738),
    ],
  })),
}));

/* --- 1 --- */
body.push(h1('１', '本書の目的'));
body.push(p('本書は、中札内村から受領した朱書き校正紙による修正依頼と、提供資料および現行公表版との突合せ（影響範囲試算）により確定した修正必要箇所を一覧にまとめたものである。'));
body.push(p('修正箇所は、①自治体からの赤字修正、②影響範囲試算で確定した修正必要箇所、③村への照会または方針判断を要する論点の三層に整理した。①および②のうち根拠が確定しているものは本文（docx）およびシミュレーション（xlsx）へ反映済みである。'));

/* --- 2 --- */
body.push(h1('２', '自治体からの赤字修正箇所'));
body.push(p('朱書き校正紙で指示のあった9件。いずれも語句・年次の置換であり、試算値への影響はない。「町」→「村」および「県」→「道」は他団体の様式を流用した際の残存と考えられる。'));
body.push(table(
  [700, 1750, 2650, 2650, 1888],
  ['頁', '箇所', '修正前', '修正後', '対応'],
  [
    ['表紙', '発行課名', '中札内村役場　建設課', '中札内村役場　施設課', st(DONE)],
    ['p.2', '（2）本村の下水道処理の歩み', '処理場は1箇所で、町の中心部', '処理場は1箇所で、村の中心部', st(DONE)],
    ['p.2', '同上', 'これまでの整備により、町全体', 'これまでの整備により、村全体', st(DONE)],
    ['p.2', '地方公営企業法の適用時期', '令和5（2023）年4月から移行', '令和4（2022）年4月から移行', st(DONE)],
    ['p.10', '（3）組織の状況・組織図', '町長', '村長', st(DONE)],
    ['p.26', '（2）投資財源の予測', '国（県）補助金等を活用', '国（道）補助金等を活用', st(DONE)],
    ['p.34', '❸収支計画の説明【資本的収入】', '●国（県）補助金', '●国（道）補助金', st(DONE)],
    ['p.34', '同上・本文', '国庫（県）補助対象事業', '国庫（道）補助対象事業', st(DONE)],
    ['p.38', '理解促進に向けた広報及び啓発活動', '町ホームページ、SNS等', '村ホームページ、SNS等', st(DONE)],
  ]
));

body.push(h2('■　あわせて受領した資料提供依頼（コメント3件）'));
body.push(p('本文中の未解決コメント3件はいずれも「資料提供願います」であり、提供資料と1対1で対応する。過不足はない。'));
body.push(table(
  [700, 2400, 3200, 3338],
  ['頁', '本文の箇所', '受領資料', '内容'],
  [
    ['p.3', '《中札内村下水道処理区域図》', '区域.pdf', '中札内村都市計画図 1:5,000／中札内処理区・中札内浄化センター一般平面図'],
    ['p.19', '4．施設の状況（1）管渠の状況', '管渠延長.pdf', '年度別・管径別の管渠延長表（平成4〜令和5年度／合計25,270.66m）'],
    ['p.20', '（2）施設の状況　■処理場の概要', '施設の概要.pdf', '中札内村下水道事業の概要（全体計画・事業計画／処理場諸元）'],
  ]
));
body.push(note('※　目次の「本文完成後調整」は既存の注記であり、本文確定後に頁番号を更新する最終工程にあたる。'));

/* --- 3 --- */
body.push(h1('３', '影響範囲試算で確定した修正必要箇所'));

body.push(h2('３-１　朱書き漏れの是正（2件）'));
body.push(p('本文を全文検索した結果、朱書きが入っていない同種の誤記が2件残っていた。これらを是正したことで、本文中の「町」「県」は残存ゼロとなった。'));
body.push(table(
  [700, 1750, 2650, 2650, 1888],
  ['頁', '箇所', '修正前', '修正後', '対応'],
  [
    ['p.2', '（2）本村の下水道処理の歩み', '町は札内川流域に位置しており', '村は札内川流域に位置しており', st(DONE)],
    ['p.20', '（2）施設の状況', '本町における処理場は1施設で運用し', '本村における処理場は1施設で運用し', st(DONE)],
  ]
));

body.push(h2('３-２　公表版との照合により確定した修正（6件）'));
body.push(p('現行公表版（令和7年2月策定）および添付の経営比較分析表（令和5年度決算）との突合せにより根拠が確定したもの。いずれも反映済みである。'));
body.push(table(
  [700, 1750, 2650, 2650, 1888],
  ['頁／箇所', '項目', '修正前', '修正後', '対応'],
  [
    ['p.2', '地方公営企業法の適用区分\n（論点A）', '全部適用へ移行し', '一部適用（財務規定等の適用）へ移行し', st(DONE)],
    ['p.2', '処理場の供用開始', '平成9（1997）年4月1日に供用を開始', '平成9（1997）年3月に供用を開始', st(DONE)],
    ['p.20', '処理場の稼働年数', '平成22（2010）年度から供用開始し、\n令和6年度現在で15年', '平成8年度（平成9年3月）から\n供用開始し、令和6年度現在で28年', st(DONE)],
    ['Excel\n建設改良費', '令和7年度の事業費', '27,424千円\n（管渠更新0＋処理場27,424）', '93,100千円\n（管渠更新12,500＋処理場80,600）', st(DONE)],
    ['Excel\n下水道現況予測', '汚水処理原価の分母', '一般家庭等のみ 244,695㎥\n→ 446.36円/㎥', '総有収水量 362,295㎥\n→ 301.48円/㎥', st(DONE)],
    ['Excel\nシミュレーションまとめ', 'パターン1の使用料単価\n（論点E）', '180.8円（改定額10.3円）', '187.55円（改定額17.05円）', st(DONE)],
  ]
));

body.push(h2('■　確定の根拠'));
body.push(table(
  [2000, 7638],
  ['項目', '根拠'],
  [
    ['適用区分', '公表版1頁「法適の区分＝一部適用」、4頁「公営企業法の一部適用であり」、経営比較分析表の全体総括「令和4年度に公営企業会計に移行した」の三箇所が一致。朱書きは年次のみを修正しており、区分そのものの誤りは指摘されていなかった。'],
    ['供用開始', '公表版1頁「供用開始年度（供用開始後年数）＝平成8年度（28年）」。提供資料「施設の概要.pdf」の供用年月日「平成9年3月」は平成8年度内にあたり整合する。本文p.2の「4月1日」は平成9年度となってしまうため月表記に改めた。'],
    ['建設改良費', '公表版の投資・財政計画（様式第2号・資本的支出）の令和7年度は93,100千円。本文p.25も93,100千円であり、Excelの27,424千円のみが食い違っていた。修正後の財源は国庫40,300／企業債46,550／自己財源6,250となり、本文p.26の財源表と完全に一致する。10年間の投資総額も1,406,952千円＝約14.1億円となり本文記述と整合した。'],
    ['原価の分母', '公表版3頁「将来有収水量（工場排水以外＋工場排水）に使用料単価を乗じて推計」。工場排水は日平均490㎥/日＝年間117,600㎥で、これを含めた総有収水量を分母とするのが総務省定義。修正後の301.48円/㎥は令和5年度実績294.74円/㎥（経営比較分析表⑥）と連続する。'],
    ['使用料単価', 'パターン2〜4は各「予測(2)」シートへのリンクであるのに対し、パターン1のみ「シミュレーションまとめ」F8に180.8が直接入力されていた。改定率10%であれば170.5×1.10＝187.55円が正しい（180.8円では実質6.0%の改定にとどまる）。使用料収入の計算自体は170.5×110%で正しく処理されており、誤っていたのは表示値のみ。'],
  ]
));

body.push(h2('３-３　提供資料の反映に伴い修正が必要な箇所（未着手）'));
body.push(p('提供資料と現行本文を突き合わせた結果、p.19およびp.20には他団体（青森県三戸町）の記載がそのまま残存していることが判明した。図表の差し替えだけでなく、本文の数値記述も連動して書き換える必要がある。'));
body.push(table(
  [700, 2100, 2900, 2450, 1488],
  ['頁', '箇所', '現行の記載', '正しい内容', '対応'],
  [
    ['p.3', '処理区域図', '白紙（見出しのみ）', '区域.pdfを貼り込み', st(TODO)],
    ['p.19', '総管渠延長', '171,347.05m（管更生含む）', '25,270.66m（提供資料の合計値）', st(TODO)],
    ['p.19', '布設地区・年代', '約26%の昭和世代に布設された笹尾・城山地区', '該当なし。笹尾・城山は村内に存在せず、管渠は全て平成4年度以降の布設', st(TODO)],
    ['p.19', '耐用年数の到達', '本計画期間中の10年間で耐用年数を迎える管渠は全体の約26%', '0m（0%）。法定耐用年数50年で最古管の到達は令和24年度', st(TODO)],
    ['p.19', '下水道普及率', '村内での下水道普及率はほぼ100%', '69.48%（経営比較分析表・令和5年度決算）', st(TODO)],
    ['p.19', '整備延長グラフ', '1976〜2024年のデータ', '平成4〜令和5年度のデータで再作成', st(TODO)],
    ['p.20', '処理場の概要（表）', '三戸浄化センター／3,180㎥日／\n計画処理人口4,120人', '中札内浄化センター／1,520㎥日\n（760×2池）／計画区域内人口2,661人', st(TODO)],
    ['p.20', '主要施設一覧', '三戸浄化センターの設備構成', '中札内浄化センターの実データ（未入手）', st(ASK)],
    ['p.25', '管渠の経過年数', '管渠は古いもので約40年が過ぎています', '約33年（最古は平成4年度布設）', st(TODO)],
    ['p.25', '耐用年数の到達', '計画期間中に全体の約1／4が耐用年数を迎える', '0%', st(TODO)],
  ]
));
body.push(note('※　管渠老朽化率は経営比較分析表で0.00%（類似団体平均0.07%）であり、耐用年数超過の管渠が存在しないことを裏付けている。また分析欄にも「下水道管渠は供用開始後20年以上経過しているが、現状では不明水も見られず、目視点検においても大きな異常は認められない」と記載されている。'));

body.push(h2('３-４　シミュレーション（Excel）側に残る修正'));
body.push(table(
  [2300, 5850, 1488],
  ['シート／箇所', '内容', '対応'],
  [
    ['予測　A3:D3', '推計の自治体設定が「コード2441・青森県三戸町」のまま。12〜14行の人口は中札内村の値が直接入力されており推計式とは切れているが、設定行が残っていると出典説明と矛盾する。整理方法は論点Gの判断に連動。', st(HOLD)],
    ['下水過去', '行政区域内人口11,023→8,699人など三戸町の実績値。他シートからの参照はゼロのため、財政計算への波及はなく整理のみで足りる。', st(TODO)],
    ['簡水過去', '同じく三戸町の実績値。ただし「予測(2)P1-10%」〜「P4-30%」の31行・36行から参照されており、簡易水道の推計に波及する。中札内村の実績値が必要。', st(ASK)],
    ['下水道現況予測', '1行目が「3事業合算」（本村は特定環境保全公共下水道の1事業のみ）。13〜14行の普及率・水洗化率が#DIV/0!。外部リンクが小松市・東員町のブックを指しており、約960箇所ずつ参照している。', st(TODO)],
    ['建設改良費(2)', 'C列・D列（令和5・6年度）および37行以下に#REF!が残存。', st(TODO)],
  ]
));

/* --- 4 --- */
body.push(h1('４', '新たに判明した論点'));
body.push(p('朱書きでは指摘されていないが、提供資料および公表版との突合せにより判明した不整合。AおよびEは村の指示により修正済み。その他は方針判断または照会を要する。'));
body.push(table(
  [560, 2300, 5290, 1488],
  ['', '論点', '内容', '対応'],
  [
    ['A', '法適用区分の誤り', '本文p.2は「全部適用」としているが、公表版は三箇所で「一部適用」。区分そのものが誤っていた。', st(DONE)],
    ['B', '汚水処理原価が高い原因の説明', '本文p.18は「現状の施設能力に対して稼働が正しく賄えていない状況」とするが、公表版の分析欄は施設利用率71.78%（類似団体平均42.09%）をもって「適正規模での施設運営であるといえる」としており、記述が逆になっている。', st(HOLD)],
    ['C', '使用料改定の時期と率', '公表版は令和8年度から料金収入を61,683→98,928千円（＋60.4%）とする引上げを織り込み済み。改定案は令和9年度から10〜30%であり、前提の大幅な下方修正かつ1年の後ろ倒しとなる。社会資本整備総合交付金の支給要件（使用料適正化ロードマップ）に関わるため、改定理由の明記または方針のすり合わせが必要。', st(HOLD)],
    ['D', '他会計補助金の前提', '公表版は令和7年度55,771千円とするのに対し、改定案のExcelは88,514千円（公表版の令和6年度値）を置いており、差は＋32,743千円。繰入金は経常収支比率を直接押し上げるため、シミュレーションの合否判定に影響する。', st(HOLD)],
    ['E', 'パターン1の使用料単価', '「シミュレーションまとめ」F8のみ手入力の180.8円。170.5×110%＝187.55円が正しい。計算自体は正しく回っており、誤りは表示値のみ。', st(DONE)],
    ['F', '処理能力の逆算値の不一致', '令和5年度の有収水量357,117㎥を有収率101.76%で割り戻すと総処理水量は約962㎥/日。施設利用率71.78%で割ると1日処理能力は約1,340㎥/日となり、提供資料の1,520㎥/日（760×2池）と約12%ずれる。認可能力と現有能力の違いか、分析表の算定根拠の問題か。', st(ASK)],
    ['G', '将来人口の推計根拠の矛盾', '同じ節（Ⅲ-1 有収水量の予測）の中で、説明文は「国立社会保障・人口問題研究所による推計」、前提条件ボックスは「第7期中札内村まちづくり計画より村独自推計」としている。公表版3頁は村独自推計であり、誤りは説明文のほうと考えられる。あわせて普及率・水洗化率の平均年度も令和4〜6年度（説明文）と令和3〜5年度（前提条件）で食い違っている。', st(HOLD)],
  ]
));

/* --- 5 --- */
body.push(h1('５', '対応状況サマリ'));
body.push(table(
  [2400, 1400, 5838],
  ['区分', '件数', '内容'],
  [
    [['本文（docx）修正済', { bold: true }], ['14箇所', { align: AlignmentType.CENTER, color: GREEN, bold: true }], '朱書きどおり9件＋朱書き漏れ2件＋論点A・供用開始年3件'],
    [['Excel修正済', { bold: true }], ['109セル', { align: AlignmentType.CENTER, color: GREEN, bold: true }], 'パターン1使用料単価3セル＋令和7年度建設改良費6セル＋汚水処理原価の分母100セル'],
    [['資料反映が必要', { bold: true }], ['10箇所', { align: AlignmentType.CENTER, color: RED, bold: true }], 'p.3・p.19・p.20・p.25の本文および図表。うち1件は資料未入手'],
    [['Excel整理が必要', { bold: true }], ['5項目', { align: AlignmentType.CENTER, color: RED, bold: true }], '三戸町データの残存、#REF!、外部リンク、事業数の表記'],
    [['方針判断待ち', { bold: true }], ['5論点', { align: AlignmentType.CENTER, color: AMBER, bold: true }], '論点B・C・D・G および 予測シートの自治体設定'],
    [['村への照会中', { bold: true }], ['3件', { align: AlignmentType.CENTER, color: AMBER, bold: true }], '主要施設一覧／管渠延長の基準時点／全体計画人口の出典'],
  ]
));

/* --- 6 --- */
body.push(h1('６', '村への確認事項'));
body.push(p('着手前に回答を要する項目。いずれもp.19・p.20に関係する。'));
body.push(table(
  [560, 2400, 6678],
  ['', '項目', '内容'],
  [
    ['1', '中札内浄化センターの主要施設一覧', '敷地面積・放流水質・設備構成（ポンプ台数・池数・脱水機等）。現行p.20の表は三戸浄化センターの諸元であり、提供資料にも公表版にも代替となるデータがない。施設台帳または竣工図書からの提供を依頼したい。'],
    ['2', '管渠延長25,270.66mの基準時点', '提供資料の経過年数は令和7年度基準であるのに対し、本文は「令和5年度末時点」としており、時点を揃える必要がある。あわせて事業計画上の総延長27,560mとの差（約2,290m）の内訳を確認したい。'],
    ['3', '全体計画人口2,810人の出典', '本文p.2の記載。提供資料の計画区域内人口2,661人（事業計画・令和7年度目標）との関係を確認したい。'],
  ]
));
body.push(note('※　供用開始年月・工場排水の取扱い・水洗化率については、公表版および経営比較分析表により確定したため照会は不要となった。'));

/* --- 7 --- */
body.push(h1('７', '今後の作業手順と留意事項'));
body.push(p('本文の図表はすべてExcelから貼り付けた画像（EMF 27点）であり、リンクされた表ではない。したがってExcelを修正するたびにWordへの貼り直しが発生するため、Excelを確定させてからWordに貼るのが手戻りのない唯一の順序となる。'));
body.push(table(
  [700, 8938],
  ['順', '作業'],
  [
    ['1', '修正版シミュレーション（xlsx）をExcelで開き、全再計算を実行する。令和7年度の建設改良費が＋65,676千円となるため、減価償却費・支払利息・企業債残高が令和7年度以降で動く。経常収支比率・経費回収率・各パターンの判定を確認する。'],
    ['2', '論点B・C・D・Gの方針を決定する。特に論点C（使用料改定の時期と率）は公表版の前提を大幅に見直すことになるため、村との協議が必要。'],
    ['3', '村への照会3件の回答を受領し、p.20の処理場概要および主要施設一覧を確定する。'],
    ['4', '区域.pdfをp.3に貼り込み、管渠延長.pdfからグラフを再作成のうえp.19の本文（延長・耐用年数・普及率）を書き換える。'],
    ['5', '再計算後のExcelから、p.25（投資スケジュール）・p.26（投資財源）・p.32（シミュレーション検証）の各表を貼り直す。'],
    ['6', '目次の頁番号を更新し（「本文完成後調整」）、本文中のコメント3件を解決する。'],
  ]
));

body.push(new Paragraph({ spacing: { before: 300, after: 0 }, children: [] }));
body.push(new Paragraph({
  spacing: { before: 200, after: 0 },
  border: { top: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 6 } },
  children: [new TextRun({
    text: '修正版の本文（docx）およびシミュレーション（xlsx）、修正スクリプト、変更履歴は別途納品済み。',
    font: JP, size: 17, color: GREY,
  })],
}));

/* ================= 出力 ================= */
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: JP, size: 21, color: INK }, paragraph: { spacing: { line: 300 } } },
    },
  },
  sections: [{
    properties: { page: { margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } } },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 0 },
          children: [new TextRun({ text: '中札内村下水道事業経営戦略（案）　修正箇所一覧', font: JP, size: 16, color: GREY })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0 },
          children: [new TextRun({ children: [PageNumber.CURRENT, ' / ', PageNumber.TOTAL_PAGES], font: JP, size: 16, color: GREY })],
        })],
      }),
    },
    children: body.filter(Boolean),
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(process.argv[2] || 'out.docx', buf);
  console.log('wrote', process.argv[2] || 'out.docx', buf.length, 'bytes');
});
