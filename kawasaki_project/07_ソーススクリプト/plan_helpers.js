/**
 * 川崎町高齢者保健福祉計画・第10期介護保険事業計画
 * 計画書素案 v1.0
 *
 * 金ヶ崎町方式踏襲：全8章構成、第9期体系継承＋認知症基本法対応の独立章、
 * 章カラー識別、図表・KPI表、根拠データ接続、仮置き値の明示
 *
 * アンケート未回収のため、調査結果由来部分は【アンケート結果反映後追記】で
 * 明示し、確定数値（実績データ確認サマリー由来）を最大限活用する。
 */
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, convertInchesToTwip,
  HeadingLevel, LevelFormat
} = require('docx');

// ===== カラー定義（章カラー方式：8章で識別） =====
const C = {
  // ベース
  navy:    "1F3864",
  blue:    "2F5597",
  lblue:   "DAE3F3",
  pale:    "EBF3FB",
  // アクセント
  orange:  "ED7D31",
  lorange: "FCE4D6",
  green:   "548235",
  lgreen:  "E2EFDA",
  red:     "C00000",
  // ベース・テキスト
  gray:    "595959",
  lgray:   "F2F2F2",
  black:   "000000",
  white:   "FFFFFF",
  yellow:  "FFE699",
};

// 章カラー（章扉・章ナビ・節見出しの差し色）
// 各章の色相を十分に離して識別性を確保
const CH = {
  1: { main: "1F3864", sub: "DAE3F3", name: "計画の策定にあたって" },         // ダークネイビー
  2: { main: "375623", sub: "E2EFDA", name: "川崎町の高齢者を取り巻く現状" }, // ダークグリーン
  3: { main: "2E75B6", sub: "DAE3F3", name: "第9期計画の取組実績と評価" },   // ミドルブルー
  4: { main: "C55A11", sub: "FCE4D6", name: "計画の基本理念と基本目標" },   // バーンオレンジ
  5: { main: "7030A0", sub: "E4D6F0", name: "施策の展開" },                 // パープル
  6: { main: "9333B0", sub: "EAD5F0", name: "認知症施策推進計画" },         // マゼンタ
  7: { main: "0F4F73", sub: "D6E4F0", name: "介護保険サービス見込量と保険料" }, // ダークティール（第1章ネイビーと差別化）
  8: { main: "404040", sub: "E2E8F0", name: "計画の推進体制と評価" },       // チャコール（旧グレーから濃度UP）
};

const FONT = "游ゴシック";

// 章コンテキスト：chapterTitle(N) 呼び出し時に記憶され、後続のsection/subsectionで暗黙的に使用される
// これにより、見出しが追加されても章番号を間違える事故が起きにくい
let _currentChapter = null;
function getCurrentChapter() { return _currentChapter; }
function setCurrentChapter(n) { _currentChapter = n; }

// ===== ヘルパー関数 =====

function text(t, opts = {}) {
  return new TextRun({
    text: t,
    font: FONT,
    size: opts.size || 21,
    bold: opts.bold || false,
    italics: opts.italics || false,
    color: opts.color || C.black,
    underline: opts.underline || undefined,
  });
}

function p(t, opts = {}) {
  return new Paragraph({
    spacing: opts.spacing || { before: 50, after: 80, line: 320 },
    alignment: opts.alignment || AlignmentType.LEFT,
    indent: opts.indent || (opts.noIndent ? undefined : { firstLine: 200 }),
    children: [text(t, opts)],
  });
}

// 章扉（ページブレーク付き・大きな章番号と章名）
// 章番号を内部状態に記憶し、後続のsection/subsectionで暗黙的に使えるようにする
function chapterTitle(num) {
  _currentChapter = num;  // 章番号を記憶
  const c = CH[num];
  return [
    new Paragraph({
      spacing: { before: 0, after: 0, line: 240 },
      pageBreakBefore: true,
      children: [text("", { size: 20 })],
    }),
    new Paragraph({
      spacing: { before: 1200, after: 200, line: 320 },
      alignment: AlignmentType.CENTER,
      children: [text(`第 ${num} 章`, { size: 48, bold: true, color: c.main })],
    }),
    new Paragraph({
      spacing: { before: 0, after: 240, line: 320 },
      alignment: AlignmentType.CENTER,
      border: {
        top: { style: BorderStyle.SINGLE, size: 24, color: c.main },
        bottom: { style: BorderStyle.SINGLE, size: 24, color: c.main },
      },
      shading: { type: ShadingType.SOLID, fill: c.sub },
      children: [text(c.name, { size: 32, bold: true, color: c.main })],
    }),
    new Paragraph({ spacing: { before: 200, after: 200 }, children: [text("")] }),
  ];
}

// 節見出し（第N-M節）
// 章番号は引数省略時に内部記憶の値を使う（呼び出し側の事故防止）
// タイトル本文も章カラーで統一（番号と本文の色を一致させ視覚的連動感を高める）
function section(chNum, secNum, title) {
  // chNum が文字列なら章番号省略呼び出し（section(secNum, title)）と解釈
  if (typeof chNum === 'string') {
    title = secNum;
    secNum = chNum;
    chNum = _currentChapter;
  }
  if (chNum == null) chNum = _currentChapter || 1;
  const c = CH[chNum];
  return new Paragraph({
    spacing: { before: 360, after: 160, line: 320 },
    pageBreakBefore: false,
    border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: c.main } },
    children: [
      text(`${chNum}-${secNum}　`, { size: 26, bold: true, color: c.main }),
      text(title, { size: 26, bold: true, color: c.main }),  // 本文も章カラーで統一
    ],
  });
}

// 項見出し（小見出し）
// 章番号は引数省略時に内部記憶の値を使う
function subsection(title, chNum) {
  if (chNum == null) chNum = _currentChapter || 1;
  const c = CH[chNum] || CH[1];
  return new Paragraph({
    spacing: { before: 220, after: 120, line: 300 },
    shading: { type: ShadingType.SOLID, fill: c.sub },
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 4, color: c.main },
    },
    children: [text("　▌ " + title, { size: 22, bold: true, color: c.main })],
  });
}

// 箇条書き
function bullet(t, opts = {}) {
  return new Paragraph({
    spacing: { before: 30, after: 30, line: 300 },
    indent: { left: convertInchesToTwip(opts.lv ? 0.5 : 0.3) },
    children: [text("・" + t, { size: 20 })],
  });
}

// 番号付き
function numItem(num, t) {
  return new Paragraph({
    spacing: { before: 40, after: 40, line: 300 },
    indent: { left: convertInchesToTwip(0.3) },
    children: [
      text(`${num}　`, { size: 20, bold: true }),
      text(t, { size: 20 }),
    ],
  });
}

// プレースホルダー（仮置き・調査後追記）
function placeholder(t) {
  return new Paragraph({
    spacing: { before: 80, after: 80, line: 280 },
    indent: { left: convertInchesToTwip(0.2) },
    shading: { type: ShadingType.SOLID, fill: "FFF2CC" },
    border: { 
      top: { style: BorderStyle.SINGLE, size: 12, color: C.orange },
      bottom: { style: BorderStyle.SINGLE, size: 12, color: C.orange },
    },
    children: [text("⚠ " + t, { size: 19, color: C.orange, italics: true, bold: true })],
  });
}

// 重要事実ボックス
function fact(t) {
  return new Paragraph({
    spacing: { before: 80, after: 80, line: 280 },
    indent: { left: convertInchesToTwip(0.2) },
    shading: { type: ShadingType.SOLID, fill: C.lgreen },
    border: {
      top: { style: BorderStyle.SINGLE, size: 6, color: C.green },
      bottom: { style: BorderStyle.SINGLE, size: 12, color: C.green },
    },
    children: [text("● " + t, { size: 19, color: C.green, bold: true })],
  });
}

// 出典注記
function source(t) {
  return new Paragraph({
    spacing: { before: 30, after: 80, line: 260 },
    alignment: AlignmentType.RIGHT,
    children: [text("出典：" + t, { size: 16, italics: true, color: C.gray })],
  });
}

// セル
function tcell(t, opts = {}) {
  return new TableCell({
    children: [new Paragraph({
      spacing: { before: 40, after: 40, line: 260 },
      alignment: opts.align || AlignmentType.LEFT,
      children: [text(t, {
        size: opts.size || 18,
        bold: opts.bold || false,
        color: opts.color || C.black,
      })],
    })],
    shading: opts.fill ? { type: ShadingType.SOLID, fill: opts.fill } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    width: opts.width ? { size: opts.width, type: WidthType.PERCENTAGE } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
  });
}

// ヘッダセル
function thcell(t, opts = {}) {
  return tcell(t, {
    ...opts,
    bold: true,
    color: C.white,
    fill: opts.fill || C.navy,
    align: AlignmentType.CENTER,
  });
}

// シンプルな2列キーバリュー表
function kvTable(rows, widths = [25, 75]) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: rows.map(([k, v]) => new TableRow({
      children: [
        tcell(k, { bold: true, fill: C.lblue, width: widths[0], color: C.navy }),
        tcell(v, { width: widths[1] }),
      ],
    })),
  });
}

// 空段落（スペーサー）
function spacer() {
  return new Paragraph({ spacing: { before: 80, after: 80 }, children: [text("")] });
}

// 画像挿入（PNG）：matplotlibグラフ等の埋め込み
const { ImageRun } = require('docx');
const fsImg = require('fs');
function image(path, opts = {}) {
  const data = fsImg.readFileSync(path);
  return new Paragraph({
    spacing: { before: 120, after: 60 },
    alignment: AlignmentType.CENTER,
    children: [
      new ImageRun({
        data: data,
        transformation: {
          width: opts.width || 480,
          height: opts.height || 300,
        },
      }),
    ],
  });
}

// 図のキャプション
function caption(t) {
  return new Paragraph({
    spacing: { before: 0, after: 200, line: 280 },
    alignment: AlignmentType.CENTER,
    children: [text(t, { size: 18, bold: true, color: C.navy })],
  });
}

module.exports = {
  C, CH, FONT,
  text, p, chapterTitle, section, subsection,
  bullet, numItem, placeholder, fact, source,
  tcell, thcell, kvTable, spacer, image, caption,
  getCurrentChapter, setCurrentChapter,
  // re-exports for build script
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, convertInchesToTwip,
};
