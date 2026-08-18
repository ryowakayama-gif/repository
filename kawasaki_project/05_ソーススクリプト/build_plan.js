/**
 * 川崎町第10期計画書素案 v1.0 - 統合ビルダー
 */
const fs = require('fs');
const H = require('./plan_helpers');
const {
  C, CH, FONT,
  Document, Packer, Paragraph, TextRun,
  Header, Footer, AlignmentType,
  PageNumber, BorderStyle,
} = H;

// 各章を読み込む
const part1 = require('./plan_part1');  // 表紙〜目次〜第1-4章
const part2 = require('./plan_part2');  // 第5-6章
const part7 = require('./plan_part7');  // 第7章（詳細化版）
const part8 = require('./plan_part8');  // 第8章

const allSections = [...part1.S, ...part2.S, ...part7.S, ...part8.S];

console.log("Total blocks/paragraphs/tables:", allSections.length);

const doc = new Document({
  creator: "ビズアップ公共コンサルティング株式会社",
  title: "川崎町第10期介護保険事業計画 計画書素案v1.0",
  description: "川崎町高齢者保健福祉計画・第10期介護保険事業計画 計画書素案",
  styles: {
    default: {
      document: { run: { font: FONT, size: 21 } },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },  // A4
        margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({
            text: "川崎町高齢者保健福祉計画・第10期介護保険事業計画　計画書素案v1.0",
            font: FONT, size: 16, color: C.gray,
          })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "─ ", font: FONT, size: 16, color: C.gray }),
            new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: C.gray }),
            new TextRun({ text: " ─", font: FONT, size: 16, color: C.gray }),
          ],
        })],
      }),
    },
    children: allSections,
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/claude/kawasaki_work/川崎町_計画書素案_v1.0.docx", buffer);
  console.log("Build done: 川崎町_計画書素案_v1.0.docx");
}).catch(err => {
  console.error("Build error:", err);
  process.exit(1);
});
