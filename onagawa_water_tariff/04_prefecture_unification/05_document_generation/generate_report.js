/**
 * 統合最終報告書 Word生成スクリプト
 * ==================================
 *
 * 本スクリプトは、統合最終報告書(宮城県水道料金体系統一化_検討報告書.docx)を
 * docxライブラリで生成する骨格コードです。
 *
 * 実行方法:
 *   npm install docx
 *   node generate_report.js
 *
 * 【出力構成】
 * - 表紙
 * - 目次
 * - エグゼクティブサマリー
 * - 第1章 検討の背景と目的
 * - 第2章 検討の進め方と方法論
 * - 第3章 県内臨海4団体の料金体系(確定版)
 * - 第4章 料金水準の比較と構造分析(全国・県内位置づけ含む)
 * - 第5章 決定的発見 ― 塩竈市『生産用水用』区分
 * - 第6章 水産加工業への影響定量化
 * - 第7章 政策選択肢の評価
 * - 第8章 推奨政策パッケージ
 * - 第9章 実施ロードマップ
 * - 第10章 残された課題と次のアクション
 * - 参考資料・出典一覧
 *
 * 【必要な事前準備】
 * 04_chart_generation/generate_charts.py を実行して
 * 04_chart_generation/charts/ に PNG画像を生成しておくこと。
 */

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType,
  ShadingType, ImageRun, PageNumber, Header, Footer, PageBreak
} = require('docx');

// -----------------------------------------------------------
// ヘルパー関数群
// -----------------------------------------------------------

const border = { style: BorderStyle.SINGLE, size: 4, color: "999999" };
const borders = { top: border, bottom: border, left: border, right: border };

/** ヘッダーセル生成 */
function headerCell(text, width, opts = {}) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: opts.shade || "1F4E78", type: ShadingType.CLEAR },
    margins: { top: 90, bottom: 90, left: 90, right: 90 },
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({
        text, bold: true, color: "FFFFFF",
        font: "游ゴシック", size: 17
      })]
    })]
  });
}

/** 本文セル生成 */
function bodyCell(text, width, opts = {}) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: opts.shade ? { fill: opts.shade, type: ShadingType.CLEAR } : undefined,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({
        text, bold: opts.bold || false,
        font: "游ゴシック", size: 17,
        color: opts.color || "000000"
      })]
    })]
  });
}

/** 見出し生成 */
function heading(text, level) {
  return new Paragraph({
    heading: level,
    children: [new TextRun({ text, font: "游ゴシック" })]
  });
}

/** 通常段落生成 */
function para(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60, line: 340 },
    alignment: opts.align || AlignmentType.LEFT,
    children: [new TextRun({
      text, font: "游明朝",
      size: opts.size || 22,
      bold: opts.bold || false,
      color: opts.color || "000000"
    })]
  });
}

/** 箇条書き生成 */
function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    spacing: { before: 40, after: 40, line: 320 },
    children: [new TextRun({ text, font: "游明朝", size: 22 })]
  });
}

function blank() { return new Paragraph({ children: [new TextRun({ text: "" })] }); }
function pageBreak() { return new Paragraph({ children: [new PageBreak()] }); }

/** 画像挿入 */
function image(imgPath, width, height) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new ImageRun({
      data: fs.readFileSync(imgPath),
      transformation: { width, height },
      type: "png"
    })]
  });
}

/** 図のキャプション */
function caption(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 40, after: 200 },
    children: [new TextRun({
      text, font: "游ゴシック", size: 18,
      italics: true, color: "555555"
    })]
  });
}

/** 強調ボックス(重要な発見用) */
function highlightBox(title, content) {
  const boxBorder = { style: BorderStyle.SINGLE, size: 12, color: "C00000" };
  const innerBorders = { top: boxBorder, bottom: boxBorder, left: boxBorder, right: boxBorder };
  return [
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [9360],
      rows: [new TableRow({
        children: [new TableCell({
          borders: innerBorders,
          width: { size: 9360, type: WidthType.DXA },
          shading: { fill: "FFF2F2", type: ShadingType.CLEAR },
          margins: { top: 200, bottom: 200, left: 240, right: 240 },
          children: [
            new Paragraph({
              spacing: { after: 100 },
              children: [new TextRun({
                text: title, bold: true, font: "游ゴシック",
                size: 24, color: "C00000"
              })]
            }),
            new Paragraph({
              spacing: { line: 340 },
              children: [new TextRun({ text: content, font: "游明朝", size: 22 })]
            })
          ]
        })]
      })]
    })
  ];
}

// -----------------------------------------------------------
// 本文コンテンツ(サンプル: 表紙 + エグゼクティブサマリー)
// 実際の統合報告書はこの構造で全10章を記述
// -----------------------------------------------------------

const chartsDir = path.join(__dirname, '..', '04_chart_generation', 'charts');

const children = [
  // ==== 表紙 ====
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 2400, after: 200 },
    children: [new TextRun({
      text: "宮城県水道料金体系統一化",
      bold: true, size: 40, font: "游ゴシック"
    })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    children: [new TextRun({
      text: "に関する検討報告書",
      bold: true, size: 40, font: "游ゴシック"
    })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 400, after: 600 },
    children: [new TextRun({
      text: "― 県内臨海4団体の料金体系分析と水産加工業への影響を踏まえた政策提言 ―",
      size: 22, font: "游ゴシック", color: "555555"
    })]
  }),

  pageBreak(),

  // ==== エグゼクティブサマリー ====
  heading("エグゼクティブサマリー", HeadingLevel.HEADING_1),
  para("宮城県内の水道料金体系統一化に向けて、県内臨海水産加工業を主要産業とする4団体(女川町・気仙沼市・石巻地方広域水道企業団・塩竈市)の料金体系を詳細に分析した。"),
  blank(),

  para("主要な発見", { bold: true, size: 24 }),
  bullet("女川町の現行料金は他3団体の約1/3水準"),
  bullet("単純統一は女川町水産加工業に+190%超の壊滅的負担増"),
  bullet("塩竈市には既に『生産用水用』(105円/㎥)という産業用優遇区分がある"),
  bullet("塩竈モデルの全県導入により女川町影響を+3%に抑制可能"),
  blank(),

  // グラフ挿入例
  image(path.join(chartsDir, 'report_chart3_scenarios.png'), 580, 295),
  caption("図 政策シナリオ別の影響比較"),

  // ==== 以下、第1章~第10章の内容を続けて記述 ====
  // (実際の統合報告書では579段落、25ページに及ぶ)
];

// -----------------------------------------------------------
// ドキュメント設定
// -----------------------------------------------------------

const doc = new Document({
  creator: "宮城県水道料金検討",
  title: "宮城県水道料金体系統一化に関する検討報告書",
  styles: {
    default: { document: { run: { font: "游明朝", size: 22 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal",
        next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, font: "游ゴシック", color: "1F4E78" },
        paragraph: {
          spacing: { before: 360, after: 240 },
          outlineLevel: 0,
          border: {
            bottom: { style: BorderStyle.SINGLE, size: 18, color: "1F4E78", space: 4 }
          }
        }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal",
        next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "游ゴシック", color: "2E75B6" },
        paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 }
      }
    ]
  },
  numbering: {
    config: [{
      reference: "bullets",
      levels: [
        {
          level: 0, format: LevelFormat.BULLET, text: "●",
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: { indent: { left: 540, hanging: 270 } },
            run: { font: "游明朝", size: 22 }
          }
        }
      ]
    }]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },  // A4
        margin: { top: 1440, right: 1134, bottom: 1440, left: 1134 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({
            text: "宮城県水道料金体系統一化 検討報告書",
            size: 18, font: "游ゴシック", color: "888888"
          })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            children: ["- ", PageNumber.CURRENT, " -"],
            size: 20, font: "游ゴシック"
          })]
        })]
      })
    },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  const outputPath = path.join(__dirname, '宮城県水道料金体系統一化_検討報告書_sample.docx');
  fs.writeFileSync(outputPath, buffer);
  console.log(`Document created: ${outputPath}`);
});
