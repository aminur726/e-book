#!/usr/bin/env node
/**
 * ═══════════════════════════════════════════════════════════════
 *  অদম্য প্রেস — Professional DOCX Generator
 *  
 *  Generates Word documents that match the PDF ebook design.
 *  Supports V2 (Navy-Gold) and V4 (B&W Clean) templates.
 *
 *  Usage: node generate_docx.js input.json output.docx [v2|v4]
 *
 *  input.json = {
 *    info: { title_bn, title_en, author, tagline_bn, ... },
 *    pages: [ { page_num, title_bn, title_en, quote, ... } ],
 *    header_left, header_right, footer_text,
 *    watermark_text
 *  }
 * ═══════════════════════════════════════════════════════════════
 */
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
  BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageNumber, PageBreak, TabStopPosition, TabStopType,
  ImageRun, Tab
} = require("docx");

// ═══════════════════════════════════════════════════════════════
// COLOR THEMES
// ═══════════════════════════════════════════════════════════════
const THEMES = {
  v2: {
    name: "Navy Gold Elegant",
    primary:    "0B1D3A",  // deep navy
    accent:     "C9A84C",  // gold
    accent2:    "DFC06E",  // light gold
    text:       "111111",
    textLight:  "3D3D5C",
    textMuted:  "888888",
    textFaint:  "BBBBBB",
    quoteBg:    "FDF8E8",  // cream
    quoteBorder:"C9A84C",
    sepColor:   "DFC06E",
    coverBg:    "0B1D3A",
    coverText:  "FFFFFF",
    coverAccent:"C9A84C",
  },
  v4: {
    name: "B&W Clean Book",
    primary:    "000000",
    accent:     "333333",
    accent2:    "555555",
    text:       "111111",
    textLight:  "1A1A1A",
    textMuted:  "666666",
    textFaint:  "BBBBBB",
    quoteBg:    "F5F5F5",
    quoteBorder:"888888",
    sepColor:   "999999",
    coverBg:    "111111",
    coverText:  "FFFFFF",
    coverAccent:"CCCCCC",
  }
};

// Bangla numerals
const BN = ["১","২","৩","৪","৫","৬","৭","৮","৯","১০"];

// ═══════════════════════════════════════════════════════════════
// A5 page dimensions in DXA (1 inch = 1440 DXA)
// A5 = 148mm × 210mm = 5.827" × 8.268"
// ═══════════════════════════════════════════════════════════════
const A5_W = 8391;   // 148mm in DXA
const A5_H = 11906;  // 210mm in DXA
const MARGIN_TOP    = 680;   // ~12mm
const MARGIN_BOTTOM = 624;   // ~11mm
const MARGIN_LR     = 737;   // ~13mm
const CONTENT_W = A5_W - MARGIN_LR * 2; // usable width

// ═══════════════════════════════════════════════════════════════
// HELPER: Create a thin horizontal rule using a single-row table
// ═══════════════════════════════════════════════════════════════
function horizontalRule(color, thickness = 4) {
  const border = { style: BorderStyle.SINGLE, size: thickness, color: color };
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    borders: {
      top: border,
      bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: {
              top: border,
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE },
            },
            width: { size: CONTENT_W, type: WidthType.DXA },
            children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [] })],
          }),
        ],
      }),
    ],
  });
}

// ═══════════════════════════════════════════════════════════════
// HELPER: Quote block — shaded table cell with left border
// ═══════════════════════════════════════════════════════════════
function quoteBlock(text, author, theme) {
  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: author ? 40 : 60 },
      children: [
        new TextRun({
          text: `"${text}"`,
          italics: true,
          font: "Noto Serif Bengali",
          size: 17, // 8.5pt
          color: theme.primary,
        }),
      ],
    }),
  ];

  if (author) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [
          new TextRun({
            text: `— ${author}`,
            font: "Noto Sans Bengali",
            size: 14, // 7pt
            color: theme.textMuted,
            italics: true,
          }),
        ],
      })
    );
  }

  const leftBorder = {
    style: BorderStyle.SINGLE,
    size: 12, // thick left border
    color: theme.quoteBorder,
  };
  const noBorder = { style: BorderStyle.NONE };

  return new Table({
    width: { size: CONTENT_W - 400, type: WidthType.DXA },
    columnWidths: [CONTENT_W - 400],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: {
              top: noBorder,
              bottom: noBorder,
              left: leftBorder,
              right: noBorder,
            },
            shading: { fill: theme.quoteBg, type: ShadingType.CLEAR },
            width: { size: CONTENT_W - 400, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 200, right: 160 },
            children: children,
          }),
        ],
      }),
    ],
  });
}

// ═══════════════════════════════════════════════════════════════
// BUILD COVER SECTION
// ═══════════════════════════════════════════════════════════════
function buildCover(info, theme) {
  // Cover uses a bordered table filling the page
  const noBorder = { style: BorderStyle.NONE };
  const goldBorder = {
    style: BorderStyle.SINGLE,
    size: 8,
    color: theme.coverAccent,
  };

  const coverContent = [
    // Top spacer
    new Paragraph({ spacing: { before: 2400, after: 0 }, children: [] }),

    // Title BN
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 80 },
      children: [
        new TextRun({
          text: info.title_bn || "",
          bold: true,
          font: "Noto Sans Bengali",
          size: 44, // 22pt
          color: theme.coverAccent,
        }),
      ],
    }),

    // Title EN
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120 },
      children: [
        new TextRun({
          text: info.title_en || "",
          italics: true,
          font: "Noto Serif Bengali",
          size: 20, // 10pt
          color: theme.accent2,
        }),
      ],
    }),

    // Gold rule
    horizontalRule(theme.coverAccent, 4),

    // Tagline BN
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 40 },
      children: [
        new TextRun({
          text: info.tagline_bn || "",
          font: "Noto Sans Bengali",
          size: 17,
          color: theme.coverText,
        }),
      ],
    }),

    // Tagline EN
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
      children: [
        new TextRun({
          text: info.tagline_en || "",
          font: "Noto Sans Bengali",
          size: 14,
          color: theme.textMuted,
        }),
      ],
    }),

    // Subtitle
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 400 },
      children: [
        new TextRun({
          text: info.subtitle_bn || "",
          font: "Noto Sans Bengali",
          size: 14,
          color: theme.textMuted,
        }),
      ],
    }),

    // Bottom spacer
    new Paragraph({ spacing: { before: 1200, after: 0 }, children: [] }),

    // Author
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 160 },
      children: [
        new TextRun({
          text: info.author || "",
          bold: true,
          font: "Noto Sans Bengali",
          size: 16,
          color: theme.coverText,
        }),
      ],
    }),

    // Press BN
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 20 },
      children: [
        new TextRun({
          text: "অদম্য প্রেস",
          bold: true,
          font: "Noto Sans Bengali",
          size: 18,
          color: theme.coverAccent,
        }),
      ],
    }),

    // Press EN
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0 },
      children: [
        new TextRun({
          text: "Odommo Press",
          font: "Noto Sans Bengali",
          size: 11,
          color: theme.textMuted,
        }),
      ],
    }),
  ];

  return {
    properties: {
      page: {
        size: { width: A5_W, height: A5_H },
        margin: { top: 340, bottom: 340, left: 340, right: 340 }, // tight margins for cover
        borders: {
          pageBorderTop: goldBorder,
          pageBorderBottom: goldBorder,
          pageBorderLeft: goldBorder,
          pageBorderRight: goldBorder,
        },
      },
    },
    children: coverContent,
  };
}

// ═══════════════════════════════════════════════════════════════
// BUILD CHAPTER PAGES SECTION
// ═══════════════════════════════════════════════════════════════
function buildChapters(pages, info, settings, theme) {
  const children = [];

  pages.forEach((pg, idx) => {
    // Page break before each chapter (except first — it follows cover section)
    if (idx > 0) {
      children.push(
        new Paragraph({
          children: [new PageBreak()],
        })
      );
    }

    // ── Page number (right-aligned) ──
    children.push(
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 0, after: 100 },
        children: [
          new TextRun({
            text: `পৃষ্ঠা ${pg.page_num}`,
            font: "Noto Sans Bengali",
            size: 14,
            color: theme.textMuted,
          }),
        ],
      })
    );

    // ── Title (Bangla) ──
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 80, after: 40 },
        children: [
          new TextRun({
            text: pg.title_bn || "",
            bold: true,
            font: "Noto Sans Bengali",
            size: 26, // 13pt
            color: theme.primary,
          }),
        ],
      })
    );

    // ── Title (English subtitle) ──
    if (pg.title_en) {
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 80 },
          children: [
            new TextRun({
              text: pg.title_en,
              italics: true,
              font: "Noto Serif Bengali",
              size: 16, // 8pt
              color: theme.textMuted,
            }),
          ],
        })
      );
    }

    // ── Decorative rule ──
    children.push(horizontalRule(theme.accent, 4));

    // ── Diamond separator (V2 only) ──
    if (theme === THEMES.v2) {
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 40, after: 60 },
          children: [
            new TextRun({
              text: "◆",
              font: "Noto Sans Bengali",
              size: 10,
              color: theme.accent,
            }),
          ],
        })
      );
    }

    // ── Quote block ──
    if (pg.quote) {
      children.push(quoteBlock(pg.quote, pg.quote_author, theme));
    }

    // ── Intro paragraph ──
    if (pg.intro) {
      children.push(
        new Paragraph({
          alignment: AlignmentType.JUSTIFY,
          spacing: { before: 100, after: 60 },
          children: [
            new TextRun({
              text: pg.intro,
              font: "Noto Sans Bengali",
              size: 16, // 8pt
              color: theme.text,
            }),
          ],
        })
      );
    }

    // ── Thin separator before tips ──
    children.push(horizontalRule(theme.sepColor, 2));

    // ── 10 Tips ──
    const tips = pg.tips || [];
    tips.slice(0, 10).forEach((tip, i) => {
      children.push(
        new Paragraph({
          alignment: AlignmentType.JUSTIFY,
          spacing: { before: 50, after: 50 },
          indent: { left: 360, hanging: 360 }, // hanging indent for number
          children: [
            new TextRun({
              text: `${BN[i]}.  `,
              bold: true,
              font: "Noto Sans Bengali",
              size: 16,
              color: theme.primary,
            }),
            new TextRun({
              text: tip,
              font: "Noto Sans Bengali",
              size: 16,
              color: theme.textLight,
            }),
          ],
        })
      );
    });

    // ── Closing ──
    if (pg.closing) {
      children.push(horizontalRule(theme.sepColor, 2));
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 80, after: 60 },
          children: [
            new TextRun({
              text: pg.closing,
              italics: true,
              font: "Noto Serif Bengali",
              size: 15,
              color: theme.textMuted,
            }),
          ],
        })
      );
    }

    // ── Footer line (book title + press) ──
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 0 },
        children: [
          new TextRun({
            text: settings.footer_text || `${info.title_bn}  •  অদম্য প্রেস`,
            font: "Noto Sans Bengali",
            size: 10,
            color: theme.textFaint,
          }),
        ],
      })
    );
  });

  return {
    properties: {
      page: {
        size: { width: A5_W, height: A5_H },
        margin: {
          top: MARGIN_TOP,
          bottom: MARGIN_BOTTOM,
          left: MARGIN_LR,
          right: MARGIN_LR,
        },
      },
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            spacing: { after: 40 },
            tabStops: [
              { type: TabStopType.RIGHT, position: CONTENT_W },
            ],
            children: [
              new TextRun({
                text: settings.header_left || info.title_bn || "",
                font: "Noto Sans Bengali",
                size: 12,
                color: theme.textMuted,
              }),
              new TextRun({ children: [new Tab()] }),
              new TextRun({
                text: settings.header_right || "অদম্য প্রেস",
                font: "Noto Sans Bengali",
                size: 12,
                color: theme.textMuted,
              }),
            ],
          }),
        ],
      }),
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: settings.footer_text || "অদম্য প্রেস | Odommo Press",
                font: "Noto Sans Bengali",
                size: 11,
                color: theme.textFaint,
              }),
            ],
          }),
        ],
      }),
    },
    children: children,
  };
}

// ═══════════════════════════════════════════════════════════════
// BUILD END PAGE SECTION
// ═══════════════════════════════════════════════════════════════
function buildEndPage(theme) {
  const goldBorder = {
    style: BorderStyle.SINGLE,
    size: 6,
    color: theme.coverAccent,
  };

  return {
    properties: {
      page: {
        size: { width: A5_W, height: A5_H },
        margin: { top: 500, bottom: 500, left: 500, right: 500 },
        borders: {
          pageBorderTop: goldBorder,
          pageBorderBottom: goldBorder,
          pageBorderLeft: goldBorder,
          pageBorderRight: goldBorder,
        },
      },
    },
    children: [
      new Paragraph({ spacing: { before: 3600, after: 0 }, children: [] }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 240 },
        children: [
          new TextRun({
            text: "— সমাপ্ত —",
            bold: true,
            font: "Noto Sans Bengali",
            size: 32,
            color: theme.coverAccent,
          }),
        ],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 40 },
        children: [
          new TextRun({
            text: "অদম্য প্রেস",
            bold: true,
            font: "Noto Sans Bengali",
            size: 20,
            color: theme.accent2,
          }),
        ],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: "Odommo Press",
            font: "Noto Sans Bengali",
            size: 12,
            color: theme.textMuted,
          }),
        ],
      }),
    ],
  };
}

// ═══════════════════════════════════════════════════════════════
// MAIN: Read input → Build document → Write output
// ═══════════════════════════════════════════════════════════════
async function main() {
  const args = process.argv.slice(2);
  if (args.length < 2) {
    console.error("Usage: node generate_docx.js input.json output.docx [v2|v4]");
    process.exit(1);
  }

  const inputPath = args[0];
  const outputPath = args[1];
  const designKey = (args[2] || "v2").toLowerCase();
  const theme = THEMES[designKey] || THEMES.v2;

  // Read input data
  const data = JSON.parse(fs.readFileSync(inputPath, "utf8"));
  const { info, pages, header_left, header_right, footer_text, watermark_text } = data;

  const settings = {
    header_left: header_left || "",
    header_right: header_right || "",
    footer_text: footer_text || "",
    watermark_text: watermark_text || "",
  };

  console.error(`[docx-gen] Design: ${theme.name} | Pages: ${pages.length}`);

  // Build document with 3 sections: cover, chapters, end page
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: "Noto Sans Bengali",
            size: 16, // 8pt default
          },
        },
      },
    },
    sections: [
      buildCover(info, theme),
      buildChapters(pages, info, settings, theme),
      buildEndPage(theme),
    ],
  });

  // Pack and write
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  console.error(`[docx-gen] Written: ${outputPath} (${(buffer.length/1024).toFixed(0)} KB)`);
}

main().catch((err) => {
  console.error(`[docx-gen] Error: ${err.message}`);
  process.exit(1);
});
