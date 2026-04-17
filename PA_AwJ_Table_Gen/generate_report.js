const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  BorderStyle, WidthType, ShadingType, UnderlineType, VerticalAlign
} = require('docx');
const fs = require('fs');

const data = JSON.parse(process.argv[2]);

// COL1: 2.9cm (fits "Explanation" on one line at 11pt), COL2: 13.0cm, total 15.9cm
const COL1 = Math.round(2.21 * 567);   // 1644 DXA
const COL2 = Math.round(13.69 * 567);  // 7371 DXA

// Table Grid: single solid line, Auto color, 0.5pt = size 4 (half-points)
const gridBorder = { style: BorderStyle.SINGLE, size: 4, color: "auto" };
const gridBorders = {
  top: gridBorder,
  bottom: gridBorder,
  left: gridBorder,
  right: gridBorder,
  insideHorizontal: gridBorder,
  insideVertical: gridBorder
};

// Table Grid paragraph style: single line spacing, 0pt space after
const cellParaSpacing = {
  line: 240,        // single = 240
  lineRule: "auto",
  after: 0
};

// Title paragraph: 1.08 line spacing, 8pt after (document default)
const titleParaSpacing = {
  line: 259,
  lineRule: "auto",
  after: 160
};

function makeRun(text, opts = {}) {
  return new TextRun({
    text,
    font: { name: "Calibri", eastAsia: "SimSun" },
    size: 22,
    ...opts
  });
}

function makeEntryBlock(title, explanation) {
  const titleParagraph = new Paragraph({
    spacing: { before: 0, after: 0, line: 259, lineRule: "auto" },
    children: [
      makeRun(title, { bold: true, underline: { type: UnderlineType.SINGLE } })
    ]
  });

  const headerRow = new TableRow({
    children: [
      new TableCell({
        borders: gridBorders,
        columnSpan: 2,
        shading: { fill: "002060", type: ShadingType.CLEAR },
        margins: { top: 0, bottom: 0, left: 108, right: 108 },
        children: [new Paragraph({
          spacing: cellParaSpacing,
          children: [makeRun("Justification", { bold: true, color: "FFFFFF" })]
        })]
      })
    ]
  });

  const dataRow = new TableRow({
    children: [
      new TableCell({
        borders: gridBorders,
        width: { size: COL1, type: WidthType.DXA },
        verticalAlign: VerticalAlign.TOP,
        margins: { top: 60, bottom: 60, left: 108, right: 108 },
        children: [new Paragraph({
          spacing: cellParaSpacing,
          indent: { right: -114 }, // -0.2cm = -114 DXA
          children: [makeRun("Explanation")]
        })]
      }),
      new TableCell({
        borders: gridBorders,
        width: { size: COL2, type: WidthType.DXA },
        verticalAlign: VerticalAlign.TOP,
        margins: { top: 60, bottom: 60, left: 108, right: 108 },
        children: [
          new Paragraph({
            spacing: cellParaSpacing,
            children: [makeRun(explanation)]
          }),
          new Paragraph({ spacing: cellParaSpacing, children: [] })
        ]
      })
    ]
  });

  return [
    titleParagraph,
    new Table({
      style: "TableGrid",
      layout: "fixed",
      columnWidths: [COL1, COL2],
      rows: [headerRow, dataRow]
    }),
    new Paragraph({ spacing: { before: 0, after: 0, line: 240, lineRule: "auto" }, children: [] })
  ];
}

const children = [];
for (const row of data) {
  const title = row['Misconfiguration'] || '';
  const explanation = row['CSTP Justification'] || '';
  children.push(...makeEntryBlock(title, explanation));
}

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: { name: "Calibri", eastAsia: "SimSun" }, size: 22 },
        paragraph: { spacing: { line: 259, lineRule: "auto", after: 160 } }
      }
    },
    tableStyles: [
      {
        id: "TableGrid",
        name: "Table Grid",
        basedOn: "TableNormal",
        run: { font: { name: "Calibri", eastAsia: "SimSun" }, size: 22 },
        paragraph: { spacing: { line: 240, lineRule: "auto", after: 0 } },
        table: {
          borders: {
            top:    { style: BorderStyle.SINGLE, size: 4, color: "auto" },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: "auto" },
            left:   { style: BorderStyle.SINGLE, size: 4, color: "auto" },
            right:  { style: BorderStyle.SINGLE, size: 4, color: "auto" },
            insideH:{ style: BorderStyle.SINGLE, size: 4, color: "auto" },
            insideV:{ style: BorderStyle.SINGLE, size: 4, color: "auto" }
          }
        }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  const outPath = process.argv[3] || 'output.docx';
  fs.writeFileSync(outPath, buffer);
  console.log('Done: ' + outPath);
}).catch(err => {
  console.error(err);
  process.exit(1);
});