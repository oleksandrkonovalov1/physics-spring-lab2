#!/usr/bin/env node

import {
  Document, Packer, Paragraph, TextRun, Header, ImageRun,
  AlignmentType, PageNumber, HeadingLevel,
  PageBreak, BorderStyle, ShadingType, VerticalAlign,
  Table, TableRow, TableCell, WidthType,
  TabStopType,
  Math as OfficeMath, MathRun, MathFraction,
  MathSuperScript, MathSubScript, MathFunction, MathRoundBrackets,
} from "docx";
import { readFileSync, writeFileSync, existsSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const MM_TO_DXA = 56.693;
const PT_TO_HALF_PT = 2;
const FONT = "Times New Roman";
const BODY_SIZE = 14 * PT_TO_HALF_PT;
const TITLE_SIZE = 14 * PT_TO_HALF_PT;
const LINE_SPACING_15 = 360;
const FIRST_LINE_INDENT = Math.round(12.5 * MM_TO_DXA);

const margins = {
  top: Math.round(20 * MM_TO_DXA),
  bottom: Math.round(20 * MM_TO_DXA),
  left: Math.round(30 * MM_TO_DXA),
  right: Math.round(15 * MM_TO_DXA),
};

const PAGE_WIDTH_DXA = 11906 - margins.left - margins.right;

// ── Text Helpers ────────────────────────────────────────────────────────────

function titleRun(text, opts = {}) {
  return new TextRun({ text, font: FONT, size: TITLE_SIZE, bold: opts.bold ?? false, ...opts });
}

function bodyRun(text, opts = {}) {
  return new TextRun({ text, font: FONT, size: BODY_SIZE, ...opts });
}

function centeredParagraph(runs, spacing = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto", ...spacing },
    children: Array.isArray(runs) ? runs : [runs],
  });
}

function emptyLine() {
  return centeredParagraph(titleRun(""));
}

function bodyParagraph(text) {
  return new Paragraph({
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    indent: { firstLine: FIRST_LINE_INDENT },
    alignment: AlignmentType.JUSTIFIED,
    children: [bodyRun(text)],
  });
}

function sectionHeading(number, title) {
  const text = number ? `${number} ${title}`.toUpperCase() : title.toUpperCase();
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 120, line: LINE_SPACING_15, lineRule: "auto" },
    keepNext: true,
    children: [new TextRun({ text, font: FONT, size: BODY_SIZE, bold: true })],
  });
}

function subsectionHeading(number, title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 120, after: 60, line: LINE_SPACING_15, lineRule: "auto" },
    indent: { firstLine: FIRST_LINE_INDENT },
    keepNext: true,
    children: [new TextRun({ text: `${number} ${title}`, font: FONT, size: BODY_SIZE, bold: true })],
  });
}

function tableCaption(number, title) {
  return new Paragraph({
    spacing: { before: 120, after: 60, line: LINE_SPACING_15, lineRule: "auto" },
    indent: { firstLine: FIRST_LINE_INDENT },
    children: [bodyRun(`Таблиця ${number} — ${title}`)],
  });
}

function figureCaption(number, title) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 120, line: LINE_SPACING_15, lineRule: "auto" },
    children: [bodyRun(`Рисунок ${number} — ${title}`)],
  });
}

// ── Title Page ───────────────────────────────────────────────────────────────

const year = new Date().getFullYear();

function tabbedLine(leftText, rightText) {
  return new Paragraph({
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    tabStops: [{ type: TabStopType.RIGHT, position: PAGE_WIDTH_DXA }],
    children: [
      titleRun(leftText),
      new TextRun({ children: ["\t"], font: FONT, size: TITLE_SIZE }),
      titleRun(rightText),
    ],
  });
}

const titlePage = [
  centeredParagraph(titleRun("Міністерство освіти і науки України")),
  centeredParagraph(titleRun("Харківський національний університет радіоелектроніки")),
  emptyLine(),
  centeredParagraph(titleRun("Кафедра фізики")),
  emptyLine(), emptyLine(), emptyLine(), emptyLine(),
  centeredParagraph(titleRun("ЗВІТ", { bold: true })),
  centeredParagraph(titleRun("з лабораторної роботи № 2")),
  centeredParagraph(titleRun("з дисципліни «Фізика»")),
  centeredParagraph(titleRun("на тему: «Дослідження явища самоіндукції»")),
  emptyLine(),
  centeredParagraph(titleRun("Варіант 10")),
  emptyLine(), emptyLine(),
  tabbedLine("Виконав:", "Перевірив:"),
  tabbedLine("ст. гр. ПЗПІ-25-6", "Приймачов Ю. Д."),
  new Paragraph({
    spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
    children: [titleRun("Коновалов О. С.")],
  }),
  emptyLine(), emptyLine(), emptyLine(), emptyLine(), emptyLine(), emptyLine(),
  centeredParagraph(titleRun(`Харків — ${year}`)),
];

// ── Table Helpers ────────────────────────────────────────────────────────────

const TABLE_BORDERS = {
  top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
  bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
  left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
  right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
  insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
};

const CELL_MARGINS = { top: 40, bottom: 40, left: 80, right: 80 };

function headerCell(text) {
  return new TableCell({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
      children: [bodyRun(text, { bold: true })],
    })],
    verticalAlign: VerticalAlign.CENTER,
    margins: CELL_MARGINS,
  });
}

function dataCell(text, alignment = AlignmentType.CENTER) {
  return new TableCell({
    children: [new Paragraph({
      alignment,
      spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
      children: [bodyRun(text)],
    })],
    verticalAlign: VerticalAlign.CENTER,
    margins: CELL_MARGINS,
  });
}

// ── Experimental Data and Table Builders ─────────────────────────────────────

const S = 16e-4, N = 150, l = 20e-2, R = 50, U = 220;
const I0 = U / R;

const t_data = [0.0, 0.7, 1.5, 2.7, 3.6, 4.7, 5.8, 7.0, 8.2, 9.4, 10.6, 11.7, 12.9];
const I_data = [4.4, 3.7661, 3.1527, 2.4148, 1.977, 1.5483, 1.2125, 0.9287, 0.7113, 0.5448, 0.4173, 0.3268, 0.2503];

const ln_I = I_data.map(i => Math.log(i));
const L_values = t_data.slice(1).map((t, idx) => -R * t / Math.log(I_data[idx + 1] / I0));
const eps_values = I_data.map(i => R * i);
const L_mean = L_values.reduce((a, b) => a + b, 0) / L_values.length;

function buildInputTable() {
  return new Table({
    rows: [
      new TableRow({
        children: [headerCell("Варіант"), headerCell("S, ×10⁻⁴ м²"), headerCell("N"), headerCell("l, ×10⁻² м"), headerCell("R, Ом")],
        tableHeader: true,
      }),
      new TableRow({
        children: [dataCell("10"), dataCell("16"), dataCell("150"), dataCell("20"), dataCell("50")],
      }),
    ],
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: TABLE_BORDERS,
  });
}

function buildResultsTable() {
  const header = new TableRow({
    children: [headerCell("t, с"), headerCell("I, А"), headerCell("ln I"), headerCell("L, Гн"), headerCell("ε, В")],
    tableHeader: true,
  });

  const rows = t_data.map((t, i) => {
    const L_str = i === 0 ? "—" : L_values[i - 1].toFixed(2);
    return new TableRow({
      children: [dataCell(t.toFixed(1)), dataCell(I_data[i].toFixed(4)), dataCell(ln_I[i].toFixed(4)), dataCell(L_str), dataCell(eps_values[i].toFixed(2))],
    });
  });

  const avgRow = new TableRow({
    children: [
      new TableCell({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" },
          children: [bodyRun("Середнє L", { bold: true })],
        })],
        columnSpan: 3,
        verticalAlign: VerticalAlign.CENTER,
        margins: CELL_MARGINS,
      }),
      dataCell(`${L_mean.toFixed(2)} Гн`),
      dataCell(""),
    ],
  });

  return new Table({
    rows: [header, ...rows, avgRow],
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: TABLE_BORDERS,
  });
}

// ── OMML Formula Helpers ─────────────────────────────────────────────────────

function numberedFormula(mathChildren, number) {
  const centerPos = Math.round(PAGE_WIDTH_DXA / 2);
  return new Paragraph({
    spacing: { before: 120, after: 120, line: LINE_SPACING_15, lineRule: "auto" },
    tabStops: [
      { type: TabStopType.CENTER, position: centerPos },
      { type: TabStopType.RIGHT, position: PAGE_WIDTH_DXA },
    ],
    children: [
      new TextRun({ children: ["\t"], font: FONT, size: BODY_SIZE }),
      new OfficeMath({ children: mathChildren }),
      new TextRun({ children: ["\t"], font: FONT, size: BODY_SIZE }),
      bodyRun(`(${number})`),
    ],
  });
}

function unnumberedFormula(mathChildren) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 60, line: LINE_SPACING_15, lineRule: "auto" },
    children: [new OfficeMath({ children: mathChildren })],
  });
}

// Formula (1): I = I₀ · e^(-Rt/L)
function formulaCurrentDecay() {
  return numberedFormula([
    new MathRun("I = "),
    new MathSubScript({ children: [new MathRun("I")], subScript: [new MathRun("0")] }),
    new MathRun(" · "),
    new MathSuperScript({
      children: [new MathRun("e")],
      superScript: [
        new MathRun("−"),
        new MathFraction({ numerator: [new MathRun("R·t")], denominator: [new MathRun("L")] }),
      ],
    }),
  ], 1);
}

// Formula (2): L = -R·t / ln(I/I₀)
function formulaInductance() {
  return numberedFormula([
    new MathRun("L = −"),
    new MathFraction({
      numerator: [new MathRun("R · t")],
      denominator: [
        new MathFunction({
          name: [new MathRun("ln")],
          children: [
            new MathRoundBrackets({
              children: [
                new MathFraction({
                  numerator: [new MathRun("I")],
                  denominator: [
                    new MathSubScript({ children: [new MathRun("I")], subScript: [new MathRun("0")] }),
                  ],
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  ], 2);
}

// Formula (3): ε_C = R · I₀ · e^(-Rt/L) = R · I(t)
function formulaEMF() {
  return numberedFormula([
    new MathSubScript({ children: [new MathRun("ε")], subScript: [new MathRun("C")] }),
    new MathRun(" = R · "),
    new MathSubScript({ children: [new MathRun("I")], subScript: [new MathRun("0")] }),
    new MathRun(" · "),
    new MathSuperScript({
      children: [new MathRun("e")],
      superScript: [
        new MathRun("−"),
        new MathFraction({ numerator: [new MathRun("R·t")], denominator: [new MathRun("L")] }),
      ],
    }),
    new MathRun(" = R · I(t)"),
  ], 3);
}

// Numerical substitution for L
function formulaLSubstitution() {
  return unnumberedFormula([
    new MathRun("L = −"),
    new MathFraction({
      numerator: [new MathRun("50 · 0,7")],
      denominator: [
        new MathFunction({
          name: [new MathRun("ln")],
          children: [
            new MathRoundBrackets({
              children: [
                new MathFraction({ numerator: [new MathRun("3,7661")], denominator: [new MathRun("4,4")] }),
              ],
            }),
          ],
        }),
      ],
    }),
    new MathRun(" = 225,00 Гн"),
  ]);
}

// ── Image Helper ─────────────────────────────────────────────────────────────

function embedImage(filename, widthPx, heightPx) {
  const filePath = join(__dirname, filename);
  if (!existsSync(filePath)) {
    console.warn(`Warning: ${filename} not found, skipping image`);
    return centeredParagraph(bodyRun(`[${filename} — файл не знайдено]`));
  }
  const data = readFileSync(filePath);
  const maxWidthEmu = Math.round(PAGE_WIDTH_DXA * 914400 / 1440);
  const scale = Math.min(1, maxWidthEmu / (widthPx * 9525));
  const w = Math.round(widthPx * 9525 * scale);
  const h = Math.round(heightPx * 9525 * scale);
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 60, line: LINE_SPACING_15, lineRule: "auto" },
    children: [
      new ImageRun({ data, transformation: { width: w / 9525, height: h / 9525 }, type: "png" }),
    ],
  });
}

// ── Body Sections ─────────────────────────────────────────────────────────────

const body = [
  sectionHeading("1", "Мета роботи"),
  bodyParagraph(
    "Дослідити явище самоіндукції, яке виникає при змінах сили струму в котушці індуктивності. " +
    "Визначити індуктивність соленоїда та ЕРС самоіндукції при розмиканні кола.",
  ),

  sectionHeading("2", "Експериментальні дослідження"),

  subsectionHeading("2.1", "Прилади та приладдя"),
  bodyParagraph("Для виконання лабораторної роботи використовувалася програма комп'ютерної симуляції «Самоіндукція»."),

  subsectionHeading("2.2", "Вхідні дані"),
  tableCaption("1", "Вихідні дані"),
  buildInputTable(),
  bodyParagraph(`U = ${U} В.`),

  subsectionHeading("2.3", "Результати вимірювань"),
  tableCaption("2", "Результати вимірювань"),
  buildResultsTable(),

  sectionHeading("3", "Обробка результатів"),

  subsectionHeading("3.1", "Закон спаду струму"),
  bodyParagraph("При розмиканні кола з індуктивністю сила струму спадає за експоненціальним законом:"),
  formulaCurrentDecay(),

  subsectionHeading("3.2", "Обчислення індуктивності"),
  bodyParagraph("З формули (1) знаходимо індуктивність:"),
  formulaInductance(),
  bodyParagraph("Числова підстановка для t = 0,7 с:"),
  formulaLSubstitution(),
  bodyParagraph(`Середнє значення індуктивності: L = ${L_mean.toFixed(2)} Гн.`),
  bodyParagraph("Індуктивність залишається сталою для всіх моментів часу, що підтверджує її характеристику як властивість контуру."),

  subsectionHeading("3.3", "Графіки I = f(t) та ln I = f(t)"),
  embedImage("graph_I_and_lnI.png", 1400, 500),
  figureCaption("1", "Залежності I = f(t) та ln I = f(t)"),
  bodyParagraph(`Лінійність графіка ln I = f(t) підтверджує експоненціальний закон спаду струму. Нахил прямої: k = ${(-R / L_mean).toFixed(4)}.`),

  subsectionHeading("3.4", "Обчислення ЕРС самоіндукції"),
  bodyParagraph("ЕРС самоіндукції визначається за формулою:"),
  formulaEMF(),
  unnumberedFormula([
    new MathSubScript({ children: [new MathRun("ε")], subScript: [new MathRun("max")] }),
    new MathRun(`(t = 0) = ${R} · ${I0} = ${(R * I0).toFixed(0)} В`),
  ]),
  unnumberedFormula([
    new MathSubScript({ children: [new MathRun("ε")], subScript: [new MathRun("min")] }),
    new MathRun(`(t = ${t_data.at(-1)}) = ${R} · ${I_data.at(-1)} = ${eps_values.at(-1).toFixed(2)} В`),
  ]),

  subsectionHeading("3.5", "Графік ε = f(t)"),
  embedImage("graph_eps.png", 1000, 500),
  figureCaption("2", "Залежність ε = f(t)"),

  sectionHeading("", "Висновок"),
  bodyParagraph(
    "В результаті лабораторних досліджень підтверджено, що після розмикання кола з індуктивністю " +
    "сила струму спадає за експоненціальним законом I = I₀·e^(−Rt/L). " +
    "Лінійність графіка ln I = f(t) є прямим підтвердженням цього закону.",
  ),
  bodyParagraph(
    `Індуктивність соленоїда L = ${L_mean.toFixed(2)} Гн залишається сталою для всіх моментів часу, ` +
    "що є характеристикою контуру, а не струму.",
  ),
  bodyParagraph(
    `ЕРС самоіндукції максимальна в початковий момент розмикання (ε = ${(R * I0).toFixed(0)} В) і спадає за тим же експоненціальним законом. ` +
    `Початкова сила струму I₀ = U/R = ${I0} А не залежить від параметрів соленоїда.`,
  ),
];

// ── Document Assembly ─────────────────────────────────────────────────────────

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: FONT, size: BODY_SIZE, language: { value: "uk-UA" } },
        paragraph: { spacing: { after: 0, line: LINE_SPACING_15, lineRule: "auto" } },
      },
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal",
        quickFormat: true,
        run: { size: BODY_SIZE, bold: true, font: FONT },
        paragraph: { spacing: { before: 240, after: 120, line: LINE_SPACING_15, lineRule: "auto" }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal",
        quickFormat: true,
        run: { size: BODY_SIZE, bold: true, font: FONT },
        paragraph: { spacing: { before: 120, after: 60, line: LINE_SPACING_15, lineRule: "auto" }, outlineLevel: 1 },
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { ...margins, header: 708, footer: 708 },
      },
      titlePage: true,
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: BODY_SIZE })],
          }),
        ],
      }),
    },
    children: [
      ...titlePage,
      new Paragraph({ children: [new PageBreak()] }),
      ...body,
    ],
  }],
});

const OUTPUT = join(__dirname, "Звіт_ЛР2_Коновалов_ПЗПІ-25-6.docx");
const buffer = await Packer.toBuffer(doc);
writeFileSync(OUTPUT, buffer);
console.log(`Created: ${OUTPUT}`);
