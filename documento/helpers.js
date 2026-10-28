const fs = require("fs");
const path = require("path");
const {
  Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, ImageRun, PageBreak, VerticalAlign,
} = require("docx");

const FUENTE = "Times New Roman";
const ANCHO_TEXTO = 9070; // A4 (11906) menos 3 cm + 3 cm (1701 + 1701 aprox.)

// Texto con **negrita**, *cursiva* y [[resaltado]] para pendientes
function runs(texto, base = {}) {
  const partes = texto.split(/(\*\*[^*]+\*\*|\*[^*]+\*|\[\[[^\]]+\]\])/g).filter(Boolean);
  return partes.map((t) => {
    if (t.startsWith("**")) return new TextRun({ text: t.slice(2, -2), bold: true, ...base });
    if (t.startsWith("[[")) return new TextRun({ text: t.slice(2, -2), highlight: "yellow", ...base });
    if (t.startsWith("*")) return new TextRun({ text: t.slice(1, -1), italics: true, ...base });
    return new TextRun({ text: t, ...base });
  });
}

const p = (texto, opts = {}) => new Paragraph({
  alignment: opts.align || AlignmentType.JUSTIFIED,
  children: runs(texto, opts.run || {}),
  indent: opts.indent,
  spacing: opts.spacing,
  keepNext: opts.keepNext,
});

const h1 = (texto, pageBreakBefore = true) => new Paragraph({
  heading: HeadingLevel.HEADING_1, pageBreakBefore, children: [new TextRun(texto)],
});
const h2 = (texto) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(texto)] });
const h3 = (texto) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(texto)] });

const vinetas = (lista) => lista.map((t) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED, numbering: { reference: "vinetas", level: 0 }, children: runs(t),
}));

const salto = () => new Paragraph({ children: [new PageBreak()] });

function leyenda(tipo, numero, titulo) {
  return new Paragraph({
    alignment: AlignmentType.LEFT, keepNext: true, spacing: { before: 120, after: 60, line: 240, lineRule: "auto", lineRule: "auto" },
    children: [new TextRun({ text: `${tipo} ${numero}: `, bold: true, size: 22 }),
      new TextRun({ text: titulo, size: 22 })],
  });
}
const fuente = (texto) => new Paragraph({
  alignment: AlignmentType.LEFT, spacing: { before: 40, after: 200, line: 240, lineRule: "auto", lineRule: "auto" },
  children: [new TextRun({ text: "Fuente: " + texto, italics: true, size: 20 })],
});

const borde = { style: BorderStyle.SINGLE, size: 4, color: "8C8C8C" };
const bordes = { top: borde, bottom: borde, left: borde, right: borde };

function celda(texto, ancho, { cab = false, negrita = false, der = false, sombra, unir = true } = {}) {
  return new TableCell({
    width: { size: ancho, type: WidthType.DXA }, borders: bordes, verticalAlign: VerticalAlign.CENTER,
    shading: cab ? { type: ShadingType.CLEAR, fill: "DCE6F2", color: "auto" }
      : sombra ? { type: ShadingType.CLEAR, fill: sombra, color: "auto" } : undefined,
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: [new Paragraph({
      alignment: der ? AlignmentType.RIGHT : AlignmentType.LEFT, keepNext: unir, keepLines: true, spacing: { before: 0, after: 0, line: 240, lineRule: "auto", lineRule: "auto" },
      children: runs(String(texto), { size: 20, bold: cab || negrita, font: FUENTE }),
    })],
  });
}

// filas: arrays de celdas; una fila cuyo primer valor empieza por "!" va en negrita
function tabla(cabecera, filas, anchos, { numericas = true } = {}) {
  const total = anchos.reduce((a, b) => a + b, 0);
  const rows = [new TableRow({ tableHeader: true, children: cabecera.map((c, i) => celda(c, anchos[i], { cab: true, der: numericas && i > 0 })) })];
  filas.forEach((f, n) => {
    const unir = filas.length <= 18 ? n < filas.length - 1 : n < 2;
    const negrita = String(f[0]).startsWith("!");
    rows.push(new TableRow({
      cantSplit: true,
      children: f.map((c, i) => celda(i === 0 && negrita ? String(c).slice(1) : c, anchos[i],
        { negrita, der: numericas && i > 0, sombra: negrita ? "F2F2F2" : undefined, unir })),
    }));
  });
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: anchos, rows });
}

function imagen(fichero, anchoCm) {
  const ruta = path.join(__dirname, "..", "figuras", fichero);
  const buf = fs.readFileSync(ruta);
  const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20);
  const px = anchoCm * 37.8;
  return new Paragraph({
    alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 60, after: 0, line: 240, lineRule: "auto", lineRule: "auto" },
    children: [new ImageRun({ type: "png", data: buf, transformation: { width: px, height: px * h / w } })],
  });
}

module.exports = { FUENTE, ANCHO_TEXTO, runs, p, h1, h2, h3, vinetas, salto, leyenda, fuente, tabla, imagen, celda };
