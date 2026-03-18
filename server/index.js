const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const cors = require("cors");
const PDFDocument = require("pdfkit");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, AlignmentType, WidthType, BorderStyle, ShadingType,
  VerticalAlign, HeadingLevel
} = require("docx");

const app = express();
app.use(cors());
app.use(express.json());

// Ensure folders exist
const uploadsDir = path.join(__dirname, "uploads");
const lettersDir = path.join(__dirname, "letters");
if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir);
if (!fs.existsSync(lettersDir)) fs.mkdirSync(lettersDir);

const upload = multer({ dest: uploadsDir });

// ─── Parse column headers to detect subject structure ────────────────────────
function extractSubjects(headers) {
  // headers is the array from row index 6 (0-based)
  // Each subject has a -TH column and optionally a -PR or -L-PR column
  const subjects = [];
  const seen = new Set();

  for (const col of headers) {
    if (!col || col === "Sr.No" || col === "PRN" || col === "Name of the Student") continue;
    if (col.startsWith("Overall") || col.startsWith("Total")) continue;

    const upper = col.toString().toUpperCase();

    // Detect TH columns → derive subject name
    if (upper.includes("-TH") || upper.endsWith("TH")) {
      let base = col.toString().replace(/-TH$/i, "").replace(/TH$/i, "").trim();
      // Normalise: "OE-V IE" → keep as-is
      if (!seen.has(base)) {
        seen.add(base);
        // Find matching PR column
        const prCol = headers.find(h => {
          if (!h) return false;
          const hStr = h.toString();
          return (
            hStr.toUpperCase().includes(base.toUpperCase()) &&
            (hStr.toUpperCase().includes("-PR") || hStr.toUpperCase().includes("L-PR"))
          );
        }) || null;

        subjects.push({ name: base, th: col.toString(), pr: prCol });
      }
    }
  }

  // Also add standalone PR-only subjects (PBL etc.)
  for (const col of headers) {
    if (!col) continue;
    const upper = col.toString().toUpperCase();
    if ((upper.includes("-PR") || upper.endsWith("-PR")) && !upper.includes("-TH")) {
      const base = col.toString().replace(/-PR$/i, "").replace(/L-PR$/i, "").trim();
      const alreadyCovered = subjects.some(s => s.pr === col.toString());
      if (!alreadyCovered && !seen.has(base)) {
        seen.add(base);
        subjects.push({ name: base, th: null, pr: col.toString() });
      }
    }
  }

  return subjects;
}

// ─── Parse Excel ─────────────────────────────────────────────────────────────
app.post("/upload", upload.single("file"), (req, res) => {
  try {
    const workbook = XLSX.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    // Find header row (row with "PRN")
    let headerRowIdx = -1;
    for (let i = 0; i < raw.length; i++) {
      if (raw[i].includes("PRN") || raw[i].includes("Name of the Student")) {
        headerRowIdx = i;
        break;
      }
    }
    if (headerRowIdx === -1) return res.status(400).json({ error: "Could not find header row in Excel" });

    const headers = raw[headerRowIdx];
    const subjects = extractSubjects(headers);

    // Read meta info from rows above header
    let division = "A", semester = "I", academicYear = "2025-2026", uptoDate = "", attendanceDuration = "";
    for (let i = 0; i < headerRowIdx; i++) {
      const rowStr = raw[i].join(" ");
      const divMatch = rowStr.match(/Division\s*:\s*(\w+)/i);
      if (divMatch) division = divMatch[1];
      const semMatch = rowStr.match(/Semester\s*:\s*(\w+)/i);
      if (semMatch) semester = semMatch[1];
      const ayMatch = rowStr.match(/Academic Year\s*:\s*([\d\-\/]+)/i);
      if (ayMatch) academicYear = ayMatch[1];
      // Extract full duration string e.g. "07-July-2025 to 06-November-2025"
      const durMatch = rowStr.match(/(\d{2}[-\/]\w+[-\/]\d{4})\s+to\s+(\d{2}[-\/]\w+[-\/]\d{4})/i);
      if (durMatch) {
        attendanceDuration = durMatch[1] + " to " + durMatch[2];
        uptoDate = durMatch[2];
      }
    }

    // Data rows start after header + 1 (totals row)
    const dataStartIdx = headerRowIdx + 2;
    const students = [];

    for (let i = dataStartIdx; i < raw.length; i++) {
      const row = raw[i];
      if (!row[0] || String(row[0]).trim() === "") continue;

      const student = {};
      headers.forEach((h, idx) => {
        if (h) student[h] = row[idx] !== undefined ? String(row[idx]).trim() : "";
      });

      // Parse overall attendance
      const overallRaw = student["Overall Att."] || student["Overall"] || "";
      const overall = parseFloat(overallRaw.replace("%", "").trim()) || 0;

      if (overall < 75) {
        student._overall = overall;
        students.push(student);
      }
    }

    res.json({ students, subjects, division, semester, academicYear, uptoDate, attendanceDuration });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ─── Shared: build subject rows ───────────────────────────────────────────────
function buildRows(student, subjects) {
  return subjects.map((sub, idx) => {
    const th = sub.th ? (student[sub.th] || "-") : "-";
    const pr = sub.pr ? (student[sub.pr] || "-") : "-";
    return { idx: idx + 1, name: sub.name, th, pr };
  });
}

function avg(rows, key) {
  const vals = rows.map(r => parseFloat(r[key])).filter(v => !isNaN(v));
  if (!vals.length) return "-";
  return (vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(1) + "%";
}

// ─── Generate PDF (pdfkit — no Chrome/puppeteer needed) ──────────────────────
app.post("/generate/pdf", async (req, res) => {
  const { student, subjects, division, semester, uptoDate, attendanceDuration, classTeacher, academicCoordinator } = req.body;
  const rows = buildRows(student, subjects);
  const avgTH = avg(rows, "th");
  const avgPR = avg(rows, "pr");
  const total = student["Overall Att."] || student["Overall"] || "-";
  const prn = student["PRN"] || "student";
  const name = student["Name of the Student"] || "";
  const filePath = path.join(__dirname, "letters", `${prn}.pdf`);

  const doc = new PDFDocument({ size: "A4", margin: 36 });
  const stream = fs.createWriteStream(filePath);
  doc.pipe(stream);

  const W = 595 - 72; // page width minus margins
  const LEFT = 36;
  const BLUE = "#1a237e";
  const GREY = "#f2f2f2";
  const HBLUE = "#d9e1f2";

  // ── Helper: draw a rect-bordered row ──
  function drawRect(x, y, w, h, fill) {
    if (fill) { doc.rect(x, y, w, h).fill(fill).stroke(); }
    else { doc.rect(x, y, w, h).stroke(); }
  }

  function cellText(text, x, y, w, h, opts = {}) {
    doc.font(opts.bold ? "Helvetica-Bold" : "Helvetica")
       .fontSize(opts.size || 9)
       .fillColor(opts.color || "black");
    const textY = y + (h - (opts.size || 9)) / 2;
    const align = opts.align || "left";
    doc.text(String(text), x + 4, textY, { width: w - 8, align, lineBreak: false });
  }

  // ── HEADER TABLE ──
  let y = 36;
  const headerH = 56;
  const col1 = 70, col2 = W - 70 - 150, col3 = 150;

  // Logo cell
  drawRect(LEFT, y, col1, headerH * 3, null);
  try {
    doc.image(path.join(__dirname, "logo.png"), LEFT + 5, y + 5, { width: col1 - 10, height: headerH * 3 - 10, fit: [col1 - 10, headerH * 3 - 10] });
  } catch(e) {}

  // College name cell
  drawRect(LEFT + col1, y, col2, headerH * 3, null);
  doc.font("Helvetica").fontSize(9).fillColor("black")
     .text("Pimpri Chinchwad Education Trust's", LEFT + col1 + 4, y + 18, { width: col2 - 8, align: "center" });
  doc.font("Helvetica-Bold").fontSize(10).fillColor("black")
     .text("Pimpri Chinchwad College of Engineering", LEFT + col1 + 4, y + 32, { width: col2 - 8, align: "center" });

  // Record cells
  const rcX = LEFT + col1 + col2;
  drawRect(rcX, y, col3, headerH, null);
  cellText("Record No.: ACAD/R/23", rcX, y, col3, headerH, { size: 8 });
  drawRect(rcX, y + headerH, col3, headerH, null);
  cellText("Revision: 01", rcX, y + headerH, col3, headerH, { size: 8 });
  drawRect(rcX, y + headerH * 2, col3, headerH, null);
  cellText("Date: 28/08/2024", rcX, y + headerH * 2, col3, headerH, { size: 8 });

  y += headerH * 3;
  // Title row
  const titleH = 22;
  drawRect(LEFT, y, W, titleH, HBLUE);
  cellText("Letter to Parents of Poor Performing Students", LEFT, y, W, titleH, { bold: true, size: 10, align: "center" });
  y += titleH + 8;

  // ── META INFO ──
  doc.font("Helvetica").fontSize(9).fillColor("black")
     .text(`Department: Computer Engineering     Academic Year: 2025-2026     Semester: ${semester || "I"} / II`, LEFT, y);
  y += 18;

  // ── BODY ──
  doc.font("Helvetica").fontSize(9).text("To,", LEFT, y); y += 13;
  doc.text("Dear Sir,", LEFT, y); y += 13;
  doc.font("Helvetica").fontSize(9)
     .text("We are sorry to inform you that attendance of your ward ", LEFT, y, { continued: true })
     .font("Helvetica-Bold").text(`${name} `, { continued: true })
     .font("Helvetica").text("PRN No. ", { continued: true })
     .font("Helvetica-Bold").text(`${prn} `, { continued: true })
     .font("Helvetica").text("Year ", { continued: true })
     .font("Helvetica-Bold").text("B.Tech ", { continued: true })
     .font("Helvetica").text("Div ", { continued: true })
     .font("Helvetica-Bold").text(`${division || "A"} `, { continued: true })
     .font("Helvetica").text("is poor.");
  y += 18;

  doc.font("Helvetica").fontSize(9)
     .text("1. Subject wise attendance from ", LEFT, y, { continued: true })
     .font("Helvetica-Bold").text(`${attendanceDuration || uptoDate || ""}`, { continued: true })
     .font("Helvetica").text(" is as follows.");
  y += 14;

  // ── ATTENDANCE TABLE ──
  const c = [30, 170, 90, 90]; // col widths
  const rowH = 18;
  const tLeft = LEFT;

  // Header row
  drawRect(tLeft, y, c[0], rowH, HBLUE);
  drawRect(tLeft + c[0], y, c[1], rowH, HBLUE);
  drawRect(tLeft + c[0] + c[1], y, c[2], rowH, HBLUE);
  drawRect(tLeft + c[0] + c[1] + c[2], y, c[3], rowH, HBLUE);
  cellText("Sr. No", tLeft, y, c[0], rowH, { bold: true, align: "center" });
  cellText("Subject", tLeft + c[0], y, c[1], rowH, { bold: true });
  cellText("Theory Att. (%)", tLeft + c[0] + c[1], y, c[2], rowH, { bold: true, align: "center", size: 8 });
  cellText("Practical Att. (%)", tLeft + c[0] + c[1] + c[2], y, c[3], rowH, { bold: true, align: "center", size: 8 });
  y += rowH;

  rows.forEach(r => {
    drawRect(tLeft, y, c[0], rowH, null);
    drawRect(tLeft + c[0], y, c[1], rowH, null);
    drawRect(tLeft + c[0] + c[1], y, c[2], rowH, null);
    drawRect(tLeft + c[0] + c[1] + c[2], y, c[3], rowH, null);
    cellText(r.idx, tLeft, y, c[0], rowH, { align: "center" });
    cellText(r.name, tLeft + c[0], y, c[1], rowH);
    cellText(r.th, tLeft + c[0] + c[1], y, c[2], rowH, { align: "center" });
    cellText(r.pr, tLeft + c[0] + c[1] + c[2], y, c[3], rowH, { align: "center" });
    y += rowH;
  });

  // Average row
  drawRect(tLeft, y, c[0] + c[1], rowH, GREY);
  drawRect(tLeft + c[0] + c[1], y, c[2], rowH, GREY);
  drawRect(tLeft + c[0] + c[1] + c[2], y, c[3], rowH, GREY);
  cellText("Average Attendance (%)", tLeft, y, c[0] + c[1], rowH, { bold: true });
  cellText(avgTH, tLeft + c[0] + c[1], y, c[2], rowH, { bold: true, align: "center" });
  cellText(avgPR, tLeft + c[0] + c[1] + c[2], y, c[3], rowH, { bold: true, align: "center" });
  y += rowH;

  // Total row
  drawRect(tLeft, y, c[0] + c[1], rowH, GREY);
  drawRect(tLeft + c[0] + c[1], y, c[2] + c[3], rowH, GREY);
  cellText("Total Attendance (%)", tLeft, y, c[0] + c[1], rowH, { bold: true });
  cellText(total, tLeft + c[0] + c[1], y, c[2] + c[3], rowH, { bold: true, align: "center" });
  y += rowH + 12;

  // ── WARNING TEXT ──
  doc.font("Helvetica").fontSize(9).fillColor("black")
     .text(
       "If he/she fails to improve attendance and to satisfy the minimum criteria of 75% attendance in theory and practical's conducted, by college, he/she shall not be eligible to appear for Final SA in Semester I / II Theory Examination.",
       LEFT, y, { width: W, align: "justify" }
     );
  y += 50;

  // ── SIGNATURES ──
  const sigW = W / 3;
  const sigs = [
    { title: "Class Teacher", name: classTeacher || "" },
    { title: "Academic Coordinator", name: academicCoordinator || "" },
    { title: "Head of the Department", name: "Dr. Sonali Patil" },
  ];
  sigs.forEach((s, i) => {
    const sx = LEFT + i * sigW;
    doc.font("Helvetica-Bold").fontSize(9).fillColor("black")
       .text(s.title, sx, y, { width: sigW, align: "center" });
    doc.font("Helvetica").fontSize(9)
       .text(s.name, sx, y + 13, { width: sigW, align: "center" });
  });

  doc.end();

  await new Promise((resolve, reject) => {
    stream.on("finish", resolve);
    stream.on("error", reject);
  });

  res.download(filePath);
});

// ─── Generate DOCX ────────────────────────────────────────────────────────────
app.post("/generate/docx", async (req, res) => {
  const { student, subjects, division, semester, uptoDate, attendanceDuration, classTeacher, academicCoordinator } = req.body;

  const rows = buildRows(student, subjects);
  const avgTH = avg(rows, "th");
  const avgPR = avg(rows, "pr");
  const total = student["Overall Att."] || student["Overall"] || "-";
  const logoBuffer = fs.readFileSync("logo.png");

  // Border helper
  const border = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const noBorder = { style: BorderStyle.NONE, size: 0 };
  const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

  // Cell helper
  const cell = (text, opts = {}) => new TableCell({
    borders: opts.borders || borders,
    width: opts.width || { size: 2000, type: WidthType.DXA },
    shading: opts.shading,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    rowSpan: opts.rowSpan,
    columnSpan: opts.columnSpan,
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({ text: String(text), bold: opts.bold || false, size: opts.size || 20, font: "Arial" })]
    })]
  });

  // ── Header table (logo | college name | record info) ──
  const headerTable = new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [1100, 6360, 1900],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders,
            width: { size: 1100, type: WidthType.DXA },
            rowSpan: 3,
            verticalAlign: VerticalAlign.CENTER,
            margins: { top: 60, bottom: 60, left: 60, right: 60 },
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new ImageRun({ data: logoBuffer, transformation: { width: 60, height: 60 }, type: "png" })]
            })]
          }),
          new TableCell({
            borders,
            width: { size: 6360, type: WidthType.DXA },
            rowSpan: 3,
            verticalAlign: VerticalAlign.CENTER,
            margins: { top: 60, bottom: 60, left: 100, right: 100 },
            children: [
              new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Pimpri Chinchwad Education Trust's", size: 20, font: "Arial" })] }),
              new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Pimpri Chinchwad College of Engineering", bold: true, size: 22, font: "Arial" })] }),
            ]
          }),
          cell("Record No.: ACAD/R/23", { size: 18, width: { size: 1900, type: WidthType.DXA } }),
        ]
      }),
      new TableRow({ children: [cell("Revision: 01", { size: 18, width: { size: 1900, type: WidthType.DXA } })] }),
      new TableRow({ children: [cell("Date: 28/08/2024", { size: 18, width: { size: 1900, type: WidthType.DXA } })] }),
      new TableRow({
        children: [
          new TableCell({
            borders,
            width: { size: 9360, type: WidthType.DXA },
            columnSpan: 3,
            margins: { top: 60, bottom: 60, left: 100, right: 100 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Letter to Parents of Poor Performing Students", bold: true, size: 22, font: "Arial" })] })]
          })
        ]
      })
    ]
  });

  // ── Attendance table ──
  const headerShading = { fill: "D9E1F2", type: ShadingType.CLEAR };
  const summaryShading = { fill: "F2F2F2", type: ShadingType.CLEAR };
  const colWidths = [800, 3200, 2100, 2100]; // Reduced last col since only 4 cols

  const attTableRows = [
    new TableRow({
      tableHeader: true,
      children: [
        cell("Sr. No", { bold: true, shading: headerShading, align: AlignmentType.CENTER, width: { size: colWidths[0], type: WidthType.DXA } }),
        cell("Subject", { bold: true, shading: headerShading, width: { size: colWidths[1], type: WidthType.DXA } }),
        cell("Theory Attendance (%)", { bold: true, shading: headerShading, align: AlignmentType.CENTER, width: { size: colWidths[2], type: WidthType.DXA } }),
        cell("Practical Attendance (%)", { bold: true, shading: headerShading, align: AlignmentType.CENTER, width: { size: colWidths[3], type: WidthType.DXA } }),
      ]
    }),
    ...rows.map(r => new TableRow({
      children: [
        cell(r.idx, { align: AlignmentType.CENTER, width: { size: colWidths[0], type: WidthType.DXA } }),
        cell(r.name, { width: { size: colWidths[1], type: WidthType.DXA } }),
        cell(r.th, { align: AlignmentType.CENTER, width: { size: colWidths[2], type: WidthType.DXA } }),
        cell(r.pr, { align: AlignmentType.CENTER, width: { size: colWidths[3], type: WidthType.DXA } }),
      ]
    })),
    new TableRow({
      children: [
        new TableCell({
          borders, columnSpan: 2, shading: summaryShading,
          width: { size: colWidths[0] + colWidths[1], type: WidthType.DXA },
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [new Paragraph({ children: [new TextRun({ text: "Average Attendance (%)", bold: true, size: 20, font: "Arial" })] })]
        }),
        cell(avgTH, { bold: true, shading: summaryShading, align: AlignmentType.CENTER, width: { size: colWidths[2], type: WidthType.DXA } }),
        cell(avgPR, { bold: true, shading: summaryShading, align: AlignmentType.CENTER, width: { size: colWidths[3], type: WidthType.DXA } }),
      ]
    }),
    new TableRow({
      children: [
        new TableCell({
          borders, columnSpan: 2, shading: summaryShading,
          width: { size: colWidths[0] + colWidths[1], type: WidthType.DXA },
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [new Paragraph({ children: [new TextRun({ text: "Total Attendance (%)", bold: true, size: 20, font: "Arial" })] })]
        }),
        new TableCell({
          borders, columnSpan: 2, shading: summaryShading,
          width: { size: colWidths[2] + colWidths[3], type: WidthType.DXA },
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: total, bold: true, size: 20, font: "Arial" })] })]
        }),
      ]
    })
  ];

  const attTable = new Table({ width: { size: 9200, type: WidthType.DXA }, columnWidths: colWidths, rows: attTableRows });

  // ── Signature table (borderless, 3 columns) ──
  const nb = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const noBordersAll = { top: nb, bottom: nb, left: nb, right: nb, insideHorizontal: nb, insideVertical: nb };
  const sigTable = new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [3120, 3120, 3120],
    borders: noBordersAll,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: { top: nb, bottom: nb, left: nb, right: nb },
            width: { size: 3120, type: WidthType.DXA },
            margins: { top: 60, bottom: 60, left: 0, right: 60 },
            children: [
              new Paragraph({ children: [new TextRun({ text: "Class Teacher", bold: true, size: 20, font: "Arial" })] }),
              new Paragraph({ children: [new TextRun({ text: classTeacher || "", size: 20, font: "Arial" })] }),
            ]
          }),
          new TableCell({
            borders: { top: nb, bottom: nb, left: nb, right: nb },
            width: { size: 3120, type: WidthType.DXA },
            margins: { top: 60, bottom: 60, left: 0, right: 60 },
            children: [
              new Paragraph({ children: [new TextRun({ text: "Academic Coordinator", bold: true, size: 20, font: "Arial" })] }),
              new Paragraph({ children: [new TextRun({ text: academicCoordinator || "", size: 20, font: "Arial" })] }),
            ]
          }),
          new TableCell({
            borders: { top: nb, bottom: nb, left: nb, right: nb },
            width: { size: 3120, type: WidthType.DXA },
            margins: { top: 60, bottom: 60, left: 0, right: 0 },
            children: [
              new Paragraph({ children: [new TextRun({ text: "Head of the Department", bold: true, size: 20, font: "Arial" })] }),
              new Paragraph({ children: [new TextRun({ text: "Dr. Sonali Patil", size: 20, font: "Arial" })] }),
            ]
          }),
        ]
      })
    ]
  });

  const p = (text, opts = {}) => new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, size: opts.size || 20, bold: opts.bold || false, font: "Arial" })]
  });

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 720, right: 720, bottom: 720, left: 720 }
        }
      },
      children: [
        headerTable,
        p(""),
        p(`Department: Computer Engineering          Academic Year: 2025-2026          Semester: ${semester || "I"} / II`),
        p(""),
        p("To,"),
        p("Dear Sir,"),
        p(""),
        new Paragraph({
          spacing: { before: 80, after: 80 },
          children: [
            new TextRun({ text: "We are sorry to inform you that attendance of your ward ", size: 20, font: "Arial" }),
            new TextRun({ text: student["Name of the Student"] || "", bold: true, size: 20, font: "Arial" }),
            new TextRun({ text: " PRN No. ", size: 20, font: "Arial" }),
            new TextRun({ text: student["PRN"] || "", bold: true, size: 20, font: "Arial" }),
            new TextRun({ text: " Year ", size: 20, font: "Arial" }),
            new TextRun({ text: "B.Tech", bold: true, size: 20, font: "Arial" }),
            new TextRun({ text: " Div ", size: 20, font: "Arial" }),
            new TextRun({ text: division || "A", bold: true, size: 20, font: "Arial" }),
            new TextRun({ text: " is poor.", size: 20, font: "Arial" }),
          ]
        }),
        p(""),
        new Paragraph({
          spacing: { before: 80, after: 80 },
          children: [
            new TextRun({ text: "1. Subject wise attendance from ", size: 20, font: "Arial" }),
            new TextRun({ text: attendanceDuration || uptoDate || "", bold: true, size: 20, font: "Arial" }),
            new TextRun({ text: " is as follows.", size: 20, font: "Arial" }),
          ]
        }),
        p(""),
        attTable,
        p(""),
        p("If he/she fails to improve attendance and to satisfy the minimum criteria of 75% attendance in theory and practical's conducted, by college, he/she shall not be eligible to appear for Final SA in Semester I / II Theory Examination."),
        p(""),
        p(""),
        sigTable,
      ]
    }]
  });

  const prn = student["PRN"] || "student";
  const filePath = path.join(__dirname, "letters", `${prn}.docx`);
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(filePath, buffer);

  res.download(filePath);
});

app.listen(5000, () => console.log("✅ Server running on http://localhost:5000"));
