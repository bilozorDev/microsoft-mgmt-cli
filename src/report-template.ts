import ExcelJS from "exceljs";

function colLetter(n: number): string {
  return String.fromCharCode(64 + n);
}

export interface ReportOptions {
  sheetName: string;
  title: string;
  tenant: string;
  summary: string;
  columns: { header: string; width: number }[];
  rows: string[][];
}

/**
 * Generates a styled Excel workbook for M365 admin reports.
 * Layout: title (A1), company (A2), date/summary (A3-A4),
 * divider (row 3 bottom), spacer (row 4), table header (row 5),
 * data (row 6+).
 */
export async function generateReport(opts: ReportOptions): Promise<Buffer> {
  const { sheetName, title, tenant, summary, columns, rows } = opts;
  const colCount = columns.length;
  const lastCol = colLetter(colCount);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(sheetName, {
    views: [{ showGridLines: false }],
  });

  // Column widths
  ws.columns = columns.map((c) => ({ width: c.width }));

  // --- Title (row 1) ---
  ws.mergeCells(`A1:${lastCol}1`);
  const titleCell = ws.getCell("A1");
  titleCell.value = title;
  titleCell.font = { size: 18, bold: true, color: { argb: "FF1B3A5C" } };
  titleCell.alignment = { vertical: "middle" };

  // --- Company name (row 2) ---
  ws.mergeCells(`A2:${lastCol}2`);
  ws.getCell("A2").value = tenant;
  ws.getCell("A2").font = { size: 11, color: { argb: "FF666666" } };

  // --- Date (row 3) ---
  ws.mergeCells(`A3:${lastCol}3`);
  const reportDate = new Date().toLocaleDateString("en-US", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });
  const dateCell = ws.getCell("A3");
  dateCell.value = `Report generated: ${reportDate}`;
  dateCell.font = { size: 11, color: { argb: "FF666666" } };

  // --- Divider (row 3 bottom border) ---
  for (let col = 1; col <= colCount; col++) {
    ws.getCell(3, col).border = {
      bottom: { style: "medium", color: { argb: "FF2B5797" } },
    };
  }

  // --- Summary (row 4) ---
  ws.mergeCells(`A4:${lastCol}4`);
  ws.getCell("A4").value = summary;
  ws.getCell("A4").font = { size: 11, italic: true, color: { argb: "FF666666" } };

  // --- Table header (row 5) ---
  const headerRow = ws.getRow(5);
  columns.forEach((c, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = c.header;
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF2B5797" },
    };
    cell.alignment = { vertical: "middle" };
    cell.border = {
      top: { style: "thin", color: { argb: "FFB0B0B0" } },
      bottom: { style: "thin", color: { argb: "FFB0B0B0" } },
      left: { style: "thin", color: { argb: "FFB0B0B0" } },
      right: { style: "thin", color: { argb: "FFB0B0B0" } },
    };
  });

  // --- Data rows (row 6+) ---
  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: "thin", color: { argb: "FFB0B0B0" } },
    bottom: { style: "thin", color: { argb: "FFB0B0B0" } },
    left: { style: "thin", color: { argb: "FFB0B0B0" } },
    right: { style: "thin", color: { argb: "FFB0B0B0" } },
  };
  const altFill: ExcelJS.Fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE8EDF2" },
  };
  const dataFont: Partial<ExcelJS.Font> = {
    size: 11,
    name: "Calibri",
    family: 2,
  };

  rows.forEach((values, idx) => {
    const row = ws.getRow(6 + idx);
    values.forEach((v, ci) => {
      const cell = row.getCell(ci + 1);
      cell.value = v;
      cell.font = dataFont;
      cell.border = thinBorder;
      if (idx % 2 === 0) {
        cell.fill = altFill;
      }
    });
  });

  return (await wb.xlsx.writeBuffer()) as unknown as Buffer;
}
