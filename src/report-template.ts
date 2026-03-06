import { join } from "path";
import ExcelJS from "exceljs";

function getLogoPath(): string {
  return join(import.meta.dir, "assets", "logo.png");
}

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
 * Generates a branded Excel workbook matching the Profulgent report template.
 * Layout: logo (cols A-B rows 1-4), title (C1), company (C2), date (C3),
 * summary (C4), contact (row 6), divider (row 7), table header (row 9),
 * data (row 10+).
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

  // --- Logo (cols A-B, rows 1-4) ---
  const logoId = wb.addImage({
    filename: getLogoPath(),
    extension: "png",
  });
  ws.addImage(logoId, {
    tl: { col: 0, row: 0 } as any,
    br: { col: 2, row: 4 } as any,
  });

  // --- Title (row 1) ---
  ws.mergeCells(`C1:${lastCol}1`);
  const titleCell = ws.getCell("C1");
  titleCell.value = title;
  titleCell.font = { size: 18, bold: true, color: { argb: "FF1B3A5C" } };
  titleCell.alignment = { vertical: "middle" };

  // --- Company name (row 2) ---
  ws.mergeCells(`C2:${lastCol}2`);
  ws.getCell("C2").value = tenant;
  ws.getCell("C2").font = { size: 11, color: { argb: "FF666666" } };

  // --- Date (row 3) ---
  ws.mergeCells(`C3:${lastCol}3`);
  const reportDate = new Date().toLocaleDateString("en-US", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });
  ws.getCell("C3").value = `Report generated: ${reportDate}`;
  ws.getCell("C3").font = { size: 11, color: { argb: "FF666666" } };

  // --- Summary (row 4) ---
  ws.mergeCells(`C4:${lastCol}4`);
  ws.getCell("C4").value = summary;
  ws.getCell("C4").font = { size: 11, italic: true, color: { argb: "FF666666" } };

  // --- Row 5: spacer ---

  // --- Contact info (row 6) ---
  ws.mergeCells(`A6:${lastCol}6`);
  ws.getCell("A6").value =
    "Profulgent · Helpdesk +1 732 242 9345 · support@profulgent.net";
  ws.getCell("A6").font = { size: 9, color: { argb: "FF999999" } };
  ws.getCell("A6").alignment = { horizontal: "center" };

  // --- Divider (row 7) ---
  for (let col = 1; col <= colCount; col++) {
    ws.getCell(7, col).border = {
      bottom: { style: "medium", color: { argb: "FF2B5797" } },
    };
  }

  // --- Row 8: spacer ---

  // --- Table header (row 9) ---
  const headerRow = ws.getRow(9);
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

  // --- Data rows (row 10+) ---
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
    const row = ws.getRow(10 + idx);
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
