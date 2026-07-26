// Shared ExcelJS styling helpers.
import type ExcelJS from "exceljs";
import type { CellValue } from "./workbook";

export const HIGHLIGHT_UNMATCHED = {
  type: "pattern" as const,
  pattern: "solid" as const,
  fgColor: { argb: "FFFFE8B0" },
};

export const HIGHLIGHT_DUP = {
  type: "pattern" as const,
  pattern: "solid" as const,
  fgColor: { argb: "FFFFCCCC" },
};

export const HEADER_STYLE = {
  font: { bold: true, color: { argb: "FFFFFFFF" } },
  fill: { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FF1F4033" } },
  alignment: { vertical: "middle" as const, horizontal: "left" as const },
};

export function styleHeader(ws: ExcelJS.Worksheet) {
  const row = ws.getRow(1);
  row.eachCell((cell) => {
    cell.font = HEADER_STYLE.font;
    cell.fill = HEADER_STYLE.fill;
    cell.alignment = HEADER_STYLE.alignment;
  });
  row.height = 22;
  row.commit();
}

export function autoWidth(ws: ExcelJS.Worksheet, headers: string[], rows: CellValue[][]) {
  headers.forEach((h, i) => {
    const col = ws.getColumn(i + 1);
    const dataMax = rows.reduce((m, r) => Math.max(m, String(r[i] ?? "").length), 0);
    col.width = Math.min(Math.max(h.length + 2, dataMax + 2, 12), 42);
  });
}
