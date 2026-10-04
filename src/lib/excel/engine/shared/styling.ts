// Shared output styling helpers (SheetJS).
//
// SheetJS Community Edition limitation: it cannot write cell fills, fonts,
// alignment, row heights or frozen panes. The former ExcelJS header style and
// row-highlight fills therefore have no runtime equivalent. Only column widths
// (`!cols`) are supported and preserved.
import type { ColInfo } from "xlsx";
import type { CellValue } from "./workbook";

/** Auto column widths — same rule as the former ExcelJS autoWidth. */
export function columnWidths(headers: string[], rows: CellValue[][]): ColInfo[] {
  return headers.map((h, i) => {
    const dataMax = rows.reduce((m, r) => Math.max(m, String(r[i] ?? "").length), 0);
    return { wch: Math.min(Math.max(h.length + 2, dataMax + 2, 12), 42) };
  });
}
