// Shared output styling helpers (SheetJS).
//
// SheetJS Community Edition limitation: it cannot write cell fills, fonts,
// alignment, row heights or frozen panes. Header styling and
// row-highlight fills therefore have no runtime equivalent. Only column widths
// (`!cols`) are supported and preserved.
import type { ColInfo } from "xlsx";
import type { CellValue } from "./workbook";

/** Auto column widths — clamp(max(header+2, data+2, 12), 42). */
export function columnWidths(headers: string[], rows: CellValue[][]): ColInfo[] {
  return headers.map((h, i) => {
    const dataMax = rows.reduce((m, r) => Math.max(m, String(r[i] ?? "").length), 0);
    return { wch: Math.min(Math.max(h.length + 2, dataMax + 2, 12), 42) };
  });
}
