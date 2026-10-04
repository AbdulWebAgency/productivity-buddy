import type { SheetHandle } from "@/lib/excel/engine/shared/workbook";

export function detectHeaderRow(worksheet: SheetHandle, scanRows = 15): number {
  let bestRow = 1;
  let bestScore = -1;

  const maxRow = Math.min(scanRows, worksheet.rowCount);

  for (let r = 1; r <= maxRow; r++) {
    let nonEmpty = 0;
    let textCells = 0;

    for (let c = 1; c <= worksheet.columnCount; c++) {
      const value = worksheet.getCell(r, c).value;

      if (value == null || value === "") continue;

      nonEmpty++;

      if (typeof value === "string") {
        textCells++;
      }
    }

    // Ignore rows that only contain one title cell
    if (nonEmpty < 2) continue;

    const score = nonEmpty + textCells * 2;

    if (score > bestScore) {
      bestScore = score;
      bestRow = r;
    }
  }

  return bestRow;
}
