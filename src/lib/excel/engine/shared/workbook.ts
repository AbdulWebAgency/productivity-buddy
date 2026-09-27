// Workbook I/O: reading, grid extraction, sheet writing, sheet-name safety.
import type ExcelJS from "exceljs";
import type { SheetMeta } from "../../types";
import { HIGHLIGHT_UNMATCHED, autoWidth, styleHeader } from "./styling";
import { detectHeaderRow } from "@/lib/excel/header-detection";

export type CellValue = string | number | boolean | null;
export type SheetGrid = { headers: string[]; rows: CellValue[][] };

export async function readWorkbook(buffer: ArrayBuffer): Promise<ExcelJS.Workbook> {
  const ExcelJSModule = await import("exceljs");
  const wb = new ExcelJSModule.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

export function cellToValue(v: ExcelJS.CellValue): CellValue {
  if (v == null) return null;
  if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") return v;
  if (v instanceof Date) return v.toISOString();
  if (typeof v === "object") {
    if ("text" in v && typeof v.text === "string") return v.text;
    if ("richText" in v && Array.isArray(v.richText)) return v.richText.map((r) => r.text).join("");
    if ("result" in v) return cellToValue(v.result as ExcelJS.CellValue);
    if ("formula" in v) return null;
  }
  return String(v);
}

export function extractSheetMeta(wb: ExcelJS.Workbook): SheetMeta {
  const sheets = wb.worksheets.map((ws) => {
    const headers: string[] = [];
    const first = ws.getRow(1);
    first.eachCell({ includeEmpty: false }, (cell, col) => {
      headers[col - 1] = String(cell.value ?? "").trim();
    });
    const sampleRows: CellValue[][] = [];
    for (let r = 2; r <= Math.min(4, ws.rowCount); r++) {
      const row = ws.getRow(r);
      const out: CellValue[] = [];
      for (let c = 1; c <= headers.length; c++) {
        out.push(cellToValue(row.getCell(c).value));
      }
      sampleRows.push(out);
    }
    return {
      name: ws.name,
      rowCount: ws.rowCount,
      columnCount: ws.columnCount,
      headers: headers.filter(Boolean),
      sampleRows,
    };
  });
  return { sheets };
}

export function sheetToGrid(ws: ExcelJS.Worksheet): SheetGrid {
  const headers: string[] = [];
  const headerRow = detectHeaderRow(ws);
  const first = ws.getRow(headerRow);
  const maxCol = ws.columnCount;
  for (let c = 1; c <= maxCol; c++) {
    headers.push(String(first.getCell(c).value ?? "").trim());
  }
  const rows: CellValue[][] = [];
  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const out: CellValue[] = [];
    let any = false;
    for (let c = 1; c <= maxCol; c++) {
      const v = cellToValue(row.getCell(c).value);
      if (v !== null && v !== "") any = true;
      out.push(v);
    }
    if (any) rows.push(out);
  }
  return { headers, rows };
}

export function writeGridToSheet(
  wb: ExcelJS.Workbook,
  sheetName: string,
  grid: SheetGrid,
  opts: { highlightRows?: Set<number>; headerStyle?: boolean } = {},
): ExcelJS.Worksheet {
  const existing = wb.getWorksheet(sheetName);
  if (existing) wb.removeWorksheet(existing.id);
  const ws = wb.addWorksheet(sheetName, {
    views: [{ state: "frozen", ySplit: 1 }],
  });
  ws.addRow(grid.headers);
  grid.rows.forEach((r) => ws.addRow(r));
  if (opts.headerStyle !== false) styleHeader(ws);
  autoWidth(ws, grid.headers, grid.rows);
  if (opts.highlightRows) {
    opts.highlightRows.forEach((rowIdx) => {
      const row = ws.getRow(rowIdx + 2); // +1 header, +1 1-indexed
      row.eachCell((cell) => {
        cell.fill = HIGHLIGHT_UNMATCHED;
      });
    });
  }
  return ws;
}

export function safeSheetName(name: string): string {
  // Excel sheet names: max 31 chars, no : \ / ? * [ ]
  return name
    .replace(/\.[^.]+$/, "")
    .replace(/[\\/:?*[\]]/g, " ")
    .slice(0, 25);
}
