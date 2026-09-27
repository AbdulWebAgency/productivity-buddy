// Workbook compatibility boundary.
//
// Phase 1 of the ExcelJS -> SheetJS migration. SheetJS (`xlsx`) is isolated
// behind WorkbookHandle / SheetHandle / RowHandle / CellHandle, which keep the
// engine's 1-indexed (row, col) calling convention. SheetJS is 0-indexed
// internally; all translation happens here.
//
// Transitional: callers not yet migrated (runPlan, highlight, registration
// functions) still pass ExcelJS objects. `extractSheetMeta`, `sheetToGrid`,
// `writeGridToSheet` and `cellToValue` accept BOTH shapes; the legacy ExcelJS
// reader is exposed as `readLegacyWorkbook` until Phase 2 removes it.
import ExcelJS from "exceljs";
import * as XLSX from "xlsx";

import type { SheetMeta } from "../../types";
import { HIGHLIGHT_UNMATCHED, autoWidth, styleHeader } from "./styling";
import { detectHeaderRow } from "@/lib/excel/header-detection";

export type CellValue = string | number | boolean | null;
export type SheetGrid = { headers: string[]; rows: CellValue[][] };

/** Raw cell value exposed by the adapter. Formula cells mirror ExcelJS's
 *  `{ formula, result }` shape so HyperFormula recalculation can reuse it. */
export type RawCellValue =
  | string
  | number
  | boolean
  | Date
  | null
  | { formula: string; result?: string | number | boolean | Date | null };

export interface CellHandle {
  readonly row: number; // 1-indexed
  readonly col: number; // 1-indexed
  readonly value: RawCellValue;
  readonly formula: string | null;
  readonly result: string | number | boolean | Date | null;
}

export interface RowHandle {
  readonly number: number; // 1-indexed
  getCell(col1Indexed: number): CellHandle;
}

export type SheetVisibility = "visible" | "hidden" | "veryHidden";

export interface SheetHandle {
  readonly kind: "sheetjs-sheet";
  readonly name: string;
  readonly rowCount: number;
  readonly columnCount: number;
  readonly hidden: boolean;
  readonly visibility: SheetVisibility;
  readonly hasMergedCells: boolean;
  getRow(row1Indexed: number): RowHandle;
  getCell(row1Indexed: number, col1Indexed: number): CellHandle;
  /** Column widths in characters (1-indexed column -> width). */
  getColumnWidth(col1Indexed: number): number | undefined;
  /** Underlying SheetJS worksheet — boundary-internal, do not use in ops. */
  readonly raw: XLSX.WorkSheet;
}

export interface WorkbookHandle {
  readonly kind: "sheetjs";
  readonly worksheets: SheetHandle[];
  creator: string | undefined;
  created: Date | undefined;
  getWorksheet(name: string): SheetHandle | undefined;
  removeWorksheet(name: string): void;
  /** Move sheet to a new 0-based position in sheet order. */
  moveWorksheet(name: string, index: number): void;
  /** Underlying SheetJS workbook — boundary-internal, do not use in ops. */
  readonly raw: XLSX.WorkBook;
}

// ---------------------------------------------------------------------------
// SheetJS adapter implementation
// ---------------------------------------------------------------------------

function isHandle(wb: unknown): wb is WorkbookHandle {
  return !!wb && typeof wb === "object" && (wb as { kind?: string }).kind === "sheetjs";
}
function isSheetHandle(ws: unknown): ws is SheetHandle {
  return !!ws && typeof ws === "object" && (ws as { kind?: string }).kind === "sheetjs-sheet";
}

function rawFromSheetCell(c: XLSX.CellObject | undefined): RawCellValue {
  if (!c) return null;
  let v: string | number | boolean | Date | null;
  if (c.t === "z" || c.v === undefined) v = null;
  else if (c.t === "e") v = null; // error cells surface as empty, matching cellToValue
  else v = c.v as string | number | boolean | Date;
  if (c.f) return { formula: c.f, result: v };
  return v;
}

function makeCell(ws: XLSX.WorkSheet, r: number, c: number): CellHandle {
  const safeR = Math.max(1, Math.floor(r));
  const safeC = Math.max(1, Math.floor(c));
  const addr = XLSX.utils.encode_cell({ r: safeR - 1, c: safeC - 1 });
  const cell = (ws as Record<string, XLSX.CellObject | undefined>)[addr];
  const value = rawFromSheetCell(cell);
  const isF = !!value && typeof value === "object" && !(value instanceof Date);
  return {
    row: safeR,
    col: safeC,
    value,
    formula: isF ? (value as { formula: string }).formula : null,
    result: isF ? ((value as { result?: CellHandle["result"] }).result ?? null) : (value as CellHandle["result"]),
  };
}

function makeSheet(book: XLSX.WorkBook, name: string): SheetHandle {
  const ws = book.Sheets[name];
  const range = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : null;
  const meta = () => {
    const i = book.SheetNames.indexOf(name);
    return book.Workbook?.Sheets?.[i]?.Hidden ?? 0;
  };
  return {
    kind: "sheetjs-sheet",
    name,
    raw: ws,
    // Match ExcelJS: last used row/col number (1-indexed), 0 when empty.
    get rowCount() {
      return range ? range.e.r + 1 : 0;
    },
    get columnCount() {
      return range ? range.e.c + 1 : 0;
    },
    get visibility(): SheetVisibility {
      const h = meta();
      return h === 2 ? "veryHidden" : h === 1 ? "hidden" : "visible";
    },
    get hidden() {
      return meta() !== 0;
    },
    get hasMergedCells() {
      return (ws["!merges"]?.length ?? 0) > 0;
    },
    getRow(r) {
      return { number: r, getCell: (c) => makeCell(ws, r, c) };
    },
    getCell(r, c) {
      return makeCell(ws, r, c);
    },
    getColumnWidth(c) {
      const col = ws["!cols"]?.[c - 1];
      return col?.wch ?? col?.width;
    },
  };
}

function wrap(book: XLSX.WorkBook): WorkbookHandle {
  book.Props = book.Props ?? {};
  return {
    kind: "sheetjs",
    raw: book,
    get worksheets() {
      return book.SheetNames.map((n) => makeSheet(book, n));
    },
    get creator() {
      return book.Props?.Author;
    },
    set creator(v) {
      book.Props = { ...(book.Props ?? {}), Author: v };
    },
    get created() {
      const d = book.Props?.CreatedDate;
      return d ? new Date(d) : undefined;
    },
    set created(v) {
      book.Props = { ...(book.Props ?? {}), CreatedDate: v };
    },
    getWorksheet(name) {
      return book.Sheets[name] ? makeSheet(book, name) : undefined;
    },
    removeWorksheet(name) {
      const i = book.SheetNames.indexOf(name);
      if (i < 0) return;
      book.SheetNames.splice(i, 1);
      delete book.Sheets[name];
      book.Workbook?.Sheets?.splice(i, 1);
    },
    moveWorksheet(name, index) {
      const i = book.SheetNames.indexOf(name);
      if (i < 0) return;
      const [n] = book.SheetNames.splice(i, 1);
      const target = Math.max(0, Math.min(index, book.SheetNames.length));
      book.SheetNames.splice(target, 0, n);
      const metas = book.Workbook?.Sheets;
      if (metas && metas.length > i) {
        const [m] = metas.splice(i, 1);
        metas.splice(target, 0, m);
      }
    },
  };
}

/** Read XLSX bytes with SheetJS, preserving formulas and dates. */
export async function readWorkbook(buffer: ArrayBuffer | Uint8Array): Promise<WorkbookHandle> {
  const data = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  const book = XLSX.read(data, {
    type: "array",
    cellFormula: true,
    cellDates: true,
    cellStyles: true, // needed for column widths (!cols) and hidden rows/cols
    bookVBA: false,
  });
  return wrap(book);
}

export function createWorkbook(): WorkbookHandle {
  const book = XLSX.utils.book_new();
  book.Workbook = { Sheets: [] };
  return wrap(book);
}

/** Deep, independent copy (mutations on the clone never touch the source). */
export function cloneWorkbook(wb: WorkbookHandle): WorkbookHandle {
  return wrap(structuredClone(wb.raw));
}

/** Serialize to XLSX bytes. */
export async function writeWorkbook(wb: WorkbookHandle): Promise<Buffer> {
  const out = XLSX.write(wb.raw, { type: "array", bookType: "xlsx", compression: true }) as ArrayBuffer;
  return Buffer.from(new Uint8Array(out));
}

/** Legacy ExcelJS reader for not-yet-migrated callers (Phase 2 removes). */
export async function readLegacyWorkbook(buffer: ArrayBuffer): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

// ---------------------------------------------------------------------------
// Shape-agnostic helpers (accept SheetJS handles or legacy ExcelJS objects)
// ---------------------------------------------------------------------------

export function cellToValue(v: ExcelJS.CellValue | RawCellValue): CellValue {
  if (v == null) return null;
  if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") return v;
  if (v instanceof Date) return v.toISOString();
  if (typeof v === "object") {
    if ("text" in v && typeof v.text === "string") return v.text;
    if ("richText" in v && Array.isArray(v.richText)) return v.richText.map((r) => r.text).join("");
    if ("result" in v && v.result !== undefined) return cellToValue(v.result as ExcelJS.CellValue);
    if ("formula" in v) return null;
  }
  return String(v);
}

type AnySheet = ExcelJS.Worksheet | SheetHandle;
type AnyWorkbook = ExcelJS.Workbook | WorkbookHandle;

function cellAt(ws: AnySheet, r: number, c: number): ExcelJS.CellValue | RawCellValue {
  return isSheetHandle(ws) ? ws.getCell(r, c).value : ws.getRow(r).getCell(c).value;
}

export function extractSheetMeta(wb: AnyWorkbook): SheetMeta {
  const list: AnySheet[] = wb.worksheets as AnySheet[];
  const sheets = list.map((ws) => {
    const headers: string[] = [];
    for (let c = 1; c <= ws.columnCount; c++) {
      const v = cellAt(ws, 1, c);
      if (v == null || v === "") continue;
      headers[c - 1] = String(isSheetHandle(ws) ? cellToValue(v as RawCellValue) ?? "" : v).trim();
    }
    const sampleRows: CellValue[][] = [];
    for (let r = 2; r <= Math.min(4, ws.rowCount); r++) {
      const out: CellValue[] = [];
      for (let c = 1; c <= headers.length; c++) out.push(cellToValue(cellAt(ws, r, c)));
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

export function sheetToGrid(ws: AnySheet): SheetGrid {
  // SheetHandle is structurally compatible with what detectHeaderRow reads
  // (rowCount, columnCount, getRow(r).getCell(c).value).
  const headerRow = detectHeaderRow(ws as unknown as ExcelJS.Worksheet);
  const maxCol = ws.columnCount;
  const headers: string[] = [];
  for (let c = 1; c <= maxCol; c++) {
    const v = cellAt(ws, headerRow, c);
    headers.push(String((isSheetHandle(ws) ? cellToValue(v as RawCellValue) : v) ?? "").trim());
  }
  const rows: CellValue[][] = [];
  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const out: CellValue[] = [];
    let any = false;
    for (let c = 1; c <= maxCol; c++) {
      const v = cellToValue(cellAt(ws, r, c));
      if (v !== null && v !== "") any = true;
      out.push(v);
    }
    if (any) rows.push(out);
  }
  return { headers, rows };
}

function columnWidths(headers: string[], rows: CellValue[][]): XLSX.ColInfo[] {
  // Same rule as styling.autoWidth.
  return headers.map((h, i) => {
    const dataMax = rows.reduce((m, r) => Math.max(m, String(r[i] ?? "").length), 0);
    return { wch: Math.min(Math.max(h.length + 2, dataMax + 2, 12), 42) };
  });
}

export function writeGridToSheet(
  wb: WorkbookHandle,
  sheetName: string,
  grid: SheetGrid,
  opts?: { highlightRows?: Set<number>; headerStyle?: boolean },
): SheetHandle;
export function writeGridToSheet(
  wb: ExcelJS.Workbook,
  sheetName: string,
  grid: SheetGrid,
  opts?: { highlightRows?: Set<number>; headerStyle?: boolean },
): ExcelJS.Worksheet;
export function writeGridToSheet(
  wb: AnyWorkbook,
  sheetName: string,
  grid: SheetGrid,
  opts: { highlightRows?: Set<number>; headerStyle?: boolean } = {},
): SheetHandle | ExcelJS.Worksheet {
  if (isHandle(wb)) {
    // Excel hard limit 31 chars + forbidden characters. Name already produced
    // by callers is kept when valid, so existing output naming is unchanged.
    const name = sheetName.replace(/[\\/:?*[\]]/g, " ").slice(0, 31) || "Sheet";
    wb.removeWorksheet(name);
    const ws = XLSX.utils.aoa_to_sheet([grid.headers, ...grid.rows]);
    ws["!cols"] = columnWidths(grid.headers, grid.rows);
    // Styling (header fill, highlight fill, frozen pane) is not supported by
    // SheetJS CE and is intentionally out of scope for Phase 1.
    XLSX.utils.book_append_sheet(wb.raw, ws, name);
    const book = wb.raw;
    book.Workbook = book.Workbook ?? {};
    book.Workbook.Sheets = book.Workbook.Sheets ?? [];
    while (book.Workbook.Sheets.length < book.SheetNames.length) book.Workbook.Sheets.push({});
    book.Workbook.Sheets[book.SheetNames.indexOf(name)] = { name, Hidden: 0 };
    return makeSheet(book, name);
  }

  const existing = wb.getWorksheet(sheetName);
  if (existing) wb.removeWorksheet(existing.id);
  const ws = wb.addWorksheet(sheetName, { views: [{ state: "frozen", ySplit: 1 }] });
  ws.addRow(grid.headers);
  grid.rows.forEach((r) => ws.addRow(r));
  if (opts.headerStyle !== false) styleHeader(ws);
  autoWidth(ws, grid.headers, grid.rows);
  if (opts.highlightRows) {
    opts.highlightRows.forEach((rowIdx) => {
      ws.getRow(rowIdx + 2).eachCell((cell) => {
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
