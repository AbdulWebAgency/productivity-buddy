// Deterministic worksheet classification. No AI.
// Used by the inspector to understand what each worksheet actually is
// before headers / keys are detected.
import type ExcelJS from "exceljs";

export type WorksheetType = "DATA" | "PIVOT" | "SUMMARY" | "DOCUMENTATION" | "EMPTY" | "UNKNOWN";

export type SheetClassification = {
  worksheetType: WorksheetType;
  confidence: number; // 0-100
  containsPivotIndicators: boolean;
  containsMergedCells: boolean;
  hidden: boolean;
  estimatedDataStartRow: number;
  estimatedDataEndRow: number;
  reasons: string[];
};

const PIVOT_LABELS = [
  "grand total",
  "row labels",
  "column labels",
  "count of",
  "sum of",
  "average of",
  "min of",
  "max of",
  "values",
  "pivottable",
];

const SUMMARY_LABELS = ["total", "subtotal", "summary", "kpi", "average", "overall", "percentage", "%"];

const DOC_LABELS = [
  "instruction",
  "instructions",
  "readme",
  "read me",
  "notes",
  "note",
  "legend",
  "definitions",
  "glossary",
  "about",
  "how to",
  "documentation",
  "changelog",
  "disclaimer",
];

const SCAN_ROWS = 60;

function cellText(v: ExcelJS.CellValue): string {
  if (v == null) return "";
  if (typeof v === "string") return v;
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  if (v instanceof Date) return v.toISOString();
  if (typeof v === "object") {
    const o = v as unknown as Record<string, unknown>;
    if (typeof o.text === "string") return o.text;
    if (Array.isArray(o.richText)) return o.richText.map((r: { text?: string }) => r.text ?? "").join("");
    if ("result" in o) return cellText(o.result as ExcelJS.CellValue);
    if (typeof o.formula === "string") return "";
  }
  return "";
}

type RowStat = { row: number; filled: number; text: number; numeric: number; firstCol: number; lastCol: number };

export function classifySheet(ws: ExcelJS.Worksheet): SheetClassification {
  const reasons: string[] = [];
  const hidden = ws.state === "hidden" || ws.state === "veryHidden";

  const maxRow = Math.min(ws.rowCount, SCAN_ROWS);
  const maxCol = Math.max(1, Math.min(ws.columnCount, 80));

  const stats: RowStat[] = [];
  let pivotHits = 0;
  let summaryHits = 0;
  let docHits = 0;
  let longTextCells = 0;
  let totalFilled = 0;

  for (let r = 1; r <= maxRow; r++) {
    const row = ws.getRow(r);
    let filled = 0;
    let text = 0;
    let numeric = 0;
    let firstCol = 0;
    let lastCol = 0;
    for (let c = 1; c <= maxCol; c++) {
      const raw = row.getCell(c).value;
      const s = cellText(raw).trim();
      if (s === "") continue;
      filled++;
      if (!firstCol) firstCol = c;
      lastCol = c;
      const lower = s.toLowerCase();
      if (typeof raw === "number") numeric++;
      else {
        text++;
        if (s.length > 80) longTextCells++;
      }
      if (PIVOT_LABELS.some((p) => lower.startsWith(p) || lower === p)) pivotHits++;
      if (SUMMARY_LABELS.some((p) => lower === p || lower.startsWith(p + " "))) summaryHits++;
      if (DOC_LABELS.some((p) => lower.includes(p))) docHits++;
    }
    totalFilled += filled;
    if (filled > 0) stats.push({ row: r, filled, text, numeric, firstCol, lastCol });
  }

  const containsMergedCells = (() => {
    const m = (ws as unknown as { model?: { merges?: unknown[] } }).model?.merges;
    return Array.isArray(m) && m.length > 0;
  })();

  const containsPivotIndicators = pivotHits > 0;

  // EMPTY
  if (stats.length === 0 || totalFilled < 3) {
    return {
      worksheetType: "EMPTY",
      confidence: 95,
      containsPivotIndicators,
      containsMergedCells,
      hidden,
      estimatedDataStartRow: 0,
      estimatedDataEndRow: 0,
      reasons: ["almost no usable content"],
    };
  }

  // Tabular consistency: modal filled-cell width across populated rows
  const widthCounts = new Map<number, number>();
  stats.forEach((s) => widthCounts.set(s.filled, (widthCounts.get(s.filled) ?? 0) + 1));
  let modalWidth = 1;
  let modalCount = 0;
  widthCounts.forEach((count, width) => {
    if (count > modalCount || (count === modalCount && width > modalWidth)) {
      modalCount = count;
      modalWidth = width;
    }
  });
  const consistency = modalCount / stats.length; // 0..1

  // Estimated data band: the contiguous run of rows matching the modal width
  const bandRows = stats.filter((s) => Math.abs(s.filled - modalWidth) <= 1);
  const estimatedDataStartRow = bandRows[0]?.row ?? stats[0].row;
  const lastBand = bandRows[bandRows.length - 1]?.row ?? stats[stats.length - 1].row;
  // extend the end row to the real sheet end when the band runs to the scan limit
  const estimatedDataEndRow = lastBand >= maxRow ? ws.rowCount : lastBand;

  const recordRows = Math.max(0, (estimatedDataEndRow - estimatedDataStartRow));

  // PIVOT
  if (pivotHits >= 2 || (pivotHits === 1 && modalWidth <= 4)) {
    reasons.push(`pivot labels found (${pivotHits})`);
    return {
      worksheetType: "PIVOT",
      confidence: Math.min(95, 60 + pivotHits * 10),
      containsPivotIndicators,
      containsMergedCells,
      hidden,
      estimatedDataStartRow,
      estimatedDataEndRow,
      reasons,
    };
  }

  // DOCUMENTATION
  const docNameHit = DOC_LABELS.some((p) => ws.name.toLowerCase().includes(p));
  if (docNameHit || longTextCells >= 3 || (modalWidth <= 2 && stats.length <= 30 && docHits > 0)) {
    reasons.push(docNameHit ? "sheet name looks like documentation" : "long free-text content, narrow layout");
    return {
      worksheetType: "DOCUMENTATION",
      confidence: docNameHit ? 85 : 70,
      containsPivotIndicators,
      containsMergedCells,
      hidden,
      estimatedDataStartRow,
      estimatedDataEndRow,
      reasons,
    };
  }

  // SUMMARY: few rows, mostly totals/stats
  const summaryNameHit = /summary|overview|dashboard|report/i.test(ws.name);
  if ((summaryHits >= 2 && recordRows < 15) || (summaryNameHit && recordRows < 25)) {
    reasons.push(summaryNameHit ? "sheet name looks like a summary" : `aggregate labels found (${summaryHits})`);
    return {
      worksheetType: "SUMMARY",
      confidence: 75,
      containsPivotIndicators,
      containsMergedCells,
      hidden,
      estimatedDataStartRow,
      estimatedDataEndRow,
      reasons,
    };
  }

  // DATA: repeated row structure, several columns, real records
  if (modalWidth >= 2 && consistency >= 0.5 && recordRows >= 2) {
    const confidence = Math.round(
      Math.min(98, 40 + consistency * 35 + Math.min(modalWidth, 10) * 1.5 + Math.min(recordRows, 50) * 0.3),
    );
    reasons.push(`${modalWidth} consistent columns, ~${recordRows} records`);
    return {
      worksheetType: "DATA",
      confidence,
      containsPivotIndicators,
      containsMergedCells,
      hidden,
      estimatedDataStartRow,
      estimatedDataEndRow,
      reasons,
    };
  }

  reasons.push("layout does not match a known worksheet shape");
  return {
    worksheetType: "UNKNOWN",
    confidence: 30,
    containsPivotIndicators,
    containsMergedCells,
    hidden,
    estimatedDataStartRow,
    estimatedDataEndRow,
    reasons,
  };
}
