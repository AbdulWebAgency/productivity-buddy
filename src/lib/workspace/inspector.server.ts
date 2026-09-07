// Server-only Excel inspector: analyze a workbook beyond sheet meta.
import type ExcelJS from "exceljs";
import { detectHeaderRow } from "@/lib/excel/header-detection";
import { KEY_HINTS, normalizeHeader } from "@/lib/excel/engine/shared/headers";
import { classifySheet, type WorksheetType } from "./sheet-classifier";

export type InspectorSheetReport = {
  name: string;
  rows: number;
  columns: number;
  headers: string[];
  blankRows: number;
  formulaCells: number;
  likelyKeys: string[]; // header names likely to be primary keys
  duplicateKeyValues: number; // count of duplicate values in the top likely key
  // --- Sprint: worksheet understanding (additive, backward compatible) ---
  worksheetType: WorksheetType;
  confidence: number; // 0-100
  detectedHeaderRow: number; // 0 when header detection was skipped
  estimatedDataStartRow: number;
  estimatedDataEndRow: number;
  containsPivotIndicators: boolean;
  containsMergedCells: boolean;
  hidden: boolean;
  likelyPrimaryTable: boolean;
  rank: number; // ranking score for DATA sheets, 0 otherwise
};

export type InspectorReport = {
  sheets: InspectorSheetReport[];
  warnings: string[];
  primaryDataSheet: string | null;
};

// Inspector-specific scoring: sums matching hint lengths (weights compound
// matches) — kept distinct from the max-based scoreKey used by the planner.
function scoreHeader(h: string): number {
  const n = normalizeHeader(h);
  let s = 0;
  for (const hint of KEY_HINTS) if (n.includes(hint)) s += hint.length;
  return s;
}

export function inspectWorkbook(wb: ExcelJS.Workbook): InspectorReport {
  const warnings: string[] = [];

  const sheets: InspectorSheetReport[] = wb.worksheets.map((ws) => {
    const cls = classifySheet(ws);
    const analyzeHeaders = cls.worksheetType === "DATA" || cls.worksheetType === "UNKNOWN";

    const base = {
      name: ws.name,
      worksheetType: cls.worksheetType,
      confidence: cls.confidence,
      estimatedDataStartRow: cls.estimatedDataStartRow,
      estimatedDataEndRow: cls.estimatedDataEndRow,
      containsPivotIndicators: cls.containsPivotIndicators,
      containsMergedCells: cls.containsMergedCells,
      hidden: cls.hidden,
    };

    if (!analyzeHeaders) {
      // Skip expensive header/key detection for pivot, summary, doc and empty sheets.
      return {
        ...base,
        rows: Math.max(0, cls.estimatedDataEndRow - cls.estimatedDataStartRow),
        columns: 0,
        headers: [],
        blankRows: 0,
        formulaCells: 0,
        likelyKeys: [],
        duplicateKeyValues: 0,
        detectedHeaderRow: 0,
        likelyPrimaryTable: false,
        rank: 0,
      };
    }

    const headers: string[] = [];
    const headerRow = detectHeaderRow(ws);
    const first = ws.getRow(headerRow);
    for (let c = 1; c <= ws.columnCount; c++) {
      headers.push(String(first.getCell(c).value ?? "").trim());
    }
    let blankRows = 0;
    let formulaCells = 0;
    // Collect column values for uniqueness scoring
    const perColValues: Map<number, string[]> = new Map();
    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      let any = false;
      for (let c = 1; c <= headers.length; c++) {
        const cell = row.getCell(c);
        const v = cell.value;
        if (v != null && v !== "") any = true;
        if (v && typeof v === "object" && "formula" in v) formulaCells++;
        const s = v == null ? "" : typeof v === "object" ? JSON.stringify(v) : String(v);
        if (!perColValues.has(c)) perColValues.set(c, []);
        perColValues.get(c)!.push(s);
      }
      if (!any) blankRows++;
    }

    // Rank likely keys: high uniqueness + hint match
    const scored = headers
      .map((h, i) => {
        const vals = (perColValues.get(i + 1) ?? []).filter((v) => v !== "");
        const uniq = new Set(vals).size;
        const uniqueness = vals.length ? uniq / vals.length : 0;
        return { h, i, uniqueness, uniqCount: uniq, total: vals.length, hint: scoreHeader(h) };
      })
      .filter((x) => x.h)
      .sort((a, b) => b.hint * 10 + b.uniqueness * 5 - (a.hint * 10 + a.uniqueness * 5));
    const likelyKeys = scored
      .filter((x) => x.uniqueness >= 0.85 || x.hint > 0)
      .slice(0, 3)
      .map((x) => x.h);

    let duplicateKeyValues = 0;
    let keyUniqueness = 0;
    if (likelyKeys[0]) {
      const idx = headers.indexOf(likelyKeys[0]) + 1;
      const vals = (perColValues.get(idx) ?? []).filter((v) => v !== "");
      duplicateKeyValues = vals.length - new Set(vals).size;
      keyUniqueness = vals.length ? new Set(vals).size / vals.length : 0;
    }

    if (blankRows > 0) warnings.push(`${ws.name}: ${blankRows} blank rows`);
    if (duplicateKeyValues > 0 && likelyKeys[0]) {
      warnings.push(`${ws.name}: ${duplicateKeyValues} duplicate values in "${likelyKeys[0]}"`);
    }

    const rows = Math.max(0, ws.rowCount - headerRow);
    const meaningfulHeaders = headers.filter(Boolean).length;
    const blankRatio = rows > 0 ? blankRows / rows : 1;
    const rank =
      cls.worksheetType === "DATA"
        ? Math.round(
            Math.min(60, Math.log10(Math.max(1, rows)) * 20) + // record volume
              Math.min(15, meaningfulHeaders * 1.5) + // meaningful headers
              keyUniqueness * 20 + // primary key confidence
              (cls.confidence / 100) * 10 - // tabular consistency
              blankRatio * 20, // blank-row penalty
          )
        : 0;

    return {
      ...base,
      rows,
      columns: meaningfulHeaders,
      headers: headers.filter(Boolean),
      blankRows,
      formulaCells,
      likelyKeys,
      duplicateKeyValues,
      detectedHeaderRow: headerRow,
      likelyPrimaryTable: false,
      rank,
    };
  });

  // Rank DATA sheets and mark the recommended primary sheet.
  const dataSheets = sheets.filter((s) => s.worksheetType === "DATA" && !s.hidden);
  const ranked = [...dataSheets].sort((a, b) => b.rank - a.rank);
  const primary = ranked[0] ?? sheets.find((s) => s.worksheetType === "UNKNOWN" && s.headers.length > 0) ?? null;
  if (primary) primary.likelyPrimaryTable = true;

  const nonData = sheets.filter((s) => s.worksheetType === "PIVOT" || s.worksheetType === "SUMMARY");
  if (primary && nonData.length) {
    warnings.push(
      `Using "${primary.name}" as the data sheet; ignoring ${nonData.map((s) => `"${s.name}" (${s.worksheetType.toLowerCase()})`).join(", ")}`,
    );
  }

  return { sheets, warnings, primaryDataSheet: primary?.name ?? null };
}
