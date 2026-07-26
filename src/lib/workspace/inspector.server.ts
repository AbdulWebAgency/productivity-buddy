// Server-only Excel inspector: analyze a workbook beyond sheet meta.
import type ExcelJS from "exceljs";
import { detectHeaderRow } from "@/lib/excel/header-detection";
import { KEY_HINTS, normalizeHeader } from "@/lib/excel/engine/shared/headers";

export type InspectorReport = {
  sheets: {
    name: string;
    rows: number;
    columns: number;
    headers: string[];
    blankRows: number;
    formulaCells: number;
    likelyKeys: string[]; // header names likely to be primary keys
    duplicateKeyValues: number; // count of duplicate values in the top likely key
  }[];
  warnings: string[];
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
  const sheets = wb.worksheets.map((ws) => {
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
    if (likelyKeys[0]) {
      const idx = headers.indexOf(likelyKeys[0]) + 1;
      const vals = (perColValues.get(idx) ?? []).filter((v) => v !== "");
      duplicateKeyValues = vals.length - new Set(vals).size;
    }

    if (blankRows > 0) warnings.push(`${ws.name}: ${blankRows} blank rows`);
    if (duplicateKeyValues > 0 && likelyKeys[0]) {
      warnings.push(`${ws.name}: ${duplicateKeyValues} duplicate values in "${likelyKeys[0]}"`);
    }

    return {
      name: ws.name,
      rows: Math.max(0, ws.rowCount - headerRow),
      columns: headers.filter(Boolean).length,
      headers: headers.filter(Boolean),
      blankRows,
      formulaCells,
      likelyKeys,
      duplicateKeyValues,
    };
  });
  return { sheets, warnings };
}
