import { describe, expect, test } from "bun:test";
import * as XLSX from "xlsx";
import {
  createWorkbook,
  extractSheetMeta,
  readWorkbook,
  sheetToGrid,
  writeGridToSheet,
  writeWorkbook,
} from "../src/lib/excel/engine/shared/workbook";
import { detectHeaderRow } from "../src/lib/excel/header-detection";
import { classifySheet } from "../src/lib/workspace/sheet-classifier";
import { inspectWorkbook } from "../src/lib/workspace/inspector.server";

function book(sheets: Record<string, unknown[][]>, hidden: number[] = []): ArrayBuffer {
  const b = XLSX.utils.book_new();
  for (const [n, aoa] of Object.entries(sheets)) XLSX.utils.book_append_sheet(b, XLSX.utils.aoa_to_sheet(aoa), n);
  b.Workbook = { Sheets: Object.keys(sheets).map((_, i) => ({ Hidden: (hidden[i] ?? 0) as 0 | 1 | 2 })) };
  return XLSX.write(b, { type: "array", bookType: "xlsx" }) as ArrayBuffer;
}

const titled = [
  ["Student Report 2024"],
  [],
  ["Name", "Reg No", "Email", "Score"],
  ...Array.from({ length: 30 }, (_, i) => [`S${i}`, `R${i}`, `s${i}@x.io`, i]),
];

describe("SheetJS consumers", () => {
  test("header detection skips title rows through SheetHandle", async () => {
    const ws = (await readWorkbook(book({ Data: titled }))).worksheets[0];
    expect(detectHeaderRow(ws)).toBe(3);
  });

  test("header detection fallback is row 1 when no row has 2+ cells", async () => {
    const ws = (await readWorkbook(book({ Data: [["only"], ["one"]] }))).worksheets[0];
    expect(detectHeaderRow(ws)).toBe(1);
  });

  test("sheetToGrid uses detected header row", async () => {
    const g = sheetToGrid((await readWorkbook(book({ Data: titled }))).worksheets[0]);
    expect(g.headers).toEqual(["Name", "Reg No", "Email", "Score"]);
    expect(g.rows.length).toBe(30);
    expect(g.rows[0]).toEqual(["S0", "R0", "s0@x.io", 0]);
  });

  test("extractSheetMeta (registration path) via readWorkbook", async () => {
    const wb = await readWorkbook(book({ A: [["Id", "Name"], [1, "x"], [2, "y"]], B: [["k"]] }));
    const meta = extractSheetMeta(wb);
    expect(meta.sheets.map((s) => s.name)).toEqual(["A", "B"]);
    expect(meta.sheets[0]).toEqual({ name: "A", rowCount: 3, columnCount: 2, headers: ["Id", "Name"], sampleRows: [[1, "x"], [2, "y"]] });
  });

  test("classifier + inspector on SheetJS handles", async () => {
    const wb = await readWorkbook(
      book(
        {
          Pivot: [["Row Labels", "Sum of Score"], ["A", 10], ["Grand Total", 10]],
          Data: titled,
          Hidden: titled,
        },
        [0, 0, 1],
      ),
    );
    expect(classifySheet(wb.getWorksheet("Pivot")!).worksheetType).toBe("PIVOT");
    expect(classifySheet(wb.getWorksheet("Data")!).worksheetType).toBe("DATA");
    expect(classifySheet(wb.getWorksheet("Hidden")!).hidden).toBe(true);
    const rep = inspectWorkbook(wb);
    expect(rep.primaryDataSheet).toBe("Data");
    const data = rep.sheets.find((s) => s.name === "Data")!;
    expect(data.detectedHeaderRow).toBe(3);
    expect(data.headers).toEqual(["Name", "Reg No", "Email", "Score"]);
  });

  test("writeGridToSheet round trip keeps data, order and widths", async () => {
    const wb = createWorkbook();
    writeGridToSheet(wb, "Out", { headers: ["Name", "N"], rows: [["abc", 1]] }, { highlightRows: new Set([0]) });
    const back = await readWorkbook(await writeWorkbook(wb));
    expect(sheetToGrid(back.worksheets[0])).toEqual({ headers: ["Name", "N"], rows: [["abc", 1]] });
    expect(back.worksheets[0].getColumnWidth(1)).toBeCloseTo(12, 0);
  });

  test("SheetJS is the only spreadsheet I/O dependency", async () => {
    const pkg = await Bun.file("package.json").json();
    const deps = Object.keys({ ...pkg.dependencies, ...pkg.devDependencies });
    expect(deps).toContain("xlsx");
    expect(deps.filter((d) => /^excel/i.test(d))).toEqual([]);
  });
});
