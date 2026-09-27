import { describe, expect, test } from "bun:test";
import * as XLSX from "xlsx";
import {
  cloneWorkbook,
  createWorkbook,
  extractSheetMeta,
  readWorkbook,
  safeSheetName,
  sheetToGrid,
  writeGridToSheet,
  writeWorkbook,
} from "../src/lib/excel/engine/shared/workbook";

function buildFixture(): ArrayBuffer {
  const book = XLSX.utils.book_new();
  const data = XLSX.utils.aoa_to_sheet([
    ["Name", "Reg", "Score"],
    ["Ana", "R1", 10],
    ["Ben", "R2", 20],
  ]);
  data["C4"] = { t: "n", v: 30, f: "SUM(C2:C3)" };
  data["E6"] = { t: "s", v: "far" };
  data["!ref"] = "A1:E6";
  data["!merges"] = [XLSX.utils.decode_range("A8:B8")];
  XLSX.utils.book_append_sheet(book, data, "Data");
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([["x"]]), "Secret");
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([["y"]]), "Deep");
  book.Workbook = { Sheets: [{ Hidden: 0 }, { Hidden: 1 }, { Hidden: 2 }] };
  book.Props = { Author: "Tester", CreatedDate: new Date("2024-01-02T00:00:00Z") };
  const out = XLSX.write(book, { type: "array", bookType: "xlsx" }) as ArrayBuffer;
  return out;
}

describe("workbook boundary", () => {
  test("readWorkbook: sheets, order, metadata", async () => {
    const wb = await readWorkbook(buildFixture());
    expect(wb.worksheets.map((s) => s.name)).toEqual(["Data", "Secret", "Deep"]);
    expect(wb.creator).toBe("Tester");
    expect(wb.created?.toISOString().slice(0, 10)).toBe("2024-01-02");
  });

  test("1-indexed getRow/getCell, counts, sparse cells", async () => {
    const ws = (await readWorkbook(buildFixture())).getWorksheet("Data")!;
    expect(ws.getCell(1, 1).value).toBe("Name");
    expect(ws.getRow(2).getCell(2).value).toBe("R1");
    expect(ws.getCell(1, 1).row).toBe(1);
    expect(ws.rowCount).toBe(6);
    expect(ws.columnCount).toBe(5);
    expect(ws.getCell(5, 4).value).toBeNull();
    expect(ws.getCell(500, 500).value).toBeNull();
    expect(ws.getCell(0, 0).value).toBe("Name"); // clamped, never throws
  });

  test("hidden / veryHidden and merged detection", async () => {
    const wb = await readWorkbook(buildFixture());
    expect(wb.getWorksheet("Data")!.hidden).toBe(false);
    expect(wb.getWorksheet("Secret")!.visibility).toBe("hidden");
    expect(wb.getWorksheet("Deep")!.visibility).toBe("veryHidden");
    expect(wb.getWorksheet("Data")!.hasMergedCells).toBe(true);
    expect(wb.getWorksheet("Secret")!.hasMergedCells).toBe(false);
  });

  test("formula preserved with result", async () => {
    const c = (await readWorkbook(buildFixture())).getWorksheet("Data")!.getCell(4, 3);
    expect(c.formula).toBe("SUM(C2:C3)");
    expect(c.result).toBe(30);
    expect(c.value).toEqual({ formula: "SUM(C2:C3)", result: 30 });
  });

  test("sheetToGrid and extractSheetMeta on handles", async () => {
    const wb = await readWorkbook(buildFixture());
    const grid = sheetToGrid(wb.getWorksheet("Data")!);
    expect(grid.headers.slice(0, 3)).toEqual(["Name", "Reg", "Score"]);
    expect(grid.rows[0].slice(0, 3)).toEqual(["Ana", "R1", 10]);
    expect(grid.rows[2][2]).toBe(30);
    const meta = extractSheetMeta(wb);
    expect(meta.sheets[0].headers).toEqual(["Name", "Reg", "Score"]);
  });

  test("createWorkbook + writeGridToSheet + widths + round trip", async () => {
    const wb = createWorkbook();
    wb.creator = "Productivity Buddy";
    writeGridToSheet(wb, "Result", { headers: ["Id", "Label"], rows: [[1, "a very long label here"], [2, null]] });
    writeGridToSheet(wb, "Result", { headers: ["Id"], rows: [[9]] }); // replace
    const ws = writeGridToSheet(wb, "Other", { headers: ["Name", "X"], rows: [["abcdefghijklmnopqrstuvwxyz", 1]] });
    expect(wb.worksheets.map((s) => s.name)).toEqual(["Result", "Other"]);
    expect(ws.getColumnWidth(1)).toBe(28);
    expect(ws.getColumnWidth(2)).toBe(12);

    const back = await readWorkbook(await writeWorkbook(wb));
    expect(back.worksheets.map((s) => s.name)).toEqual(["Result", "Other"]);
    expect(back.getWorksheet("Result")!.getCell(2, 1).value).toBe(9);
    expect(back.getWorksheet("Other")!.getColumnWidth(1)).toBeCloseTo(28, 0);
    expect(back.creator).toBe("Productivity Buddy");
  });

  test("moveWorksheet and removeWorksheet", () => {
    const wb = createWorkbook();
    writeGridToSheet(wb, "A", { headers: ["h"], rows: [] });
    writeGridToSheet(wb, "B", { headers: ["h"], rows: [] });
    wb.moveWorksheet("B", 0);
    expect(wb.worksheets.map((s) => s.name)).toEqual(["B", "A"]);
    wb.removeWorksheet("B");
    expect(wb.worksheets.map((s) => s.name)).toEqual(["A"]);
  });

  test("safe sheet names", () => {
    expect(safeSheetName("Only in data/2024?.xlsx")).toBe("Only in data 2024 ");
    expect(safeSheetName("x".repeat(40)).length).toBe(25);
    const wb = createWorkbook();
    const ws = writeGridToSheet(wb, "a:b*c[" + "z".repeat(40), { headers: ["h"], rows: [] });
    expect(ws.name.length).toBeLessThanOrEqual(31);
    expect(ws.name).not.toMatch(/[\\/:?*[\]]/);
  });

  test("cloneWorkbook is independent", async () => {
    const wb = await readWorkbook(buildFixture());
    const copy = cloneWorkbook(wb);
    writeGridToSheet(copy, "New", { headers: ["h"], rows: [] });
    copy.removeWorksheet("Secret");
    expect(wb.worksheets.map((s) => s.name)).toEqual(["Data", "Secret", "Deep"]);
    expect(copy.worksheets.map((s) => s.name)).toEqual(["Data", "Deep", "New"]);
    expect(copy.getWorksheet("Deep")!.visibility).toBe("veryHidden");
  });
});
