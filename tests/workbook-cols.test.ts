import { describe, expect, test } from "bun:test";
import * as XLSX from "xlsx";
import { readWorkbook, stripZeroOutlineLevels, writeWorkbook } from "../src/lib/excel/engine/shared/workbook";

type Cols = XLSX.ColInfo[];

function bookWithCols(cols: Cols): XLSX.WorkBook {
  const ws = XLSX.utils.aoa_to_sheet([["A", "B"], [1, 2]]);
  ws["!cols"] = cols;
  const b = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(b, ws, "S");
  return b;
}
const colsOf = (b: XLSX.WorkBook): Cols => (b.Sheets.S["!cols"] ?? []) as Cols;
// Uncompressed zip, so the sheet XML is readable as plain text.
const storedXml = (b: XLSX.WorkBook) =>
  Buffer.from(XLSX.write(b, { type: "array", bookType: "xlsx", compression: false }) as ArrayBuffer).toString("latin1");

describe("zero outline levels are not written as an invalid <col level> attribute", () => {
  test("control: SheetJS itself echoes level on <col> when level is set", () => {
    expect(storedXml(bookWithCols([{ wch: 10, level: 0 }]))).toContain(' level="0"');
  });

  test("stripZeroOutlineLevels removes the attribute from the output XML", () => {
    const b = bookWithCols([
      { wch: 10, level: 0 },
      { wch: 12, level: 0 },
    ]);
    stripZeroOutlineLevels(b);
    expect(storedXml(b)).not.toContain(' level="');
    expect(colsOf(b)[0].wch).toBe(10); // widths untouched
  });

  test("real outline levels are left alone", () => {
    const b = bookWithCols([
      { wch: 10, level: 0 },
      { wch: 10, level: 2 },
    ]);
    stripZeroOutlineLevels(b);
    expect(colsOf(b)[0].level).toBeUndefined();
    expect(colsOf(b)[1].level).toBe(2);
  });

  test("a file read from disk and written back no longer carries level 0", async () => {
    const bytes = XLSX.write(bookWithCols([{ wch: 10, level: 0 }]), { type: "array", bookType: "xlsx" }) as ArrayBuffer;
    const h = await readWorkbook(bytes);
    expect((h.raw.Sheets.S["!cols"] as Cols)[0].level).toBe(0);
    await writeWorkbook(h);
    expect((h.raw.Sheets.S["!cols"] as Cols)[0].level).toBeUndefined();
  });
});