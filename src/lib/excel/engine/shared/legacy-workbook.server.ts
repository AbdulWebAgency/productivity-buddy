// TEMPORARY legacy ExcelJS reader. No runtime callers remain; kept isolated
// here (not re-exported) until the final cleanup pass removes it together
// with the `exceljs` dependency. Do not import from normal runtime code.
import ExcelJS from "exceljs";

export async function readLegacyWorkbook(buffer: ArrayBuffer): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}
