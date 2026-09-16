import { describe, expect, it } from "bun:test";
import { resolveKeyColumn, resolveKeyAcrossFiles } from "../src/lib/excel/engine/shared/key-resolution";

describe("resolveKeyColumn", () => {
  it("prefers an exact normalized match over any fuzzy candidate", () => {
    const r = resolveKeyColumn(["Old Student ID", "Student ID"], "student id");
    expect(r).toMatchObject({ status: "resolved", header: "Student ID", confidence: 100 });
  });

  it("does not match across token boundaries", () => {
    expect(resolveKeyColumn(["Filename", "Surname"], "Name").status).toBe("not_found");
    expect(resolveKeyColumn(["Valid Until", "Grid Ref"], "ID").status).toBe("not_found");
  });

  it("penalises extra tokens instead of taking header order", () => {
    const r = resolveKeyColumn(["Reg No Verified Flag", "Reg No"], "Reg No");
    expect(r).toMatchObject({ status: "resolved", header: "Reg No" });
  });

  it("reports ambiguity when two headers are equally plausible", () => {
    const r = resolveKeyColumn(["Old Student ID", "New Student ID"], "Student ID");
    expect(r.status).toBe("ambiguous");
  });

  it("returns not_found for an unknown column", () => {
    expect(resolveKeyColumn(["Name", "Email"], "Invoice Number").status).toBe("not_found");
  });
});

describe("resolveKeyAcrossFiles", () => {
  it("accepts files that agree", () => {
    const out = resolveKeyAcrossFiles(
      [["Name", "Reg No"], ["Reg No", "Email"]],
      "Reg No",
      ["a.xlsx", "b.xlsx"],
    );
    expect(out.map((o) => o.index)).toEqual([1, 0]);
  });

  it("rejects files that resolve to different columns", () => {
    expect(() =>
      resolveKeyAcrossFiles(
        [["Student ID"], ["Old Student ID"]],
        "Student ID",
        ["a.xlsx", "b.xlsx"],
      ),
    ).toThrow(/different columns/);
  });
});
