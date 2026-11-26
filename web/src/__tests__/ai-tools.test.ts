import { describe, it, expect } from "vitest";
import {
  MAX_CELLS, parseA1Range, rectToA1, usedRange, findSheetIndex, checkSheetName,
  readRange, writeCells, clearCells,
} from "../lib/ai-tools";
import { makeSheet } from "../lib/persistence";
import type { Sheet, WorkbookData } from "../types";

const sheetWith = (cells: Record<string, string>, name = "Sheet 1"): Sheet => ({ ...makeSheet(name), cells });

describe("parseA1Range", () => {
  it("single cell", () => expect(parseA1Range("B3")).toEqual({ minC: 1, maxC: 1, minR: 2, maxR: 2 }));
  it("range", () => expect(parseA1Range("A1:C10")).toEqual({ minC: 0, maxC: 2, minR: 0, maxR: 9 }));
  it("reversed corners are normalised", () => expect(parseA1Range("C10:A1")).toEqual({ minC: 0, maxC: 2, minR: 0, maxR: 9 }));
  it("lower case and $ and spaces", () => expect(parseA1Range(" $a$1 : b2 ")).toEqual({ minC: 0, maxC: 1, minR: 0, maxR: 1 }));
  it("multi-letter columns", () => expect(parseA1Range("AA1")).toEqual({ minC: 26, maxC: 26, minR: 0, maxR: 0 }));
  it("rejects garbage", () => expect(() => parseA1Range("hello")).toThrow(/not an A1 reference/));
  it("rejects row 0", () => expect(() => parseA1Range("A0")).toThrow(/not an A1 reference/));
  it("rejects empty", () => expect(() => parseA1Range("")).toThrow(/A1 notation/));
  it("rejects sheet prefix", () => expect(() => parseA1Range("Sheet2!A1")).toThrow(/"sheet" argument/));
  it("rejects three parts", () => expect(() => parseA1Range("A1:B2:C3")).toThrow());
  it("rejects past column ZZ", () => expect(() => parseA1Range("AAA1")).toThrow(/sheet limit/));
  it("rejects past the row limit", () => expect(() => parseA1Range("A100001")).toThrow(/sheet limit/));
});

describe("rectToA1", () => {
  it("single", () => expect(rectToA1(parseA1Range("B2"))).toBe("B2"));
  it("range", () => expect(rectToA1(parseA1Range("b2:a1"))).toBe("A1:B2"));
});

describe("usedRange", () => {
  it("empty sheet", () => expect(usedRange(sheetWith({}))).toBeNull());
  it("bounds every cell", () => {
    const r = usedRange(sheetWith({ B2: "x", D1: "y", A5: "z" }));
    expect(r && rectToA1(r)).toBe("A1:D5");
  });
});

describe("findSheetIndex", () => {
  const wb: WorkbookData = { sheets: [sheetWith({}, "Sheet 1"), sheetWith({}, "Data")], activeSheet: 1 };
  it("defaults to active", () => expect(findSheetIndex(wb)).toBe(1));
  it("case-insensitive", () => expect(findSheetIndex(wb, "sheet 1")).toBe(0));
  it("lists sheets when missing", () => expect(() => findSheetIndex(wb, "Q4")).toThrow('Sheet "Q4" does not exist. Sheets: "Sheet 1", "Data".'));
});

describe("checkSheetName", () => {
  const wb: WorkbookData = { sheets: [sheetWith({}, "Data")], activeSheet: 0 };
  it("trims", () => expect(checkSheetName(wb, "  Q1 ")).toBe("Q1"));
  it("rejects duplicates", () => expect(() => checkSheetName(wb, "data")).toThrow(/already exists/));
  it("allows the sheet's own name when renaming", () => expect(checkSheetName(wb, "DATA", 0)).toBe("DATA"));
  it("rejects empty", () => expect(() => checkSheetName(wb, " ")).toThrow(/non-empty/));
});

describe("readRange", () => {
  const s = sheetWith({ A1: "Item", B1: "Price", A2: "Tea", B2: "3.5", B3: "=B2*2", A3: "007" });
  it("returns values and formulas", () => {
    expect(readRange(s, "A1:B3")).toEqual({
      range: "A1:B3",
      values: [["Item", "Price"], ["Tea", 3.5], ["007", 7]],
      formulas: [[null, null], [null, null], [null, "=B2*2"]],
    });
  });
  it("empty cells are empty strings", () => expect(readRange(s, "C1").values).toEqual([[""]]));
  it("resolves cross-sheet refs through the lookup", () => {
    const t = sheetWith({ A1: "=Data!A1+1" });
    expect(readRange(t, "A1", (name, ref) => (name === "Data" && ref === "A1" ? "41" : "")).values).toEqual([[42]]);
  });
  it("caps the size", () => {
    expect(() => readRange(s, "A1:J1001")).toThrow(`at most ${MAX_CELLS} cells`);
    expect(() => readRange(s, "A1:J1000")).not.toThrow();
  });
});

describe("writeCells", () => {
  it("writes from a start cell and grows the sheet", () => {
    const s = { ...sheetWith({ A1: "keep" }), rowCount: 2, colCount: 2 };
    const r = writeCells(s, "B2", [["x", 1, true], [null, "=SUM(C2:C2)", false]]);
    expect(r.range).toBe("B2:D3");
    expect(r.cells).toBe(6);
    expect(r.sheet.cells).toEqual({ A1: "keep", B2: "x", C2: "1", D2: "TRUE", C3: "=SUM(C2:C2)", D3: "FALSE" });
    expect(r.sheet.colCount).toBe(4);
    expect(r.sheet.rowCount).toBe(3);
    expect(s.cells).toEqual({ A1: "keep" });
  });
  it("empty string and null clear cells", () => {
    const r = writeCells(sheetWith({ A1: "a", B1: "b" }), "A1", [["", null]]);
    expect(r.sheet.cells).toEqual({});
  });
  it("accepts a range of the same shape", () => {
    expect(writeCells(sheetWith({}), "A1:B2", [[1, 2], [3, 4]]).range).toBe("A1:B2");
  });
  it("rejects a range of another shape", () => {
    expect(() => writeCells(sheetWith({}), "A1:B3", [[1, 2], [3, 4]])).toThrow(/3 rows x 2 columns but "values" is 2 x 2/);
  });
  it("rejects non-2D values", () => {
    expect(() => writeCells(sheetWith({}), "A1", [1, 2])).toThrow(/2D array/);
    expect(() => writeCells(sheetWith({}), "A1", [])).toThrow(/2D array/);
    expect(() => writeCells(sheetWith({}), "A1", [[{ a: 1 }]])).toThrow(/A1 must be a string/);
  });
  it("caps the size", () => {
    const rows = Array.from({ length: 101 }, () => Array.from({ length: 100 }, () => 1));
    expect(() => writeCells(sheetWith({}), "A1", rows)).toThrow(`at most ${MAX_CELLS} cells`);
  });
  it("refuses to go past column ZZ", () => {
    expect(() => writeCells(sheetWith({}), "ZZ1", [[1, 2]])).toThrow(/sheet limit/);
  });
});

describe("clearCells", () => {
  it("clears contents only", () => {
    const s = { ...sheetWith({ A1: "a", B2: "b", C3: "c" }), formats: { A1: { bold: true } } };
    const r = clearCells(s, "A1:B2");
    expect(r).toMatchObject({ range: "A1:B2", cleared: 2 });
    expect(r.sheet.cells).toEqual({ C3: "c" });
    expect(r.sheet.formats).toEqual({ A1: { bold: true } });
  });
});
