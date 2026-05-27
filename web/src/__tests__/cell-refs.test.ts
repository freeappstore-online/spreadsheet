import { describe, it, expect } from "vitest";
import { colLabel, cellId, parseCellRef, clamp } from "../lib/cell-refs";

describe("colLabel", () => {
  it("A=0", () => expect(colLabel(0)).toBe("A"));
  it("Z=25", () => expect(colLabel(25)).toBe("Z"));
  it("AA=26", () => expect(colLabel(26)).toBe("AA"));
  it("AZ=51", () => expect(colLabel(51)).toBe("AZ"));
  it("BA=52", () => expect(colLabel(52)).toBe("BA"));
});

describe("cellId", () => {
  it("A1", () => expect(cellId(0, 0)).toBe("A1"));
  it("Z1", () => expect(cellId(25, 0)).toBe("Z1"));
  it("A10", () => expect(cellId(0, 9)).toBe("A10"));
  it("AA1", () => expect(cellId(26, 0)).toBe("AA1"));
});

describe("parseCellRef", () => {
  it("A1", () => expect(parseCellRef("A1")).toEqual({ col: 0, row: 0 }));
  it("Z1", () => expect(parseCellRef("Z1")).toEqual({ col: 25, row: 0 }));
  it("AA1", () => expect(parseCellRef("AA1")).toEqual({ col: 26, row: 0 }));
  it("B10", () => expect(parseCellRef("B10")).toEqual({ col: 1, row: 9 }));
  it("invalid empty", () => expect(parseCellRef("")).toBeNull());
  it("invalid number", () => expect(parseCellRef("123")).toBeNull());
  it("lowercase rejected", () => expect(parseCellRef("a1")).toBeNull());
  it("roundtrip", () => {
    for (let c = 0; c < 30; c++) {
      for (let r = 0; r < 10; r++) {
        const id = cellId(c, r);
        expect(parseCellRef(id)).toEqual({ col: c, row: r });
      }
    }
  });
});

describe("clamp", () => {
  it("within range", () => expect(clamp(5, 10)).toBe(5));
  it("below zero", () => expect(clamp(-1, 10)).toBe(0));
  it("at max", () => expect(clamp(10, 10)).toBe(9));
  it("above max", () => expect(clamp(15, 10)).toBe(9));
});
