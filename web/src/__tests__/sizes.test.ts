import { describe, it, expect } from "vitest";
import { rowAt, rowOffsets, setSize, shiftSizes } from "../lib/sizes";

describe("shiftSizes", () => {
  it("moves sizes at and after an inserted index", () => {
    expect(shiftSizes({ 0: 50, 2: 80, 3: 90 }, 2, 1)).toEqual({ 0: 50, 3: 80, 4: 90 });
  });
  it("drops the deleted index and moves later ones back", () => {
    expect(shiftSizes({ 0: 50, 2: 80, 3: 90 }, 2, -1)).toEqual({ 0: 50, 2: 90 });
  });
  it("keeps a missing map missing", () => {
    expect(shiftSizes(undefined, 0, 1)).toBeUndefined();
  });
});

describe("setSize", () => {
  it("sets and resets a size", () => {
    expect(setSize(undefined, 1, 60)).toEqual({ 1: 60 });
    expect(setSize({ 1: 60, 2: 70 }, 1, null)).toEqual({ 2: 70 });
  });
});

describe("rowOffsets / rowAt", () => {
  const offsets = rowOffsets({ 1: 50 }, 4, 28); // rows: 28, 50, 28, 28
  it("sums heights", () => {
    expect(offsets).toEqual([0, 28, 78, 106, 134]);
  });
  it("finds the row under a position", () => {
    expect(rowAt(offsets, 0)).toBe(0);
    expect(rowAt(offsets, 27)).toBe(0);
    expect(rowAt(offsets, 28)).toBe(1);
    expect(rowAt(offsets, 77)).toBe(1);
    expect(rowAt(offsets, 78)).toBe(2);
    expect(rowAt(offsets, 10_000)).toBe(3);
    expect(rowAt(offsets, -5)).toBe(0);
  });
});
