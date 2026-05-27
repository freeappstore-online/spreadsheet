import { describe, it, expect } from "vitest";
import { parseCsvLine } from "../lib/csv";

describe("parseCsvLine", () => {
  it("simple comma-separated", () => {
    expect(parseCsvLine("a,b,c", ",")).toEqual(["a", "b", "c"]);
  });
  it("quoted field with comma", () => {
    expect(parseCsvLine('"hello, world",b', ",")).toEqual(["hello, world", "b"]);
  });
  it("escaped quotes", () => {
    expect(parseCsvLine('"say ""hi""",b', ",")).toEqual(['say "hi"', "b"]);
  });
  it("tab-separated", () => {
    expect(parseCsvLine("a\tb\tc", "\t")).toEqual(["a", "b", "c"]);
  });
  it("empty fields", () => {
    expect(parseCsvLine("a,,c", ",")).toEqual(["a", "", "c"]);
  });
  it("single field", () => {
    expect(parseCsvLine("hello", ",")).toEqual(["hello"]);
  });
  it("empty line", () => {
    expect(parseCsvLine("", ",")).toEqual([""]);
  });
  it("quoted empty field", () => {
    expect(parseCsvLine('"",b', ",")).toEqual(["", "b"]);
  });
});
