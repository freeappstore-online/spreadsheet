import { describe, it, expect } from "vitest";
import { evaluateFormula, computeDisplay, splitArgs } from "../lib/formula-engine";
import { expandRange } from "../lib/cell-refs";

const cells: Record<string, string> = {
  A1: "10", A2: "20", A3: "30",
  B1: "hello", B2: "5", B3: "",
  C1: "=A1+A2", C2: "=SUM(A1:A3)",
};

function eval_(formula: string, c = cells): string {
  return evaluateFormula(formula, c, new Set(), "Z99");
}

describe("arithmetic functions", () => {
  it("SUM of range", () => expect(eval_("=SUM(A1:A3)")).toBe("60"));
  it("SUM of individual refs", () => expect(eval_("=SUM(A1,A2,A3)")).toBe("60"));
  it("SUM of empty range", () => expect(eval_("=SUM(D1:D3)")).toBe("0"));
  it("AVERAGE", () => expect(eval_("=AVERAGE(A1:A3)")).toBe("20"));
  it("AVG alias", () => expect(eval_("=AVG(A1:A3)")).toBe("20"));
  it("MIN", () => expect(eval_("=MIN(A1:A3)")).toBe("10"));
  it("MAX", () => expect(eval_("=MAX(A1:A3)")).toBe("30"));
  it("COUNT", () => expect(eval_("=COUNT(A1:A3)")).toBe("3"));
  it("COUNTA", () => expect(eval_("=COUNTA(A1:B3)")).toBe("5"));
});

describe("math functions", () => {
  it("ABS positive", () => expect(eval_("=ABS(A1)")).toBe("10"));
  it("ABS negative", () => expect(eval_("=ABS(-5)", {})).toBe("5"));
  it("ROUND", () => expect(eval_("=ROUND(3.456,2)", {})).toBe("3.46"));
  it("ROUND no decimals", () => expect(eval_("=ROUND(3.7)", {})).toBe("4"));
  it("FLOOR", () => expect(eval_("=FLOOR(3.9)", {})).toBe("3"));
  it("CEIL", () => expect(eval_("=CEIL(3.1)", {})).toBe("4"));
  it("CEILING alias", () => expect(eval_("=CEILING(3.1)", {})).toBe("4"));
  it("SQRT", () => expect(eval_("=SQRT(9)", {})).toBe("3"));
  it("SQRT negative", () => expect(eval_("=SQRT(-1)", {})).toBe("#NUM!"));
  it("POWER", () => expect(eval_("=POWER(2,3)", {})).toBe("8"));
  it("POW alias", () => expect(eval_("=POW(2,10)", {})).toBe("1024"));
  it("MOD", () => expect(eval_("=MOD(10,3)", {})).toBe("1"));
  it("MOD div zero", () => expect(eval_("=MOD(10,0)", {})).toBe("#DIV/0!"));
  it("INT", () => expect(eval_("=INT(3.9)", {})).toBe("3"));
  it("INT negative", () => expect(eval_("=INT(-3.9)", {})).toBe("-3"));
  it("SIGN positive", () => expect(eval_("=SIGN(42)", {})).toBe("1"));
  it("SIGN negative", () => expect(eval_("=SIGN(-5)", {})).toBe("-1"));
  it("SIGN zero", () => expect(eval_("=SIGN(0)", {})).toBe("0"));
  it("PI", () => expect(eval_("=PI()")).toBe(String(Math.PI)));
});

describe("string functions", () => {
  it("LEN", () => expect(eval_("=LEN(B1)")).toBe("5"));
  it("UPPER", () => expect(eval_("=UPPER(B1)")).toBe("HELLO"));
  it("LOWER", () => expect(eval_("=LOWER(B1)")).toBe("hello"));
  it("TRIM", () => expect(eval_("=TRIM(A1)", { A1: "  hi  " })).toBe("hi"));
  it("LEFT", () => expect(eval_("=LEFT(B1,3)")).toBe("hel"));
  it("RIGHT", () => expect(eval_("=RIGHT(B1,2)")).toBe("lo"));
  it("MID", () => expect(eval_("=MID(B1,2,3)")).toBe("ell"));
  it("CONCATENATE", () => expect(eval_("=CONCATENATE(B1,A1)")).toBe("hello10"));
  it("CONCAT alias", () => expect(eval_("=CONCAT(B1,A1)")).toBe("hello10"));
  it("SUBSTITUTE with cell refs", () => {
    const c = { A1: "hello world" };
    expect(eval_('=SUBSTITUTE(A1,world,earth)', c)).toBe("hello earth");
  });
  it('SUBSTITUTE with string literals', () => {
    const c = { A1: "hello world" };
    expect(eval_('=SUBSTITUTE(A1,"world","earth")', c)).toBe("hello earth");
  });
  it("UPPER with string literal", () => {
    expect(eval_('=UPPER("hello")', {})).toBe("HELLO");
  });
  it("CONCAT mixing literals and refs", () => {
    const c = { A1: "world" };
    expect(eval_('=CONCAT("hello ",A1)', c)).toBe("hello world");
  });
  it("LEN with literal containing comma", () => {
    expect(eval_('=LEN("a,b,c")', {})).toBe("5");
  });
  it("escaped quote in literal", () => {
    expect(eval_('=LEN("a""b")', {})).toBe("3");
  });
  it("VALUE", () => expect(eval_("=VALUE(A1)")).toBe("10"));
  it("VALUE non-number", () => expect(eval_("=VALUE(B1)")).toBe("#VALUE!"));
});

describe("logic", () => {
  it("IF true branch", () => expect(eval_("=IF(1,A1,A2)")).toBe("10"));
  it("IF false branch", () => expect(eval_("=IF(0,A1,A2)")).toBe("20"));
  it("IF with comparison", () => expect(eval_("=IF(A1>15,A1,A2)")).toBe("20"));
  it("IF 2-arg true", () => expect(eval_("=IF(1,A1)")).toBe("10"));
  it("IF 2-arg false", () => expect(eval_("=IF(0,A1)")).toBe("0"));
  it("IF too few args", () => expect(eval_("=IF(1)")).toBe("#ARG!"));
});

describe("expressions", () => {
  it("simple arithmetic", () => expect(eval_("=1+2")).toBe("3"));
  it("multiplication", () => expect(eval_("=3*4")).toBe("12"));
  it("cell reference arithmetic", () => expect(eval_("=A1+A2")).toBe("30"));
  it("parentheses", () => expect(eval_("=(A1+A2)*2")).toBe("60"));
  it("division by zero", () => expect(eval_("=1/0")).toBe("#DIV/0!"));
  it("formula referencing formula", () => expect(eval_("=C1")).toBe("30"));
  it("nested formula reference", () => expect(eval_("=C2")).toBe("60"));
});

describe("error handling", () => {
  it("circular reference", () => {
    expect(evaluateFormula("=A1", { A1: "=A1" }, new Set(), "A1")).toBe("#CIRC!");
  });
  it("mutual circular", () => {
    const c = { A1: "=B1", B1: "=A1" };
    expect(evaluateFormula("=A1", c, new Set(), "A1")).toBe("#CIRC!");
  });
  it("unknown function", () => expect(eval_("=NOTAFUNCTION(1)")).toBe("#NAME!"));
  it("unmatched parens", () => expect(eval_("=SUM(A1:A3")).toBe("#PAREN!"));
  it("invalid expression", () => expect(eval_("=A1+++")).toBe("#ERROR!"));
});

describe("splitArgs", () => {
  it("simple", () => expect(splitArgs("A1,B1")).toEqual(["A1", "B1"]));
  it("nested parens", () => expect(splitArgs("A1,IF(B1>0,1,2)")).toEqual(["A1", "IF(B1>0,1,2)"]));
  it("single arg", () => expect(splitArgs("A1:A3")).toEqual(["A1:A3"]));
  it("empty string", () => expect(splitArgs("")).toEqual([]));
});

describe("expandRange", () => {
  it("single column range", () => expect(expandRange("A1:A3")).toEqual(["A1", "A2", "A3"]));
  it("single row range", () => expect(expandRange("A1:C1")).toEqual(["A1", "B1", "C1"]));
  it("2D range", () => expect(expandRange("A1:B2")).toEqual(["A1", "B1", "A2", "B2"]));
  it("reversed range", () => expect(expandRange("B2:A1")).toEqual(["A1", "B1", "A2", "B2"]));
  it("invalid range", () => expect(expandRange("A1")).toEqual([]));
});

describe("cross-sheet references", () => {
  const sheet1: Record<string, string> = { A1: "10" };
  const sheet2: Record<string, string> = { B5: "42", B6: "8" };
  const lookup = (name: string, ref: string) => {
    if (name.toLowerCase() === "sheet2") return sheet2[ref] ?? "";
    if (name.toLowerCase() === "sheet1") return sheet1[ref] ?? "";
    return "";
  };
  it("reads single cell from other sheet", () => {
    expect(evaluateFormula("=Sheet2!B5", sheet1, new Set(), "Z1", lookup)).toBe("42");
  });
  it("arithmetic across sheets", () => {
    expect(evaluateFormula("=Sheet2!B5+A1", sheet1, new Set(), "Z1", lookup)).toBe("52");
  });
  it("SUM with cross-sheet refs", () => {
    expect(evaluateFormula("=SUM(Sheet2!B5, Sheet2!B6)", sheet1, new Set(), "Z1", lookup)).toBe("50");
  });
  it("unknown sheet returns 0", () => {
    expect(evaluateFormula("=Nothere!A1", sheet1, new Set(), "Z1", lookup)).toBe("0");
  });
  it("case-insensitive sheet name", () => {
    expect(evaluateFormula("=SHEET2!B5", sheet1, new Set(), "Z1", lookup)).toBe("42");
  });
});

describe("computeDisplay", () => {
  it("plain value", () => expect(computeDisplay("A1", cells)).toBe("10"));
  it("formula", () => expect(computeDisplay("C1", cells)).toBe("30"));
  it("empty cell", () => expect(computeDisplay("Z1", cells)).toBe(""));
});
