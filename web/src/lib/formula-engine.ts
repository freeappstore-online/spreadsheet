import { expandRange } from "./cell-refs";

// Optional callback to look up cells from other sheets (for Sheet!A1 references).
// Returns the raw cell value (which may itself be a formula).
export type SheetLookup = (sheetName: string, cellRef: string) => string;

// Match qualified ref: SheetName!A1 (sheet name is alphanumeric + underscore, must start with letter)
const QUALIFIED_REF_RE = /\b([A-Za-z_][\w]*)!([A-Z]+\d+)\b/g;

// String literal (double quotes; escape with "")
function isStringLiteral(s: string): boolean {
  const t = s.trim();
  return t.length >= 2 && t.startsWith('"') && t.endsWith('"');
}

function stripQuotes(s: string): string {
  const t = s.trim();
  if (!isStringLiteral(t)) return t;
  return t.slice(1, -1).replace(/""/g, '"');
}

export function evaluateFormula(
  formula: string,
  cells: Record<string, string>,
  visited: Set<string>,
  currentCell: string,
  sheetLookup?: SheetLookup,
): string {
  if (visited.has(currentCell)) return "#CIRC!";
  visited.add(currentCell);
  const expr = formula.slice(1).trim();

  const resolveRef = (ref: string): number => {
    // Cross-sheet ref?
    if (ref.includes("!") && sheetLookup) {
      const [sheetName, cellRef] = ref.split("!");
      if (!sheetName || !cellRef) return 0;
      const raw = sheetLookup(sheetName, cellRef);
      if (!raw) return 0;
      // Note: cross-sheet circular detection uses qualified name as the visited key
      const qualifiedKey = `${sheetName}!${cellRef}`;
      if (visited.has(qualifiedKey)) return 0;
      const val = raw.startsWith("=")
        ? evaluateFormula(raw, cells, new Set([...visited, qualifiedKey]), qualifiedKey, sheetLookup)
        : raw;
      const num = parseFloat(val);
      return isNaN(num) ? 0 : num;
    }
    const raw = cells[ref] ?? "";
    if (!raw) return 0;
    const val = raw.startsWith("=")
      ? evaluateFormula(raw, cells, new Set(visited), ref, sheetLookup)
      : raw;
    const num = parseFloat(val);
    return isNaN(num) ? 0 : num;
  };

  const resolveRefs = (arg: string): number[] => {
    arg = arg.trim();
    if (isStringLiteral(arg)) return [parseFloat(stripQuotes(arg)) || 0];
    if (arg.includes(":")) return expandRange(arg).map(resolveRef);
    return [resolveRef(arg)];
  };

  const fnStart = expr.match(/^([A-Z]+)\(/i);
  if (fnStart && expr.endsWith(")")) {
    const fn = fnStart[1]!.toUpperCase();
    const inner = expr.slice(fnStart[0].length, -1);
    const args = splitArgs(inner);
    const evalArg = (a: string) => {
      const trimmed = a.trim();
      if (isStringLiteral(trimmed)) return stripQuotes(trimmed);
      return evaluateExpression(trimmed, cells, visited, currentCell, sheetLookup);
    };
    const numArg = (a: string) => { const v = parseFloat(evalArg(a)); return isNaN(v) ? 0 : v; };

    switch (fn) {
      case "SUM": return String(args.flatMap(resolveRefs).reduce((a, b) => a + b, 0));
      case "AVG": case "AVERAGE": {
        const nums = args.flatMap(resolveRefs);
        return nums.length === 0 ? "#DIV/0!" : String(nums.reduce((a, b) => a + b, 0) / nums.length);
      }
      case "MIN": { const n = args.flatMap(resolveRefs); return n.length ? String(Math.min(...n)) : "0"; }
      case "MAX": { const n = args.flatMap(resolveRefs); return n.length ? String(Math.max(...n)) : "0"; }
      case "COUNT": return String(args.flatMap(resolveRefs).filter((n) => n !== 0).length);
      case "COUNTA": return String(args.flatMap((a) => {
        a = a.trim();
        if (isStringLiteral(a)) return [stripQuotes(a)];
        if (a.includes(":")) return expandRange(a).map((r) => cells[r] ?? "");
        return [cells[a] ?? ""];
      }).filter((v) => v !== "").length);
      case "IF": {
        if (args.length < 2) return "#ARG!";
        const cond = evalArg(args[0]!);
        const condNum = parseFloat(cond);
        const truthy = !isNaN(condNum) ? condNum !== 0 : cond.length > 0;
        if (truthy) return args[1] ? evalArg(args[1]) : "1";
        return args[2] ? evalArg(args[2]) : "0";
      }
      case "ABS": return args.length < 1 ? "#ARG!" : String(Math.abs(numArg(args[0]!)));
      case "ROUND": {
        const val = numArg(args[0] ?? "0");
        const dec = args[1] ? numArg(args[1]) : 0;
        const f = Math.pow(10, dec);
        return String(Math.round(val * f) / f);
      }
      case "FLOOR": return String(Math.floor(numArg(args[0] ?? "0")));
      case "CEIL": case "CEILING": return String(Math.ceil(numArg(args[0] ?? "0")));
      case "SQRT": { const v = numArg(args[0] ?? "0"); return v < 0 ? "#NUM!" : String(Math.sqrt(v)); }
      case "POWER": case "POW": return String(Math.pow(numArg(args[0] ?? "0"), numArg(args[1] ?? "1")));
      case "MOD": { const d = numArg(args[1] ?? "0"); return d === 0 ? "#DIV/0!" : String(numArg(args[0] ?? "0") % d); }
      case "LEN": return String(evalArg(args[0] ?? "").length);
      case "UPPER": return evalArg(args[0] ?? "").toUpperCase();
      case "LOWER": return evalArg(args[0] ?? "").toLowerCase();
      case "TRIM": return evalArg(args[0] ?? "").trim();
      case "LEFT": return evalArg(args[0] ?? "").slice(0, numArg(args[1] ?? "1"));
      case "RIGHT": { const s = evalArg(args[0] ?? ""); return s.slice(-numArg(args[1] ?? "1")); }
      case "MID": { const s = evalArg(args[0] ?? ""); return s.slice(numArg(args[1] ?? "1") - 1, numArg(args[1] ?? "1") - 1 + numArg(args[2] ?? "1")); }
      case "CONCATENATE": case "CONCAT": return args.map((a) => evalArg(a)).join("");
      case "SUBSTITUTE": {
        const text = evalArg(args[0] ?? "");
        const old = evalArg(args[1] ?? "");
        const repl = evalArg(args[2] ?? "");
        if (!old) return text;
        return text.split(old).join(repl);
      }
      case "TEXT": return evalArg(args[0] ?? "");
      case "VALUE": { const v = parseFloat(evalArg(args[0] ?? "")); return isNaN(v) ? "#VALUE!" : String(v); }
      case "INT": return String(Math.trunc(numArg(args[0] ?? "0")));
      case "SIGN": { const v = numArg(args[0] ?? "0"); return String(v > 0 ? 1 : v < 0 ? -1 : 0); }
      case "PI": return String(Math.PI);
      case "NOW": return new Date().toLocaleString();
      case "TODAY": return new Date().toLocaleDateString();
      default: return "#NAME!";
    }
  }

  let depth = 0;
  for (const ch of expr) { if (ch === "(") depth++; if (ch === ")") depth--; }
  if (depth !== 0) return "#PAREN!";

  return evaluateExpression(expr, cells, visited, currentCell, sheetLookup);
}

// Split function args on commas at depth 0, respecting parens AND string literals.
export function splitArgs(str: string): string[] {
  const args: string[] = [];
  let depth = 0;
  let current = "";
  let inString = false;
  for (let i = 0; i < str.length; i++) {
    const ch = str[i]!;
    if (inString) {
      current += ch;
      if (ch === '"') {
        // Escape: "" inside string literal
        if (str[i + 1] === '"') { current += '"'; i++; }
        else inString = false;
      }
      continue;
    }
    if (ch === '"') { inString = true; current += ch; continue; }
    if (ch === "(") depth++;
    else if (ch === ")") depth--;
    if (ch === "," && depth === 0) { args.push(current); current = ""; }
    else current += ch;
  }
  if (current) args.push(current);
  return args;
}

export function evaluateExpression(
  expr: string,
  cells: Record<string, string>,
  visited: Set<string>,
  _currentCell: string,
  sheetLookup?: SheetLookup,
): string {
  let resolved = expr.trim();

  // If the entire expression is a string literal, return its content
  if (isStringLiteral(resolved)) return stripQuotes(resolved);

  // Replace cross-sheet refs FIRST (before bare cell ref replacement)
  resolved = resolved.replace(QUALIFIED_REF_RE, (_, sheetName: string, ref: string) => {
    if (!sheetLookup) return "0";
    const qualifiedKey = `${sheetName}!${ref}`;
    if (visited.has(qualifiedKey)) return "0";
    const raw = sheetLookup(sheetName, ref);
    if (!raw) return "0";
    if (raw.startsWith("=")) {
      return evaluateFormula(raw, cells, new Set([...visited, qualifiedKey]), qualifiedKey, sheetLookup);
    }
    return raw;
  });

  // Replace bare cell refs
  resolved = resolved.replace(/\b([A-Z]+\d+)\b/gi, (match) => {
    const ref = match.toUpperCase();
    const raw = cells[ref] ?? "";
    if (!raw) return "0";
    if (raw.startsWith("=")) return evaluateFormula(raw, cells, new Set(visited), ref, sheetLookup);
    return raw;
  });

  try {
    if (/^[\d\s+\-*/().%<>=!&|]+$/.test(resolved)) {
      resolved = resolved.replace(/(?<!=)=(?!=)/g, "==");
      const result = new Function(`"use strict"; return (${resolved})`)() as number | boolean;
      if (typeof result === "boolean") return result ? "1" : "0";
      if (typeof result === "number") {
        if (!isFinite(result)) return "#DIV/0!";
        return String(Math.round(result * 1e10) / 1e10);
      }
      return String(result);
    }
    return resolved;
  } catch { return "#ERROR!"; }
}

export function computeDisplay(
  id: string,
  cells: Record<string, string>,
  sheetLookup?: SheetLookup,
): string {
  const raw = cells[id];
  if (!raw) return "";
  if (!raw.startsWith("=")) return raw;
  return evaluateFormula(raw, cells, new Set(), id, sheetLookup);
}
