import { expandRange } from "./cell-refs";

export function evaluateFormula(
  formula: string,
  cells: Record<string, string>,
  visited: Set<string>,
  currentCell: string,
): string {
  if (visited.has(currentCell)) return "#CIRC!";
  visited.add(currentCell);
  const expr = formula.slice(1).trim();

  const resolveRef = (ref: string): number => {
    const raw = cells[ref] ?? "";
    if (!raw) return 0;
    const val = raw.startsWith("=")
      ? evaluateFormula(raw, cells, new Set(visited), ref)
      : raw;
    const num = parseFloat(val);
    return isNaN(num) ? 0 : num;
  };

  const resolveRefs = (arg: string): number[] => {
    arg = arg.trim();
    if (arg.includes(":")) return expandRange(arg).map(resolveRef);
    return [resolveRef(arg)];
  };

  const fnStart = expr.match(/^([A-Z]+)\(/i);
  if (fnStart && expr.endsWith(")")) {
    const fn = fnStart[1]!.toUpperCase();
    const inner = expr.slice(fnStart[0].length, -1);
    const args = splitArgs(inner);
    const evalArg = (a: string) => evaluateExpression(a.trim(), cells, visited, currentCell);
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

  return evaluateExpression(expr, cells, visited, currentCell);
}

export function splitArgs(str: string): string[] {
  const args: string[] = [];
  let depth = 0;
  let current = "";
  for (const ch of str) {
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
): string {
  let resolved = expr.trim();
  resolved = resolved.replace(/\b([A-Z]+\d+)\b/gi, (match) => {
    const ref = match.toUpperCase();
    const raw = cells[ref] ?? "";
    if (!raw) return "0";
    if (raw.startsWith("=")) return evaluateFormula(raw, cells, new Set(visited), ref);
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

export function computeDisplay(id: string, cells: Record<string, string>): string {
  const raw = cells[id];
  if (!raw) return "";
  if (!raw.startsWith("=")) return raw;
  return evaluateFormula(raw, cells, new Set(), id);
}
