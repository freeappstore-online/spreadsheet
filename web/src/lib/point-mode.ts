import { cellId } from "./cell-refs";
import type { PointMode } from "../types";

export const REF_TRIGGERS = new Set(["=", "(", ",", "+", "-", "*", "/", "<", ">", "^", "%", "&", "|", "!", ":"]);

export function isRefPosition(value: string, cursorPos: number): boolean {
  if (!value.startsWith("=")) return false;
  if (cursorPos === 0) return false;
  for (let i = cursorPos - 1; i >= 0; i--) {
    const ch = value[i]!;
    if (ch === " ") continue;
    return REF_TRIGGERS.has(ch);
  }
  return true;
}

export function buildRefString(pm: PointMode): string {
  const a = pm.anchor;
  const b = pm.active;
  if (a.col === b.col && a.row === b.row) return cellId(a.col, a.row);
  return `${cellId(Math.min(a.col, b.col), Math.min(a.row, b.row))}:${cellId(Math.max(a.col, b.col), Math.max(a.row, b.row))}`;
}

export function splice(str: string, start: number, end: number, insert: string): string {
  return str.slice(0, start) + insert + str.slice(end);
}
