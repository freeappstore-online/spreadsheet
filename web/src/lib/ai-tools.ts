// Pure range logic behind the local AI tools (see hooks/useAiAccess.ts).
// Errors are thrown as instructions: the AI reads them as-is.

import type { Sheet, WorkbookData } from "../types";
import { cellId, parseCellRef } from "./cell-refs";
import { computeDisplay, type SheetLookup } from "./formula-engine";

/** Most cells one tool call may read, write or clear. */
export const MAX_CELLS = 10_000;
/** Largest sheet the AI may grow: columns A..ZZ, 100,000 rows. */
export const MAX_COLS = 702;
export const MAX_ROWS = 100_000;

export interface Rect { minC: number; maxC: number; minR: number; maxR: number }

export type CellInput = string | number | boolean | null;

const CELL_RE = /^\$?([A-Z]{1,3})\$?(\d{1,7})$/;

function parseCell(text: string, input: string): { col: number; row: number } {
  const m = CELL_RE.exec(text.trim().toUpperCase());
  const p = m ? parseCellRef(`${m[1]}${m[2]}`) : null;
  if (!p || p.row < 0) {
    throw new Error(`"${input}" is not an A1 reference. Use a cell like "B2" or a range like "A1:C10".`);
  }
  if (p.col >= MAX_COLS || p.row >= MAX_ROWS) {
    throw new Error(`"${input}" is outside the sheet limit (columns A to ZZ, rows 1 to ${MAX_ROWS}).`);
  }
  return p;
}

/** Parse "B2" or "A1:C10" (case-insensitive, "$" allowed) into a rectangle. */
export function parseA1Range(input: string): Rect {
  if (typeof input !== "string" || !input.trim()) {
    throw new Error('Give a range in A1 notation, like "A1:C10" or "B2".');
  }
  if (input.includes("!")) {
    throw new Error(`Pass the sheet name in the "sheet" argument, not in the range ("${input}").`);
  }
  const parts = input.split(":");
  if (parts.length > 2) throw new Error(`"${input}" is not an A1 range. Use "A1:C10".`);
  const a = parseCell(parts[0]!, input);
  const b = parts[1] === undefined ? a : parseCell(parts[1], input);
  return {
    minC: Math.min(a.col, b.col), maxC: Math.max(a.col, b.col),
    minR: Math.min(a.row, b.row), maxR: Math.max(a.row, b.row),
  };
}

export function rectToA1(r: Rect): string {
  const a = cellId(r.minC, r.minR);
  const b = cellId(r.maxC, r.maxR);
  return a === b ? a : `${a}:${b}`;
}

const rectCells = (r: Rect) => (r.maxC - r.minC + 1) * (r.maxR - r.minR + 1);

function checkSize(r: Rect, what: string) {
  const n = rectCells(r);
  if (n > MAX_CELLS) {
    throw new Error(`${rectToA1(r)} is ${n} cells; ${what} at most ${MAX_CELLS} cells per call. Split it into smaller ranges.`);
  }
}

/** Smallest rectangle holding every non-empty cell, or null for an empty sheet. */
export function usedRange(sheet: Sheet): Rect | null {
  let r: Rect | null = null;
  for (const id of Object.keys(sheet.cells)) {
    const p = parseCellRef(id);
    if (!p) continue;
    r = r
      ? { minC: Math.min(r.minC, p.col), maxC: Math.max(r.maxC, p.col), minR: Math.min(r.minR, p.row), maxR: Math.max(r.maxR, p.row) }
      : { minC: p.col, maxC: p.col, minR: p.row, maxR: p.row };
  }
  return r;
}

/** Index of the sheet named `name` (case-insensitive), or the active sheet when omitted. */
export function findSheetIndex(wb: WorkbookData, name?: string): number {
  if (name === undefined || name === null || name === "") return wb.activeSheet;
  const idx = wb.sheets.findIndex((s) => s.name.toLowerCase() === String(name).toLowerCase());
  if (idx < 0) {
    throw new Error(`Sheet "${name}" does not exist. Sheets: ${wb.sheets.map((s) => `"${s.name}"`).join(", ")}.`);
  }
  return idx;
}

/** Check a new sheet name: non-empty, at most 64 characters, unique (ignoring `exceptIdx`). */
export function checkSheetName(wb: WorkbookData, name: unknown, exceptIdx = -1): string {
  const trimmed = typeof name === "string" ? name.trim() : "";
  if (!trimmed) throw new Error("Give the sheet a non-empty name.");
  if (trimmed.length > 64) throw new Error("Sheet names can be at most 64 characters.");
  if (wb.sheets.some((s, i) => i !== exceptIdx && s.name.toLowerCase() === trimmed.toLowerCase())) {
    throw new Error(`A sheet named "${trimmed}" already exists.`);
  }
  return trimmed;
}

// Plain numbers come back as JSON numbers; everything else as the displayed text.
function jsonValue(display: string): string | number {
  if (display === "") return "";
  const n = Number(display);
  return Number.isFinite(n) && String(n) === display ? n : display;
}

export interface ReadResult {
  range: string;
  /** Computed values, row by row. Empty cells are "". */
  values: (string | number)[][];
  /** The formula in each cell ("=..."), or null where the cell holds a plain value or nothing. */
  formulas: (string | null)[][];
}

export function readRange(sheet: Sheet, range: string, lookup?: SheetLookup): ReadResult {
  const r = parseA1Range(range);
  checkSize(r, "read");
  const values: (string | number)[][] = [];
  const formulas: (string | null)[][] = [];
  for (let row = r.minR; row <= r.maxR; row++) {
    const vRow: (string | number)[] = [];
    const fRow: (string | null)[] = [];
    for (let col = r.minC; col <= r.maxC; col++) {
      const id = cellId(col, row);
      const raw = sheet.cells[id] ?? "";
      vRow.push(jsonValue(computeDisplay(id, sheet.cells, lookup)));
      fRow.push(raw.startsWith("=") ? raw : null);
    }
    values.push(vRow);
    formulas.push(fRow);
  }
  return { range: rectToA1(r), values, formulas };
}

function cellText(v: unknown, at: string): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v;
  if (typeof v === "number") {
    if (!Number.isFinite(v)) throw new Error(`The value for ${at} is not a finite number.`);
    return String(v);
  }
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  throw new Error(`The value for ${at} must be a string, number, boolean or null, not ${Array.isArray(v) ? "an array" : typeof v}.`);
}

/**
 * Write a 2D array of values starting at the top-left of `range`. If `range` spans more than one
 * cell, `values` must have exactly its shape. "" or null clears a cell. Grows the sheet if needed.
 */
export function writeCells(sheet: Sheet, range: string, values: unknown): { sheet: Sheet; range: string; cells: number } {
  if (!Array.isArray(values) || values.length === 0 || !values.every(Array.isArray)) {
    throw new Error('"values" must be a non-empty 2D array of rows, e.g. [["Name", "Total"], ["A", 3]].');
  }
  const rows = values as unknown[][];
  const width = Math.max(...rows.map((row) => row.length));
  if (width === 0) throw new Error('"values" has no cells. Each row needs at least one value.');
  const given = parseA1Range(range);
  const isSingle = given.minC === given.maxC && given.minR === given.maxR;
  const target: Rect = { minC: given.minC, minR: given.minR, maxC: given.minC + width - 1, maxR: given.minR + rows.length - 1 };
  if (!isSingle && (target.maxC !== given.maxC || target.maxR !== given.maxR)) {
    throw new Error(
      `${rectToA1(given)} is ${given.maxR - given.minR + 1} rows x ${given.maxC - given.minC + 1} columns but "values" is ${rows.length} x ${width}. ` +
      `Match the shape, or pass only the start cell ("${cellId(given.minC, given.minR)}").`,
    );
  }
  if (target.maxC >= MAX_COLS || target.maxR >= MAX_ROWS) {
    throw new Error(`Writing ${rows.length} x ${width} from ${cellId(given.minC, given.minR)} goes past the sheet limit (columns A to ZZ, rows 1 to ${MAX_ROWS}).`);
  }
  checkSize(target, "write");
  const cells = { ...sheet.cells };
  let count = 0;
  rows.forEach((row, dr) => {
    row.forEach((v, dc) => {
      const id = cellId(target.minC + dc, target.minR + dr);
      const text = cellText(v, id);
      if (text === "") delete cells[id];
      else cells[id] = text;
      count++;
    });
  });
  return {
    sheet: {
      ...sheet,
      cells,
      colCount: Math.max(sheet.colCount, target.maxC + 1),
      rowCount: Math.max(sheet.rowCount, target.maxR + 1),
    },
    range: rectToA1(target),
    cells: count,
  };
}

/** Remove the contents (not the formatting) of every cell in `range`. */
export function clearCells(sheet: Sheet, range: string): { sheet: Sheet; range: string; cleared: number } {
  const r = parseA1Range(range);
  checkSize(r, "clear");
  const cells = { ...sheet.cells };
  let cleared = 0;
  for (let row = r.minR; row <= r.maxR; row++) {
    for (let col = r.minC; col <= r.maxC; col++) {
      const id = cellId(col, row);
      if (id in cells) { delete cells[id]; cleared++; }
    }
  }
  return { sheet: { ...sheet, cells }, range: rectToA1(r), cleared };
}
