import { useState, useEffect, useCallback, useRef, useMemo, useLayoutEffect } from "react";

// ── Types ──────────────────────────────────────────────────────────────

interface SheetData {
  cells: Record<string, string>;
  colCount: number;
  rowCount: number;
}

interface PointMode {
  refStart: number;
  refEnd: number;
  anchor: { col: number; row: number };
  active: { col: number; row: number };
}

// ── Persistence ────────────────────────────────────────────────────────

const STORAGE_KEY = "spreadsheet_data";
const DEFAULT_COLS = 26;
const DEFAULT_ROWS = 50;
const MAX_UNDO = 50;

function loadSheet(): SheetData {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw) as Partial<SheetData>;
      return {
        cells: parsed.cells ?? {},
        colCount: parsed.colCount ?? DEFAULT_COLS,
        rowCount: parsed.rowCount ?? DEFAULT_ROWS,
      };
    }
  } catch { /* ignore */ }
  return { cells: {}, colCount: DEFAULT_COLS, rowCount: DEFAULT_ROWS };
}

function saveSheet(data: SheetData) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
}

// ── Cell reference helpers ──────────────────────────────────────────────

function colLabel(index: number): string {
  let label = "";
  let i = index;
  while (i >= 0) {
    label = String.fromCharCode(65 + (i % 26)) + label;
    i = Math.floor(i / 26) - 1;
  }
  return label;
}

function cellId(col: number, row: number): string {
  return `${colLabel(col)}${row + 1}`;
}

function parseCellRef(ref: string): { col: number; row: number } | null {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;
  const letters = match[1]!;
  const rowNum = parseInt(match[2]!, 10);
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return { col: col - 1, row: rowNum - 1 };
}

function clamp(v: number, max: number) { return Math.max(0, Math.min(v, max - 1)); }

// ── Formula engine ──────────────────────────────────────────────────────

function expandRange(rangeStr: string): string[] {
  const [startStr, endStr] = rangeStr.split(":");
  if (!startStr || !endStr) return [];
  const start = parseCellRef(startStr);
  const end = parseCellRef(endStr);
  if (!start || !end) return [];
  const refs: string[] = [];
  for (let r = Math.min(start.row, end.row); r <= Math.max(start.row, end.row); r++) {
    for (let c = Math.min(start.col, end.col); c <= Math.max(start.col, end.col); c++) {
      refs.push(cellId(c, r));
    }
  }
  return refs;
}

function evaluateFormula(
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

  // Match function calls — find the outermost FUNC(...) pattern
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

  // Unmatched parens or plain expression
  let depth = 0;
  for (const ch of expr) { if (ch === "(") depth++; if (ch === ")") depth--; }
  if (depth !== 0) return "#PAREN!";

  return evaluateExpression(expr, cells, visited, currentCell);
}

function splitArgs(str: string): string[] {
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

function evaluateExpression(
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

function computeDisplay(id: string, cells: Record<string, string>): string {
  const raw = cells[id];
  if (!raw) return "";
  if (!raw.startsWith("=")) return raw;
  return evaluateFormula(raw, cells, new Set(), id);
}

// ── CSV parsing ────────────────────────────────────────────────────────

function parseCsvLine(line: string, sep: string): string[] {
  const result: string[] = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i]!;
    if (inQuotes) {
      if (ch === '"' && line[i + 1] === '"') { current += '"'; i++; }
      else if (ch === '"') inQuotes = false;
      else current += ch;
    } else {
      if (ch === '"') inQuotes = true;
      else if (ch === sep) { result.push(current); current = ""; }
      else current += ch;
    }
  }
  result.push(current);
  return result;
}

// ── Point mode helpers ─────────────────────────────────────────────────

const REF_TRIGGERS = new Set(["=", "(", ",", "+", "-", "*", "/", "<", ">", "^", "%", "&", "|", "!", ":"]);

function isRefPosition(value: string, cursorPos: number): boolean {
  if (!value.startsWith("=")) return false;
  if (cursorPos === 0) return false;
  for (let i = cursorPos - 1; i >= 0; i--) {
    const ch = value[i]!;
    if (ch === " ") continue;
    return REF_TRIGGERS.has(ch);
  }
  return true;
}

function buildRefString(pm: PointMode): string {
  const a = pm.anchor;
  const b = pm.active;
  if (a.col === b.col && a.row === b.row) return cellId(a.col, a.row);
  return `${cellId(Math.min(a.col, b.col), Math.min(a.row, b.row))}:${cellId(Math.max(a.col, b.col), Math.max(a.row, b.row))}`;
}

function splice(str: string, start: number, end: number, insert: string): string {
  return str.slice(0, start) + insert + str.slice(end);
}

// ── App ────────────────────────────────────────────────────────────────

export function App() {
  const [sheet, setSheet] = useState<SheetData>(loadSheet);
  const [selectedCell, setSelectedCell] = useState<string>("A1");
  const [editingCell, setEditingCell] = useState<string | null>(null);
  const [editValue, setEditValue] = useState("");
  const [formulaBarValue, setFormulaBarValue] = useState("");
  const [formulaBarFocused, setFormulaBarFocused] = useState(false);
  const [pointMode, setPointMode] = useState<PointMode | null>(null);
  const [clipboard, setClipboard] = useState<{ type: "copy" | "cut"; id: string; value: string } | null>(null);
  const [, setUndoStack] = useState<Record<string, string>[]>([]);

  const editInputRef = useRef<HTMLInputElement>(null);
  const formulaBarRef = useRef<HTMLInputElement>(null);
  const gridRef = useRef<HTMLDivElement>(null);
  const pendingCursorPos = useRef<number | null>(null);

  // ── Persistence ─────────────────────────────────────────────────────

  useEffect(() => { saveSheet(sheet); }, [sheet]);

  // ── Focus & cursor: run synchronously before paint ──────────────────

  useLayoutEffect(() => {
    if (editingCell && editInputRef.current) {
      editInputRef.current.focus();
    }
  }, [editingCell]);

  useLayoutEffect(() => {
    if (pendingCursorPos.current !== null && editInputRef.current) {
      const pos = pendingCursorPos.current;
      editInputRef.current.setSelectionRange(pos, pos);
      pendingCursorPos.current = null;
    }
  });

  // Auto-focus grid on mount
  useEffect(() => { gridRef.current?.focus(); }, []);

  // ── Derived state ───────────────────────────────────────────────────

  const displayValues = useMemo(() => {
    const result: Record<string, string> = {};
    for (const id of Object.keys(sheet.cells)) {
      result[id] = computeDisplay(id, sheet.cells);
    }
    return result;
  }, [sheet.cells]);

  const pointHighlight = useMemo<Set<string>>(() => {
    if (!pointMode) return new Set();
    const a = pointMode.anchor, b = pointMode.active;
    const s = new Set<string>();
    for (let r = Math.min(a.row, b.row); r <= Math.max(a.row, b.row); r++) {
      for (let c = Math.min(a.col, b.col); c <= Math.max(a.col, b.col); c++) {
        s.add(cellId(c, r));
      }
    }
    return s;
  }, [pointMode]);

  // ── Undo ────────────────────────────────────────────────────────────

  const pushUndo = useCallback(() => {
    setUndoStack((prev) => {
      const next = [...prev, { ...sheet.cells }];
      if (next.length > MAX_UNDO) next.shift();
      return next;
    });
  }, [sheet.cells]);

  const undo = useCallback(() => {
    setUndoStack((prev) => {
      if (prev.length === 0) return prev;
      const next = [...prev];
      const restored = next.pop()!;
      setSheet((s) => ({ ...s, cells: restored }));
      return next;
    });
  }, []);

  // ── Cell mutations ──────────────────────────────────────────────────

  const writeCell = useCallback((id: string, value: string) => {
    pushUndo();
    setSheet((prev) => {
      const next = { ...prev, cells: { ...prev.cells } };
      if (value === "") delete next.cells[id];
      else next.cells[id] = value;
      return next;
    });
  }, [pushUndo]);

  // ── Edit lifecycle ──────────────────────────────────────────────────

  const startEditing = useCallback((id: string, initial?: string) => {
    setEditingCell(id);
    setEditValue(initial ?? sheet.cells[id] ?? "");
    setPointMode(null);
  }, [sheet.cells]);

  const commitEdit = useCallback(() => {
    if (!editingCell) return;
    writeCell(editingCell, editValue);
    setEditingCell(null);
    setPointMode(null);
    setEditValue("");
  }, [editingCell, editValue, writeCell]);

  const cancelEdit = useCallback(() => {
    setEditingCell(null);
    setPointMode(null);
    setEditValue("");
  }, []);

  const focusGrid = useCallback(() => {
    requestAnimationFrame(() => gridRef.current?.focus());
  }, []);

  // ── Selection ───────────────────────────────────────────────────────

  const select = useCallback((id: string) => {
    setSelectedCell(id);
    setFormulaBarValue(sheet.cells[id] ?? "");
  }, [sheet.cells]);

  const move = useCallback((dCol: number, dRow: number) => {
    const parsed = parseCellRef(selectedCell);
    if (!parsed) return;
    const id = cellId(
      clamp(parsed.col + dCol, sheet.colCount),
      clamp(parsed.row + dRow, sheet.rowCount),
    );
    select(id);
  }, [selectedCell, sheet.colCount, sheet.rowCount, select]);

  // Commit current edit, then move selection and focus grid
  const commitAndMove = useCallback((dCol: number, dRow: number) => {
    commitEdit();
    move(dCol, dRow);
    focusGrid();
  }, [commitEdit, move, focusGrid]);

  // Sync formula bar when idle (not editing, not typing in formula bar)
  useEffect(() => {
    if (!formulaBarFocused && !editingCell) {
      setFormulaBarValue(sheet.cells[selectedCell] ?? "");
    }
  }, [selectedCell, sheet.cells, formulaBarFocused, editingCell]);

  // Sync formula bar live while editing inline
  useEffect(() => {
    if (editingCell) setFormulaBarValue(editValue);
  }, [editingCell, editValue]);

  // ── Point mode logic ────────────────────────────────────────────────

  const doPointMode = useCallback((dCol: number, dRow: number, extend: boolean) => {
    if (!editingCell) return;
    const origin = parseCellRef(editingCell);
    if (!origin) return;

    if (pointMode) {
      // Move existing reference
      const newActive = {
        col: clamp(pointMode.active.col + dCol, sheet.colCount),
        row: clamp(pointMode.active.row + dRow, sheet.rowCount),
      };
      const newAnchor = extend ? pointMode.anchor : newActive;
      const pm: PointMode = { ...pointMode, anchor: newAnchor, active: newActive };
      const ref = buildRefString(pm);
      const val = splice(editValue, pointMode.refStart, pointMode.refEnd, ref);
      pm.refEnd = pm.refStart + ref.length;
      setPointMode(pm);
      setEditValue(val);
      pendingCursorPos.current = pm.refEnd;
    } else {
      // Enter point mode: insert new reference at cursor
      const cursor = editInputRef.current?.selectionStart ?? editValue.length;
      if (!isRefPosition(editValue, cursor)) return;
      const tc = clamp(origin.col + dCol, sheet.colCount);
      const tr = clamp(origin.row + dRow, sheet.rowCount);
      const ref = cellId(tc, tr);
      const val = splice(editValue, cursor, cursor, ref);
      setPointMode({
        refStart: cursor,
        refEnd: cursor + ref.length,
        anchor: { col: tc, row: tr },
        active: { col: tc, row: tr },
      });
      setEditValue(val);
      pendingCursorPos.current = cursor + ref.length;
    }
  }, [editingCell, editValue, pointMode, sheet.colCount, sheet.rowCount]);

  // Insert/replace a ref by clicking a cell
  const insertRefByClick = useCallback((targetId: string, extend: boolean) => {
    const target = parseCellRef(targetId);
    if (!target) return;

    if (extend && pointMode) {
      const pm: PointMode = { ...pointMode, active: target };
      const ref = buildRefString(pm);
      const val = splice(editValue, pointMode.refStart, pointMode.refEnd, ref);
      pm.refEnd = pm.refStart + ref.length;
      setPointMode(pm);
      setEditValue(val);
      pendingCursorPos.current = pm.refEnd;
    } else {
      const cursor = pointMode ? pointMode.refStart : (editInputRef.current?.selectionStart ?? editValue.length);
      const removeEnd = pointMode ? pointMode.refEnd : cursor;
      const ref = targetId;
      const val = splice(editValue, cursor, removeEnd, ref);
      const end = cursor + ref.length;
      setPointMode({ refStart: cursor, refEnd: end, anchor: target, active: target });
      setEditValue(val);
      pendingCursorPos.current = end;
    }
  }, [editValue, pointMode]);

  // ── Cell mouseDown: unified interaction handler ─────────────────────
  // Using mouseDown (not click) so we can preventDefault to keep the
  // editing input focused during formula reference insertion.

  const handleCellMouseDown = useCallback((id: string, e: React.MouseEvent) => {
    // Formula edit → click another cell = insert reference
    if (editingCell && editValue.startsWith("=") && id !== editingCell) {
      e.preventDefault(); // keep input focused
      insertRefByClick(id, e.shiftKey);
      return;
    }

    // Non-formula edit → click another cell = commit + select
    if (editingCell && id !== editingCell) {
      commitEdit();
    }

    select(id);
    if (!editingCell) {
      // Ensure grid gets focus when selecting (not editing)
      gridRef.current?.focus();
    }
  }, [editingCell, editValue, insertRefByClick, commitEdit, select]);

  // ── Keyboard: grid (navigation when not editing) ────────────────────

  const handleGridKeyDown = useCallback((e: React.KeyboardEvent) => {
    // When editing, the input captures keys via stopPropagation.
    // This handler only fires when the grid itself has focus (not editing).
    const meta = e.metaKey || e.ctrlKey;

    if (meta && e.key === "z") { e.preventDefault(); undo(); return; }

    if (meta && e.key === "c") {
      e.preventDefault();
      setClipboard({ type: "copy", id: selectedCell, value: sheet.cells[selectedCell] ?? "" });
      return;
    }
    if (meta && e.key === "x") {
      e.preventDefault();
      setClipboard({ type: "cut", id: selectedCell, value: sheet.cells[selectedCell] ?? "" });
      return;
    }
    if (meta && e.key === "v" && clipboard) {
      e.preventDefault();
      writeCell(selectedCell, clipboard.value);
      if (clipboard.type === "cut") { writeCell(clipboard.id, ""); setClipboard(null); }
      return;
    }

    switch (e.key) {
      case "ArrowUp":    e.preventDefault(); move(0, -1); break;
      case "ArrowDown":  e.preventDefault(); move(0, 1); break;
      case "ArrowLeft":  e.preventDefault(); move(-1, 0); break;
      case "ArrowRight": e.preventDefault(); move(1, 0); break;
      case "Tab":        e.preventDefault(); move(e.shiftKey ? -1 : 1, 0); break;
      case "Enter":
      case "F2":
        e.preventDefault();
        startEditing(selectedCell);
        break;
      case "Delete":
      case "Backspace":
        writeCell(selectedCell, "");
        setFormulaBarValue("");
        break;
      default:
        // Any printable character starts editing with that character
        if (e.key.length === 1 && !meta) {
          e.preventDefault();
          startEditing(selectedCell, e.key);
        }
    }
  }, [selectedCell, sheet.cells, clipboard, move, startEditing, writeCell, undo]);

  // ── Keyboard: inline cell input ─────────────────────────────────────

  const handleInputKeyDown = useCallback((e: React.KeyboardEvent) => {
    // Stop propagation for ALL keys so the grid handler never double-fires.
    // Native behaviors (Ctrl+C/V/Z on the text) still work since
    // stopPropagation doesn't affect the target element's own defaults.
    e.stopPropagation();

    if (e.key === "Enter") {
      e.preventDefault();
      commitAndMove(0, e.shiftKey ? -1 : 1);
      return;
    }
    if (e.key === "Tab") {
      e.preventDefault();
      commitAndMove(e.shiftKey ? -1 : 1, 0);
      return;
    }
    if (e.key === "Escape") {
      e.preventDefault();
      cancelEdit();
      focusGrid();
      return;
    }

    // Formula point mode: arrow keys insert / move cell references
    const isArrow = e.key === "ArrowUp" || e.key === "ArrowDown" ||
                    e.key === "ArrowLeft" || e.key === "ArrowRight";
    if (isArrow && editValue.startsWith("=")) {
      const cursor = editInputRef.current?.selectionStart ?? editValue.length;
      if (pointMode || isRefPosition(editValue, cursor)) {
        e.preventDefault();
        const dCol = e.key === "ArrowLeft" ? -1 : e.key === "ArrowRight" ? 1 : 0;
        const dRow = e.key === "ArrowUp" ? -1 : e.key === "ArrowDown" ? 1 : 0;
        doPointMode(dCol, dRow, e.shiftKey);
        return;
      }
      // Not at a ref position → let arrow move the text cursor normally
    }

    // Any non-modifier key exits point mode (the typed character goes in normally)
    if (pointMode && e.key !== "Shift" && e.key !== "Control" && e.key !== "Meta" && e.key !== "Alt") {
      setPointMode(null);
    }
  }, [editValue, pointMode, commitAndMove, cancelEdit, focusGrid, doPointMode]);

  // ── Formula bar ─────────────────────────────────────────────────────

  const handleFormulaBarKeyDown = useCallback((e: React.KeyboardEvent) => {
    if (e.key === "Enter") {
      e.preventDefault();
      if (editingCell) {
        commitEdit();
      } else if (selectedCell) {
        pushUndo();
        setSheet((prev) => {
          const next = { ...prev, cells: { ...prev.cells } };
          if (formulaBarValue === "") delete next.cells[selectedCell];
          else next.cells[selectedCell] = formulaBarValue;
          return next;
        });
      }
      formulaBarRef.current?.blur();
      gridRef.current?.focus();
    } else if (e.key === "Escape") {
      setFormulaBarValue(sheet.cells[selectedCell] ?? "");
      if (editingCell) cancelEdit();
      formulaBarRef.current?.blur();
      gridRef.current?.focus();
    }
  }, [editingCell, selectedCell, formulaBarValue, sheet.cells, commitEdit, cancelEdit, pushUndo]);

  // ── Clear all ───────────────────────────────────────────────────────

  const clearAll = useCallback(() => {
    if (editingCell) cancelEdit();
    pushUndo();
    setSheet({ cells: {}, colCount: DEFAULT_COLS, rowCount: DEFAULT_ROWS });
    setSelectedCell("A1");
    setFormulaBarValue("");
    gridRef.current?.focus();
  }, [editingCell, cancelEdit, pushUndo]);

  // ── CSV export / import ──────────────────────────────────────────────

  const exportCsv = useCallback(() => {
    let maxRow = 0, maxCol = 0;
    for (const id of Object.keys(sheet.cells)) {
      const p = parseCellRef(id);
      if (p) { maxRow = Math.max(maxRow, p.row); maxCol = Math.max(maxCol, p.col); }
    }
    const rows: string[] = [];
    for (let r = 0; r <= maxRow; r++) {
      const cols: string[] = [];
      for (let c = 0; c <= maxCol; c++) {
        const raw = sheet.cells[cellId(c, r)] ?? "";
        const val = raw.startsWith("=") ? computeDisplay(cellId(c, r), sheet.cells) : raw;
        cols.push(val.includes(",") || val.includes('"') || val.includes("\n")
          ? `"${val.replace(/"/g, '""')}"` : val);
      }
      rows.push(cols.join(","));
    }
    const blob = new Blob([rows.join("\n")], { type: "text/csv" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "spreadsheet.csv";
    a.click();
    URL.revokeObjectURL(a.href);
  }, [sheet.cells]);

  const importCsv = useCallback(() => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".csv,.tsv,.txt";
    input.onchange = () => {
      const file = input.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        const text = reader.result as string;
        const sep = text.includes("\t") ? "\t" : ",";
        const lines = text.split(/\r?\n/);
        pushUndo();
        const cells: Record<string, string> = {};
        let maxCol = 0;
        for (let r = 0; r < lines.length; r++) {
          const row = lines[r]!;
          if (!row.trim()) continue;
          const vals = parseCsvLine(row, sep);
          maxCol = Math.max(maxCol, vals.length - 1);
          for (let c = 0; c < vals.length; c++) {
            const v = vals[c]!.trim();
            if (v) cells[cellId(c, r)] = v;
          }
        }
        setSheet({
          cells,
          colCount: Math.max(DEFAULT_COLS, maxCol + 1),
          rowCount: Math.max(DEFAULT_ROWS, lines.length),
        });
        setSelectedCell("A1");
        setFormulaBarValue(cells["A1"] ?? "");
      };
      reader.readAsText(file);
    };
    input.click();
  }, [pushUndo]);

  // ── Render helpers ──────────────────────────────────────────────────

  const isError = (val: string) => val.startsWith("#");
  const COL_WIDTH = 100;
  const ROW_HEIGHT = 28;
  const HEADER_WIDTH = 44;

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100dvh" }}>
      {/* Top bar */}
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: "0.75rem",
          padding: "0.5rem 0.75rem",
          borderBottom: "1px solid var(--color-line)",
          background: "var(--color-panel)",
          flexShrink: 0,
        }}
      >
        <span style={{ fontFamily: "'Fraunces', Georgia, serif", fontWeight: 700, fontSize: "1.125rem", whiteSpace: "nowrap" }}>
          spreadsheet
        </span>
        <div style={{ flex: 1 }} />
        {[
          { label: "Import CSV", action: importCsv },
          { label: "Export CSV", action: exportCsv },
          { label: "Clear all", action: clearAll },
        ].map(({ label, action }) => (
          <button
            key={label}
            onMouseDown={(e) => e.preventDefault()}
            onClick={action}
            style={{
              border: "1px solid var(--color-line)",
              borderRadius: "var(--radius-btn)",
              background: "transparent",
              color: "var(--color-muted)",
              padding: "0.25rem 0.75rem",
              fontSize: "0.75rem",
              fontWeight: 600,
              cursor: "pointer",
            }}
          >
            {label}
          </button>
        ))}
        <a
          href="https://freeappstore.online"
          target="_blank"
          rel="noopener noreferrer"
          style={{ color: "var(--color-muted)", textDecoration: "none", fontSize: "0.75rem", whiteSpace: "nowrap" }}
        >
          FreeAppStore
        </a>
      </div>

      {/* Formula bar */}
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: "0.5rem",
          padding: "0.25rem 0.75rem",
          borderBottom: "1px solid var(--color-line)",
          background: "var(--color-panel)",
          flexShrink: 0,
          minHeight: "2rem",
        }}
      >
        <span style={{ fontWeight: 600, fontSize: "0.8125rem", color: "var(--color-accent)", minWidth: "3rem", textAlign: "center" }}>
          {selectedCell}
        </span>
        <span style={{ color: "var(--color-line)" }}>|</span>
        <input
          ref={formulaBarRef}
          value={formulaBarValue}
          onChange={(e) => {
            setFormulaBarValue(e.target.value);
            if (editingCell) {
              setEditValue(e.target.value);
              setPointMode(null);
            }
          }}
          onFocus={() => {
            setFormulaBarFocused(true);
            if (!editingCell) startEditing(selectedCell);
          }}
          onBlur={() => setFormulaBarFocused(false)}
          onKeyDown={handleFormulaBarKeyDown}
          placeholder="Enter a value or formula (e.g. =SUM(A1:A10))"
          style={{
            flex: 1,
            border: "none",
            outline: "none",
            background: "transparent",
            fontSize: "0.875rem",
            color: "var(--color-ink)",
            padding: "0.25rem 0",
          }}
        />
      </div>

      {/* Grid */}
      <div
        ref={gridRef}
        tabIndex={0}
        onKeyDown={handleGridKeyDown}
        style={{ flex: 1, overflow: "auto", outline: "none", position: "relative" }}
      >
        <div
          style={{
            display: "grid",
            gridTemplateColumns: `${HEADER_WIDTH}px repeat(${sheet.colCount}, ${COL_WIDTH}px)`,
            gridTemplateRows: `${ROW_HEIGHT}px repeat(${sheet.rowCount}, ${ROW_HEIGHT}px)`,
            width: "fit-content",
            minWidth: "100%",
          }}
        >
          {/* Corner */}
          <div style={{
            position: "sticky", top: 0, left: 0, zIndex: 3,
            background: "var(--color-panel)",
            borderBottom: "2px solid var(--color-line)",
            borderRight: "1px solid var(--color-line)",
          }} />

          {/* Column headers */}
          {Array.from({ length: sheet.colCount }, (_, c) => (
            <div
              key={`ch-${c}`}
              style={{
                position: "sticky", top: 0, zIndex: 2,
                display: "flex", alignItems: "center", justifyContent: "center",
                background: "var(--color-panel)",
                borderBottom: "2px solid var(--color-line)",
                borderRight: "1px solid var(--color-line)",
                fontSize: "0.6875rem", fontWeight: 600, color: "var(--color-muted)",
                userSelect: "none",
              }}
            >
              {colLabel(c)}
            </div>
          ))}

          {/* Rows */}
          {Array.from({ length: sheet.rowCount }, (_, r) => (
            <>
              {/* Row header */}
              <div
                key={`rh-${r}`}
                style={{
                  position: "sticky", left: 0, zIndex: 1,
                  display: "flex", alignItems: "center", justifyContent: "center",
                  background: "var(--color-panel)",
                  borderBottom: "1px solid var(--color-line)",
                  borderRight: "1px solid var(--color-line)",
                  fontSize: "0.6875rem", fontWeight: 500, color: "var(--color-muted)",
                  userSelect: "none",
                }}
              >
                {r + 1}
              </div>

              {/* Cells */}
              {Array.from({ length: sheet.colCount }, (_, c) => {
                const id = cellId(c, r);
                const isSelected = selectedCell === id;
                const isEditing = editingCell === id;
                const isPointed = pointHighlight.has(id);
                const display = displayValues[id] ?? "";

                return (
                  <div
                    key={id}
                    onMouseDown={(e) => handleCellMouseDown(id, e)}
                    onDoubleClick={() => {
                      if (!editingCell) startEditing(id);
                    }}
                    style={{
                      position: "relative",
                      borderBottom: "1px solid var(--color-line)",
                      borderRight: "1px solid var(--color-line)",
                      outline: isSelected
                        ? "2px solid var(--color-accent)"
                        : isPointed
                          ? "2px solid #5b8cd6"
                          : "none",
                      outlineOffset: "-1px",
                      background: isPointed && !isSelected ? "rgba(91,140,214,0.12)" : undefined,
                      zIndex: isSelected || isPointed ? 1 : 0,
                      cursor: "cell",
                    }}
                  >
                    {isEditing ? (
                      <input
                        ref={editInputRef}
                        value={editValue}
                        onChange={(e) => {
                          if (pointMode) setPointMode(null);
                          setEditValue(e.target.value);
                        }}
                        onKeyDown={handleInputKeyDown}
                        style={{
                          position: "absolute", inset: 0,
                          border: "none", outline: "none",
                          background: "var(--color-paper)",
                          padding: "0 4px", fontSize: "0.8125rem",
                          color: "var(--color-ink)",
                          width: "100%", height: "100%",
                        }}
                      />
                    ) : (
                      <div
                        style={{
                          padding: "0 4px",
                          fontSize: "0.8125rem",
                          lineHeight: `${ROW_HEIGHT}px`,
                          overflow: "hidden",
                          textOverflow: "ellipsis",
                          whiteSpace: "nowrap",
                          color: isError(display) ? "#d94040" : "var(--color-ink)",
                          textAlign: !isNaN(parseFloat(display)) && display !== "" && !isError(display) ? "right" : "left",
                          fontWeight: isError(display) ? 600 : 400,
                        }}
                      >
                        {display}
                      </div>
                    )}
                  </div>
                );
              })}
            </>
          ))}
        </div>
      </div>
    </div>
  );
}

export default App;
