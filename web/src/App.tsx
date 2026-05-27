import { useState, useEffect, useCallback, useRef, useMemo } from "react";

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

function clampCol(c: number, max: number) { return Math.max(0, Math.min(c, max - 1)); }
function clampRow(r: number, max: number) { return Math.max(0, Math.min(r, max - 1)); }

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
    if (arg.includes(":")) {
      return expandRange(arg).map(resolveRef);
    }
    return [resolveRef(arg)];
  };

  const fnMatch = expr.match(/^([A-Z]+)\((.+)\)$/i);
  if (fnMatch) {
    const fn = fnMatch[1]!.toUpperCase();
    const argsStr = fnMatch[2]!;
    const args = splitArgs(argsStr);

    switch (fn) {
      case "SUM": {
        const nums = args.flatMap(resolveRefs);
        return String(nums.reduce((a, b) => a + b, 0));
      }
      case "AVG":
      case "AVERAGE": {
        const nums = args.flatMap(resolveRefs);
        if (nums.length === 0) return "#DIV/0!";
        return String(nums.reduce((a, b) => a + b, 0) / nums.length);
      }
      case "MIN": {
        const nums = args.flatMap(resolveRefs);
        if (nums.length === 0) return "0";
        return String(Math.min(...nums));
      }
      case "MAX": {
        const nums = args.flatMap(resolveRefs);
        if (nums.length === 0) return "0";
        return String(Math.max(...nums));
      }
      case "COUNT": {
        const nums = args.flatMap(resolveRefs);
        return String(nums.filter((n) => n !== 0).length);
      }
      case "IF": {
        if (args.length < 3) return "#ARG!";
        const condition = evaluateExpression(args[0]!, cells, visited, currentCell);
        const condNum = parseFloat(condition);
        const truthy = !isNaN(condNum) ? condNum !== 0 : condition.length > 0;
        const branch = truthy ? args[1]! : args[2]!;
        return evaluateExpression(branch.trim(), cells, visited, currentCell);
      }
      case "ABS": {
        if (args.length < 1) return "#ARG!";
        const val = parseFloat(evaluateExpression(args[0]!, cells, visited, currentCell));
        return isNaN(val) ? "#VALUE!" : String(Math.abs(val));
      }
      case "ROUND": {
        if (args.length < 1) return "#ARG!";
        const val = parseFloat(evaluateExpression(args[0]!, cells, visited, currentCell));
        const decimals = args[1] ? parseInt(evaluateExpression(args[1], cells, visited, currentCell), 10) : 0;
        if (isNaN(val)) return "#VALUE!";
        const factor = Math.pow(10, decimals);
        return String(Math.round(val * factor) / factor);
      }
      default:
        return "#NAME!";
    }
  }

  return evaluateExpression(expr, cells, visited, currentCell);
}

function splitArgs(str: string): string[] {
  const args: string[] = [];
  let depth = 0;
  let current = "";
  for (const ch of str) {
    if (ch === "(") depth++;
    else if (ch === ")") depth--;
    if (ch === "," && depth === 0) {
      args.push(current);
      current = "";
    } else {
      current += ch;
    }
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
    if (raw.startsWith("=")) {
      return evaluateFormula(raw, cells, new Set(visited), ref);
    }
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
  } catch {
    return "#ERROR!";
  }
}

function computeDisplay(id: string, cells: Record<string, string>): string {
  const raw = cells[id];
  if (!raw) return "";
  if (!raw.startsWith("=")) return raw;
  return evaluateFormula(raw, cells, new Set(), id);
}

// ── Point mode helpers ─────────────────────────────────────────────────

const REF_TRIGGER_CHARS = new Set(["=", "(", ",", "+", "-", "*", "/", "<", ">", "^", "%", "&", "|", "!"]);

function isRefPosition(value: string, cursorPos: number): boolean {
  if (!value.startsWith("=")) return false;
  if (cursorPos === 0) return false;
  for (let i = cursorPos - 1; i >= 0; i--) {
    const ch = value[i]!;
    if (ch === " ") continue;
    return REF_TRIGGER_CHARS.has(ch);
  }
  return true;
}

function buildRefString(pm: PointMode): string {
  const a = pm.anchor;
  const b = pm.active;
  if (a.col === b.col && a.row === b.row) {
    return cellId(a.col, a.row);
  }
  const minC = Math.min(a.col, b.col);
  const maxC = Math.max(a.col, b.col);
  const minR = Math.min(a.row, b.row);
  const maxR = Math.max(a.row, b.row);
  return `${cellId(minC, minR)}:${cellId(maxC, maxR)}`;
}

function spliceString(str: string, start: number, end: number, insert: string): string {
  return str.slice(0, start) + insert + str.slice(end);
}

// ── App ────────────────────────────────────────────────────────────────

export function App() {
  const [sheet, setSheet] = useState<SheetData>(loadSheet);
  const [selectedCell, setSelectedCell] = useState<string | null>(null);
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

  useEffect(() => { saveSheet(sheet); }, [sheet]);

  useEffect(() => {
    if (pendingCursorPos.current !== null && editInputRef.current) {
      const pos = pendingCursorPos.current;
      editInputRef.current.setSelectionRange(pos, pos);
      pendingCursorPos.current = null;
    }
  });

  const displayValues = useMemo(() => {
    const result: Record<string, string> = {};
    for (const id of Object.keys(sheet.cells)) {
      result[id] = computeDisplay(id, sheet.cells);
    }
    return result;
  }, [sheet.cells]);

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

  const updateCell = useCallback((id: string, value: string) => {
    pushUndo();
    setSheet((prev) => {
      const next = { ...prev, cells: { ...prev.cells } };
      if (value === "") {
        delete next.cells[id];
      } else {
        next.cells[id] = value;
      }
      return next;
    });
  }, [pushUndo]);

  const exitPointMode = useCallback(() => { setPointMode(null); }, []);

  const cancelEdit = useCallback(() => {
    setEditingCell(null);
    setPointMode(null);
    setEditValue("");
  }, []);

  const commitFormulaBar = useCallback(() => {
    if (selectedCell) {
      pushUndo();
      setSheet((prev) => {
        const next = { ...prev, cells: { ...prev.cells } };
        if (formulaBarValue === "") {
          delete next.cells[selectedCell];
        } else {
          next.cells[selectedCell] = formulaBarValue;
        }
        return next;
      });
    }
  }, [selectedCell, formulaBarValue, pushUndo]);

  const startEditing = useCallback((id: string, initialValue?: string) => {
    setEditingCell(id);
    setEditValue(initialValue ?? sheet.cells[id] ?? "");
    setPointMode(null);
  }, [sheet.cells]);

  const selectCell = useCallback((id: string) => {
    setSelectedCell(id);
    setFormulaBarValue(sheet.cells[id] ?? "");
  }, [sheet.cells]);

  const moveAndSelect = useCallback((dCol: number, dRow: number) => {
    const base = selectedCell ?? "A1";
    const parsed = parseCellRef(base);
    if (!parsed) return;
    const newCol = clampCol(parsed.col + dCol, sheet.colCount);
    const newRow = clampRow(parsed.row + dRow, sheet.rowCount);
    const newId = cellId(newCol, newRow);
    selectCell(newId);
  }, [selectedCell, sheet.colCount, sheet.rowCount, selectCell]);

  const commitAndMove = useCallback((dCol: number, dRow: number) => {
    if (editingCell) {
      updateCell(editingCell, editValue);
      setEditingCell(null);
      setPointMode(null);
    }
    moveAndSelect(dCol, dRow);
  }, [editingCell, editValue, updateCell, moveAndSelect]);

  // Sync formula bar when not editing from it
  useEffect(() => {
    if (selectedCell && !formulaBarFocused && !editingCell) {
      setFormulaBarValue(sheet.cells[selectedCell] ?? "");
    }
  }, [selectedCell, sheet.cells, formulaBarFocused, editingCell]);

  // Sync formula bar with inline edit
  useEffect(() => {
    if (editingCell) {
      setFormulaBarValue(editValue);
    }
  }, [editingCell, editValue]);

  // ── Point mode: insert or move a cell reference ─────────────────────

  const enterOrMovePointMode = useCallback((dCol: number, dRow: number, extend: boolean) => {
    if (!editingCell) return;
    const editCell = parseCellRef(editingCell);
    if (!editCell) return;

    if (pointMode) {
      const newActive = {
        col: clampCol(pointMode.active.col + dCol, sheet.colCount),
        row: clampRow(pointMode.active.row + dRow, sheet.rowCount),
      };
      const newAnchor = extend ? pointMode.anchor : newActive;
      const newPM: PointMode = { ...pointMode, anchor: newAnchor, active: newActive };
      const ref = buildRefString(newPM);
      const newVal = spliceString(editValue, pointMode.refStart, pointMode.refEnd, ref);
      const newEnd = pointMode.refStart + ref.length;
      setPointMode({ ...newPM, refEnd: newEnd });
      setEditValue(newVal);
      pendingCursorPos.current = newEnd;
    } else {
      const cursorPos = editInputRef.current?.selectionStart ?? editValue.length;
      if (!isRefPosition(editValue, cursorPos)) return;

      const targetCol = clampCol(editCell.col + dCol, sheet.colCount);
      const targetRow = clampRow(editCell.row + dRow, sheet.rowCount);
      const ref = cellId(targetCol, targetRow);
      const newVal = spliceString(editValue, cursorPos, cursorPos, ref);
      const newEnd = cursorPos + ref.length;
      setPointMode({
        refStart: cursorPos,
        refEnd: newEnd,
        anchor: { col: targetCol, row: targetRow },
        active: { col: targetCol, row: targetRow },
      });
      setEditValue(newVal);
      pendingCursorPos.current = newEnd;
    }
  }, [editingCell, editValue, pointMode, sheet.colCount, sheet.rowCount]);

  // ── Click-to-insert-ref during formula editing ──────────────────────

  const handleCellClick = useCallback((id: string, shiftKey: boolean) => {
    if (editingCell && editValue.startsWith("=") && id !== editingCell) {
      const clicked = parseCellRef(id);
      if (!clicked) return;

      if (shiftKey && pointMode) {
        const newPM: PointMode = { ...pointMode, active: clicked };
        const ref = buildRefString(newPM);
        const newVal = spliceString(editValue, pointMode.refStart, pointMode.refEnd, ref);
        const newEnd = pointMode.refStart + ref.length;
        setPointMode({ ...newPM, refEnd: newEnd });
        setEditValue(newVal);
        pendingCursorPos.current = newEnd;
      } else {
        const cursorPos = editInputRef.current?.selectionStart ?? editValue.length;
        const insertAt = pointMode ? pointMode.refStart : cursorPos;
        const removeEnd = pointMode ? pointMode.refEnd : cursorPos;
        const ref = id;
        const newVal = spliceString(editValue, insertAt, removeEnd, ref);
        const newEnd = insertAt + ref.length;
        setPointMode({
          refStart: insertAt,
          refEnd: newEnd,
          anchor: clicked,
          active: clicked,
        });
        setEditValue(newVal);
        pendingCursorPos.current = newEnd;
      }
      editInputRef.current?.focus();
      return;
    }

    if (editingCell && editingCell !== id) {
      updateCell(editingCell, editValue);
      setEditingCell(null);
      setPointMode(null);
    }
    selectCell(id);
  }, [editingCell, editValue, pointMode, selectCell, updateCell]);

  // ── Keyboard: grid-level (when not editing inline) ──────────────────

  const handleGridKeyDown = useCallback((e: React.KeyboardEvent) => {
    const meta = e.metaKey || e.ctrlKey;

    // Undo
    if (meta && e.key === "z") {
      e.preventDefault();
      undo();
      return;
    }

    // Copy / Cut / Paste
    if (meta && e.key === "c" && selectedCell && !editingCell) {
      e.preventDefault();
      setClipboard({ type: "copy", id: selectedCell, value: sheet.cells[selectedCell] ?? "" });
      return;
    }
    if (meta && e.key === "x" && selectedCell && !editingCell) {
      e.preventDefault();
      setClipboard({ type: "cut", id: selectedCell, value: sheet.cells[selectedCell] ?? "" });
      return;
    }
    if (meta && e.key === "v" && selectedCell && !editingCell && clipboard) {
      e.preventDefault();
      updateCell(selectedCell, clipboard.value);
      if (clipboard.type === "cut") {
        updateCell(clipboard.id, "");
        setClipboard(null);
      }
      return;
    }

    // If editing, delegate to inline handler
    if (editingCell) {
      handleEditKeyDown(e);
      return;
    }

    // Navigation
    if (e.key === "ArrowUp") { e.preventDefault(); moveAndSelect(0, -1); }
    else if (e.key === "ArrowDown") { e.preventDefault(); moveAndSelect(0, 1); }
    else if (e.key === "ArrowLeft") { e.preventDefault(); moveAndSelect(-1, 0); }
    else if (e.key === "ArrowRight") { e.preventDefault(); moveAndSelect(1, 0); }
    else if (e.key === "Tab") { e.preventDefault(); moveAndSelect(e.shiftKey ? -1 : 1, 0); }
    else if (e.key === "Enter") {
      e.preventDefault();
      if (selectedCell) startEditing(selectedCell);
    }
    else if (e.key === "F2") {
      e.preventDefault();
      if (selectedCell) startEditing(selectedCell);
    }
    else if (e.key === "Delete" || e.key === "Backspace") {
      if (selectedCell) {
        updateCell(selectedCell, "");
        setFormulaBarValue("");
      }
    }
    else if (e.key.length === 1 && !meta) {
      if (selectedCell) {
        startEditing(selectedCell, e.key);
        e.preventDefault();
      }
    }
  }, [editingCell, selectedCell, clipboard, sheet.cells, moveAndSelect, startEditing, updateCell, undo]);

  // ── Keyboard: inline cell editing ───────────────────────────────────

  const handleEditKeyDown = useCallback((e: React.KeyboardEvent) => {
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
      return;
    }

    // Formula point mode: arrow keys insert/move refs
    if (editValue.startsWith("=") && (
      e.key === "ArrowUp" || e.key === "ArrowDown" ||
      e.key === "ArrowLeft" || e.key === "ArrowRight"
    )) {
      const inPointMode = !!pointMode;
      const cursorPos = editInputRef.current?.selectionStart ?? editValue.length;
      if (inPointMode || isRefPosition(editValue, cursorPos)) {
        e.preventDefault();
        e.stopPropagation();
        const dCol = e.key === "ArrowLeft" ? -1 : e.key === "ArrowRight" ? 1 : 0;
        const dRow = e.key === "ArrowUp" ? -1 : e.key === "ArrowDown" ? 1 : 0;
        enterOrMovePointMode(dCol, dRow, e.shiftKey);
        return;
      }
    }

    // Any non-navigation key exits point mode
    if (pointMode && e.key !== "Shift") {
      exitPointMode();
    }
  }, [editValue, pointMode, commitAndMove, cancelEdit, enterOrMovePointMode, exitPointMode]);

  const handleFormulaBarKeyDown = useCallback((e: React.KeyboardEvent) => {
    if (e.key === "Enter") {
      e.preventDefault();
      commitFormulaBar();
      formulaBarRef.current?.blur();
      gridRef.current?.focus();
    } else if (e.key === "Escape") {
      if (selectedCell) setFormulaBarValue(sheet.cells[selectedCell] ?? "");
      formulaBarRef.current?.blur();
      gridRef.current?.focus();
    }
  }, [commitFormulaBar, selectedCell, sheet.cells]);

  useEffect(() => {
    if (editingCell && editInputRef.current) {
      editInputRef.current.focus();
    }
  }, [editingCell]);

  const clearAll = useCallback(() => {
    pushUndo();
    setSheet({ cells: {}, colCount: DEFAULT_COLS, rowCount: DEFAULT_ROWS });
    setSelectedCell(null);
    setEditingCell(null);
    setEditValue("");
    setFormulaBarValue("");
    setPointMode(null);
  }, [pushUndo]);

  // ── Point mode highlight set ────────────────────────────────────────

  const pointHighlight = useMemo<Set<string>>(() => {
    if (!pointMode) return new Set();
    const a = pointMode.anchor;
    const b = pointMode.active;
    const minC = Math.min(a.col, b.col);
    const maxC = Math.max(a.col, b.col);
    const minR = Math.min(a.row, b.row);
    const maxR = Math.max(a.row, b.row);
    const s = new Set<string>();
    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        s.add(cellId(c, r));
      }
    }
    return s;
  }, [pointMode]);

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
        <span
          style={{
            fontFamily: "'Fraunces', Georgia, serif",
            fontWeight: 700,
            fontSize: "1.125rem",
            whiteSpace: "nowrap",
          }}
        >
          spreadsheet
        </span>
        <div style={{ flex: 1 }} />
        <button
          onClick={clearAll}
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
          Clear all
        </button>
        <a
          href="https://freeappstore.online"
          target="_blank"
          rel="noopener noreferrer"
          style={{
            color: "var(--color-muted)",
            textDecoration: "none",
            fontSize: "0.75rem",
            whiteSpace: "nowrap",
          }}
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
        <span
          style={{
            fontWeight: 600,
            fontSize: "0.8125rem",
            color: "var(--color-accent)",
            minWidth: "3rem",
            textAlign: "center",
          }}
        >
          {selectedCell ?? ""}
        </span>
        <span style={{ color: "var(--color-line)" }}>|</span>
        <input
          ref={formulaBarRef}
          value={formulaBarValue}
          onChange={(e) => {
            setFormulaBarValue(e.target.value);
            if (editingCell) setEditValue(e.target.value);
          }}
          onFocus={() => setFormulaBarFocused(true)}
          onBlur={() => {
            setFormulaBarFocused(false);
            if (!editingCell) commitFormulaBar();
          }}
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
        style={{
          flex: 1,
          overflow: "auto",
          outline: "none",
          position: "relative",
        }}
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
          <div
            style={{
              position: "sticky",
              top: 0,
              left: 0,
              zIndex: 3,
              background: "var(--color-panel)",
              borderBottom: "2px solid var(--color-line)",
              borderRight: "1px solid var(--color-line)",
            }}
          />

          {/* Column headers */}
          {Array.from({ length: sheet.colCount }, (_, c) => (
            <div
              key={`ch-${c}`}
              style={{
                position: "sticky",
                top: 0,
                zIndex: 2,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                background: "var(--color-panel)",
                borderBottom: "2px solid var(--color-line)",
                borderRight: "1px solid var(--color-line)",
                fontSize: "0.6875rem",
                fontWeight: 600,
                color: "var(--color-muted)",
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
                  position: "sticky",
                  left: 0,
                  zIndex: 1,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  background: "var(--color-panel)",
                  borderBottom: "1px solid var(--color-line)",
                  borderRight: "1px solid var(--color-line)",
                  fontSize: "0.6875rem",
                  fontWeight: 500,
                  color: "var(--color-muted)",
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
                    onClick={(e) => handleCellClick(id, e.shiftKey)}
                    onDoubleClick={() => startEditing(id)}
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
                      background: isPointed && !isSelected
                        ? "rgba(91, 140, 214, 0.1)"
                        : undefined,
                      zIndex: isSelected ? 1 : isPointed ? 1 : 0,
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
                        onKeyDown={(e) => {
                          handleEditKeyDown(e);
                        }}
                        style={{
                          position: "absolute",
                          inset: 0,
                          border: "none",
                          outline: "none",
                          background: "var(--color-paper)",
                          padding: "0 4px",
                          fontSize: "0.8125rem",
                          color: "var(--color-ink)",
                          width: "100%",
                          height: "100%",
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
                          textAlign:
                            !isNaN(parseFloat(display)) && display !== "" && !isError(display)
                              ? "right"
                              : "left",
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
