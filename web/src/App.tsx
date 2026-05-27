import { useState, useEffect, useCallback, useRef, useMemo, useLayoutEffect } from "react";
import type { CellFormat, Sheet, WorkbookData, PointMode } from "./types";
import { DEFAULT_COLS, DEFAULT_ROWS, MAX_UNDO } from "./constants";
import { colLabel, cellId, parseCellRef, clamp } from "./lib/cell-refs";
import { computeDisplay } from "./lib/formula-engine";
import { parseCsvLine } from "./lib/csv";
import { isRefPosition, buildRefString, splice } from "./lib/point-mode";
import { makeSheet, loadWorkbook, saveWorkbook } from "./lib/persistence";

// ── App ────────────────────────────────────────────────────────────────

export function App() {
  const [workbook, setWorkbook] = useState<WorkbookData>(loadWorkbook);
  const sheet = workbook.sheets[workbook.activeSheet]!;
  const setSheet = useCallback((updater: Sheet | ((prev: Sheet) => Sheet)) => {
    setWorkbook((wb) => {
      const sheets = [...wb.sheets];
      const idx = wb.activeSheet;
      sheets[idx] = typeof updater === "function" ? updater(sheets[idx]!) : updater;
      return { ...wb, sheets };
    });
  }, []);
  const [selectedCell, setSelectedCell] = useState<string>("A1");
  const [selectionEnd, setSelectionEnd] = useState<string | null>(null);
  const [editingCell, setEditingCell] = useState<string | null>(null);
  const [editValue, setEditValue] = useState("");
  const [formulaBarValue, setFormulaBarValue] = useState("");
  const [formulaBarFocused, setFormulaBarFocused] = useState(false);
  const [pointMode, setPointMode] = useState<PointMode | null>(null);
  const [clipboard, setClipboard] = useState<{ type: "copy" | "cut"; ids: string[]; values: Record<string, string> } | null>(null);
  const [colWidths, setColWidths] = useState<Record<number, number>>({});
  const [contextMenu, setContextMenu] = useState<{ x: number; y: number; cellId: string } | null>(null);
  const [findBar, setFindBar] = useState<{ open: boolean; query: string; replace: string; showReplace: boolean }>({ open: false, query: "", replace: "", showReplace: false });
  const [showHelp, setShowHelp] = useState(false);
  const [colorPicker, setColorPicker] = useState<{ type: "bg" | "color"; x: number; y: number } | null>(null);
  const [findMatches, setFindMatches] = useState<string[]>([]);
  const [findIndex, setFindIndex] = useState(0);

  const editInputRef = useRef<HTMLInputElement>(null);
  const formulaBarRef = useRef<HTMLInputElement>(null);
  const gridRef = useRef<HTMLDivElement>(null);
  const findInputRef = useRef<HTMLInputElement>(null);
  const pendingCursorPos = useRef<number | null>(null);
  const resizingCol = useRef<{ col: number; startX: number; startW: number } | null>(null);
  const undoStackRef = useRef<{ cells: Record<string, string>; formats: Record<string, CellFormat> }[]>([]);
  const redoStackRef = useRef<{ cells: Record<string, string>; formats: Record<string, CellFormat> }[]>([]);

  // ── Persistence ─────────────────────────────────────────────────────

  useEffect(() => { saveWorkbook(workbook); }, [workbook]);

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

  // Close popups on Escape or outside click
  useEffect(() => {
    const handleKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        if (colorPicker) { setColorPicker(null); e.stopPropagation(); }
        if (contextMenu) { setContextMenu(null); e.stopPropagation(); }
        if (showHelp) { setShowHelp(false); e.stopPropagation(); }
      }
    };
    const handleClick = () => {
      setColorPicker(null);
      setContextMenu(null);
    };
    window.addEventListener("keydown", handleKey, true);
    window.addEventListener("mousedown", handleClick);
    return () => {
      window.removeEventListener("keydown", handleKey, true);
      window.removeEventListener("mousedown", handleClick);
    };
  }, [colorPicker, contextMenu, showHelp]);

  // ── Derived state ───────────────────────────────────────────────────

  const displayValues = useMemo(() => {
    const result: Record<string, string> = {};
    for (const id of Object.keys(sheet.cells)) {
      result[id] = computeDisplay(id, sheet.cells);
    }
    return result;
  }, [sheet.cells]);

  const selectionRange = useMemo<{ minC: number; maxC: number; minR: number; maxR: number; cells: Set<string> } | null>(() => {
    if (!selectionEnd) return null;
    const a = parseCellRef(selectedCell);
    const b = parseCellRef(selectionEnd);
    if (!a || !b) return null;
    const minC = Math.min(a.col, b.col), maxC = Math.max(a.col, b.col);
    const minR = Math.min(a.row, b.row), maxR = Math.max(a.row, b.row);
    const s = new Set<string>();
    for (let r = minR; r <= maxR; r++)
      for (let c = minC; c <= maxC; c++)
        s.add(cellId(c, r));
    return { minC, maxC, minR, maxR, cells: s };
  }, [selectedCell, selectionEnd]);

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

  // Status bar: SUM / AVG / COUNT for selection
  const selectionStats = useMemo(() => {
    const ids = selectionRange ? [...selectionRange.cells] : [selectedCell];
    const nums: number[] = [];
    let count = 0;
    for (const id of ids) {
      const d = displayValues[id] ?? "";
      if (d === "") continue;
      count++;
      const n = parseFloat(d);
      if (!isNaN(n)) nums.push(n);
    }
    if (nums.length === 0) return { sum: 0, avg: 0, count, numCount: 0 };
    const sum = nums.reduce((a, b) => a + b, 0);
    return { sum, avg: sum / nums.length, count, numCount: nums.length };
  }, [selectedCell, selectionRange, displayValues]);

  // ── Undo / Redo ─────────────────────────────────────────────────────

  const pushUndo = useCallback(() => {
    undoStackRef.current = [...undoStackRef.current, { cells: { ...sheet.cells }, formats: { ...sheet.formats } }];
    if (undoStackRef.current.length > MAX_UNDO) undoStackRef.current.shift();
    redoStackRef.current = [];
  }, [sheet.cells, sheet.formats]);

  const undo = useCallback(() => {
    if (undoStackRef.current.length === 0) return;
    redoStackRef.current = [...redoStackRef.current, { cells: { ...sheet.cells }, formats: { ...sheet.formats } }];
    const restored = undoStackRef.current.pop()!;
    setSheet((s) => ({ ...s, cells: restored.cells, formats: restored.formats }));
  }, [sheet.cells, sheet.formats, setSheet]);

  const redo = useCallback(() => {
    if (redoStackRef.current.length === 0) return;
    undoStackRef.current = [...undoStackRef.current, { cells: { ...sheet.cells }, formats: { ...sheet.formats } }];
    const restored = redoStackRef.current.pop()!;
    setSheet((s) => ({ ...s, cells: restored.cells, formats: restored.formats }));
  }, [sheet.cells, sheet.formats, setSheet]);

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
    setSelectionEnd(null);
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

  const extendSelection = useCallback((dCol: number, dRow: number) => {
    const end = selectionEnd ?? selectedCell;
    const parsed = parseCellRef(end);
    if (!parsed) return;
    const newId = cellId(
      clamp(parsed.col + dCol, sheet.colCount),
      clamp(parsed.row + dRow, sheet.rowCount),
    );
    setSelectionEnd(newId);
  }, [selectedCell, selectionEnd, sheet.colCount, sheet.rowCount]);

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

    if (e.shiftKey && !editingCell) {
      setSelectionEnd(id);
    } else {
      select(id);
    }
    if (!editingCell) {
      gridRef.current?.focus();
    }
  }, [editingCell, editValue, insertRefByClick, commitEdit, select]);

  // ── Keyboard: grid (navigation when not editing) ────────────────────

  const handleGridKeyDown = useCallback((e: React.KeyboardEvent) => {
    // When editing, the input captures keys via stopPropagation.
    // This handler only fires when the grid itself has focus (not editing).
    const meta = e.metaKey || e.ctrlKey;

    if (meta && e.key === "z" && !e.shiftKey) { e.preventDefault(); undo(); return; }
    if (meta && (e.key === "y" || (e.key === "z" && e.shiftKey))) { e.preventDefault(); redo(); return; }
    if (meta && e.key === "f") { e.preventDefault(); setFindBar((p) => ({ ...p, open: true })); requestAnimationFrame(() => findInputRef.current?.focus()); return; }
    if (meta && e.key === "h") { e.preventDefault(); setFindBar((p) => ({ ...p, open: true, showReplace: true })); requestAnimationFrame(() => findInputRef.current?.focus()); return; }
    if (meta && e.key === "b") {
      e.preventDefault();
      const cur = sheet.formats[selectedCell]?.bold;
      const ids = selectionRange ? [...selectionRange.cells] : [selectedCell];
      setSheet((prev) => { const f = { ...prev.formats }; for (const id of ids) { f[id] = { ...f[id], bold: !cur }; } return { ...prev, formats: f }; });
      return;
    }
    if (meta && e.key === "i") {
      e.preventDefault();
      const cur = sheet.formats[selectedCell]?.italic;
      const ids = selectionRange ? [...selectionRange.cells] : [selectedCell];
      setSheet((prev) => { const f = { ...prev.formats }; for (const id of ids) { f[id] = { ...f[id], italic: !cur }; } return { ...prev, formats: f }; });
      return;
    }
    if (meta && e.key === "/") { e.preventDefault(); setShowHelp((p) => !p); return; }

    if (meta && e.key === "c") {
      e.preventDefault();
      const ids = selectionRange ? [...selectionRange.cells] : [selectedCell];
      const vals: Record<string, string> = {};
      for (const id of ids) { const v = sheet.cells[id]; if (v) vals[id] = v; }
      setClipboard({ type: "copy", ids, values: vals });
      return;
    }
    if (meta && e.key === "x") {
      e.preventDefault();
      const ids = selectionRange ? [...selectionRange.cells] : [selectedCell];
      const vals: Record<string, string> = {};
      for (const id of ids) { const v = sheet.cells[id]; if (v) vals[id] = v; }
      setClipboard({ type: "cut", ids, values: vals });
      return;
    }
    if (meta && e.key === "v" && clipboard) {
      e.preventDefault();
      const target = parseCellRef(selectedCell);
      const first = parseCellRef(clipboard.ids[0] ?? "A1");
      if (target && first) {
        pushUndo();
        const dC = target.col - first.col, dR = target.row - first.row;
        setSheet((prev) => {
          const next = { ...prev, cells: { ...prev.cells } };
          for (const [srcId, val] of Object.entries(clipboard.values)) {
            const src = parseCellRef(srcId);
            if (src) {
              const destId = cellId(src.col + dC, src.row + dR);
              next.cells[destId] = val;
            }
          }
          if (clipboard.type === "cut") {
            for (const id of clipboard.ids) delete next.cells[id];
          }
          return next;
        });
        if (clipboard.type === "cut") setClipboard(null);
      }
      return;
    }
    if (meta && e.key === "a") {
      e.preventDefault();
      setSelectionEnd(cellId(sheet.colCount - 1, sheet.rowCount - 1));
      return;
    }

    switch (e.key) {
      case "ArrowUp":    e.preventDefault(); e.shiftKey ? extendSelection(0, -1) : move(0, -1); break;
      case "ArrowDown":  e.preventDefault(); e.shiftKey ? extendSelection(0, 1) : move(0, 1); break;
      case "ArrowLeft":  e.preventDefault(); e.shiftKey ? extendSelection(-1, 0) : move(-1, 0); break;
      case "ArrowRight": e.preventDefault(); e.shiftKey ? extendSelection(1, 0) : move(1, 0); break;
      case "Tab":        e.preventDefault(); move(e.shiftKey ? -1 : 1, 0); break;
      case "Enter":
      case "F2":
        e.preventDefault();
        startEditing(selectedCell);
        break;
      case "Delete":
      case "Backspace": {
        const ids = selectionRange ? [...selectionRange.cells] : [selectedCell];
        pushUndo();
        setSheet((prev) => {
          const next = { ...prev, cells: { ...prev.cells } };
          for (const id of ids) delete next.cells[id];
          return next;
        });
        setFormulaBarValue("");
        break;
      }
      default:
        // Any printable character starts editing with that character
        if (e.key.length === 1 && !meta) {
          e.preventDefault();
          startEditing(selectedCell, e.key);
        }
    }
  }, [selectedCell, selectionRange, sheet, clipboard, move, extendSelection, startEditing, writeCell, pushUndo, undo, redo, setSheet]);

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

  // ── Context menu actions ─────────────────────────────────────────────

  const shiftEntries = useCallback(<T,>(map: Record<string, T>, test: (p: { col: number; row: number }) => string | null): Record<string, T> => {
    const result: Record<string, T> = {};
    for (const [id, val] of Object.entries(map)) {
      const p = parseCellRef(id);
      if (!p) continue;
      const newId = test(p);
      if (newId) result[newId] = val;
    }
    return result;
  }, []);

  const insertRow = useCallback((row: number, direction: "above" | "below") => {
    const targetRow = direction === "below" ? row + 1 : row;
    pushUndo();
    setSheet((prev) => {
      const shift = (p: { col: number; row: number }) => p.row >= targetRow ? cellId(p.col, p.row + 1) : cellId(p.col, p.row);
      return { ...prev, cells: shiftEntries(prev.cells, shift), formats: shiftEntries(prev.formats, shift), rowCount: prev.rowCount + 1 };
    });
  }, [pushUndo, setSheet, shiftEntries]);

  const deleteRow = useCallback((row: number) => {
    pushUndo();
    setSheet((prev) => {
      const shift = (p: { col: number; row: number }) => p.row === row ? null : p.row > row ? cellId(p.col, p.row - 1) : cellId(p.col, p.row);
      return { ...prev, cells: shiftEntries(prev.cells, shift), formats: shiftEntries(prev.formats, shift), rowCount: Math.max(1, prev.rowCount - 1) };
    });
  }, [pushUndo, setSheet, shiftEntries]);

  const insertCol = useCallback((col: number, direction: "left" | "right") => {
    const targetCol = direction === "right" ? col + 1 : col;
    pushUndo();
    setSheet((prev) => {
      const shift = (p: { col: number; row: number }) => p.col >= targetCol ? cellId(p.col + 1, p.row) : cellId(p.col, p.row);
      return { ...prev, cells: shiftEntries(prev.cells, shift), formats: shiftEntries(prev.formats, shift), colCount: prev.colCount + 1 };
    });
  }, [pushUndo, setSheet, shiftEntries]);

  const deleteCol = useCallback((col: number) => {
    pushUndo();
    setSheet((prev) => {
      const shift = (p: { col: number; row: number }) => p.col === col ? null : p.col > col ? cellId(p.col - 1, p.row) : cellId(p.col, p.row);
      return { ...prev, cells: shiftEntries(prev.cells, shift), formats: shiftEntries(prev.formats, shift), colCount: Math.max(1, prev.colCount - 1) };
    });
  }, [pushUndo, setSheet, shiftEntries]);

  const sortByColumn = useCallback((col: number, ascending: boolean) => {
    pushUndo();
    setSheet((prev) => {
      const rows: { row: number; sortVal: string }[] = [];
      for (let r = 0; r < prev.rowCount; r++) {
        const id = cellId(col, r);
        const raw = prev.cells[id] ?? "";
        const display = raw.startsWith("=") ? computeDisplay(id, prev.cells) : raw;
        rows.push({ row: r, sortVal: display });
      }
      rows.sort((a, b) => {
        const na = parseFloat(a.sortVal), nb = parseFloat(b.sortVal);
        const aEmpty = a.sortVal === "", bEmpty = b.sortVal === "";
        if (aEmpty && bEmpty) return 0;
        if (aEmpty) return 1;
        if (bEmpty) return -1;
        if (!isNaN(na) && !isNaN(nb)) return ascending ? na - nb : nb - na;
        return ascending ? a.sortVal.localeCompare(b.sortVal) : b.sortVal.localeCompare(a.sortVal);
      });
      const nextCells: Record<string, string> = {};
      const nextFormats: Record<string, CellFormat> = {};
      for (let newR = 0; newR < rows.length; newR++) {
        const oldR = rows[newR]!.row;
        for (let c = 0; c < prev.colCount; c++) {
          const oldId = cellId(c, oldR);
          const newId = cellId(c, newR);
          const val = prev.cells[oldId];
          if (val) nextCells[newId] = val;
          const fmt = prev.formats[oldId];
          if (fmt) nextFormats[newId] = fmt;
        }
      }
      return { ...prev, cells: nextCells, formats: nextFormats };
    });
    setContextMenu(null);
  }, [pushUndo]);

  // ── Find / Replace ─────────────────────────────────────────────────

  useEffect(() => {
    if (!findBar.open || !findBar.query) { setFindMatches([]); return; }
    const q = findBar.query.toLowerCase();
    const matches: string[] = [];
    for (const [id, raw] of Object.entries(sheet.cells)) {
      const display = displayValues[id] ?? raw;
      if (raw.toLowerCase().includes(q) || display.toLowerCase().includes(q)) {
        matches.push(id);
      }
    }
    matches.sort((a, b) => {
      const pa = parseCellRef(a)!, pb = parseCellRef(b)!;
      return pa.row !== pb.row ? pa.row - pb.row : pa.col - pb.col;
    });
    setFindMatches(matches);
    setFindIndex(0);
    if (matches.length > 0) select(matches[0]!);
  }, [findBar.open, findBar.query, sheet.cells, displayValues, select]);

  const findNext = useCallback(() => {
    if (findMatches.length === 0) return;
    const next = (findIndex + 1) % findMatches.length;
    setFindIndex(next);
    select(findMatches[next]!);
  }, [findMatches, findIndex, select]);

  const findPrev = useCallback(() => {
    if (findMatches.length === 0) return;
    const prev = (findIndex - 1 + findMatches.length) % findMatches.length;
    setFindIndex(prev);
    select(findMatches[prev]!);
  }, [findMatches, findIndex, select]);

  const replaceOne = useCallback(() => {
    if (findMatches.length === 0) return;
    const id = findMatches[findIndex]!;
    const raw = sheet.cells[id] ?? "";
    const q = findBar.query;
    const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi");
    writeCell(id, raw.replace(re, findBar.replace));
  }, [findMatches, findIndex, findBar.query, findBar.replace, sheet.cells, writeCell]);

  const replaceAll = useCallback(() => {
    if (!findBar.query) return;
    pushUndo();
    const escaped = findBar.query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    setSheet((prev) => {
      const next = { ...prev, cells: { ...prev.cells } };
      for (const [id, raw] of Object.entries(next.cells)) {
        const replaced = raw.replace(new RegExp(escaped, "gi"), findBar.replace);
        if (replaced !== raw) next.cells[id] = replaced;
      }
      return next;
    });
  }, [findBar.query, findBar.replace, pushUndo, setSheet]);

  const closeFindBar = useCallback(() => {
    setFindBar({ open: false, query: "", replace: "", showReplace: false });
    setFindMatches([]);
    gridRef.current?.focus();
  }, []);

  // ── Clear all ───────────────────────────────────────────────────────

  const clearAll = useCallback(() => {
    if (editingCell) cancelEdit();
    pushUndo();
    setSheet((s) => ({ ...s, cells: {}, formats: {} }));
    setSelectedCell("A1");
    setFormulaBarValue("");
    gridRef.current?.focus();
  }, [editingCell, cancelEdit, pushUndo, setSheet]);

  // ── Cell formatting ─────────────────────────────────────────────────

  const applyFormat = useCallback((update: Partial<CellFormat>) => {
    const ids = selectionRange ? [...selectionRange.cells] : [selectedCell];
    setSheet((prev) => {
      const formats = { ...prev.formats };
      for (const id of ids) {
        const existing = formats[id] ?? {};
        const merged = { ...existing, ...update };
        const isEmpty = !merged.bold && !merged.italic && !merged.align && !merged.bg && !merged.color;
        if (isEmpty) delete formats[id];
        else formats[id] = merged;
      }
      return { ...prev, formats };
    });
  }, [selectedCell, selectionRange, setSheet]);

  const toggleBold = useCallback(() => {
    const current = sheet.formats[selectedCell]?.bold;
    applyFormat({ bold: !current });
  }, [selectedCell, sheet.formats, applyFormat]);

  const toggleItalic = useCallback(() => {
    const current = sheet.formats[selectedCell]?.italic;
    applyFormat({ italic: !current });
  }, [selectedCell, sheet.formats, applyFormat]);

  // ── Sheet tabs ──────────────────────────────────────────────────────

  const addSheet = useCallback(() => {
    setWorkbook((wb) => {
      const name = `Sheet ${wb.sheets.length + 1}`;
      return { sheets: [...wb.sheets, makeSheet(name)], activeSheet: wb.sheets.length };
    });
    setSelectedCell("A1");
    setSelectionEnd(null);
    setFormulaBarValue("");
  }, []);

  const switchSheet = useCallback((index: number) => {
    if (editingCell) commitEdit();
    setWorkbook((wb) => ({ ...wb, activeSheet: index }));
    setSelectedCell("A1");
    setSelectionEnd(null);
    setFormulaBarValue("");
  }, [editingCell, commitEdit]);

  const renameSheet = useCallback((index: number) => {
    const current = workbook.sheets[index]?.name ?? "";
    const name = prompt("Sheet name:", current);
    if (name && name !== current) {
      setWorkbook((wb) => {
        const sheets = [...wb.sheets];
        sheets[index] = { ...sheets[index]!, name };
        return { ...wb, sheets };
      });
    }
  }, [workbook.sheets]);

  const deleteSheet = useCallback((index: number) => {
    if (workbook.sheets.length <= 1) return;
    setWorkbook((wb) => {
      const sheets = wb.sheets.filter((_, i) => i !== index);
      const activeSheet = wb.activeSheet >= sheets.length ? sheets.length - 1 : wb.activeSheet;
      return { sheets, activeSheet };
    });
    setSelectedCell("A1");
    setSelectionEnd(null);
  }, [workbook.sheets.length]);

  // ── Column resize ───────────────────────────────────────────────────

  const getColWidth = useCallback((c: number) => colWidths[c] ?? COL_WIDTH, [colWidths]);

  const handleColResizeStart = useCallback((col: number, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startW = colWidths[col] ?? COL_WIDTH;
    resizingCol.current = { col, startX, startW };

    const onMove = (ev: MouseEvent) => {
      if (!resizingCol.current) return;
      const diff = ev.clientX - resizingCol.current.startX;
      const newW = Math.max(40, resizingCol.current.startW + diff);
      setColWidths((prev) => ({ ...prev, [resizingCol.current!.col]: newW }));
    };
    const onUp = () => {
      resizingCol.current = null;
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
  }, [colWidths]);

  const handleColResizeDoubleClick = useCallback((col: number) => {
    let maxW = 50;
    for (let r = 0; r < sheet.rowCount; r++) {
      const id = cellId(col, r);
      const d = displayValues[id] ?? "";
      maxW = Math.max(maxW, d.length * 8 + 16);
    }
    setColWidths((prev) => ({ ...prev, [col]: Math.min(maxW, 400) }));
  }, [sheet.rowCount, displayValues]);

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
        setSheet((prev) => ({
          ...prev,
          cells,
          formats: {},
          colCount: Math.max(DEFAULT_COLS, maxCol + 1),
          rowCount: Math.max(DEFAULT_ROWS, lines.length),
        }));
        setSelectedCell("A1");
        setFormulaBarValue(cells["A1"] ?? "");
      };
      reader.readAsText(file);
    };
    input.click();
  }, [pushUndo]);

  // ── Number formatting ────────────────────────────────────────────────

  const formatDisplay = useCallback((val: string): string => {
    if (!val || val.startsWith("#")) return val;
    const num = parseFloat(val);
    if (isNaN(num) || val !== String(num)) return val;
    if (Number.isInteger(num) && Math.abs(num) >= 1000) {
      return num.toLocaleString("en-US");
    }
    if (!Number.isInteger(num)) {
      return num.toLocaleString("en-US", { maximumFractionDigits: 10 });
    }
    return val;
  }, []);

  // ── Virtualization ───────────────────────────────────────────────────

  const [scrollTop, setScrollTop] = useState(0);
  const [gridHeight, setGridHeight] = useState(600);
  const BUFFER = 5;
  const COL_WIDTH = 100;
  const ROW_HEIGHT = 28;
  const HEADER_WIDTH = 44;

  const firstVisible = Math.floor(scrollTop / ROW_HEIGHT);
  const visibleCount = Math.ceil(gridHeight / ROW_HEIGHT);
  const startRow = Math.max(0, firstVisible - BUFFER);
  const endRow = Math.min(sheet.rowCount - 1, firstVisible + visibleCount + BUFFER);

  // Always include the editing row in the rendered range
  const editingRow = editingCell ? (parseCellRef(editingCell)?.row ?? null) : null;

  const handleGridScroll = useCallback((e: React.UIEvent<HTMLDivElement>) => {
    setScrollTop(e.currentTarget.scrollTop);
  }, []);

  // Track grid container height
  useEffect(() => {
    const el = gridRef.current;
    if (!el) return;
    const ro = new ResizeObserver((entries) => {
      for (const entry of entries) setGridHeight(entry.contentRect.height);
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  // Scroll selected cell into view
  useEffect(() => {
    const parsed = parseCellRef(selectedCell);
    if (!parsed || !gridRef.current) return;
    const cellTop = (parsed.row + 1) * ROW_HEIGHT;
    const cellBottom = cellTop + ROW_HEIGHT;
    const viewTop = gridRef.current.scrollTop + ROW_HEIGHT;
    const viewBottom = gridRef.current.scrollTop + gridHeight;
    if (cellTop < viewTop) gridRef.current.scrollTop = cellTop - ROW_HEIGHT;
    else if (cellBottom > viewBottom) gridRef.current.scrollTop = cellBottom - gridHeight;
  }, [selectedCell, gridHeight]);

  // ── Render helpers ──────────────────────────────────────────────────

  const isError = (val: string) => val.startsWith("#");

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

      {/* Formatting toolbar */}
      <div
        style={{
          display: "flex", alignItems: "center", gap: "0.25rem",
          padding: "0.25rem 0.75rem",
          borderBottom: "1px solid var(--color-line)",
          background: "var(--color-panel)", flexShrink: 0, flexWrap: "wrap",
        }}
      >
        {[
          { label: "B", title: "Bold (Ctrl+B)", active: sheet.formats[selectedCell]?.bold, action: toggleBold, style: { fontWeight: 800 } as React.CSSProperties },
          { label: "I", title: "Italic (Ctrl+I)", active: sheet.formats[selectedCell]?.italic, action: toggleItalic, style: { fontStyle: "italic" } as React.CSSProperties },
        ].map(({ label, title, active, action, style }) => (
          <button
            key={label}
            title={title}
            onMouseDown={(e) => e.preventDefault()}
            onClick={action}
            style={{
              width: "1.75rem", height: "1.75rem",
              border: active ? "1px solid var(--color-accent)" : "1px solid var(--color-line)",
              borderRadius: "0.25rem",
              background: active ? "rgba(192,133,82,0.12)" : "transparent",
              color: active ? "var(--color-accent)" : "var(--color-muted)",
              cursor: "pointer", fontSize: "0.8125rem",
              display: "flex", alignItems: "center", justifyContent: "center",
              ...style,
            }}
          >
            {label}
          </button>
        ))}
        <span style={{ width: 1, height: "1rem", background: "var(--color-line)", margin: "0 0.25rem" }} />
        {(["left", "center", "right"] as const).map((a) => (
          <button
            key={a}
            title={`Align ${a}`}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => applyFormat({ align: sheet.formats[selectedCell]?.align === a ? undefined : a })}
            style={{
              width: "1.75rem", height: "1.75rem",
              border: sheet.formats[selectedCell]?.align === a ? "1px solid var(--color-accent)" : "1px solid var(--color-line)",
              borderRadius: "0.25rem",
              background: sheet.formats[selectedCell]?.align === a ? "rgba(192,133,82,0.12)" : "transparent",
              color: sheet.formats[selectedCell]?.align === a ? "var(--color-accent)" : "var(--color-muted)",
              cursor: "pointer", fontSize: "0.6875rem", fontWeight: 600,
              display: "flex", alignItems: "center", justifyContent: "center",
            }}
          >
            {a === "left" ? "⫷" : a === "center" ? "⫶" : "⫸"}
          </button>
        ))}
        <span style={{ width: 1, height: "1rem", background: "var(--color-line)", margin: "0 0.25rem" }} />
        {/* Background color */}
        <div style={{ position: "relative" }}>
          <button
            title="Background color"
            onMouseDown={(e) => e.preventDefault()}
            onClick={(e) => {
              const rect = e.currentTarget.getBoundingClientRect();
              setColorPicker(colorPicker?.type === "bg" ? null : { type: "bg", x: rect.left, y: rect.bottom + 4 });
            }}
            style={{
              width: "1.75rem", height: "1.75rem",
              border: "1px solid var(--color-line)", borderRadius: "0.25rem",
              background: "transparent", cursor: "pointer",
              display: "flex", alignItems: "center", justifyContent: "center",
              fontSize: "0.75rem", color: "var(--color-muted)",
            }}
          >
            <span style={{ display: "block", width: 12, height: 12, borderRadius: 2, background: sheet.formats[selectedCell]?.bg || "var(--color-line)" }} />
          </button>
        </div>
        {/* Text color */}
        <div style={{ position: "relative" }}>
          <button
            title="Text color"
            onMouseDown={(e) => e.preventDefault()}
            onClick={(e) => {
              const rect = e.currentTarget.getBoundingClientRect();
              setColorPicker(colorPicker?.type === "color" ? null : { type: "color", x: rect.left, y: rect.bottom + 4 });
            }}
            style={{
              width: "1.75rem", height: "1.75rem",
              border: "1px solid var(--color-line)", borderRadius: "0.25rem",
              background: "transparent", cursor: "pointer",
              display: "flex", alignItems: "center", justifyContent: "center",
              fontSize: "0.8125rem", fontWeight: 700,
              color: sheet.formats[selectedCell]?.color || "var(--color-muted)",
            }}
          >
            A
          </button>
        </div>
        <span style={{ width: 1, height: "1rem", background: "var(--color-line)", margin: "0 0.25rem" }} />
        <button
          title="Keyboard shortcuts (?)"
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => setShowHelp((p) => !p)}
          style={{
            width: "1.75rem", height: "1.75rem",
            border: "1px solid var(--color-line)", borderRadius: "0.25rem",
            background: "transparent", color: "var(--color-muted)", cursor: "pointer",
            fontSize: "0.8125rem", fontWeight: 600,
            display: "flex", alignItems: "center", justifyContent: "center",
          }}
        >
          ?
        </button>
      </div>

      {/* Color picker popup */}
      {colorPicker && (
        <div
          onMouseDown={(e) => e.stopPropagation()}
          style={{
            position: "fixed", left: colorPicker.x, top: colorPicker.y, zIndex: 100,
            background: "var(--color-paper)", border: "1px solid var(--color-line)",
            borderRadius: "var(--radius-btn)", boxShadow: "0 4px 16px rgba(0,0,0,0.12)",
            padding: "0.5rem", display: "grid", gridTemplateColumns: "repeat(6, 1.5rem)", gap: "0.25rem",
          }}
        >
          {[
            undefined, "#ef4444", "#f97316", "#eab308", "#22c55e", "#3b82f6",
            "#8b5cf6", "#ec4899", "#14b8a6", "#6b7280", "#1e293b", "#fefce8",
            "#fee2e2", "#dbeafe", "#dcfce7", "#f3e8ff", "#fef3c7",
          ].map((c, i) => (
            <button
              key={i}
              onClick={() => {
                applyFormat({ [colorPicker.type]: c });
                setColorPicker(null);
              }}
              style={{
                width: "1.5rem", height: "1.5rem",
                border: c ? "1px solid var(--color-line)" : "2px dashed var(--color-muted)",
                borderRadius: "0.25rem",
                background: c ?? "transparent",
                cursor: "pointer",
              }}
              title={c ?? "Clear"}
            />
          ))}
        </div>
      )}

      {/* Grid — virtualized rows */}
      <div
        ref={gridRef}
        tabIndex={0}
        onKeyDown={handleGridKeyDown}
        onScroll={handleGridScroll}
        onClick={() => setContextMenu(null)}
        onContextMenu={(e) => e.preventDefault()}
        style={{ flex: 1, overflow: "auto", outline: "none", position: "relative" }}
      >
        {(() => {
          const totalW = HEADER_WIDTH + Array.from({ length: sheet.colCount }, (_, c) => getColWidth(c)).reduce((a, b) => a + b, 0);
          const totalH = (sheet.rowCount + 1) * ROW_HEIGHT;
          const findSet = new Set(findMatches);
          const currentMatch = findMatches[findIndex];

          return (
            <div style={{ width: totalW, height: totalH, position: "relative", minWidth: "100%" }}>
              {/* Column headers — sticky */}
              <div style={{ position: "sticky", top: 0, zIndex: 10, display: "flex", height: ROW_HEIGHT }}>
                <div style={{
                  position: "sticky", left: 0, zIndex: 3, width: HEADER_WIDTH, flexShrink: 0,
                  background: "var(--color-panel)",
                  borderBottom: "2px solid var(--color-line)",
                  borderRight: "1px solid var(--color-line)",
                }} />
                {Array.from({ length: sheet.colCount }, (_, c) => {
                  const colSel = selectionRange
                    ? c >= selectionRange.minC && c <= selectionRange.maxC
                    : parseCellRef(selectedCell)?.col === c;
                  const w = getColWidth(c);
                  return (
                    <div
                      key={c}
                      style={{
                        width: w, flexShrink: 0, position: "relative",
                        display: "flex", alignItems: "center", justifyContent: "center",
                        background: colSel ? "var(--color-line)" : "var(--color-panel)",
                        borderBottom: "2px solid var(--color-line)",
                        borderRight: "1px solid var(--color-line)",
                        fontSize: "0.6875rem", fontWeight: 600,
                        color: colSel ? "var(--color-ink)" : "var(--color-muted)",
                        userSelect: "none",
                      }}
                    >
                      {colLabel(c)}
                      <div
                        onMouseDown={(e) => handleColResizeStart(c, e)}
                        onDoubleClick={() => handleColResizeDoubleClick(c)}
                        style={{ position: "absolute", right: -2, top: 0, bottom: 0, width: 5, cursor: "col-resize", zIndex: 4 }}
                      />
                    </div>
                  );
                })}
              </div>

              {/* Virtualized data rows */}
              {Array.from({ length: endRow - startRow + 1 }, (_, i) => {
                const r = startRow + i;
                if (editingRow !== null && r !== editingRow && r === endRow && editingRow > endRow) return null;
                const rowSel = selectionRange
                  ? r >= selectionRange.minR && r <= selectionRange.maxR
                  : parseCellRef(selectedCell)?.row === r;
                return (
                  <div key={r} style={{ position: "absolute", top: (r + 1) * ROW_HEIGHT, left: 0, height: ROW_HEIGHT, display: "flex", width: totalW }}>
                    <div style={{
                      position: "sticky", left: 0, zIndex: 1, width: HEADER_WIDTH, flexShrink: 0,
                      display: "flex", alignItems: "center", justifyContent: "center",
                      background: rowSel ? "var(--color-line)" : "var(--color-panel)",
                      borderBottom: "1px solid var(--color-line)",
                      borderRight: "1px solid var(--color-line)",
                      fontSize: "0.6875rem", fontWeight: 500,
                      color: rowSel ? "var(--color-ink)" : "var(--color-muted)",
                      userSelect: "none",
                    }}>
                      {r + 1}
                    </div>
                    {Array.from({ length: sheet.colCount }, (_, c) => {
                      const id = cellId(c, r);
                      const isSelected = selectedCell === id;
                      const isEditing = editingCell === id;
                      const isPointed = pointHighlight.has(id);
                      const inRange = selectionRange?.cells.has(id) ?? false;
                      const display = displayValues[id] ?? "";
                      const fmt = sheet.formats[id];
                      const isNum = !isNaN(parseFloat(display)) && display !== "" && !isError(display);
                      const isFindMatch = findSet.has(id);
                      const isCurrentMatch = id === currentMatch;
                      const zebraColor = r % 2 === 1 ? "var(--color-zebra)" : undefined;
                      const w = getColWidth(c);

                      return (
                        <div
                          key={c}
                          onMouseDown={(e) => handleCellMouseDown(id, e)}
                          onDoubleClick={() => { if (!editingCell) startEditing(id); }}
                          onContextMenu={(e) => {
                            e.preventDefault(); e.stopPropagation();
                            select(id);
                            setContextMenu({ x: e.clientX, y: e.clientY, cellId: id });
                          }}
                          style={{
                            width: w, flexShrink: 0, position: "relative",
                            borderBottom: "1px solid var(--color-line)",
                            borderRight: "1px solid var(--color-line)",
                            outline: isSelected ? "2px solid var(--color-accent)" : isPointed ? "2px solid #5b8cd6" : "none",
                            outlineOffset: "-1px",
                            background: isCurrentMatch ? "rgba(255,180,0,0.35)"
                              : isFindMatch ? "rgba(255,220,80,0.2)"
                              : isPointed && !isSelected ? "rgba(91,140,214,0.12)"
                              : inRange && !isSelected ? "rgba(192,133,82,0.10)"
                              : fmt?.bg ?? zebraColor,
                            zIndex: isSelected || isPointed ? 1 : 0,
                            cursor: "cell",
                          }}
                        >
                          {isEditing ? (
                            <input
                              ref={editInputRef}
                              value={editValue}
                              onChange={(e) => { if (pointMode) setPointMode(null); setEditValue(e.target.value); }}
                              onKeyDown={handleInputKeyDown}
                              style={{
                                position: "absolute", inset: 0, border: "none", outline: "none",
                                background: "var(--color-paper)", padding: "0 4px", fontSize: "0.8125rem",
                                color: "var(--color-ink)", width: "100%", height: "100%",
                              }}
                            />
                          ) : (
                            <div style={{
                              padding: "0 4px", fontSize: "0.8125rem", lineHeight: `${ROW_HEIGHT}px`,
                              overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
                              color: isError(display) ? "#d94040" : fmt?.color ?? "var(--color-ink)",
                              textAlign: fmt?.align ?? (isNum ? "right" : "left"),
                              fontWeight: fmt?.bold ? 700 : isError(display) ? 600 : 400,
                              fontStyle: fmt?.italic ? "italic" : undefined,
                            }}>
                              {formatDisplay(display)}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                );
              })}

              {/* Render editing row if off-screen */}
              {editingRow !== null && (editingRow < startRow || editingRow > endRow) && (() => {
                const r = editingRow;
                return (
                  <div key={`edit-${r}`} style={{ position: "absolute", top: (r + 1) * ROW_HEIGHT, left: 0, height: ROW_HEIGHT, display: "flex", width: totalW }}>
                    <div style={{ position: "sticky", left: 0, zIndex: 1, width: HEADER_WIDTH, flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center", background: "var(--color-panel)", borderBottom: "1px solid var(--color-line)", borderRight: "1px solid var(--color-line)", fontSize: "0.6875rem", color: "var(--color-muted)", userSelect: "none" }}>
                      {r + 1}
                    </div>
                    {Array.from({ length: sheet.colCount }, (_, c) => {
                      const id = cellId(c, r);
                      const isEditing = editingCell === id;
                      return (
                        <div key={c} style={{ width: getColWidth(c), flexShrink: 0, position: "relative", borderBottom: "1px solid var(--color-line)", borderRight: "1px solid var(--color-line)" }}>
                          {isEditing && (
                            <input ref={editInputRef} value={editValue}
                              onChange={(e) => { if (pointMode) setPointMode(null); setEditValue(e.target.value); }}
                              onKeyDown={handleInputKeyDown}
                              style={{ position: "absolute", inset: 0, border: "none", outline: "none", background: "var(--color-paper)", padding: "0 4px", fontSize: "0.8125rem", color: "var(--color-ink)", width: "100%", height: "100%" }}
                            />
                          )}
                        </div>
                      );
                    })}
                  </div>
                );
              })()}
            </div>
          );
        })()}
      </div>

      {/* Find bar */}
      {findBar.open && (
        <div
          style={{
            display: "flex", alignItems: "center", gap: "0.375rem",
            padding: "0.375rem 0.75rem",
            borderTop: "1px solid var(--color-line)",
            background: "var(--color-panel)", flexShrink: 0, flexWrap: "wrap",
          }}
        >
          <input
            ref={findInputRef}
            value={findBar.query}
            onChange={(e) => setFindBar((p) => ({ ...p, query: e.target.value }))}
            onKeyDown={(e) => {
              if (e.key === "Enter") { e.preventDefault(); e.shiftKey ? findPrev() : findNext(); }
              if (e.key === "Escape") closeFindBar();
            }}
            placeholder="Find..."
            style={{
              border: "1px solid var(--color-line)", borderRadius: "var(--radius-btn)",
              padding: "0.25rem 0.5rem", fontSize: "0.8125rem", background: "var(--color-paper)",
              color: "var(--color-ink)", outline: "none", width: "10rem",
            }}
          />
          {findBar.showReplace && (
            <input
              value={findBar.replace}
              onChange={(e) => setFindBar((p) => ({ ...p, replace: e.target.value }))}
              onKeyDown={(e) => { if (e.key === "Escape") closeFindBar(); }}
              placeholder="Replace..."
              style={{
                border: "1px solid var(--color-line)", borderRadius: "var(--radius-btn)",
                padding: "0.25rem 0.5rem", fontSize: "0.8125rem", background: "var(--color-paper)",
                color: "var(--color-ink)", outline: "none", width: "10rem",
              }}
            />
          )}
          <span style={{ fontSize: "0.6875rem", color: "var(--color-muted)" }}>
            {findMatches.length > 0 ? `${findIndex + 1}/${findMatches.length}` : findBar.query ? "0 results" : ""}
          </span>
          {[
            { label: "↑", action: findPrev },
            { label: "↓", action: findNext },
            ...(findBar.showReplace ? [
              { label: "Replace", action: replaceOne },
              { label: "All", action: replaceAll },
            ] : []),
            { label: "×", action: closeFindBar },
          ].map(({ label, action }) => (
            <button
              key={label}
              onClick={action}
              style={{
                border: "1px solid var(--color-line)", borderRadius: "var(--radius-btn)",
                background: "transparent", color: "var(--color-muted)",
                padding: "0.125rem 0.5rem", fontSize: "0.75rem", cursor: "pointer", fontWeight: 600,
              }}
            >
              {label}
            </button>
          ))}
        </div>
      )}

      {/* Sheet tabs */}
      <div
        style={{
          display: "flex", alignItems: "center",
          borderTop: "1px solid var(--color-line)",
          background: "var(--color-panel)", flexShrink: 0,
          overflowX: "auto", minHeight: "1.75rem",
        }}
      >
        {workbook.sheets.map((s, i) => (
          <button
            key={i}
            onClick={() => switchSheet(i)}
            onDoubleClick={() => renameSheet(i)}
            onContextMenu={(e) => {
              e.preventDefault();
              if (workbook.sheets.length > 1 && confirm(`Delete "${s.name}"?`)) deleteSheet(i);
            }}
            style={{
              padding: "0.25rem 0.75rem",
              fontSize: "0.6875rem", fontWeight: i === workbook.activeSheet ? 700 : 400,
              border: "none", borderRight: "1px solid var(--color-line)",
              background: i === workbook.activeSheet ? "var(--color-paper)" : "transparent",
              color: i === workbook.activeSheet ? "var(--color-ink)" : "var(--color-muted)",
              cursor: "pointer", whiteSpace: "nowrap",
            }}
          >
            {s.name}
          </button>
        ))}
        <button
          onClick={addSheet}
          style={{
            padding: "0.25rem 0.5rem", fontSize: "0.8125rem", fontWeight: 600,
            border: "none", background: "transparent",
            color: "var(--color-muted)", cursor: "pointer",
          }}
        >
          +
        </button>
      </div>

      {/* Status bar */}
      <div
        style={{
          display: "flex", alignItems: "center", gap: "1.25rem",
          padding: "0.25rem 0.75rem",
          borderTop: "1px solid var(--color-line)",
          background: "var(--color-panel)", flexShrink: 0,
          minHeight: "1.5rem", fontSize: "0.6875rem", color: "var(--color-muted)",
        }}
      >
        {selectionStats.numCount > 1 && (
          <>
            <span>Sum: <b style={{ color: "var(--color-ink)" }}>{selectionStats.sum.toLocaleString("en-US", { maximumFractionDigits: 4 })}</b></span>
            <span>Avg: <b style={{ color: "var(--color-ink)" }}>{selectionStats.avg.toLocaleString("en-US", { maximumFractionDigits: 4 })}</b></span>
          </>
        )}
        {selectionStats.count > 0 && (
          <span>Count: <b style={{ color: "var(--color-ink)" }}>{selectionStats.count}</b></span>
        )}
        <div style={{ flex: 1 }} />
        <span>{sheet.rowCount} rows × {sheet.colCount} cols</span>
      </div>

      {/* Context menu */}
      {contextMenu && (() => {
        const p = parseCellRef(contextMenu.cellId);
        if (!p) return null;
        const items = [
          { label: "Insert row above", action: () => { insertRow(p.row, "above"); setContextMenu(null); } },
          { label: "Insert row below", action: () => { insertRow(p.row, "below"); setContextMenu(null); } },
          { label: "Delete row", action: () => { deleteRow(p.row); setContextMenu(null); } },
          { label: "─", action: () => {} },
          { label: "Insert column left", action: () => { insertCol(p.col, "left"); setContextMenu(null); } },
          { label: "Insert column right", action: () => { insertCol(p.col, "right"); setContextMenu(null); } },
          { label: "Delete column", action: () => { deleteCol(p.col); setContextMenu(null); } },
          { label: "─", action: () => {} },
          { label: "Sort A → Z", action: () => sortByColumn(p.col, true) },
          { label: "Sort Z → A", action: () => sortByColumn(p.col, false) },
        ];
        return (
          <div
            onMouseDown={(e) => e.stopPropagation()}
            style={{
              position: "fixed", left: contextMenu.x, top: contextMenu.y, zIndex: 100,
              background: "var(--color-paper)", border: "1px solid var(--color-line)",
              borderRadius: "var(--radius-btn)", boxShadow: "0 4px 16px rgba(0,0,0,0.12)",
              padding: "0.25rem 0", minWidth: "10rem",
            }}
          >
            {items.map((item, i) =>
              item.label === "─" ? (
                <div key={i} style={{ borderTop: "1px solid var(--color-line)", margin: "0.25rem 0" }} />
              ) : (
                <button
                  key={item.label}
                  onClick={item.action}
                  style={{
                    display: "block", width: "100%", textAlign: "left",
                    background: "transparent", border: "none", cursor: "pointer",
                    padding: "0.375rem 0.75rem", fontSize: "0.8125rem",
                    color: "var(--color-ink)",
                  }}
                  onMouseEnter={(e) => { e.currentTarget.style.background = "var(--color-panel)"; }}
                  onMouseLeave={(e) => { e.currentTarget.style.background = "transparent"; }}
                >
                  {item.label}
                </button>
              )
            )}
          </div>
        );
      })()}

      {/* Help dialog */}
      {showHelp && (
        <div
          onClick={() => setShowHelp(false)}
          style={{
            position: "fixed", inset: 0, zIndex: 200,
            background: "rgba(0,0,0,0.4)",
            display: "flex", alignItems: "center", justifyContent: "center",
          }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{
              background: "var(--color-paper)", borderRadius: "var(--radius-card)",
              padding: "1.5rem", maxWidth: "28rem", width: "90vw",
              maxHeight: "80vh", overflow: "auto",
              boxShadow: "0 8px 32px rgba(0,0,0,0.2)",
            }}
          >
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1rem" }}>
              <h2 style={{ margin: 0, fontSize: "1.125rem" }}>Keyboard Shortcuts</h2>
              <button onClick={() => setShowHelp(false)} style={{ border: "none", background: "transparent", fontSize: "1.25rem", cursor: "pointer", color: "var(--color-muted)" }}>×</button>
            </div>
            {[
              ["Navigation", [
                ["Arrow keys", "Move selection"],
                ["Shift + Arrow", "Extend selection"],
                ["Tab / Shift+Tab", "Move right / left"],
                ["Enter / Shift+Enter", "Move down / up"],
                ["Ctrl+A", "Select all"],
              ]],
              ["Editing", [
                ["Type any character", "Start editing cell"],
                ["F2 / Enter", "Edit selected cell"],
                ["Escape", "Cancel edit"],
                ["Delete / Backspace", "Clear cell(s)"],
              ]],
              ["Formulas", [
                ["= then Arrow keys", "Insert cell reference"],
                ["Shift + Arrow (in formula)", "Extend to range (A1:B3)"],
                ["Click cell (in formula)", "Insert cell reference"],
              ]],
              ["Formatting", [
                ["Ctrl+B", "Bold"],
                ["Ctrl+I", "Italic"],
              ]],
              ["Tools", [
                ["Ctrl+C / X / V", "Copy / Cut / Paste"],
                ["Ctrl+Z", "Undo"],
                ["Ctrl+Shift+Z / Ctrl+Y", "Redo"],
                ["Ctrl+F", "Find"],
                ["Ctrl+H", "Find & Replace"],
                ["Right-click", "Context menu (insert/delete/sort)"],
                ["Ctrl+/", "This help"],
              ]],
            ].map(([title, shortcuts]) => (
              <div key={title as string} style={{ marginBottom: "0.75rem" }}>
                <div style={{ fontSize: "0.75rem", fontWeight: 700, color: "var(--color-accent)", marginBottom: "0.25rem", textTransform: "uppercase", letterSpacing: "0.05em" }}>
                  {title as string}
                </div>
                {(shortcuts as string[][]).map(([key, desc]) => (
                  <div key={key} style={{ display: "flex", justifyContent: "space-between", padding: "0.125rem 0", fontSize: "0.8125rem" }}>
                    <code style={{ background: "var(--color-panel)", padding: "0.05rem 0.375rem", borderRadius: "0.25rem", fontSize: "0.75rem" }}>{key}</code>
                    <span style={{ color: "var(--color-muted)" }}>{desc}</span>
                  </div>
                ))}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
