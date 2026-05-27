import { useState, useEffect, useCallback, useRef, useMemo, useLayoutEffect } from "react";
import type { CellFormat, ConditionalRule, Sheet, WorkbookData, PointMode } from "../types";
import { DEFAULT_COLS, DEFAULT_ROWS, MAX_UNDO, COL_WIDTH, ROW_HEIGHT, HEADER_WIDTH } from "../constants";
import { cellId, parseCellRef, clamp, expandRange } from "../lib/cell-refs";
import { computeDisplay } from "../lib/formula-engine";
import { parseCsvLine } from "../lib/csv";
import { isRefPosition, buildRefString, splice } from "../lib/point-mode";
import { makeSheet, loadWorkbook, saveWorkbook } from "../lib/persistence";

export function useSpreadsheet() {
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
  const [contextMenu, setContextMenu] = useState<{ x: number; y: number; cellId: string; type?: "cell" | "col" | "row" } | null>(null);
  const [findBar, setFindBar] = useState<{ open: boolean; query: string; replace: string; showReplace: boolean }>({ open: false, query: "", replace: "", showReplace: false });
  const [showHelp, setShowHelp] = useState(false);
  const [colorPicker, setColorPicker] = useState<{ type: "bg" | "color"; x: number; y: number } | null>(null);
  const [findMatches, setFindMatches] = useState<string[]>([]);
  const [findIndex, setFindIndex] = useState(0);
  const [frozenRows, setFrozenRows] = useState(0);
  const [autoFillTarget, setAutoFillTarget] = useState<string | null>(null);
  const [condFormatDialog, setCondFormatDialog] = useState(false);

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

  // Lookup function for cross-sheet refs (Sheet2!A1).
  // Sheet names are matched case-insensitively (Excel convention).
  const sheetLookup = useCallback((sheetName: string, ref: string): string => {
    const found = workbook.sheets.find((s) => s.name.toLowerCase() === sheetName.toLowerCase());
    return found?.cells[ref] ?? "";
  }, [workbook.sheets]);

  const displayValues = useMemo(() => {
    const result: Record<string, string> = {};
    for (const id of Object.keys(sheet.cells)) {
      result[id] = computeDisplay(id, sheet.cells, sheetLookup);
    }
    return result;
  }, [sheet.cells, sheetLookup]);

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
  }, [pushUndo, setSheet]);

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

  const handleCellMouseDown = useCallback((id: string, e: React.MouseEvent) => {
    // Formula edit -> click another cell = insert reference
    if (editingCell && editValue.startsWith("=") && id !== editingCell) {
      e.preventDefault(); // keep input focused
      insertRefByClick(id, e.shiftKey);
      return;
    }

    // Non-formula edit -> click another cell = commit + select
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
  }, [selectedCell, selectionRange, sheet, clipboard, move, extendSelection, startEditing, pushUndo, undo, redo, setSheet]);

  // ── Keyboard: inline cell input ─────────────────────────────────────

  const handleInputKeyDown = useCallback((e: React.KeyboardEvent) => {
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
    }

    // Any non-modifier key exits point mode
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
  }, [editingCell, selectedCell, formulaBarValue, sheet.cells, commitEdit, cancelEdit, pushUndo, setSheet]);

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
        const display = raw.startsWith("=") ? computeDisplay(id, prev.cells, sheetLookup) : raw;
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
  }, [pushUndo, setSheet, sheetLookup]);

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

  // ── Freeze panes ────────────────────────────────────────────────────

  const toggleFreezeFirstRow = useCallback(() => {
    setFrozenRows((n) => (n > 0 ? 0 : 1));
  }, []);

  // ── Conditional formatting ──────────────────────────────────────────

  const addConditionalRule = useCallback((rule: ConditionalRule) => {
    pushUndo();
    setSheet((prev) => ({
      ...prev,
      conditionalRules: [...(prev.conditionalRules ?? []), rule],
    }));
  }, [pushUndo, setSheet]);

  const removeConditionalRule = useCallback((index: number) => {
    pushUndo();
    setSheet((prev) => ({
      ...prev,
      conditionalRules: (prev.conditionalRules ?? []).filter((_, i) => i !== index),
    }));
  }, [pushUndo, setSheet]);

  // Compute effective conditional format for a cell. Last matching rule wins.
  const evalConditional = useCallback((cellRef: string, displayVal: string): { bg?: string; color?: string; bold?: boolean } | null => {
    const rules = sheet.conditionalRules ?? [];
    if (rules.length === 0) return null;
    const cellPos = parseCellRef(cellRef);
    if (!cellPos) return null;

    let result: { bg?: string; color?: string; bold?: boolean } | null = null;
    for (const rule of rules) {
      // Check if cell is in range
      const range = expandRange(rule.range);
      const inRange = range.length > 0 ? range.includes(cellRef) : rule.range === cellRef;
      if (!inRange) continue;

      const num = parseFloat(displayVal);
      const ruleNum = parseFloat(rule.value);
      let matches = false;

      switch (rule.type) {
        case "greater": matches = !isNaN(num) && !isNaN(ruleNum) && num > ruleNum; break;
        case "less": matches = !isNaN(num) && !isNaN(ruleNum) && num < ruleNum; break;
        case "equal": matches = displayVal === rule.value; break;
        case "contains": matches = displayVal.toLowerCase().includes(rule.value.toLowerCase()); break;
        case "between": {
          const [lo, hi] = rule.value.split("..").map((v) => parseFloat(v));
          matches = !isNaN(num) && lo !== undefined && hi !== undefined && !isNaN(lo) && !isNaN(hi) && num >= lo && num <= hi;
          break;
        }
      }

      if (matches) result = rule.format;
    }
    return result;
  }, [sheet.conditionalRules]);

  // ── Auto-fill ───────────────────────────────────────────────────────
  // Detect a series from source cells and extend it to target cells.

  const detectSeries = useCallback((values: string[]): ((index: number) => string) | null => {
    if (values.length === 0) return null;
    const nums = values.map((v) => parseFloat(v));
    const allNumeric = nums.every((n) => !isNaN(n));

    if (allNumeric && values.length >= 2) {
      const step = nums[1]! - nums[0]!;
      const consistent = nums.every((n, i) => i === 0 || Math.abs(n - nums[i - 1]! - step) < 1e-9);
      if (consistent) {
        return (i) => {
          const v = nums[nums.length - 1]! + step * (i + 1);
          return Number.isInteger(v) ? String(v) : String(Math.round(v * 1e10) / 1e10);
        };
      }
    }

    if (allNumeric && values.length === 1) {
      return (i) => String(nums[0]! + i + 1);
    }

    const trailingNumRe = /^(.*?)(\d+)$/;
    const matches = values.map((v) => v.match(trailingNumRe));
    const allTextNum = matches.every((m) => m !== null);
    if (allTextNum && values.length >= 1) {
      const prefix = matches[0]![1]!;
      const samePrefix = matches.every((m) => m![1] === prefix);
      if (samePrefix) {
        const lastNum = parseInt(matches[matches.length - 1]![2]!, 10);
        const step = values.length >= 2 ? parseInt(matches[1]![2]!, 10) - parseInt(matches[0]![2]!, 10) : 1;
        return (i) => `${prefix}${lastNum + step * (i + 1)}`;
      }
    }

    return (i) => values[i % values.length]!;
  }, []);

  const performAutoFill = useCallback((targetId: string) => {
    if (!selectionRange) {
      // Single cell source — extend to target
      const src = parseCellRef(selectedCell);
      const tgt = parseCellRef(targetId);
      if (!src || !tgt) return;
      const value = sheet.cells[selectedCell] ?? "";
      if (!value) return;

      const minC = Math.min(src.col, tgt.col), maxC = Math.max(src.col, tgt.col);
      const minR = Math.min(src.row, tgt.row), maxR = Math.max(src.row, tgt.row);
      const isVertical = src.col === tgt.col;
      const isHorizontal = src.row === tgt.row;
      if (!isVertical && !isHorizontal) return;

      pushUndo();
      setSheet((prev) => {
        const cells = { ...prev.cells };
        const sourceValues = [value];
        const fillFn = detectSeries(sourceValues);
        if (!fillFn) return prev;

        if (isVertical) {
          let idx = 0;
          for (let r = src.row; r !== tgt.row; r += src.row < tgt.row ? 1 : -1) {
            if (r === src.row) continue;
            cells[cellId(src.col, r)] = fillFn(idx++);
          }
          cells[cellId(tgt.col, tgt.row)] = fillFn(idx);
        } else {
          let idx = 0;
          for (let c = src.col; c !== tgt.col; c += src.col < tgt.col ? 1 : -1) {
            if (c === src.col) continue;
            cells[cellId(c, src.row)] = fillFn(idx++);
          }
          cells[cellId(tgt.col, tgt.row)] = fillFn(idx);
        }
        return { ...prev, cells };
      });
      setSelectionEnd(targetId);
      void [minC, maxC, minR, maxR];
      return;
    }

    // Range source — extend the pattern
    const tgt = parseCellRef(targetId);
    if (!tgt) return;
    const srcMinC = selectionRange.minC, srcMaxC = selectionRange.maxC;
    const srcMinR = selectionRange.minR, srcMaxR = selectionRange.maxR;

    const fillDown = tgt.row > srcMaxR && tgt.col >= srcMinC && tgt.col <= srcMaxC;
    const fillUp = tgt.row < srcMinR && tgt.col >= srcMinC && tgt.col <= srcMaxC;
    const fillRight = tgt.col > srcMaxC && tgt.row >= srcMinR && tgt.row <= srcMaxR;
    const fillLeft = tgt.col < srcMinC && tgt.row >= srcMinR && tgt.row <= srcMaxR;

    if (!fillDown && !fillUp && !fillRight && !fillLeft) return;

    pushUndo();
    setSheet((prev) => {
      const cells = { ...prev.cells };
      if (fillDown || fillUp) {
        for (let c = srcMinC; c <= srcMaxC; c++) {
          const sourceValues: string[] = [];
          for (let r = srcMinR; r <= srcMaxR; r++) {
            sourceValues.push(prev.cells[cellId(c, r)] ?? "");
          }
          const fillFn = detectSeries(sourceValues);
          if (!fillFn) continue;
          if (fillDown) {
            for (let r = srcMaxR + 1, i = 0; r <= tgt.row; r++, i++) {
              cells[cellId(c, r)] = fillFn(i);
            }
          } else {
            const reversed = sourceValues.slice().reverse();
            const reverseFill = detectSeries(reversed);
            if (!reverseFill) continue;
            for (let r = srcMinR - 1, i = 0; r >= tgt.row; r--, i++) {
              cells[cellId(c, r)] = reverseFill(i);
            }
          }
        }
      } else {
        for (let r = srcMinR; r <= srcMaxR; r++) {
          const sourceValues: string[] = [];
          for (let c = srcMinC; c <= srcMaxC; c++) {
            sourceValues.push(prev.cells[cellId(c, r)] ?? "");
          }
          const fillFn = detectSeries(sourceValues);
          if (!fillFn) continue;
          if (fillRight) {
            for (let c = srcMaxC + 1, i = 0; c <= tgt.col; c++, i++) {
              cells[cellId(c, r)] = fillFn(i);
            }
          } else {
            const reversed = sourceValues.slice().reverse();
            const reverseFill = detectSeries(reversed);
            if (!reverseFill) continue;
            for (let c = srcMinC - 1, i = 0; c >= tgt.col; c--, i++) {
              cells[cellId(c, r)] = reverseFill(i);
            }
          }
        }
      }
      return { ...prev, cells };
    });
    setSelectionEnd(targetId);
  }, [selectedCell, selectionRange, sheet.cells, pushUndo, detectSeries]);

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
        const val = raw.startsWith("=") ? computeDisplay(cellId(c, r), sheet.cells, sheetLookup) : raw;
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
  }, [sheet.cells, sheetLookup]);

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
  }, [pushUndo, setSheet]);

  // ── Number formatting ────────────────────────────────────────────────

  const formatDisplay = useCallback((val: string, numFmt?: CellFormat["numFmt"]): string => {
    if (!val || val.startsWith("#")) return val;
    const num = parseFloat(val);
    if (isNaN(num)) return val;

    if (numFmt === "currency") {
      return num.toLocaleString("en-US", { style: "currency", currency: "USD" });
    }
    if (numFmt === "percent") {
      return (num * 100).toLocaleString("en-US", { maximumFractionDigits: 2 }) + "%";
    }
    if (numFmt === "decimal2") {
      return num.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }

    // Auto formatting (only if val is exactly a number string)
    if (val !== String(num)) return val;
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

  // ── Formula bar change handler ──────────────────────────────────────

  const handleFormulaBarChange = useCallback((value: string) => {
    setFormulaBarValue(value);
    if (editingCell) {
      setEditValue(value);
      setPointMode(null);
    }
  }, [editingCell]);

  const handleFormulaBarFocus = useCallback(() => {
    setFormulaBarFocused(true);
  }, []);

  const handleFormulaBarBlur = useCallback(() => {
    setFormulaBarFocused(false);
  }, []);

  // ── Edit input change handler ──────────────────────────────────────

  const handleEditInputChange = useCallback((value: string) => {
    if (pointMode) setPointMode(null);
    setEditValue(value);
  }, [pointMode]);

  return {
    // State
    workbook,
    sheet,
    selectedCell,
    selectionEnd,
    editingCell,
    editValue,
    formulaBarValue,
    pointMode,
    contextMenu,
    findBar,
    showHelp,
    colorPicker,
    findMatches,
    findIndex,
    clipboard,
    frozenRows,
    autoFillTarget,
    setAutoFillTarget,
    setSelectionEnd,
    toggleFreezeFirstRow,
    performAutoFill,
    condFormatDialog,
    setCondFormatDialog,
    addConditionalRule,
    removeConditionalRule,
    evalConditional,

    // Derived
    displayValues,
    selectionRange,
    pointHighlight,
    selectionStats,

    // Refs
    editInputRef,
    formulaBarRef,
    gridRef,
    findInputRef,

    // Virtualization
    startRow,
    endRow,
    editingRow,
    ROW_HEIGHT,
    COL_WIDTH,
    HEADER_WIDTH,

    // Callbacks
    handleCellMouseDown,
    handleGridKeyDown,
    handleInputKeyDown,
    handleFormulaBarKeyDown,
    handleGridScroll,
    handleColResizeStart,
    handleColResizeDoubleClick,
    handleFormulaBarChange,
    handleFormulaBarFocus,
    handleFormulaBarBlur,
    handleEditInputChange,
    startEditing,
    select,
    getColWidth,
    formatDisplay,

    // Actions
    importCsv,
    exportCsv,
    clearAll,
    toggleBold,
    toggleItalic,
    applyFormat,
    setColorPicker,
    setShowHelp,
    setContextMenu,
    setFindBar,
    findNext,
    findPrev,
    replaceOne,
    replaceAll,
    closeFindBar,
    addSheet,
    switchSheet,
    renameSheet,
    deleteSheet,
    insertRow,
    deleteRow,
    insertCol,
    deleteCol,
    sortByColumn,
  };
}

export type SpreadsheetState = ReturnType<typeof useSpreadsheet>;
