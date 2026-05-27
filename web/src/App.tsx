import { useState, useEffect, useCallback, useRef, useMemo } from "react";

// ── Types ──────────────────────────────────────────────────────────────

interface SheetData {
  cells: Record<string, string>;
  colCount: number;
  rowCount: number;
}

// ── Persistence ────────────────────────────────────────────────────────

const STORAGE_KEY = "spreadsheet_data";
const DEFAULT_COLS = 26;
const DEFAULT_ROWS = 50;

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

  // Replace cell references with their values
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
    // Only allow safe characters: digits, operators, parens, dots, spaces, comparison
    if (/^[\d\s+\-*/().%<>=!&|]+$/.test(resolved)) {
      // Convert comparison operators
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

function computeDisplay(
  id: string,
  cells: Record<string, string>,
): string {
  const raw = cells[id];
  if (!raw) return "";
  if (!raw.startsWith("=")) return raw;
  return evaluateFormula(raw, cells, new Set(), id);
}

// ── App ────────────────────────────────────────────────────────────────

export function App() {
  const [sheet, setSheet] = useState<SheetData>(loadSheet);
  const [selectedCell, setSelectedCell] = useState<string | null>(null);
  const [editingCell, setEditingCell] = useState<string | null>(null);
  const [editValue, setEditValue] = useState("");
  const [formulaBarValue, setFormulaBarValue] = useState("");
  const [formulaBarFocused, setFormulaBarFocused] = useState(false);
  const editInputRef = useRef<HTMLInputElement>(null);
  const formulaBarRef = useRef<HTMLInputElement>(null);
  const gridRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    saveSheet(sheet);
  }, [sheet]);

  const displayValues = useMemo(() => {
    const result: Record<string, string> = {};
    for (const id of Object.keys(sheet.cells)) {
      result[id] = computeDisplay(id, sheet.cells);
    }
    return result;
  }, [sheet.cells]);

  const updateCell = useCallback((id: string, value: string) => {
    setSheet((prev) => {
      const next = { ...prev, cells: { ...prev.cells } };
      if (value === "") {
        delete next.cells[id];
      } else {
        next.cells[id] = value;
      }
      return next;
    });
  }, []);

  const commitEdit = useCallback(() => {
    if (editingCell) {
      updateCell(editingCell, editValue);
      setEditingCell(null);
    }
  }, [editingCell, editValue, updateCell]);

  const commitFormulaBar = useCallback(() => {
    if (selectedCell) {
      updateCell(selectedCell, formulaBarValue);
    }
  }, [selectedCell, formulaBarValue, updateCell]);

  const startEditing = useCallback((id: string) => {
    setEditingCell(id);
    setEditValue(sheet.cells[id] ?? "");
  }, [sheet.cells]);

  const selectCell = useCallback((id: string) => {
    if (editingCell && editingCell !== id) {
      commitEdit();
    }
    setSelectedCell(id);
    setFormulaBarValue(sheet.cells[id] ?? "");
  }, [editingCell, commitEdit, sheet.cells]);

  // Sync formula bar when selecting cells
  useEffect(() => {
    if (selectedCell && !formulaBarFocused) {
      setFormulaBarValue(sheet.cells[selectedCell] ?? "");
    }
  }, [selectedCell, sheet.cells, formulaBarFocused]);

  const moveSelection = useCallback((dCol: number, dRow: number) => {
    if (!selectedCell) return;
    const parsed = parseCellRef(selectedCell);
    if (!parsed) return;
    const newCol = Math.max(0, Math.min(parsed.col + dCol, sheet.colCount - 1));
    const newRow = Math.max(0, Math.min(parsed.row + dRow, sheet.rowCount - 1));
    const newId = cellId(newCol, newRow);
    selectCell(newId);
  }, [selectedCell, sheet.colCount, sheet.rowCount, selectCell]);

  const handleCellKeyDown = useCallback((e: React.KeyboardEvent) => {
    if (editingCell) {
      if (e.key === "Enter") {
        e.preventDefault();
        commitEdit();
        moveSelection(0, 1);
      } else if (e.key === "Tab") {
        e.preventDefault();
        commitEdit();
        moveSelection(e.shiftKey ? -1 : 1, 0);
      } else if (e.key === "Escape") {
        setEditingCell(null);
      }
      return;
    }

    if (e.key === "Enter") {
      e.preventDefault();
      if (selectedCell) startEditing(selectedCell);
    } else if (e.key === "Tab") {
      e.preventDefault();
      moveSelection(e.shiftKey ? -1 : 1, 0);
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      moveSelection(0, -1);
    } else if (e.key === "ArrowDown") {
      e.preventDefault();
      moveSelection(0, 1);
    } else if (e.key === "ArrowLeft") {
      e.preventDefault();
      moveSelection(-1, 0);
    } else if (e.key === "ArrowRight") {
      e.preventDefault();
      moveSelection(1, 0);
    } else if (e.key === "Delete" || e.key === "Backspace") {
      if (selectedCell) {
        updateCell(selectedCell, "");
        setFormulaBarValue("");
      }
    } else if (e.key.length === 1 && !e.metaKey && !e.ctrlKey) {
      if (selectedCell) {
        setEditingCell(selectedCell);
        setEditValue(e.key);
      }
    }
  }, [editingCell, selectedCell, commitEdit, moveSelection, startEditing, updateCell]);

  const handleFormulaBarKeyDown = useCallback((e: React.KeyboardEvent) => {
    if (e.key === "Enter") {
      e.preventDefault();
      commitFormulaBar();
      formulaBarRef.current?.blur();
      gridRef.current?.focus();
    } else if (e.key === "Escape") {
      if (selectedCell) {
        setFormulaBarValue(sheet.cells[selectedCell] ?? "");
      }
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
    setSheet({ cells: {}, colCount: DEFAULT_COLS, rowCount: DEFAULT_ROWS });
    setSelectedCell(null);
    setEditingCell(null);
    setEditValue("");
    setFormulaBarValue("");
  }, []);

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
          onChange={(e) => setFormulaBarValue(e.target.value)}
          onFocus={() => setFormulaBarFocused(true)}
          onBlur={() => {
            setFormulaBarFocused(false);
            commitFormulaBar();
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
        onKeyDown={handleCellKeyDown}
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
                const display = displayValues[id] ?? "";

                return (
                  <div
                    key={id}
                    onClick={() => selectCell(id)}
                    onDoubleClick={() => startEditing(id)}
                    style={{
                      position: "relative",
                      borderBottom: "1px solid var(--color-line)",
                      borderRight: "1px solid var(--color-line)",
                      outline: isSelected ? `2px solid var(--color-accent)` : "none",
                      outlineOffset: "-1px",
                      zIndex: isSelected ? 1 : 0,
                      cursor: "cell",
                    }}
                  >
                    {isEditing ? (
                      <input
                        ref={editInputRef}
                        value={editValue}
                        onChange={(e) => setEditValue(e.target.value)}
                        onBlur={commitEdit}
                        onKeyDown={(e) => {
                          if (e.key === "Enter") {
                            e.preventDefault();
                            commitEdit();
                            moveSelection(0, 1);
                          } else if (e.key === "Tab") {
                            e.preventDefault();
                            commitEdit();
                            moveSelection(e.shiftKey ? -1 : 1, 0);
                          } else if (e.key === "Escape") {
                            setEditingCell(null);
                          }
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
                          color: isError(display)
                            ? "#d94040"
                            : !isNaN(parseFloat(display)) && display !== ""
                              ? "var(--color-ink)"
                              : "var(--color-ink)",
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
