import { useCallback } from "react";
import { colLabel, cellId, parseCellRef } from "../lib/cell-refs";
import type { SpreadsheetState } from "../hooks/useSpreadsheet";

const isError = (val: string) => val.startsWith("#");

export function Grid({ s }: { s: SpreadsheetState }) {
  const totalW = s.HEADER_WIDTH + Array.from({ length: s.sheet.colCount }, (_, c) => s.getColWidth(c)).reduce((a, b) => a + b, 0);
  const totalH = (s.sheet.rowCount + 1) * s.ROW_HEIGHT;
  const findSet = new Set(s.findMatches);
  const currentMatch = s.findMatches[s.findIndex];

  // Auto-fill drag handle
  const startAutoFill = useCallback((e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const gridEl = s.gridRef.current;
    if (!gridEl) return;
    const rect = gridEl.getBoundingClientRect();

    const onMove = (ev: MouseEvent) => {
      const x = ev.clientX - rect.left + gridEl.scrollLeft;
      const y = ev.clientY - rect.top + gridEl.scrollTop;
      if (y < s.ROW_HEIGHT || x < s.HEADER_WIDTH) return;
      const r = Math.floor((y - s.ROW_HEIGHT) / s.ROW_HEIGHT);
      let cx = s.HEADER_WIDTH;
      let c = 0;
      for (; c < s.sheet.colCount; c++) {
        cx += s.getColWidth(c);
        if (x < cx) break;
      }
      const targetId = cellId(Math.min(c, s.sheet.colCount - 1), Math.min(r, s.sheet.rowCount - 1));
      s.setAutoFillTarget(targetId);
    };
    const onUp = () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
      // Read target from closure-fresh state via DOM attribute fallback
      const tgt = (gridEl as HTMLDivElement & { _autoFillTarget?: string })._autoFillTarget;
      if (tgt) s.performAutoFill(tgt);
      s.setAutoFillTarget(null);
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
  }, [s]);

  // Determine auto-fill anchor cell (bottom-right of current selection / cell)
  const anchorCell = s.selectionRange
    ? cellId(s.selectionRange.maxC, s.selectionRange.maxR)
    : s.selectedCell;
  const anchor = parseCellRef(anchorCell);

  // Auto-fill preview range (cells from selection edge to target)
  const fillPreview = new Set<string>();
  if (s.autoFillTarget) {
    const tgt = parseCellRef(s.autoFillTarget);
    const src = s.selectionRange ?? (anchor ? { minC: anchor.col, maxC: anchor.col, minR: anchor.row, maxR: anchor.row } : null);
    if (tgt && src) {
      if (tgt.row > src.maxR && tgt.col >= src.minC && tgt.col <= src.maxC) {
        for (let r = src.maxR + 1; r <= tgt.row; r++)
          for (let c = src.minC; c <= src.maxC; c++)
            fillPreview.add(cellId(c, r));
      } else if (tgt.row < src.minR && tgt.col >= src.minC && tgt.col <= src.maxC) {
        for (let r = tgt.row; r < src.minR; r++)
          for (let c = src.minC; c <= src.maxC; c++)
            fillPreview.add(cellId(c, r));
      } else if (tgt.col > src.maxC && tgt.row >= src.minR && tgt.row <= src.maxR) {
        for (let c = src.maxC + 1; c <= tgt.col; c++)
          for (let r = src.minR; r <= src.maxR; r++)
            fillPreview.add(cellId(c, r));
      } else if (tgt.col < src.minC && tgt.row >= src.minR && tgt.row <= src.maxR) {
        for (let c = tgt.col; c < src.minC; c++)
          for (let r = src.minR; r <= src.maxR; r++)
            fillPreview.add(cellId(c, r));
      }
    }
  }

  // Stash target on grid element for the mouseup handler closure
  if (s.gridRef.current) {
    (s.gridRef.current as HTMLDivElement & { _autoFillTarget?: string | null })._autoFillTarget = s.autoFillTarget;
  }

  return (
    <div
      ref={s.gridRef}
      tabIndex={0}
      onKeyDown={s.handleGridKeyDown}
      onScroll={s.handleGridScroll}
      onClick={() => s.setContextMenu(null)}
      onContextMenu={(e) => e.preventDefault()}
      style={{ flex: 1, overflow: "auto", outline: "none", position: "relative" }}
    >
      <div style={{ width: totalW, height: totalH, position: "relative", minWidth: "100%" }}>
        {/* Column headers -- sticky */}
        <div style={{ position: "sticky", top: 0, zIndex: 10, display: "flex", height: s.ROW_HEIGHT }}>
          <div style={{
            position: "sticky", left: 0, zIndex: 3, width: s.HEADER_WIDTH, flexShrink: 0,
            background: "var(--color-panel)",
            borderBottom: "2px solid var(--color-line)",
            borderRight: "1px solid var(--color-line)",
          }} />
          {Array.from({ length: s.sheet.colCount }, (_, c) => {
            const colSel = s.selectionRange
              ? c >= s.selectionRange.minC && c <= s.selectionRange.maxC
              : parseCellRef(s.selectedCell)?.col === c;
            const w = s.getColWidth(c);
            return (
              <div
                key={c}
                onMouseDown={(e) => {
                  if (e.button !== 0) return;
                  s.select(cellId(c, 0));
                  s.setSelectionEnd(cellId(c, s.sheet.rowCount - 1));
                  s.gridRef.current?.focus();
                }}
                onContextMenu={(e) => {
                  e.preventDefault(); e.stopPropagation();
                  s.select(cellId(c, 0));
                  s.setSelectionEnd(cellId(c, s.sheet.rowCount - 1));
                  s.setContextMenu({ x: e.clientX, y: e.clientY, cellId: cellId(c, 0), type: "col" });
                }}
                style={{
                  width: w, flexShrink: 0, position: "relative",
                  display: "flex", alignItems: "center", justifyContent: "center",
                  background: colSel ? "var(--color-line)" : "var(--color-panel)",
                  borderBottom: "2px solid var(--color-line)",
                  borderRight: "1px solid var(--color-line)",
                  fontSize: "0.6875rem", fontWeight: 600,
                  color: colSel ? "var(--color-ink)" : "var(--color-muted)",
                  userSelect: "none",
                  cursor: "pointer",
                }}
              >
                {colLabel(c)}
                <div
                  onMouseDown={(e) => s.handleColResizeStart(c, e)}
                  onDoubleClick={() => s.handleColResizeDoubleClick(c)}
                  style={{ position: "absolute", right: -2, top: 0, bottom: 0, width: 5, cursor: "col-resize", zIndex: 4 }}
                />
              </div>
            );
          })}
        </div>

        {/* Virtualized data rows */}
        {Array.from({ length: s.endRow - s.startRow + 1 }, (_, i) => {
          const r = s.startRow + i;
          if (s.editingRow !== null && r !== s.editingRow && r === s.endRow && s.editingRow > s.endRow) return null;
          // Skip rows that are covered by the frozen row stickied at the top
          if (s.frozenRows > 0 && r < s.frozenRows) return null;
          const rowSel = s.selectionRange
            ? r >= s.selectionRange.minR && r <= s.selectionRange.maxR
            : parseCellRef(s.selectedCell)?.row === r;
          return (
            <div key={r} style={{ position: "absolute", top: (r + 1) * s.ROW_HEIGHT, left: 0, height: s.ROW_HEIGHT, display: "flex", width: totalW }}>
              <div
                onMouseDown={(e) => {
                  if (e.button !== 0) return;
                  s.select(cellId(0, r));
                  s.setSelectionEnd(cellId(s.sheet.colCount - 1, r));
                  s.gridRef.current?.focus();
                }}
                onContextMenu={(e) => {
                  e.preventDefault(); e.stopPropagation();
                  s.select(cellId(0, r));
                  s.setSelectionEnd(cellId(s.sheet.colCount - 1, r));
                  s.setContextMenu({ x: e.clientX, y: e.clientY, cellId: cellId(0, r), type: "row" });
                }}
                style={{
                  position: "sticky", left: 0, zIndex: 1, width: s.HEADER_WIDTH, flexShrink: 0,
                  display: "flex", alignItems: "center", justifyContent: "center",
                  background: rowSel ? "var(--color-line)" : "var(--color-panel)",
                  borderBottom: "1px solid var(--color-line)",
                  borderRight: "1px solid var(--color-line)",
                  fontSize: "0.6875rem", fontWeight: 500,
                  color: rowSel ? "var(--color-ink)" : "var(--color-muted)",
                  userSelect: "none", cursor: "pointer",
                }}
              >
                {r + 1}
              </div>
              {Array.from({ length: s.sheet.colCount }, (_, c) => {
                const id = cellId(c, r);
                const isSelected = s.selectedCell === id;
                const isEditing = s.editingCell === id;
                const isPointed = s.pointHighlight.has(id);
                const inRange = s.selectionRange?.cells.has(id) ?? false;
                const display = s.displayValues[id] ?? "";
                const fmt = s.sheet.formats[id];
                const isNum = !isNaN(parseFloat(display)) && display !== "" && !isError(display);
                const isFindMatch = findSet.has(id);
                const isCurrentMatch = id === currentMatch;
                const isFillPreview = fillPreview.has(id);
                const isAnchor = id === anchorCell;
                const zebraColor = r % 2 === 1 ? "var(--color-zebra)" : undefined;
                const w = s.getColWidth(c);

                return (
                  <div
                    key={c}
                    onMouseDown={(e) => s.handleCellMouseDown(id, e)}
                    onDoubleClick={() => { if (!s.editingCell) s.startEditing(id); }}
                    onContextMenu={(e) => {
                      e.preventDefault(); e.stopPropagation();
                      s.select(id);
                      s.setContextMenu({ x: e.clientX, y: e.clientY, cellId: id });
                    }}
                    style={{
                      width: w, flexShrink: 0, position: "relative",
                      borderBottom: "1px solid var(--color-line)",
                      borderRight: "1px solid var(--color-line)",
                      outline: isSelected ? "2px solid var(--color-accent)" : isPointed ? "2px solid #5b8cd6" : "none",
                      outlineOffset: "-1px",
                      background: isCurrentMatch ? "rgba(255,180,0,0.35)"
                        : isFindMatch ? "rgba(255,220,80,0.2)"
                        : isFillPreview ? "rgba(192,133,82,0.20)"
                        : isPointed && !isSelected ? "rgba(91,140,214,0.12)"
                        : inRange && !isSelected ? "rgba(192,133,82,0.10)"
                        : fmt?.bg ?? zebraColor,
                      zIndex: isSelected || isPointed || isAnchor ? 2 : 0,
                      cursor: "cell",
                    }}
                  >
                    {isEditing ? (
                      <input
                        ref={s.editInputRef}
                        value={s.editValue}
                        onChange={(e) => s.handleEditInputChange(e.target.value)}
                        onKeyDown={s.handleInputKeyDown}
                        style={{
                          position: "absolute", inset: 0, border: "none", outline: "none",
                          background: "var(--color-paper)", padding: "0 4px", fontSize: "0.8125rem",
                          color: "var(--color-ink)", width: "100%", height: "100%",
                        }}
                      />
                    ) : (
                      <div style={{
                        padding: "0 4px", fontSize: "0.8125rem", lineHeight: `${s.ROW_HEIGHT}px`,
                        overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
                        color: isError(display) ? "#d94040" : fmt?.color ?? "var(--color-ink)",
                        textAlign: fmt?.align ?? (isNum ? "right" : "left"),
                        fontWeight: fmt?.bold ? 700 : isError(display) ? 600 : 400,
                        fontStyle: fmt?.italic ? "italic" : undefined,
                      }}>
                        {s.formatDisplay(display)}
                      </div>
                    )}
                    {isAnchor && !isEditing && (
                      <div
                        onMouseDown={startAutoFill}
                        title="Drag to fill"
                        style={{
                          position: "absolute", right: -3, bottom: -3, width: 8, height: 8,
                          background: "var(--color-accent)",
                          border: "1px solid var(--color-paper)",
                          cursor: "crosshair", zIndex: 5,
                        }}
                      />
                    )}
                  </div>
                );
              })}
            </div>
          );
        })}

        {/* Frozen rows -- rendered as sticky elements that stay below the column header */}
        {s.frozenRows > 0 && Array.from({ length: s.frozenRows }, (_, r) => {
          const rowSel = s.selectionRange
            ? r >= s.selectionRange.minR && r <= s.selectionRange.maxR
            : parseCellRef(s.selectedCell)?.row === r;
          return (
            <div key={`frozen-${r}`} style={{ position: "sticky", top: s.ROW_HEIGHT * (r + 1), zIndex: 9, height: s.ROW_HEIGHT, display: "flex", width: totalW, marginTop: -s.ROW_HEIGHT }}>
              <div style={{
                position: "sticky", left: 0, zIndex: 1, width: s.HEADER_WIDTH, flexShrink: 0,
                display: "flex", alignItems: "center", justifyContent: "center",
                background: rowSel ? "var(--color-line)" : "var(--color-panel)",
                borderBottom: "2px solid var(--color-accent)",
                borderRight: "1px solid var(--color-line)",
                fontSize: "0.6875rem", fontWeight: 500,
                color: rowSel ? "var(--color-ink)" : "var(--color-muted)",
                userSelect: "none",
              }}>
                {r + 1}
              </div>
              {Array.from({ length: s.sheet.colCount }, (_, c) => {
                const id = cellId(c, r);
                const isSelected = s.selectedCell === id;
                const isPointed = s.pointHighlight.has(id);
                const inRange = s.selectionRange?.cells.has(id) ?? false;
                const display = s.displayValues[id] ?? "";
                const fmt = s.sheet.formats[id];
                const isNum = !isNaN(parseFloat(display)) && display !== "" && !isError(display);
                const w = s.getColWidth(c);

                return (
                  <div
                    key={c}
                    onMouseDown={(e) => s.handleCellMouseDown(id, e)}
                    onDoubleClick={() => { if (!s.editingCell) s.startEditing(id); }}
                    onContextMenu={(e) => {
                      e.preventDefault(); e.stopPropagation();
                      s.select(id);
                      s.setContextMenu({ x: e.clientX, y: e.clientY, cellId: id });
                    }}
                    style={{
                      width: w, flexShrink: 0, position: "relative",
                      borderBottom: "2px solid var(--color-accent)",
                      borderRight: "1px solid var(--color-line)",
                      outline: isSelected ? "2px solid var(--color-accent)" : isPointed ? "2px solid #5b8cd6" : "none",
                      outlineOffset: "-1px",
                      background: isPointed && !isSelected ? "rgba(91,140,214,0.12)"
                        : inRange && !isSelected ? "rgba(192,133,82,0.10)"
                        : fmt?.bg ?? "var(--color-paper)",
                      cursor: "cell",
                    }}
                  >
                    <div style={{
                      padding: "0 4px", fontSize: "0.8125rem", lineHeight: `${s.ROW_HEIGHT}px`,
                      overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
                      color: isError(display) ? "#d94040" : fmt?.color ?? "var(--color-ink)",
                      textAlign: fmt?.align ?? (isNum ? "right" : "left"),
                      fontWeight: fmt?.bold ? 700 : isError(display) ? 600 : 400,
                      fontStyle: fmt?.italic ? "italic" : undefined,
                    }}>
                      {s.formatDisplay(display)}
                    </div>
                  </div>
                );
              })}
            </div>
          );
        })}

        {/* Render editing row if off-screen */}
        {s.editingRow !== null && (s.editingRow < s.startRow || s.editingRow > s.endRow) && (() => {
          const r = s.editingRow;
          return (
            <div key={`edit-${r}`} style={{ position: "absolute", top: (r + 1) * s.ROW_HEIGHT, left: 0, height: s.ROW_HEIGHT, display: "flex", width: totalW }}>
              <div style={{ position: "sticky", left: 0, zIndex: 1, width: s.HEADER_WIDTH, flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center", background: "var(--color-panel)", borderBottom: "1px solid var(--color-line)", borderRight: "1px solid var(--color-line)", fontSize: "0.6875rem", color: "var(--color-muted)", userSelect: "none" }}>
                {r + 1}
              </div>
              {Array.from({ length: s.sheet.colCount }, (_, c) => {
                const id = cellId(c, r);
                const isEditing = s.editingCell === id;
                return (
                  <div key={c} style={{ width: s.getColWidth(c), flexShrink: 0, position: "relative", borderBottom: "1px solid var(--color-line)", borderRight: "1px solid var(--color-line)" }}>
                    {isEditing && (
                      <input ref={s.editInputRef} value={s.editValue}
                        onChange={(e) => s.handleEditInputChange(e.target.value)}
                        onKeyDown={s.handleInputKeyDown}
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
    </div>
  );
}
