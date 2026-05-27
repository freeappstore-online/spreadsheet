import { colLabel, cellId, parseCellRef } from "../lib/cell-refs";
import type { SpreadsheetState } from "../hooks/useSpreadsheet";

const isError = (val: string) => val.startsWith("#");

export function Grid({ s }: { s: SpreadsheetState }) {
  const totalW = s.HEADER_WIDTH + Array.from({ length: s.sheet.colCount }, (_, c) => s.getColWidth(c)).reduce((a, b) => a + b, 0);
  const totalH = (s.sheet.rowCount + 1) * s.ROW_HEIGHT;
  const findSet = new Set(s.findMatches);
  const currentMatch = s.findMatches[s.findIndex];

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
          const rowSel = s.selectionRange
            ? r >= s.selectionRange.minR && r <= s.selectionRange.maxR
            : parseCellRef(s.selectedCell)?.row === r;
          return (
            <div key={r} style={{ position: "absolute", top: (r + 1) * s.ROW_HEIGHT, left: 0, height: s.ROW_HEIGHT, display: "flex", width: totalW }}>
              <div style={{
                position: "sticky", left: 0, zIndex: 1, width: s.HEADER_WIDTH, flexShrink: 0,
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
                        : isPointed && !isSelected ? "rgba(91,140,214,0.12)"
                        : inRange && !isSelected ? "rgba(192,133,82,0.10)"
                        : fmt?.bg ?? zebraColor,
                      zIndex: isSelected || isPointed ? 1 : 0,
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
