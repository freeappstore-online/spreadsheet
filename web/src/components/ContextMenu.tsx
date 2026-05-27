import { parseCellRef } from "../lib/cell-refs";
import type { SpreadsheetState } from "../hooks/useSpreadsheet";

export function ContextMenu({ s }: { s: SpreadsheetState }) {
  if (!s.contextMenu) return null;

  const p = parseCellRef(s.contextMenu.cellId);
  if (!p) return null;

  const items = [
    { label: "Insert row above", action: () => { s.insertRow(p.row, "above"); s.setContextMenu(null); } },
    { label: "Insert row below", action: () => { s.insertRow(p.row, "below"); s.setContextMenu(null); } },
    { label: "Delete row", action: () => { s.deleteRow(p.row); s.setContextMenu(null); } },
    { label: "---", action: () => {} },
    { label: "Insert column left", action: () => { s.insertCol(p.col, "left"); s.setContextMenu(null); } },
    { label: "Insert column right", action: () => { s.insertCol(p.col, "right"); s.setContextMenu(null); } },
    { label: "Delete column", action: () => { s.deleteCol(p.col); s.setContextMenu(null); } },
    { label: "---2", action: () => {} },
    { label: "Sort A → Z", action: () => s.sortByColumn(p.col, true) },
    { label: "Sort Z → A", action: () => s.sortByColumn(p.col, false) },
  ];

  return (
    <div
      onMouseDown={(e) => e.stopPropagation()}
      style={{
        position: "fixed", left: s.contextMenu.x, top: s.contextMenu.y, zIndex: 100,
        background: "var(--color-paper)", border: "1px solid var(--color-line)",
        borderRadius: "var(--radius-btn)", boxShadow: "0 4px 16px rgba(0,0,0,0.12)",
        padding: "0.25rem 0", minWidth: "10rem",
      }}
    >
      {items.map((item, i) =>
        item.label.startsWith("---") ? (
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
}
