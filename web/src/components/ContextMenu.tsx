import { parseCellRef } from "../lib/cell-refs";
import type { SpreadsheetState } from "../hooks/useSpreadsheet";

export function ContextMenu({ s }: { s: SpreadsheetState }) {
  if (!s.contextMenu) return null;

  const p = parseCellRef(s.contextMenu.cellId);
  if (!p) return null;

  const type = s.contextMenu.type ?? "cell";
  const close = () => s.setContextMenu(null);

  const colItems = [
    { label: "Insert column left", action: () => { s.insertCol(p.col, "left"); close(); } },
    { label: "Insert column right", action: () => { s.insertCol(p.col, "right"); close(); } },
    { label: "Delete column", action: () => { s.deleteCol(p.col); close(); } },
    { label: "---", action: () => {} },
    { label: "Sort A → Z", action: () => s.sortByColumn(p.col, true) },
    { label: "Sort Z → A", action: () => s.sortByColumn(p.col, false) },
  ];

  const rowItems = [
    { label: "Insert row above", action: () => { s.insertRow(p.row, "above"); close(); } },
    { label: "Insert row below", action: () => { s.insertRow(p.row, "below"); close(); } },
    { label: "Delete row", action: () => { s.deleteRow(p.row); close(); } },
    { label: "---", action: () => {} },
    {
      label: s.frozenRows > 0 ? "Unfreeze first row" : "Freeze first row",
      action: () => { s.toggleFreezeFirstRow(); close(); },
    },
  ];

  const cellItems = [
    { label: "Insert row above", action: () => { s.insertRow(p.row, "above"); close(); } },
    { label: "Insert row below", action: () => { s.insertRow(p.row, "below"); close(); } },
    { label: "Delete row", action: () => { s.deleteRow(p.row); close(); } },
    { label: "---", action: () => {} },
    { label: "Insert column left", action: () => { s.insertCol(p.col, "left"); close(); } },
    { label: "Insert column right", action: () => { s.insertCol(p.col, "right"); close(); } },
    { label: "Delete column", action: () => { s.deleteCol(p.col); close(); } },
    { label: "---2", action: () => {} },
    { label: "Sort A → Z", action: () => s.sortByColumn(p.col, true) },
    { label: "Sort Z → A", action: () => s.sortByColumn(p.col, false) },
    { label: "---3", action: () => {} },
    {
      label: s.frozenRows > 0 ? "Unfreeze first row" : "Freeze first row",
      action: () => { s.toggleFreezeFirstRow(); close(); },
    },
  ];

  const items = type === "col" ? colItems : type === "row" ? rowItems : cellItems;

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
