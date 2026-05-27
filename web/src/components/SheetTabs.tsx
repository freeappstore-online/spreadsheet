import type { SpreadsheetState } from "../hooks/useSpreadsheet";

export function SheetTabs({ s }: { s: SpreadsheetState }) {
  return (
    <div
      style={{
        display: "flex", alignItems: "center",
        borderTop: "1px solid var(--color-line)",
        background: "var(--color-panel)", flexShrink: 0,
        overflowX: "auto", minHeight: "1.75rem",
      }}
    >
      {s.workbook.sheets.map((sh, i) => (
        <button
          key={i}
          onClick={() => s.switchSheet(i)}
          onDoubleClick={() => s.renameSheet(i)}
          onContextMenu={(e) => {
            e.preventDefault();
            if (s.workbook.sheets.length > 1 && confirm(`Delete "${sh.name}"?`)) s.deleteSheet(i);
          }}
          style={{
            padding: "0.25rem 0.75rem",
            fontSize: "0.6875rem", fontWeight: i === s.workbook.activeSheet ? 700 : 400,
            border: "none", borderRight: "1px solid var(--color-line)",
            background: i === s.workbook.activeSheet ? "var(--color-paper)" : "transparent",
            color: i === s.workbook.activeSheet ? "var(--color-ink)" : "var(--color-muted)",
            cursor: "pointer", whiteSpace: "nowrap",
          }}
        >
          {sh.name}
        </button>
      ))}
      <button
        onClick={s.addSheet}
        style={{
          padding: "0.25rem 0.5rem", fontSize: "0.8125rem", fontWeight: 600,
          border: "none", background: "transparent",
          color: "var(--color-muted)", cursor: "pointer",
        }}
      >
        +
      </button>
    </div>
  );
}
