import type { SpreadsheetState } from "../hooks/useSpreadsheet";

export function StatusBar({ s }: { s: SpreadsheetState }) {
  return (
    <div
      style={{
        display: "flex", alignItems: "center", gap: "1.25rem",
        padding: "0.25rem 0.75rem",
        borderTop: "1px solid var(--color-line)",
        background: "var(--color-panel)", flexShrink: 0,
        minHeight: "1.5rem", fontSize: "0.6875rem", color: "var(--color-muted)",
      }}
    >
      {s.selectionStats.numCount > 1 && (
        <>
          <span>Sum: <b style={{ color: "var(--color-ink)" }}>{s.selectionStats.sum.toLocaleString("en-US", { maximumFractionDigits: 4 })}</b></span>
          <span>Avg: <b style={{ color: "var(--color-ink)" }}>{s.selectionStats.avg.toLocaleString("en-US", { maximumFractionDigits: 4 })}</b></span>
        </>
      )}
      {s.selectionStats.count > 0 && (
        <span>Count: <b style={{ color: "var(--color-ink)" }}>{s.selectionStats.count}</b></span>
      )}
      <div style={{ flex: 1 }} />
      <span>{s.sheet.rowCount} rows × {s.sheet.colCount} cols</span>
    </div>
  );
}
