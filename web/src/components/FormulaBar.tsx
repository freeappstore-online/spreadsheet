import type { SpreadsheetState } from "../hooks/useSpreadsheet";

export function FormulaBar({ s }: { s: SpreadsheetState }) {
  return (
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
        {s.selectedCell}
      </span>
      <span style={{ color: "var(--color-line)" }}>|</span>
      <input
        ref={s.formulaBarRef}
        value={s.formulaBarValue}
        onChange={(e) => s.handleFormulaBarChange(e.target.value)}
        onFocus={s.handleFormulaBarFocus}
        onBlur={s.handleFormulaBarBlur}
        onKeyDown={s.handleFormulaBarKeyDown}
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
  );
}
