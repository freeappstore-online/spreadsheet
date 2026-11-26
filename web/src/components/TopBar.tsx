import type { SpreadsheetState } from "../hooks/useSpreadsheet";
import type { useAiAccess } from "../hooks/useAiAccess";
import { AiAccess } from "./AiAccess";

export function TopBar({ s, aiAccess }: { s: SpreadsheetState; aiAccess: ReturnType<typeof useAiAccess> }) {
  return (
    <div
      style={{
        display: "flex",
        alignItems: "center",
        flexWrap: "wrap",
        gap: "0.5rem 0.75rem",
        padding: "0.5rem 0.75rem",
        borderBottom: "1px solid var(--color-line)",
        background: "var(--color-panel)",
        flexShrink: 0,
      }}
    >
      <span style={{ fontFamily: "'Fraunces', Georgia, serif", fontWeight: 700, fontSize: "1.125rem", whiteSpace: "nowrap" }}>
        spreadsheet
      </span>
      <AiAccess {...aiAccess} />
      <div style={{ flex: 1 }} />
      {[
        { label: "Import CSV", action: s.importCsv },
        { label: "Export CSV", action: s.exportCsv },
        { label: "Clear all", action: s.clearAll },
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
            whiteSpace: "nowrap",
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
  );
}
