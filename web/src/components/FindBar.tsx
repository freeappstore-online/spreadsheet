import type { SpreadsheetState } from "../hooks/useSpreadsheet";

export function FindBar({ s }: { s: SpreadsheetState }) {
  return (
    <div
      style={{
        display: "flex", alignItems: "center", gap: "0.375rem",
        padding: "0.375rem 0.75rem",
        borderTop: "1px solid var(--color-line)",
        background: "var(--color-panel)", flexShrink: 0, flexWrap: "wrap",
      }}
    >
      <input
        ref={s.findInputRef}
        value={s.findBar.query}
        onChange={(e) => s.setFindBar((p) => ({ ...p, query: e.target.value }))}
        onKeyDown={(e) => {
          if (e.key === "Enter") { e.preventDefault(); e.shiftKey ? s.findPrev() : s.findNext(); }
          if (e.key === "Escape") s.closeFindBar();
        }}
        placeholder="Find..."
        style={{
          border: "1px solid var(--color-line)", borderRadius: "var(--radius-btn)",
          padding: "0.25rem 0.5rem", fontSize: "0.8125rem", background: "var(--color-paper)",
          color: "var(--color-ink)", outline: "none", width: "10rem",
        }}
      />
      {s.findBar.showReplace && (
        <input
          value={s.findBar.replace}
          onChange={(e) => s.setFindBar((p) => ({ ...p, replace: e.target.value }))}
          onKeyDown={(e) => { if (e.key === "Escape") s.closeFindBar(); }}
          placeholder="Replace..."
          style={{
            border: "1px solid var(--color-line)", borderRadius: "var(--radius-btn)",
            padding: "0.25rem 0.5rem", fontSize: "0.8125rem", background: "var(--color-paper)",
            color: "var(--color-ink)", outline: "none", width: "10rem",
          }}
        />
      )}
      <span style={{ fontSize: "0.6875rem", color: "var(--color-muted)" }}>
        {s.findMatches.length > 0 ? `${s.findIndex + 1}/${s.findMatches.length}` : s.findBar.query ? "0 results" : ""}
      </span>
      {[
        { label: "↑", action: s.findPrev },
        { label: "↓", action: s.findNext },
        ...(s.findBar.showReplace ? [
          { label: "Replace", action: s.replaceOne },
          { label: "All", action: s.replaceAll },
        ] : []),
        { label: "×", action: s.closeFindBar },
      ].map(({ label, action }) => (
        <button
          key={label}
          onClick={action}
          style={{
            border: "1px solid var(--color-line)", borderRadius: "var(--radius-btn)",
            background: "transparent", color: "var(--color-muted)",
            padding: "0.125rem 0.5rem", fontSize: "0.75rem", cursor: "pointer", fontWeight: 600,
          }}
        >
          {label}
        </button>
      ))}
    </div>
  );
}
