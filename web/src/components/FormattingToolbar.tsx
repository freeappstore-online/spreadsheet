import type { SpreadsheetState } from "../hooks/useSpreadsheet";

export function FormattingToolbar({ s }: { s: SpreadsheetState }) {
  return (
    <div
      style={{
        display: "flex", alignItems: "center", gap: "0.25rem",
        padding: "0.25rem 0.75rem",
        borderBottom: "1px solid var(--color-line)",
        background: "var(--color-panel)", flexShrink: 0, flexWrap: "wrap",
      }}
    >
      {[
        { label: "B", title: "Bold (Ctrl+B)", active: s.sheet.formats[s.selectedCell]?.bold, action: s.toggleBold, style: { fontWeight: 800 } as React.CSSProperties },
        { label: "I", title: "Italic (Ctrl+I)", active: s.sheet.formats[s.selectedCell]?.italic, action: s.toggleItalic, style: { fontStyle: "italic" } as React.CSSProperties },
      ].map(({ label, title, active, action, style }) => (
        <button
          key={label}
          title={title}
          onMouseDown={(e) => e.preventDefault()}
          onClick={action}
          style={{
            width: "1.75rem", height: "1.75rem",
            border: active ? "1px solid var(--color-accent)" : "1px solid var(--color-line)",
            borderRadius: "0.25rem",
            background: active ? "rgba(192,133,82,0.12)" : "transparent",
            color: active ? "var(--color-accent)" : "var(--color-muted)",
            cursor: "pointer", fontSize: "0.8125rem",
            display: "flex", alignItems: "center", justifyContent: "center",
            ...style,
          }}
        >
          {label}
        </button>
      ))}
      <span style={{ width: 1, height: "1rem", background: "var(--color-line)", margin: "0 0.25rem" }} />
      {(["left", "center", "right"] as const).map((a) => (
        <button
          key={a}
          title={`Align ${a}`}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => s.applyFormat({ align: s.sheet.formats[s.selectedCell]?.align === a ? undefined : a })}
          style={{
            width: "1.75rem", height: "1.75rem",
            border: s.sheet.formats[s.selectedCell]?.align === a ? "1px solid var(--color-accent)" : "1px solid var(--color-line)",
            borderRadius: "0.25rem",
            background: s.sheet.formats[s.selectedCell]?.align === a ? "rgba(192,133,82,0.12)" : "transparent",
            color: s.sheet.formats[s.selectedCell]?.align === a ? "var(--color-accent)" : "var(--color-muted)",
            cursor: "pointer", fontSize: "0.6875rem", fontWeight: 600,
            display: "flex", alignItems: "center", justifyContent: "center",
          }}
        >
          {a === "left" ? "⯷" : a === "center" ? "⯶" : "⯸"}
        </button>
      ))}
      <span style={{ width: 1, height: "1rem", background: "var(--color-line)", margin: "0 0.25rem" }} />
      {/* Background color */}
      <div style={{ position: "relative" }}>
        <button
          title="Background color"
          onMouseDown={(e) => e.preventDefault()}
          onClick={(e) => {
            const rect = e.currentTarget.getBoundingClientRect();
            s.setColorPicker(s.colorPicker?.type === "bg" ? null : { type: "bg", x: rect.left, y: rect.bottom + 4 });
          }}
          style={{
            width: "1.75rem", height: "1.75rem",
            border: "1px solid var(--color-line)", borderRadius: "0.25rem",
            background: "transparent", cursor: "pointer",
            display: "flex", alignItems: "center", justifyContent: "center",
            fontSize: "0.75rem", color: "var(--color-muted)",
          }}
        >
          <span style={{ display: "block", width: 12, height: 12, borderRadius: 2, background: s.sheet.formats[s.selectedCell]?.bg || "var(--color-line)" }} />
        </button>
      </div>
      {/* Text color */}
      <div style={{ position: "relative" }}>
        <button
          title="Text color"
          onMouseDown={(e) => e.preventDefault()}
          onClick={(e) => {
            const rect = e.currentTarget.getBoundingClientRect();
            s.setColorPicker(s.colorPicker?.type === "color" ? null : { type: "color", x: rect.left, y: rect.bottom + 4 });
          }}
          style={{
            width: "1.75rem", height: "1.75rem",
            border: "1px solid var(--color-line)", borderRadius: "0.25rem",
            background: "transparent", cursor: "pointer",
            display: "flex", alignItems: "center", justifyContent: "center",
            fontSize: "0.8125rem", fontWeight: 700,
            color: s.sheet.formats[s.selectedCell]?.color || "var(--color-muted)",
          }}
        >
          A
        </button>
      </div>
      <span style={{ width: 1, height: "1rem", background: "var(--color-line)", margin: "0 0.25rem" }} />
      <button
        title="Keyboard shortcuts (?)"
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => s.setShowHelp((p) => !p)}
        style={{
          width: "1.75rem", height: "1.75rem",
          border: "1px solid var(--color-line)", borderRadius: "0.25rem",
          background: "transparent", color: "var(--color-muted)", cursor: "pointer",
          fontSize: "0.8125rem", fontWeight: 600,
          display: "flex", alignItems: "center", justifyContent: "center",
        }}
      >
        ?
      </button>
    </div>
  );
}
