import type { SpreadsheetState } from "../hooks/useSpreadsheet";

export function HelpDialog({ s }: { s: SpreadsheetState }) {
  return (
    <div
      onClick={() => s.setShowHelp(false)}
      style={{
        position: "fixed", inset: 0, zIndex: 200,
        background: "rgba(0,0,0,0.4)",
        display: "flex", alignItems: "center", justifyContent: "center",
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: "var(--color-paper)", borderRadius: "var(--radius-card)",
          padding: "1.5rem", maxWidth: "28rem", width: "90vw",
          maxHeight: "80vh", overflow: "auto",
          boxShadow: "0 8px 32px rgba(0,0,0,0.2)",
        }}
      >
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1rem" }}>
          <h2 style={{ margin: 0, fontSize: "1.125rem" }}>Keyboard Shortcuts</h2>
          <button onClick={() => s.setShowHelp(false)} style={{ border: "none", background: "transparent", fontSize: "1.25rem", cursor: "pointer", color: "var(--color-muted)" }}>×</button>
        </div>
        {[
          ["Navigation", [
            ["Arrow keys", "Move selection"],
            ["Shift + Arrow", "Extend selection"],
            ["Tab / Shift+Tab", "Move right / left"],
            ["Enter / Shift+Enter", "Move down / up"],
            ["Ctrl+A", "Select all"],
          ]],
          ["Editing", [
            ["Type any character", "Start editing cell"],
            ["F2 / Enter", "Edit selected cell"],
            ["Escape", "Cancel edit"],
            ["Delete / Backspace", "Clear cell(s)"],
          ]],
          ["Formulas", [
            ["= then Arrow keys", "Insert cell reference"],
            ["Shift + Arrow (in formula)", "Extend to range (A1:B3)"],
            ["Click cell (in formula)", "Insert cell reference"],
          ]],
          ["Formatting", [
            ["Ctrl+B", "Bold"],
            ["Ctrl+I", "Italic"],
          ]],
          ["Tools", [
            ["Ctrl+C / X / V", "Copy / Cut / Paste"],
            ["Ctrl+Z", "Undo"],
            ["Ctrl+Shift+Z / Ctrl+Y", "Redo"],
            ["Ctrl+F", "Find"],
            ["Ctrl+H", "Find & Replace"],
            ["Right-click", "Context menu (insert/delete/sort)"],
            ["Ctrl+/", "This help"],
          ]],
        ].map(([title, shortcuts]) => (
          <div key={title as string} style={{ marginBottom: "0.75rem" }}>
            <div style={{ fontSize: "0.75rem", fontWeight: 700, color: "var(--color-accent)", marginBottom: "0.25rem", textTransform: "uppercase", letterSpacing: "0.05em" }}>
              {title as string}
            </div>
            {(shortcuts as string[][]).map(([key, desc]) => (
              <div key={key} style={{ display: "flex", justifyContent: "space-between", padding: "0.125rem 0", fontSize: "0.8125rem" }}>
                <code style={{ background: "var(--color-panel)", padding: "0.05rem 0.375rem", borderRadius: "0.25rem", fontSize: "0.75rem" }}>{key}</code>
                <span style={{ color: "var(--color-muted)" }}>{desc}</span>
              </div>
            ))}
          </div>
        ))}
      </div>
    </div>
  );
}
