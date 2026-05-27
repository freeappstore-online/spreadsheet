import type { SpreadsheetState } from "../hooks/useSpreadsheet";

export function ColorPicker({ s }: { s: SpreadsheetState }) {
  if (!s.colorPicker) return null;

  return (
    <div
      onMouseDown={(e) => e.stopPropagation()}
      style={{
        position: "fixed", left: s.colorPicker.x, top: s.colorPicker.y, zIndex: 100,
        background: "var(--color-paper)", border: "1px solid var(--color-line)",
        borderRadius: "var(--radius-btn)", boxShadow: "0 4px 16px rgba(0,0,0,0.12)",
        padding: "0.5rem", display: "grid", gridTemplateColumns: "repeat(6, 1.5rem)", gap: "0.25rem",
      }}
    >
      {[
        undefined, "#ef4444", "#f97316", "#eab308", "#22c55e", "#3b82f6",
        "#8b5cf6", "#ec4899", "#14b8a6", "#6b7280", "#1e293b", "#fefce8",
        "#fee2e2", "#dbeafe", "#dcfce7", "#f3e8ff", "#fef3c7",
      ].map((c, i) => (
        <button
          key={i}
          onClick={() => {
            s.applyFormat({ [s.colorPicker!.type]: c });
            s.setColorPicker(null);
          }}
          style={{
            width: "1.5rem", height: "1.5rem",
            border: c ? "1px solid var(--color-line)" : "2px dashed var(--color-muted)",
            borderRadius: "0.25rem",
            background: c ?? "transparent",
            cursor: "pointer",
          }}
          title={c ?? "Clear"}
        />
      ))}
    </div>
  );
}
