import { useState } from "react";
import type { SpreadsheetState } from "../hooks/useSpreadsheet";
import type { CondRuleType } from "../types";

export function ConditionalFormatDialog({ s }: { s: SpreadsheetState }) {
  const [range, setRange] = useState(() => {
    if (s.selectionRange) {
      const { minC, maxC, minR, maxR } = s.selectionRange;
      const cellId = (c: number, r: number) => String.fromCharCode(65 + c) + (r + 1);
      return `${cellId(minC, minR)}:${cellId(maxC, maxR)}`;
    }
    return s.selectedCell;
  });
  const [type, setType] = useState<CondRuleType>("greater");
  const [value, setValue] = useState("");
  const [bg, setBg] = useState("#fee2e2");
  const [color, setColor] = useState("");
  const [bold, setBold] = useState(false);

  const close = () => s.setCondFormatDialog(false);
  const rules = s.sheet.conditionalRules ?? [];

  const add = () => {
    if (!value.trim() || !range.trim()) return;
    s.addConditionalRule({
      range: range.trim(),
      type,
      value: value.trim(),
      format: {
        ...(bg ? { bg } : {}),
        ...(color ? { color } : {}),
        ...(bold ? { bold } : {}),
      },
    });
    setValue("");
  };

  return (
    <div
      onClick={close}
      style={{
        position: "fixed", inset: 0, zIndex: 200, background: "rgba(0,0,0,0.4)",
        display: "flex", alignItems: "center", justifyContent: "center",
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: "var(--color-paper)", borderRadius: "var(--radius-card)",
          padding: "1.5rem", maxWidth: "32rem", width: "90vw",
          maxHeight: "80vh", overflow: "auto",
          boxShadow: "0 8px 32px rgba(0,0,0,0.2)",
        }}
      >
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1rem" }}>
          <h2 style={{ margin: 0, fontSize: "1.125rem" }}>Conditional formatting</h2>
          <button onClick={close} style={{ border: "none", background: "transparent", fontSize: "1.25rem", cursor: "pointer", color: "var(--color-muted)" }}>×</button>
        </div>

        {/* Existing rules */}
        {rules.length > 0 && (
          <div style={{ marginBottom: "1rem" }}>
            <div style={{ fontSize: "0.75rem", fontWeight: 700, color: "var(--color-accent)", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: "0.5rem" }}>
              Active rules
            </div>
            {rules.map((rule, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: "0.5rem", padding: "0.375rem 0.5rem", borderRadius: "0.25rem", background: "var(--color-panel)", marginBottom: "0.25rem" }}>
                <div style={{ width: 18, height: 18, borderRadius: 3, background: rule.format.bg || "var(--color-line)", border: "1px solid var(--color-line)" }} />
                <div style={{ flex: 1, fontSize: "0.8125rem" }}>
                  <code style={{ background: "var(--color-paper)", padding: "0 0.25rem", borderRadius: "0.125rem" }}>{rule.range}</code>
                  {" "}{rule.type}{" "}
                  <code style={{ background: "var(--color-paper)", padding: "0 0.25rem", borderRadius: "0.125rem" }}>{rule.value}</code>
                </div>
                <button
                  onClick={() => s.removeConditionalRule(i)}
                  style={{ border: "1px solid var(--color-line)", background: "transparent", borderRadius: "0.25rem", padding: "0.125rem 0.5rem", cursor: "pointer", color: "var(--color-muted)", fontSize: "0.75rem" }}
                >
                  Remove
                </button>
              </div>
            ))}
          </div>
        )}

        {/* Add new rule */}
        <div style={{ fontSize: "0.75rem", fontWeight: 700, color: "var(--color-accent)", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: "0.5rem" }}>
          Add rule
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "auto 1fr", gap: "0.5rem", alignItems: "center", marginBottom: "1rem" }}>
          <label style={{ fontSize: "0.8125rem", color: "var(--color-muted)" }}>Range</label>
          <input value={range} onChange={(e) => setRange(e.target.value)} placeholder="A1:B10"
            style={{ padding: "0.375rem", border: "1px solid var(--color-line)", borderRadius: "0.25rem", background: "var(--color-paper)", color: "var(--color-ink)", fontSize: "0.8125rem", outline: "none" }} />

          <label style={{ fontSize: "0.8125rem", color: "var(--color-muted)" }}>Condition</label>
          <div style={{ display: "flex", gap: "0.5rem" }}>
            <select value={type} onChange={(e) => setType(e.target.value as CondRuleType)}
              style={{ padding: "0.375rem", border: "1px solid var(--color-line)", borderRadius: "0.25rem", background: "var(--color-paper)", color: "var(--color-ink)", fontSize: "0.8125rem" }}>
              <option value="greater">is greater than</option>
              <option value="less">is less than</option>
              <option value="equal">equals</option>
              <option value="contains">contains</option>
              <option value="between">is between (e.g. 1..10)</option>
            </select>
            <input value={value} onChange={(e) => setValue(e.target.value)} placeholder={type === "between" ? "1..10" : "100"}
              style={{ flex: 1, padding: "0.375rem", border: "1px solid var(--color-line)", borderRadius: "0.25rem", background: "var(--color-paper)", color: "var(--color-ink)", fontSize: "0.8125rem", outline: "none" }} />
          </div>

          <label style={{ fontSize: "0.8125rem", color: "var(--color-muted)" }}>Bg color</label>
          <div style={{ display: "flex", gap: "0.25rem", flexWrap: "wrap" }}>
            {["", "#fee2e2", "#fef3c7", "#dcfce7", "#dbeafe", "#f3e8ff", "#ffe4e6"].map((c) => (
              <button key={c || "none"} type="button" onClick={() => setBg(c)}
                style={{ width: 20, height: 20, borderRadius: 3, border: bg === c ? "2px solid var(--color-accent)" : "1px solid var(--color-line)", background: c || "transparent", cursor: "pointer" }}
                title={c || "none"} />
            ))}
          </div>

          <label style={{ fontSize: "0.8125rem", color: "var(--color-muted)" }}>Text color</label>
          <div style={{ display: "flex", gap: "0.25rem", flexWrap: "wrap" }}>
            {["", "#ef4444", "#f97316", "#eab308", "#22c55e", "#3b82f6", "#8b5cf6"].map((c) => (
              <button key={c || "none"} type="button" onClick={() => setColor(c)}
                style={{ width: 20, height: 20, borderRadius: 3, border: color === c ? "2px solid var(--color-accent)" : "1px solid var(--color-line)", background: c || "transparent", cursor: "pointer" }}
                title={c || "none"} />
            ))}
          </div>

          <label style={{ fontSize: "0.8125rem", color: "var(--color-muted)" }}>Style</label>
          <label style={{ fontSize: "0.8125rem", display: "flex", alignItems: "center", gap: "0.25rem", cursor: "pointer" }}>
            <input type="checkbox" checked={bold} onChange={(e) => setBold(e.target.checked)} />
            Bold
          </label>
        </div>

        <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
          <button onClick={close} style={{ padding: "0.375rem 1rem", border: "1px solid var(--color-line)", borderRadius: "var(--radius-btn)", background: "transparent", color: "var(--color-muted)", cursor: "pointer", fontSize: "0.8125rem" }}>
            Close
          </button>
          <button onClick={add} style={{ padding: "0.375rem 1rem", border: "none", borderRadius: "var(--radius-btn)", background: "var(--color-accent)", color: "#fff", cursor: "pointer", fontSize: "0.8125rem", fontWeight: 600 }}>
            Add rule
          </button>
        </div>
      </div>
    </div>
  );
}
