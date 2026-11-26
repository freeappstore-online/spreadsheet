import type { BridgeStatus } from "@freeappstore/sdk/bridge";

const DESKTOP_URL = "https://github.com/freeappstore-online/platform/tree/main/desktop#readme";

const STATUS: Record<BridgeStatus, { label: string; dot: string }> = {
  off: { label: "Off", dot: "var(--color-muted)" },
  connecting: { label: "Desktop app not running", dot: "#d4a017" },
  connected: { label: "Connected", dot: "#3a9d5d" },
  rejected: { label: "Rejected by the desktop app", dot: "#c0392b" },
};

export function AiAccess({ allowed, setAllowed, status }: { allowed: boolean; setAllowed: (b: boolean) => void; status: BridgeStatus }) {
  const shown = allowed ? STATUS[status] : STATUS.off;
  return (
    <span style={{ display: "inline-flex", alignItems: "center", flexWrap: "wrap", gap: "0.25rem 0.5rem", fontSize: "0.75rem", color: "var(--color-muted)", minWidth: 0 }}>
      <label
        title="Let an AI on this computer (e.g. Claude Code) read and edit this workbook through the FreeAppStore desktop app. Your data stays in this browser."
        style={{ display: "inline-flex", alignItems: "center", gap: "0.375rem", cursor: "pointer", fontWeight: 600, whiteSpace: "nowrap" }}
      >
        <input
          type="checkbox"
          role="switch"
          checked={allowed}
          onChange={(e) => setAllowed(e.target.checked)}
          onMouseDown={(e) => e.preventDefault()}
          style={{ accentColor: "var(--color-accent)", margin: 0 }}
        />
        Allow AI access
      </label>
      {allowed && (
        <span aria-live="polite" style={{ display: "inline-flex", alignItems: "center", gap: "0.25rem", whiteSpace: "nowrap" }}>
          <span aria-hidden style={{ width: 7, height: 7, borderRadius: "50%", background: shown.dot }} />
          {shown.label}
        </span>
      )}
      {allowed && status !== "connected" && (
        <a href={DESKTOP_URL} target="_blank" rel="noopener noreferrer" style={{ color: "var(--color-muted)" }}>
          Install and open the FreeAppStore desktop app
        </a>
      )}
    </span>
  );
}
