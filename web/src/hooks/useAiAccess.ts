import { useEffect, useRef, useState } from "react";
import { exposeTools, type AppTool, type BridgeStatus, type ToolBridge } from "@freeappstore/sdk/bridge";
import type { SpreadsheetState } from "./useSpreadsheet";
import { MAX_CELLS } from "../lib/ai-tools";

const STORAGE_KEY = "spreadsheet_ai_access";

const sheetArg = {
  type: "string",
  description: "Sheet name (case-insensitive). Omit for the sheet the user is looking at.",
};

const FUNCTIONS =
  "SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, IF, ROUND, ABS, FLOOR, CEILING, INT, MOD, POWER, SQRT, SIGN, PI, " +
  "LEN, UPPER, LOWER, TRIM, LEFT, RIGHT, MID, CONCAT, SUBSTITUTE, VALUE, TEXT, NOW, TODAY";

function tools(ai: SpreadsheetState["ai"]): AppTool[] {
  return [
    {
      name: "list_sheets",
      description:
        "List the workbook's sheets with each one's used range (the smallest range holding every non-empty cell, null if empty), " +
        "filled cell count and grid size, plus the sheet and cell selection the user is looking at. Call this first.",
      inputSchema: { type: "object", properties: {} },
      handler: () => ai.listSheets(),
    },
    {
      name: "read_range",
      description:
        `Read a range in A1 notation, e.g. "A1:D20" or "B3". Returns "values" (computed results, row by row; plain numbers as numbers, ` +
        `empty cells as "") and "formulas" (the "=..." formula in each cell, or null). At most ${MAX_CELLS} cells per call; ` +
        "use list_sheets to find the used range.",
      inputSchema: {
        type: "object",
        properties: { sheet: sheetArg, range: { type: "string", description: 'A1 range, e.g. "A1:D20", or one cell, e.g. "B3".' } },
        required: ["range"],
      },
      handler: ({ sheet, range }) => ai.readRange(sheet, range),
    },
    {
      name: "set_cells",
      description:
        "Write values into cells, row by row, starting at the top-left of `range`. Strings starting with \"=\" are formulas, " +
        `e.g. "=SUM(B2:B9)" or "=A2*1.2"; use upper-case cell references. Functions: ${FUNCTIONS}. Other sheets: Sheet2!A1 ` +
        "(only for sheet names without spaces). An empty string or null clears a cell. The edit shows live, is saved in the browser " +
        `and the user can undo it (Cmd/Ctrl+Z). At most ${MAX_CELLS} cells per call; the sheet grows if needed.`,
      inputSchema: {
        type: "object",
        properties: {
          sheet: sheetArg,
          range: {
            type: "string",
            description: 'Start cell, e.g. "A1", or a range whose size matches `values` exactly, e.g. "A1:B2".',
          },
          values: {
            type: "array",
            description: 'Rows of cell values, e.g. [["Item", "Price"], ["Tea", 3.5], ["Total", "=SUM(B2:B2)"]].',
            items: { type: "array", items: { type: ["string", "number", "boolean", "null"] } },
          },
        },
        required: ["range", "values"],
      },
      handler: ({ sheet, range, values }) => ai.setCells(sheet, range, values),
    },
    {
      name: "clear_range",
      description: `Clear the contents of every cell in an A1 range (formatting stays). Undoable. At most ${MAX_CELLS} cells per call.`,
      inputSchema: {
        type: "object",
        properties: { sheet: sheetArg, range: { type: "string", description: 'A1 range, e.g. "A1:D20".' } },
        required: ["range"],
      },
      handler: ({ sheet, range }) => ai.clearRange(sheet, range),
    },
    {
      name: "add_sheet",
      description:
        'Add an empty sheet at the end of the workbook. Names must be unique; use letters, digits and "_" (no spaces) ' +
        "if formulas on other sheets will refer to it.",
      inputSchema: {
        type: "object",
        properties: { name: { type: "string", description: 'New sheet name. Default "Sheet <n>".' } },
      },
      handler: ({ name }) => ai.addSheet(name),
    },
    {
      name: "rename_sheet",
      description: "Rename a sheet. Formulas on other sheets that use the old name are not updated.",
      inputSchema: {
        type: "object",
        properties: {
          sheet: { type: "string", description: "Current sheet name." },
          new_name: { type: "string", description: "New, unique sheet name." },
        },
        required: ["sheet", "new_name"],
      },
      handler: ({ sheet, new_name }) => ai.renameSheet(sheet, new_name),
    },
  ];
}

function loadAllowed(): boolean {
  try { return localStorage.getItem(STORAGE_KEY) === "1"; } catch { return false; }
}

/** "Allow AI access" toggle: offers the sheet tools to the FreeAppStore desktop app while on. */
export function useAiAccess(ai: SpreadsheetState["ai"]) {
  const [allowed, setAllowed] = useState(loadAllowed);
  const [status, setStatus] = useState<BridgeStatus>("off");
  const bridgeRef = useRef<ToolBridge | null>(null);

  // The bridge is created on the first switch-on and kept; `ai` is stable.
  useEffect(() => {
    try { localStorage.setItem(STORAGE_KEY, allowed ? "1" : "0"); } catch { /* ignore */ }
    if (!allowed) {
      bridgeRef.current?.setEnabled(false);
      return;
    }
    if (bridgeRef.current) bridgeRef.current.setEnabled(true);
    else bridgeRef.current = exposeTools({ appId: "spreadsheet", tools: tools(ai), onStatus: setStatus });
  }, [allowed, ai]);

  useEffect(() => () => {
    bridgeRef.current?.close();
    bridgeRef.current = null;
  }, []);

  return { allowed, setAllowed, status };
}
