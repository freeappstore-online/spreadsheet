import type { Sheet, WorkbookData } from "../types";
import { STORAGE_KEY, DEFAULT_COLS, DEFAULT_ROWS } from "../constants";

export function makeSheet(name: string): Sheet {
  return { name, cells: {}, formats: {}, colCount: DEFAULT_COLS, rowCount: DEFAULT_ROWS };
}

export function loadWorkbook(): WorkbookData {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (parsed.sheets) return parsed as WorkbookData;
      return {
        sheets: [{
          name: "Sheet 1",
          cells: parsed.cells ?? {},
          formats: {},
          colCount: parsed.colCount ?? DEFAULT_COLS,
          rowCount: parsed.rowCount ?? DEFAULT_ROWS,
        }],
        activeSheet: 0,
      };
    }
  } catch { /* ignore */ }
  return { sheets: [makeSheet("Sheet 1")], activeSheet: 0 };
}

export function saveWorkbook(data: WorkbookData) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
}
