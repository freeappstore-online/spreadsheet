export interface CellFormat {
  bold?: boolean;
  italic?: boolean;
  align?: "left" | "center" | "right";
  bg?: string;
  color?: string;
  numFmt?: "currency" | "percent" | "decimal2";
}

export type CondRuleType = "greater" | "less" | "equal" | "contains" | "between";

export interface ConditionalRule {
  range: string;        // e.g. "A1:B10"
  type: CondRuleType;
  value: string;        // for between, "a..b"
  format: { bg?: string; color?: string; bold?: boolean };
}

export interface Sheet {
  name: string;
  cells: Record<string, string>;
  formats: Record<string, CellFormat>;
  conditionalRules?: ConditionalRule[];
  colCount: number;
  rowCount: number;
}

export interface WorkbookData {
  sheets: Sheet[];
  activeSheet: number;
}

export interface PointMode {
  refStart: number;
  refEnd: number;
  anchor: { col: number; row: number };
  active: { col: number; row: number };
}
