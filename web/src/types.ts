export interface CellFormat {
  bold?: boolean;
  italic?: boolean;
  align?: "left" | "center" | "right";
  bg?: string;
  color?: string;
}

export interface Sheet {
  name: string;
  cells: Record<string, string>;
  formats: Record<string, CellFormat>;
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
