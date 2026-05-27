export function colLabel(index: number): string {
  let label = "";
  let i = index;
  while (i >= 0) {
    label = String.fromCharCode(65 + (i % 26)) + label;
    i = Math.floor(i / 26) - 1;
  }
  return label;
}

export function cellId(col: number, row: number): string {
  return `${colLabel(col)}${row + 1}`;
}

export function parseCellRef(ref: string): { col: number; row: number } | null {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;
  const letters = match[1]!;
  const rowNum = parseInt(match[2]!, 10);
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return { col: col - 1, row: rowNum - 1 };
}

export function clamp(v: number, max: number) {
  return Math.max(0, Math.min(v, max - 1));
}

export function expandRange(rangeStr: string): string[] {
  const [startStr, endStr] = rangeStr.split(":");
  if (!startStr || !endStr) return [];
  const start = parseCellRef(startStr);
  const end = parseCellRef(endStr);
  if (!start || !end) return [];
  const refs: string[] = [];
  for (let r = Math.min(start.row, end.row); r <= Math.max(start.row, end.row); r++) {
    for (let c = Math.min(start.col, end.col); c <= Math.max(start.col, end.col); c++) {
      refs.push(cellId(c, r));
    }
  }
  return refs;
}
