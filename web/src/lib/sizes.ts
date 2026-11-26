/** Custom column widths / row heights, keyed by index; missing means the default size. */
export type Sizes = Record<number, number>;

/**
 * Moves sizes after inserting (`delta` 1) or deleting (`delta` -1) a row or column at `at`,
 * so each custom size stays with its row or column. A deleted index loses its size.
 */
export function shiftSizes(sizes: Sizes | undefined, at: number, delta: 1 | -1): Sizes | undefined {
  if (!sizes) return sizes;
  const next: Sizes = {};
  for (const [key, size] of Object.entries(sizes)) {
    const i = Number(key);
    if (delta === -1 && i === at) continue;
    next[i >= at ? i + delta : i] = size;
  }
  return next;
}

/** Sets one size, or removes it (back to the default) when `size` is null. */
export function setSize(sizes: Sizes | undefined, index: number, size: number | null): Sizes {
  const next = { ...sizes };
  if (size === null) delete next[index];
  else next[index] = size;
  return next;
}

/** offsets[i] = total height of rows 0..i-1; offsets[count] is the height of all rows. */
export function rowOffsets(heights: Sizes | undefined, count: number, fallback: number): number[] {
  const offsets = new Array<number>(count + 1);
  offsets[0] = 0;
  for (let r = 0; r < count; r++) offsets[r + 1] = offsets[r]! + (heights?.[r] ?? fallback);
  return offsets;
}

/** The row containing vertical position `y` (0 = top of row 0), clamped to the rows that exist. */
export function rowAt(offsets: number[], y: number): number {
  const count = offsets.length - 1;
  let lo = 0, hi = count - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (offsets[mid]! <= y) lo = mid;
    else hi = mid - 1;
  }
  return Math.max(0, lo);
}
