import { SheetLayout } from "./SheetLayout";
import { MergeRegion } from "./types";

/** Resolves each print area from a worksheet print area string. */
export class PrintAreaRanges {
  static read(printArea: string | undefined): MergeRegion[] {
    if (!printArea || printArea.trim() === "") return [];
    const ranges: MergeRegion[] = [];
    const normalized = printArea.trim();
    for (const part of normalized.split(",")) {
      const ends = part.replace(/\$/g, "").split(":");
      const first = PrintAreaRanges.parseCellAddress(ends[0]);
      const last = PrintAreaRanges.parseCellAddress(ends[ends.length - 1]);
      ranges.push(new MergeRegion(
        first.row, first.col,
        last.row - first.row + 1,
        last.col - first.col + 1,
      ));
    }
    return ranges;
  }

  private static parseCellAddress(ref: string): { row: number; col: number } {
    const m = ref.match(/^([A-Z]+)(\d+)$/);
    if (!m) return { row: 0, col: 0 };
    let col = 0;
    for (let i = 0; i < m[1].length; i++) {
      col = col * 26 + (m[1].charCodeAt(i) - 64);
    }
    return { row: parseInt(m[2], 10) - 1, col: col - 1 };
  }
}
