/**
 * Worksheet content rectangle clipped to a page.  Floating objects may extend
 * beyond the last used column/row so we add a 72 pt (1") margin in that case.
 */
export interface SheetLayoutBounds {
  ColumnStartPt: number[];
  RowStartPt: number[];
  LastColumn: number;
  LastRow: number;
}

export interface PageRange {
  StartColumn: number;
  EndColumn: number;
  StartRow: number;
  EndRow: number;
  Sheet: SheetLayoutBounds;
}

export class PageContentClip {
  static Resolve(page: PageRange, floatingObjects: boolean): { left: number; top: number; width: number; height: number } {
    const width = page.Sheet.ColumnStartPt[page.EndColumn + 1] - page.Sheet.ColumnStartPt[page.StartColumn];
    const height = page.Sheet.RowStartPt[page.EndRow + 1] - page.Sheet.RowStartPt[page.StartRow];
    let w = width;
    let h = height;
    if (floatingObjects) {
      if (page.EndColumn >= page.Sheet.LastColumn) w += 72;
      if (page.EndRow >= page.Sheet.LastRow) h += 72;
    }
    return { left: 0, top: 0, width: Math.max(w, 0), height: Math.max(h, 0) };
  }
}
