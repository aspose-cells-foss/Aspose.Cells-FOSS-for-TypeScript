/** A single physical PDF page definition. */
export class PageLayout {
  sheetIndex = 0;
  sheet!: any; // SheetLayout reference

  startRow = 0;
  endRow = 0;
  startColumn = 0;
  endColumn = 0;

  scaleFactor = 1;

  pageWidthPt = 0;
  pageHeightPt = 0;

  /** Top-left of printable content area (margins), in points. */
  contentOriginXPt = 0;
  contentOriginYPt = 0;

  /** 1-based page index and total, for header/footer substitution. */
  pageNumber = 0;
  totalPages = 0;
}
