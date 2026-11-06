import { ExtractedWorksheetData, MergeRegionInfo } from "./WorksheetDataExtractor";
import { RenderUnits } from "./RenderUnits";

/**
 * Resolved geometry for one worksheet: per-column widths and per-row heights in points, their
 * cumulative offsets, the used cell range, and a merge lookup.
 */
export class SheetLayout {
  static readonly HorizontalPaddingPt = 1.5;
  static readonly VerticalPaddingPt = 1.0;

  private static readonly DefaultColumnWidthChars = 8.43;
  private static readonly DefaultRowHeightPt = 15;
  /** Default column width as Excel chars (the .NET-side DefaultColumnWidthChars). */
  private static readonly DefaultColumnWidthCharsValue = 8.43;

  sheet!: ExtractedWorksheetData;
  firstRow = 0;
  firstColumn = 0;
  lastRow = -1;
  lastColumn = -1;

  columnWidthPt: number[] = [];
  rowHeightPt: number[] = [];
  columnHidden: boolean[] = [];
  rowHidden: boolean[] = [];

  /** Cumulative left edge offsets (length lastColumn+2). */
  columnStartPt: number[] = [0];
  /** Cumulative top edge offsets (length lastRow+2). */
  rowStartPt: number[] = [0];

  // -- Page-local coordinate helpers (the .NET ContentX / ContentY pattern). --
  // Every Draw method MUST go through these so even pages that break at a non-zero
  // start column/row get coordinates relative to the page's own origin, not the
  // sheet's origin. This is the canonical fix for the "even pages are shifted right" bug.

  /** x coordinate of column `col` in a page that starts at `startCol`. */
  contentX(startCol: number, col: number): number {
    return this.columnStartPt[col] - this.columnStartPt[startCol];
  }
  /** y coordinate of row `row` in a page that starts at `startRow`. */
  contentY(startRow: number, row: number): number {
    return this.rowStartPt[row] - this.rowStartPt[startRow];
  }

  private mergeTopLeft = new Map<string, MergeRegionInfo>();
  private mergeCovered = new Set<string>();

  get isEmpty(): boolean {
    return this.lastRow < this.firstRow || this.lastColumn < this.firstColumn;
  }

  static build(data: ExtractedWorksheetData): SheetLayout {
    const layout = new SheetLayout();
    layout.sheet = data;
    layout.computeUsedRange(data);
    layout.buildMergeLookup(data);
    layout.computeColumns(data);
    layout.computeRows(data);
    layout.computeCumulativeOffsets();
    return layout;
  }

  private computeUsedRange(data: ExtractedWorksheetData): void {
    let minRow = Infinity, minCol = Infinity, maxRow = -1, maxCol = -1;

    for (const cell of data.cells.values()) {
      if (cell.value === null || cell.value === undefined) continue;
      if (cell.rowIndex < minRow) minRow = cell.rowIndex;
      if (cell.columnIndex < minCol) minCol = cell.columnIndex;
      if (cell.rowIndex > maxRow) maxRow = cell.rowIndex;
      if (cell.columnIndex > maxCol) maxCol = cell.columnIndex;
    }

    for (const merge of data.merges) {
      if (merge.firstRow < minRow) minRow = merge.firstRow;
      if (merge.firstColumn < minCol) minCol = merge.firstColumn;
      const lastR = merge.firstRow + Math.max(1, merge.totalRows) - 1;
      const lastC = merge.firstColumn + Math.max(1, merge.totalColumns) - 1;
      if (lastR > maxRow) maxRow = lastR;
      if (lastC > maxCol) maxCol = lastC;
    }

    // Anchored pictures live outside the cell grid — expand the used range to include
    // their twoCellAnchor rectangles so Paginator lays out enough columns/rows for them.
    for (const p of data.pictures) {
      const minR = Math.min(p.fromRow, p.toRow);
      const maxR = Math.max(p.fromRow, p.toRow);
      const minC = Math.min(p.fromCol, p.toCol);
      const maxC = Math.max(p.fromCol, p.toCol);
      if (minR < minRow) minRow = minR;
      if (maxR > maxRow) maxRow = maxR;
      if (minC < minCol) minCol = minC;
      if (maxC > maxCol) maxCol = maxC;
    }

    if (maxRow < 0 || maxCol < 0) {
      this.firstRow = 0; this.firstColumn = 0;
      this.lastRow = -1; this.lastColumn = -1;
      return;
    }

    this.firstRow = 0;
    this.firstColumn = 0;
    this.lastRow = maxRow;
    this.lastColumn = maxCol;
  }

  private buildMergeLookup(data: ExtractedWorksheetData): void {
    for (const merge of data.merges) {
      const rows = Math.max(1, merge.totalRows);
      const cols = Math.max(1, merge.totalColumns);
      this.mergeTopLeft.set(SheetLayout.key(merge.firstRow, merge.firstColumn), merge);

      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          if (r === 0 && c === 0) continue;
          this.mergeCovered.add(SheetLayout.key(merge.firstRow + r, merge.firstColumn + c));
        }
      }
    }
  }

  private computeColumns(data: ExtractedWorksheetData): void {
    if (this.isEmpty) {
      this.columnWidthPt = [];
      this.columnHidden = [];
      return;
    }

    const count = this.lastColumn + 1;
    this.columnWidthPt = new Array(count);
    this.columnHidden = new Array(count).fill(false);

    // worksheet._columnWidths are Excel character widths (both OOXML and HTML loaders store them
    // that way now). Convert each to pt via the ECMA-376 formula, matching .NET RenderUnits.
    // Sheet-level defaultColWidth takes priority over the fabricated 8.43 fallback.
    // See .NET SheetLayout.cs lines 194-198.
    const defaultPt = data.defaultColumnWidthChars !== undefined && data.defaultColumnWidthChars > 0
      ? RenderUnits.columnWidthCharsToPoints(data.defaultColumnWidthChars, 7.5)
      : RenderUnits.defaultColumnWidthCharsToPoints(SheetLayout.DefaultColumnWidthChars, 7.5);

    for (let c = 0; c < count; c++) {
      this.columnWidthPt[c] = defaultPt;
    }

    for (const [colIdx, widthChars] of data.columnWidths) {
      const c = Math.max(0, Math.min(colIdx, this.lastColumn));
      if (c < 0 || c >= count) continue;
      this.columnWidthPt[c] = RenderUnits.columnWidthCharsToPoints(widthChars, 7.5);
    }
  }

  private computeRows(data: ExtractedWorksheetData): void {
    if (this.isEmpty) {
      this.rowHeightPt = [];
      this.rowHidden = [];
      return;
    }

    const count = this.lastRow + 1;
    this.rowHeightPt = new Array(count);
    this.rowHidden = new Array(count).fill(false);

    for (let r = 0; r < count; r++) {
      this.rowHeightPt[r] = SheetLayout.DefaultRowHeightPt;
    }

    for (const [rowIdx, height] of data.rowHeights) {
      const r = Math.max(0, Math.min(rowIdx, this.lastRow));
      if (r < 0 || r >= count) continue;
      this.rowHeightPt[r] = height;
    }

    for (const hiddenRow of data.hiddenRows) {
      if (hiddenRow >= 0 && hiddenRow <= this.lastRow) {
        this.rowHidden[hiddenRow] = true;
        this.rowHeightPt[hiddenRow] = 0;
      }
    }
  }

  private computeCumulativeOffsets(): void {
    if (this.isEmpty) {
      this.columnStartPt = [0];
      this.rowStartPt = [0];
      return;
    }

    this.columnStartPt = new Array(this.lastColumn + 2).fill(0);
    for (let c = 0; c <= this.lastColumn; c++) {
      this.columnStartPt[c + 1] = this.columnStartPt[c] + this.columnWidthPt[c];
    }

    this.rowStartPt = new Array(this.lastRow + 2).fill(0);
    for (let r = 0; r <= this.lastRow; r++) {
      this.rowStartPt[r + 1] = this.rowStartPt[r] + this.rowHeightPt[r];
    }
  }

  isMergeCovered(row: number, col: number): boolean {
    return this.mergeCovered.has(SheetLayout.key(row, col));
  }

  tryGetMergeOrigin(row: number, col: number): MergeRegionInfo | undefined {
    return this.mergeTopLeft.get(SheetLayout.key(row, col));
  }

  private static key(row: number, col: number): string {
    return `${row},${col}`;
  }
}
