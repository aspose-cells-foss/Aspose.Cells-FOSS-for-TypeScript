import type { Worksheet } from "../index";

/**
 * Data extracted from a TypeScript Worksheet for the rendering pipeline.
 * This is the bridge between the existing TS data model and the .NET Core model concepts.
 */
export interface ExtractedCell {
  rowIndex: number;
  columnIndex: number;
  value: unknown;
  styleIndex: number;
}

export interface ExtractedWorksheetData {
  name: string;
  index: number;
  firstRow: number;
  lastRow: number;
  firstColumn: number;
  lastColumn: number;
  cells: Map<string, ExtractedCell>; // key = "row,col"
  merges: MergeRegionInfo[];
  columnWidths: Map<number, number>; // colIndex -> width in points-like units
  rowHeights: Map<number, number>;   // rowIndex -> height
  hiddenRows: Set<number>;
  hiddenColumns: Set<number>;
}

export interface MergeRegionInfo {
  firstRow: number;
  firstColumn: number;
  totalRows: number;
  totalColumns: number;
}

/** Extracts worksheet data into a flat structure consumable by rendering. */
export class WorksheetDataExtractor {
  static extract(worksheet: Worksheet): ExtractedWorksheetData {
    const cells = new Map<string, ExtractedCell>();
    let minRow = Infinity, maxRow = -1;
    let minCol = Infinity, maxCol = -1;

    // Extract cells - Cells class has iterator and values() method
    const allCells = (worksheet as any).cells as Iterable<any>;
    for (const cell of allCells) {
      const row = cell.row as number;
      const col = cell.col as number;
      if (row === undefined || col === undefined) continue;

      cells.set(`${row},${col}`, {
        rowIndex: row,
        columnIndex: col,
        value: cell.value ?? null,
        styleIndex: cell.styleIndex ?? 0,
      });

      if (row < minRow) minRow = row;
      if (row > maxRow) maxRow = row;
      if (col < minCol) minCol = minCol;
      if (col > maxCol) maxCol = col;
    }

    if (maxRow < 0) { minRow = 0; minCol = 0; maxRow = -1; maxCol = -1; }

    // Extract merged cells
    const merges: MergeRegionInfo[] = [];
    const mergedCells = (worksheet as any)._mergedCells ?? [];
    for (const merge of mergedCells) {
      // merge is like "A1:B3"
      const parts = merge.split(":");
      if (parts.length === 2) {
        const start = WorksheetDataExtractor.parseRef(parts[0]);
        const end = WorksheetDataExtractor.parseRef(parts[1]);
        merges.push({
          firstRow: start.row,
          firstColumn: start.col,
          totalRows: end.row - start.row + 1,
          totalColumns: end.col - start.col + 1,
        });
      }
    }

    // Also expand merge regions into the used range
    for (const m of merges) {
      if (m.firstRow < minRow) minRow = m.firstRow;
      if (m.firstColumn < minCol) minCol = m.firstColumn;
      const lastR = m.firstRow + m.totalRows - 1;
      const lastC = m.firstColumn + m.totalColumns - 1;
      if (lastR > maxRow) maxRow = lastR;
      if (lastC > maxCol) maxCol = lastC;
    }

    return {
      name: worksheet.name,
      index: worksheet.index,
      firstRow: 0, // Always anchored at A1 like Excel
      lastRow: maxRow,
      firstColumn: 0,
      lastColumn: maxCol,
      cells,
      merges,
      columnWidths: (worksheet as any)._columnWidths ?? new Map(),
      rowHeights: (worksheet as any)._rowHeights ?? new Map(),
      hiddenRows: (worksheet as any)._hiddenRows ?? new Set(),
      hiddenColumns: new Set<number>(),
    };
  }

  private static parseRef(ref: string): { row: number; col: number } {
    const m = ref.match(/^([A-Z]+)(\d+)$/);
    if (!m) return { row: 0, col: 0 };
    let col = 0;
    for (let i = 0; i < m[1].length; i++) {
      col = col * 26 + (m[1].charCodeAt(i) - 64);
    }
    return { row: parseInt(m[2], 10) - 1, col: col - 1 };
  }
}
