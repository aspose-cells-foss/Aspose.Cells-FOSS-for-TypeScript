import type { Worksheet } from "../index";
import type { ImageInfo, ShapeInfo } from "../types";

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

export interface MergeRegionInfo {
  firstRow: number;
  firstColumn: number;
  totalRows: number;
  totalColumns: number;
}

/** Anchored picture usable by the rendering pipeline. */
export interface ExtractedPicture {
  /** Buffer decoded from the worksheet media (PNG/JPEG/etc.). */
  data: Buffer;
  /** Image extension for pdfkit ("png", "jpg", ...). */
  ext: string;
  /** EMU offsets for the twoCellAnchor (914400 EMU/in, /12700 = pt). */
  fromRow: number;
  fromRowOff: number;
  fromCol: number;
  fromColOff: number;
  toRow: number;
  toRowOff: number;
  toCol: number;
  toColOff: number;
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
  pictures: ExtractedPicture[];
  shapes: ShapeInfo[];
  columnWidths: Map<number, number>; // colIndex -> width in Excel chars
  rowHeights: Map<number, number>;   // rowIndex -> height in points
  defaultColumnWidthChars?: number;  // from sheetFormatPr@defaultColWidth, Excel chars
  defaultRowHeightPt?: number;       // from sheetFormatPr@defaultRowHeight, in points
  hiddenRows: Set<number>;
  hiddenColumns: Set<number>;
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
      if (col < minCol) minCol = col;
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

    const shapes = (worksheet as any).shapes as ShapeInfo[] | undefined;
    const pictures: ExtractedPicture[] = [];
    const imagesByName = new Map<string, ImageInfo>();
    for (const img of worksheet.images as ImageInfo[]) imagesByName.set(img.name, img);
    if (Array.isArray(shapes)) {
      for (const shape of shapes) {
        if ((shape as any)?.type !== "picture") continue;
        const imgName = (shape as any).imageName ?? (shape as any).imageEmbedId;
        const img: ImageInfo | undefined = imgName ? imagesByName.get(imgName) : undefined;
        if (!img?.data || img.data.length === 0) continue;
        pictures.push({
          data: img.data,
          ext: (img.ext || "png").toLowerCase() === "jpg" ? "jpeg" : (img.ext || "png").toLowerCase(),
          fromRow: shape.fromRow ?? 0,
          fromRowOff: shape.fromRowOff ?? 0,
          fromCol: shape.fromCol ?? 0,
          fromColOff: shape.fromColOff ?? 0,
          toRow: shape.toRow ?? shape.fromRow ?? 0,
          toRowOff: shape.toRowOff ?? 0,
          toCol: shape.toCol ?? shape.fromCol ?? 0,
          toColOff: shape.toColOff ?? 0,
        });
      }
    }

    // Also expand picture anchors into the used range so PageLayouts include them
    for (const p of pictures) {
      const minR = Math.min(p.fromRow, p.toRow);
      const maxR = Math.max(p.fromRow, p.toRow);
      const minC = Math.min(p.fromCol, p.toCol);
      const maxC = Math.max(p.fromCol, p.toCol);
      if (minR < minRow) minRow = minR;
      if (maxR > maxRow) maxRow = maxR;
      if (minC < minCol) minCol = minC;
      if (maxC > maxCol) maxCol = maxC;
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
      pictures,
      shapes: Array.isArray(shapes) ? shapes.slice() : [],
      columnWidths: (worksheet as any)._columnWidths ?? new Map(),
      rowHeights: (worksheet as any)._rowHeights ?? new Map(),
      defaultColumnWidthChars: worksheet.defaultColumnWidth,
      defaultRowHeightPt: worksheet.defaultRowHeight,
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
