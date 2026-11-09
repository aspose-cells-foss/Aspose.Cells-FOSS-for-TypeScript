import { Cell } from "./cell";
import { CellAddress } from "./Core/CellAddress";
import { CellArea } from "./CellArea";
import { CellsException } from "./CellsException";
import { ColumnCollection } from "./ColumnCollection";
import { createIndexerProxy } from "./IndexerProxy";
import { MergeRegion } from "./Core/MergeRegion";
import { RowCollection } from "./RowCollection";
import { Style } from "./style";
import { Worksheet } from "./worksheet";

/** Provides access to worksheet cells, rows, columns, and merged ranges. */
export class Cells {
  /** Allows A1-style bracket access such as cells["A1"]. */
  [key: string]: any;

  private readonly worksheet: Worksheet;
  private readonly rows: RowCollection;
  private readonly columns: ColumnCollection;

  public constructor(worksheet: Worksheet) {
    this.worksheet = worksheet;

    const rows = new RowCollection(worksheet);
    this.rows = createIndexerProxy(rows, (key) =>
      typeof key === "number" ? rows.get(key) : undefined,
    );

    const columns = new ColumnCollection(worksheet);
    this.columns = createIndexerProxy(columns, (key) =>
      typeof key === "number" ? columns.get(key) : undefined,
    );
  }

  /** Gets a cell by A1 reference or by zero-based row and column indices. */
  public get(cellName: string): Cell;
  public get(row: number, column: number): Cell;
  public get(rowOrName: number | string, column?: number): Cell {
    if (typeof rowOrName === "string") {
      return this.getByName(rowOrName);
    }

    if (column === undefined) {
      throw new CellsException("A column index is required when addressing a cell by row index.");
    }

    if (rowOrName < 0 || column < 0) {
      throw new CellsException("Row and column indices must be non-negative.");
    }

    return new Cell(this.worksheet, new CellAddress(rowOrName, column));
  }

  /** Gets a cell by A1 reference such as A1 or BC12. */
  public getByName(cellName: string): Cell {
    try {
      return new Cell(this.worksheet, CellAddress.Parse(cellName));
    } catch (error) {
      throw new CellsException(
        error instanceof Error ? error.message : `Cell reference '${cellName}' is invalid.`,
        error,
      );
    }
  }

  /** Gets row-level settings for the worksheet. */
  public get Rows(): RowCollection {
    return this.rows;
  }

  /** Gets or sets the default style of the worksheet cells scope. */
  public get Style(): Style {
    return Style.FromCore(this.worksheet.Workbook.Model.DefaultStyle);
  }

  public set Style(value: Style) {
    if (value === null || value === undefined) {
      throw new Error("value");
    }

    this.worksheet.Workbook.Model.DefaultStyle = value.ToCore();
  }

  /** Gets column-level settings for the worksheet. */
  public get Columns(): ColumnCollection {
    return this.columns;
  }

  /** Gets the current merged-cell regions in worksheet order. */
  public get MergedCells(): readonly CellArea[] {
    return this.worksheet.Model.MergeRegions.map(
      (region) =>
        new CellArea(
          region.FirstRow,
          region.FirstColumn,
          region.TotalRows,
          region.TotalColumns,
        ),
    );
  }

  /** Merges a rectangular cell region using zero-based coordinates. */
  public Merge(firstRow: number, firstColumn: number, totalRows: number, totalColumns: number): void {
    if (firstRow < 0 || firstColumn < 0) {
      throw new CellsException("Merge origin indices must be non-negative.");
    }

    if (totalRows <= 0 || totalColumns <= 0) {
      throw new CellsException("Merge range dimensions must be positive.");
    }

    const candidate = new MergeRegion(firstRow, firstColumn, totalRows, totalColumns);
    for (const existing of this.worksheet.Model.MergeRegions) {
      if (Cells.overlaps(existing, candidate)) {
        throw new CellsException("Merge ranges must not overlap.");
      }
    }

    this.worksheet.Model.MergeRegions.push(candidate);
  }

  private static overlaps(left: MergeRegion, right: MergeRegion): boolean {
    const leftLastRow = left.FirstRow + left.TotalRows - 1;
    const leftLastColumn = left.FirstColumn + left.TotalColumns - 1;
    const rightLastRow = right.FirstRow + right.TotalRows - 1;
    const rightLastColumn = right.FirstColumn + right.TotalColumns - 1;

    return (
      left.FirstRow <= rightLastRow &&
      right.FirstRow <= leftLastRow &&
      left.FirstColumn <= rightLastColumn &&
      right.FirstColumn <= leftLastColumn
    );
  }
}
