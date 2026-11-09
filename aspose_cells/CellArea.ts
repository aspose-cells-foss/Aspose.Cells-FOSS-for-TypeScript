/**
 * Represents a rectangular cell area.
 * Ported from Aspose.Cells_FOSS.CellArea.
 */

interface ParsedCellAddress {
  readonly RowIndex: number;
  readonly ColumnIndex: number;
}

function parseCellAddress(cellReference: string): ParsedCellAddress {
  if (cellReference === null || cellReference === undefined || cellReference.trim().length === 0) {
    throw new Error("Cell reference must be non-empty.");
  }

  const reference = cellReference.trim();
  let index = 0;
  let column = 0;

  while (index < reference.length && isLetter(reference[index])) {
    const letter = reference[index].toUpperCase();
    if (letter < "A" || letter > "Z") {
      throw new Error(`Cell reference '${cellReference}' is invalid.`);
    }
    column = column * 26 + (letter.charCodeAt(0) - 65 + 1);
    index++;
  }

  if (column === 0 || index === reference.length) {
    throw new Error(`Cell reference '${cellReference}' is invalid.`);
  }

  let row = 0;
  while (index < reference.length && isDigit(reference[index])) {
    row = row * 10 + (reference.charCodeAt(index) - 48);
    index++;
  }

  if (index !== reference.length || row <= 0) {
    throw new Error(`Cell reference '${cellReference}' is invalid.`);
  }

  return { RowIndex: row - 1, ColumnIndex: column - 1 };
}

function isLetter(ch: string): boolean {
  return (ch >= "A" && ch <= "Z") || (ch >= "a" && ch <= "z");
}

function isDigit(ch: string): boolean {
  return ch >= "0" && ch <= "9";
}

export class CellArea {
  public StartRow: number;
  public StartColumn: number;
  public EndRow: number;
  public EndColumn: number;

  public constructor(firstRow: number = 0, firstColumn: number = 0, totalRows: number = 1, totalColumns: number = 1) {
    if (firstRow < 0) {
      throw new RangeError("firstRow");
    }
    if (firstColumn < 0) {
      throw new RangeError("firstColumn");
    }
    if (totalRows <= 0) {
      throw new RangeError("totalRows");
    }
    if (totalColumns <= 0) {
      throw new RangeError("totalColumns");
    }

    this.StartRow = firstRow;
    this.StartColumn = firstColumn;
    this.EndRow = firstRow + totalRows - 1;
    this.EndColumn = firstColumn + totalColumns - 1;
  }

  /** Gets the start row. */
  public get FirstRow(): number {
    return this.StartRow;
  }

  /** Gets the start column. */
  public get FirstColumn(): number {
    return this.StartColumn;
  }

  /** Gets the total number of rows in this area. */
  public get TotalRows(): number {
    return this.EndRow - this.StartRow + 1;
  }

  /** Gets the total number of columns in this area. */
  public get TotalColumns(): number {
    return this.EndColumn - this.StartColumn + 1;
  }

  /** Creates a cell area from zero-based row/column bounds. */
  public static CreateCellArea(startRow: number, startColumn: number, endRow: number, endColumn: number): CellArea;
  /** Creates a cell area from two A1-style cell references. */
  public static CreateCellArea(startCellName: string, endCellName: string): CellArea;
  public static CreateCellArea(
    startRow: number | string,
    startColumn: number | string,
    endRow?: number,
    endColumn?: number,
  ): CellArea {
    if (typeof startRow === "string" || typeof startColumn === "string") {
      if (typeof startRow !== "string" || typeof startColumn !== "string") {
        throw new Error("Both cell references must be strings.");
      }
      const start = parseCellAddress(startRow);
      const end = parseCellAddress(startColumn);
      return CellArea.CreateCellArea(
        Math.min(start.RowIndex, end.RowIndex),
        Math.min(start.ColumnIndex, end.ColumnIndex),
        Math.max(start.RowIndex, end.RowIndex),
        Math.max(start.ColumnIndex, end.ColumnIndex),
      );
    }

    if (endRow === undefined || endColumn === undefined) {
      throw new Error("End row and end column are required.");
    }
    if (startRow < 0) {
      throw new RangeError("startRow");
    }
    if (startColumn < 0) {
      throw new RangeError("startColumn");
    }
    if (endRow < startRow) {
      throw new RangeError("endRow");
    }
    if (endColumn < startColumn) {
      throw new RangeError("endColumn");
    }

    const area = new CellArea();
    area.StartRow = startRow;
    area.StartColumn = startColumn;
    area.EndRow = endRow;
    area.EndColumn = endColumn;
    return area;
  }

  /** Compares two cell areas by their top-left corner, then their bottom-right corner. */
  public CompareTo(obj: CellArea | null | undefined): number {
    if (obj === null || obj === undefined) {
      return 1;
    }

    if (this.StartRow !== obj.StartRow) {
      return compareNumbers(this.StartRow, obj.StartRow);
    }
    if (this.StartColumn !== obj.StartColumn) {
      return compareNumbers(this.StartColumn, obj.StartColumn);
    }
    if (this.EndRow !== obj.EndRow) {
      return compareNumbers(this.EndRow, obj.EndRow);
    }
    return compareNumbers(this.EndColumn, obj.EndColumn);
  }

  /** Determines whether the specified area equals this instance. */
  public Equals(other: CellArea | null | undefined): boolean {
    if (other === null || other === undefined) {
      return false;
    }
    return (
      this.StartRow === other.StartRow &&
      this.StartColumn === other.StartColumn &&
      this.EndRow === other.EndRow &&
      this.EndColumn === other.EndColumn
    );
  }

  /** Lowercase alias of Equals for generic assertion helpers. */
  public equals(other: CellArea | null | undefined): boolean {
    return this.Equals(other);
  }

  public ToString(): string {
    return `R${this.StartRow}C${this.StartColumn}:R${this.EndRow}C${this.EndColumn}`;
  }

  public toString(): string {
    return this.ToString();
  }

  /** Creates a copy of this area. */
  public clone(): CellArea {
    const clone = new CellArea();
    clone.StartRow = this.StartRow;
    clone.StartColumn = this.StartColumn;
    clone.EndRow = this.EndRow;
    clone.EndColumn = this.EndColumn;
    return clone;
  }
}

function compareNumbers(left: number, right: number): number {
  if (left < right) {
    return -1;
  }
  if (left > right) {
    return 1;
  }
  return 0;
}
