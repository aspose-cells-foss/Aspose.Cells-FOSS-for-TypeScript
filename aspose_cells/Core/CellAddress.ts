export class CellAddress {
  public readonly RowIndex: number;
  public readonly ColumnIndex: number;

  public constructor(rowIndex: number, columnIndex: number) {
    if (rowIndex < 0) throw new RangeError("rowIndex");
    if (columnIndex < 0) throw new RangeError("columnIndex");
    this.RowIndex = rowIndex;
    this.ColumnIndex = columnIndex;
  }

  public static Parse(cellReference: string): CellAddress {
    if (cellReference == null || cellReference.trim().length === 0) {
      throw new Error("Cell reference must be non-empty.");
    }

    const reference = cellReference.trim();
    let index = 0;
    let column = 0;

    // Convert the A1 column prefix into a zero-based index using base-26 letters.
    while (index < reference.length && CellAddress.isLetter(reference.charCodeAt(index))) {
      const letter = reference.charCodeAt(index);
      const upper = letter >= 97 && letter <= 122 ? letter - 32 : letter;
      if (upper < 65 || upper > 90) {
        throw new Error(`Cell reference '${cellReference}' is invalid.`);
      }

      column = column * 26 + (upper - 65 + 1);
      index++;
    }

    if (column === 0 || index === reference.length) {
      throw new Error(`Cell reference '${cellReference}' is invalid.`);
    }

    let row = 0;
    while (index < reference.length && CellAddress.isDigit(reference.charCodeAt(index))) {
      row = row * 10 + (reference.charCodeAt(index) - 48);
      index++;
    }

    if (index !== reference.length || row <= 0) {
      throw new Error(`Cell reference '${cellReference}' is invalid.`);
    }

    return new CellAddress(row - 1, column - 1);
  }

  public Equals(other: CellAddress | null): boolean {
    if (other == null) {
      return false;
    }

    return this.RowIndex === other.RowIndex && this.ColumnIndex === other.ColumnIndex;
  }

  public hashKey(): number {
    return (Math.imul(this.RowIndex, 397) ^ this.ColumnIndex) | 0;
  }

  public ToString(): string {
    return CellAddress.ColumnIndexToName(this.ColumnIndex) + (this.RowIndex + 1).toString();
  }

  public clone(): CellAddress {
    return new CellAddress(this.RowIndex, this.ColumnIndex);
  }

  private static ColumnIndexToName(columnIndex: number): string {
    let index = columnIndex + 1;
    const characters: string[] = [];
    // Convert the zero-based column index back to Excel's repeated-letter form.
    while (index > 0) {
      index--;
      characters.push(String.fromCharCode(65 + (index % 26)));
      index = Math.floor(index / 26);
    }

    return characters.reverse().join("");
  }

  private static isLetter(code: number): boolean {
    return (code >= 65 && code <= 90) || (code >= 97 && code <= 122);
  }

  private static isDigit(code: number): boolean {
    return code >= 48 && code <= 57;
  }
}
