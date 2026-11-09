export class MergeRegion {
  public readonly FirstRow: number;
  public readonly FirstColumn: number;
  public readonly TotalRows: number;
  public readonly TotalColumns: number;

  public constructor(firstRow: number, firstColumn: number, totalRows: number, totalColumns: number) {
    this.FirstRow = firstRow;
    this.FirstColumn = firstColumn;
    this.TotalRows = totalRows;
    this.TotalColumns = totalColumns;
  }

  public clone(): MergeRegion {
    return new MergeRegion(this.FirstRow, this.FirstColumn, this.TotalRows, this.TotalColumns);
  }
}
