import { CellsException } from "./CellsException";
import { Column } from "./Column";
import { Worksheet } from "./worksheet";

/** Represents a collection of column objects. */
export class ColumnCollection {
  [key: string]: any;

  private readonly worksheet: Worksheet;

  public constructor(worksheet: Worksheet) {
    this.worksheet = worksheet;
  }

  /** Gets the number of stored column ranges. */
  public get Count(): number {
    return this.worksheet.Model.Columns.length;
  }

  /** Gets the element at the specified zero-based index. */
  public get(index: number): Column {
    if (index < 0) {
      throw new CellsException("Column index must be non-negative.");
    }

    return new Column(this.worksheet, index);
  }

  /** Alias of get. */
  public at(index: number): Column {
    return this.get(index);
  }

  public *[Symbol.iterator](): IterableIterator<Column> {
    const indexes = this.worksheet.Model.Columns
      .map((model) => model.MinColumnIndex)
      .sort((left, right) => left - right);
    for (const index of indexes) {
      yield this.get(index);
    }
  }
}
