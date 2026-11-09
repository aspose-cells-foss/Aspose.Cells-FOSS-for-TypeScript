import { CellsException } from "./CellsException";
import { Row } from "./Row";
import { Worksheet } from "./worksheet";

/** Represents a collection of row objects. */
export class RowCollection {
  [key: string]: any;

  private readonly worksheet: Worksheet;

  public constructor(worksheet: Worksheet) {
    this.worksheet = worksheet;
  }

  /** Gets the number of explicitly stored row records. */
  public get Count(): number {
    return this.worksheet.Model.Rows.size;
  }

  /** Gets the element at the specified zero-based index. */
  public get(index: number): Row {
    if (index < 0) {
      throw new CellsException("Row index must be non-negative.");
    }

    return new Row(this.worksheet, index);
  }

  /** Alias of get. */
  public at(index: number): Row {
    return this.get(index);
  }

  public *[Symbol.iterator](): IterableIterator<Row> {
    const indexes = Array.from(this.worksheet.Model.Rows.keys()).sort((left, right) => left - right);
    for (const index of indexes) {
      yield this.get(index);
    }
  }
}
