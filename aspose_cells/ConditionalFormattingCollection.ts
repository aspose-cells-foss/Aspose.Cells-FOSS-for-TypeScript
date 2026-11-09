import { CellsException } from "./CellsException";
import { CellArea } from "./CellArea";
import { FormatConditionCollection, getNextPriority } from "./FormatConditionCollection";
import { ConditionalFormattingModel } from "./Core/ConditionalFormattingModel";

/** Represents a collection of conditional formatting objects. */
export class ConditionalFormattingCollection {
  private readonly collections: ConditionalFormattingModel[];

  public constructor(collections: ConditionalFormattingModel[]) {
    this.collections = collections;
  }

  public get Count(): number {
    return this.collections.length;
  }

  public get(index: number): FormatConditionCollection {
    if (index < 0 || index >= this.collections.length) {
      throw new CellsException("Conditional formatting index was out of range.");
    }

    return new FormatConditionCollection(this.collections, this.collections[index]);
  }

  public at(index: number): FormatConditionCollection {
    return this.get(index);
  }

  public *[Symbol.iterator](): IterableIterator<FormatConditionCollection> {
    for (let index = 0; index < this.collections.length; index++) {
      yield this.get(index);
    }
  }

  public Add(): number {
    this.collections.push(new ConditionalFormattingModel());
    return this.collections.length - 1;
  }

  public RemoveAt(index: number): void {
    if (index < 0 || index >= this.collections.length) {
      throw new CellsException("Conditional formatting index was out of range.");
    }

    this.collections.splice(index, 1);
  }

  public RemoveArea(startRow: number, startColumn: number, totalRows: number, totalColumns: number): void {
    const area = CellArea.CreateCellArea(
      startRow,
      startColumn,
      startRow + totalRows - 1,
      startColumn + totalColumns - 1,
    );

    for (let index = this.collections.length - 1; index >= 0; index--) {
      const collection = new FormatConditionCollection(this.collections, this.collections[index]);
      collection.RemoveAreaByArea(area);
    }
  }

  /** Returns the next unused conditional formatting priority. */
  public static getNextPriority(collections: readonly ConditionalFormattingModel[]): number {
    return getNextPriority(collections);
  }
}
