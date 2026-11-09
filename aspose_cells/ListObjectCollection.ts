import { CellsException } from "./CellsException";
import { CellAddress } from "./Core/CellAddress";
import { ListColumnModel } from "./Core/ListColumnModel";
import { ListObject } from "./ListObject";
import { ListObjectModel } from "./Core/ListObjectModel";

// Implemented by serializer module.
/** Represents the collection of Excel tables on a worksheet. */
export class ListObjectCollection {
  [key: string]: any;

  private readonly models: ListObjectModel[];

  public constructor(models: ListObjectModel[]) {
    this.models = models;
  }

  public get Count(): number {
    return this.models.length;
  }

  public get(index: number): ListObject {
    if (index < 0 || index >= this.models.length) {
      throw new CellsException(`Table index ${index} is out of range.`);
    }

    return new ListObject(this.models[index], this.models);
  }

  public at(index: number): ListObject {
    return this.get(index);
  }

  public getByName(name: string): ListObject {
    for (const model of this.models) {
      if (model.DisplayName.toLowerCase() === name.toLowerCase()) {
        return new ListObject(model, this.models);
      }
    }

    throw new CellsException(`No table with display name '${name}' was found.`);
  }

  public RemoveAt(index: number): void {
    if (index < 0 || index >= this.models.length) {
      throw new CellsException(`Table index ${index} is out of range.`);
    }

    this.models.splice(index, 1);
  }

  public Add(startRow: number, startColumn: number, endRow: number, endColumn: number, hasHeaders: boolean): number;
  public Add(startCellName: string, endCellName: string, hasHeaders: boolean): number;
  public Add(
    startRowOrName: number | string,
    startColumnOrEndName: number | string,
    endRowOrHasHeaders: number | boolean,
    endColumnOrNothing?: number | boolean,
    hasHeaders?: boolean,
  ): number {
    let startRow: number;
    let startColumn: number;
    let endRow: number;
    let endColumn: number;
    let headers: boolean;

    if (typeof startRowOrName === "string") {
      const start = CellAddress.Parse(startRowOrName);
      const end = CellAddress.Parse(startColumnOrEndName as string);
      startRow = start.RowIndex;
      startColumn = start.ColumnIndex;
      endRow = end.RowIndex;
      endColumn = end.ColumnIndex;
      headers = typeof endRowOrHasHeaders === "boolean" ? endRowOrHasHeaders : true;
    } else {
      startRow = startRowOrName;
      startColumn = startColumnOrEndName as number;
      endRow = endRowOrHasHeaders as number;
      endColumn = endColumnOrNothing as number;
      headers = hasHeaders ?? true;
    }

    if (startRow < 0 || startColumn < 0) {
      throw new CellsException("Table origin indices must be non-negative.");
    }

    if (endRow < startRow || endColumn < startColumn) {
      throw new CellsException("Table range end must not precede the start.");
    }

    const model = new ListObjectModel();
    model.DisplayName = this.nextDisplayName();
    model.Name = model.DisplayName;
    model.StartRow = startRow;
    model.StartColumn = startColumn;
    model.EndRow = endRow;
    model.EndColumn = endColumn;
    model.ShowHeaderRow = headers;

    for (let column = startColumn; column <= endColumn; column++) {
      model.Columns.push(new ListColumnModel(column - startColumn + 1, `Column${column - startColumn + 1}`));
    }

    this.models.push(model);
    return this.models.length - 1;
  }

  public *[Symbol.iterator](): IterableIterator<ListObject> {
    for (let index = 0; index < this.models.length; index++) {
      yield this.get(index);
    }
  }

  private nextDisplayName(): string {
    let suffix = this.models.length + 1;
    while (this.models.some((model) => model.DisplayName.toLowerCase() === `table${suffix}`.toLowerCase())) {
      suffix++;
    }

    return `Table${suffix}`;
  }
}
