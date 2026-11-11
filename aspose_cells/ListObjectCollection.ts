import { CellsException } from "./CellsException";
import { CellAddress } from "./Core/CellAddress";
import { findCellRecord } from "./CellStore";
import { ListColumnModel } from "./Core/ListColumnModel";
import { ListObject } from "./ListObject";
import { ListObjectModel } from "./Core/ListObjectModel";
import { WorksheetModel } from "./Core/WorksheetModel";

// Implemented by serializer module.
/** Represents the collection of Excel tables on a worksheet. */
export class ListObjectCollection {
  [key: string]: any;

  private readonly models: ListObjectModel[];

  public constructor(models: ListObjectModel[], private readonly worksheetModel?: WorksheetModel) {
    this.models = models;
  }

  public get Count(): number {
    return this.models.length;
  }

  public get(index: number): ListObject {
    if (index < 0 || index >= this.models.length) {
      throw new CellsException(`Table index ${index} is out of range.`);
    }

    return new ListObject(this.models[index], this.models, this.worksheetModel);
  }

  public at(index: number): ListObject {
    return this.get(index);
  }

  public getByName(name: string): ListObject {
    for (const model of this.models) {
      if (model.DisplayName.toLowerCase() === name.toLowerCase()) {
        return new ListObject(model, this.models, this.worksheetModel);
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

    ListObjectCollection.validateRange(startRow, startColumn, endRow, endColumn);
    ListObjectCollection.validateNoOverlap(this.models, startRow, startColumn, endRow, endColumn, -1);

    const tableNumber = this.models.length + 1;
    const displayName = `Table${tableNumber}`;
    const model = new ListObjectModel();
    model.DisplayName = displayName;
    model.Name = displayName;
    model.StartRow = startRow;
    model.StartColumn = startColumn;
    model.EndRow = endRow;
    model.EndColumn = endColumn;
    model.ShowHeaderRow = headers;
    model.HasAutoFilter = headers;

    const columnCount = endColumn - startColumn + 1;
    for (let offset = 0; offset < columnCount; offset++) {
      const columnName = ListObjectCollection.resolveColumnName(
        this.worksheetModel,
        startRow,
        startColumn + offset,
        headers,
        offset + 1,
      );
      model.Columns.push(new ListColumnModel(offset + 1, columnName));
    }

    ListObjectCollection.validateUniqueDisplayName(this.models, model.DisplayName, -1);
    this.models.push(model);
    return this.models.length - 1;
  }

  public *[Symbol.iterator](): IterableIterator<ListObject> {
    for (let index = 0; index < this.models.length; index++) {
      yield this.get(index);
    }
  }

  public static validateRange(startRow: number, startColumn: number, endRow: number, endColumn: number): void {
    if (startRow < 0) {
      throw new CellsException("Table startRow must be non-negative.");
    }

    if (startColumn < 0) {
      throw new CellsException("Table startColumn must be non-negative.");
    }

    if (endRow < startRow) {
      throw new CellsException("Table endRow must be greater than or equal to startRow.");
    }

    if (endColumn < startColumn) {
      throw new CellsException("Table endColumn must be greater than or equal to startColumn.");
    }
  }

  public static validateDisplayName(displayName: string): void {
    if (displayName === null || displayName === undefined || displayName.length === 0) {
      throw new CellsException("Table DisplayName must be non-empty.");
    }

    const first = displayName.charAt(0);
    const isLetter = (first >= "A" && first <= "Z") || (first >= "a" && first <= "z");
    if (!isLetter && first !== "_" && first !== "\\") {
      throw new CellsException(`Table DisplayName '${displayName}' must start with a letter or underscore.`);
    }
  }

  public static validateUniqueDisplayName(
    existing: readonly ListObjectModel[],
    displayName: string,
    skipIndex: number,
  ): void {
    for (let index = 0; index < existing.length; index++) {
      if (index === skipIndex) {
        continue;
      }

      if (existing[index].DisplayName.toLowerCase() === displayName.toLowerCase()) {
        throw new CellsException(`A table with display name '${displayName}' already exists in this worksheet.`);
      }
    }
  }

  public static validateNoOverlap(
    existing: readonly ListObjectModel[],
    startRow: number,
    startColumn: number,
    endRow: number,
    endColumn: number,
    skipIndex: number,
  ): void {
    for (let index = 0; index < existing.length; index++) {
      if (index === skipIndex) {
        continue;
      }

      const other = existing[index];
      if (
        ListObjectCollection.rangesOverlap(
          startRow,
          startColumn,
          endRow,
          endColumn,
          other.StartRow,
          other.StartColumn,
          other.EndRow,
          other.EndColumn,
        )
      ) {
        throw new CellsException(`The table range overlaps with existing table '${other.DisplayName}'.`);
      }
    }
  }

  public static rebuildColumns(model: ListObjectModel, worksheetModel?: WorksheetModel): void {
    model.Columns.length = 0;
    const columnCount = model.EndColumn - model.StartColumn + 1;
    for (let offset = 0; offset < columnCount; offset++) {
      const columnName = ListObjectCollection.resolveColumnName(
        worksheetModel,
        model.StartRow,
        model.StartColumn + offset,
        model.ShowHeaderRow,
        offset + 1,
      );
      model.Columns.push(new ListColumnModel(offset + 1, columnName));
    }
  }

  private static rangesOverlap(
    r1s: number,
    c1s: number,
    r1e: number,
    c1e: number,
    r2s: number,
    c2s: number,
    r2e: number,
    c2e: number,
  ): boolean {
    return r1s <= r2e && r2s <= r1e && c1s <= c2e && c2s <= c1e;
  }

  private static resolveColumnName(
    worksheetModel: WorksheetModel | undefined,
    headerRow: number,
    columnIndex: number,
    hasHeaders: boolean,
    defaultColumnNumber: number,
  ): string {
    if (hasHeaders && worksheetModel !== undefined) {
      const record = findCellRecord(worksheetModel, headerRow, columnIndex);
      if (record !== undefined && typeof record.Value === "string" && record.Value.length > 0) {
        return record.Value;
      }
    }

    return `Column${defaultColumnNumber}`;
  }
}
