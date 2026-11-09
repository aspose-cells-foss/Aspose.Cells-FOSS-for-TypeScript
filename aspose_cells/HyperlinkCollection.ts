import { CellsException } from "./CellsException";
import { CellAddress } from "./Core/CellAddress";
import { Hyperlink } from "./hyperlink";
import { HyperlinkModel } from "./Core/HyperlinkModel";

// Implemented by serializer module.
/** Represents a collection of worksheet hyperlinks. */
export class HyperlinkCollection {
  [key: string]: any;

  private readonly models: HyperlinkModel[];

  public constructor(models: HyperlinkModel[]) {
    this.models = models;
  }

  public get Count(): number {
    return this.models.length;
  }

  /** Aspose compatibility: capacity mirrors Count because the collection is backed by an array. */
  public get Capacity(): number {
    return this.models.length;
  }

  public set Capacity(_value: number) {
    // Retained for Aspose-compatible API shape.
  }

  public get(index: number): Hyperlink {
    if (index < 0 || index >= this.models.length) {
      throw new CellsException("Hyperlink index was out of range.");
    }

    return new Hyperlink(this.models[index], (model) => this.removeModel(model));
  }

  public at(index: number): Hyperlink {
    return this.get(index);
  }

  public Add(
    firstRow: number,
    firstColumn: number,
    totalRows: number,
    totalColumns: number,
    address: string,
  ): number;
  public Add(cellName: string, totalRows: number, totalColumns: number, address: string): number;
  public Add(
    startCellName: string,
    endCellName: string,
    address: string,
    textToDisplay?: string,
    screenTip?: string,
  ): number;
  public Add(
    startCellNameOrFirstRow: string | number,
    second: number | string,
    third: number | string,
    fourth: string | number | undefined,
    fifth?: string,
    sixth?: string,
  ): number {
    if (typeof startCellNameOrFirstRow === "number") {
      if (
        typeof second !== "number" ||
        typeof third !== "number" ||
        typeof fourth !== "number" ||
        fifth === undefined
      ) {
        throw new CellsException("Hyperlink range arguments are invalid.");
      }

      return this.addRange(startCellNameOrFirstRow, second, third, fourth, fifth);
    }

    const startCellName = startCellNameOrFirstRow;
    if (typeof second === "number") {
      const anchor = CellAddress.Parse(startCellName);
      return this.addRange(anchor.RowIndex, anchor.ColumnIndex, second, third as number, fourth as string);
    }

    const start = CellAddress.Parse(startCellName);
    const end = CellAddress.Parse(second);
    const firstRow = Math.min(start.RowIndex, end.RowIndex);
    const firstColumn = Math.min(start.ColumnIndex, end.ColumnIndex);
    const lastRow = Math.max(start.RowIndex, end.RowIndex);
    const lastColumn = Math.max(start.ColumnIndex, end.ColumnIndex);
    const index = this.addRange(
      firstRow,
      firstColumn,
      lastRow - firstRow + 1,
      lastColumn - firstColumn + 1,
      fourth as string,
    );
    const model = this.models[index];
    model.TextToDisplay = normalizeText(fifth);
    model.ScreenTip = normalizeText(sixth);
    return index;
  }

  public AddRange(
    firstRow: number,
    firstColumn: number,
    totalRows: number,
    totalColumns: number,
    address: string,
  ): number {
    return this.addRange(firstRow, firstColumn, totalRows, totalColumns, address);
  }

  public RemoveAt(index: number): void {
    if (index < 0 || index >= this.models.length) {
      throw new CellsException("Hyperlink index was out of range.");
    }

    this.models.splice(index, 1);
  }

  public Clear(): void {
    this.models.length = 0;
  }

  public *[Symbol.iterator](): IterableIterator<Hyperlink> {
    for (let index = 0; index < this.models.length; index++) {
      yield this.get(index);
    }
  }

  private addRange(
    firstRow: number,
    firstColumn: number,
    totalRows: number,
    totalColumns: number,
    address: string,
  ): number {
    if (firstRow < 0 || firstColumn < 0) {
      throw new CellsException("Hyperlink origin indices must be non-negative.");
    }

    if (totalRows <= 0 || totalColumns <= 0) {
      throw new CellsException("Hyperlink range dimensions must be positive.");
    }

    const model = new HyperlinkModel();
    model.FirstRow = firstRow;
    model.FirstColumn = firstColumn;
    model.TotalRows = totalRows;
    model.TotalColumns = totalColumns;
    assignAddress(model, address);

    for (const existing of this.models) {
      if (overlaps(existing, model)) {
        throw new CellsException("Hyperlink ranges must not overlap.");
      }
    }

    this.models.push(model);
    return this.models.length - 1;
  }

  private removeModel(model: HyperlinkModel): void {
    const index = this.models.indexOf(model);
    if (index >= 0) {
      this.models.splice(index, 1);
    }
  }
}

function normalizeText(value: string | null | undefined): string | null {
  return value === null || value === undefined || value.length === 0 ? null : value;
}

function assignAddress(model: HyperlinkModel, address: string): void {
  if (address === null || address === undefined || address.trim().length === 0) {
    throw new CellsException("Hyperlink address must be non-empty.");
  }

  const normalized = address.trim();
  if (normalized.startsWith("#")) {
    model.Address = null;
    model.SubAddress = normalized.substring(1);
    return;
  }

  if (normalized.indexOf("!") >= 0) {
    model.Address = null;
    model.SubAddress = normalized;
    return;
  }

  model.Address = normalized;
  model.SubAddress = null;
}

function overlaps(left: HyperlinkModel, right: HyperlinkModel): boolean {
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
