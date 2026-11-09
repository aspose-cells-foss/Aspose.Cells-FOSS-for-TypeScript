import { CellsException } from "./CellsException";
import { ListColumnModel } from "./Core/ListColumnModel";
import { ListObjectModel } from "./Core/ListObjectModel";
import { TableStyleType, TotalsCalculation } from "./enums";

// Implemented by serializer module.
/** Represents a column of an Excel table. */
export class ListColumn {
  public constructor(private readonly model: ListColumnModel) {}

  public get Model(): ListColumnModel {
    return this.model;
  }

  public get Id(): number {
    return this.model.Id;
  }

  public get Name(): string {
    return this.model.Name;
  }

  public set Name(value: string) {
    this.model.Name = value ?? "";
  }

  public get TotalsCalculation(): TotalsCalculation {
    return ListColumn.totalsCalculationFromString(this.model.TotalsRowFunction);
  }

  public set TotalsCalculation(value: TotalsCalculation) {
    this.model.TotalsRowFunction = ListColumn.totalsCalculationToString(value);
  }

  public get TotalsRowFunction(): string {
    return this.model.TotalsRowFunction;
  }

  public set TotalsRowFunction(value: string) {
    this.model.TotalsRowFunction = value ?? "";
  }

  public get TotalsRowLabel(): string {
    return this.model.TotalsRowLabel;
  }

  public set TotalsRowLabel(value: string) {
    this.model.TotalsRowLabel = value ?? "";
  }

  public get TotalsRowFormula(): string {
    return this.model.TotalsRowFormula;
  }

  public set TotalsRowFormula(value: string) {
    this.model.TotalsRowFormula = value ?? "";
  }

  private static totalsCalculationFromString(value: string): TotalsCalculation {
    switch (value) {
      case "sum":
        return TotalsCalculation.Sum;
      case "count":
        return TotalsCalculation.Count;
      case "average":
        return TotalsCalculation.Average;
      case "max":
        return TotalsCalculation.Max;
      case "min":
        return TotalsCalculation.Min;
      case "stdDev":
        return TotalsCalculation.StdDev;
      case "var":
        return TotalsCalculation.Var;
      case "countNums":
        return TotalsCalculation.CountNums;
      case "custom":
        return TotalsCalculation.Custom;
      default:
        return TotalsCalculation.None;
    }
  }

  private static totalsCalculationToString(value: TotalsCalculation): string {
    switch (value) {
      case TotalsCalculation.Sum:
        return "sum";
      case TotalsCalculation.Count:
        return "count";
      case TotalsCalculation.Average:
        return "average";
      case TotalsCalculation.Max:
        return "max";
      case TotalsCalculation.Min:
        return "min";
      case TotalsCalculation.StdDev:
        return "stdDev";
      case TotalsCalculation.Var:
        return "var";
      case TotalsCalculation.CountNums:
        return "countNums";
      case TotalsCalculation.Custom:
        return "custom";
      default:
        return "none";
    }
  }
}

// Implemented by serializer module.
/** Represents a collection of table columns. */
export class ListColumnCollection {
  private readonly models: ListColumnModel[];

  public constructor(models: ListColumnModel[]) {
    this.models = models;
  }

  public get Count(): number {
    return this.models.length;
  }

  public get(index: number): ListColumn {
    if (index < 0 || index >= this.models.length) {
      throw new CellsException(`List column index ${index} is out of range.`);
    }

    return new ListColumn(this.models[index]);
  }

  public at(index: number): ListColumn {
    return this.get(index);
  }

  public *[Symbol.iterator](): IterableIterator<ListColumn> {
    for (let index = 0; index < this.models.length; index++) {
      yield this.get(index);
    }
  }
}

// Implemented by serializer module.
/** Represents an Excel table. */
export class ListObject {
  public readonly ListColumns: ListColumnCollection;
  private readonly model: ListObjectModel;

  public constructor(model: ListObjectModel, private readonly owner?: ListObjectModel[]) {
    this.model = model;
    this.ListColumns = new ListColumnCollection(model.Columns);
  }

  public get Model(): ListObjectModel {
    return this.model;
  }

  public get DisplayName(): string {
    return this.model.DisplayName;
  }

  public set DisplayName(value: string) {
    this.model.DisplayName = value ?? "";
  }

  public get Name(): string {
    return this.model.Name;
  }

  public set Name(value: string) {
    this.model.Name = value ?? "";
  }

  public get Comment(): string {
    return this.model.Comment;
  }

  public set Comment(value: string) {
    this.model.Comment = value ?? "";
  }

  public get StartRow(): number {
    return this.model.StartRow;
  }

  public get StartColumn(): number {
    return this.model.StartColumn;
  }

  public get EndRow(): number {
    return this.model.EndRow;
  }

  public get EndColumn(): number {
    return this.model.EndColumn;
  }

  public get ShowHeaderRow(): boolean {
    return this.model.ShowHeaderRow;
  }

  public set ShowHeaderRow(value: boolean) {
    this.model.ShowHeaderRow = value;
  }

  public get ShowTotals(): boolean {
    return this.model.ShowTotals;
  }

  public set ShowTotals(value: boolean) {
    this.model.ShowTotals = value;
  }

  public get TableStyleType(): TableStyleType {
    const name = this.model.TableStyleName;
    if (name === null || name.length === 0) {
      return TableStyleType.None;
    }

    const parsed = (TableStyleType as unknown as Record<string, number>)[name];
    return typeof parsed === "number" && parsed !== TableStyleType.None && parsed !== TableStyleType.Custom
      ? parsed
      : TableStyleType.Custom;
  }

  public set TableStyleType(value: TableStyleType) {
    if (value === TableStyleType.Custom) {
      return;
    }

    this.model.TableStyleName =
      value === TableStyleType.None ? "" : String((TableStyleType as unknown as Record<number, string>)[value] ?? "");
  }

  public get TableStyleName(): string {
    return this.model.TableStyleName;
  }

  public set TableStyleName(value: string) {
    this.model.TableStyleName = value ?? "";
  }

  public get ShowTableStyleFirstColumn(): boolean {
    return this.model.ShowFirstColumn;
  }

  public set ShowTableStyleFirstColumn(value: boolean) {
    this.model.ShowFirstColumn = value;
  }

  public get ShowTableStyleLastColumn(): boolean {
    return this.model.ShowLastColumn;
  }

  public set ShowTableStyleLastColumn(value: boolean) {
    this.model.ShowLastColumn = value;
  }

  public get ShowTableStyleRowStripes(): boolean {
    return this.model.ShowRowStripes;
  }

  public set ShowTableStyleRowStripes(value: boolean) {
    this.model.ShowRowStripes = value;
  }

  public get ShowTableStyleColumnStripes(): boolean {
    return this.model.ShowColumnStripes;
  }

  public set ShowTableStyleColumnStripes(value: boolean) {
    this.model.ShowColumnStripes = value;
  }

  /** Enables the AutoFilter for this table. */
  public ShowAutoFilter(): void {
    this.model.HasAutoFilter = true;
  }

  /** Disables the AutoFilter for this table. */
  public RemoveAutoFilter(): void {
    this.model.HasAutoFilter = false;
  }

  /** Resizes the table to the supplied range. */
  public Resize(startRow: number, startColumn: number, endRow: number, endColumn: number, hasHeaders: boolean): void {
    if (startRow < 0 || startColumn < 0) {
      throw new CellsException("Table origin indices must be non-negative.");
    }

    if (endRow < startRow || endColumn < startColumn) {
      throw new CellsException("Table range end must not precede the start.");
    }

    this.model.StartRow = startRow;
    this.model.StartColumn = startColumn;
    this.model.EndRow = endRow;
    this.model.EndColumn = endColumn;
    this.model.ShowHeaderRow = hasHeaders;
  }

  /** Converts the table to a normal cell range. */
  public ConvertToRange(): void {
    if (this.owner === undefined) {
      return;
    }

    const index = this.owner.indexOf(this.model);
    if (index >= 0) {
      this.owner.splice(index, 1);
    }
  }
}
