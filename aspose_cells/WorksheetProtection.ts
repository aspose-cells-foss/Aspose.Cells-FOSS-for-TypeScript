import { WorksheetProtectionModel } from "./Core/WorksheetProtectionModel";

/** Represents worksheet protection. */
export class WorksheetProtection {
  private readonly model: WorksheetProtectionModel;

  public constructor(model: WorksheetProtectionModel) {
    this.model = model;
  }

  public get IsProtected(): boolean {
    return this.model.IsProtected;
  }

  public set IsProtected(value: boolean) {
    this.model.IsProtected = value;
  }

  public get Objects(): boolean {
    return this.model.Objects;
  }

  public set Objects(value: boolean) {
    this.model.Objects = value;
    this.markProtectedWhenEnabled(value);
  }

  public get Scenarios(): boolean {
    return this.model.Scenarios;
  }

  public set Scenarios(value: boolean) {
    this.model.Scenarios = value;
    this.markProtectedWhenEnabled(value);
  }

  public get FormatCells(): boolean {
    return this.model.FormatCells;
  }

  public set FormatCells(value: boolean) {
    this.model.FormatCells = value;
    this.markProtectedWhenEnabled(value);
  }

  public get FormatColumns(): boolean {
    return this.model.FormatColumns;
  }

  public set FormatColumns(value: boolean) {
    this.model.FormatColumns = value;
    this.markProtectedWhenEnabled(value);
  }

  public get FormatRows(): boolean {
    return this.model.FormatRows;
  }

  public set FormatRows(value: boolean) {
    this.model.FormatRows = value;
    this.markProtectedWhenEnabled(value);
  }

  public get InsertColumns(): boolean {
    return this.model.InsertColumns;
  }

  public set InsertColumns(value: boolean) {
    this.model.InsertColumns = value;
    this.markProtectedWhenEnabled(value);
  }

  public get InsertRows(): boolean {
    return this.model.InsertRows;
  }

  public set InsertRows(value: boolean) {
    this.model.InsertRows = value;
    this.markProtectedWhenEnabled(value);
  }

  public get InsertHyperlinks(): boolean {
    return this.model.InsertHyperlinks;
  }

  public set InsertHyperlinks(value: boolean) {
    this.model.InsertHyperlinks = value;
    this.markProtectedWhenEnabled(value);
  }

  public get DeleteColumns(): boolean {
    return this.model.DeleteColumns;
  }

  public set DeleteColumns(value: boolean) {
    this.model.DeleteColumns = value;
    this.markProtectedWhenEnabled(value);
  }

  public get DeleteRows(): boolean {
    return this.model.DeleteRows;
  }

  public set DeleteRows(value: boolean) {
    this.model.DeleteRows = value;
    this.markProtectedWhenEnabled(value);
  }

  public get SelectLockedCells(): boolean {
    return this.model.SelectLockedCells;
  }

  public set SelectLockedCells(value: boolean) {
    this.model.SelectLockedCells = value;
    this.markProtectedWhenEnabled(value);
  }

  public get Sort(): boolean {
    return this.model.Sort;
  }

  public set Sort(value: boolean) {
    this.model.Sort = value;
    this.markProtectedWhenEnabled(value);
  }

  public get AutoFilter(): boolean {
    return this.model.AutoFilter;
  }

  public set AutoFilter(value: boolean) {
    this.model.AutoFilter = value;
    this.markProtectedWhenEnabled(value);
  }

  public get PivotTables(): boolean {
    return this.model.PivotTables;
  }

  public set PivotTables(value: boolean) {
    this.model.PivotTables = value;
    this.markProtectedWhenEnabled(value);
  }

  public get SelectUnlockedCells(): boolean {
    return this.model.SelectUnlockedCells;
  }

  public set SelectUnlockedCells(value: boolean) {
    this.model.SelectUnlockedCells = value;
    this.markProtectedWhenEnabled(value);
  }

  /** Clears all protection flags. */
  public Reset(): void {
    this.model.Clear();
  }

  private markProtectedWhenEnabled(value: boolean): void {
    if (value) {
      this.model.IsProtected = true;
    }
  }
}
