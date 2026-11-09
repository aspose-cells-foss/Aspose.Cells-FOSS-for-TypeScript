import type { FilterColumn as LegacyFilterColumn } from "./types"
import { cellRef, escapeXml } from "./util"

export class LegacyAutoFilter {
  private _range: string
  private _columns: LegacyFilterColumn[] = []

  constructor(range: string) {
    this._range = range
  }

  get range(): string {
    return this._range
  }

  set range(value: string) {
    this._range = value
  }

  get columns(): LegacyFilterColumn[] {
    return this._columns
  }

  addFilterColumn(col: number, filters: (string | number)[], blank?: boolean) {
    this._columns.push({ col, filters, blank })
  }

  removeFilterColumn(col: number) {
    const idx = this._columns.findIndex((c) => c.col === col)
    if (idx >= 0) {
      this._columns.splice(idx, 1)
    }
  }

  clear() {
    this._columns = []
  }

  toXml(): string {
    let xml = `<autoFilter ref="${this._range}">`
    for (const col of this._columns) {
      xml += `<filterColumn colId="${col.col}">`
      if (col.blank) {
        xml += "<blank/>"
      }
      for (const filter of col.filters) {
        xml += `<filters><filter val="${escapeXml(String(filter))}"/></filters>`
      }
      xml += "</filterColumn>"
    }
    xml += "</autoFilter>"
    return xml
  }
}

import { AutoFilterColorFilterModel } from "./Core/AutoFilterColorFilterModel";
import { AutoFilterCustomFilterModel } from "./Core/AutoFilterCustomFilterModel";
import { AutoFilterDynamicFilterModel } from "./Core/AutoFilterDynamicFilterModel";
import { AutoFilterModel } from "./Core/AutoFilterModel";
import { AutoFilterSortConditionModel } from "./Core/AutoFilterSortConditionModel";
import { AutoFilterSortStateModel } from "./Core/AutoFilterSortStateModel";
import { AutoFilterTop10Model } from "./Core/AutoFilterTop10Model";
import { CellsException } from "./CellsException";
import { FilterColumnModel } from "./Core/FilterColumnModel";
import { FilterOperatorType } from "./enums";

/** Trims an optional text value, returning an empty string when nothing remains. */
function normalizeOptionalText(value: string | null | undefined): string {
  if (value === null || value === undefined || value.trim().length === 0) {
    return "";
  }

  return value.trim();
}

/** Requires a non-null text value. */
function normalizeRequiredText(value: string, parameterName: string): string {
  if (value === null || value === undefined) {
    throw new Error(parameterName);
  }

  return value;
}

/** Normalizes an optional cell or range reference. */
function normalizeOptionalRange(value: string | null | undefined): string {
  return normalizeOptionalText(value);
}

/** Normalizes a required cell or range reference. */
function normalizeRequiredRange(value: string, parameterName: string): string {
  const normalized = normalizeOptionalText(value);
  if (normalized.length === 0) {
    throw new CellsException(parameterName + " must be a valid cell or range reference.");
  }

  return normalized;
}

/** Maps an operator to its SpreadsheetML attribute name, mirroring AutoFilterSupport.ToOperatorName. */
function toOperatorName(operatorType: FilterOperatorType): string {
  switch (operatorType) {
    case FilterOperatorType.LessThan:
      return "lessThan";
    case FilterOperatorType.LessOrEqual:
      return "lessThanOrEqual";
    case FilterOperatorType.NotEqual:
      return "notEqual";
    case FilterOperatorType.GreaterOrEqual:
      return "greaterThanOrEqual";
    case FilterOperatorType.GreaterThan:
      return "greaterThan";
    default:
      return "";
  }
}

/** Parses an operator name, defaulting to Equal, mirroring AutoFilterSupport.ParseOperatorOrDefault. */
function parseOperatorOrDefault(value: string): FilterOperatorType {
  if (value === null || value === undefined || value.trim().length === 0) {
    return FilterOperatorType.Equal;
  }

  switch (value.trim()) {
    case "lessThan":
      return FilterOperatorType.LessThan;
    case "lessThanOrEqual":
      return FilterOperatorType.LessOrEqual;
    case "notEqual":
      return FilterOperatorType.NotEqual;
    case "greaterThanOrEqual":
      return FilterOperatorType.GreaterOrEqual;
    case "greaterThan":
      return FilterOperatorType.GreaterThan;
    default:
      return FilterOperatorType.Equal;
  }
}

/** Represents a collection of filter value objects. */
export class FilterValueCollection {
  private readonly values: string[];

  public constructor(values: string[]) {
    this.values = values;
  }

  public get Count(): number {
    return this.values.length;
  }

  public get(index: number): string {
    if (index < 0 || index >= this.values.length) {
      throw new CellsException("Filter value index was out of range.");
    }

    return this.values[index];
  }

  public at(index: number): string {
    return this.get(index);
  }

  public Add(value: string): number {
    const normalized = normalizeRequiredText(value, "value");
    this.values.push(normalized);
    return this.values.length - 1;
  }

  public RemoveAt(index: number): void {
    if (index < 0 || index >= this.values.length) {
      throw new CellsException("Filter value index was out of range.");
    }

    this.values.splice(index, 1);
  }

  public Clear(): void {
    this.values.length = 0;
  }

  public *[Symbol.iterator](): IterableIterator<string> {
    for (let index = 0; index < this.values.length; index++) {
      yield this.values[index];
    }
  }
}

/** Represents auto filter custom filter. */
export class AutoFilterCustomFilter {
  public constructor(private readonly model: AutoFilterCustomFilterModel) {}

  public get Operator(): FilterOperatorType {
    return parseOperatorOrDefault(this.model.Operator);
  }

  public set Operator(value: FilterOperatorType) {
    this.model.Operator = toOperatorName(value);
  }

  public get Value(): string {
    return this.model.Value;
  }

  public set Value(value: string) {
    this.model.Value = normalizeRequiredText(value, "Value");
  }
}

/** Represents a collection of auto filter custom filter objects. */
export class AutoFilterCustomFilterCollection {
  private readonly models: AutoFilterCustomFilterModel[];
  private readonly columnModel: FilterColumnModel;

  public constructor(models: AutoFilterCustomFilterModel[], columnModel: FilterColumnModel) {
    this.models = models;
    this.columnModel = columnModel;
  }

  public get MatchAll(): boolean {
    return this.columnModel.CustomFiltersAnd;
  }

  public set MatchAll(value: boolean) {
    this.columnModel.CustomFiltersAnd = value;
  }

  public get Count(): number {
    return this.models.length;
  }

  public get(index: number): AutoFilterCustomFilter {
    if (index < 0 || index >= this.models.length) {
      throw new CellsException("Custom filter index was out of range.");
    }

    return new AutoFilterCustomFilter(this.models[index]);
  }

  public at(index: number): AutoFilterCustomFilter {
    return this.get(index);
  }

  public Add(operatorType: FilterOperatorType, value: string): number {
    if (this.models.length >= 2) {
      throw new CellsException("Custom filters support at most two filter conditions.");
    }

    const model = new AutoFilterCustomFilterModel();
    model.Operator = toOperatorName(operatorType);
    model.Value = normalizeRequiredText(value, "value");
    this.models.push(model);
    return this.models.length - 1;
  }

  public RemoveAt(index: number): void {
    if (index < 0 || index >= this.models.length) {
      throw new CellsException("Custom filter index was out of range.");
    }

    this.models.splice(index, 1);
  }

  public Clear(): void {
    this.models.length = 0;
    this.columnModel.CustomFiltersAnd = false;
  }

  public *[Symbol.iterator](): IterableIterator<AutoFilterCustomFilter> {
    for (let index = 0; index < this.models.length; index++) {
      yield this.get(index);
    }
  }
}

/** Represents auto filter color filter. */
export class AutoFilterColorFilter {
  public constructor(private readonly model: AutoFilterColorFilterModel) {}

  public get Enabled(): boolean {
    return this.model.Enabled;
  }

  public set Enabled(value: boolean) {
    if (!value) {
      this.model.Clear();
      return;
    }

    this.model.Enabled = true;
  }

  public get DifferentialStyleId(): number | null {
    return this.model.DifferentialStyleId;
  }

  public set DifferentialStyleId(value: number | null) {
    if (value !== null && value !== undefined && value < 0) {
      throw new CellsException("Differential style id must be zero or greater.");
    }

    this.model.DifferentialStyleId = value;
    if (value !== null && value !== undefined) {
      this.model.Enabled = true;
    }
  }

  public get CellColor(): boolean {
    return this.model.CellColor;
  }

  public set CellColor(value: boolean) {
    this.model.CellColor = value;
    if (value) {
      this.model.Enabled = true;
    }
  }

  public Clear(): void {
    this.model.Clear();
  }
}

/** Represents auto filter dynamic filter. */
export class AutoFilterDynamicFilter {
  public constructor(private readonly model: AutoFilterDynamicFilterModel) {}

  public get Enabled(): boolean {
    return this.model.Enabled;
  }

  public set Enabled(value: boolean) {
    if (!value) {
      this.model.Clear();
      return;
    }

    this.model.Enabled = true;
  }

  public get Type(): string {
    return this.model.Type;
  }

  public set Type(value: string) {
    this.model.Type = normalizeOptionalText(value);
    if (this.model.Type.length > 0) {
      this.model.Enabled = true;
    }
  }

  public get Value(): number | null {
    return this.model.Value;
  }

  public set Value(value: number | null) {
    this.model.Value = value;
    if (value !== null && value !== undefined) {
      this.model.Enabled = true;
    }
  }

  public get MaxValue(): number | null {
    return this.model.MaxValue;
  }

  public set MaxValue(value: number | null) {
    this.model.MaxValue = value;
    if (value !== null && value !== undefined) {
      this.model.Enabled = true;
    }
  }

  public Clear(): void {
    this.model.Clear();
  }
}

/** Represents auto filter top10. */
export class AutoFilterTop10 {
  public constructor(private readonly model: AutoFilterTop10Model) {}

  public get Enabled(): boolean {
    return this.model.Enabled;
  }

  public set Enabled(value: boolean) {
    if (!value) {
      this.model.Clear();
      return;
    }

    this.model.Enabled = true;
  }

  public get Top(): boolean {
    return this.model.Top;
  }

  public set Top(value: boolean) {
    this.model.Top = value;
    this.model.Enabled = true;
  }

  public get Percent(): boolean {
    return this.model.Percent;
  }

  public set Percent(value: boolean) {
    this.model.Percent = value;
    this.model.Enabled = true;
  }

  public get Value(): number | null {
    return this.model.Value;
  }

  public set Value(value: number | null) {
    this.model.Value = value;
    if (value !== null && value !== undefined) {
      this.model.Enabled = true;
    }
  }

  public get FilterValue(): number | null {
    return this.model.FilterValue;
  }

  public set FilterValue(value: number | null) {
    this.model.FilterValue = value;
    if (value !== null && value !== undefined) {
      this.model.Enabled = true;
    }
  }

  public Clear(): void {
    this.model.Clear();
  }
}

/** Represents auto filter sort condition. */
export class AutoFilterSortCondition {
  public constructor(private readonly model: AutoFilterSortConditionModel) {}

  public get Ref(): string {
    return this.model.Ref;
  }

  public set Ref(value: string) {
    this.model.Ref = normalizeRequiredRange(value, "Ref");
  }

  public get Descending(): boolean {
    return this.model.Descending;
  }

  public set Descending(value: boolean) {
    this.model.Descending = value;
  }

  public get SortBy(): string {
    return this.model.SortBy;
  }

  public set SortBy(value: string) {
    this.model.SortBy = normalizeOptionalText(value);
  }

  public get CustomList(): string {
    return this.model.CustomList;
  }

  public set CustomList(value: string) {
    this.model.CustomList = normalizeOptionalText(value);
  }

  public get DifferentialStyleId(): number | null {
    return this.model.DifferentialStyleId;
  }

  public set DifferentialStyleId(value: number | null) {
    if (value !== null && value !== undefined && value < 0) {
      throw new CellsException("Differential style id must be zero or greater.");
    }

    this.model.DifferentialStyleId = value;
  }

  public get IconSet(): string {
    return this.model.IconSet;
  }

  public set IconSet(value: string) {
    this.model.IconSet = normalizeOptionalText(value);
  }

  public get IconId(): number | null {
    return this.model.IconId;
  }

  public set IconId(value: number | null) {
    if (value !== null && value !== undefined && value < 0) {
      throw new CellsException("Icon id must be zero or greater.");
    }

    this.model.IconId = value;
  }
}

/** Represents a collection of auto filter sort condition objects. */
export class AutoFilterSortConditionCollection {
  private readonly models: AutoFilterSortConditionModel[];

  public constructor(models: AutoFilterSortConditionModel[]) {
    this.models = models;
  }

  public get Count(): number {
    return this.models.length;
  }

  public get(index: number): AutoFilterSortCondition {
    if (index < 0 || index >= this.models.length) {
      throw new CellsException("Sort condition index was out of range.");
    }

    return new AutoFilterSortCondition(this.models[index]);
  }

  public at(index: number): AutoFilterSortCondition {
    return this.get(index);
  }

  public Add(reference: string): number {
    const model = new AutoFilterSortConditionModel();
    model.Ref = normalizeRequiredRange(reference, "reference");
    this.models.push(model);
    return this.models.length - 1;
  }

  public RemoveAt(index: number): void {
    if (index < 0 || index >= this.models.length) {
      throw new CellsException("Sort condition index was out of range.");
    }

    this.models.splice(index, 1);
  }

  public Clear(): void {
    this.models.length = 0;
  }

  public *[Symbol.iterator](): IterableIterator<AutoFilterSortCondition> {
    for (let index = 0; index < this.models.length; index++) {
      yield this.get(index);
    }
  }
}

// Implemented by serializer module.
/** Represents an auto filter column. */
export class FilterColumn {
  public readonly Filters: FilterValueCollection;
  public readonly CustomFilters: AutoFilterCustomFilterCollection;
  public readonly ColorFilter: AutoFilterColorFilter;
  public readonly DynamicFilter: AutoFilterDynamicFilter;
  public readonly Top10: AutoFilterTop10;

  public constructor(private readonly model: FilterColumnModel) {
    this.Filters = new FilterValueCollection(model.Filters);
    this.CustomFilters = new AutoFilterCustomFilterCollection(model.CustomFilters, model);
    this.ColorFilter = new AutoFilterColorFilter(model.ColorFilter);
    this.DynamicFilter = new AutoFilterDynamicFilter(model.DynamicFilter);
    this.Top10 = new AutoFilterTop10(model.Top10);
  }

  public get Model(): FilterColumnModel {
    return this.model;
  }

  public get ColumnIndex(): number {
    return this.model.ColumnIndex;
  }

  public set ColumnIndex(value: number) {
    this.model.ColumnIndex = value;
  }

  public get HiddenButton(): boolean {
    return this.model.HiddenButton;
  }

  public set HiddenButton(value: boolean) {
    this.model.HiddenButton = value;
  }

  public Clear(): void {
    this.model.ClearCriteria();
  }
}

// Implemented by serializer module.
/** Represents a collection of auto filter columns. */
export class FilterColumnCollection {
  private readonly models: FilterColumnModel[];

  public constructor(models: FilterColumnModel[]) {
    this.models = models;
  }

  public get Count(): number {
    return this.models.length;
  }

  public get(index: number): FilterColumn {
    if (index < 0 || index >= this.models.length) {
      throw new CellsException(`Filter column index ${index} is out of range.`);
    }

    return new FilterColumn(this.models[index]);
  }

  public at(index: number): FilterColumn {
    return this.get(index);
  }

  public Add(columnIndex: number): number {
    if (columnIndex < 0) {
      throw new CellsException("Filter column index must be non-negative.");
    }

    const model = new FilterColumnModel();
    model.ColumnIndex = columnIndex;
    this.models.push(model);
    return this.models.length - 1;
  }

  public RemoveAt(index: number): void {
    if (index < 0 || index >= this.models.length) {
      throw new CellsException(`Filter column index ${index} is out of range.`);
    }

    this.models.splice(index, 1);
  }

  public Clear(): void {
    this.models.length = 0;
  }

  public *[Symbol.iterator](): IterableIterator<FilterColumn> {
    for (let index = 0; index < this.models.length; index++) {
      yield this.get(index);
    }
  }
}

// Implemented by serializer module.
/** Represents the sort state of an auto filter. */
export class AutoFilterSortState {
  public readonly SortConditions: AutoFilterSortConditionCollection;

  public constructor(private readonly model: AutoFilterSortStateModel) {
    this.SortConditions = new AutoFilterSortConditionCollection(model.Conditions);
  }

  public get Model(): AutoFilterSortStateModel {
    return this.model;
  }

  public get ColumnSort(): boolean {
    return this.model.ColumnSort;
  }

  public set ColumnSort(value: boolean) {
    this.model.ColumnSort = value;
  }

  public get CaseSensitive(): boolean {
    return this.model.CaseSensitive;
  }

  public set CaseSensitive(value: boolean) {
    this.model.CaseSensitive = value;
  }

  public get SortMethod(): string {
    return this.model.SortMethod;
  }

  public set SortMethod(value: string) {
    this.model.SortMethod = normalizeOptionalText(value);
  }

  public get Ref(): string {
    return this.model.Ref;
  }

  public set Ref(value: string) {
    this.model.Ref = normalizeOptionalRange(value);
  }

  /** Sort conditions are represented by the serializer module through the core model. */
  public get Conditions(): readonly AutoFilterSortConditionModel[] {
    return this.model.Conditions;
  }

  public Clear(): void {
    this.model.Clear();
  }
}

// Implemented by serializer module.
/** Represents auto filter settings for a worksheet. */
export class AutoFilter {
  private readonly model: AutoFilterModel;
  public readonly FilterColumns: FilterColumnCollection;
  public readonly SortState: AutoFilterSortState;

  public constructor(model: AutoFilterModel) {
    this.model = model;
    this.FilterColumns = new FilterColumnCollection(model.FilterColumns);
    this.SortState = new AutoFilterSortState(model.SortState);
  }

  public get Model(): AutoFilterModel {
    return this.model;
  }

  public get Range(): string {
    return this.model.Range;
  }

  public set Range(value: string) {
    this.model.Range = value === null || value === undefined ? "" : value.trim();
  }

  public Clear(): void {
    this.model.Clear();
  }
}
