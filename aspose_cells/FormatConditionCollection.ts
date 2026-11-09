import { CellArea } from "./CellArea";
import { CellsException } from "./CellsException";
import { FormatCondition } from "./FormatCondition";
import { FormatConditionType, OperatorType } from "./enums";
import { ConditionalFormattingModel } from "./Core/ConditionalFormattingModel";
import { FormatConditionModel } from "./Core/FormatConditionModel";
import { StyleValue } from "./Core/StyleValue";
import { ColorValue } from "./Core/ColorValue";
import { toApiCellArea, toCoreCellArea } from "./CellAreaConvert";
import { normalizeFormula } from "./FormatCondition";

/** Represents a collection of format condition objects. */
export class FormatConditionCollection {
  private readonly owner: ConditionalFormattingModel[];
  private readonly model: ConditionalFormattingModel;

  public constructor(owner: ConditionalFormattingModel[], model: ConditionalFormattingModel) {
    this.owner = owner;
    this.model = model;
  }

  public get Model(): ConditionalFormattingModel {
    return this.model;
  }

  public get Count(): number {
    return this.model.Conditions.length;
  }

  public get RangeCount(): number {
    return this.model.Areas.length;
  }

  public get(index: number): FormatCondition {
    if (index < 0 || index >= this.model.Conditions.length) {
      throw new CellsException("Format condition index was out of range.");
    }

    return new FormatCondition(this.owner, this.model, this.model.Conditions[index]);
  }

  public at(index: number): FormatCondition {
    return this.get(index);
  }

  public *[Symbol.iterator](): IterableIterator<FormatCondition> {
    for (let index = 0; index < this.model.Conditions.length; index++) {
      yield this.get(index);
    }
  }

  public Add(
    area: CellArea,
    type: FormatConditionType,
    operatorType: OperatorType,
    formula1: string,
    formula2: string,
  ): number {
    this.AddArea(area);
    return this.AddCondition(type, operatorType, formula1, formula2);
  }

  public AddCondition(type: FormatConditionType, operatorType?: OperatorType, formula1?: string, formula2?: string): number {
    const condition = new FormatConditionModel();
    condition.Type = type;
    condition.Operator = operatorType ?? OperatorType.None;
    condition.Formula1 = normalizeFormula(formula1);
    condition.Formula2 = normalizeFormula(formula2);
    condition.Priority = getNextPriority(this.owner);
    condition.Style = StyleValue.Default.clone();
    initializeDefaults(condition);
    this.model.Conditions.push(condition);
    return this.model.Conditions.length - 1;
  }

  public AddArea(area: CellArea): void {
    validateArea(area);
    const areas = this.model.Areas.map(toApiCellArea);
    areas.push(area);
    FormatConditionCollection.sortAreas(areas);
    replaceAreas(this.model, areas);
  }

  public GetCellArea(index: number): CellArea {
    if (index < 0 || index >= this.model.Areas.length) {
      throw new CellsException("Conditional formatting area index was out of range.");
    }

    return toApiCellArea(this.model.Areas[index]);
  }

  public RemoveArea(index: number): void {
    if (index < 0 || index >= this.model.Areas.length) {
      throw new CellsException("Conditional formatting area index was out of range.");
    }

    this.model.Areas.splice(index, 1);
    removeCollectionIfEmpty(this.owner, this.model);
  }

  public RemoveAreaByBounds(startRow: number, startColumn: number, totalRows: number, totalColumns: number): void {
    this.RemoveAreaByArea(
      CellArea.CreateCellArea(startRow, startColumn, startRow + totalRows - 1, startColumn + totalColumns - 1),
    );
  }

  /** Removes the given area from every range of this collection. */
  public RemoveAreaByArea(area: CellArea): void {
    validateArea(area);
    const areas = this.model.Areas.map(toApiCellArea);
    replaceAreas(this.model, subtractAreas(areas, area));
    removeCollectionIfEmpty(this.owner, this.model);
  }

  public RemoveCondition(index: number): void {
    if (index < 0 || index >= this.model.Conditions.length) {
      throw new CellsException("Format condition index was out of range.");
    }

    this.model.Conditions.splice(index, 1);
    removeCollectionIfEmpty(this.owner, this.model);
  }

  public static removeCondition(
    owner: ConditionalFormattingModel[],
    collection: ConditionalFormattingModel,
    model: FormatConditionModel,
  ): void {
    const position = collection.Conditions.indexOf(model);
    if (position >= 0) {
      collection.Conditions.splice(position, 1);
    }

    removeCollectionIfEmpty(owner, collection);
  }

  public static areasOverlap(left: CellArea, right: CellArea): boolean {
    return (
      left.FirstRow <= right.FirstRow + right.TotalRows - 1 &&
      right.FirstRow <= left.FirstRow + left.TotalRows - 1 &&
      left.FirstColumn <= right.FirstColumn + right.TotalColumns - 1 &&
      right.FirstColumn <= left.FirstColumn + left.TotalColumns - 1
    );
  }

  public static compareAreas(left: CellArea, right: CellArea): number {
    if (left.FirstRow !== right.FirstRow) {
      return left.FirstRow - right.FirstRow;
    }

    if (left.FirstColumn !== right.FirstColumn) {
      return left.FirstColumn - right.FirstColumn;
    }

    if (left.TotalRows !== right.TotalRows) {
      return left.TotalRows - right.TotalRows;
    }

    return left.TotalColumns - right.TotalColumns;
  }

  public static sortAreas(areas: CellArea[]): void {
    areas.sort(FormatConditionCollection.compareAreas);
  }
}

export function getNextPriority(collections: readonly ConditionalFormattingModel[]): number {
  let maxPriority = 0;
  for (const collection of collections) {
    for (const condition of collection.Conditions) {
      maxPriority = Math.max(maxPriority, condition.Priority);
    }
  }

  return maxPriority + 1;
}

function initializeDefaults(condition: FormatConditionModel): void {
  switch (condition.Type) {
    case FormatConditionType.DuplicateValues:
      condition.Duplicate = true;
      break;
    case FormatConditionType.UniqueValues:
      condition.Duplicate = false;
      break;
    case FormatConditionType.Top10:
      condition.Top = true;
      condition.Rank = 10;
      break;
    case FormatConditionType.Bottom10:
      condition.Top = false;
      condition.Rank = 10;
      break;
    case FormatConditionType.AboveAverage:
      condition.Above = true;
      break;
    case FormatConditionType.BelowAverage:
      condition.Above = false;
      break;
    case FormatConditionType.ColorScale:
      condition.ColorScaleCount = 2;
      break;
    case FormatConditionType.DataBar:
      condition.BarColor = new ColorValue(255, 99, 142, 198);
      break;
    case FormatConditionType.IconSet:
      condition.IconSetType = "3TrafficLights1";
      break;
    default:
      break;
  }
}

function removeCollectionIfEmpty(
  owner: ConditionalFormattingModel[],
  collection: ConditionalFormattingModel,
): void {
  if (collection.Areas.length === 0 || collection.Conditions.length === 0) {
    const position = owner.indexOf(collection);
    if (position >= 0) {
      owner.splice(position, 1);
    }
  }
}

function replaceAreas(model: ConditionalFormattingModel, areas: readonly CellArea[]): void {
  model.Areas.length = 0;
  for (const area of areas) {
    model.Areas.push(toCoreCellArea(area));
  }
}

function validateArea(area: CellArea): void {
  if (area.FirstRow < 0 || area.FirstColumn < 0 || area.TotalRows <= 0 || area.TotalColumns <= 0) {
    throw new CellsException("Conditional formatting area must be a positive cell range.");
  }
}

function subtractAreas(sourceAreas: readonly CellArea[], removal: CellArea): CellArea[] {
  const remaining: CellArea[] = [];
  for (const source of sourceAreas) {
    subtractArea(source, removal, remaining);
  }

  FormatConditionCollection.sortAreas(remaining);
  return remaining;
}

function subtractArea(source: CellArea, removal: CellArea, output: CellArea[]): void {
  if (!FormatConditionCollection.areasOverlap(source, removal)) {
    output.push(source);
    return;
  }

  const sourceLastRow = source.FirstRow + source.TotalRows - 1;
  const sourceLastColumn = source.FirstColumn + source.TotalColumns - 1;
  const removalLastRow = removal.FirstRow + removal.TotalRows - 1;
  const removalLastColumn = removal.FirstColumn + removal.TotalColumns - 1;

  const overlapFirstRow = Math.max(source.FirstRow, removal.FirstRow);
  const overlapFirstColumn = Math.max(source.FirstColumn, removal.FirstColumn);
  const overlapLastRow = Math.min(sourceLastRow, removalLastRow);
  const overlapLastColumn = Math.min(sourceLastColumn, removalLastColumn);

  addIfNonEmpty(output, source.FirstRow, source.FirstColumn, overlapFirstRow - 1, sourceLastColumn);
  addIfNonEmpty(output, overlapLastRow + 1, source.FirstColumn, sourceLastRow, sourceLastColumn);
  addIfNonEmpty(output, overlapFirstRow, source.FirstColumn, overlapLastRow, overlapFirstColumn - 1);
  addIfNonEmpty(output, overlapFirstRow, overlapLastColumn + 1, overlapLastRow, sourceLastColumn);
}

function addIfNonEmpty(
  areas: CellArea[],
  firstRow: number,
  firstColumn: number,
  lastRow: number,
  lastColumn: number,
): void {
  if (lastRow < firstRow || lastColumn < firstColumn) {
    return;
  }

  areas.push(CellArea.CreateCellArea(firstRow, firstColumn, lastRow, lastColumn));
}
