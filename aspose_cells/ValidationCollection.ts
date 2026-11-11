import { CellsException } from "./CellsException";
import { CellArea } from "./CellArea";
import { Validation } from "./validation";
import { ValidationModel } from "./Core/ValidationModel";

// Implemented by serializer module.
/** Represents a collection of data validations. */
export class ValidationCollection {
  [key: string]: any;

  private readonly models: ValidationModel[];

  public constructor(models: ValidationModel[]) {
    this.models = models;
  }

  public get Count(): number {
    return this.models.length;
  }

  public get(index: number): Validation {
    if (index < 0 || index >= this.models.length) {
      throw new CellsException("Validation index was out of range.");
    }

    return new Validation(this.models, this.models[index]);
  }

  public at(index: number): Validation {
    return this.get(index);
  }

  public Add(area: CellArea): number {
    ValidationCollection.validateArea(area);
    ValidationCollection.ensureNoOverlap(this.models, null, area);

    const model = new ValidationModel();
    model.Areas.push(new CellArea(area.FirstRow, area.FirstColumn, area.TotalRows, area.TotalColumns));
    this.models.push(model);
    return this.models.length - 1;
  }

  public GetValidationInCell(row: number, column: number): Validation | null {
    if (row < 0 || column < 0) {
      throw new CellsException("Row and column indices must be non-negative.");
    }

    for (const model of this.models) {
      for (const area of model.Areas) {
        if (ValidationCollection.contains(area, row, column)) {
          return new Validation(this.models, model);
        }
      }
    }

    return null;
  }

  public RemoveACell(row: number, column: number): void {
    if (row < 0 || column < 0) {
      throw new CellsException("Row and column indices must be non-negative.");
    }

    this.RemoveArea(new CellArea(row, column, 1, 1));
  }

  public RemoveArea(cellArea: CellArea): void {
    ValidationCollection.validateArea(cellArea);

    for (let index = this.models.length - 1; index >= 0; index--) {
      const model = this.models[index];
      ValidationCollection.replaceAreas(model, ValidationCollection.subtractAreas(model.Areas, cellArea));
      if (model.Areas.length === 0) {
        this.models.splice(index, 1);
      }
    }
  }

  public static AddAreaToValidation(owner: ValidationModel[], model: ValidationModel, area: CellArea): void {
    ValidationCollection.validateArea(area);
    ValidationCollection.ensureNoOverlap(owner, model, area);

    for (const current of model.Areas) {
      if (ValidationCollection.areasOverlap(current, area)) {
        throw new CellsException("Validation areas must not overlap.");
      }
    }

    model.Areas.push(new CellArea(area.FirstRow, area.FirstColumn, area.TotalRows, area.TotalColumns));
    ValidationCollection.sortAreas(model.Areas);
  }

  public static RemoveAreaFromValidation(owner: ValidationModel[], model: ValidationModel, area: CellArea): void {
    ValidationCollection.validateArea(area);
    ValidationCollection.replaceAreas(model, ValidationCollection.subtractAreas(model.Areas, area));
    if (model.Areas.length === 0) {
      const index = owner.indexOf(model);
      if (index >= 0) {
        owner.splice(index, 1);
      }
    }
  }

  public *[Symbol.iterator](): IterableIterator<Validation> {
    for (let index = 0; index < this.models.length; index++) {
      yield this.get(index);
    }
  }

  private static replaceAreas(model: ValidationModel, areas: readonly CellArea[]): void {
    model.Areas.length = 0;
    for (const area of areas) {
      model.Areas.push(area);
    }
  }

  private static validateArea(area: CellArea | null | undefined): void {
    if (
      area === null ||
      area === undefined ||
      area.FirstRow < 0 ||
      area.FirstColumn < 0 ||
      area.TotalRows <= 0 ||
      area.TotalColumns <= 0
    ) {
      throw new CellsException("Validation area must be a positive cell range.");
    }
  }

  private static ensureNoOverlap(
    validations: readonly ValidationModel[],
    currentValidation: ValidationModel | null,
    candidate: CellArea,
  ): void {
    for (const model of validations) {
      if (model === currentValidation) {
        continue;
      }

      for (const area of model.Areas) {
        if (ValidationCollection.areasOverlap(area, candidate)) {
          throw new CellsException("Validation areas must not overlap.");
        }
      }
    }
  }

  private static subtractAreas(sourceAreas: readonly CellArea[], removal: CellArea): CellArea[] {
    const remaining: CellArea[] = [];
    for (const source of sourceAreas) {
      ValidationCollection.subtractArea(source, removal, remaining);
    }

    ValidationCollection.sortAreas(remaining);
    return remaining;
  }

  private static subtractArea(source: CellArea, removal: CellArea, output: CellArea[]): void {
    if (!ValidationCollection.areasOverlap(source, removal)) {
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

    ValidationCollection.addIfNonEmpty(output, source.FirstRow, source.FirstColumn, overlapFirstRow - 1, sourceLastColumn);
    ValidationCollection.addIfNonEmpty(output, overlapLastRow + 1, source.FirstColumn, sourceLastRow, sourceLastColumn);
    ValidationCollection.addIfNonEmpty(output, overlapFirstRow, source.FirstColumn, overlapLastRow, overlapFirstColumn - 1);
    ValidationCollection.addIfNonEmpty(output, overlapFirstRow, overlapLastColumn + 1, overlapLastRow, sourceLastColumn);
  }

  private static addIfNonEmpty(
    areas: CellArea[],
    firstRow: number,
    firstColumn: number,
    lastRow: number,
    lastColumn: number,
  ): void {
    if (lastRow < firstRow || lastColumn < firstColumn) {
      return;
    }

    areas.push(new CellArea(firstRow, firstColumn, lastRow - firstRow + 1, lastColumn - firstColumn + 1));
  }

  private static areasOverlap(left: CellArea, right: CellArea): boolean {
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

  private static contains(area: CellArea, row: number, column: number): boolean {
    return (
      row >= area.FirstRow &&
      row < area.FirstRow + area.TotalRows &&
      column >= area.FirstColumn &&
      column < area.FirstColumn + area.TotalColumns
    );
  }

  private static sortAreas(areas: CellArea[]): void {
    areas.sort(ValidationCollection.compareAreas);
  }

  private static compareAreas(left: CellArea, right: CellArea): number {
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
}
