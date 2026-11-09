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

    return new Validation(this.models[index]);
  }

  public at(index: number): Validation {
    return this.get(index);
  }

  public Add(area: CellArea): number {
    if (area === null || area === undefined) {
      throw new CellsException("area");
    }

    const model = new ValidationModel();
    model.Areas.push(
      new CellArea(area.FirstRow, area.FirstColumn, area.TotalRows, area.TotalColumns),
    );
    this.models.push(model);
    return this.models.length - 1;
  }

  public GetValidationInCell(row: number, column: number): Validation | null {
    if (row < 0 || column < 0) {
      throw new CellsException("Row and column indices must be non-negative.");
    }

    for (const model of this.models) {
      for (const area of model.Areas) {
        if (area.StartRow <= row && area.EndRow >= row && area.StartColumn <= column && area.EndColumn >= column) {
          return new Validation(model);
        }
      }
    }

    return null;
  }

  public RemoveACell(row: number, column: number): void {
    this.RemoveArea(CellArea.CreateCellArea(row, column, row, column));
  }

  public RemoveArea(cellArea: CellArea): void {
    if (cellArea === null || cellArea === undefined) {
      throw new CellsException("cellArea");
    }

    for (let index = this.models.length - 1; index >= 0; index--) {
      const model = this.models[index];
      const remaining = model.Areas.filter((area) => !ValidationCollection.overlaps(area, cellArea));
      if (remaining.length === 0) {
        this.models.splice(index, 1);
        continue;
      }

      model.Areas.length = 0;
      for (const area of remaining) {
        model.Areas.push(area);
      }
    }
  }

  public *[Symbol.iterator](): IterableIterator<Validation> {
    for (let index = 0; index < this.models.length; index++) {
      yield this.get(index);
    }
  }

  private static overlaps(area: CellArea, target: CellArea): boolean {
    return (
      area.StartRow <= target.EndRow &&
      area.EndRow >= target.StartRow &&
      area.StartColumn <= target.EndColumn &&
      area.EndColumn >= target.StartColumn
    );
  }
}
