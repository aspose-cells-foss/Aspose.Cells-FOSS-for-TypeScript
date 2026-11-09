import { CellsException } from "./CellsException";
import { ColumnRangeModel } from "./Core/ColumnRangeModel";
import { Worksheet } from "./worksheet";

/** Represents a worksheet column. */
export class Column {
  private readonly worksheet: Worksheet;
  private readonly index: number;

  public constructor(worksheet: Worksheet, index: number) {
    this.worksheet = worksheet;
    this.index = index;
  }

  /** Gets or sets the column width in characters. */
  public get Width(): number | null {
    const model = this.findModel();
    return model === undefined ? null : model.Width;
  }

  public set Width(value: number | null) {
    if (value !== null && value !== undefined && value <= 0) {
      throw new CellsException("Column width must be positive.");
    }

    const model = this.getOrCreateModel();
    model.Width = value;
    this.normalize(model);
  }

  /** Gets or sets whether the column is hidden. */
  public get IsHidden(): boolean {
    const model = this.findModel();
    return model === undefined ? false : model.Hidden;
  }

  public set IsHidden(value: boolean) {
    const model = this.getOrCreateModel();
    model.Hidden = value;
    this.normalize(model);
  }

  private findModel(): ColumnRangeModel | undefined {
    const columns = this.worksheet.Model.Columns;
    for (let index = columns.length - 1; index >= 0; index--) {
      const model = columns[index];
      if (model.MinColumnIndex <= this.index && model.MaxColumnIndex >= this.index) {
        return model;
      }
    }

    return undefined;
  }

  private getOrCreateModel(): ColumnRangeModel {
    const existing = this.findModel();
    if (
      existing !== undefined &&
      existing.MinColumnIndex === this.index &&
      existing.MaxColumnIndex === this.index
    ) {
      return existing;
    }

    const created = new ColumnRangeModel();
    created.MinColumnIndex = this.index;
    created.MaxColumnIndex = this.index;
    this.worksheet.Model.Columns.push(created);
    return created;
  }

  private normalize(model: ColumnRangeModel): void {
    if (
      (model.Width === null || model.Width === undefined) &&
      !model.Hidden &&
      (model.StyleIndex === null || model.StyleIndex === undefined)
    ) {
      const index = this.worksheet.Model.Columns.indexOf(model);
      if (index >= 0) {
        this.worksheet.Model.Columns.splice(index, 1);
      }
    }
  }
}
