import { CellsException } from "./CellsException";
import { RowModel } from "./Core/RowModel";
import { Worksheet } from "./worksheet";

/** Represents a worksheet row. */
export class Row {
  private readonly worksheet: Worksheet;
  private readonly index: number;

  public constructor(worksheet: Worksheet, index: number) {
    this.worksheet = worksheet;
    this.index = index;
  }

  /** Gets or sets the row height in points. */
  public get Height(): number | null {
    const model = this.tryGetModel();
    return model === undefined ? null : model.Height;
  }

  public set Height(value: number | null) {
    if (value !== null && value !== undefined && value <= 0) {
      throw new CellsException("Row height must be positive.");
    }

    const model = this.getOrCreateModel();
    model.Height = value;
    this.normalize();
  }

  /** Gets or sets whether the row is hidden. */
  public get IsHidden(): boolean {
    const model = this.tryGetModel();
    return model === undefined ? false : model.Hidden;
  }

  public set IsHidden(value: boolean) {
    const model = this.getOrCreateModel();
    model.Hidden = value;
    this.normalize();
  }

  private tryGetModel(): RowModel | undefined {
    return this.worksheet.Model.Rows.get(this.index);
  }

  private getOrCreateModel(): RowModel {
    const existing = this.tryGetModel();
    if (existing !== undefined) {
      return existing;
    }

    const created = new RowModel();
    this.worksheet.Model.Rows.set(this.index, created);
    return created;
  }

  private normalize(): void {
    const model = this.tryGetModel();
    if (
      model !== undefined &&
      (model.Height === null || model.Height === undefined) &&
      !model.Hidden &&
      (model.StyleIndex === null || model.StyleIndex === undefined)
    ) {
      this.worksheet.Model.Rows.delete(this.index);
    }
  }
}
