import { CellsException } from "./CellsException";
import { CellAddress } from "./Core/CellAddress";
import { Comment } from "./comment";
import { CommentModel } from "./Core/CommentModel";

// Implemented by serializer module.
/** Represents a collection of cell comments. */
export class CommentCollection {
  private readonly models: CommentModel[];

  public constructor(models: CommentModel[]) {
    this.models = models;
  }

  public get Count(): number {
    return this.models.length;
  }

  public get(index: number): Comment;
  public get(cellName: string): Comment | null;
  public get(indexOrCellName: number | string): Comment | null {
    return typeof indexOrCellName === "number"
      ? this.getByIndex(indexOrCellName)
      : this.getByName(indexOrCellName);
  }

  public at(index: number): Comment {
    return this.getByIndex(index);
  }

  public getByName(cellName: string): Comment | null {
    const address = CellAddress.Parse(cellName);
    for (const model of this.models) {
      if (model.Row === address.RowIndex && model.Column === address.ColumnIndex) {
        return new Comment(model);
      }
    }

    return null;
  }

  public Add(row: number, column: number): Comment;
  public Add(cellName: string): Comment;
  public Add(rowOrCellName: number | string, column?: number): Comment {
    const row = typeof rowOrCellName === "number" ? rowOrCellName : CellAddress.Parse(rowOrCellName).RowIndex;
    const columnIndex =
      typeof rowOrCellName === "number" ? (column ?? 0) : CellAddress.Parse(rowOrCellName).ColumnIndex;

    if (row < 0 || columnIndex < 0) {
      throw new CellsException("Comment coordinates must be non-negative.");
    }

    for (const existing of this.models) {
      if (existing.Row === row && existing.Column === columnIndex) {
        throw new CellsException(`A comment already exists at row ${row}, column ${columnIndex}.`);
      }
    }

    const model = new CommentModel();
    model.Row = row;
    model.Column = columnIndex;
    this.models.push(model);
    return new Comment(model);
  }

  public RemoveAt(index: number): void;
  public RemoveAt(cellName: string): void;
  public RemoveAt(indexOrCellName: number | string): void {
    if (typeof indexOrCellName === "string") {
      const address = CellAddress.Parse(indexOrCellName);
      for (let index = 0; index < this.models.length; index++) {
        const model = this.models[index];
        if (model.Row === address.RowIndex && model.Column === address.ColumnIndex) {
          this.models.splice(index, 1);
          return;
        }
      }

      return;
    }

    if (indexOrCellName < 0 || indexOrCellName >= this.models.length) {
      throw new CellsException(`Comment index ${indexOrCellName} is out of range.`);
    }

    this.models.splice(indexOrCellName, 1);
  }

  public *[Symbol.iterator](): IterableIterator<Comment> {
    for (let index = 0; index < this.models.length; index++) {
      yield this.getByIndex(index);
    }
  }

  private getByIndex(index: number): Comment {
    if (index < 0 || index >= this.models.length) {
      throw new CellsException("Comment index was out of range.");
    }

    return new Comment(this.models[index]);
  }
}
