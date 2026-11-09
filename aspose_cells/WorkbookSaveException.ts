import { CellsException } from "./CellsException";

/** Represents an error that occurs during workbook save. */
export class WorkbookSaveException extends CellsException {
  public constructor(message: string, innerException?: unknown) {
    super(message, innerException);
    this.name = "WorkbookSaveException";
  }
}
