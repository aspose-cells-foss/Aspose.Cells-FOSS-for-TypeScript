import { CellsException } from "./CellsException";

/** Represents an error that occurs during workbook load. */
export class WorkbookLoadException extends CellsException {
  public constructor(message: string, innerException?: unknown) {
    super(message, innerException);
    this.name = "WorkbookLoadException";
  }
}
