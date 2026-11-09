import { CellsException } from "./CellsException";

/** Represents an error raised for an invalid file format. */
export class InvalidFileFormatException extends CellsException {
  public constructor(message: string, innerException?: unknown) {
    super(message, innerException);
    this.name = "InvalidFileFormatException";
  }
}
