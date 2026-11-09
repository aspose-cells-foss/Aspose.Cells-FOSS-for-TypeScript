/**
 * Represents an error that occurs during cells operations.
 * Ported from Aspose.Cells_FOSS.CellsException.
 */
export class CellsException extends Error {
  public readonly innerException: unknown;

  public constructor(message: string, innerException?: unknown) {
    super(message);
    this.name = "CellsException";
    this.innerException = innerException ?? null;
    Object.setPrototypeOf(this, new.target.prototype);
  }
}
