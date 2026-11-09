import { CellsException } from "./CellsException";

/** Represents an error raised while working with styles. */
export class StyleException extends CellsException {
  public constructor(message: string) {
    super(message);
    this.name = "StyleException";
  }
}
