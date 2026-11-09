import { CellsException } from "./CellsException";

/** Represents an error raised while working with formulas. */
export class FormulaException extends CellsException {
  public constructor(message: string) {
    super(message);
    this.name = "FormulaException";
  }
}
