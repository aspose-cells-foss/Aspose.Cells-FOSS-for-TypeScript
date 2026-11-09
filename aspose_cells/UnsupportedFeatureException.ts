import { CellsException } from "./CellsException";

/** Represents an error raised for an unsupported feature. */
export class UnsupportedFeatureException extends CellsException {
  public constructor(message: string) {
    super(message);
    this.name = "UnsupportedFeatureException";
  }
}
