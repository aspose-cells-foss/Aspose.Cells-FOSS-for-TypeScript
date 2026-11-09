import { CellArea } from "../CellArea";
import { FormatConditionModel } from "./FormatConditionModel";

export class ConditionalFormattingModel {
  public readonly Areas: CellArea[];
  public readonly Conditions: FormatConditionModel[];

  public constructor() {
    this.Areas = [];
    this.Conditions = [];
  }
}
