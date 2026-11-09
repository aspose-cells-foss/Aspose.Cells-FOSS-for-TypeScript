import { SaveOptions } from "./SaveOptions";
import { SaveFormat } from "./enums";

/** Options controlling XLSX-to-PDF rendering. */
export class PdfSaveOptions extends SaveOptions {
  public OnePagePerSheet: boolean = false;
  public AllColumnsInOnePagePerSheet: boolean = false;
  public DefaultFont: string | null = null;
  public ConditionalFormattingEvaluationDate: Date | null = null;
  public UseType3TextOptimization: boolean = true;
  public UseTrueTypeSubsetPrototype: boolean = true;

  public constructor() {
    super();
    this.SaveFormat = SaveFormat.Pdf;
  }
}
