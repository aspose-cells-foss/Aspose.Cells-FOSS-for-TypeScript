import { RenderColor } from "./RenderColor";
import { PdfTextDocumentSession } from "./PdfTextDocumentSession";

/** Bundles services shared across layout and rendering passes for a single export. */
export class RenderContext {
  defaultFont?: string;
  colors: RenderColor;
  culture = "en-US";
  conditionalFormattingEvaluationDate: Date;
  maxDigitWidthPixels = 7.5;
  fontDerivedRowHeightPt = 15;

  /** Optional text optimization - stores PDF text runs for post-processing. */
  pdfTextSession?: PdfTextDocumentSession;
  enableWorksheetTextOptimization = false;

  constructor(defaultFont?: string, themeXml?: string) {
    this.defaultFont = defaultFont;
    this.colors = RenderColor.fromWorkbook(themeXml);
    this.conditionalFormattingEvaluationDate = new Date();
  }
}
