/** Core type definitions shared across the rendering module. */

export enum PageOrientation {
  Default = 0,
  Portrait = 1,
  Landscape = 2,
}

export class MergeRegion {
  constructor(
    public firstRow: number,
    public firstColumn: number,
    public totalRows: number,
    public totalColumns: number,
  ) {}
}

export class PageMargins {
  left = 0.7;
  right = 0.7;
  top = 0.75;
  bottom = 0.75;
  header = 0.3;
  footer = 0.3;
}

export interface PageSetupData {
  paperSize: number;
  orientation: PageOrientation;
  margins: PageMargins;
  printArea?: string;
  scale?: number;
  fitToPage?: boolean;
  fitToWidth?: number;
  fitToHeight?: number;
  horizontalPageBreaks: number[];
  verticalPageBreaks: number[];
  pageOrderOverThenDown?: boolean;
}

export class PdfSaveOptions {
  onePagePerSheet = false;
  allColumnsInOnePagePerSheet = false;
  defaultFont?: string;
  conditionalFormattingEvaluationDate?: Date;
}
