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

// PdfSaveOptions is now defined in ../types (the canonical location), which also
// carries the SaveOptions base class and HtmlSaveOptions. Re-exported from index.ts.
