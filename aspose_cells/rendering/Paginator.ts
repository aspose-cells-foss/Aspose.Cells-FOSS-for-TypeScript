import { PageLayout } from "./PageLayout";
import { SheetLayout } from "./SheetLayout";
import { PaperSizes } from "./PaperSizes";
import { RenderUnits } from "./RenderUnits";
import type { PdfSaveOptions, PageSetupData } from "./types";
import { PageOrientation, MergeRegion } from "./types";

interface Span {
  start: number;
  end: number;
}

/**
 * Turns a sheet's resolved SheetLayout plus its page setup into a flat list of PageLayout pages,
 * applying scale/fit-to-page and greedy row/column splitting.
 */
export class Paginator {
  static paginate(
    layout: SheetLayout,
    sheetIndex: number,
    options: PdfSaveOptions,
    pageSetup: PageSetupData,
    printAreas: MergeRegion[],
  ): PageLayout[] {
    const areas = printAreas.length > 0 ? printAreas : [
      new MergeRegion(
        layout.firstRow, layout.firstColumn,
        layout.lastRow - layout.firstRow + 1,
        layout.lastColumn - layout.firstColumn + 1,
      ),
    ];

    const pages: PageLayout[] = [];
    for (const area of areas) {
      pages.push(...Paginator.paginateArea(layout, sheetIndex, options, pageSetup, area));
    }
    return pages;
  }

  private static paginateArea(
    layout: SheetLayout,
    sheetIndex: number,
    options: PdfSaveOptions,
    pageSetup: PageSetupData,
    area: MergeRegion,
  ): PageLayout[] {
    if (layout.isEmpty) return [];

    let { widthPt: pageWidthPt, heightPt: pageHeightPt } =
      PaperSizes.getDimensionsPoints(pageSetup.paperSize);

    if (pageSetup.orientation === PageOrientation.Landscape) {
      const swap = pageWidthPt;
      pageWidthPt = pageHeightPt;
      pageHeightPt = swap;
    }

    const leftPt = RenderUnits.inchesToPoints(pageSetup.margins.left);
    const rightPt = RenderUnits.inchesToPoints(pageSetup.margins.right);
    const topPt = RenderUnits.inchesToPoints(pageSetup.margins.top);
    const bottomPt = RenderUnits.inchesToPoints(pageSetup.margins.bottom);

    const usableWidthPt = Math.max(1, pageWidthPt - leftPt - rightPt);
    const usableHeightPt = Math.max(1, pageHeightPt - topPt - bottomPt);

    const firstRow = area.firstRow;
    const firstCol = area.firstColumn;
    const lastRow = firstRow + area.totalRows - 1;
    const lastCol = firstCol + area.totalColumns - 1;

    const totalWidthPt = layout.columnStartPt[lastCol + 1] - layout.columnStartPt[firstCol];
    const totalHeightPt = layout.rowStartPt[lastRow + 1] - layout.rowStartPt[firstRow];

    const forceOnePageWide = options.onePagePerSheet || options.allColumnsInOnePagePerSheet;
    const forceOnePageTall = options.onePagePerSheet;

    const scale = Paginator.resolveScale(
      pageSetup, options,
      totalWidthPt, totalHeightPt,
      usableWidthPt, usableHeightPt,
    );

    const colBreaks: Span[] = forceOnePageWide
      ? [{ start: firstCol, end: lastCol }]
      : Paginator.splitAxis(
          firstCol, lastCol,
          usableWidthPt / scale,
          (i) => layout.columnWidthPt[i] ?? 0,
          pageSetup.verticalPageBreaks ?? [],
        );

    const rowBreaks: Span[] = forceOnePageTall
      ? [{ start: firstRow, end: lastRow }]
      : Paginator.splitAxis(
          firstRow, lastRow,
          usableHeightPt / scale,
          (i) => layout.rowHeightPt[i] ?? 0,
          pageSetup.horizontalPageBreaks ?? [],
        );

    const pages: PageLayout[] = [];
    const pageCount = rowBreaks.length * colBreaks.length;

    for (let pageIndex = 0; pageIndex < pageCount; pageIndex++) {
      let rowIndex: number, colIndex: number;
      if (pageSetup.pageOrderOverThenDown) {
        rowIndex = Math.floor(pageIndex / colBreaks.length);
        colIndex = pageIndex % colBreaks.length;
      } else {
        rowIndex = pageIndex % rowBreaks.length;
        colIndex = Math.floor(pageIndex / rowBreaks.length);
      }
      const rowSpan = rowBreaks[rowIndex];
      const colSpan = colBreaks[colIndex];

      const page = new PageLayout();
      page.sheetIndex = sheetIndex;
      page.sheet = layout;
      page.startRow = rowSpan.start;
      page.endRow = rowSpan.end;
      page.startColumn = colSpan.start;
      page.endColumn = colSpan.end;
      page.scaleFactor = scale;
      page.pageWidthPt = pageWidthPt;
      page.pageHeightPt = pageHeightPt;
      page.contentOriginXPt = leftPt;
      page.contentOriginYPt = topPt;
      pages.push(page);
    }

    return pages;
  }

  private static resolveScale(
    pageSetup: PageSetupData,
    options: PdfSaveOptions,
    totalWidthPt: number,
    totalHeightPt: number,
    usableWidthPt: number,
    usableHeightPt: number,
  ): number {
    if (options.onePagePerSheet) {
      const w = totalWidthPt > 0 ? usableWidthPt / totalWidthPt : 1;
      const h = totalHeightPt > 0 ? usableHeightPt / totalHeightPt : 1;
      return Math.min(1, Math.min(w, h));
    }
    if (options.allColumnsInOnePagePerSheet) {
      const w = totalWidthPt > 0 ? usableWidthPt / totalWidthPt : 1;
      return Math.min(1, w);
    }
    if (pageSetup.fitToPage) {
      let scale = Infinity;
      if (pageSetup.fitToWidth && totalWidthPt > 0) {
        scale = Math.min(scale, (pageSetup.fitToWidth || 1) * usableWidthPt / totalWidthPt);
      }
      if (pageSetup.fitToHeight && totalHeightPt > 0) {
        scale = Math.min(scale, (pageSetup.fitToHeight || 1) * usableHeightPt / totalHeightPt);
      }
      if (scale === Infinity) return 1;
      return Math.min(1, scale);
    }
    if (pageSetup.scale && pageSetup.scale > 0) {
      return pageSetup.scale / 100;
    }
    return 1;
  }

  /**
   * Greedily groups indices first..last into spans that each fit within capacity.
   */
  private static splitAxis(
    first: number,
    last: number,
    capacity: number,
    size: (i: number) => number,
    manualBreaks: number[],
  ): Span[] {
    const breakSet = new Set(manualBreaks);
    const spans: Span[] = [];
    let start = first;
    let accumulated = 0;

    for (let i = first; i <= last; i++) {
      const cellSize = size(i);
      const manualBreakHere = i > start && breakSet.has(i);
      const overflow = i > start && accumulated + cellSize > capacity;

      if (manualBreakHere || overflow) {
        spans.push({ start, end: i - 1 });
        start = i;
        accumulated = 0;
      }
      accumulated += cellSize;
    }
    spans.push({ start, end: last });
    return spans;
  }
}
