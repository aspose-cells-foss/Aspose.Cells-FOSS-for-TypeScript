import type { Workbook, Worksheet } from "../index";
import { RenderContext } from "./RenderContext";
import { SheetLayout } from "./SheetLayout";
import { Paginator } from "./Paginator";
import { PageLayout } from "./PageLayout";
import { PaperSizes } from "./PaperSizes";
import { PdfSaveOptions } from "./types";
import { WorksheetDataExtractor } from "./WorksheetDataExtractor";
import { PrintAreaRanges } from "./PrintAreaRanges";
import { PdfTextDocumentSession } from "./PdfTextDocumentSession";
import { SheetPdfRenderer } from "./SheetPdfRenderer";

/**
 * Entry point for XLSX-to-PDF export. Orchestrates layout, pagination, and rendering
 * for every visible worksheet using pdfkit.
 */
export class WorkbookPdfExporter {
  static async export(
    workbook: Workbook,
    outputPath: string,
    options?: PdfSaveOptions,
  ): Promise<void> {
    const opts = options ?? new PdfSaveOptions();
    const context = new RenderContext(opts.defaultFont);

    if (opts.conditionalFormattingEvaluationDate) {
      context.conditionalFormattingEvaluationDate = opts.conditionalFormattingEvaluationDate;
    }

    // Always enable text optimization for smaller PDFs
    context.pdfTextSession = new PdfTextDocumentSession();
    context.enableWorksheetTextOptimization = true;

    const pages = WorkbookPdfExporter.buildPages(workbook, context, opts);

    for (let i = 0; i < pages.length; i++) {
      pages[i].pageNumber = i + 1;
      pages[i].totalPages = pages.length;
    }

    const fs = await import("fs");
    const PDFDocument = (await import("pdfkit")).default;

    const doc = new PDFDocument({
      autoFirstPage: false,
      info: {
        Creator: "Aspose.Cells FOSS (TypeScript)",
        Producer: "Aspose.Cells FOSS (pdfkit)",
      },
    });

    const stream = fs.createWriteStream(outputPath);

    // 先建立完成通知，避免 doc/stream 提前触发 error 或 finish 时 Promise 永不结算
    const completion = new Promise<void>((resolve, reject) => {
      stream.on("finish", () => resolve());
      stream.on("error", reject);
      doc.on("error", reject);
    });

    doc.pipe(stream);

    if (pages.length === 0) {
      // A valid PDF needs at least one page - emit blank A4
      const { widthPt, heightPt } = PaperSizes.getDimensionsPoints(9);
      doc.addPage({ size: [widthPt, heightPt] });
    }

    const renderer = new SheetPdfRenderer(doc as any, context);
    for (const page of pages) {
      doc.addPage({ size: [page.pageWidthPt, page.pageHeightPt] });
      renderer.renderPage(page);
    }

    doc.end();

    return completion;
  }

  private static buildPages(
    workbook: Workbook,
    context: RenderContext,
    options: PdfSaveOptions,
  ): PageLayout[] {
    const pages: PageLayout[] = [];
    const sheets = workbook.worksheets;

    for (let sheetIndex = 0; sheetIndex < sheets.length; sheetIndex++) {
      const sheet = sheets.get(sheetIndex);
      if (!sheet) continue;

      const sheetData = WorksheetDataExtractor.extract(sheet);
      const layout = SheetLayout.build(sheetData);

      if (layout.isEmpty) continue;

      const pageSetup = WorkbookPdfExporter.extractPageSetup(sheet);
      const printAreas = PrintAreaRanges.read(pageSetup.printArea);

      const sheetPages = Paginator.paginate(
        layout, sheetIndex, options, pageSetup, printAreas,
      );
      pages.push(...sheetPages);
    }

    return pages;
  }

  private static extractPageSetup(_sheet: Worksheet) {
    return {
      paperSize: 9, // A4
      orientation: 1, // Portrait
      margins: {
        left: 0.7, right: 0.7, top: 0.75, bottom: 0.75,
        header: 0.3, footer: 0.3,
      },
      printArea: undefined,
      scale: undefined,
      fitToPage: false,
      fitToWidth: undefined,
      fitToHeight: undefined,
      horizontalPageBreaks: [] as number[],
      verticalPageBreaks: [] as number[],
    };
  }
}
