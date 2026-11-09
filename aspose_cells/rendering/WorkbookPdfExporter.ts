import type { LegacyWorkbook, LegacyWorksheet } from "../index";
import { RenderContext } from "./RenderContext";
import { SheetLayout } from "./SheetLayout";
import { Paginator } from "./Paginator";
import { PageLayout } from "./PageLayout";
import { PaperSizes } from "./PaperSizes";
import { PdfSaveOptions } from "../types";
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
    workbook: LegacyWorkbook,
    outputPath: string,
    options?: PdfSaveOptions,
  ): Promise<void> {
    const opts = options ?? new PdfSaveOptions();
    const context = new RenderContext(opts.DefaultFont);

    if (opts.ConditionalFormattingEvaluationDate) {
      context.conditionalFormattingEvaluationDate = opts.ConditionalFormattingEvaluationDate;
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
      // A valid PDF needs at least one page - emit blank Letter page
      const { widthPt, heightPt } = PaperSizes.getDimensionsPoints(1);
      doc.addPage({ size: [widthPt, heightPt] });
    }

    const renderer = new SheetPdfRenderer(doc as any, context, workbook.worksheets.styles);
    for (const page of pages) {
      doc.addPage({ size: [page.pageWidthPt, page.pageHeightPt] });
      renderer.renderPage(page);
    }

    doc.end();

    return completion;
  }

  private static buildPages(
    workbook: LegacyWorkbook,
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

  private static extractPageSetup(_sheet: LegacyWorksheet) {
    // Defaults must match what the HTML/XLSX writers emit for this codebase:
    // <pageSetup orientation="portrait" paperSize="1"/> and
    // <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
    return {
      paperSize: 1, // Letter
      orientation: 1, // Portrait
      margins: {
        left: 0.75, right: 0.75, top: 1, bottom: 1,
        header: 0.5, footer: 0.5,
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
