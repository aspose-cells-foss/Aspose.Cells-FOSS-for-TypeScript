import { SaveFormat } from "../../aspose_cells/enums";
import { PdfExportAdapter } from "../../aspose_cells/PdfExportAdapter";
import { Paginator } from "../../aspose_cells/rendering/Paginator";
import { PrintAreaRanges } from "../../aspose_cells/rendering/PrintAreaRanges";
import { SheetLayout } from "../../aspose_cells/rendering/SheetLayout";
import { WorksheetDataExtractor } from "../../aspose_cells/rendering/WorksheetDataExtractor";
import { Workbook } from "../../aspose_cells/workbook";
import { AssertEx, MemoryStream } from "../Shared/TestInfrastructure";
import { guardedTest, requireInput, RASTER_REASON } from "./Support";

/**
 * Port of Aspose.Cells_FOSS.PdfRenderingTests.PrintAreaRenderingTests.
 */
export class PrintAreaRenderingTests {
  /**
   * Pagination treats every print-area range independently, keeps the explicit
   * scale when fit-to-width is set, and clamps the second range to its last row.
   */
  static RangesAndScalingModes(): void {
    const workbook = new Workbook();
    const sheet = workbook.Worksheets[0];
    sheet.Cells.get(0, 40).PutValue("Outside print area");
    sheet.PageSetup.PrintArea = "B2:D5,F10:G12";
    sheet.PageSetup.FitToPagesWide = 1;
    sheet.PageSetup.Scale = 69;

    let stream = new MemoryStream();
    workbook.Save(stream, SaveFormat.Xlsx);
    let loaded = new Workbook(stream);
    let pages = PrintAreaRenderingTests.paginate(loaded);
    AssertEx.equal(2, pages.length);
    AssertEx.equal(1, pages[0].startColumn);
    AssertEx.equal(3, pages[0].endColumn);
    AssertEx.equal(1, pages[0].startRow);
    AssertEx.equal(4, pages[0].endRow);
    AssertEx.equal(9, pages[1].startRow);
    AssertEx.equal(6, pages[1].endColumn);
    AssertEx.equal(0.69, pages[0].scaleFactor);

    sheet.PageSetup.FitToPagesWide = 1;
    AssertEx.true(workbook.Model.Worksheets[0].PageSetup.FitToPage);
    sheet.PageSetup.PrintArea = "B2:D50";

    stream = new MemoryStream();
    workbook.Save(stream, SaveFormat.Xlsx);
    loaded = new Workbook(stream);
    AssertEx.true(loaded.Model.Worksheets[0].PageSetup.FitToPage);
    pages = PrintAreaRenderingTests.paginate(loaded);
    AssertEx.equal(1, pages.length);
    AssertEx.equal(49, pages[0].endRow);
  }

  /** Fixture 0011 print layout: page slices and row heights. */
  static FixtureLayout(): void {
    requireInput("0011.xlsx");
    guardedTest([], RASTER_REASON);
  }

  /** Fixture 0011 chart formatting: point colors, data labels and side legends. */
  static ChartFormatting(): void {
    requireInput("0011.xlsx");
    guardedTest([], RASTER_REASON);
  }

  private static paginate(workbook: Workbook) {
    const legacy = PdfExportAdapter.toLegacyWorkbook(workbook.Model);
    const worksheet = legacy.worksheets.get(0);
    if (!worksheet) {
      throw new Error("Expected the workbook to contain a worksheet.");
    }

    const pageSetup = (worksheet as unknown as { __pdfPageSetup: { printArea: string | null } }).__pdfPageSetup;
    const layout = SheetLayout.build(WorksheetDataExtractor.extract(worksheet));
    return Paginator.paginate(
      layout,
      0,
      { OnePagePerSheet: false, AllColumnsInOnePagePerSheet: false } as never,
      pageSetup as never,
      PrintAreaRanges.read(pageSetup.printArea ?? undefined),
    );
  }
}
