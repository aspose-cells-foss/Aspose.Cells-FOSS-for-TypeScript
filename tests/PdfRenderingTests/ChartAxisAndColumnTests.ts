import { guardedTest, RASTER_REASON } from "./Support";

/**
 * Port of Aspose.Cells_FOSS.PdfRenderingTests.ChartAxisAndColumnTests.
 *
 * These checks compare parsed chart axes and column geometry against pixels
 * measured from the Excel reference. That requires the .NET chart parser and
 * SkiaSharp rasterizer, neither of which exists in the TypeScript port.
 */
export class ChartAxisAndColumnTests {
  /** Legacy value axes and explicit axis overrides. */
  static LegacyValueAxes(): void {
    guardedTest(["0011.xlsx"], RASTER_REASON);
  }

  /** Column borders separate series that touch at their edges. */
  static ColumnBorders(): void {
    guardedTest(["0011.xlsx"], RASTER_REASON);
  }

  /** League columns keep the value-axis gap from the Excel reference. */
  static LeagueAxisGap(): void {
    guardedTest(["0011.xlsx"], RASTER_REASON);
  }
}
