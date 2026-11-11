import { skipTest } from "../Shared/TestInfrastructure";
import { guardedTest, RASTER_REASON } from "./Support";

/**
 * Port of Aspose.Cells_FOSS.PdfRenderingTests.ChartDetailRenderingTests.
 *
 * The upstream checks inspect rasterized chart pixels using SkiaSharp. The
 * TypeScript renderer emits PDF text and vector operators, so the fidelity
 * checks cannot be reproduced without the .NET rasterizer.
 */
export class ChartDetailRenderingTests {
  /** Stacked columns with zero and near-zero segments keep distinct slot geometry. */
  static StackedZeroAndTinySegments(): void {
    skipTest(RASTER_REASON);
  }

  /** Multi-line category labels keep their intended wrap inside the plot band. */
  static MultilineCategoryLabels(): void {
    guardedTest(["0011.xlsx"], RASTER_REASON);
  }

  /** Value axis headroom and gridline placement for fixture 0011. */
  static AxisHeadroomAndGridlines(): void {
    guardedTest(["0011.xlsx"], RASTER_REASON);
  }
}
