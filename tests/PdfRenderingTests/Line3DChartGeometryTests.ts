import { skipTest } from "../Shared/TestInfrastructure";
import { CHART_INTERNALS_REASON } from "./Support";

/**
 * Port of Aspose.Cells_FOSS.PdfRenderingTests.Line3DChartGeometryTests.
 *
 * The upstream verifies the projected camera and ribbon topology using
 * Line3DProjection and SkiaSharp points. Those types are not part of the
 * TypeScript port, so the geometry checks are reported as skipped.
 */
export class Line3DChartGeometryTests {
  static Verify(_chart: unknown): void {
    skipTest(CHART_INTERNALS_REASON);
  }
}
