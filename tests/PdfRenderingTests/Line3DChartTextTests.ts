import { skipTest } from "../Shared/TestInfrastructure";
import { CHART_INTERNALS_REASON, requireInputFile } from "./Support";

/**
 * Port of Aspose.Cells_FOSS.PdfRenderingTests.Line3DChartTextTests.
 *
 * The upstream renders pages into SkiaSharp bitmaps and inspects PDF text
 * session runs, page clips and embedded font programs. These internals are
 * specific to the .NET renderer.
 */
export class Line3DChartTextTests {
  static Verify(input: string, _pathOptions: unknown, _textOptions: unknown): void {
    requireInputFile(input);
    skipTest(CHART_INTERNALS_REASON);
  }
}
