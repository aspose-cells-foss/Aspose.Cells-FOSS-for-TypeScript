import { skipTest } from "../Shared/TestInfrastructure";
import { CHART_INTERNALS_REASON, requireInputFile } from "./Support";

/**
 * Port of Aspose.Cells_FOSS.PdfRenderingTests.Line3DCategoryChartTests.
 *
 * The upstream parses RawChartXml with ChartXmlParser, projects a 3D camera
 * and inspects rendered ribbons. The TypeScript port has neither the chart
 * XML parser entry point nor the Line3DProjection geometry.
 */
export class Line3DCategoryChartTests {
  static Verify(input: string): void {
    requireInputFile(input);
    skipTest(CHART_INTERNALS_REASON);
  }
}
