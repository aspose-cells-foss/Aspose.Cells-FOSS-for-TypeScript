import { ChartType } from "./ChartType";
import { UnsupportedFeatureException } from "./UnsupportedFeatureException";

const CHART_NS_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const DRAWING_NS_URI = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL_NS_URI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/**
 * Ported from Aspose.Cells_FOSS/ChartXmlTemplates.cs.
 * Produces the chart part XML for charts created programmatically.
 */
export class ChartXmlTemplates {
  public static Build(type: ChartType, dataRange: string, chartTitle: string): string {
    switch (type) {
      case ChartType.Column:
        return ChartXmlTemplates.buildWithAxes(dataRange, chartTitle, ChartXmlTemplates.buildBarInner("col", "clustered"));
      case ChartType.Bar:
        return ChartXmlTemplates.buildWithAxes(dataRange, chartTitle, ChartXmlTemplates.buildBarInner("bar", "clustered"));
      case ChartType.Line:
        return ChartXmlTemplates.buildWithAxes(dataRange, chartTitle, ChartXmlTemplates.buildLineInner("standard"));
      case ChartType.Area:
        return ChartXmlTemplates.buildWithAxes(dataRange, chartTitle, ChartXmlTemplates.buildAreaInner("standard"));
      case ChartType.Pie:
        return ChartXmlTemplates.buildPieStyle(dataRange, chartTitle, "pieChart");
      case ChartType.Doughnut:
        return ChartXmlTemplates.buildPieStyle(dataRange, chartTitle, "doughnutChart", 50);
      case ChartType.Scatter:
        return ChartXmlTemplates.buildScatterStyle(dataRange, chartTitle);
      case ChartType.Bubble:
        return ChartXmlTemplates.buildBubbleStyle(dataRange, chartTitle);
      case ChartType.Radar:
        return ChartXmlTemplates.buildRadarStyle(dataRange, chartTitle);
      case ChartType.Stock:
        return ChartXmlTemplates.buildStockStyle(dataRange, chartTitle);
      case ChartType.Column3D:
        return ChartXmlTemplates.buildWithAxes3D(dataRange, chartTitle, ChartXmlTemplates.buildBar3DInner("col", "clustered"));
      case ChartType.Bar3D:
        return ChartXmlTemplates.buildWithAxes3D(dataRange, chartTitle, ChartXmlTemplates.buildBar3DInner("bar", "clustered"));
      case ChartType.Line3D:
        return ChartXmlTemplates.buildWithAxes3D(dataRange, chartTitle, ChartXmlTemplates.buildLine3DInner());
      case ChartType.Area3D:
        return ChartXmlTemplates.buildWithAxes3D(dataRange, chartTitle, ChartXmlTemplates.buildArea3DInner("standard"));
      case ChartType.Pie3D:
        return ChartXmlTemplates.buildPieStyle(dataRange, chartTitle, "pie3DChart");
      case ChartType.Surface3D:
        return ChartXmlTemplates.buildSurface3DStyle(dataRange, chartTitle, false);
      case ChartType.SurfaceWireframe3D:
        return ChartXmlTemplates.buildSurface3DStyle(dataRange, chartTitle, true);
      case ChartType.Contour:
        return ChartXmlTemplates.buildContourStyle(dataRange, chartTitle);
      default:
        throw new UnsupportedFeatureException(
          "Programmatic chart creation is not supported for chart type '" + type + "'. " +
            "This chart type can only be loaded from an existing XLSX file.",
        );
    }
  }

  private static header(): string {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<c:chartSpace xmlns:c="' + CHART_NS_URI + '"' +
      ' xmlns:a="' + DRAWING_NS_URI + '"' +
      ' xmlns:r="' + REL_NS_URI + '">\r\n' +
      '  <c:lang val="en-US"/>\r\n'
    );
  }

  private static footer(): string {
    return "</c:chartSpace>";
  }

  private static titleXml(chartTitle: string): string {
    if (chartTitle === null || chartTitle === undefined || chartTitle.length === 0) {
      return '    <c:autoTitleDeleted val="1"/>\r\n';
    }

    const escaped = chartTitle.split("&").join("&amp;").split("<").join("&lt;").split(">").join("&gt;");
    return (
      "    <c:title>" +
      "<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" + escaped + "</a:t></a:r></a:p></c:rich></c:tx>" +
      '<c:overlay val="0"/>' +
      "</c:title>\r\n" +
      '    <c:autoTitleDeleted val="0"/>\r\n'
    );
  }

  private static seriesWithVal(dataRange: string, idx: number = 0): string {
    return (
      "        <c:ser>" +
      '<c:idx val="' + String(idx) + '"/>' +
      '<c:order val="' + String(idx) + '"/>' +
      "<c:val><c:numRef><c:f>" + dataRange + "</c:f>" +
      "<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"0\"/></c:numCache>" +
      "</c:numRef></c:val>" +
      "</c:ser>\r\n"
    );
  }

  private static catAxValAx(): string {
    return (
      "      <c:catAx>" +
      '<c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
      '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/>' +
      "</c:catAx>\r\n" +
      "      <c:valAx>" +
      '<c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
      '<c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/>' +
      "</c:valAx>\r\n"
    );
  }

  // Series axis (depth axis) required by all 3D chart types that declare three axIds.
  // deleted=true hides the axis (used by bar3D/line3D/area3D where the depth tick marks
  // are not normally shown); deleted=false makes it visible (surface/contour charts).
  private static serAx(deleted: boolean): string {
    const deleteVal = deleted ? "1" : "0";
    return (
      "      <c:serAx>" +
      '<c:axId val="3"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
      '<c:delete val="' + deleteVal + '"/><c:axPos val="b"/><c:crossAx val="2"/>' +
      "</c:serAx>\r\n"
    );
  }

  private static valAxValAx(): string {
    return (
      "      <c:valAx>" +
      '<c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
      '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/>' +
      "</c:valAx>\r\n" +
      "      <c:valAx>" +
      '<c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
      '<c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/>' +
      "</c:valAx>\r\n"
    );
  }

  private static legend(): string {
    return '    <c:legend><c:legendPos val="b"/></c:legend>\r\n' + '    <c:plotVisOnly val="1"/>\r\n';
  }

  // --- Per-type inner element builders ---

  private static buildBarInner(barDir: string, grouping: string): string {
    return (
      "      <c:barChart>" +
      '<c:barDir val="' + barDir + '"/>' +
      '<c:grouping val="' + grouping + '"/>' +
      '<c:varyColors val="0"/>\r\n' +
      ChartXmlTemplates.seriesWithVal("__DATA__") +
      '        <c:axId val="1"/><c:axId val="2"/>' +
      "</c:barChart>\r\n"
    );
  }

  private static buildLineInner(grouping: string): string {
    return (
      "      <c:lineChart>" +
      '<c:grouping val="' + grouping + '"/>' +
      '<c:varyColors val="0"/>\r\n' +
      ChartXmlTemplates.seriesWithVal("__DATA__") +
      '        <c:axId val="1"/><c:axId val="2"/>' +
      "</c:lineChart>\r\n"
    );
  }

  private static buildAreaInner(grouping: string): string {
    return (
      "      <c:areaChart>" +
      '<c:grouping val="' + grouping + '"/>' +
      '<c:varyColors val="0"/>\r\n' +
      ChartXmlTemplates.seriesWithVal("__DATA__") +
      '        <c:axId val="1"/><c:axId val="2"/>' +
      "</c:areaChart>\r\n"
    );
  }

  private static buildBar3DInner(barDir: string, grouping: string): string {
    return (
      "      <c:bar3DChart>" +
      '<c:barDir val="' + barDir + '"/>' +
      '<c:grouping val="' + grouping + '"/>' +
      '<c:varyColors val="0"/>\r\n' +
      ChartXmlTemplates.seriesWithVal("__DATA__") +
      '        <c:axId val="1"/><c:axId val="2"/><c:axId val="3"/>' +
      "</c:bar3DChart>\r\n"
    );
  }

  private static buildLine3DInner(): string {
    return (
      "      <c:line3DChart>" +
      '<c:grouping val="standard"/>' +
      '<c:varyColors val="0"/>\r\n' +
      ChartXmlTemplates.seriesWithVal("__DATA__") +
      '        <c:axId val="1"/><c:axId val="2"/><c:axId val="3"/>' +
      "</c:line3DChart>\r\n"
    );
  }

  private static buildArea3DInner(grouping: string): string {
    return (
      "      <c:area3DChart>" +
      '<c:grouping val="' + grouping + '"/>' +
      '<c:varyColors val="0"/>\r\n' +
      ChartXmlTemplates.seriesWithVal("__DATA__") +
      '        <c:axId val="1"/><c:axId val="2"/><c:axId val="3"/>' +
      "</c:area3DChart>\r\n"
    );
  }

  // --- Full document builders ---

  private static buildWithAxes(dataRange: string, chartTitle: string, chartInner: string): string {
    const inner = chartInner.split("__DATA__").join(dataRange);
    return (
      ChartXmlTemplates.header() +
      "  <c:chart>\r\n" +
      ChartXmlTemplates.titleXml(chartTitle) +
      "    <c:plotArea>\r\n" +
      "      <c:layout/>\r\n" +
      inner +
      ChartXmlTemplates.catAxValAx() +
      "    </c:plotArea>\r\n" +
      ChartXmlTemplates.legend() +
      "  </c:chart>\r\n" +
      ChartXmlTemplates.footer()
    );
  }

  private static buildWithAxes3D(dataRange: string, chartTitle: string, chartInner: string): string {
    const inner = chartInner.split("__DATA__").join(dataRange);
    return (
      ChartXmlTemplates.header() +
      "  <c:chart>\r\n" +
      ChartXmlTemplates.titleXml(chartTitle) +
      '    <c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:perspective val="30"/></c:view3D>\r\n' +
      "    <c:plotArea>\r\n" +
      "      <c:layout/>\r\n" +
      inner +
      ChartXmlTemplates.catAxValAx() +
      ChartXmlTemplates.serAx(true) +
      "    </c:plotArea>\r\n" +
      ChartXmlTemplates.legend() +
      "  </c:chart>\r\n" +
      ChartXmlTemplates.footer()
    );
  }

  private static buildPieStyle(dataRange: string, chartTitle: string, elementName: string, holeSize: number = 0): string {
    const holeSizeAttr = holeSize > 0 ? ' <c:holeSize val="' + String(holeSize) + '"/>' : "";
    return (
      ChartXmlTemplates.header() +
      "  <c:chart>\r\n" +
      ChartXmlTemplates.titleXml(chartTitle) +
      "    <c:plotArea>\r\n" +
      "      <c:layout/>\r\n" +
      "      <c:" + elementName + ">" +
      '<c:varyColors val="1"/>\r\n' +
      "        <c:ser>" +
      '<c:idx val="0"/><c:order val="0"/>' +
      "<c:val><c:numRef><c:f>" + dataRange + "</c:f>" +
      "<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"0\"/></c:numCache>" +
      "</c:numRef></c:val>" +
      "</c:ser>\r\n" +
      holeSizeAttr +
      "      </c:" + elementName + ">\r\n" +
      "    </c:plotArea>\r\n" +
      ChartXmlTemplates.legend() +
      "  </c:chart>\r\n" +
      ChartXmlTemplates.footer()
    );
  }

  private static buildScatterStyle(dataRange: string, chartTitle: string): string {
    return (
      ChartXmlTemplates.header() +
      "  <c:chart>\r\n" +
      ChartXmlTemplates.titleXml(chartTitle) +
      "    <c:plotArea>\r\n" +
      "      <c:layout/>\r\n" +
      "      <c:scatterChart>" +
      '<c:scatterStyle val="marker"/>' +
      '<c:varyColors val="0"/>\r\n' +
      "        <c:ser>" +
      '<c:idx val="0"/><c:order val="0"/>' +
      "<c:yVal><c:numRef><c:f>" + dataRange + "</c:f>" +
      "<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"0\"/></c:numCache>" +
      "</c:numRef></c:yVal>" +
      "</c:ser>\r\n" +
      '        <c:axId val="1"/><c:axId val="2"/>' +
      "</c:scatterChart>\r\n" +
      ChartXmlTemplates.valAxValAx() +
      "    </c:plotArea>\r\n" +
      ChartXmlTemplates.legend() +
      "  </c:chart>\r\n" +
      ChartXmlTemplates.footer()
    );
  }

  private static buildBubbleStyle(dataRange: string, chartTitle: string): string {
    return (
      ChartXmlTemplates.header() +
      "  <c:chart>\r\n" +
      ChartXmlTemplates.titleXml(chartTitle) +
      "    <c:plotArea>\r\n" +
      "      <c:layout/>\r\n" +
      "      <c:bubbleChart>" +
      '<c:varyColors val="0"/>' +
      '<c:showNegBubbles val="0"/>\r\n' +
      "        <c:ser>" +
      '<c:idx val="0"/><c:order val="0"/>' +
      "<c:yVal><c:numRef><c:f>" + dataRange + "</c:f>" +
      "<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"0\"/></c:numCache>" +
      "</c:numRef></c:yVal>" +
      "<c:bubbleSize><c:numRef><c:f>" + dataRange + "</c:f>" +
      "<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"0\"/></c:numCache>" +
      "</c:numRef></c:bubbleSize>" +
      "</c:ser>\r\n" +
      '        <c:axId val="1"/><c:axId val="2"/>' +
      "</c:bubbleChart>\r\n" +
      ChartXmlTemplates.valAxValAx() +
      "    </c:plotArea>\r\n" +
      ChartXmlTemplates.legend() +
      "  </c:chart>\r\n" +
      ChartXmlTemplates.footer()
    );
  }

  private static buildRadarStyle(dataRange: string, chartTitle: string): string {
    return (
      ChartXmlTemplates.header() +
      "  <c:chart>\r\n" +
      ChartXmlTemplates.titleXml(chartTitle) +
      "    <c:plotArea>\r\n" +
      "      <c:layout/>\r\n" +
      "      <c:radarChart>" +
      '<c:radarStyle val="marker"/>' +
      '<c:varyColors val="0"/>\r\n' +
      ChartXmlTemplates.seriesWithVal(dataRange) +
      '        <c:axId val="1"/><c:axId val="2"/>' +
      "</c:radarChart>\r\n" +
      ChartXmlTemplates.catAxValAx() +
      "    </c:plotArea>\r\n" +
      ChartXmlTemplates.legend() +
      "  </c:chart>\r\n" +
      ChartXmlTemplates.footer()
    );
  }

  private static buildStockStyle(dataRange: string, chartTitle: string): string {
    // Stock chart requires 4 series (Open, High, Low, Close)
    let result = "";
    result += ChartXmlTemplates.header();
    result += "  <c:chart>\r\n";
    result += ChartXmlTemplates.titleXml(chartTitle);
    result += "    <c:plotArea>\r\n";
    result += "      <c:layout/>\r\n";
    result += "      <c:stockChart>\r\n";
    for (let index = 0; index < 4; index++) {
      result += ChartXmlTemplates.seriesWithVal(dataRange, index);
    }

    result += '        <c:axId val="1"/><c:axId val="2"/>';
    result += "</c:stockChart>\r\n";
    result += ChartXmlTemplates.catAxValAx();
    result += "    </c:plotArea>\r\n";
    result += ChartXmlTemplates.legend();
    result += "  </c:chart>\r\n";
    result += ChartXmlTemplates.footer();
    return result;
  }

  private static buildSurface3DStyle(dataRange: string, chartTitle: string, wireframe: boolean): string {
    const wireframeAttr = wireframe ? '<c:wireframe val="1"/>' : '<c:wireframe val="0"/>';
    return (
      ChartXmlTemplates.header() +
      "  <c:chart>\r\n" +
      ChartXmlTemplates.titleXml(chartTitle) +
      '    <c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:perspective val="30"/></c:view3D>\r\n' +
      "    <c:plotArea>\r\n" +
      "      <c:layout/>\r\n" +
      "      <c:surface3DChart>" +
      wireframeAttr +
      '<c:varyColors val="0"/>\r\n' +
      ChartXmlTemplates.seriesWithVal(dataRange) +
      '        <c:axId val="1"/><c:axId val="2"/><c:axId val="3"/>' +
      "</c:surface3DChart>\r\n" +
      ChartXmlTemplates.catAxValAx() +
      ChartXmlTemplates.serAx(false) +
      "    </c:plotArea>\r\n" +
      ChartXmlTemplates.legend() +
      "  </c:chart>\r\n" +
      ChartXmlTemplates.footer()
    );
  }

  private static buildContourStyle(dataRange: string, chartTitle: string): string {
    return (
      ChartXmlTemplates.header() +
      "  <c:chart>\r\n" +
      ChartXmlTemplates.titleXml(chartTitle) +
      "    <c:plotArea>\r\n" +
      "      <c:layout/>\r\n" +
      "      <c:surfaceChart>" +
      '<c:wireframe val="0"/>' +
      ChartXmlTemplates.seriesWithVal(dataRange) +
      '        <c:axId val="1"/><c:axId val="2"/><c:axId val="3"/>' +
      "</c:surfaceChart>\r\n" +
      ChartXmlTemplates.catAxValAx() +
      ChartXmlTemplates.serAx(false) +
      "    </c:plotArea>\r\n" +
      ChartXmlTemplates.legend() +
      "  </c:chart>\r\n" +
      ChartXmlTemplates.footer()
    );
  }
}
