import * as path from "path";
import { ChartType, SaveFormat, Workbook } from "../../aspose_cells";
import { ZipReader } from "../../aspose_cells/Packaging/Archive";
import { AssertEx, getOrCreateOutputDirectory, MemoryStream, TestCase, TestRunner } from "../Shared/TestInfrastructure";
import { ChartScenarioFactory } from "../Shared/ChartScenarioFactory";

const DATA_RANGE = "Sheet1!$B$1:$B$5";

function columnChartSavesBarElementWithColDirection(): void {
  const chartXml = getChartXml(ChartType.Column);
  AssertEx.contains("<c:barChart", chartXml, "Expected c:barChart element");
  AssertEx.contains('barDir val="col"', chartXml, "Expected barDir=col for column chart");
}

function barChartSavesBarElementWithBarDirection(): void {
  const chartXml = getChartXml(ChartType.Bar);
  AssertEx.contains("<c:barChart", chartXml, "Expected c:barChart element");
  AssertEx.contains('barDir val="bar"', chartXml, "Expected barDir=bar for bar chart");
}

function lineChartSavesLineChartElement(): void {
  AssertEx.contains("<c:lineChart", getChartXml(ChartType.Line), "Expected c:lineChart element");
}

function areaChartSavesAreaChartElement(): void {
  AssertEx.contains("<c:areaChart", getChartXml(ChartType.Area), "Expected c:areaChart element");
}

function pieChartSavesPieChartElement(): void {
  AssertEx.contains("<c:pieChart", getChartXml(ChartType.Pie), "Expected c:pieChart element");
}

function doughnutChartSavesDoughnutChartElement(): void {
  AssertEx.contains("<c:doughnutChart", getChartXml(ChartType.Doughnut), "Expected c:doughnutChart element");
}

function scatterChartSavesScatterChartElement(): void {
  AssertEx.contains("<c:scatterChart", getChartXml(ChartType.Scatter), "Expected c:scatterChart element");
}

function radarChartSavesRadarChartElement(): void {
  AssertEx.contains("<c:radarChart", getChartXml(ChartType.Radar), "Expected c:radarChart element");
}

function bubbleChartSavesBubbleChartElement(): void {
  AssertEx.contains("<c:bubbleChart", getChartXml(ChartType.Bubble), "Expected c:bubbleChart element");
}

function stockChartSavesStockChartElement(): void {
  AssertEx.contains("<c:stockChart", getChartXml(ChartType.Stock), "Expected c:stockChart element");
}

function column3DChartSavesBar3DElementWithCol(): void {
  const chartXml = getChartXml(ChartType.Column3D);
  AssertEx.contains("<c:bar3DChart", chartXml, "Expected c:bar3DChart element");
  AssertEx.contains('barDir val="col"', chartXml, "Expected barDir=col for column3D chart");
}

function bar3DChartSavesBar3DElementWithBar(): void {
  const chartXml = getChartXml(ChartType.Bar3D);
  AssertEx.contains("<c:bar3DChart", chartXml, "Expected c:bar3DChart element");
  AssertEx.contains('barDir val="bar"', chartXml, "Expected barDir=bar for bar3D chart");
}

function line3DChartSavesLine3DElement(): void {
  AssertEx.contains("<c:line3DChart", getChartXml(ChartType.Line3D), "Expected c:line3DChart element");
}

function area3DChartSavesArea3DElement(): void {
  AssertEx.contains("<c:area3DChart", getChartXml(ChartType.Area3D), "Expected c:area3DChart element");
}

function pie3DChartSavesPie3DElement(): void {
  AssertEx.contains("<c:pie3DChart", getChartXml(ChartType.Pie3D), "Expected c:pie3DChart element");
}

function surface3DChartSavesSurface3DElement(): void {
  AssertEx.contains("<c:surface3DChart", getChartXml(ChartType.Surface3D), "Expected c:surface3DChart element");
}

function contourChartSavesSurfaceElement(): void {
  const chartXml = getChartXml(ChartType.Contour);
  AssertEx.contains("<c:surfaceChart", chartXml, "Expected c:surfaceChart element");

  const surfaceChartStart = chartXml.indexOf("<c:surfaceChart>");
  const surfaceChartEnd = chartXml.indexOf("</c:surfaceChart>");
  AssertEx.true(surfaceChartStart >= 0 && surfaceChartEnd > surfaceChartStart, "Expected c:surfaceChart block");

  const surfaceChartXml = chartXml.substring(surfaceChartStart, surfaceChartEnd);
  AssertEx.true(surfaceChartXml.indexOf("<c:axId") >= 0, "Expected c:surfaceChart to contain c:axId elements");
  AssertEx.true(
    surfaceChartXml.indexOf("<c:varyColors") < 0,
    "Contour surfaceChart must not serialize c:varyColors.",
  );
}

function chartDrawingXmlContainsGraphicFrame(): void {
  const stream = saveChartWorkbook(ChartType.Column);
  const drawingXml = readZipEntry(stream, "xl/drawings/drawing1.xml");
  AssertEx.contains("graphicFrame", drawingXml, "Drawing XML must contain graphicFrame");
  AssertEx.contains(
    "http://schemas.openxmlformats.org/drawingml/2006/chart",
    drawingXml,
    "Drawing XML must reference chart namespace URI",
  );
}

function chartContentTypesIncludesChartOverride(): void {
  const stream = saveChartWorkbook(ChartType.Line);
  const contentTypes = readZipEntry(stream, "[Content_Types].xml");
  AssertEx.contains("/xl/charts/chart1.xml", contentTypes, "Content_Types.xml must have chart1.xml override");
  AssertEx.contains(
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
    contentTypes,
    "Content_Types.xml must have chart content type",
  );
}

function chartDrawingRelsLinksToChartFile(): void {
  const stream = saveChartWorkbook(ChartType.Pie);
  const relsXml = readZipEntry(stream, "xl/drawings/_rels/drawing1.xml.rels");
  AssertEx.contains("../charts/chart1.xml", relsXml, "Drawing rels must reference ../charts/chart1.xml");
  AssertEx.contains(
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
    relsXml,
    "Drawing rels must use chart relationship type",
  );
}

function multipleChartsOnSameSheetAllSaved(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  sheet.Charts.Add(ChartType.Column, "Sheet1!$B$1:$B$5", 0, 0, 10, 8);
  sheet.Charts.Add(ChartType.Line, "Sheet1!$C$1:$C$5", 12, 0, 22, 8);

  const stream = new MemoryStream();
  workbook.Save(stream, SaveFormat.Xlsx);
  const reloaded = new Workbook(stream);
  AssertEx.equal(2, reloaded.Worksheets[0].Charts.Count, "Both charts must survive round-trip");
  AssertEx.equal(ChartType.Column, reloaded.Worksheets[0].Charts[0].ChartType);
  AssertEx.equal(ChartType.Line, reloaded.Worksheets[0].Charts[1].ChartType);

  const contentTypes = readZipEntry(stream, "[Content_Types].xml");
  AssertEx.contains("/xl/charts/chart1.xml", contentTypes);
  AssertEx.contains("/xl/charts/chart2.xml", contentTypes);
}

function chartPreservesDataRangeInXml(): void {
  const workbook = ChartScenarioFactory.CreateChartWorkbook(ChartType.Column);
  const stream = new MemoryStream();
  workbook.Save(stream, SaveFormat.Xlsx);
  const chartXml = readZipEntry(stream, "xl/charts/chart1.xml");
  AssertEx.contains(DATA_RANGE, chartXml, "Chart XML must preserve the data range formula");
}

function getChartXml(type: ChartType): string {
  return readZipEntry(saveChartWorkbook(type), "xl/charts/chart1.xml");
}

function saveChartWorkbook(type: ChartType): MemoryStream {
  const workbook = ChartScenarioFactory.CreateChartWorkbook(type);
  const outputDir = getOrCreateOutputDirectory("chart-openxml-roundtrips");
  workbook.Save(path.join(outputDir, chartTypeName(type) + ".xlsx"));

  const stream = new MemoryStream();
  workbook.Save(stream, SaveFormat.Xlsx);
  return stream;
}

function chartTypeName(type: ChartType): string {
  return String((ChartType as unknown as Record<number, string>)[type] ?? type);
}

function readZipEntry(stream: MemoryStream, entryPath: string): string {
  const archive = new ZipReader(stream.toBuffer());
  const normalized = entryPath.replace(/^\/+/, "");
  if (!archive.hasEntry(normalized)) {
    throw new Error("Missing zip entry: " + entryPath);
  }

  return archive.readText(normalized);
}

TestRunner.runAndExit("Chart.OpenXmlTests", [
  new TestCase("column_chart_saves_bar_element_with_col_direction", columnChartSavesBarElementWithColDirection),
  new TestCase("bar_chart_saves_bar_element_with_bar_direction", barChartSavesBarElementWithBarDirection),
  new TestCase("line_chart_saves_line_chart_element", lineChartSavesLineChartElement),
  new TestCase("area_chart_saves_area_chart_element", areaChartSavesAreaChartElement),
  new TestCase("pie_chart_saves_pie_chart_element", pieChartSavesPieChartElement),
  new TestCase("doughnut_chart_saves_doughnut_chart_element", doughnutChartSavesDoughnutChartElement),
  new TestCase("scatter_chart_saves_scatter_chart_element", scatterChartSavesScatterChartElement),
  new TestCase("radar_chart_saves_radar_chart_element", radarChartSavesRadarChartElement),
  new TestCase("bubble_chart_saves_bubble_chart_element", bubbleChartSavesBubbleChartElement),
  new TestCase("stock_chart_saves_stock_chart_element", stockChartSavesStockChartElement),
  new TestCase("column3d_chart_saves_bar3d_element_with_col", column3DChartSavesBar3DElementWithCol),
  new TestCase("bar3d_chart_saves_bar3d_element_with_bar", bar3DChartSavesBar3DElementWithBar),
  new TestCase("line3d_chart_saves_line3d_element", line3DChartSavesLine3DElement),
  new TestCase("area3d_chart_saves_area3d_element", area3DChartSavesArea3DElement),
  new TestCase("pie3d_chart_saves_pie3d_element", pie3DChartSavesPie3DElement),
  new TestCase("surface3d_chart_saves_surface3d_element", surface3DChartSavesSurface3DElement),
  new TestCase("contour_chart_saves_surface_element", contourChartSavesSurfaceElement),
  new TestCase("chart_drawing_xml_contains_graphic_frame", chartDrawingXmlContainsGraphicFrame),
  new TestCase("chart_content_types_includes_chart_override", chartContentTypesIncludesChartOverride),
  new TestCase("chart_drawing_rels_links_to_chart_file", chartDrawingRelsLinksToChartFile),
  new TestCase("multiple_charts_on_same_sheet_all_saved", multipleChartsOnSameSheetAllSaved),
  new TestCase("chart_preserves_data_range_in_xml", chartPreservesDataRangeInXml),
]);
