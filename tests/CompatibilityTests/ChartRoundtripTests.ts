import * as path from "path";
import { Workbook } from "../../aspose_cells";
import {
  XmlNode,
  attribute,
  descendants,
  documentRoot,
  elementChildren,
  firstElement,
  parseXml,
} from "../../aspose_cells/Xml/XmlReader";
import {
  AssertEx,
  TemporaryDirectory,
  ZipPackageHelper,
  getOrCreateOutputDirectory,
} from "../Shared/TestInfrastructure";
import {
  appendOutSuffix,
  childLocalNames,
  findElementIndex,
  listXlsxFiles,
  resolveRepositoryFile,
} from "./CompatSupport";

function roundtripAllChartFilesInDirectory(subDirectory: string, expectAtLeastOneChart: boolean): void {
  const inputDir = resolveRepositoryFile("Input", subDirectory);
  const files = listXlsxFiles(inputDir);
  const outputDir = getOrCreateOutputDirectory(path.join("chart-roundtrips", subDirectory));

  for (const filePath of files) {
    const fileName = path.basename(filePath);
    const original = new Workbook(filePath);
    const originalChartCount = original.Worksheets.get(0).Charts.Count;

    if (expectAtLeastOneChart) {
      AssertEx.true(originalChartCount > 0, "Expected at least one chart in: " + fileName);
    }

    const outputPath = path.join(outputDir, appendOutSuffix(fileName));
    original.Save(outputPath);

    const reloaded = new Workbook(outputPath);
    AssertEx.equal(
      originalChartCount,
      reloaded.Worksheets.get(0).Charts.Count,
      "Chart count changed after round-trip for: " + fileName,
    );
  }
}

export function excelInputAreaChartRoundtrip(): void {
  roundtripAllChartFilesInDirectory("AreaChart", true);
}

export function excelInputBarChartRoundtrip(): void {
  roundtripAllChartFilesInDirectory("BarChart", true);
}

export function excelInputColumnChartRoundtrip(): void {
  roundtripAllChartFilesInDirectory("ColumnChart", true);
}

export function excelInputLineChartRoundtrip(): void {
  roundtripAllChartFilesInDirectory("linechart", true);
}

export function excelInputPieChartRoundtrip(): void {
  roundtripAllChartFilesInDirectory("PieChart", true);
}

export function excelInputXyChartRoundtrip(): void {
  roundtripAllChartFilesInDirectory("XYChart", true);
}

export function excelInputRadarChartRoundtrip(): void {
  roundtripAllChartFilesInDirectory("RadarChart", true);
}

export function excelInputSurfaceChartRoundtrip(): void {
  roundtripAllChartFilesInDirectory("SurfaceChart", true);
}

export function excelInputSurfaceContourRoundtripRemovesInvalidSurfaceChartVaryColors(): void {
  const inputPath = resolveRepositoryFile("Input", path.join("SurfaceChart", "surfacecontour.xlsx"));
  const workbook = new Workbook(inputPath);

  const temp = new TemporaryDirectory(path.join("chart-roundtrips", "surfacechart-normalized"));
  const roundtripPath = temp.getPath("surfacecontourout.xlsx");
  workbook.Save(roundtripPath);

  const chartXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/charts/chart1.xml");
  const chartRoot = documentRoot(parseXml(chartXml));
  AssertEx.notNull(chartRoot, "chart1.xml must have a root element.");
  const surfaceCharts = descendants(chartRoot, "surfaceChart");
  AssertEx.true(surfaceCharts.length > 0, "surfacecontour roundtrip must contain c:surfaceChart.");
  AssertEx.null(
    firstElement(surfaceCharts[0], "varyColors"),
    "surfacecontour roundtrip must remove invalid c:surfaceChart/c:varyColors.",
  );
}

export function excelInputStockChartRoundtrip(): void {
  roundtripAllChartFilesInDirectory("StockChart", true);
}

export function excelInputFunnelChartRoundtrip(): void {
  roundtripAllChartFilesInDirectory("FunnelChart", false);
}

export function excelInputHistogramChartRoundtrip(): void {
  roundtripAllChartFilesInDirectory("HistogramChart", false);
}

export function excelInputBoxChartRoundtrip(): void {
  roundtripAllChartFilesInDirectory("BoxChart", false);
}

export function excelInputSunburstChartRoundtrip(): void {
  roundtripAllChartFilesInDirectory("SunburstChart", false);
}

export function excelInputTreemapChartRoundtrip(): void {
  roundtripAllChartFilesInDirectory("TreemapChart", false);
}

export function excelInputMapChartRoundtrip(): void {
  roundtripAllChartFilesInDirectory("MapChart", false);
}

export function excelInputMapChartSamplesRoundtripRemovesInvalidChartexValueColors(): void {
  const inputPath = resolveRepositoryFile("Input", path.join("MapChart", "Map Chart samples.xlsx"));
  const workbook = new Workbook(inputPath);

  const temp = new TemporaryDirectory(path.join("chart-roundtrips", "mapchart-normalized"));
  const roundtripPath = temp.getPath("Map Chart samplesout.xlsx");
  workbook.Save(roundtripPath);

  const chartXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/charts/chartEx2.xml");
  const chartRoot = documentRoot(parseXml(chartXml));
  AssertEx.notNull(chartRoot, "chartEx2.xml must have a root element.");
  const seriesList = descendants(chartRoot, "series");
  AssertEx.true(seriesList.length > 0, "Map Chart samples roundtrip must contain cx:series.");
  const series = seriesList[0];
  AssertEx.null(
    firstElement(series, "valueColors"),
    "Map Chart samples roundtrip must remove invalid cx:series/cx:valueColors.",
  );
  AssertEx.null(
    firstElement(series, "valueColorPositions"),
    "Map Chart samples roundtrip must remove invalid cx:series/cx:valueColorPositions.",
  );
}

export function excelInputMapChartSamples1RoundtripPreservesShapeHyperlinkRelationships(): void {
  const inputPath = resolveRepositoryFile("Input", path.join("MapChart", "Map Chart samples1.xlsx"));
  const workbook = new Workbook(inputPath);

  const temp = new TemporaryDirectory(path.join("chart-roundtrips", "mapchart-shape-hyperlink"));
  const roundtripPath = temp.getPath("Map Chart samples1out.xlsx");
  workbook.Save(roundtripPath);

  const drawingXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/drawings/drawing1.xml");
  const drawingRoot = documentRoot(parseXml(drawingXml));
  AssertEx.notNull(drawingRoot, "drawing1.xml must have a root element.");

  let nextButton: XmlNode | null = null;
  for (const shape of descendants(drawingRoot, "sp")) {
    const nvSpPr = firstElement(shape, "nvSpPr");
    const cNvPr = firstElement(nvSpPr, "cNvPr");
    if ((attribute(cNvPr, "name") ?? "") === "Next Button") {
      nextButton = shape;
      break;
    }
  }

  AssertEx.notNull(nextButton, "Map Chart samples1 roundtrip must preserve the Next Button shape.");

  const hyperlinkElements = descendants(nextButton, "hlinkClick");
  AssertEx.true(
    hyperlinkElements.length > 0,
    "Map Chart samples1 roundtrip must preserve the Next Button hyperlink.",
  );
  const hyperlinkElement = hyperlinkElements[0];
  const hyperlinkRelationshipId = attribute(hyperlinkElement, "id");
  AssertEx.true(
    hyperlinkRelationshipId !== null && hyperlinkRelationshipId.length > 0,
    "Map Chart samples1 roundtrip hyperlink must keep a drawing relationship id.",
  );

  const drawingRelsXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/drawings/_rels/drawing1.xml.rels");
  const drawingRelsRoot = documentRoot(parseXml(drawingRelsXml));
  AssertEx.notNull(drawingRelsRoot, "drawing1.xml.rels must have a root element.");

  let hyperlinkRelationship: XmlNode | null = null;
  for (const relationship of elementChildren(drawingRelsRoot, "Relationship")) {
    if ((attribute(relationship, "Id") ?? "") === hyperlinkRelationshipId) {
      hyperlinkRelationship = relationship;
      break;
    }
  }

  AssertEx.notNull(
    hyperlinkRelationship,
    "Map Chart samples1 roundtrip must write the shape hyperlink relationship.",
  );
  AssertEx.equal(
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    attribute(hyperlinkRelationship!, "Type"),
  );
  AssertEx.equal("#'3-color Map Chart'!A1", attribute(hyperlinkRelationship!, "Target"));
}

export function excelInputComboChartRoundtrip(): void {
  roundtripAllChartFilesInDirectory("combochart", false);
}

export function excelInputCombochart61RoundtripPreservesChartImageRelationships(): void {
  const inputPath = resolveRepositoryFile("Input", path.join("allcharts", "combochart6-1.xlsx"));
  const workbook = new Workbook(inputPath);

  const temp = new TemporaryDirectory("compat-input-combochart6-1");
  const roundtripPath = temp.getPath("combochart6-1out.xlsx");
  workbook.Save(roundtripPath);

  const chartXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/charts/chart1.xml");
  AssertEx.contains(
    'r:embed="rId3"',
    chartXml,
    "combochart6-1 chart text formatting must keep the chart image relationship reference.",
  );

  const chartRelsXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/charts/_rels/chart1.xml.rels");
  AssertEx.contains(
    'Id="rId3"',
    chartRelsXml,
    "combochart6-1 chart relationships must preserve the embedded image relationship id.",
  );
  AssertEx.contains(
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    chartRelsXml,
    "combochart6-1 chart relationships must preserve the embedded image relationship type.",
  );

  const chartRelsRoot = documentRoot(parseXml(chartRelsXml));
  AssertEx.notNull(chartRelsRoot, "chart1.xml.rels must have a root element.");

  let imageRelationship: XmlNode | null = null;
  for (const relationship of elementChildren(chartRelsRoot, "Relationship")) {
    if ((attribute(relationship, "Id") ?? "") === "rId3") {
      imageRelationship = relationship;
      break;
    }
  }

  AssertEx.notNull(imageRelationship, "combochart6-1 must keep chart relationship rId3.");

  const target = attribute(imageRelationship!, "Target") ?? "";
  AssertEx.true(target.startsWith("../media/"), "combochart6-1 chart image relationship must target xl/media.");

  const normalizedTarget = "xl/" + target.substring(3).split("\\").join("/");
  AssertEx.true(
    ZipPackageHelper.entryExists(roundtripPath, normalizedTarget),
    "combochart6-1 must keep the chart companion image part referenced by rId3.",
  );
}

export function excelInputWaterfallChartRoundtrip(): void {
  roundtripAllChartFilesInDirectory("waterfallchart", false);
}

export function excelInputWaterfallStackedV2RoundtripNormalizesDrawingChartAndStyles(): void {
  const inputPath = resolveRepositoryFile(
    "Input",
    path.join("waterfallchart", "Waterfall_Stacked_V2.xlsx"),
  );
  const workbook = new Workbook(inputPath);

  const temp = new TemporaryDirectory(path.join("chart-roundtrips", "waterfallchart-normalized"));
  const roundtripPath = temp.getPath("Waterfall_Stacked_V2out.xlsx");
  workbook.Save(roundtripPath);

  const drawingXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/drawings/drawing2.xml");
  const drawingRoot = documentRoot(parseXml(drawingXml));
  AssertEx.notNull(drawingRoot, "drawing2.xml must have a root element.");

  const drawingIds = new Set<string>();
  for (const cNvPr of descendants(drawingRoot, "cNvPr")) {
    const id = attribute(cNvPr, "id") ?? "";
    AssertEx.true(
      !drawingIds.has(id),
      "Waterfall_Stacked_V2 drawing cNvPr/@id must be unique. Duplicate id: " + id,
    );
    drawingIds.add(id);
  }

  for (const chartEntry of ["xl/charts/chart1.xml", "xl/charts/chart2.xml", "xl/charts/chart4.xml"]) {
    const chartXml = ZipPackageHelper.readEntryText(roundtripPath, chartEntry);
    AssertEx.true(
      chartXml.indexOf("c15:showLeaderLines") < 0,
      "Waterfall_Stacked_V2 must remove unsupported c15:showLeaderLines from " + chartEntry,
    );
  }

  const stylesXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/styles.xml");
  const stylesRoot = documentRoot(parseXml(stylesXml));
  AssertEx.notNull(stylesRoot, "styles.xml must have a root element.");

  const fontsElement = firstElement(stylesRoot, "fonts");
  AssertEx.notNull(fontsElement, "Waterfall_Stacked_V2 must contain a fonts collection.");
  const fonts = elementChildren(fontsElement, "font");
  AssertEx.true(fonts.length > 0, "Waterfall_Stacked_V2 must contain at least one font.");

  for (let index = 0; index < Math.min(5, fonts.length); index++) {
    const childNames = childLocalNames(fonts[index]);
    const sizeIndex = findElementIndex(childNames, "sz");
    const colorIndex = findElementIndex(childNames, "color");
    const nameIndex = findElementIndex(childNames, "name");
    const familyIndex = findElementIndex(childNames, "family");
    AssertEx.true(sizeIndex >= 0, "Font must contain sz.");
    AssertEx.true(colorIndex < 0 || sizeIndex < colorIndex, "Font sz must appear before color.");
    AssertEx.true(nameIndex < 0 || sizeIndex < nameIndex, "Font sz must appear before name.");
    AssertEx.true(familyIndex < 0 || nameIndex < familyIndex, "Font name must appear before family.");
  }
}

export function excelInputIcWaterfallTemplateRoundtripRemovesUnsupportedChartLevelExtensions(): void {
  const inputPath = resolveRepositoryFile(
    "Input",
    path.join("waterfallchart", "IC-Waterfall-Chart-Template-2016-8568.xlsx"),
  );
  const workbook = new Workbook(inputPath);

  const temp = new TemporaryDirectory(path.join("chart-roundtrips", "waterfallchart-template-normalized"));
  const roundtripPath = temp.getPath("IC-Waterfall-Chart-Template-2016-8568out.xlsx");
  workbook.Save(roundtripPath);

  for (const chartEntry of ["xl/charts/chart1.xml", "xl/charts/chart2.xml"]) {
    const chartXml = ZipPackageHelper.readEntryText(roundtripPath, chartEntry);
    AssertEx.true(
      chartXml.indexOf("c:showDLblsOverMax") < 0,
      "IC waterfall template must remove unsupported chart-level c:showDLblsOverMax from " + chartEntry,
    );
    AssertEx.true(
      chartXml.indexOf("c16r3:dataDisplayOptions16") < 0,
      "IC waterfall template must remove unsupported chart-level c16r3:dataDisplayOptions16 from " + chartEntry,
    );

    const chartRoot = documentRoot(parseXml(chartXml));
    AssertEx.notNull(chartRoot, chartEntry + " must have a root element.");
    const chartElement = firstElement(chartRoot, "chart");
    AssertEx.notNull(chartElement, chartEntry + " must contain c:chart.");
    AssertEx.null(
      firstElement(chartElement, "extLst"),
      "IC waterfall template must remove unsupported chart-level c:extLst from " + chartEntry,
    );
  }
}

export function excelInputAllChartsRoundtrip(): void {
  roundtripAllChartFilesInDirectory("allcharts", false);
}

export function excelInputSparklineRoundtripNoCrash(): void {
  roundtripAllChartFilesInDirectory("Sparkline", false);
}
