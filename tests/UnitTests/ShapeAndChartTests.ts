import * as path from "path";
import {
  AutoShapeType,
  ChartType,
  SaveFormat,
  Workbook,
} from "../../aspose_cells";
import { AssertEx, getOrCreateOutputDirectory, MemoryStream } from "../Shared/TestInfrastructure";
import { ChartScenarioFactory } from "../Shared/ChartScenarioFactory";
import { PictureScenarioFactory } from "../Shared/PictureScenarioFactory";

export function shapeApisMutateExpectedSettings(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];

  AssertEx.equal(0, sheet.Shapes.Count);

  const idx = sheet.Shapes.Add(1, 1, 5, 3, AutoShapeType.Rectangle);
  AssertEx.equal(0, idx);
  AssertEx.equal(1, sheet.Shapes.Count);

  const shape = sheet.Shapes.get(0);
  AssertEx.equal("Shape 1", shape.Name);
  AssertEx.equal(1, shape.UpperLeftRow);
  AssertEx.equal(1, shape.UpperLeftColumn);
  AssertEx.equal(5, shape.LowerRightRow);
  AssertEx.equal(3, shape.LowerRightColumn);
  AssertEx.equal(AutoShapeType.Rectangle, shape.AutoShapeType);
  AssertEx.equal("rect", shape.GeometryType);

  shape.Name = "MyRect";
  AssertEx.equal("MyRect", sheet.Shapes.get(0).Name);

  shape.UpperLeftRow = 0;
  shape.UpperLeftColumn = 0;
  shape.LowerRightRow = 4;
  shape.LowerRightColumn = 2;
  AssertEx.equal(0, sheet.Shapes.get(0).UpperLeftRow);
  AssertEx.equal(4, sheet.Shapes.get(0).LowerRightRow);

  shape.AutoShapeType = AutoShapeType.RightArrow;
  AssertEx.equal(AutoShapeType.RightArrow, sheet.Shapes.get(0).AutoShapeType);
  AssertEx.equal("rightArrow", sheet.Shapes.get(0).GeometryType);

  shape.GeometryType = "star12";
  AssertEx.equal(AutoShapeType.Star12Point, sheet.Shapes.get(0).AutoShapeType);

  const idx2 = sheet.Shapes.Add(7, 1, 9, 3, AutoShapeType.Ellipse);
  AssertEx.equal(1, idx2);
  AssertEx.equal(2, sheet.Shapes.Count);
  AssertEx.equal(AutoShapeType.Ellipse, sheet.Shapes.get(1).AutoShapeType);

  sheet.Shapes.RemoveAt(0);
  AssertEx.equal(1, sheet.Shapes.Count);
  AssertEx.equal(AutoShapeType.Ellipse, sheet.Shapes.get(0).AutoShapeType);

  AssertEx.throws(() => {
    sheet.Shapes.Add(-1, 0, 3, 3, AutoShapeType.Rectangle);
  });
  AssertEx.throws(() => {
    sheet.Shapes.Add(0, 2, 3, 1, AutoShapeType.Rectangle);
  });
  AssertEx.throws(() => {
    sheet.Shapes.RemoveAt(99);
  });
  AssertEx.throws(() => {
    sheet.Shapes.get(99);
  });
}

export function chartCollectionIsEmptyOnNewWorksheet(): void {
  const workbook = new Workbook();
  AssertEx.equal(0, workbook.Worksheets[0].Charts.Count);
}

export function chartAddColumnChartReturnsIndexZero(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  const idx = sheet.Charts.Add(ChartType.Column, "Sheet1!$B$1:$B$5", 5, 0, 15, 8);
  AssertEx.equal(0, idx);
  AssertEx.equal(1, sheet.Charts.Count);
}

export function chartAddMultipleChartsIncrementsCount(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  sheet.Charts.Add(ChartType.Column, "Sheet1!$B$1:$B$5", 0, 0, 10, 8);
  sheet.Charts.Add(ChartType.Line, "Sheet1!$C$1:$C$5", 12, 0, 22, 8);
  AssertEx.equal(2, sheet.Charts.Count);
  AssertEx.equal(ChartType.Column, sheet.Charts.get(0).ChartType);
  AssertEx.equal(ChartType.Line, sheet.Charts.get(1).ChartType);
}

export function chartAddSetsCorrectPosition(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  sheet.Charts.Add(ChartType.Bar, "Sheet1!$B$1:$B$5", 7, 2, 20, 12);
  const chart = sheet.Charts.get(0);
  AssertEx.equal(7, chart.UpperLeftRow);
  AssertEx.equal(2, chart.UpperLeftColumn);
  AssertEx.equal(20, chart.LowerRightRow);
  AssertEx.equal(12, chart.LowerRightColumn);
}

export function chartAddSetsCorrectChartType(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  sheet.Charts.Add(ChartType.Pie, "Sheet1!$B$1:$B$5", 0, 0, 10, 8);
  AssertEx.equal(ChartType.Pie, sheet.Charts.get(0).ChartType);
}

export function chartAddUnsupportedChartexTypeThrows(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  AssertEx.throws(() => {
    sheet.Charts.Add(ChartType.Waterfall, "Sheet1!$B$1:$B$5", 0, 0, 10, 8);
  });
  AssertEx.throws(() => {
    sheet.Charts.Add(ChartType.Treemap, "Sheet1!$B$1:$B$5", 0, 0, 10, 8);
  });
  AssertEx.throws(() => {
    sheet.Charts.Add(ChartType.Funnel, "Sheet1!$B$1:$B$5", 0, 0, 10, 8);
  });
}

export function chartIndexOutOfRangeThrows(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  AssertEx.throws(() => {
    sheet.Charts.get(0);
  });
  sheet.Charts.Add(ChartType.Column, "Sheet1!$B$1:$B$5", 0, 0, 10, 8);
  AssertEx.throws(() => {
    sheet.Charts.get(1);
  });
  AssertEx.throws(() => {
    sheet.Charts.get(-1);
  });
}

export function chartRoundtripColumnPreservesCountAndType(): void {
  roundtripChart(ChartType.Column);
}

export function chartRoundtripBarPreservesCountAndType(): void {
  roundtripChart(ChartType.Bar);
}

export function chartRoundtripLinePreservesCountAndType(): void {
  roundtripChart(ChartType.Line);
}

export function chartRoundtripAreaPreservesCountAndType(): void {
  roundtripChart(ChartType.Area);
}

export function chartRoundtripPiePreservesCountAndType(): void {
  roundtripChart(ChartType.Pie);
}

export function chartRoundtripDoughnutPreservesCountAndType(): void {
  roundtripChart(ChartType.Doughnut);
}

export function chartRoundtripScatterPreservesCountAndType(): void {
  roundtripChart(ChartType.Scatter);
}

export function chartRoundtripRadarPreservesCountAndType(): void {
  roundtripChart(ChartType.Radar);
}

export function chartRoundtripBubblePreservesCountAndType(): void {
  roundtripChart(ChartType.Bubble);
}

export function chartRoundtripStockPreservesCountAndType(): void {
  roundtripChart(ChartType.Stock);
}

export function chartRoundtripColumn3DPreservesCountAndType(): void {
  roundtripChart(ChartType.Column3D);
}

export function chartRoundtripBar3DPreservesCountAndType(): void {
  roundtripChart(ChartType.Bar3D);
}

export function chartRoundtripLine3DPreservesCountAndType(): void {
  roundtripChart(ChartType.Line3D);
}

export function chartRoundtripArea3DPreservesCountAndType(): void {
  roundtripChart(ChartType.Area3D);
}

export function chartRoundtripPie3DPreservesCountAndType(): void {
  roundtripChart(ChartType.Pie3D);
}

export function chartRoundtripSurface3DPreservesCountAndType(): void {
  roundtripChart(ChartType.Surface3D);
}

export function chartRoundtripContourPreservesCountAndType(): void {
  roundtripChart(ChartType.Contour);
}

function roundtripChart(type: ChartType): void {
  const outputDir = getOrCreateOutputDirectory("chart-unit-roundtrips");
  const outputPath = path.join(outputDir, ChartType[type] + ".xlsx");

  const workbook = ChartScenarioFactory.CreateChartWorkbook(type);
  workbook.Save(outputPath);

  const reloaded = new Workbook(outputPath);
  ChartScenarioFactory.AssertChartWorkbook(reloaded, type);
}

export function chartWorksheetWithPictureAndChartSavesBoth(): void {
  const jpegBytes = PictureScenarioFactory.LoadJpeg();
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  sheet.Pictures.Add(0, 0, 5, 5, jpegBytes);
  sheet.Charts.Add(ChartType.Column, "Sheet1!$B$1:$B$5", 6, 0, 20, 10);
  AssertEx.equal(1, sheet.Pictures.Count);
  AssertEx.equal(1, sheet.Charts.Count);

  const ms = new MemoryStream();
  workbook.Save(ms, SaveFormat.Xlsx);
  ms.position = 0;
  const reloaded = new Workbook(ms);
  AssertEx.equal(1, reloaded.Worksheets[0].Pictures.Count);
  AssertEx.equal(1, reloaded.Worksheets[0].Charts.Count);
  AssertEx.equal(ChartType.Column, reloaded.Worksheets[0].Charts.get(0).ChartType);
}
