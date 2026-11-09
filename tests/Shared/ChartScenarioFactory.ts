import { ChartType, Workbook } from "../../aspose_cells";
import { AssertEx } from "./TestInfrastructure";

/** Scenario factory for chart workbooks. */
export class ChartScenarioFactory {
  private static readonly QuarterLabels: string[] = ["Q1", "Q2", "Q3", "Q4", "Q5"];
  private static readonly QuarterValues: number[] = [10, 30, 20, 50, 40];

  public static CreateChartWorkbook(type: ChartType, sheetName: string = "Sheet1"): Workbook {
    const workbook = new Workbook();
    const sheet = workbook.Worksheets[0];
    sheet.Name = sheetName;

    // Populate data in A1:B5
    sheet.Cells["A1"].PutValue(ChartScenarioFactory.QuarterLabels[0]); sheet.Cells["B1"].PutValue(ChartScenarioFactory.QuarterValues[0]);
    sheet.Cells["A2"].PutValue(ChartScenarioFactory.QuarterLabels[1]); sheet.Cells["B2"].PutValue(ChartScenarioFactory.QuarterValues[1]);
    sheet.Cells["A3"].PutValue(ChartScenarioFactory.QuarterLabels[2]); sheet.Cells["B3"].PutValue(ChartScenarioFactory.QuarterValues[2]);
    sheet.Cells["A4"].PutValue(ChartScenarioFactory.QuarterLabels[3]); sheet.Cells["B4"].PutValue(ChartScenarioFactory.QuarterValues[3]);
    sheet.Cells["A5"].PutValue(ChartScenarioFactory.QuarterLabels[4]); sheet.Cells["B5"].PutValue(ChartScenarioFactory.QuarterValues[4]);

    const dataRange = sheetName + "!$B$1:$B$5";
    sheet.Charts.Add(type, dataRange, 7, 0, 22, 10);
    return workbook;
  }

  public static AssertChartWorkbook(workbook: Workbook, expectedType: ChartType): void {
    const sheet = workbook.Worksheets[0];
    AssertEx.equal(1, sheet.Charts.Count, "Expected exactly 1 chart");
    const chart = sheet.Charts.get(0);
    AssertEx.equal(expectedType, chart.ChartType, "Wrong chart type after round-trip");
    AssertEx.true(chart.Name.length > 0, "Chart name should not be empty");
    AssertEx.equal(7, chart.UpperLeftRow, "Wrong UpperLeftRow");
    AssertEx.equal(0, chart.UpperLeftColumn, "Wrong UpperLeftColumn");
    AssertEx.equal(22, chart.LowerRightRow, "Wrong LowerRightRow");
    AssertEx.equal(10, chart.LowerRightColumn, "Wrong LowerRightColumn");
  }
}
