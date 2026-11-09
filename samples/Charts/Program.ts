import { ChartType, Workbook } from "../../aspose_cells";
import { sampleOutputPath } from "../Shared";

const outputPath = sampleOutputPath("charts-sample.xlsx");

const workbook = new Workbook();
const sheet = workbook.Worksheets[0];
sheet.Name = "Charts";

sheet.Cells["A1"].PutValue("Month");
sheet.Cells["B1"].PutValue("Sales");
sheet.Cells["C1"].PutValue("Profit");

for (let month = 1; month <= 12; month++) {
  sheet.Cells.get(month, 0).PutValue("Month " + month);
  sheet.Cells.get(month, 1).PutValue(month * 1000 + month * 50);
  sheet.Cells.get(month, 2).PutValue(month * 300 + month * 20);
}

const salesChartIndex = sheet.Charts.Add(ChartType.Column, "Charts!$B$1:$B$13", 0, 4, 18, 8);
const salesChart = sheet.Charts[salesChartIndex];
sheet.Cells["E1"].PutValue("Sales Chart: " + salesChart.Name);

const profitChartIndex = sheet.Charts.Add(ChartType.Line, "Charts!$C$1:$C$13", 0, 9, 18, 13);
const profitChart = sheet.Charts[profitChartIndex];
sheet.Cells["J1"].PutValue("Profit Chart: " + profitChart.Name);

workbook.Save(outputPath);

const loaded = new Workbook(outputPath);
const loadedSheet = loaded.Worksheets["Charts"];

console.log("Saved: " + outputPath);
console.log("Chart count: " + loadedSheet.Charts.Count);
console.log("First chart: " + loadedSheet.Charts[0].Name + " (Type: " + ChartType[loadedSheet.Charts[0].ChartType] + ")");
console.log("Second chart: " + loadedSheet.Charts[1].Name + " (Type: " + ChartType[loadedSheet.Charts[1].ChartType] + ")");
console.log("First chart anchor: R" + loadedSheet.Charts[0].UpperLeftRow + "C" + loadedSheet.Charts[0].UpperLeftColumn + " to R" + loadedSheet.Charts[0].LowerRightRow + "C" + loadedSheet.Charts[0].LowerRightColumn);
