import { CellArea, Color, FillPattern, FormatConditionType, OperatorType, Workbook } from "../../aspose_cells";
import { sampleOutputPath } from "../Shared";

const outputPath = sampleOutputPath("conditional-formatting-sample.xlsx");

const workbook = new Workbook();
const sheet = workbook.Worksheets[0];
sheet.Name = "Conditional Formatting";

for (let index = 0; index < 10; index++) {
  sheet.Cells.get(index, 0).PutValue(index + 1);
  sheet.Cells.get(index, 1).PutValue((index + 1) * 10);
  sheet.Cells.get(index, 2).PutValue((index + 1) * 10);
  sheet.Cells.get(index, 3).PutValue((index + 1) * 10);
  sheet.Cells.get(index, 4).PutValue((index + 1) * 10);
}

const betweenCollection = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
betweenCollection.AddArea(CellArea.CreateCellArea("A1", "A10"));
const betweenRule = betweenCollection.get(betweenCollection.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "3", "7"));
betweenRule.Priority = 1;
betweenRule.StopIfTrue = true;
const betweenStyle = betweenRule.Style;
betweenStyle.Pattern = FillPattern.Solid;
betweenStyle.ForegroundColor = Color.FromArgb(255, 255, 199, 206);
betweenStyle.Font.IsBold = true;
betweenStyle.Font.Color = Color.FromArgb(255, 156, 0, 6);
betweenRule.Style = betweenStyle;

const expressionCollection = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
expressionCollection.AddArea(CellArea.CreateCellArea("B1", "B10"));
const expressionRule = expressionCollection.get(expressionCollection.AddCondition(FormatConditionType.Expression, OperatorType.None, "MOD(B1,20)=0", ""));
expressionRule.Priority = 2;
const expressionStyle = expressionRule.Style;
expressionStyle.Font.IsItalic = true;
expressionStyle.Font.Color = Color.FromArgb(255, 0, 0, 255);
expressionRule.Style = expressionStyle;

const colorScaleCollection = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
colorScaleCollection.AddArea(CellArea.CreateCellArea("C1", "C10"));
const colorScaleRule = colorScaleCollection.get(colorScaleCollection.AddCondition(FormatConditionType.ColorScale));
colorScaleRule.ColorScaleCount = 3;
colorScaleRule.MinColor = Color.FromArgb(255, 248, 105, 107);
colorScaleRule.MidColor = Color.FromArgb(255, 255, 235, 132);
colorScaleRule.MaxColor = Color.FromArgb(255, 99, 190, 123);

const dataBarCollection = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
dataBarCollection.AddArea(CellArea.CreateCellArea("D1", "D10"));
const dataBarRule = dataBarCollection.get(dataBarCollection.AddCondition(FormatConditionType.DataBar));
dataBarRule.BarColor = Color.FromArgb(255, 99, 142, 198);
dataBarRule.NegativeBarColor = Color.FromArgb(255, 255, 0, 0);
dataBarRule.ShowBorder = true;
dataBarRule.Direction = "left-to-right";

const iconSetCollection = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
iconSetCollection.AddArea(CellArea.CreateCellArea("E1", "E10"));
const iconSetRule = iconSetCollection.get(iconSetCollection.AddCondition(FormatConditionType.IconSet));
iconSetRule.IconSetType = "4Arrows";
iconSetRule.ReverseIcons = true;
iconSetRule.ShowIconOnly = true;

workbook.Save(outputPath);

const loaded = new Workbook(outputPath);
const loadedSheet = loaded.Worksheets["Conditional Formatting"];

console.log("Saved: " + outputPath);
console.log("Conditional formatting collections: " + loadedSheet.ConditionalFormattings.Count);
console.log("A-column rule type: " + FormatConditionType[loadedSheet.ConditionalFormattings.get(0).get(0).Type as unknown as number]);
console.log("C-column rule type: " + FormatConditionType[loadedSheet.ConditionalFormattings.get(2).get(0).Type as unknown as number]);
console.log("E-column icon set: " + loadedSheet.ConditionalFormattings.get(4).get(0).IconSetType);
