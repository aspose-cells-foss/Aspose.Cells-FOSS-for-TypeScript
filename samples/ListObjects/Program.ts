import { TableStyleType, Workbook } from "../../aspose_cells";
import { sampleOutputPath } from "../Shared";

function boolText(value: boolean): string {
  return value ? "True" : "False";
}

const outputPath = sampleOutputPath("listobjects-sample.xlsx");

const workbook = new Workbook();
const sheet = workbook.Worksheets[0];
sheet.Name = "Data";

sheet.Cells["A1"].PutValue("Product");
sheet.Cells["B1"].PutValue("Category");
sheet.Cells["C1"].PutValue("Price");
sheet.Cells["D1"].PutValue("Quantity");

sheet.Cells["A2"].PutValue("Laptop");
sheet.Cells["B2"].PutValue("Electronics");
sheet.Cells["C2"].PutValue(999.99);
sheet.Cells["D2"].PutValue(50);

sheet.Cells["A3"].PutValue("Mouse");
sheet.Cells["B3"].PutValue("Electronics");
sheet.Cells["C3"].PutValue(29.99);
sheet.Cells["D3"].PutValue(200);

sheet.Cells["A4"].PutValue("Desk Chair");
sheet.Cells["B4"].PutValue("Furniture");
sheet.Cells["C4"].PutValue(149.99);
sheet.Cells["D4"].PutValue(75);

sheet.Cells["A5"].PutValue("Monitor");
sheet.Cells["B5"].PutValue("Electronics");
sheet.Cells["C5"].PutValue(349.99);
sheet.Cells["D5"].PutValue(100);

const tableIndex = sheet.ListObjects.Add("A1", "D5", true);
const table = sheet.ListObjects[tableIndex];
table.DisplayName = "Products";
table.Comment = "Product inventory table";
table.TableStyleType = TableStyleType.TableStyleMedium2;
table.ShowTableStyleFirstColumn = true;
table.ShowTableStyleLastColumn = true;
table.ShowTableStyleRowStripes = true;
table.ShowTotals = true;

sheet.Cells["A6"].PutValue("Total");
sheet.Cells["D6"].PutValue("=SUM(D2:D5)");

workbook.Save(outputPath);

const loaded = new Workbook(outputPath);
const loadedSheet = loaded.Worksheets["Data"];
const loadedTable = loadedSheet.ListObjects["Products"];

console.log("Saved: " + outputPath);
console.log("Table count: " + loadedSheet.ListObjects.Count);
console.log("Table name: " + loadedTable.DisplayName);
console.log("Table comment: " + loadedTable.Comment);
console.log("Table style: " + TableStyleType[loadedTable.TableStyleType]);
console.log("Show headers: " + boolText(loadedTable.ShowHeaderRow));
console.log("Show totals: " + boolText(loadedTable.ShowTotals));
console.log("Column count: " + loadedTable.ListColumns.Count);
console.log("First column: " + loadedTable.ListColumns.get(0).Name);
