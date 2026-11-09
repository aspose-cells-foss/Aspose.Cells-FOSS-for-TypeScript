import { Cell, Workbook } from "../../aspose_cells";
import { sampleOutputPath } from "../Shared";

function getValueTypeName(cell: Cell): string {
  const value = cell.Value;
  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "string") {
    return "String";
  }

  if (typeof value === "boolean") {
    return "Boolean";
  }

  if (value instanceof Date) {
    return "DateTime";
  }

  if (typeof value === "number") {
    return "Double";
  }

  return "";
}

const outputPath = sampleOutputPath("cell-data-roundtrip.xlsx");
const timestamp = new Date(Date.UTC(2024, 4, 6, 7, 8, 9));

const workbook = new Workbook();
const sheet = workbook.Worksheets[0];
sheet.Cells["A1"].PutValue("Hello");
sheet.Cells["B1"].PutValue(123);
sheet.Cells["C1"].PutValue(true);
sheet.Cells["D1"].PutValue(12.5);
sheet.Cells["E1"].PutValue(timestamp);
sheet.Cells["F1"].PutValue(10);
sheet.Cells["G1"].PutValue(20);
sheet.Cells["G1"].Formula = "=F1*2";
workbook.Save(outputPath);

const loaded = new Workbook(outputPath);
const loadedSheet = loaded.Worksheets[0];

console.log(loadedSheet.Cells["A1"].StringValue);
console.log(getValueTypeName(loadedSheet.Cells["B1"]) + ":" + loadedSheet.Cells["B1"].StringValue);
console.log(getValueTypeName(loadedSheet.Cells["C1"]) + ":" + loadedSheet.Cells["C1"].StringValue);
console.log(getValueTypeName(loadedSheet.Cells["D1"]) + ":" + loadedSheet.Cells["D1"].StringValue);
console.log(getValueTypeName(loadedSheet.Cells["E1"]) + ":" + loadedSheet.Cells["E1"].StringValue);
console.log(loadedSheet.Cells["G1"].Formula + " -> " + loadedSheet.Cells["G1"].StringValue);
