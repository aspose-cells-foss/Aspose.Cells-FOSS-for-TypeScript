import { Color, FillPattern, Workbook } from "../../aspose_cells";

const workbook = new Workbook();
const cell = workbook.Worksheets[0].Cells["A1"];
cell.PutValue("Styled");

const style = cell.GetStyle();
style.Font.IsBold = true;
style.Pattern = FillPattern.Solid;
style.ForegroundColor = Color.FromArgb(255, 241, 196, 15);
cell.SetStyle(style);

console.log(`${cell.StringValue} / Bold=${cell.GetStyle().Font.IsBold ? "True" : "False"}`);
