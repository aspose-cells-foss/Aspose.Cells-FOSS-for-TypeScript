import { Color, VisibilityType, Workbook } from "../../aspose_cells";
import { sampleOutputPath } from "../Shared";

const outputPath = sampleOutputPath("worksheet-settings-sample.xlsx");

const workbook = new Workbook();
const layout = workbook.Worksheets[0];
layout.Name = "Layout";
layout.VisibilityType = VisibilityType.Hidden;
layout.TabColor = Color.FromArgb(255, 34, 68, 102);
layout.ShowGridlines = false;
layout.ShowRowColumnHeaders = false;
layout.ShowZeros = false;
layout.RightToLeft = true;
layout.Zoom = 85;
layout.Protect();
layout.Protection.Objects = true;
layout.Protection.Scenarios = true;
layout.Protection.FormatCells = true;
layout.Protection.InsertRows = true;
layout.Protection.AutoFilter = true;
layout.Protection.SelectLockedCells = true;
layout.Protection.SelectUnlockedCells = true;

layout.Cells["A1"].PutValue("Merged");
layout.Cells["C4"].PutValue(99);
layout.Cells.Rows[1].Height = 22.5;
layout.Cells.Rows[3].IsHidden = true;
layout.Cells.Columns[0].Width = 18.25;
layout.Cells.Columns[2].IsHidden = true;
layout.Cells.Merge(0, 0, 2, 2);

const visibleIndex = workbook.Worksheets.Add("Visible");
const visibleSheet = workbook.Worksheets[visibleIndex];
visibleSheet.Cells["A1"].PutValue("Visible");
workbook.Worksheets.ActiveSheetName = "Visible";

workbook.Save(outputPath);

const loaded = new Workbook(outputPath);
const loadedLayout = loaded.Worksheets["Layout"];

console.log("Saved: " + outputPath);
console.log("Active sheet: " + loaded.Worksheets.ActiveSheetName);
console.log("Layout visibility: " + VisibilityType[loadedLayout.VisibilityType]);
console.log("Merged value: " + loadedLayout.Cells["A1"].StringValue);
console.log("Row 2 height: " + (loadedLayout.Cells.Rows[1].Height ?? 0));
console.log("Column A width: " + (loadedLayout.Cells.Columns[0].Width ?? 0));
console.log("Merged regions: " + loadedLayout.Cells.MergedCells.length);
