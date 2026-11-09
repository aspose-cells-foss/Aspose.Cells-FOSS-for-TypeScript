import { Workbook } from "../../aspose_cells";
import { sampleOutputPath } from "../Shared";

function boolText(value: boolean): string {
  return value ? "True" : "False";
}

const outputPath = sampleOutputPath("comments-sample.xlsx");

const workbook = new Workbook();
const sheet = workbook.Worksheets[0];
sheet.Name = "Comments";

sheet.Cells["A1"].PutValue("Task");
sheet.Cells["B1"].PutValue("Status");
sheet.Cells["A2"].PutValue("Design UI");
sheet.Cells["B2"].PutValue("In Progress");
sheet.Cells["A3"].PutValue("Implement API");
sheet.Cells["B3"].PutValue("Pending");
sheet.Cells["A4"].PutValue("Write Tests");
sheet.Cells["B4"].PutValue("Completed");

const comment2 = sheet.Comments.Add(1, 0);
comment2.Author = "John";
comment2.Note = "Need to finalize the color scheme by Friday";
comment2.IsVisible = false;
comment2.Width = 200;
comment2.Height = 100;

const comment3 = sheet.Comments.Add("A3");
comment3.Author = "Sarah";
comment3.Note = "Waiting for database schema approval";
comment3.IsVisible = true;
comment3.Width = 180;
comment3.Height = 90;

const comment4 = sheet.Comments.Add("B4");
comment4.Author = "Mike";
comment4.Note = "All unit tests passing. Ready for review.";
comment4.IsVisible = false;

workbook.Save(outputPath);

const loaded = new Workbook(outputPath);
const loadedSheet = loaded.Worksheets["Comments"];

console.log("Saved: " + outputPath);
console.log("Comment count: " + loadedSheet.Comments.Count);
console.log("A2 comment by " + loadedSheet.Comments["A2"].Author + ": " + loadedSheet.Comments["A2"].Note);
console.log("A3 comment visible: " + boolText(loadedSheet.Comments["A3"].IsVisible));
console.log("B4 comment: " + loadedSheet.Comments["B4"].Note);
