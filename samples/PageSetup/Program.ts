import { PageOrientationType, PaperSizeType, Workbook } from "../../aspose_cells";
import { sampleOutputPath } from "../Shared";

const outputPath = sampleOutputPath("page-setup-sample.xlsx");

const workbook = new Workbook();
const sheet = workbook.Worksheets[0];
sheet.Name = "Print Sheet";
sheet.Cells["A1"].PutValue("Title");
sheet.Cells["C10"].PutValue(42);

const pageSetup = sheet.PageSetup;
pageSetup.LeftMarginInch = 0.25;
pageSetup.RightMarginInch = 0.4;
pageSetup.TopMarginInch = 0.5;
pageSetup.BottomMarginInch = 0.6;
pageSetup.HeaderMarginInch = 0.2;
pageSetup.FooterMarginInch = 0.22;
pageSetup.Orientation = PageOrientationType.Landscape;
pageSetup.PaperSize = PaperSizeType.PaperA4;
pageSetup.FirstPageNumber = 3;
pageSetup.Scale = 95;
pageSetup.FitToPagesWide = 1;
pageSetup.FitToPagesTall = 2;
pageSetup.PrintArea = "$A$1:$C$10";
pageSetup.PrintTitleRows = "$1:$2";
pageSetup.PrintTitleColumns = "$A:$B";
pageSetup.LeftHeader = "Left Header";
pageSetup.CenterHeader = "Center Header";
pageSetup.RightHeader = "Right Header";
pageSetup.LeftFooter = "Left Footer";
pageSetup.CenterFooter = "Center Footer";
pageSetup.RightFooter = "Right Footer";
pageSetup.PrintGridlines = true;
pageSetup.PrintHeadings = true;
pageSetup.CenterHorizontally = true;
pageSetup.CenterVertically = true;
pageSetup.AddHorizontalPageBreak(4);
pageSetup.AddHorizontalPageBreak(7);
pageSetup.AddVerticalPageBreak(2);

workbook.Save(outputPath);

const loaded = new Workbook(outputPath);
const loadedPageSetup = loaded.Worksheets[0].PageSetup;

console.log("Saved: " + outputPath);
console.log("Orientation: " + PageOrientationType[loadedPageSetup.Orientation]);
console.log("Paper size: " + PaperSizeType[loadedPageSetup.PaperSize]);
console.log("Print area: " + loadedPageSetup.PrintArea);
console.log("Horizontal breaks: " + loadedPageSetup.HorizontalPageBreaks.join(", "));
console.log("Vertical breaks: " + loadedPageSetup.VerticalPageBreaks.join(", "));
