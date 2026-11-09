import { TargetModeType, Workbook } from "../../aspose_cells";
import { sampleOutputPath } from "../Shared";

function getSheetIndexOrDefault(localSheetIndex: number | null): number {
  return localSheetIndex ?? -1;
}

const outputPath = sampleOutputPath("hyperlinks-and-names-sample.xlsx");

const workbook = new Workbook();
const linksSheet = workbook.Worksheets[0];
linksSheet.Name = "Links";
const scopedSheetIndex = workbook.Worksheets.Add("Scoped Sheet");
const scopedSheet = workbook.Worksheets[scopedSheetIndex];
scopedSheet.Cells["B2"].PutValue(5);

linksSheet.Cells["A1"].PutValue("Docs");
const externalLink = linksSheet.Hyperlinks[linksSheet.Hyperlinks.Add("A1", 1, 1, "https://example.com/docs?q=1")];
externalLink.TextToDisplay = "Docs";
externalLink.ScreenTip = "External docs";

linksSheet.Cells["B2"].PutValue("Jump");
const internalLink = linksSheet.Hyperlinks[linksSheet.Hyperlinks.Add("B2", 1, 1, "'Scoped Sheet'!B2")];
internalLink.TextToDisplay = "Jump";
internalLink.ScreenTip = "Jump to scoped sheet";

const mailLink = linksSheet.Hyperlinks[linksSheet.Hyperlinks.Add("C4", "D5", "mailto:test@example.com", "Mail", "Send mail")];
mailLink.ScreenTip = "Send mail";

const globalName = workbook.DefinedNames[workbook.DefinedNames.Add("GlobalRange", "='Links'!$A$1:$D$5")];
globalName.Hidden = true;
globalName.Comment = "Primary sample range";

const localName = workbook.DefinedNames[workbook.DefinedNames.Add("ScopedCell", "'Scoped Sheet'!$B$2", scopedSheetIndex)];
localName.Comment = "Sheet-scoped name";

workbook.Save(outputPath);

const loaded = new Workbook(outputPath);
const loadedLinks = loaded.Worksheets["Links"];

console.log("Saved: " + outputPath);
console.log("Hyperlinks: " + loadedLinks.Hyperlinks.Count);
console.log("First hyperlink: " + loadedLinks.Hyperlinks[0].Address + " / " + TargetModeType[loadedLinks.Hyperlinks[0].LinkType]);
console.log("Second hyperlink: " + loadedLinks.Hyperlinks[1].Address + " / " + TargetModeType[loadedLinks.Hyperlinks[1].LinkType]);
console.log("Defined names: " + loaded.DefinedNames.Count);
console.log("Global name formula: " + loaded.DefinedNames[0].Formula);
console.log("Local name scope: " + getSheetIndexOrDefault(loaded.DefinedNames[1].LocalSheetIndex));
