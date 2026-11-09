import { CellArea, OperatorType, Validation, ValidationAlertType, ValidationType, Workbook } from "../../aspose_cells";
import { sampleOutputPath } from "../Shared";

function requireValidation(validation: Validation | null, cellName: string): Validation {
  if (validation === null || validation === undefined) {
    throw new Error("Expected validation at " + cellName + ".");
  }

  return validation;
}

const outputPath = sampleOutputPath("validations-sample.xlsx");

const workbook = new Workbook();
const sheet = workbook.Worksheets[0];
sheet.Name = "Validation Sheet";

sheet.Cells["A1"].PutValue("Open");
sheet.Cells["B2"].PutValue(5);
sheet.Cells["C3"].PutValue(7);
sheet.Cells["E2"].PutValue(8);
sheet.Cells["G1"].PutValue("ABCDE");

const listValidationIndex = sheet.Validations.Add(CellArea.CreateCellArea("A1", "A3"));
const listValidation = sheet.Validations[listValidationIndex];
listValidation.Type = ValidationType.List;
listValidation.Formula1 = "\"Open,Closed\"";
listValidation.IgnoreBlank = true;
listValidation.InCellDropDown = true;
listValidation.ShowInput = true;
listValidation.InputTitle = "Status";
listValidation.InputMessage = "Pick a status";
listValidation.ShowError = true;
listValidation.ErrorTitle = "Invalid";
listValidation.ErrorMessage = "Choose from the list";

const decimalValidationIndex = sheet.Validations.Add(CellArea.CreateCellArea("B2", "C3"));
const decimalValidation = sheet.Validations[decimalValidationIndex];
decimalValidation.Type = ValidationType.Decimal;
decimalValidation.Operator = OperatorType.Between;
decimalValidation.Formula1 = "1.5";
decimalValidation.Formula2 = "9.5";
decimalValidation.ErrorTitle = "Range";
decimalValidation.ErrorMessage = "Enter 1.5-9.5";
decimalValidation.ShowError = true;
decimalValidation.AddArea(CellArea.CreateCellArea("E2", "E3"));

const customValidationIndex = sheet.Validations.Add(CellArea.CreateCellArea("G1", "G1"));
const customValidation = sheet.Validations[customValidationIndex];
customValidation.Type = ValidationType.Custom;
customValidation.AlertStyle = ValidationAlertType.Warning;
customValidation.Formula1 = "LEN(G1)<=5";
customValidation.ShowInput = true;
customValidation.InputTitle = "Code";
customValidation.InputMessage = "Up to 5 chars";

workbook.Save(outputPath);

const loaded = new Workbook(outputPath);
const loadedSheet = loaded.Worksheets["Validation Sheet"];
const a1Validation = requireValidation(loadedSheet.Validations.GetValidationInCell(0, 0), "A1");
const c2Validation = requireValidation(loadedSheet.Validations.GetValidationInCell(1, 2), "C2");
const g1Validation = requireValidation(loadedSheet.Validations.GetValidationInCell(0, 6), "G1");

console.log("Saved: " + outputPath);
console.log("Validation count: " + loadedSheet.Validations.Count);
console.log("A1 validation: " + ValidationType[a1Validation.Type] + " / " + a1Validation.Formula1);
console.log("C2 validation: " + ValidationType[c2Validation.Type] + " / " + c2Validation.Formula1 + " to " + c2Validation.Formula2);
console.log("G1 validation: " + ValidationType[g1Validation.Type]);
