import { Workbook } from "../../aspose_cells";
import { OperatorType, ValidationType } from "../../aspose_cells";
import { attribute, documentRoot, elementChildren, firstElement, XmlNode } from "../../aspose_cells/Xml/XmlReader";
import { AssertEx, TemporaryDirectory, TestCase, TestRunner } from "../Shared/TestInfrastructure";
import { ValidationScenarioFactory } from "../Shared/ScenarioFactories";
import { createOpenXmlWorkbook, OpenXmlWorkbookReader } from "../Shared/OpenXmlPackage";

class ValidationSnapshot {
  public constructor(
    public readonly Sqref: string,
    public readonly Type: string,
    public readonly Operator: string,
    public readonly ErrorStyle: string,
    public readonly AllowBlank: boolean,
    public readonly ShowDropDown: boolean,
    public readonly ShowInputMessage: boolean,
    public readonly ShowErrorMessage: boolean,
    public readonly PromptTitle: string,
    public readonly Prompt: string,
    public readonly ErrorTitle: string,
    public readonly Error: string,
    public readonly Formula1: string,
    public readonly Formula2: string,
  ) {}
}

function dataValidationsMatchOpenXmlSdk(): void {
  const temp = new TemporaryDirectory("validation-openxml-write");
  const path = temp.getPath("validations.xlsx");

  const workbook = ValidationScenarioFactory.CreateValidationWorkbook();
  workbook.Save(path);

  const loaded = new Workbook(path);
  ValidationScenarioFactory.AssertValidations(loaded);

  const snapshot = readValidationSnapshots(path, "Validation Sheet");
  AssertEx.equal(3, snapshot.length);

  assertValidation(
    getValidationSnapshot(snapshot, "A1:A3"),
    "A1:A3",
    "list",
    "",
    "stop",
    true,
    false,
    true,
    true,
    "Status",
    "Pick a status",
    "Invalid",
    "Choose from the list",
    '"Open,Closed"',
    "",
  );
  assertValidation(
    getValidationSnapshot(snapshot, "B2:C3 E2:E3"),
    "B2:C3 E2:E3",
    "decimal",
    "between",
    "stop",
    false,
    true,
    false,
    true,
    "",
    "",
    "Range",
    "Enter 1.5-9.5",
    "1.5",
    "9.5",
  );
  assertValidation(
    getValidationSnapshot(snapshot, "G1"),
    "G1",
    "custom",
    "",
    "warning",
    false,
    false,
    true,
    false,
    "Code",
    "Up to 5 chars",
    "",
    "",
    "LEN(G1)<=5",
    "",
  );
}

function openXmlCreatedDataValidationsLoadInLibrary(): void {
  const temp = new TemporaryDirectory("validation-openxml-read");
  const path = temp.getPath("openxml-validations.xlsx");

  createOpenXmlValidationWorkbook(path);

  const loaded = new Workbook(path);
  const sheet = loaded.Worksheets[0];
  AssertEx.equal("Validation Sheet", sheet.Name);
  AssertEx.equal(2, sheet.Validations.Count);

  const listValidation = sheet.Validations[0];
  AssertEx.equal(ValidationType.List, listValidation.Type);
  AssertEx.equal('"One,Two"', listValidation.Formula1);
  AssertEx.true(listValidation.IgnoreBlank);
  AssertEx.true(listValidation.ShowInput);
  AssertEx.equal("Status", listValidation.InputTitle);
  AssertEx.equal("Pick one", listValidation.InputMessage);
  AssertEx.equal(1, listValidation.Areas.length);
  AssertEx.equal(0, listValidation.Areas[0].StartRow);
  AssertEx.equal(0, listValidation.Areas[0].StartColumn);
  AssertEx.equal(1, listValidation.Areas[0].EndRow);
  AssertEx.equal(0, listValidation.Areas[0].EndColumn);

  const decimalValidation = sheet.Validations[1];
  AssertEx.equal(ValidationType.Decimal, decimalValidation.Type);
  AssertEx.equal(OperatorType.Between, decimalValidation.Operator);
  AssertEx.equal("1", decimalValidation.Formula1);
  AssertEx.equal("9", decimalValidation.Formula2);
  AssertEx.false(decimalValidation.InCellDropDown);
  AssertEx.true(decimalValidation.ShowError);
  AssertEx.equal("Range", decimalValidation.ErrorTitle);
  AssertEx.equal("Enter 1-9", decimalValidation.ErrorMessage);
  AssertEx.equal(2, decimalValidation.Areas.length);
  AssertEx.equal(1, decimalValidation.Areas[0].StartRow);
  AssertEx.equal(1, decimalValidation.Areas[0].StartColumn);
  AssertEx.equal(2, decimalValidation.Areas[0].EndRow);
  AssertEx.equal(2, decimalValidation.Areas[0].EndColumn);
  AssertEx.equal(1, decimalValidation.Areas[1].StartRow);
  AssertEx.equal(4, decimalValidation.Areas[1].StartColumn);
  AssertEx.equal(2, decimalValidation.Areas[1].EndRow);
  AssertEx.equal(4, decimalValidation.Areas[1].EndColumn);

  const snapshot = readValidationSnapshots(path, "Validation Sheet");
  AssertEx.equal(2, snapshot.length);
  assertValidation(
    getValidationSnapshot(snapshot, "A1:A2"),
    "A1:A2",
    "list",
    "",
    "stop",
    true,
    false,
    true,
    false,
    "Status",
    "Pick one",
    "",
    "",
    '"One,Two"',
    "",
  );
  assertValidation(
    getValidationSnapshot(snapshot, "B2:C3 E2:E3"),
    "B2:C3 E2:E3",
    "decimal",
    "between",
    "stop",
    false,
    true,
    false,
    true,
    "",
    "",
    "Range",
    "Enter 1-9",
    "1",
    "9",
  );
}

function readValidationSnapshots(workbookPath: string, sheetName: string): ValidationSnapshot[] {
  const reader = new OpenXmlWorkbookReader(workbookPath);
  const worksheetXml = reader.sheetXmlByName(sheetName);
  const root = documentRoot(worksheetXml);
  const validations: ValidationSnapshot[] = [];

  for (const validation of elementChildren(firstElement(root, "dataValidations"), "dataValidation")) {
    validations.push(
      new ValidationSnapshot(
        (attribute(validation, "sqref") ?? "").toUpperCase(),
        attribute(validation, "type") ?? "",
        attribute(validation, "operator") ?? "",
        attribute(validation, "errorStyle") ?? "stop",
        parseBool(attribute(validation, "allowBlank")),
        parseBool(attribute(validation, "showDropDown")),
        parseBool(attribute(validation, "showInputMessage")),
        parseBool(attribute(validation, "showErrorMessage")),
        attribute(validation, "promptTitle") ?? "",
        attribute(validation, "prompt") ?? "",
        attribute(validation, "errorTitle") ?? "",
        attribute(validation, "error") ?? "",
        textOfChild(validation, "formula1"),
        textOfChild(validation, "formula2"),
      ),
    );
  }

  return validations;
}

function textOfChild(node: XmlNode, name: string): string {
  const child = firstElement(node, name);
  return child === null ? "" : (child.textContent ?? "");
}

function createOpenXmlValidationWorkbook(path: string): void {
  createOpenXmlWorkbook(
    path,
    "Validation Sheet",
    '<?xml version="1.0" encoding="utf-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:E3"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>One</t></is></c></row></sheetData><dataValidations count="2"><dataValidation type="list" allowBlank="1" showInputMessage="1" promptTitle="Status" prompt="Pick one" sqref="A1:A2"><formula1>"One,Two"</formula1></dataValidation><dataValidation type="decimal" operator="between" showDropDown="1" showErrorMessage="1" errorTitle="Range" error="Enter 1-9" sqref="B2:C3 E2:E3"><formula1>1</formula1><formula2>9</formula2></dataValidation></dataValidations></worksheet>',
  );
}

function getValidationSnapshot(snapshots: readonly ValidationSnapshot[], sqref: string): ValidationSnapshot {
  const found = snapshots.filter((snapshot) => snapshot.Sqref === sqref);
  AssertEx.equal(1, found.length);
  return found[0];
}

function assertValidation(
  actual: ValidationSnapshot,
  sqref: string,
  type: string,
  operatorName: string,
  errorStyle: string,
  allowBlank: boolean,
  showDropDown: boolean,
  showInputMessage: boolean,
  showErrorMessage: boolean,
  promptTitle: string,
  prompt: string,
  errorTitle: string,
  error: string,
  formula1: string,
  formula2: string,
): void {
  AssertEx.equal(sqref, actual.Sqref);
  AssertEx.equal(type, actual.Type);
  AssertEx.equal(operatorName, actual.Operator);
  AssertEx.equal(errorStyle, actual.ErrorStyle);
  AssertEx.equal(allowBlank, actual.AllowBlank);
  AssertEx.equal(showDropDown, actual.ShowDropDown);
  AssertEx.equal(showInputMessage, actual.ShowInputMessage);
  AssertEx.equal(showErrorMessage, actual.ShowErrorMessage);
  AssertEx.equal(promptTitle, actual.PromptTitle);
  AssertEx.equal(prompt, actual.Prompt);
  AssertEx.equal(errorTitle, actual.ErrorTitle);
  AssertEx.equal(error, actual.Error);
  AssertEx.equal(formula1, actual.Formula1);
  AssertEx.equal(formula2, actual.Formula2);
}

function parseBool(value: string | null): boolean {
  return value === "1" || (value ?? "").toLowerCase() === "true";
}

TestRunner.runAndExit("DataValidation.OpenXmlTests", [
  new TestCase("data_validations_match_openxml_sdk", dataValidationsMatchOpenXmlSdk),
  new TestCase("openxml_created_data_validations_load_in_library", openXmlCreatedDataValidationsLoadInLibrary),
]);
