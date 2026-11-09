import { WorksheetModel } from "./Core/WorksheetModel";
import { WorksheetProtectionModel } from "./Core/WorksheetProtectionModel";
import { DiagnosticSeverity } from "./enums";
import { InvalidFileFormatException } from "./InvalidFileFormatException";
import { LoadDiagnostics } from "./LoadDiagnostics";
import { LoadIssue } from "./LoadIssue";
import { LoadOptions } from "./LoadOptions";
import { WarningInfo } from "./WarningInfo";
import { XmlElement } from "./Xml/XmlWriter";
import { XmlNode, attribute, firstElement } from "./Xml/XmlReader";

export function buildSheetProtectionElement(worksheet: WorksheetModel): XmlElement | null {
  const protection = worksheet.Protection;
  if (!protection.HasStoredState()) {
    return null;
  }

  const element = XmlElement.create("sheetProtection");
  if (protection.IsProtected) {
    element.attr("sheet", 1);
  }

  setBoolAttributeWhenTrue(element, "objects", protection.Objects);
  setBoolAttributeWhenTrue(element, "scenarios", protection.Scenarios);
  setBoolAttributeWhenTrue(element, "formatCells", protection.FormatCells);
  setBoolAttributeWhenTrue(element, "formatColumns", protection.FormatColumns);
  setBoolAttributeWhenTrue(element, "formatRows", protection.FormatRows);
  setBoolAttributeWhenTrue(element, "insertColumns", protection.InsertColumns);
  setBoolAttributeWhenTrue(element, "insertRows", protection.InsertRows);
  setBoolAttributeWhenTrue(element, "insertHyperlinks", protection.InsertHyperlinks);
  setBoolAttributeWhenTrue(element, "deleteColumns", protection.DeleteColumns);
  setBoolAttributeWhenTrue(element, "deleteRows", protection.DeleteRows);
  setBoolAttributeWhenTrue(element, "selectLockedCells", protection.SelectLockedCells);
  setBoolAttributeWhenTrue(element, "sort", protection.Sort);
  setBoolAttributeWhenTrue(element, "autoFilter", protection.AutoFilter);
  setBoolAttributeWhenTrue(element, "pivotTables", protection.PivotTables);
  setBoolAttributeWhenTrue(element, "selectUnlockedCells", protection.SelectUnlockedCells);
  setStringAttributeWhenPresent(element, "password", protection.PasswordHash);
  setStringAttributeWhenPresent(element, "algorithmName", protection.AlgorithmName);
  setStringAttributeWhenPresent(element, "hashValue", protection.HashValue);
  setStringAttributeWhenPresent(element, "saltValue", protection.SaltValue);
  setStringAttributeWhenPresent(element, "spinCount", protection.SpinCount);
  return element;
}

export function loadWorksheetProtection(
  worksheetModel: WorksheetModel,
  worksheetRoot: XmlNode,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  worksheetModel.Protection.Clear();
  const protectionElement = firstElement(worksheetRoot, "sheetProtection");
  if (protectionElement === null) {
    return;
  }

  worksheetModel.Protection.IsProtected = parseProtectionBooleanAttribute(
    attribute(protectionElement, "sheet"),
    false,
    diagnostics,
    options,
    sheetName,
    "sheet",
  );
  worksheetModel.Protection.Objects = parseProtectionBooleanAttribute(
    attribute(protectionElement, "objects"),
    false,
    diagnostics,
    options,
    sheetName,
    "objects",
  );
  worksheetModel.Protection.Scenarios = parseProtectionBooleanAttribute(
    attribute(protectionElement, "scenarios"),
    false,
    diagnostics,
    options,
    sheetName,
    "scenarios",
  );
  worksheetModel.Protection.FormatCells = parseProtectionBooleanAttribute(
    attribute(protectionElement, "formatCells"),
    false,
    diagnostics,
    options,
    sheetName,
    "formatCells",
  );
  worksheetModel.Protection.FormatColumns = parseProtectionBooleanAttribute(
    attribute(protectionElement, "formatColumns"),
    false,
    diagnostics,
    options,
    sheetName,
    "formatColumns",
  );
  worksheetModel.Protection.FormatRows = parseProtectionBooleanAttribute(
    attribute(protectionElement, "formatRows"),
    false,
    diagnostics,
    options,
    sheetName,
    "formatRows",
  );
  worksheetModel.Protection.InsertColumns = parseProtectionBooleanAttribute(
    attribute(protectionElement, "insertColumns"),
    false,
    diagnostics,
    options,
    sheetName,
    "insertColumns",
  );
  worksheetModel.Protection.InsertRows = parseProtectionBooleanAttribute(
    attribute(protectionElement, "insertRows"),
    false,
    diagnostics,
    options,
    sheetName,
    "insertRows",
  );
  worksheetModel.Protection.InsertHyperlinks = parseProtectionBooleanAttribute(
    attribute(protectionElement, "insertHyperlinks"),
    false,
    diagnostics,
    options,
    sheetName,
    "insertHyperlinks",
  );
  worksheetModel.Protection.DeleteColumns = parseProtectionBooleanAttribute(
    attribute(protectionElement, "deleteColumns"),
    false,
    diagnostics,
    options,
    sheetName,
    "deleteColumns",
  );
  worksheetModel.Protection.DeleteRows = parseProtectionBooleanAttribute(
    attribute(protectionElement, "deleteRows"),
    false,
    diagnostics,
    options,
    sheetName,
    "deleteRows",
  );
  worksheetModel.Protection.SelectLockedCells = parseProtectionBooleanAttribute(
    attribute(protectionElement, "selectLockedCells"),
    false,
    diagnostics,
    options,
    sheetName,
    "selectLockedCells",
  );
  worksheetModel.Protection.Sort = parseProtectionBooleanAttribute(
    attribute(protectionElement, "sort"),
    false,
    diagnostics,
    options,
    sheetName,
    "sort",
  );
  worksheetModel.Protection.AutoFilter = parseProtectionBooleanAttribute(
    attribute(protectionElement, "autoFilter"),
    false,
    diagnostics,
    options,
    sheetName,
    "autoFilter",
  );
  worksheetModel.Protection.PivotTables = parseProtectionBooleanAttribute(
    attribute(protectionElement, "pivotTables"),
    false,
    diagnostics,
    options,
    sheetName,
    "pivotTables",
  );
  worksheetModel.Protection.SelectUnlockedCells = parseProtectionBooleanAttribute(
    attribute(protectionElement, "selectUnlockedCells"),
    false,
    diagnostics,
    options,
    sheetName,
    "selectUnlockedCells",
  );
  worksheetModel.Protection.PasswordHash = readOptionalAttribute(protectionElement, "password");
  worksheetModel.Protection.AlgorithmName = readOptionalAttribute(protectionElement, "algorithmName");
  worksheetModel.Protection.HashValue = readOptionalAttribute(protectionElement, "hashValue");
  worksheetModel.Protection.SaltValue = readOptionalAttribute(protectionElement, "saltValue");
  worksheetModel.Protection.SpinCount = readOptionalAttribute(protectionElement, "spinCount");

  if (!worksheetModel.Protection.IsProtected && hasEffectiveProtectionState(worksheetModel.Protection)) {
    worksheetModel.Protection.IsProtected = true;
  }
}

function hasEffectiveProtectionState(protection: WorksheetProtectionModel): boolean {
  if (
    protection.Objects ||
    protection.Scenarios ||
    protection.FormatCells ||
    protection.FormatColumns ||
    protection.FormatRows ||
    protection.InsertColumns ||
    protection.InsertRows ||
    protection.InsertHyperlinks ||
    protection.DeleteColumns ||
    protection.DeleteRows ||
    protection.Sort ||
    protection.AutoFilter ||
    protection.PivotTables
  ) {
    return true;
  }

  if (
    !isNullOrWhiteSpace(protection.PasswordHash) ||
    !isNullOrWhiteSpace(protection.AlgorithmName) ||
    !isNullOrWhiteSpace(protection.HashValue) ||
    !isNullOrWhiteSpace(protection.SaltValue) ||
    !isNullOrWhiteSpace(protection.SpinCount)
  ) {
    return true;
  }

  return false;
}

function parseProtectionBooleanAttribute(
  attributeValue: string | null,
  defaultValue: boolean,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  attributeName: string,
): boolean {
  if (attributeValue === null) {
    return defaultValue;
  }

  const rawValue = attributeValue.trim();
  if (rawValue === "0" || rawValue.toLowerCase() === "false") {
    return false;
  }

  if (rawValue === "1" || rawValue.toLowerCase() === "true") {
    return true;
  }

  if (options.StrictMode) {
    throw new InvalidFileFormatException(`Worksheet protection attribute '${attributeName}' is invalid.`);
  }

  const issue = new LoadIssue(
    "WS-L009",
    DiagnosticSeverity.Warning,
    `Worksheet protection attribute '${attributeName}' is invalid and the default value was used.`,
    false,
    true,
  );
  issue.SheetName = sheetName;
  addIssue(diagnostics, options, issue);
  return defaultValue;
}

function readOptionalAttribute(element: XmlNode, attributeName: string): string | null {
  const value = (attribute(element, attributeName) ?? "").trim();
  if (value.length === 0) {
    return null;
  }

  return value;
}

function setBoolAttributeWhenTrue(element: XmlElement, attributeName: string, value: boolean): void {
  if (value) {
    element.attr(attributeName, 1);
  }
}

function setStringAttributeWhenPresent(element: XmlElement, attributeName: string, value: string | null): void {
  if (!isNullOrWhiteSpace(value)) {
    element.attr(attributeName, value);
  }
}

function isNullOrWhiteSpace(value: string | null | undefined): value is null | undefined {
  return value === null || value === undefined || value.trim().length === 0;
}

function addIssue(diagnostics: LoadDiagnostics, options: LoadOptions, issue: LoadIssue): void {
  diagnostics.Add(issue);
  if (options.WarningCallback !== null) {
    const info = new WarningInfo();
    info.Code = issue.Code;
    info.Severity = issue.Severity;
    info.Message = issue.Message;
    info.DataLossRisk = issue.DataLossRisk;
    info.PartUri = issue.PartUri;
    info.SheetName = issue.SheetName;
    info.CellRef = issue.CellRef;
    info.RowIndex = issue.RowIndex;
    options.WarningCallback.Warning(info);
  }
}
