import { CellAddress } from "./Core/CellAddress";
import { CellArea } from "./CellArea";
import { OperatorType } from "./OperatorType";
import { ValidationAlertType } from "./ValidationAlertType";
import { ValidationModel } from "./Core/ValidationModel";
import { ValidationType } from "./ValidationType";
import { WorksheetModel } from "./Core/WorksheetModel";
import { DiagnosticSeverity } from "./enums";
import { InvalidFileFormatException } from "./InvalidFileFormatException";
import { LoadDiagnostics } from "./LoadDiagnostics";
import { LoadIssue } from "./LoadIssue";
import { LoadOptions } from "./LoadOptions";
import { WarningInfo } from "./WarningInfo";
import { XmlElement } from "./Xml/XmlWriter";
import { attribute, elementChildren, firstElement, textOf, XmlNode } from "./Xml/XmlReader";
import { tryParseMergeReference } from "./XlsxWorkbookSerializerCommon";

/** Builds the worksheet "dataValidations" element, mirroring XlsxWorkbookValidations.BuildDataValidationsElement. */
export function buildDataValidationsElement(worksheet: WorksheetModel): XmlElement | null {
  const ordered = getOrderedValidations(worksheet.Validations);
  if (ordered.length === 0) {
    return null;
  }

  const container = XmlElement.create("dataValidations");
  let count = 0;

  for (let index = 0; index < ordered.length; index++) {
    const validation = ordered[index];
    if (validation.Areas.length === 0) {
      continue;
    }

    const element = XmlElement.create("dataValidation").attr("sqref", buildSqref(validation.Areas));

    const typeName = toValidationTypeName(validation.Type);
    if (!isEmptyString(typeName)) {
      element.attr("type", typeName);
    }

    const operatorName = toOperatorName(validation.Operator);
    if (!isEmptyString(operatorName)) {
      element.attr("operator", operatorName);
    }

    if (validation.AlertStyle !== ValidationAlertType.Stop) {
      element.attr("errorStyle", toAlertStyleName(validation.AlertStyle));
    }

    if (validation.IgnoreBlank) {
      element.boolAttr("allowBlank", true);
    }

    if (!validation.InCellDropDown) {
      element.boolAttr("showDropDown", true);
    }

    if (validation.ShowInput) {
      element.boolAttr("showInputMessage", true);
    }

    if (validation.ShowError) {
      element.boolAttr("showErrorMessage", true);
    }

    if (!isEmptyString(validation.ErrorTitle)) {
      element.attr("errorTitle", validation.ErrorTitle);
    }

    if (!isEmptyString(validation.ErrorMessage)) {
      element.attr("error", validation.ErrorMessage);
    }

    if (!isEmptyString(validation.InputTitle)) {
      element.attr("promptTitle", validation.InputTitle);
    }

    if (!isEmptyString(validation.InputMessage)) {
      element.attr("prompt", validation.InputMessage);
    }

    if (!isEmptyString(validation.Formula1)) {
      element.child(XmlElement.create("formula1").child(validation.Formula1));
    }

    if (!isEmptyString(validation.Formula2)) {
      element.child(XmlElement.create("formula2").child(validation.Formula2));
    }

    container.child(element);
    count++;
  }

  if (count === 0) {
    return null;
  }

  container.attr("count", count);
  return container;
}

/** Loads data validations from a worksheet root element. */
export function loadValidations(
  worksheetModel: WorksheetModel,
  worksheetRoot: XmlNode,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  worksheetModel.Validations.length = 0;

  const container = firstElement(worksheetRoot, "dataValidations");
  if (container === null) {
    return;
  }

  for (const validationElement of elementChildren(container, "dataValidation")) {
    const validation = new ValidationModel();
    const sqref = attribute(validationElement, "sqref");
    if (!tryLoadAreas(worksheetModel.Validations, validation, sqref, diagnostics, options, sheetName)) {
      continue;
    }

    validation.Type = parseValidationType(attribute(validationElement, "type"), diagnostics, options, sheetName, sqref);
    validation.Operator = parseOperatorType(attribute(validationElement, "operator"), diagnostics, options, sheetName, sqref);
    validation.AlertStyle = parseAlertStyle(attribute(validationElement, "errorStyle"), diagnostics, options, sheetName, sqref);
    validation.IgnoreBlank = parseBoolAttribute(validationElement, "allowBlank");
    validation.InCellDropDown = !parseBoolAttribute(validationElement, "showDropDown");
    validation.ShowInput = parseBoolAttribute(validationElement, "showInputMessage");
    validation.ShowError = parseBoolAttribute(validationElement, "showErrorMessage");
    validation.ErrorTitle = normalizeText(attribute(validationElement, "errorTitle"));
    validation.ErrorMessage = normalizeText(attribute(validationElement, "error"));
    validation.InputTitle = normalizeText(attribute(validationElement, "promptTitle"));
    validation.InputMessage = normalizeText(attribute(validationElement, "prompt"));
    validation.Formula1 = normalizeFormulaText(formulaText(firstElement(validationElement, "formula1")));
    validation.Formula2 = normalizeFormulaText(formulaText(firstElement(validationElement, "formula2")));

    worksheetModel.Validations.push(validation);
  }
}

/** Alias matching the loader name expected by the orchestrator. */
export const loadDataValidations = loadValidations;

function tryLoadAreas(
  existingValidations: ValidationModel[],
  candidate: ValidationModel,
  sqref: string | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): boolean {
  if (sqref === null || sqref.trim().length === 0) {
    if (options.StrictMode) {
      throw new InvalidFileFormatException("A data validation is missing sqref.");
    }

    addIssue(
      diagnostics,
      options,
      issueWithLocation(
        "DV-L002",
        DiagnosticSeverity.LossyRecoverable,
        "A data validation without any valid areas was dropped during load.",
        sheetName,
        null,
      ),
    );
    return false;
  }

  const tokens = sqref.split(/[ \t\r\n]+/).filter((token) => token.length > 0);
  let invalidAreaReported = false;
  let overlapReported = false;

  for (let index = 0; index < tokens.length; index++) {
    const token = tokens[index];
    const region = tryParseMergeReference(token);
    if (region === null) {
      if (options.StrictMode) {
        throw new InvalidFileFormatException("The data validation sqref '" + token + "' is invalid.");
      }

      if (!invalidAreaReported) {
        addIssue(
          diagnostics,
          options,
          issueWithLocation(
            "DV-L001",
            DiagnosticSeverity.LossyRecoverable,
            "One or more data validation areas were invalid and were dropped during load.",
            sheetName,
            token,
          ),
        );
        invalidAreaReported = true;
      }
      continue;
    }

    const area = new CellArea(region.FirstRow, region.FirstColumn, region.TotalRows, region.TotalColumns);
    if (conflictsWithExisting(existingValidations, candidate, area)) {
      if (options.StrictMode) {
        throw new InvalidFileFormatException("The data validation sqref '" + token + "' overlaps an existing validation.");
      }

      if (!overlapReported) {
        addIssue(
          diagnostics,
          options,
          issueWithLocation(
            "DV-L003",
            DiagnosticSeverity.LossyRecoverable,
            "Overlapping data validation areas were normalized during load.",
            sheetName,
            token,
          ),
        );
        overlapReported = true;
      }
      continue;
    }

    candidate.Areas.push(area);
  }

  candidate.Areas.sort(compareAreas);
  if (candidate.Areas.length === 0) {
    addIssue(
      diagnostics,
      options,
      issueWithLocation(
        "DV-L002",
        DiagnosticSeverity.LossyRecoverable,
        "A data validation without any valid areas was dropped during load.",
        sheetName,
        sqref,
      ),
    );
    return false;
  }

  return true;
}

function conflictsWithExisting(existingValidations: ValidationModel[], candidate: ValidationModel, area: CellArea): boolean {
  for (const validation of existingValidations) {
    for (const existingArea of validation.Areas) {
      if (areasOverlap(existingArea, area)) {
        return true;
      }
    }
  }

  for (const candidateArea of candidate.Areas) {
    if (areasOverlap(candidateArea, area)) {
      return true;
    }
  }

  return false;
}

function getOrderedValidations(validations: ValidationModel[]): ValidationModel[] {
  const ordered: ValidationModel[] = [];
  for (const validation of validations) {
    if (validation.Areas.length === 0) {
      continue;
    }

    validation.Areas.sort(compareAreas);
    ordered.push(validation);
  }

  ordered.sort((left, right) => compareAreas(left.Areas[0], right.Areas[0]));
  return ordered;
}

function compareAreas(left: CellArea, right: CellArea): number {
  if (left.StartRow !== right.StartRow) {
    return left.StartRow - right.StartRow;
  }

  if (left.StartColumn !== right.StartColumn) {
    return left.StartColumn - right.StartColumn;
  }

  if (left.TotalRows !== right.TotalRows) {
    return left.TotalRows - right.TotalRows;
  }

  return left.TotalColumns - right.TotalColumns;
}

function areasOverlap(left: CellArea, right: CellArea): boolean {
  return (
    left.StartRow <= right.EndRow &&
    right.StartRow <= left.EndRow &&
    left.StartColumn <= right.EndColumn &&
    right.StartColumn <= left.EndColumn
  );
}

function buildSqref(areas: CellArea[]): string {
  return areas.map(toAreaReference).join(" ");
}

function toAreaReference(area: CellArea): string {
  const first = new CellAddress(area.FirstRow, area.FirstColumn).ToString();
  if (area.TotalRows === 1 && area.TotalColumns === 1) {
    return first;
  }

  const last = new CellAddress(area.FirstRow + area.TotalRows - 1, area.FirstColumn + area.TotalColumns - 1).ToString();
  return first + ":" + last;
}

function parseValidationType(
  text: string | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  cellRef: string | null,
): ValidationType {
  switch (normalizeToken(text)) {
    case null:
      return ValidationType.AnyValue;
    case "whole":
      return ValidationType.WholeNumber;
    case "decimal":
      return ValidationType.Decimal;
    case "list":
      return ValidationType.List;
    case "date":
      return ValidationType.Date;
    case "time":
      return ValidationType.Time;
    case "textlength":
      return ValidationType.TextLength;
    case "custom":
      return ValidationType.Custom;
    default:
      return warnAndReturnAnyValue("type", text, diagnostics, options, sheetName, cellRef);
  }
}

function parseOperatorType(
  text: string | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  cellRef: string | null,
): OperatorType {
  switch (normalizeToken(text)) {
    case null:
      return OperatorType.None;
    case "between":
      return OperatorType.Between;
    case "equal":
      return OperatorType.Equal;
    case "greaterthan":
      return OperatorType.GreaterThan;
    case "greaterthanorequal":
      return OperatorType.GreaterOrEqual;
    case "lessthan":
      return OperatorType.LessThan;
    case "lessthanorequal":
      return OperatorType.LessOrEqual;
    case "notbetween":
      return OperatorType.NotBetween;
    case "notequal":
      return OperatorType.NotEqual;
    default:
      return warnAndReturnNone("operator", text, diagnostics, options, sheetName, cellRef);
  }
}

function parseAlertStyle(
  text: string | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  cellRef: string | null,
): ValidationAlertType {
  switch (normalizeToken(text)) {
    case null:
    case "stop":
      return ValidationAlertType.Stop;
    case "warning":
      return ValidationAlertType.Warning;
    case "information":
      return ValidationAlertType.Information;
    default:
      if (options.StrictMode) {
        throw new InvalidFileFormatException("The data validation errorStyle '" + (text ?? "") + "' is invalid.");
      }

      addIssue(
        diagnostics,
        options,
        issueWithLocation(
          "DV-R001",
          DiagnosticSeverity.Recoverable,
          "An unknown data validation errorStyle was normalized during load.",
          sheetName,
          cellRef,
          false,
        ),
      );
      return ValidationAlertType.Stop;
  }
}

function warnAndReturnAnyValue(
  attributeName: string,
  text: string | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  cellRef: string | null,
): ValidationType {
  if (options.StrictMode) {
    throw new InvalidFileFormatException("The data validation " + attributeName + " '" + (text ?? "") + "' is invalid.");
  }

  addIssue(
    diagnostics,
    options,
    issueWithLocation(
      "DV-R001",
      DiagnosticSeverity.Recoverable,
      "An unknown data validation " + attributeName + " was normalized during load.",
      sheetName,
      cellRef,
      false,
    ),
  );
  return ValidationType.AnyValue;
}

function warnAndReturnNone(
  attributeName: string,
  text: string | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  cellRef: string | null,
): OperatorType {
  if (options.StrictMode) {
    throw new InvalidFileFormatException("The data validation " + attributeName + " '" + (text ?? "") + "' is invalid.");
  }

  addIssue(
    diagnostics,
    options,
    issueWithLocation(
      "DV-R001",
      DiagnosticSeverity.Recoverable,
      "An unknown data validation " + attributeName + " was normalized during load.",
      sheetName,
      cellRef,
      false,
    ),
  );
  return OperatorType.None;
}

function toValidationTypeName(type: ValidationType): string | null {
  switch (type) {
    case ValidationType.WholeNumber:
      return "whole";
    case ValidationType.Decimal:
      return "decimal";
    case ValidationType.List:
      return "list";
    case ValidationType.Date:
      return "date";
    case ValidationType.Time:
      return "time";
    case ValidationType.TextLength:
      return "textLength";
    case ValidationType.Custom:
      return "custom";
    default:
      return null;
  }
}

function toOperatorName(type: OperatorType): string | null {
  switch (type) {
    case OperatorType.Between:
      return "between";
    case OperatorType.Equal:
      return "equal";
    case OperatorType.GreaterThan:
      return "greaterThan";
    case OperatorType.GreaterOrEqual:
      return "greaterThanOrEqual";
    case OperatorType.LessThan:
      return "lessThan";
    case OperatorType.LessOrEqual:
      return "lessThanOrEqual";
    case OperatorType.NotBetween:
      return "notBetween";
    case OperatorType.NotEqual:
      return "notEqual";
    default:
      return null;
  }
}

function toAlertStyleName(type: ValidationAlertType): string {
  switch (type) {
    case ValidationAlertType.Warning:
      return "warning";
    case ValidationAlertType.Information:
      return "information";
    default:
      return "stop";
  }
}

function isEmptyString(value: string | null | undefined): boolean {
  return value === null || value === undefined || value.length === 0;
}

function normalizeToken(value: string | null): string | null {
  if (value === null) {
    return null;
  }

  const trimmed = value.trim();
  return trimmed.length === 0 ? null : trimmed.toLowerCase();
}

function normalizeFormulaText(value: string | null): string | null {
  if (value === null) {
    return null;
  }

  const trimmed = value.trim();
  if (trimmed.length === 0) {
    return null;
  }

  return trimmed[0] === "=" ? trimmed.substring(1) : trimmed;
}

function normalizeText(value: string | null): string | null {
  return value === null || value.length === 0 ? null : value;
}

function formulaText(element: XmlNode | null): string | null {
  return element === null ? null : textOf(element);
}

function parseBoolAttribute(node: XmlNode, name: string): boolean {
  const value = attribute(node, name);
  if (value === null) {
    return false;
  }

  return value === "1" || value.toLowerCase() === "true";
}

function issueWithLocation(
  code: string,
  severity: DiagnosticSeverity,
  message: string,
  sheetName: string,
  cellRef: string | null,
  dataLossRisk = true,
): LoadIssue {
  const issue = new LoadIssue(code, severity, message, true, dataLossRisk);
  issue.SheetName = sheetName;
  issue.CellRef = cellRef;
  return issue;
}

function addIssue(diagnostics: LoadDiagnostics, options: LoadOptions, issue: LoadIssue): void {
  diagnostics.Add(issue);
  const callback = options.WarningCallback;
  if (callback !== null) {
    const warning = new WarningInfo();
    warning.Code = issue.Code;
    warning.Severity = issue.Severity;
    warning.Message = issue.Message;
    warning.DataLossRisk = issue.DataLossRisk;
    warning.PartUri = issue.PartUri;
    warning.SheetName = issue.SheetName;
    warning.CellRef = issue.CellRef;
    warning.RowIndex = issue.RowIndex;
    callback.Warning(warning);
  }
}
