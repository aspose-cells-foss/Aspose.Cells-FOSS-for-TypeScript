import { CellAddress } from "./Core/CellAddress";
import { DefinedNameModel } from "./Core/DefinedNameModel";
import { WorkbookModel } from "./Core/WorkbookModel";
import { DiagnosticSeverity } from "./enums";
import { CellsException } from "./CellsException";
import { InvalidFileFormatException } from "./InvalidFileFormatException";
import { LoadDiagnostics } from "./LoadDiagnostics";
import { LoadIssue } from "./LoadIssue";
import { LoadOptions } from "./LoadOptions";
import { addIssue } from "./XlsxWorkbookArchiveHelpers";
import { XmlElement } from "./Xml/XmlWriter";
import { XmlNode, attribute, firstElement, textOf } from "./Xml/XmlReader";
import { buildPageSetupDefinedNames } from "./XlsxWorkbookPageSetup";

const PRINT_AREA_DEFINED_NAME = "_xlnm.Print_Area";
const PRINT_TITLES_DEFINED_NAME = "_xlnm.Print_Titles";
const FILTER_DATABASE_DEFINED_NAME = "_xlnm._FilterDatabase";

/** Builds the definedNames element covering page setup, autofilter and user defined names. */
export function buildDefinedNames(model: WorkbookModel): XmlElement | null {
  const definedNames: XmlElement[] = [];
  const pageSetupResult = buildPageSetupDefinedNames(model) as unknown as XmlElement | XmlElement[] | null;
  for (const element of collectDefinedNameElements(pageSetupResult)) {
    definedNames.push(element);
  }

  for (const element of buildAutoFilterDefinedNames(model)) {
    definedNames.push(element);
  }

  for (let index = 0; index < model.DefinedNames.length; index++) {
    const definedName = model.DefinedNames[index]!;
    const element = XmlElement.create("definedName")
      .attr("name", definedName.Name)
      .child(definedName.Formula);
    if (definedName.LocalSheetIndex !== null) {
      element.attr("localSheetId", definedName.LocalSheetIndex.toString());
    }

    if (definedName.Hidden) {
      element.attr("hidden", 1);
    }

    if (definedName.Comment.length > 0) {
      element.attr("comment", definedName.Comment);
    }

    definedNames.push(element);
  }

  if (definedNames.length === 0) {
    return null;
  }

  return XmlElement.create("definedNames").children(definedNames);
}

/** Loads workbook-scoped and sheet-scoped user defined names. */
export function loadWorkbookDefinedNames(
  workbookRoot: XmlNode,
  workbookModel: WorkbookModel,
  sheetCount: number,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
): void {
  const container = firstElement(workbookRoot, "definedNames");
  const elements: XmlNode[] = [];
  const containerChildren = container === null ? null : container.childNodes;
  if (containerChildren !== null) {
    for (let index = 0; index < containerChildren.length; index++) {
      const child = containerChildren[index]!;
      const name = child.nodeName;
      if (name === "definedName" || name.endsWith(":definedName")) {
        elements.push(child);
      }
    }
  }

  for (const element of elements) {
    const name = (attribute(element, "name") ?? "").trim();
    if (isReservedDefinedName(name)) {
      continue;
    }

    if (name.length === 0) {
      handleInvalidDefinedName(options, "Workbook defined name is missing a valid name.");
      addInvalidDefinedNameIssue(
        diagnostics,
        options,
        "Workbook defined name is missing a valid name and was ignored.",
      );
      continue;
    }

    let localSheetIndex: number | null = null;
    const localSheetAttribute = attribute(element, "localSheetId");
    if (localSheetAttribute !== null) {
      localSheetIndex = parseStrictInt(localSheetAttribute);
      if (localSheetIndex === null || localSheetIndex < 0 || localSheetIndex >= sheetCount) {
        handleInvalidDefinedName(
          options,
          "Workbook defined name '" + name + "' has an invalid localSheetId.",
        );
        addInvalidDefinedNameIssue(
          diagnostics,
          options,
          "Workbook defined name '" + name + "' has an invalid localSheetId and was ignored.",
        );
        continue;
      }
    }

    let formula: string;
    try {
      formula = normaliseDefinedNameFormula(textOf(element));
    } catch (error) {
      if (error instanceof CellsException) {
        if (options.StrictMode) {
          throw new InvalidFileFormatException("Workbook defined name '" + name + "' is invalid.", error);
        }

        addInvalidDefinedNameIssue(
          diagnostics,
          options,
          "Workbook defined name '" + name + "' has an invalid formula and was ignored.",
        );
        continue;
      }

      throw error;
    }

    if (containsDuplicate(workbookModel.DefinedNames, name, localSheetIndex)) {
      handleInvalidDefinedName(
        options,
        "Workbook defined name '" + name + "' is duplicated in the same scope.",
      );
      addInvalidDefinedNameIssue(
        diagnostics,
        options,
        "Workbook defined name '" + name + "' duplicates an existing scope and was ignored.",
      );
      continue;
    }

    const model = new DefinedNameModel();
    model.Name = name;
    model.Formula = formula;
    model.LocalSheetIndex = localSheetIndex;
    model.Hidden = parseBoolAttribute(element, "hidden");
    model.Comment = normaliseDefinedNameComment(attribute(element, "comment"));
    workbookModel.DefinedNames.push(model);
  }
}

function collectDefinedNameElements(value: XmlElement | XmlElement[] | null): XmlElement[] {
  if (value === null) {
    return [];
  }

  if (Array.isArray(value)) {
    return value;
  }

  const result: XmlElement[] = [];
  for (const child of value.childrenItems) {
    if (typeof child !== "string") {
      result.push(child);
    }
  }

  return result;
}

function containsDuplicate(
  definedNames: readonly DefinedNameModel[],
  name: string,
  localSheetIndex: number | null,
): boolean {
  for (const existing of definedNames) {
    if (existing.Name.toLowerCase() !== name.toLowerCase()) {
      continue;
    }

    if (sameScope(existing.LocalSheetIndex, localSheetIndex)) {
      return true;
    }
  }

  return false;
}

function sameScope(left: number | null, right: number | null): boolean {
  if (left === null && right === null) {
    return true;
  }

  return left !== null && right !== null && left === right;
}

function buildAutoFilterDefinedNames(model: WorkbookModel): XmlElement[] {
  const definedNames: XmlElement[] = [];
  for (let sheetIndex = 0; sheetIndex < model.Worksheets.length; sheetIndex++) {
    const worksheet = model.Worksheets[sheetIndex]!;
    if (worksheet.AutoFilter.Range.length === 0) {
      continue;
    }

    definedNames.push(
      XmlElement.create("definedName")
        .attr("name", FILTER_DATABASE_DEFINED_NAME)
        .attr("localSheetId", sheetIndex)
        .attr("hidden", 1)
        .child(qualifyAutoFilterRange(worksheet.Name, worksheet.AutoFilter.Range)),
    );
  }

  return definedNames;
}

function qualifyAutoFilterRange(sheetName: string, range: string): string {
  const parts = range.split(":");
  if (parts.length === 1) {
    return quoteWorksheetName(sheetName) + "!" + toAbsoluteCellReference(parts[0]!);
  }

  if (parts.length !== 2) {
    throw new CellsException("AutoFilter range is invalid.");
  }

  return (
    quoteWorksheetName(sheetName) +
    "!" +
    toAbsoluteCellReference(parts[0]!) +
    ":" +
    toAbsoluteCellReference(parts[1]!)
  );
}

function quoteWorksheetName(sheetName: string): string {
  return "'" + sheetName.replace(/'/g, "''") + "'";
}

function toAbsoluteCellReference(value: string): string {
  let address: CellAddress;
  try {
    address = CellAddress.Parse(value);
  } catch {
    throw new CellsException("AutoFilter range is invalid.");
  }

  const reference = address.ToString();
  let splitIndex = 0;
  while (splitIndex < reference.length && /[A-Za-z]/.test(reference[splitIndex]!)) {
    splitIndex++;
  }

  return "$" + reference.substring(0, splitIndex) + "$" + reference.substring(splitIndex);
}

/** Trims a defined name formula and removes a leading '='. */
export function normaliseDefinedNameFormula(formula: string | null | undefined): string {
  let normalized = (formula ?? "").trim();
  if (normalized.startsWith("=")) {
    normalized = normalized.substring(1).trim();
  }

  if (normalized.length === 0) {
    throw new CellsException("Defined name formula must be non-empty.");
  }

  return normalized;
}

/** Trims a defined name comment. */
export function normaliseDefinedNameComment(comment: string | null | undefined): string {
  return (comment ?? "").trim();
}

function isReservedDefinedName(name: string): boolean {
  const lower = name.toLowerCase();
  return (
    lower === PRINT_AREA_DEFINED_NAME.toLowerCase() ||
    lower === PRINT_TITLES_DEFINED_NAME.toLowerCase() ||
    lower === FILTER_DATABASE_DEFINED_NAME.toLowerCase()
  );
}

function parseStrictInt(value: string): number | null {
  const trimmed = value.trim();
  if (!/^[+-]?\d+$/.test(trimmed)) {
    return null;
  }

  return parseInt(trimmed, 10);
}

function parseBoolAttribute(element: XmlNode, name: string): boolean {
  const value = attribute(element, name);
  if (value === null) {
    return false;
  }

  const normalized = value.trim().toLowerCase();
  return normalized === "1" || normalized === "true";
}

function addInvalidDefinedNameIssue(
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  message: string,
): void {
  addIssue(diagnostics, options, new LoadIssue("WB-L002", DiagnosticSeverity.Warning, message, false, true));
}

function handleInvalidDefinedName(options: LoadOptions, message: string): void {
  if (options.StrictMode) {
    throw new InvalidFileFormatException(message);
  }
}
