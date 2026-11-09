import { ColorValue } from "./Core/ColorValue";
import { PageSetupModel } from "./Core/PageSetupModel";
import { WorksheetModel } from "./Core/WorksheetModel";
import { WorksheetViewModel } from "./Core/WorksheetViewModel";
import { DiagnosticSeverity } from "./enums";
import { InvalidFileFormatException } from "./InvalidFileFormatException";
import { LoadDiagnostics } from "./LoadDiagnostics";
import { LoadIssue } from "./LoadIssue";
import { LoadOptions } from "./LoadOptions";
import { WarningInfo } from "./WarningInfo";
import { XmlElement } from "./Xml/XmlWriter";
import { XmlNode, attribute, firstElement } from "./Xml/XmlReader";

export function buildWorksheetSheetProperties(worksheet: WorksheetModel): XmlElement | null {
  const pageSetUpProperties = buildPageSetupPropertiesElement(worksheet.PageSetup);
  const tabColorElement = buildTabColorElement(worksheet.TabColor);
  if (pageSetUpProperties === null && tabColorElement === null) {
    return null;
  }

  const element = XmlElement.create("sheetPr");
  if (tabColorElement !== null) {
    element.child(tabColorElement);
  }

  if (pageSetUpProperties !== null) {
    element.child(pageSetUpProperties);
  }

  return element;
}

export function buildWorksheetViewsElement(worksheet: WorksheetModel): XmlElement | null {
  if (worksheetViewIsDefault(worksheet.View)) {
    return null;
  }

  const sheetView = XmlElement.create("sheetView").attr("workbookViewId", 0);

  if (!worksheet.View.ShowGridLines) {
    sheetView.attr("showGridLines", 0);
  }

  if (!worksheet.View.ShowRowColumnHeaders) {
    sheetView.attr("showRowColHeaders", 0);
  }

  if (!worksheet.View.ShowZeros) {
    sheetView.attr("showZeros", 0);
  }

  if (worksheet.View.RightToLeft) {
    sheetView.attr("rightToLeft", 1);
  }

  if (worksheet.View.ZoomScale !== 100) {
    sheetView.attr("zoomScale", worksheet.View.ZoomScale);
  }

  if (!isNullOrWhiteSpace(worksheet.View.ViewType)) {
    sheetView.attr("view", worksheet.View.ViewType);
  }

  if (!isNullOrWhiteSpace(worksheet.View.TopLeftCell)) {
    sheetView.attr("topLeftCell", worksheet.View.TopLeftCell);
  }

  if (worksheet.View.TabSelected) {
    sheetView.attr("tabSelected", 1);
  }

  const selection = buildSelectionElement(worksheet.View);
  if (selection !== null) {
    sheetView.child(selection);
  }

  return XmlElement.create("sheetViews").child(sheetView);
}

export function loadWorksheetViewSettings(
  worksheetModel: WorksheetModel,
  worksheetRoot: XmlNode,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  loadTabColor(worksheetModel, firstElement(worksheetRoot, "sheetPr"), diagnostics, options, sheetName);
  const sheetViews = firstElement(worksheetRoot, "sheetViews");
  loadSheetView(worksheetModel, firstElement(sheetViews, "sheetView"), diagnostics, options, sheetName);
}

function buildPageSetupPropertiesElement(pageSetup: PageSetupModel): XmlElement | null {
  if (!pageSetup.FitToPage) {
    return null;
  }

  return XmlElement.create("pageSetUpPr").attr("fitToPage", 1);
}

function buildTabColorElement(color: ColorValue | null): XmlElement | null {
  if (color === null) {
    return null;
  }

  return XmlElement.create("tabColor").attr("rgb", formatColor(color));
}

function worksheetViewIsDefault(view: WorksheetViewModel): boolean {
  return (
    view.ShowGridLines &&
    view.ShowRowColumnHeaders &&
    view.ShowZeros &&
    !view.RightToLeft &&
    view.ZoomScale === 100 &&
    isNullOrWhiteSpace(view.ViewType) &&
    isNullOrWhiteSpace(view.TopLeftCell) &&
    !view.TabSelected &&
    isNullOrWhiteSpace(view.SelectionActiveCell) &&
    isNullOrWhiteSpace(view.SelectionSqref)
  );
}

function loadTabColor(
  worksheetModel: WorksheetModel,
  sheetPropertiesElement: XmlNode | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  worksheetModel.TabColor = null;
  if (sheetPropertiesElement === null) {
    return;
  }

  const tabColor = firstElement(sheetPropertiesElement, "tabColor");
  if (tabColor === null) {
    return;
  }

  const rgb = (attribute(tabColor, "rgb") ?? "").trim();
  if (rgb.length === 0) {
    return;
  }

  const color = tryParseColor(rgb);
  if (color !== null) {
    worksheetModel.TabColor = color;
    return;
  }

  if (options.StrictMode) {
    throw new InvalidFileFormatException("Worksheet tab color is invalid.");
  }

  const issue = new LoadIssue(
    "WS-L007",
    DiagnosticSeverity.Warning,
    "Worksheet tab color is invalid and was ignored.",
    false,
    true,
  );
  issue.SheetName = sheetName;
  addIssue(diagnostics, options, issue);
}

function loadSheetView(
  worksheetModel: WorksheetModel,
  sheetViewElement: XmlNode | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  worksheetModel.View.ShowGridLines = true;
  worksheetModel.View.ShowRowColumnHeaders = true;
  worksheetModel.View.ShowZeros = true;
  worksheetModel.View.RightToLeft = false;
  worksheetModel.View.ZoomScale = 100;
  worksheetModel.View.ViewType = null;
  worksheetModel.View.TopLeftCell = null;
  worksheetModel.View.TabSelected = false;
  worksheetModel.View.SelectionActiveCell = null;
  worksheetModel.View.SelectionSqref = null;
  if (sheetViewElement === null) {
    return;
  }

  worksheetModel.View.ShowGridLines = parseBooleanViewAttribute(
    attribute(sheetViewElement, "showGridLines"),
    true,
    diagnostics,
    options,
    sheetName,
    "showGridLines",
  );
  worksheetModel.View.ShowRowColumnHeaders = parseBooleanViewAttribute(
    attribute(sheetViewElement, "showRowColHeaders"),
    true,
    diagnostics,
    options,
    sheetName,
    "showRowColHeaders",
  );
  worksheetModel.View.ShowZeros = parseBooleanViewAttribute(
    attribute(sheetViewElement, "showZeros"),
    true,
    diagnostics,
    options,
    sheetName,
    "showZeros",
  );
  worksheetModel.View.RightToLeft = parseBooleanViewAttribute(
    attribute(sheetViewElement, "rightToLeft"),
    false,
    diagnostics,
    options,
    sheetName,
    "rightToLeft",
  );
  worksheetModel.View.ZoomScale = parseZoomScale(attribute(sheetViewElement, "zoomScale"), diagnostics, options, sheetName);
  worksheetModel.View.ViewType = readOptionalTrimmedAttribute(sheetViewElement, "view");
  worksheetModel.View.TopLeftCell = readOptionalTrimmedAttribute(sheetViewElement, "topLeftCell");
  worksheetModel.View.TabSelected = parseBooleanViewAttribute(
    attribute(sheetViewElement, "tabSelected"),
    false,
    diagnostics,
    options,
    sheetName,
    "tabSelected",
  );
  loadPrimarySelection(worksheetModel.View, sheetViewElement);
}

function buildSelectionElement(view: WorksheetViewModel): XmlElement | null {
  if (isNullOrWhiteSpace(view.SelectionActiveCell) && isNullOrWhiteSpace(view.SelectionSqref)) {
    return null;
  }

  const selection = XmlElement.create("selection");
  if (!isNullOrWhiteSpace(view.SelectionActiveCell)) {
    selection.attr("activeCell", view.SelectionActiveCell);
  }

  if (!isNullOrWhiteSpace(view.SelectionSqref)) {
    selection.attr("sqref", view.SelectionSqref);
  }

  return selection;
}

function loadPrimarySelection(view: WorksheetViewModel, sheetViewElement: XmlNode): void {
  const selection = firstElement(sheetViewElement, "selection");
  if (selection === null) {
    return;
  }

  view.SelectionActiveCell = readOptionalTrimmedAttribute(selection, "activeCell");
  view.SelectionSqref = readOptionalTrimmedAttribute(selection, "sqref");
}

function readOptionalTrimmedAttribute(element: XmlNode, attributeName: string): string | null {
  const value = (attribute(element, attributeName) ?? "").trim();
  if (value.length === 0) {
    return null;
  }

  return value;
}

function parseBooleanViewAttribute(
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
    throw new InvalidFileFormatException(`Worksheet view attribute '${attributeName}' is invalid.`);
  }

  const issue = new LoadIssue(
    "WS-L008",
    DiagnosticSeverity.Warning,
    `Worksheet view attribute '${attributeName}' is invalid and the default value was used.`,
    false,
    true,
  );
  issue.SheetName = sheetName;
  addIssue(diagnostics, options, issue);
  return defaultValue;
}

function parseZoomScale(
  attributeValue: string | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): number {
  if (attributeValue === null) {
    return 100;
  }

  const value = parseInvariantInt(attributeValue);
  if (value !== null && value >= 10 && value <= 400) {
    return value;
  }

  if (options.StrictMode) {
    throw new InvalidFileFormatException("Worksheet zoomScale is invalid.");
  }

  const issue = new LoadIssue(
    "WS-L008",
    DiagnosticSeverity.Warning,
    "Worksheet zoomScale is invalid and the default value was used.",
    false,
    true,
  );
  issue.SheetName = sheetName;
  addIssue(diagnostics, options, issue);
  return 100;
}

function formatColor(color: ColorValue): string {
  return formatHex2(color.A) + formatHex2(color.R) + formatHex2(color.G) + formatHex2(color.B);
}

function formatHex2(value: number): string {
  return (value & 0xff).toString(16).toUpperCase().padStart(2, "0");
}

function tryParseColor(value: string): ColorValue | null {
  if (value.length !== 8 || !/^[0-9A-Fa-f]{8}$/.test(value)) {
    return null;
  }

  const a = parseInt(value.substring(0, 2), 16);
  const r = parseInt(value.substring(2, 2 + 2), 16);
  const g = parseInt(value.substring(4, 4 + 2), 16);
  const b = parseInt(value.substring(6, 6 + 2), 16);
  return new ColorValue(a, r, g, b);
}

function parseInvariantInt(text: string): number | null {
  const trimmed = text.trim();
  if (!/^[+-]?\d+$/.test(trimmed)) {
    return null;
  }

  const value = Number(trimmed);
  return Number.isSafeInteger(value) ? value : null;
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
