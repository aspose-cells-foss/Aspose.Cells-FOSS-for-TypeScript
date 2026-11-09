import { CellAddress } from "./Core/CellAddress";
import { PageMarginsModel } from "./Core/PageMarginsModel";
import { PageOrientation } from "./Core/PageOrientation";
import { PageSetupModel } from "./Core/PageSetupModel";
import { WorkbookModel } from "./Core/WorkbookModel";
import { WorksheetModel } from "./Core/WorksheetModel";
import { CellsException } from "./CellsException";
import { DiagnosticSeverity } from "./enums";
import { InvalidFileFormatException } from "./InvalidFileFormatException";
import { LoadDiagnostics } from "./LoadDiagnostics";
import { LoadIssue } from "./LoadIssue";
import { LoadOptions } from "./LoadOptions";
import { WarningInfo } from "./WarningInfo";
import { XmlElement } from "./Xml/XmlWriter";
import { elementChildren, firstElement, textOf, XmlNode } from "./Xml/XmlReader";

const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

const MAX_SPREADSHEET_ROW = 1048575;
const MAX_SPREADSHEET_COLUMN = 16383;

export const PRINT_AREA_DEFINED_NAME = "_xlnm.Print_Area";
export const PRINT_TITLES_DEFINED_NAME = "_xlnm.Print_Titles";

/** Holds the worksheet-scoped defined names that affect page setup. */
export class WorksheetDefinedNamesState {
  public PrintArea: string | null = null;
  public PrintTitleRows: string | null = null;
  public PrintTitleColumns: string | null = null;
}

export function buildPageSetupDefinedNames(model: WorkbookModel): XmlElement | null {
  const definedNames: XmlElement[] = [];
  for (let sheetIndex = 0; sheetIndex < model.Worksheets.length; sheetIndex++) {
    const worksheet = model.Worksheets[sheetIndex];
    const printArea = normalizePrintAreaList(worksheet.PageSetup.PrintArea, worksheet.Name);
    if (printArea.length > 0) {
      definedNames.push(
        XmlElement.create("definedName")
          .attr("name", PRINT_AREA_DEFINED_NAME)
          .attr("localSheetId", sheetIndex)
          .child(printArea),
      );
    }

    const printTitles = buildPrintTitlesDefinedNameText(worksheet.PageSetup, worksheet.Name);
    if (printTitles.length > 0) {
      definedNames.push(
        XmlElement.create("definedName")
          .attr("name", PRINT_TITLES_DEFINED_NAME)
          .attr("localSheetId", sheetIndex)
          .child(printTitles),
      );
    }
  }

  if (definedNames.length === 0) {
    return null;
  }

  const root = XmlElement.createWithNamespace("definedNames", MAIN_NS);
  for (const definedName of definedNames) {
    root.child(definedName);
  }

  return root;
}

export function loadWorksheetDefinedNames(
  workbookRoot: XmlNode,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
): Map<number, WorksheetDefinedNamesState> {
  const states = new Map<number, WorksheetDefinedNamesState>();
  const definedNamesElement = firstElement(workbookRoot, "definedNames");
  if (definedNamesElement === null) {
    return states;
  }

  for (const definedName of elementChildren(definedNamesElement, "definedName")) {
    const name = (attributeOrNull(definedName, "name") ?? "").trim();
    if (!equalsIgnoreCase(name, PRINT_AREA_DEFINED_NAME) && !equalsIgnoreCase(name, PRINT_TITLES_DEFINED_NAME)) {
      continue;
    }

    const localSheetId = parseIntAttribute(attributeOrNull(definedName, "localSheetId"));
    if (localSheetId === null || localSheetId < 0) {
      addIssue(
        diagnostics,
        options,
        makeIssue(
          "PG-L004",
          DiagnosticSeverity.Warning,
          `Defined name '${name}' is missing a valid localSheetId and was ignored.`,
          false,
          true,
          null,
        ),
      );
      continue;
    }

    let state = states.get(localSheetId);
    if (state === undefined) {
      state = new WorksheetDefinedNamesState();
      states.set(localSheetId, state);
    }

    const text = textOf(definedName).trim();
    if (text.length === 0) {
      continue;
    }

    try {
      if (equalsIgnoreCase(name, PRINT_AREA_DEFINED_NAME)) {
        state.PrintArea = normalizeLoadedPrintArea(text);
      } else {
        parseLoadedPrintTitles(text, state);
      }
    } catch (exception) {
      if (exception instanceof CellsException) {
        if (options.StrictMode) {
          throw new InvalidFileFormatException(`Defined name '${name}' is invalid.`, exception);
        }

        addIssue(
          diagnostics,
          options,
          makeIssue(
            "PG-L004",
            DiagnosticSeverity.Warning,
            `Defined name '${name}' is invalid and was ignored.`,
            false,
            true,
            null,
          ),
        );
      } else {
        throw exception;
      }
    }
  }

  return states;
}

export function applyWorksheetDefinedNames(
  worksheetModel: WorksheetModel,
  definedNamesState: WorksheetDefinedNamesState | null,
): void {
  if (definedNamesState === null) {
    return;
  }

  worksheetModel.PageSetup.PrintArea = definedNamesState.PrintArea;
  worksheetModel.PageSetup.PrintTitleRows = definedNamesState.PrintTitleRows;
  worksheetModel.PageSetup.PrintTitleColumns = definedNamesState.PrintTitleColumns;
}

export function buildSheetProperties(pageSetup: PageSetupModel): XmlElement | null {
  if (!pageSetup.FitToPage) {
    return null;
  }

  return XmlElement.createWithNamespace("sheetPr", MAIN_NS).child(
    XmlElement.create("pageSetUpPr").attr("fitToPage", 1),
  );
}

export function buildPrintOptionsElement(pageSetup: PageSetupModel): XmlElement | null {
  if (
    !pageSetup.PrintOptions.GridLines &&
    !pageSetup.PrintOptions.Headings &&
    !pageSetup.PrintOptions.HorizontalCentered &&
    !pageSetup.PrintOptions.VerticalCentered
  ) {
    return null;
  }

  const element = XmlElement.createWithNamespace("printOptions", MAIN_NS);
  if (pageSetup.PrintOptions.Headings) {
    element.attr("headings", 1);
  }

  if (pageSetup.PrintOptions.GridLines) {
    element.attr("gridLines", 1);
    element.attr("gridLinesSet", 1);
  }

  if (pageSetup.PrintOptions.HorizontalCentered) {
    element.attr("horizontalCentered", 1);
  }

  if (pageSetup.PrintOptions.VerticalCentered) {
    element.attr("verticalCentered", 1);
  }

  return element;
}

export function buildPageMarginsElement(pageSetup: PageSetupModel): XmlElement | null {
  if (marginsEqual(pageSetup.Margins, new PageMarginsModel())) {
    return null;
  }

  return XmlElement.createWithNamespace("pageMargins", MAIN_NS)
    .attr("left", formatNumber(pageSetup.Margins.Left))
    .attr("right", formatNumber(pageSetup.Margins.Right))
    .attr("top", formatNumber(pageSetup.Margins.Top))
    .attr("bottom", formatNumber(pageSetup.Margins.Bottom))
    .attr("header", formatNumber(pageSetup.Margins.Header))
    .attr("footer", formatNumber(pageSetup.Margins.Footer));
}

export function buildPageSetupElement(pageSetup: PageSetupModel): XmlElement | null {
  if (
    pageSetup.PaperSize === 0 &&
    pageSetup.Orientation === PageOrientation.Default &&
    pageSetup.FirstPageNumber === null &&
    pageSetup.Scale === null &&
    pageSetup.FitToWidth === null &&
    pageSetup.FitToHeight === null &&
    !pageSetup.PageOrderOverThenDown
  ) {
    return null;
  }

  const element = XmlElement.createWithNamespace("pageSetup", MAIN_NS);
  if (pageSetup.PageOrderOverThenDown) {
    element.attr("pageOrder", "overThenDown");
  }

  if (pageSetup.PaperSize > 0) {
    element.attr("paperSize", pageSetup.PaperSize);
  }

  if (pageSetup.Scale !== null) {
    element.attr("scale", pageSetup.Scale);
  }

  if (pageSetup.FitToWidth !== null) {
    element.attr("fitToWidth", pageSetup.FitToWidth);
  }

  if (pageSetup.FitToHeight !== null) {
    element.attr("fitToHeight", pageSetup.FitToHeight);
  }

  if (pageSetup.FirstPageNumber !== null) {
    element.attr("firstPageNumber", pageSetup.FirstPageNumber);
    element.attr("useFirstPageNumber", 1);
  }

  if (pageSetup.Orientation === PageOrientation.Portrait) {
    element.attr("orientation", "portrait");
  } else if (pageSetup.Orientation === PageOrientation.Landscape) {
    element.attr("orientation", "landscape");
  }

  return element;
}

export function buildHeaderFooterElement(pageSetup: PageSetupModel): XmlElement | null {
  const headerFooter = pageSetup.HeaderFooter;
  const hasHeaderFooter =
    isNotEmpty(headerFooter.LeftHeader) ||
    isNotEmpty(headerFooter.CenterHeader) ||
    isNotEmpty(headerFooter.RightHeader) ||
    isNotEmpty(headerFooter.LeftFooter) ||
    isNotEmpty(headerFooter.CenterFooter) ||
    isNotEmpty(headerFooter.RightFooter);

  if (!hasHeaderFooter) {
    return null;
  }

  const element = XmlElement.createWithNamespace("headerFooter", MAIN_NS);
  if (
    isNotEmpty(headerFooter.LeftHeader) ||
    isNotEmpty(headerFooter.CenterHeader) ||
    isNotEmpty(headerFooter.RightHeader)
  ) {
    element.child(
      XmlElement.create("oddHeader").child(
        composeHeaderFooterText(headerFooter.LeftHeader, headerFooter.CenterHeader, headerFooter.RightHeader),
      ),
    );
  }

  if (
    isNotEmpty(headerFooter.LeftFooter) ||
    isNotEmpty(headerFooter.CenterFooter) ||
    isNotEmpty(headerFooter.RightFooter)
  ) {
    element.child(
      XmlElement.create("oddFooter").child(
        composeHeaderFooterText(headerFooter.LeftFooter, headerFooter.CenterFooter, headerFooter.RightFooter),
      ),
    );
  }

  return element;
}

export function buildRowBreaksElement(pageSetup: PageSetupModel): XmlElement | null {
  return buildBreaksElement("rowBreaks", pageSetup.HorizontalPageBreaks, MAX_SPREADSHEET_COLUMN);
}

export function buildColumnBreaksElement(pageSetup: PageSetupModel): XmlElement | null {
  return buildBreaksElement("colBreaks", pageSetup.VerticalPageBreaks, MAX_SPREADSHEET_ROW);
}

function buildBreaksElement(elementName: string, breaks: readonly number[], maxIndex: number): XmlElement | null {
  const orderedBreaks: number[] = [];
  for (const value of breaks) {
    if (orderedBreaks.indexOf(value) < 0) {
      orderedBreaks.push(value);
    }
  }

  orderedBreaks.sort((left, right) => left - right);
  if (orderedBreaks.length === 0) {
    return null;
  }

  const root = XmlElement.createWithNamespace(elementName, MAIN_NS)
    .attr("count", orderedBreaks.length)
    .attr("manualBreakCount", orderedBreaks.length);
  for (const value of orderedBreaks) {
    root.child(XmlElement.create("brk").attr("id", value).attr("max", maxIndex).attr("man", 1));
  }

  return root;
}

export function loadWorksheetPageSetup(
  worksheetModel: WorksheetModel,
  worksheetRoot: XmlNode,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  const pageSetup = worksheetModel.PageSetup;
  const sheetPr = firstElement(worksheetRoot, "sheetPr");
  const setupProperties = sheetPr === null ? null : firstElement(sheetPr, "pageSetUpPr");
  pageSetup.FitToPage = parseBoolAttribute(
    setupProperties === null ? null : attributeOrNull(setupProperties, "fitToPage"),
  );
  loadPageMargins(pageSetup, firstElement(worksheetRoot, "pageMargins"), diagnostics, options, sheetName);
  loadPageSetupCore(pageSetup, firstElement(worksheetRoot, "pageSetup"), diagnostics, options, sheetName);
  loadPrintOptions(pageSetup, firstElement(worksheetRoot, "printOptions"));
  loadHeaderFooter(pageSetup, firstElement(worksheetRoot, "headerFooter"));
  loadBreaks(pageSetup.HorizontalPageBreaks, firstElement(worksheetRoot, "rowBreaks"), diagnostics, options, sheetName, "row");
  loadBreaks(pageSetup.VerticalPageBreaks, firstElement(worksheetRoot, "colBreaks"), diagnostics, options, sheetName, "column");
}

/** Alias for loadWorksheetPageSetup, matching the C# loader entry point name. */
export function loadPageSetup(
  worksheetModel: WorksheetModel,
  worksheetRoot: XmlNode,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  loadWorksheetPageSetup(worksheetModel, worksheetRoot, diagnostics, options, sheetName);
}

export function loadPageMargins(
  pageSetup: PageSetupModel,
  marginsElement: XmlNode | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  if (marginsElement === null) {
    return;
  }

  pageSetup.Margins.Left = parseMarginAttribute(attributeOrNull(marginsElement, "left"), pageSetup.Margins.Left, diagnostics, options, sheetName, "left");
  pageSetup.Margins.Right = parseMarginAttribute(attributeOrNull(marginsElement, "right"), pageSetup.Margins.Right, diagnostics, options, sheetName, "right");
  pageSetup.Margins.Top = parseMarginAttribute(attributeOrNull(marginsElement, "top"), pageSetup.Margins.Top, diagnostics, options, sheetName, "top");
  pageSetup.Margins.Bottom = parseMarginAttribute(attributeOrNull(marginsElement, "bottom"), pageSetup.Margins.Bottom, diagnostics, options, sheetName, "bottom");
  pageSetup.Margins.Header = parseMarginAttribute(attributeOrNull(marginsElement, "header"), pageSetup.Margins.Header, diagnostics, options, sheetName, "header");
  pageSetup.Margins.Footer = parseMarginAttribute(attributeOrNull(marginsElement, "footer"), pageSetup.Margins.Footer, diagnostics, options, sheetName, "footer");
}

export function loadPageSetupCore(
  pageSetup: PageSetupModel,
  pageSetupElement: XmlNode | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  if (pageSetupElement === null) {
    return;
  }

  pageSetup.PaperSize = parsePositiveIntAttribute(attributeOrNull(pageSetupElement, "paperSize"), diagnostics, options, sheetName, "paperSize") ?? 0;
  pageSetup.FirstPageNumber = parsePositiveIntAttribute(attributeOrNull(pageSetupElement, "firstPageNumber"), diagnostics, options, sheetName, "firstPageNumber");
  pageSetup.Scale = parseBoundedIntAttribute(attributeOrNull(pageSetupElement, "scale"), 10, 400, diagnostics, options, sheetName, "scale");
  pageSetup.FitToWidth = parseNonNegativeIntAttribute(attributeOrNull(pageSetupElement, "fitToWidth"), diagnostics, options, sheetName, "fitToWidth");
  pageSetup.FitToHeight = parseNonNegativeIntAttribute(attributeOrNull(pageSetupElement, "fitToHeight"), diagnostics, options, sheetName, "fitToHeight");
  pageSetup.Orientation = parseOrientation(attributeOrNull(pageSetupElement, "orientation"), diagnostics, options, sheetName);
  pageSetup.PageOrderOverThenDown = attributeOrNull(pageSetupElement, "pageOrder") === "overThenDown";
}

export function loadPrintOptions(pageSetup: PageSetupModel, printOptionsElement: XmlNode | null): void {
  if (printOptionsElement === null) {
    return;
  }

  pageSetup.PrintOptions.Headings = parseBoolAttribute(attributeOrNull(printOptionsElement, "headings"));
  pageSetup.PrintOptions.GridLines = parseBoolAttribute(attributeOrNull(printOptionsElement, "gridLines"));
  pageSetup.PrintOptions.HorizontalCentered = parseBoolAttribute(attributeOrNull(printOptionsElement, "horizontalCentered"));
  pageSetup.PrintOptions.VerticalCentered = parseBoolAttribute(attributeOrNull(printOptionsElement, "verticalCentered"));
}

export function loadHeaderFooter(pageSetup: PageSetupModel, headerFooterElement: XmlNode | null): void {
  if (headerFooterElement === null) {
    return;
  }

  const oddHeader = firstElement(headerFooterElement, "oddHeader");
  const headerSections = parseHeaderFooterText(oddHeader === null ? null : textOf(oddHeader));
  pageSetup.HeaderFooter.LeftHeader = headerSections.Left;
  pageSetup.HeaderFooter.CenterHeader = headerSections.Center;
  pageSetup.HeaderFooter.RightHeader = headerSections.Right;

  const oddFooter = firstElement(headerFooterElement, "oddFooter");
  const footerSections = parseHeaderFooterText(oddFooter === null ? null : textOf(oddFooter));
  pageSetup.HeaderFooter.LeftFooter = footerSections.Left;
  pageSetup.HeaderFooter.CenterFooter = footerSections.Center;
  pageSetup.HeaderFooter.RightFooter = footerSections.Right;
}

export function loadBreaks(
  target: number[],
  breaksElement: XmlNode | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  axis: string,
): void {
  target.length = 0;
  if (breaksElement === null) {
    return;
  }

  for (const breakElement of elementChildren(breaksElement, "brk")) {
    const id = parseNonNegativeIntAttribute(attributeOrNull(breakElement, "id"), diagnostics, options, sheetName, `${axis}Break`);
    if (id === null) {
      continue;
    }

    if (target.indexOf(id) < 0) {
      target.push(id);
    }
  }
}

export function loadRowBreaks(
  pageSetup: PageSetupModel,
  breaksElement: XmlNode | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  loadBreaks(pageSetup.HorizontalPageBreaks, breaksElement, diagnostics, options, sheetName, "row");
}

export function loadColumnBreaks(
  pageSetup: PageSetupModel,
  breaksElement: XmlNode | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  loadBreaks(pageSetup.VerticalPageBreaks, breaksElement, diagnostics, options, sheetName, "column");
}

export function normalizeLoadedPrintArea(value: string): string {
  const segments: string[] = [];
  for (const segment of splitReferenceList(value)) {
    segments.push(normalizeAreaReference(removeWorksheetQualifier(segment)));
  }

  return segments.join(",");
}

function composeHeaderFooterText(left: string | null, center: string | null, right: string | null): string {
  let text = "";
  if (isNotEmpty(left)) {
    text += `&L${left}`;
  }

  if (isNotEmpty(center)) {
    text += `&C${center}`;
  }

  if (isNotEmpty(right)) {
    text += `&R${right}`;
  }

  return text;
}

interface HeaderFooterSections {
  Left: string | null;
  Center: string | null;
  Right: string | null;
}

function parseHeaderFooterText(value: string | null): HeaderFooterSections {
  const sections: HeaderFooterSections = { Left: null, Center: null, Right: null };
  if (value === null || value.length === 0) {
    return sections;
  }

  let currentSection = "C";
  let builder = "";
  const text = value;
  for (let index = 0; index < text.length; index++) {
    if (text[index] === "&" && index + 1 < text.length) {
      const marker = text[index + 1].toUpperCase();
      if (marker === "L" || marker === "C" || marker === "R") {
        assignHeaderFooterSection(currentSection, builder, sections);
        builder = "";
        currentSection = marker;
        index++;
        continue;
      }
    }

    builder += text[index];
  }

  assignHeaderFooterSection(currentSection, builder, sections);
  return sections;
}

function assignHeaderFooterSection(section: string, value: string, sections: HeaderFooterSections): void {
  const text = value.length === 0 ? null : value;
  if (section === "L") {
    sections.Left = text;
  } else if (section === "R") {
    sections.Right = text;
  } else {
    sections.Center = text;
  }
}

function marginsEqual(left: PageMarginsModel, right: PageMarginsModel): boolean {
  return (
    left.Left === right.Left &&
    left.Right === right.Right &&
    left.Top === right.Top &&
    left.Bottom === right.Bottom &&
    left.Header === right.Header &&
    left.Footer === right.Footer
  );
}

function buildPrintTitlesDefinedNameText(pageSetup: PageSetupModel, sheetName: string): string {
  const segments: string[] = [];
  const normalizedRows = normalizeTitleRows(pageSetup.PrintTitleRows, sheetName);
  if (normalizedRows.length > 0) {
    segments.push(normalizedRows);
  }

  const normalizedColumns = normalizeTitleColumns(pageSetup.PrintTitleColumns, sheetName);
  if (normalizedColumns.length > 0) {
    segments.push(normalizedColumns);
  }

  return segments.join(",");
}

function normalizePrintAreaList(printArea: string | null, sheetName: string): string {
  if (isNullOrWhiteSpace(printArea)) {
    return "";
  }

  const segments: string[] = [];
  for (const segment of splitReferenceList(printArea as string)) {
    segments.push(qualifyReference(sheetName, normalizeAreaReference(removeWorksheetQualifier(segment))));
  }

  return segments.join(",");
}

function parseLoadedPrintTitles(value: string, state: WorksheetDefinedNamesState): void {
  for (const segment of splitReferenceList(value)) {
    const unqualified = removeWorksheetQualifier(segment);
    if (looksLikeRowRange(unqualified)) {
      state.PrintTitleRows = normalizeRowReference(unqualified);
    } else if (looksLikeColumnRange(unqualified)) {
      state.PrintTitleColumns = normalizeColumnReference(unqualified);
    } else {
      throw new CellsException("Print title reference is invalid.");
    }
  }
}

function normalizeTitleRows(value: string | null, sheetName: string): string {
  if (isNullOrWhiteSpace(value)) {
    return "";
  }

  return qualifyReference(sheetName, normalizeRowReference(removeWorksheetQualifier(value as string)));
}

function normalizeTitleColumns(value: string | null, sheetName: string): string {
  if (isNullOrWhiteSpace(value)) {
    return "";
  }

  return qualifyReference(sheetName, normalizeColumnReference(removeWorksheetQualifier(value as string)));
}

function splitReferenceList(value: string): string[] {
  const parts: string[] = [];
  let builder = "";
  let inQuotes = false;

  for (let index = 0; index < value.length; index++) {
    const character = value[index];
    if (character === "'") {
      builder += character;
      if (inQuotes && index + 1 < value.length && value[index + 1] === "'") {
        builder += value[index + 1];
        index++;
        continue;
      }

      inQuotes = !inQuotes;
      continue;
    }

    if (character === "," && !inQuotes) {
      const part = builder.trim();
      if (part.length > 0) {
        parts.push(part);
      }

      builder = "";
      continue;
    }

    builder += character;
  }

  const lastPart = builder.trim();
  if (lastPart.length > 0) {
    parts.push(lastPart);
  }

  return parts;
}

function removeWorksheetQualifier(value: string): string {
  const trimmed = value.trim();
  if (trimmed.length === 0) {
    throw new CellsException("Worksheet reference is invalid.");
  }

  if (trimmed[0] === "'") {
    let index = 1;
    while (index < trimmed.length) {
      if (trimmed[index] === "'") {
        if (index + 1 < trimmed.length && trimmed[index + 1] === "'") {
          index += 2;
          continue;
        }

        break;
      }

      index++;
    }

    if (index + 1 < trimmed.length && trimmed[index + 1] === "!") {
      return trimmed.substring(index + 2);
    }
  }

  const exclamation = trimmed.indexOf("!");
  return exclamation >= 0 ? trimmed.substring(exclamation + 1) : trimmed;
}

function qualifyReference(sheetName: string, reference: string): string {
  return `${quoteWorksheetName(sheetName)}!${reference}`;
}

function quoteWorksheetName(sheetName: string): string {
  return `'${sheetName.replace(/'/g, "''")}'`;
}

function normalizeAreaReference(value: string): string {
  const parts = value.split(":");
  if (parts.length === 1) {
    return toAbsoluteCellReference(parseAbsoluteCellAddress(parts[0]));
  }

  if (parts.length !== 2) {
    throw new CellsException("PrintArea must be a cell or range reference.");
  }

  const first = parseAbsoluteCellAddress(parts[0]);
  const last = parseAbsoluteCellAddress(parts[1]);
  if (last.RowIndex < first.RowIndex || last.ColumnIndex < first.ColumnIndex) {
    throw new CellsException("PrintArea range must be ordered from top-left to bottom-right.");
  }

  return `${toAbsoluteCellReference(first)}:${toAbsoluteCellReference(last)}`;
}

function normalizeRowReference(value: string): string {
  const normalized = value.replace(/\$/g, "").trim();
  const parts = normalized.split(":");
  const first = parts.length === 2 ? parseIntValue(parts[0]) : null;
  const last = parts.length === 2 ? parseIntValue(parts[1]) : null;
  if (parts.length !== 2 || first === null || last === null || first <= 0 || last < first) {
    throw new CellsException("PrintTitleRows must be a row span like '1:2'.");
  }

  return `$${first}:$${last}`;
}

function normalizeColumnReference(value: string): string {
  const normalized = value.replace(/\$/g, "").trim().toUpperCase();
  const parts = normalized.split(":");
  if (parts.length !== 2 || !isColumnName(parts[0]) || !isColumnName(parts[1])) {
    throw new CellsException("PrintTitleColumns must be a column span like 'A:B'.");
  }

  return `$${parts[0]}:$${parts[1]}`;
}

function looksLikeRowRange(value: string): boolean {
  for (const part of value.replace(/\$/g, "").split(":")) {
    if (parseIntValue(part) === null) {
      return false;
    }
  }

  return true;
}

function looksLikeColumnRange(value: string): boolean {
  for (const part of value.replace(/\$/g, "").split(":")) {
    if (!isColumnName(part)) {
      return false;
    }
  }

  return true;
}

function isColumnName(value: string): boolean {
  if (isNullOrWhiteSpace(value)) {
    return false;
  }

  for (let index = 0; index < value.length; index++) {
    if (!isLetter(value.charCodeAt(index))) {
      return false;
    }
  }

  return true;
}

function isLetter(code: number): boolean {
  return (code >= 65 && code <= 90) || (code >= 97 && code <= 122);
}

function parseAbsoluteCellAddress(value: string): CellAddress {
  const normalized = value.replace(/\$/g, "").trim();
  try {
    return CellAddress.Parse(normalized);
  } catch (exception) {
    throw new CellsException(`Cell reference '${value}' is invalid.`, exception);
  }
}

function toAbsoluteCellReference(address: CellAddress): string {
  const reference = address.ToString();
  let splitIndex = 0;
  while (splitIndex < reference.length && isLetter(reference.charCodeAt(splitIndex))) {
    splitIndex++;
  }

  return `$${reference.substring(0, splitIndex)}$${reference.substring(splitIndex)}`;
}

function parseMarginAttribute(
  value: string | null,
  fallback: number,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  attributeName: string,
): number {
  if (value === null) {
    return fallback;
  }

  const parsed = parseDoubleValue(value);
  if (parsed !== null && parsed >= 0) {
    return parsed;
  }

  if (options.StrictMode) {
    throw new InvalidFileFormatException(`The page margin attribute '${attributeName}' is invalid.`);
  }

  addIssue(
    diagnostics,
    options,
    makeIssue(
      "PG-L001",
      DiagnosticSeverity.Warning,
      `Page margin attribute '${attributeName}' is invalid and the default value was used.`,
      false,
      true,
      sheetName,
    ),
  );
  return fallback;
}

function parsePositiveIntAttribute(
  value: string | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  attributeName: string,
): number | null {
  return parseConstrainedIntAttribute(value, 1, Number.POSITIVE_INFINITY, diagnostics, options, sheetName, attributeName);
}

function parseNonNegativeIntAttribute(
  value: string | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  attributeName: string,
): number | null {
  return parseConstrainedIntAttribute(value, 0, Number.POSITIVE_INFINITY, diagnostics, options, sheetName, attributeName);
}

function parseBoundedIntAttribute(
  value: string | null,
  minimum: number,
  maximum: number,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  attributeName: string,
): number | null {
  return parseConstrainedIntAttribute(value, minimum, maximum, diagnostics, options, sheetName, attributeName);
}

function parseConstrainedIntAttribute(
  value: string | null,
  minimum: number,
  maximum: number,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  attributeName: string,
): number | null {
  if (value === null) {
    return null;
  }

  const parsed = parseIntValue(value);
  if (parsed !== null && parsed >= minimum && parsed <= maximum) {
    return parsed;
  }

  if (options.StrictMode) {
    throw new InvalidFileFormatException(`The page setup attribute '${attributeName}' is invalid.`);
  }

  addIssue(
    diagnostics,
    options,
    makeIssue(
      "PG-L002",
      DiagnosticSeverity.Warning,
      `Page setup attribute '${attributeName}' is invalid and was ignored.`,
      false,
      true,
      sheetName,
    ),
  );
  return null;
}

function parseOrientation(
  value: string | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): PageOrientation {
  if (isNullOrWhiteSpace(value)) {
    return PageOrientation.Default;
  }

  if ((value as string).toLowerCase() === "portrait") {
    return PageOrientation.Portrait;
  }

  if ((value as string).toLowerCase() === "landscape") {
    return PageOrientation.Landscape;
  }

  if (options.StrictMode) {
    throw new InvalidFileFormatException("The page setup orientation is invalid.");
  }

  addIssue(
    diagnostics,
    options,
    makeIssue(
      "PG-L003",
      DiagnosticSeverity.Warning,
      "Page setup orientation is invalid and the default orientation was used.",
      false,
      true,
      sheetName,
    ),
  );
  return PageOrientation.Default;
}

function parseBoolAttribute(value: string | null): boolean {
  if (value === null) {
    return false;
  }

  return value === "1" || value.toLowerCase() === "true";
}

function parseIntValue(value: string): number | null {
  return parseIntAttribute(value);
}

function parseIntAttribute(value: string | null): number | null {
  if (value === null) {
    return null;
  }

  const trimmed = value.trim();
  if (!/^[+-]?\d+$/.test(trimmed)) {
    return null;
  }

  const parsed = Number(trimmed);
  if (!Number.isInteger(parsed) || parsed < -2147483648 || parsed > 2147483647) {
    return null;
  }

  return parsed;
}

function parseDoubleValue(value: string): number | null {
  const trimmed = value.trim();
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(trimmed)) {
    return null;
  }

  const parsed = Number(trimmed);
  return Number.isNaN(parsed) ? null : parsed;
}

function attributeOrNull(node: XmlNode, name: string): string | null {
  const attributes = (node as unknown as { attributes?: ArrayLike<{ nodeName: string; value: string }> }).attributes;
  if (attributes === undefined) {
    return null;
  }

  for (let index = 0; index < attributes.length; index++) {
    const attributeNode = attributes[index];
    if (localNameOf(attributeNode.nodeName) === name) {
      return attributeNode.value;
    }
  }

  return null;
}

function localNameOf(nodeName: string): string {
  const separator = nodeName.indexOf(":");
  return separator >= 0 ? nodeName.substring(separator + 1) : nodeName;
}

function formatNumber(value: number): string {
  return String(value);
}

function isNotEmpty(value: string | null): boolean {
  return value !== null && value.length > 0;
}

function isNullOrWhiteSpace(value: string | null): boolean {
  return value === null || value.trim().length === 0;
}

function equalsIgnoreCase(left: string, right: string): boolean {
  return left.toLowerCase() === right.toLowerCase();
}

function makeIssue(
  code: string,
  severity: DiagnosticSeverity,
  message: string,
  repairApplied: boolean,
  dataLossRisk: boolean,
  sheetName: string | null,
): LoadIssue {
  const issue = new LoadIssue(code, severity, message, repairApplied, dataLossRisk);
  issue.SheetName = sheetName;
  return issue;
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
