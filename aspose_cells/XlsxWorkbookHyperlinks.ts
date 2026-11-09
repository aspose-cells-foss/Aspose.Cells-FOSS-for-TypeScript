import { CellAddress } from "./Core/CellAddress";
import { HyperlinkModel } from "./Core/HyperlinkModel";
import { WorksheetModel } from "./Core/WorksheetModel";
import { DiagnosticSeverity } from "./enums";
import { InvalidFileFormatException } from "./InvalidFileFormatException";
import { LoadDiagnostics } from "./LoadDiagnostics";
import { LoadIssue } from "./LoadIssue";
import { LoadOptions } from "./LoadOptions";
import { WarningInfo } from "./WarningInfo";
import { XmlElement } from "./Xml/XmlWriter";
import type { ZipReader } from "./Packaging/Archive";
import { attribute, documentRoot, elementChildren, firstElement, parseXml, XmlNode } from "./Xml/XmlReader";
import { PACKAGE_RELATIONSHIP_NS, tryParseMergeReference } from "./XlsxWorkbookSerializerCommon";

/** Relationship type used for external hyperlinks. */
export const HYPERLINK_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

/** Builds the worksheet "hyperlinks" element, mirroring XlsxWorkbookHyperlinks.BuildHyperlinksElement. */
export function buildHyperlinksElement(worksheet: WorksheetModel): XmlElement | null {
  const orderedHyperlinks = getOrderedHyperlinks(worksheet.Hyperlinks);
  if (orderedHyperlinks.length === 0) {
    return null;
  }

  let relationshipId = 1;
  const hyperlinksElement = XmlElement.create("hyperlinks");
  for (const hyperlink of orderedHyperlinks) {
    if (!shouldPersistHyperlink(hyperlink)) {
      continue;
    }

    const element = XmlElement.create("hyperlink").attr("ref", toRangeReference(hyperlink));
    if (!isEmptyString(hyperlink.Address)) {
      element.attr("r:id", "rId" + relationshipId.toString());
      relationshipId++;
    }

    if (!isEmptyString(hyperlink.SubAddress)) {
      element.attr("location", hyperlink.SubAddress);
    }

    if (!isEmptyString(hyperlink.ScreenTip)) {
      element.attr("tooltip", hyperlink.ScreenTip);
    }

    if (!isEmptyString(hyperlink.TextToDisplay)) {
      element.attr("display", hyperlink.TextToDisplay);
    }

    hyperlinksElement.child(element);
  }

  return hyperlinksElement.childCount === 0 ? null : hyperlinksElement;
}

/** Builds the worksheet relationship part for external hyperlinks. */
export function buildWorksheetHyperlinkRelationships(worksheet: WorksheetModel): XmlElement | null {
  const orderedHyperlinks = getOrderedHyperlinks(worksheet.Hyperlinks);
  const relationships = XmlElement.createWithNamespace("Relationships", PACKAGE_RELATIONSHIP_NS);
  let relationshipId = 1;

  for (const hyperlink of orderedHyperlinks) {
    if (isEmptyString(hyperlink.Address)) {
      continue;
    }

    relationships.child(
      XmlElement.create("Relationship")
        .attr("Id", "rId" + relationshipId.toString())
        .attr("Type", HYPERLINK_RELATIONSHIP_TYPE)
        .attr("Target", hyperlink.Address)
        .attr("TargetMode", "External"),
    );
    relationshipId++;
  }

  return relationships.childCount === 0 ? null : relationships;
}

/** Counts the external hyperlinks that receive a relationship id, mirroring XlsxWorkbookSerializer.CountExternalHyperlinks. */
export function countExternalHyperlinks(worksheet: WorksheetModel): number {
  let count = 0;
  const ordered = getOrderedHyperlinks(worksheet.Hyperlinks);
  for (const hyperlink of ordered) {
    if (!isEmptyString(hyperlink.Address)) {
      count++;
    }
  }

  return count;
}

/** Loads worksheet relationship targets used to resolve hyperlink addresses. */
export function loadWorksheetHyperlinkTargets(archive: ZipReader, worksheetUri: string): Map<string, string> {
  const relationshipsEntry = archive.tryReadText(getWorksheetRelationshipsUri(worksheetUri));
  if (relationshipsEntry === null) {
    return new Map<string, string>();
  }

  const document = parseXml(relationshipsEntry);
  const root = documentRoot(document);
  const targets = new Map<string, string>();
  if (root === null) {
    return targets;
  }

  for (const relationship of elementChildren(root, "Relationship")) {
    const id = attribute(relationship, "Id");
    const type = attribute(relationship, "Type");
    const target = attribute(relationship, "Target");
    if (isEmptyString(id) || isEmptyString(target) || !equalsIgnoreCase(type, HYPERLINK_RELATIONSHIP_TYPE)) {
      continue;
    }

    targets.set(id as string, target as string);
  }

  return targets;
}

/** Resolves a relationship id case-insensitively, mirroring the C# StringComparer.OrdinalIgnoreCase dictionary. */
function resolveHyperlinkTarget(targets: ReadonlyMap<string, string>, relationshipId: string): string | undefined {
  const direct = targets.get(relationshipId);
  if (direct !== undefined) {
    return direct;
  }

  const lower = relationshipId.toLowerCase();
  for (const [key, value] of targets) {
    if (key.toLowerCase() === lower) {
      return value;
    }
  }

  return undefined;
}

/** Loads hyperlinks from a worksheet root element. */
export function loadHyperlinks(
  worksheetModel: WorksheetModel,
  worksheetRoot: XmlNode,
  hyperlinkTargets: ReadonlyMap<string, string>,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  worksheetModel.Hyperlinks.length = 0;

  const container = firstElement(worksheetRoot, "hyperlinks");
  if (container === null) {
    return;
  }

  for (const hyperlinkElement of elementChildren(container, "hyperlink")) {
    const reference = attribute(hyperlinkElement, "ref");
    const candidate = tryParseHyperlinkReference(reference);
    if (candidate === null) {
      if (options.StrictMode) {
        throw new InvalidFileFormatException("The hyperlink reference '" + (reference ?? "") + "' is invalid.");
      }

      addIssue(
        diagnostics,
        options,
        issueWithLocation(
          "HL-L001",
          DiagnosticSeverity.LossyRecoverable,
          "A hyperlink with an invalid reference was dropped during load.",
          sheetName,
          reference,
        ),
      );
      continue;
    }

    if (containsOverlappingHyperlink(worksheetModel.Hyperlinks, candidate)) {
      if (options.StrictMode) {
        throw new InvalidFileFormatException("The hyperlink reference '" + (reference ?? "") + "' overlaps an existing hyperlink.");
      }

      addIssue(
        diagnostics,
        options,
        issueWithLocation(
          "HL-L003",
          DiagnosticSeverity.LossyRecoverable,
          "Overlapping hyperlink ranges were normalized during load.",
          sheetName,
          reference,
        ),
      );
      continue;
    }

    candidate.SubAddress = normalizeAttributeValue(attribute(hyperlinkElement, "location"));
    candidate.ScreenTip = normalizeAttributeValue(attribute(hyperlinkElement, "tooltip"));
    candidate.TextToDisplay = normalizeAttributeValue(attribute(hyperlinkElement, "display"));

    const relationshipId = attribute(hyperlinkElement, "id");
    if (!isEmptyString(relationshipId)) {
      const address = resolveHyperlinkTarget(hyperlinkTargets, relationshipId as string);
      if (address !== undefined) {
        candidate.Address = normalizeAttributeValue(address);
      } else if (options.StrictMode) {
        throw new InvalidFileFormatException("The hyperlink relationship '" + relationshipId + "' could not be resolved.");
      } else {
        addIssue(
          diagnostics,
          options,
          issueWithLocation(
            "HL-L002",
            DiagnosticSeverity.LossyRecoverable,
            "A hyperlink relationship was missing and the external address was cleared.",
            sheetName,
            reference,
          ),
        );
      }
    }

    if (isEmptyString(candidate.Address) && isEmptyString(candidate.SubAddress)) {
      if (options.StrictMode) {
        throw new InvalidFileFormatException(
          "The hyperlink reference '" + (reference ?? "") + "' does not define an address or sub-address.",
        );
      }

      addIssue(
        diagnostics,
        options,
        issueWithLocation(
          "HL-L004",
          DiagnosticSeverity.LossyRecoverable,
          "A hyperlink without an address or sub-address was dropped during load.",
          sheetName,
          reference,
        ),
      );
      continue;
    }

    worksheetModel.Hyperlinks.push(candidate);
  }
}

/** Returns the hyperlinks ordered by range, mirroring XlsxWorkbookHyperlinks.GetOrderedHyperlinks. */
export function getOrderedHyperlinks(hyperlinks: readonly HyperlinkModel[]): HyperlinkModel[] {
  return hyperlinks.slice().sort(compareHyperlinks);
}

function compareHyperlinks(left: HyperlinkModel, right: HyperlinkModel): number {
  if (left.FirstRow !== right.FirstRow) {
    return left.FirstRow - right.FirstRow;
  }

  if (left.FirstColumn !== right.FirstColumn) {
    return left.FirstColumn - right.FirstColumn;
  }

  if (left.TotalRows !== right.TotalRows) {
    return left.TotalRows - right.TotalRows;
  }

  return left.TotalColumns - right.TotalColumns;
}

function shouldPersistHyperlink(hyperlink: HyperlinkModel): boolean {
  return !isEmptyString(hyperlink.Address) || !isEmptyString(hyperlink.SubAddress);
}

function toRangeReference(hyperlink: HyperlinkModel): string {
  const first = new CellAddress(hyperlink.FirstRow, hyperlink.FirstColumn).ToString();
  if (hyperlink.TotalRows === 1 && hyperlink.TotalColumns === 1) {
    return first;
  }

  const last = new CellAddress(
    hyperlink.FirstRow + hyperlink.TotalRows - 1,
    hyperlink.FirstColumn + hyperlink.TotalColumns - 1,
  ).ToString();
  return first + ":" + last;
}

function getWorksheetRelationshipsUri(worksheetUri: string): string {
  const normalizedUri = worksheetUri.replace(/^\/+/, "");
  const slashIndex = normalizedUri.lastIndexOf("/");
  const directory = slashIndex >= 0 ? normalizedUri.substring(0, slashIndex + 1) : "";
  const fileName = slashIndex >= 0 ? normalizedUri.substring(slashIndex + 1) : normalizedUri;
  return "/" + directory + "_rels/" + fileName + ".rels";
}

function tryParseHyperlinkReference(reference: string | null): HyperlinkModel | null {
  const region = tryParseMergeReference(reference ?? "");
  if (region === null) {
    return null;
  }

  const hyperlink = new HyperlinkModel();
  hyperlink.FirstRow = region.FirstRow;
  hyperlink.FirstColumn = region.FirstColumn;
  hyperlink.TotalRows = region.TotalRows;
  hyperlink.TotalColumns = region.TotalColumns;
  return hyperlink;
}

function containsOverlappingHyperlink(hyperlinks: HyperlinkModel[], candidate: HyperlinkModel): boolean {
  for (const hyperlink of hyperlinks) {
    if (hyperlinksOverlap(hyperlink, candidate)) {
      return true;
    }
  }

  return false;
}

function hyperlinksOverlap(left: HyperlinkModel, right: HyperlinkModel): boolean {
  const leftLastRow = left.FirstRow + left.TotalRows - 1;
  const leftLastColumn = left.FirstColumn + left.TotalColumns - 1;
  const rightLastRow = right.FirstRow + right.TotalRows - 1;
  const rightLastColumn = right.FirstColumn + right.TotalColumns - 1;

  return (
    left.FirstRow <= rightLastRow &&
    right.FirstRow <= leftLastRow &&
    left.FirstColumn <= rightLastColumn &&
    right.FirstColumn <= leftLastColumn
  );
}

function normalizeAttributeValue(value: string | null): string | null {
  return value === null || value.length === 0 ? null : value;
}

function equalsIgnoreCase(left: string | null, right: string): boolean {
  return left !== null && left.toLowerCase() === right.toLowerCase();
}

function isEmptyString(value: string | null | undefined): boolean {
  return value === null || value === undefined || value.length === 0;
}

function issueWithLocation(
  code: string,
  severity: DiagnosticSeverity,
  message: string,
  sheetName: string,
  cellRef: string | null,
): LoadIssue {
  const issue = new LoadIssue(code, severity, message, true, true);
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
