import { CellAddress } from "./Core/CellAddress";
import { HyperlinkModel } from "./Core/HyperlinkModel";
import { ListColumnModel } from "./Core/ListColumnModel";
import { ListObjectModel } from "./Core/ListObjectModel";
import { WorkbookModel } from "./Core/WorkbookModel";
import { WorksheetModel } from "./Core/WorksheetModel";
import { DiagnosticSeverity } from "./enums";
import { InvalidFileFormatException } from "./InvalidFileFormatException";
import { LoadDiagnostics } from "./LoadDiagnostics";
import { LoadIssue } from "./LoadIssue";
import { LoadOptions } from "./LoadOptions";
import { ZipReader } from "./Packaging/Archive";
import {
  MAIN_NS,
  PACKAGE_RELATIONSHIP_NS,
  RELATIONSHIP_NS,
  tryParseCellReference,
} from "./XlsxWorkbookSerializerCommon";
import { COMMENTS_RELATIONSHIP_TYPE, VML_DRAWING_RELATIONSHIP_TYPE } from "./XlsxWorkbookComments";
import {
  attribute,
  attributeBool,
  attributeInt,
  documentRoot,
  elementChildren,
  firstElement,
  parseXml,
  XmlNode,
} from "./Xml/XmlReader";
import { XmlElement } from "./Xml/XmlWriter";

export const TABLE_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
export const TABLE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";

const HYPERLINK_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const DRAWING_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

/** Builds the table part document for a single list object. */
export function buildTableXml(model: ListObjectModel, tableId: number): XmlElement {
  const startAddress = new CellAddress(model.StartRow, model.StartColumn).ToString();
  const endAddress = new CellAddress(model.EndRow, model.EndColumn).ToString();
  const tableRef = startAddress + ":" + endAddress;

  const tableElement = XmlElement.createWithNamespace("table", MAIN_NS)
    .attr("xmlns:r", RELATIONSHIP_NS)
    .attr("id", tableId)
    .attr("name", model.Name)
    .attr("displayName", model.DisplayName)
    .attr("ref", tableRef);

  if (!model.ShowHeaderRow) {
    tableElement.attr("headerRowCount", "0");
  }

  if (model.ShowTotals) {
    tableElement.attr("totalsRowCount", "1");
  } else {
    tableElement.attr("totalsRowShown", "0");
  }

  if (model.Comment.length > 0) {
    tableElement.attr("comment", model.Comment);
  }

  if (model.HasAutoFilter && model.ShowHeaderRow) {
    tableElement.child(
      XmlElement.create("autoFilter").attr("ref", buildAutoFilterRef(model)),
    );
  }

  const tableColumns = XmlElement.create("tableColumns").attr("count", model.Columns.length);
  for (let index = 0; index < model.Columns.length; index++) {
    tableColumns.child(buildTableColumnElement(model.Columns[index]));
  }

  tableElement.child(tableColumns);

  const tableStyleInfo = buildTableStyleInfoElement(model);
  if (tableStyleInfo !== null) {
    tableElement.child(tableStyleInfo);
  }

  return tableElement;
}

/** Alias kept for the C# BuildTableDocument name. */
export function buildTableDocument(model: ListObjectModel, tableId: number): XmlElement {
  return buildTableXml(model, tableId);
}

function buildAutoFilterRef(model: ListObjectModel): string {
  const dataEndRow = model.ShowTotals ? model.EndRow - 1 : model.EndRow;
  const startAddress = new CellAddress(model.StartRow, model.StartColumn).ToString();
  const endAddress = new CellAddress(dataEndRow, model.EndColumn).ToString();
  return startAddress + ":" + endAddress;
}

function buildTableColumnElement(column: ListColumnModel): XmlElement {
  const element = XmlElement.create("tableColumn")
    .attr("id", column.Id)
    .attr("name", column.Name);

  const fn = column.TotalsRowFunction;
  if (fn.length > 0 && fn !== "none") {
    element.attr("totalsRowFunction", fn);
  }

  if (fn === "none" || fn.length === 0) {
    if (column.TotalsRowLabel.length > 0) {
      element.attr("totalsRowLabel", column.TotalsRowLabel);
    }
  }

  if (fn === "custom" && column.TotalsRowFormula.length > 0) {
    element.attr("totalsRowFormula", column.TotalsRowFormula);
  }

  return element;
}

function buildTableStyleInfoElement(model: ListObjectModel): XmlElement | null {
  if (model.TableStyleName.length === 0) {
    return null;
  }

  return XmlElement.create("tableStyleInfo")
    .attr("name", model.TableStyleName)
    .attr("showFirstColumn", model.ShowFirstColumn ? "1" : "0")
    .attr("showLastColumn", model.ShowLastColumn ? "1" : "0")
    .attr("showRowStripes", model.ShowRowStripes ? "1" : "0")
    .attr("showColumnStripes", model.ShowColumnStripes ? "1" : "0");
}

/** Builds the tableParts element attached to the worksheet root. */
export function buildTablePartsElement(
  worksheet: WorksheetModel,
  tableRIdStart: number,
): XmlElement | null {
  if (worksheet.ListObjects.length === 0) {
    return null;
  }

  const tablePartsElement = XmlElement.createWithNamespace("tableParts", MAIN_NS).attr(
    "count",
    worksheet.ListObjects.length,
  );

  for (let index = 0; index < worksheet.ListObjects.length; index++) {
    const rId = "rId" + (tableRIdStart + index);
    tablePartsElement.child(XmlElement.create("tablePart").attr("r:id", rId));
  }

  return tablePartsElement;
}

/** Builds the worksheet relationship part including tables, drawings and comments. */
export function buildWorksheetRelationshipsDocument(
  worksheet: WorksheetModel,
  tableFileOffset: number,
  drawingNumber: number,
  commentFileNumber: number,
): XmlElement | null {
  const relationships = XmlElement.createWithNamespace("Relationships", PACKAGE_RELATIONSHIP_NS);
  let rId = 1;

  const orderedHyperlinks = getOrderedHyperlinks(worksheet.Hyperlinks);
  for (let index = 0; index < orderedHyperlinks.length; index++) {
    const hyperlink = orderedHyperlinks[index];
    if (hyperlink.Address === null || hyperlink.Address.length === 0) {
      continue;
    }

    relationships.child(
      XmlElement.create("Relationship")
        .attr("Id", "rId" + rId)
        .attr("Type", HYPERLINK_RELATIONSHIP_TYPE)
        .attr("Target", hyperlink.Address)
        .attr("TargetMode", "External"),
    );
    rId++;
  }

  for (let t = 0; t < worksheet.ListObjects.length; t++) {
    const globalTableNumber = tableFileOffset + t + 1;
    relationships.child(
      XmlElement.create("Relationship")
        .attr("Id", "rId" + rId)
        .attr("Type", TABLE_RELATIONSHIP_TYPE)
        .attr("Target", "../tables/table" + globalTableNumber + ".xml"),
    );
    rId++;
  }

  if (drawingNumber > 0) {
    relationships.child(
      XmlElement.create("Relationship")
        .attr("Id", "rId" + rId)
        .attr("Type", DRAWING_RELATIONSHIP_TYPE)
        .attr("Target", "../drawings/drawing" + drawingNumber + ".xml"),
    );
    rId++;
  }

  if (commentFileNumber > 0) {
    relationships.child(
      XmlElement.create("Relationship")
        .attr("Id", "rId" + rId)
        .attr("Type", COMMENTS_RELATIONSHIP_TYPE)
        .attr("Target", "../comments" + commentFileNumber + ".xml"),
    );
    rId++;
    relationships.child(
      XmlElement.create("Relationship")
        .attr("Id", "rId" + rId)
        .attr("Type", VML_DRAWING_RELATIONSHIP_TYPE)
        .attr("Target", "../drawings/vmlDrawing" + commentFileNumber + ".vml"),
    );
    rId++;
  }

  if (relationships.childCount === 0) {
    return null;
  }

  return relationships;
}

function getOrderedHyperlinks(hyperlinks: readonly HyperlinkModel[]): HyperlinkModel[] {
  const ordered = hyperlinks.slice();
  ordered.sort((left, right) => {
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
  });
  return ordered;
}

/** Computes the running table part offset for each worksheet. */
export function computeTableFileOffsets(model: WorkbookModel): number[] {
  const offsets = new Array<number>(model.Worksheets.length);
  let running = 0;
  for (let index = 0; index < model.Worksheets.length; index++) {
    offsets[index] = running;
    running += model.Worksheets[index].ListObjects.length;
  }

  return offsets;
}

/** Counts all list objects across the workbook. */
export function computeTotalTableCount(model: WorkbookModel): number {
  let total = 0;
  for (let index = 0; index < model.Worksheets.length; index++) {
    total += model.Worksheets[index].ListObjects.length;
  }

  return total;
}

/** Loads all table parts referenced by the worksheet. */
export function loadTables(
  worksheetModel: WorksheetModel,
  worksheetRoot: XmlNode,
  archive: ZipReader,
  worksheetUri: string,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  worksheetModel.ListObjects.length = 0;

  const tablePartsElement = firstElement(worksheetRoot, "tableParts");
  if (tablePartsElement === null) {
    return;
  }

  const tableTargets = loadWorksheetTableTargets(archive, worksheetUri);

  for (const tablePartElement of elementChildren(tablePartsElement, "tablePart")) {
    const rId = attribute(tablePartElement, "id");
    if (rId === null || rId.length === 0) {
      continue;
    }

    const tableUri = tableTargets.get(rId.toLowerCase());
    if (tableUri === undefined) {
      addIssue(
        diagnostics,
        options,
        createIssue(
          "TBL-L002",
          DiagnosticSeverity.LossyRecoverable,
          "Table relationship '" + rId + "' could not be resolved and the table was dropped.",
          true,
          true,
          sheetName,
          null,
        ),
      );
      continue;
    }

    const tableText = archive.tryReadText(tableUri);
    if (tableText === null) {
      addIssue(
        diagnostics,
        options,
        createIssue(
          "TBL-R001",
          DiagnosticSeverity.Recoverable,
          "Table part '" + tableUri + "' was not found and the table was skipped.",
          true,
          false,
          sheetName,
          null,
        ),
      );
      continue;
    }

    const tableDocument = parseXml(tableText);
    const tableModel = loadTableModel(tableDocument, diagnostics, options, sheetName, worksheetModel);
    if (tableModel !== null) {
      worksheetModel.ListObjects.push(tableModel);
    }
  }
}

function loadWorksheetTableTargets(
  archive: ZipReader,
  worksheetUri: string,
): Map<string, string> {
  const targets = new Map<string, string>();
  const relsUri = getWorksheetRelsUri(worksheetUri);
  const relsText = archive.tryReadText(relsUri);
  if (relsText === null) {
    return targets;
  }

  const document = parseXml(relsText);
  const root = documentRoot(document);
  if (root === null) {
    return targets;
  }

  for (const relationship of elementChildren(root, "Relationship")) {
    const id = attribute(relationship, "Id");
    const type = attribute(relationship, "Type");
    const target = attribute(relationship, "Target");
    if (id === null || id.length === 0 || target === null || target.length === 0) {
      continue;
    }

    if (type === null || type.toLowerCase() !== TABLE_RELATIONSHIP_TYPE.toLowerCase()) {
      continue;
    }

    targets.set(id.toLowerCase(), resolvePartUri(worksheetUri, target));
  }

  return targets;
}

function loadTableModel(
  tableDocument: XmlNode,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  worksheetModel: WorksheetModel,
): ListObjectModel | null {
  const root = documentRoot(tableDocument);
  if (root === null) {
    return null;
  }

  const tableRef = attribute(root, "ref");
  if (tableRef === null || tableRef.length === 0) {
    addIssue(
      diagnostics,
      options,
      createIssue(
        "TBL-L001",
        DiagnosticSeverity.LossyRecoverable,
        "Table ref attribute is missing and the table was dropped.",
        true,
        true,
        sheetName,
        null,
      ),
    );
    return null;
  }

  const range = tryParseTableRef(tableRef);
  if (range === null) {
    if (options.StrictMode) {
      throw new InvalidFileFormatException("Table ref '" + tableRef + "' is invalid.");
    }

    addIssue(
      diagnostics,
      options,
      createIssue(
        "TBL-L001",
        DiagnosticSeverity.LossyRecoverable,
        "Table ref '" + tableRef + "' is invalid and the table was dropped.",
        true,
        true,
        sheetName,
        null,
      ),
    );
    return null;
  }

  const model = new ListObjectModel();
  model.StartRow = range.startRow;
  model.StartColumn = range.startColumn;
  model.EndRow = range.endRow;
  model.EndColumn = range.endColumn;

  model.DisplayName = attribute(root, "displayName") ?? "";
  const name = attribute(root, "name");
  model.Name = name !== null ? name : model.DisplayName;
  model.Comment = attribute(root, "comment") ?? "";

  if (model.DisplayName.length === 0) {
    model.DisplayName = model.Name;
  }

  const headerRowCount = attributeInt(root, "headerRowCount");
  model.ShowHeaderRow = headerRowCount === null || headerRowCount !== 0;

  const totalsRowCount = attributeInt(root, "totalsRowCount");
  model.ShowTotals = totalsRowCount !== null && totalsRowCount > 0;

  const autoFilterElement = firstElement(root, "autoFilter");
  model.HasAutoFilter = autoFilterElement !== null;

  loadTableColumns(model, root, diagnostics, options, sheetName);
  loadTableStyleInfo(model, root);

  return model;
}

function loadTableColumns(
  model: ListObjectModel,
  root: XmlNode,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  const tableColumnsElement = firstElement(root, "tableColumns");
  if (tableColumnsElement === null) {
    const expectedCount = model.EndColumn - model.StartColumn + 1;
    for (let columnIndex = 0; columnIndex < expectedCount; columnIndex++) {
      model.Columns.push(new ListColumnModel(columnIndex + 1, "Column" + (columnIndex + 1)));
    }

    addIssue(
      diagnostics,
      options,
      createIssue(
        "TBL-R004",
        DiagnosticSeverity.Recoverable,
        "Table columns element was missing; default column names were synthesized.",
        true,
        false,
        sheetName,
        null,
      ),
    );
    return;
  }

  for (const columnElement of elementChildren(tableColumnsElement, "tableColumn")) {
    const idAttribute = attributeInt(columnElement, "id");
    const name = attribute(columnElement, "name") ?? "";
    const columnId = idAttribute !== null && idAttribute > 0 ? idAttribute : model.Columns.length + 1;

    const columnModel = new ListColumnModel(columnId, name);
    columnModel.TotalsRowFunction = attribute(columnElement, "totalsRowFunction") ?? "none";
    columnModel.TotalsRowLabel = attribute(columnElement, "totalsRowLabel") ?? "";
    columnModel.TotalsRowFormula = attribute(columnElement, "totalsRowFormula") ?? "";

    model.Columns.push(columnModel);
  }
}

function loadTableStyleInfo(model: ListObjectModel, root: XmlNode): void {
  const styleInfoElement = firstElement(root, "tableStyleInfo");
  if (styleInfoElement === null) {
    return;
  }

  model.TableStyleName = attribute(styleInfoElement, "name") ?? "";
  model.ShowFirstColumn = attributeBool(styleInfoElement, "showFirstColumn");
  model.ShowLastColumn = attributeBool(styleInfoElement, "showLastColumn");
  model.ShowRowStripes = attributeBool(styleInfoElement, "showRowStripes");
  model.ShowColumnStripes = attributeBool(styleInfoElement, "showColumnStripes");
}

interface TableRange {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}

function tryParseTableRef(tableRef: string): TableRange | null {
  if (tableRef.trim().length === 0) {
    return null;
  }

  const parts = tableRef.split(":");
  if (parts.length !== 2) {
    return null;
  }

  const first = tryParseCellReference(parts[0]);
  const last = tryParseCellReference(parts[1]);
  if (first === null || last === null) {
    return null;
  }

  if (last.RowIndex < first.RowIndex || last.ColumnIndex < first.ColumnIndex) {
    return null;
  }

  return {
    startRow: first.RowIndex,
    startColumn: first.ColumnIndex,
    endRow: last.RowIndex,
    endColumn: last.ColumnIndex,
  };
}

function getWorksheetRelsUri(worksheetUri: string): string {
  const normalized = worksheetUri.replace(/^\/+/, "");
  const slashIndex = normalized.lastIndexOf("/");
  const directory = slashIndex >= 0 ? normalized.substring(0, slashIndex + 1) : "";
  const fileName = slashIndex >= 0 ? normalized.substring(slashIndex + 1) : normalized;
  return "/" + directory + "_rels/" + fileName + ".rels";
}

function resolvePartUri(sourceUri: string, target: string): string {
  if (target.startsWith("/")) {
    return target;
  }

  return new URL(target, "http://package" + sourceUri).pathname;
}

function createIssue(
  code: string,
  severity: DiagnosticSeverity,
  message: string,
  repairApplied: boolean,
  dataLossRisk: boolean,
  sheetName: string,
  cellRef: string | null,
): LoadIssue {
  const issue = new LoadIssue(code, severity, message, repairApplied, dataLossRisk);
  issue.SheetName = sheetName;
  issue.CellRef = cellRef;
  return issue;
}

function addIssue(
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  issue: LoadIssue,
): void {
  diagnostics.Add(issue);
  if (options.WarningCallback !== null && options.WarningCallback !== undefined) {
    const warning = {
      Code: issue.Code,
      Severity: issue.Severity,
      Message: issue.Message,
      DataLossRisk: issue.DataLossRisk,
      PartUri: issue.PartUri,
      SheetName: issue.SheetName,
      CellRef: issue.CellRef,
      RowIndex: issue.RowIndex,
    };
    options.WarningCallback.Warning(warning);
  }
}
