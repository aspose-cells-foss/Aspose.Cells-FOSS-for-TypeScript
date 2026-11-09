import { CommentModel } from "./Core/CommentModel";
import { WorksheetModel } from "./Core/WorksheetModel";
import { DiagnosticSeverity } from "./enums";
import { LoadDiagnostics } from "./LoadDiagnostics";
import { LoadIssue } from "./LoadIssue";
import { LoadOptions } from "./LoadOptions";
import { ZipReader } from "./Packaging/Archive";
import { tryParseCellReference } from "./XlsxWorkbookSerializerCommon";
import { COMMENTS_RELATIONSHIP_TYPE, VML_DRAWING_RELATIONSHIP_TYPE } from "./XlsxWorkbookComments";
import {
  attribute,
  documentRoot,
  elementChildren,
  firstElement,
  parseXml,
  textOf,
  XmlNode,
} from "./Xml/XmlReader";

/** Loads the comments and their VML drawing metadata for a worksheet. */
export function loadComments(
  worksheetModel: WorksheetModel,
  archive: ZipReader,
  worksheetUri: string,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  worksheetModel.Comments.length = 0;

  const relsUri = getWorksheetRelsUri(worksheetUri);
  const relsText = archive.tryReadText(relsUri);
  if (relsText === null) {
    return;
  }

  let relsDocument: XmlNode;
  try {
    relsDocument = parseXml(relsText);
  } catch {
    return;
  }

  const relsRoot = documentRoot(relsDocument);
  if (relsRoot === null) {
    return;
  }

  let commentsTarget: string | null = null;
  let vmlTarget: string | null = null;

  for (const relationship of elementChildren(relsRoot, "Relationship")) {
    const type = attribute(relationship, "Type");
    const target = attribute(relationship, "Target");
    if (type === null || type.length === 0 || target === null || target.length === 0) {
      continue;
    }

    if (type.toLowerCase() === COMMENTS_RELATIONSHIP_TYPE.toLowerCase()) {
      commentsTarget = resolvePartUri(worksheetUri, target);
    } else if (type.toLowerCase() === VML_DRAWING_RELATIONSHIP_TYPE.toLowerCase()) {
      vmlTarget = resolvePartUri(worksheetUri, target);
    }
  }

  if (commentsTarget === null) {
    if (vmlTarget !== null) {
      addIssue(
        diagnostics,
        options,
        createIssue(
          "COM-R003",
          DiagnosticSeverity.Recoverable,
          "A VML drawing part was found without a comments relationship; comments were skipped.",
          false,
          false,
          sheetName,
          null,
        ),
      );
    }

    return;
  }

  const commentsText = archive.tryReadText(commentsTarget);
  if (commentsText === null) {
    addIssue(
      diagnostics,
      options,
      createIssue(
        "COM-R002",
        DiagnosticSeverity.Recoverable,
        "Comments part '" + commentsTarget + "' was referenced but not found.",
        false,
        false,
        sheetName,
        null,
      ),
    );
    return;
  }

  loadCommentsXml(worksheetModel, commentsText, diagnostics, options, sheetName);

  if (vmlTarget === null) {
    addIssue(
      diagnostics,
      options,
      createIssue(
        "COM-R002",
        DiagnosticSeverity.Recoverable,
        "Comments were loaded but the VML drawing part is missing; default shape dimensions will be used.",
        true,
        false,
        sheetName,
        null,
      ),
    );
    return;
  }

  const vmlText = archive.tryReadText(vmlTarget);
  if (vmlText === null) {
    addIssue(
      diagnostics,
      options,
      createIssue(
        "COM-R002",
        DiagnosticSeverity.Recoverable,
        "VML drawing part '" + vmlTarget + "' was referenced but not found; default shape dimensions will be used.",
        true,
        false,
        sheetName,
        null,
      ),
    );
    return;
  }

  loadVmlDrawing(worksheetModel, vmlText);
}

function loadCommentsXml(
  worksheetModel: WorksheetModel,
  xmlText: string,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  let document: XmlNode;
  try {
    document = parseXml(xmlText);
  } catch {
    return;
  }

  const root = documentRoot(document);
  if (root === null) {
    return;
  }

  const authors: string[] = [];
  const authorsElement = firstElement(root, "authors");
  if (authorsElement !== null) {
    for (const authorElement of elementChildren(authorsElement, "author")) {
      authors.push(textOf(authorElement));
    }
  }

  const commentListElement = firstElement(root, "commentList");
  if (commentListElement === null) {
    return;
  }

  for (const commentElement of elementChildren(commentListElement, "comment")) {
    const cellRef = attribute(commentElement, "ref");
    if (cellRef === null || cellRef.trim().length === 0) {
      addIssue(
        diagnostics,
        options,
        createIssue(
          "COM-R001",
          DiagnosticSeverity.Recoverable,
          "A comment with a missing ref attribute was skipped.",
          false,
          false,
          sheetName,
          null,
        ),
      );
      continue;
    }

    const address = tryParseCellReference(cellRef);
    if (address === null) {
      addIssue(
        diagnostics,
        options,
        createIssue(
          "COM-R001",
          DiagnosticSeverity.Recoverable,
          "A comment with invalid ref '" + cellRef + "' was skipped.",
          false,
          false,
          sheetName,
          cellRef,
        ),
      );
      continue;
    }

    let author = "";
    const authorIdValue = attribute(commentElement, "authorId");
    const authorId = tryParseInt(authorIdValue);
    if (authorId !== null && authorId >= 0 && authorId < authors.length) {
      author = authors[authorId];
    }

    const textElement = firstElement(commentElement, "text");
    const note = readNoteText(textElement);

    const model = new CommentModel();
    model.Row = address.RowIndex;
    model.Column = address.ColumnIndex;
    model.Author = author;
    model.Note = note;
    worksheetModel.Comments.push(model);
  }
}

function readNoteText(textElement: XmlNode | null): string {
  if (textElement === null) {
    return "";
  }

  let hasRuns = false;
  let note = "";
  for (const runElement of elementChildren(textElement, "r")) {
    hasRuns = true;
    const tElement = firstElement(runElement, "t");
    if (tElement !== null) {
      note += textOf(tElement);
    }
  }

  if (!hasRuns) {
    return textOf(textElement);
  }

  return note;
}

function loadVmlDrawing(worksheetModel: WorksheetModel, vmlText: string): void {
  const commentMap = new Map<string, CommentModel>();
  for (let index = 0; index < worksheetModel.Comments.length; index++) {
    const comment = worksheetModel.Comments[index];
    const key = comment.Row * 1048576 + comment.Column;
    const keyText = key.toString();
    if (!commentMap.has(keyText)) {
      commentMap.set(keyText, comment);
    }
  }

  const shapes = extractVmlShapes(vmlText);
  for (let index = 0; index < shapes.length; index++) {
    const shapeText = shapes[index];

    if (indexOfIgnoreCase(shapeText, 'ObjectType="Note"') < 0) {
      continue;
    }

    const row = extractTagInt(shapeText, "x:Row");
    const col = extractTagInt(shapeText, "x:Column");
    if (row < 0 || col < 0) {
      continue;
    }

    const mapKey = (row * 1048576 + col).toString();
    const model = commentMap.get(mapKey);
    if (model === undefined) {
      continue;
    }

    const styleAttr = extractAttributeValue(shapeText, 'style="');
    if (styleAttr !== null) {
      model.IsVisible = indexOfIgnoreCase(styleAttr, "visibility:visible") >= 0;

      const width = parseStylePt(styleAttr, "width:");
      if (width > 0) {
        model.Width = width;
      }

      const height = parseStylePt(styleAttr, "height:");
      if (height > 0) {
        model.Height = height;
      }
    }

    model.RawVmlShapeXml = shapeText;
  }
}

function extractVmlShapes(vmlText: string): string[] {
  const shapes: string[] = [];
  let searchStart = 0;
  const closeTag = "</v:shape>";

  while (true) {
    const shapeStart = findNextVmlShape(vmlText, searchStart);
    if (shapeStart < 0) {
      break;
    }

    const shapeEnd = vmlText.indexOf(closeTag, shapeStart);
    if (shapeEnd < 0) {
      break;
    }

    const endPosition = shapeEnd + closeTag.length;
    shapes.push(vmlText.substring(shapeStart, endPosition));
    searchStart = endPosition;
  }

  return shapes;
}

function findNextVmlShape(vmlText: string, searchFrom: number): number {
  let position = searchFrom;
  while (true) {
    const index = vmlText.indexOf("<v:shape", position);
    if (index < 0) {
      return -1;
    }

    const afterTag = index + 8;
    if (afterTag >= vmlText.length) {
      return -1;
    }

    const next = vmlText.charAt(afterTag);
    if (next === " " || next === ">" || next === "\r" || next === "\n" || next === "\t") {
      return index;
    }

    position = index + 1;
  }
}

function extractTagInt(text: string, tagName: string): number {
  const openTag = "<" + tagName + ">";
  const closeTag = "</" + tagName + ">";
  const start = text.indexOf(openTag);
  if (start < 0) {
    return -1;
  }

  const valueStart = start + openTag.length;
  const end = text.indexOf(closeTag, valueStart);
  if (end < 0) {
    return -1;
  }

  const parsed = tryParseInt(text.substring(valueStart, end).trim());
  return parsed === null ? -1 : parsed;
}

function extractAttributeValue(text: string, prefix: string): string | null {
  const start = text.indexOf(prefix);
  if (start < 0) {
    return null;
  }

  const valueStart = start + prefix.length;
  const end = text.indexOf('"', valueStart);
  if (end < 0) {
    return null;
  }

  return text.substring(valueStart, end);
}

function parseStylePt(style: string, prefix: string): number {
  const start = indexOfIgnoreCase(style, prefix);
  if (start < 0) {
    return 0;
  }

  const valueStart = start + prefix.length;
  let end = style.indexOf(";", valueStart);
  if (end < 0) {
    end = style.indexOf('"', valueStart);
  }

  if (end < 0) {
    end = style.length;
  }

  let valuePart = style.substring(valueStart, end).trim();
  if (valuePart.toLowerCase().endsWith("pt")) {
    valuePart = valuePart.substring(0, valuePart.length - 2);
  }

  const pt = tryParseDouble(valuePart);
  if (pt !== null && pt > 0) {
    return roundHalfToEven(pt / 0.75);
  }

  return 0;
}

function tryParseInt(value: string | null): number | null {
  if (value === null) {
    return null;
  }

  const trimmed = value.trim();
  if (!/^[+-]?\d+$/.test(trimmed)) {
    return null;
  }

  const parsed = Number(trimmed);
  if (parsed < -2147483648 || parsed > 2147483647) {
    return null;
  }

  return parsed;
}

function tryParseDouble(value: string): number | null {
  const trimmed = value.trim();
  if (trimmed.length === 0) {
    return null;
  }

  const parsed = Number(trimmed);
  return Number.isNaN(parsed) ? null : parsed;
}

function roundHalfToEven(value: number): number {
  const floor = Math.floor(value);
  if (value - floor === 0.5) {
    return floor % 2 === 0 ? floor : floor + 1;
  }

  return Math.round(value);
}

function indexOfIgnoreCase(text: string, search: string): number {
  return text.toLowerCase().indexOf(search.toLowerCase());
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
