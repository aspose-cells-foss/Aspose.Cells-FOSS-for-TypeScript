import { CellAddress } from "./Core/CellAddress";
import { CommentModel } from "./Core/CommentModel";
import { WorkbookModel } from "./Core/WorkbookModel";
import { WorksheetModel } from "./Core/WorksheetModel";
import { ZipWriter } from "./Packaging/Archive";
import { createTextElement, MAIN_NS, writeXmlEntry } from "./XlsxWorkbookSerializerCommon";
import { XmlElement } from "./Xml/XmlWriter";

export const COMMENTS_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
export const VML_DRAWING_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
export const COMMENTS_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
export const VML_DRAWING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.vmlDrawing";

export { loadComments } from "./XlsxWorkbookCommentsLoader";

/** Builds the comments part document for a worksheet. */
export function buildCommentsDocument(worksheet: WorksheetModel): XmlElement {
  const sorted = sortComments(worksheet.Comments);

  const authors: string[] = [];
  const authorToIndex = new Map<string, number>();

  for (let index = 0; index < sorted.length; index++) {
    const author = sorted[index].Author ?? "";
    if (!authorToIndex.has(author)) {
      authorToIndex.set(author, authors.length);
      authors.push(author);
    }
  }

  const authorsElement = XmlElement.create("authors");
  for (let index = 0; index < authors.length; index++) {
    authorsElement.child(XmlElement.create("author").child(authors[index]));
  }

  const commentListElement = XmlElement.create("commentList");
  for (let index = 0; index < sorted.length; index++) {
    const comment = sorted[index];
    const author = comment.Author ?? "";
    const authorId = authorToIndex.get(author) ?? 0;

    const cellRef = new CellAddress(comment.Row, comment.Column).ToString();
    const note = comment.Note ?? "";

    const commentElement = XmlElement.create("comment")
      .attr("ref", cellRef)
      .attr("authorId", authorId)
      .child(
        XmlElement.create("text").child(
          XmlElement.create("r").child(createTextElement(note)),
        ),
      );

    commentListElement.child(commentElement);
  }

  return XmlElement.createWithNamespace("comments", MAIN_NS)
    .child(authorsElement)
    .child(commentListElement);
}

/** Writes the comments XML part and its VML drawing part. */
export function writeCommentFiles(
  archive: ZipWriter,
  worksheet: WorksheetModel,
  commentFileNumber: number,
): void {
  writeXmlEntry(archive, "xl/comments" + commentFileNumber + ".xml", buildCommentsDocument(worksheet));
  writeVmlDrawing(archive, worksheet, commentFileNumber);
}

/** Writes the VML drawing part for the worksheet comments. */
export function writeVmlDrawing(
  archive: ZipWriter,
  worksheet: WorksheetModel,
  commentFileNumber: number,
): void {
  const path = "xl/drawings/vmlDrawing" + commentFileNumber + ".vml";
  archive.addText(path, buildVmlContent(worksheet));
}

function buildVmlContent(worksheet: WorksheetModel): string {
  let content = "";
  content += '<xml xmlns:v="urn:schemas-microsoft-com:vml"';
  content += ' xmlns:o="urn:schemas-microsoft-com:office:office"';
  content += ' xmlns:x="urn:schemas-microsoft-com:office:excel">';
  content += '<o:shapelayout v:ext="edit">';
  content += '<o:idmap v:ext="edit" data="1"/>';
  content += "</o:shapelayout>";
  content += '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"';
  content += ' path="m,l,21600r21600,l21600,xe">';
  content += '<v:stroke joinstyle="miter"/>';
  content += '<v:path gradientshapeok="t" o:connecttype="rect"/>';
  content += "</v:shapetype>";

  const sorted = sortComments(worksheet.Comments);
  for (let index = 0; index < sorted.length; index++) {
    const comment = sorted[index];
    const shapeId = 1025 + index;
    if (comment.RawVmlShapeXml !== null) {
      content += patchVmlShape(comment.RawVmlShapeXml, comment, shapeId);
    } else {
      content += buildDefaultVmlShape(comment, shapeId);
    }
  }

  content += "</xml>";
  return content;
}

function patchVmlShape(raw: string, comment: CommentModel, shapeId: number): string {
  let value = replaceInlineValue(raw, 'id="_x0000_s', '"', shapeId.toString());

  const widthPt = formatPt(comment.Width);
  const heightPt = formatPt(comment.Height);
  value = replaceInlineValue(value, "width:", "pt", widthPt);
  value = replaceInlineValue(value, "height:", "pt", heightPt);

  value = setStyleVisibility(value, comment.IsVisible);
  value = replaceTagContent(value, "x:Row", comment.Row.toString());
  value = replaceTagContent(value, "x:Column", comment.Column.toString());
  value = setClientDataVisible(value, comment.IsVisible);

  return value;
}

function buildDefaultVmlShape(comment: CommentModel, shapeId: number): string {
  const widthPt = formatPt(comment.Width);
  const heightPt = formatPt(comment.Height);
  const visStyle = comment.IsVisible ? ";visibility:visible" : "";

  const anchorLeft = comment.Column + 1;
  const anchorTop = comment.Row > 0 ? comment.Row - 1 : 0;
  const anchorRight = comment.Column + 4;
  const anchorBottom = comment.Row + 5;

  let content = "";
  content += '<v:shape id="_x0000_s';
  content += shapeId.toString();
  content += '" type="#_x0000_t202"';
  content += ' style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:';
  content += widthPt;
  content += "pt;height:";
  content += heightPt;
  content += "pt;z-index:1";
  content += visStyle;
  content += '" fillcolor="#ffffe1" o:insetmode="auto">';
  content += '<v:fill color2="#ffffe1"/>';
  content += '<v:shadow on="t" color="black" obscured="t"/>';
  content += '<v:path o:connecttype="none"/>';
  content += '<v:textbox style="mso-direction-alt:auto">';
  content += '<div style="text-align:left"/>';
  content += "</v:textbox>";
  content += '<x:ClientData ObjectType="Note">';
  content += "<x:MoveWithCells/>";
  content += "<x:SizeWithCells/>";
  content += "<x:Anchor>";
  content += anchorLeft.toString();
  content += ",15,";
  content += anchorTop.toString();
  content += ",2,";
  content += anchorRight.toString();
  content += ",15,";
  content += anchorBottom.toString();
  content += ",16</x:Anchor>";
  content += "<x:AutoFill>False</x:AutoFill>";
  content += "<x:Row>";
  content += comment.Row.toString();
  content += "</x:Row>";
  content += "<x:Column>";
  content += comment.Column.toString();
  content += "</x:Column>";
  if (comment.IsVisible) {
    content += "<x:Visible/>";
  }

  content += "</x:ClientData>";
  content += "</v:shape>";
  return content;
}

/** Formats a pixel value as points using the C# "0.##" pattern. */
export function formatPt(pixels: number): string {
  const pt = pixels * 0.75;
  return formatAtMostTwoDecimals(pt);
}

function formatAtMostTwoDecimals(value: number): string {
  const rounded = Math.round(value * 100) / 100;
  let text = rounded.toFixed(2);
  if (text.indexOf(".") >= 0) {
    text = text.replace(/0+$/, "").replace(/\.$/, "");
  }

  return text;
}

/** Returns a copy of the comments sorted by row then column. */
export function sortComments(comments: readonly CommentModel[]): CommentModel[] {
  const sorted = comments.slice();
  for (let i = 0; i < sorted.length - 1; i++) {
    for (let j = i + 1; j < sorted.length; j++) {
      const a = sorted[i];
      const b = sorted[j];
      if (a.Row > b.Row || (a.Row === b.Row && a.Column > b.Column)) {
        sorted[i] = b;
        sorted[j] = a;
      }
    }
  }

  return sorted;
}

/** Replaces the value between prefix and the next occurrence of suffix. */
export function replaceInlineValue(
  text: string,
  prefix: string,
  suffix: string,
  newValue: string,
): string {
  const start = text.indexOf(prefix);
  if (start < 0) {
    return text;
  }

  const valueStart = start + prefix.length;
  const end = text.indexOf(suffix, valueStart);
  if (end < 0) {
    return text;
  }

  return text.substring(0, valueStart) + newValue + text.substring(end);
}

/** Replaces the text content of the first <tagName>...</tagName> occurrence. */
export function replaceTagContent(text: string, tagName: string, newContent: string): string {
  const openTag = "<" + tagName + ">";
  const closeTag = "</" + tagName + ">";
  const start = text.indexOf(openTag);
  if (start < 0) {
    return text;
  }

  const contentStart = start + openTag.length;
  const end = text.indexOf(closeTag, contentStart);
  if (end < 0) {
    return text;
  }

  return text.substring(0, contentStart) + newContent + text.substring(end);
}

function setStyleVisibility(raw: string, isVisible: boolean): string {
  const hasVisible = raw.indexOf("visibility:visible") >= 0;

  if (isVisible && !hasVisible) {
    const stylePrefix = 'style="';
    const styleStart = raw.indexOf(stylePrefix);
    if (styleStart >= 0) {
      const quoteStart = styleStart + stylePrefix.length;
      const quoteEnd = raw.indexOf('"', quoteStart);
      if (quoteEnd >= 0) {
        return raw.substring(0, quoteEnd) + ";visibility:visible" + raw.substring(quoteEnd);
      }
    }

    return raw;
  }

  if (!isVisible && hasVisible) {
    const withSemi = ";visibility:visible";
    let position = raw.indexOf(withSemi);
    if (position >= 0) {
      return raw.substring(0, position) + raw.substring(position + withSemi.length);
    }

    const noSemi = "visibility:visible";
    position = raw.indexOf(noSemi);
    if (position >= 0) {
      return raw.substring(0, position) + raw.substring(position + noSemi.length);
    }
  }

  return raw;
}

function setClientDataVisible(raw: string, isVisible: boolean): string {
  const hasTag = raw.indexOf("<x:Visible/>") >= 0;

  if (isVisible && !hasTag) {
    const endTag = "</x:ClientData>";
    const position = raw.indexOf(endTag);
    if (position >= 0) {
      return raw.substring(0, position) + "<x:Visible/>" + raw.substring(position);
    }
  } else if (!isVisible && hasTag) {
    return raw.split("<x:Visible/>").join("");
  }

  return raw;
}

/** Computes the comment part number for each worksheet (0 when it has no comments). */
export function computeCommentFileNumbers(model: WorkbookModel): number[] {
  const numbers = new Array<number>(model.Worksheets.length);
  let counter = 0;
  for (let index = 0; index < model.Worksheets.length; index++) {
    if (model.Worksheets[index].Comments.length > 0) {
      counter++;
      numbers[index] = counter;
    } else {
      numbers[index] = 0;
    }
  }

  return numbers;
}

/** Counts the worksheets that contain comments. */
export function computeTotalCommentCount(model: WorkbookModel): number {
  let total = 0;
  for (let index = 0; index < model.Worksheets.length; index++) {
    if (model.Worksheets[index].Comments.length > 0) {
      total++;
    }
  }

  return total;
}
