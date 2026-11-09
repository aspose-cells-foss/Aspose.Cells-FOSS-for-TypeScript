import { CellAddress } from "./Core/CellAddress";
import { CellRecord } from "./Core/CellRecord";
import { CellValueKind } from "./Core/CellValueKind";
import { FontValue } from "./Core/FontValue";
import { MergeRegion } from "./Core/MergeRegion";
import { SharedStringEntry } from "./SharedStringEntry";
import { SharedStringRepository } from "./Core/SharedStringRepository";
import { StyleValue } from "./Core/StyleValue";
import { WorkbookModel } from "./Core/WorkbookModel";
import { SheetVisibility } from "./Core/SheetVisibility";
import { ZipWriter } from "./Packaging/Archive";
import { XmlElement, xmlDocument } from "./Xml/XmlWriter";
import { buildCalculationPropertiesElement, buildBookViewsElement, buildWorkbookPropertiesElement, buildWorkbookProtectionElement } from "./XlsxWorkbookProperties";
import { buildDefinedNames } from "./XlsxWorkbookDefinedNames";
import { buildStylesheet, buildRichTextRunProperties, stylesEqual } from "./XlsxWorkbookStyles";
import { contentTypeFromExtension } from "./XlsxWorkbookPictures";
import { COMMENTS_CONTENT_TYPE, VML_DRAWING_CONTENT_TYPE } from "./XlsxWorkbookComments";
import { TABLE_CONTENT_TYPE } from "./XlsxWorkbookTables";
import { DRAWING_CONTENT_TYPE } from "./XlsxWorkbookPictures";

export const WORKSHEET_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
export const SHARED_STRINGS_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
export const STYLES_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
export const EXTERNAL_LINK_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
export const EXTERNAL_LINK_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
export const CORE_PROPERTIES_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
export const EXTENDED_PROPERTIES_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";

export const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
export const RELATIONSHIP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const PACKAGE_RELATIONSHIP_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
export const CONTENT_TYPE_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
export const XML_NS = "http://www.w3.org/XML/1998/namespace";

/** Built-in number formats that Excel treats as date/time formats. */
export const BUILT_IN_DATE_FORMATS = new Set<number>([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);

/** Returns true when a cell record must be written to the worksheet part. */
export function shouldPersistCell(workbookDefaultStyle: StyleValue, record: CellRecord): boolean {
  return (
    record.IsExplicitlyStored ||
    (record.Formula !== null && record.Formula.length > 0) ||
    record.Value !== null ||
    record.Kind !== CellValueKind.Blank ||
    !stylesEqual(record.Style, workbookDefaultStyle)
  );
}

/** Writes an XML document part into the archive. */
export function writeXmlEntry(archive: ZipWriter, path: string, document: XmlElement): void {
  archive.addText(path, xmlDocument(document));
}

export function buildContentTypes(
  model: WorkbookModel,
  hasSharedStrings: boolean,
  hasDateStyles: boolean,
  hasCoreProperties: boolean,
  hasExtendedProperties: boolean,
  totalTableCount: number,
  totalDrawingCount: number,
  imageExtensions: readonly string[],
  chartPartNames: readonly string[],
  chartContentTypes: readonly string[],
  chartCompanionPartNames: readonly string[],
  chartCompanionContentTypes: readonly string[],
  hasTheme = false,
  totalCommentCount = 0,
): XmlElement {
  const root = XmlElement.createWithNamespace("Types", CONTENT_TYPE_NS)
    .child(
      XmlElement.create("Default")
        .attr("Extension", "rels")
        .attr("ContentType", "application/vnd.openxmlformats-package.relationships+xml"),
    )
    .child(
      XmlElement.create("Default")
        .attr("Extension", "xml")
        .attr("ContentType", "application/xml"),
    )
    .child(
      XmlElement.create("Override")
        .attr("PartName", "/xl/workbook.xml")
        .attr("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
    );

  if (hasCoreProperties) {
    root.child(
      XmlElement.create("Override")
        .attr("PartName", "/docProps/core.xml")
        .attr("ContentType", "application/vnd.openxmlformats-package.core-properties+xml"),
    );
  }

  if (hasExtendedProperties) {
    root.child(
      XmlElement.create("Override")
        .attr("PartName", "/docProps/app.xml")
        .attr("ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
    );
  }

  for (let index = 0; index < model.Worksheets.length; index++) {
    root.child(
      XmlElement.create("Override")
        .attr("PartName", `/xl/worksheets/sheet${index + 1}.xml`)
        .attr("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"),
    );
  }

  for (let tableNumber = 1; tableNumber <= totalTableCount; tableNumber++) {
    root.child(
      XmlElement.create("Override")
        .attr("PartName", `/xl/tables/table${tableNumber}.xml`)
        .attr("ContentType", TABLE_CONTENT_TYPE),
    );
  }

  for (let drawingNumber = 1; drawingNumber <= totalDrawingCount; drawingNumber++) {
    root.child(
      XmlElement.create("Override")
        .attr("PartName", `/xl/drawings/drawing${drawingNumber}.xml`)
        .attr("ContentType", DRAWING_CONTENT_TYPE),
    );
  }

  for (let c = 0; c < chartPartNames.length; c++) {
    root.child(
      XmlElement.create("Override")
        .attr("PartName", chartPartNames[c])
        .attr("ContentType", chartContentTypes[c]),
    );
  }

  for (let c = 0; c < chartCompanionPartNames.length; c++) {
    root.child(
      XmlElement.create("Override")
        .attr("PartName", chartCompanionPartNames[c])
        .attr("ContentType", chartCompanionContentTypes[c]),
    );
  }

  for (const extension of imageExtensions) {
    root.child(
      XmlElement.create("Default")
        .attr("Extension", extension)
        .attr("ContentType", contentTypeFromExtension(extension)),
    );
  }

  if (hasSharedStrings) {
    root.child(
      XmlElement.create("Override")
        .attr("PartName", "/xl/sharedStrings.xml")
        .attr("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"),
    );
  }

  if (hasDateStyles) {
    root.child(
      XmlElement.create("Override")
        .attr("PartName", "/xl/styles.xml")
        .attr("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"),
    );
  }

  if (hasTheme) {
    root.child(
      XmlElement.create("Override")
        .attr("PartName", "/xl/theme/theme1.xml")
        .attr("ContentType", "application/vnd.openxmlformats-officedocument.theme+xml"),
    );
  }

  for (let n = 1; n <= totalCommentCount; n++) {
    root.child(
      XmlElement.create("Override")
        .attr("PartName", `/xl/comments${n}.xml`)
        .attr("ContentType", COMMENTS_CONTENT_TYPE),
    );
    root.child(
      XmlElement.create("Override")
        .attr("PartName", `/xl/drawings/vmlDrawing${n}.vml`)
        .attr("ContentType", VML_DRAWING_CONTENT_TYPE),
    );
  }

  for (let index = 0; index < model.ExternalLinks.length; index++) {
    root.child(
      XmlElement.create("Override")
        .attr("PartName", `/xl/externalLinks/externalLink${index + 1}.xml`)
        .attr("ContentType", EXTERNAL_LINK_CONTENT_TYPE),
    );
  }

  return root;
}

export function buildRootRelationships(
  hasCoreProperties: boolean,
  hasExtendedProperties: boolean,
): XmlElement {
  const relationships = XmlElement.createWithNamespace("Relationships", PACKAGE_RELATIONSHIP_NS).child(
    XmlElement.create("Relationship")
      .attr("Id", "rId1")
      .attr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
      .attr("Target", "xl/workbook.xml"),
  );

  let relationshipId = 2;
  if (hasCoreProperties) {
    relationships.child(
      XmlElement.create("Relationship")
        .attr("Id", `rId${relationshipId}`)
        .attr("Type", CORE_PROPERTIES_RELATIONSHIP_TYPE)
        .attr("Target", "docProps/core.xml"),
    );
    relationshipId++;
  }

  if (hasExtendedProperties) {
    relationships.child(
      XmlElement.create("Relationship")
        .attr("Id", `rId${relationshipId}`)
        .attr("Type", EXTENDED_PROPERTIES_RELATIONSHIP_TYPE)
        .attr("Target", "docProps/app.xml"),
    );
  }

  return relationships;
}

export function buildWorkbook(model: WorkbookModel, externalLinkBaseRId: number): XmlElement {
  const workbook = XmlElement.createWithNamespace("workbook", MAIN_NS).attr("xmlns:r", RELATIONSHIP_NS);

  const workbookProperties = buildWorkbookPropertiesElement(model);
  if (workbookProperties) {
    workbook.child(workbookProperties);
  }

  const workbookProtection = buildWorkbookProtectionElement(model);
  if (workbookProtection) {
    workbook.child(workbookProtection);
  }

  const bookViews = buildBookViewsElement(model);
  if (bookViews) {
    workbook.child(bookViews);
  }

  const sheets = XmlElement.create("sheets");
  for (let index = 0; index < model.Worksheets.length; index++) {
    const worksheet = model.Worksheets[index];
    const sheet = XmlElement.create("sheet")
      .attr("name", worksheet.Name)
      .attr("sheetId", index + 1)
      .attr("r:id", `rId${index + 1}`);

    if (worksheet.Visibility === SheetVisibility.Hidden) {
      sheet.attr("state", "hidden");
    } else if (worksheet.Visibility === SheetVisibility.VeryHidden) {
      sheet.attr("state", "veryHidden");
    }

    sheets.child(sheet);
  }

  workbook.child(sheets);

  if (model.ExternalLinks.length > 0) {
    const externalReferences = XmlElement.create("externalReferences");
    for (let index = 0; index < model.ExternalLinks.length; index++) {
      externalReferences.child(
        XmlElement.create("externalReference").attr("r:id", `rId${externalLinkBaseRId + index}`),
      );
    }

    workbook.child(externalReferences);
  }

  const definedNames = buildDefinedNames(model);
  if (definedNames) {
    workbook.child(definedNames);
  }

  const calculationProperties = buildCalculationPropertiesElement(model);
  if (calculationProperties) {
    workbook.child(calculationProperties);
  }

  return workbook;
}

export function buildWorkbookRelationships(
  model: WorkbookModel,
  hasSharedStrings: boolean,
  hasDateStyles: boolean,
): XmlElement {
  const relationships = XmlElement.createWithNamespace("Relationships", PACKAGE_RELATIONSHIP_NS);
  let relationshipId = 1;

  for (let index = 0; index < model.Worksheets.length; index++) {
    relationships.child(
      XmlElement.create("Relationship")
        .attr("Id", `rId${relationshipId++}`)
        .attr("Type", WORKSHEET_RELATIONSHIP_TYPE)
        .attr("Target", `worksheets/sheet${index + 1}.xml`),
    );
  }

  if (hasSharedStrings) {
    relationships.child(
      XmlElement.create("Relationship")
        .attr("Id", `rId${relationshipId++}`)
        .attr("Type", SHARED_STRINGS_RELATIONSHIP_TYPE)
        .attr("Target", "sharedStrings.xml"),
    );
  }

  if (hasDateStyles) {
    relationships.child(
      XmlElement.create("Relationship")
        .attr("Id", `rId${relationshipId++}`)
        .attr("Type", STYLES_RELATIONSHIP_TYPE)
        .attr("Target", "styles.xml"),
    );
  }

  if (model.RawThemeXml !== null && model.RawThemeXml.length > 0) {
    relationships.child(
      XmlElement.create("Relationship")
        .attr("Id", `rId${relationshipId++}`)
        .attr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        .attr("Target", "theme/theme1.xml"),
    );
  }

  for (let index = 0; index < model.ExternalLinks.length; index++) {
    relationships.child(
      XmlElement.create("Relationship")
        .attr("Id", `rId${relationshipId}`)
        .attr("Type", EXTERNAL_LINK_RELATIONSHIP_TYPE)
        .attr("Target", `externalLinks/externalLink${index + 1}.xml`),
    );
    relationshipId++;
  }

  return relationships;
}

export function toRangeReference(region: MergeRegion): string {
  const first = new CellAddress(region.FirstRow, region.FirstColumn).ToString();
  const last = new CellAddress(
    region.FirstRow + region.TotalRows - 1,
    region.FirstColumn + region.TotalColumns - 1,
  ).ToString();
  return first === last ? first : `${first}:${last}`;
}

export function tryParseMergeReference(reference: string | null | undefined): MergeRegion | null {
  if (reference === null || reference === undefined || reference.trim().length === 0) {
    return null;
  }

  const parts = reference.split(":");
  if (parts.length === 1) {
    const single = tryParseCellReference(parts[0]);
    return single ? new MergeRegion(single.RowIndex, single.ColumnIndex, 1, 1) : null;
  }

  if (parts.length !== 2) {
    return null;
  }

  const first = tryParseCellReference(parts[0]);
  const last = tryParseCellReference(parts[1]);
  if (!first || !last || last.RowIndex < first.RowIndex || last.ColumnIndex < first.ColumnIndex) {
    return null;
  }

  return new MergeRegion(
    first.RowIndex,
    first.ColumnIndex,
    last.RowIndex - first.RowIndex + 1,
    last.ColumnIndex - first.ColumnIndex + 1,
  );
}

export function tryParseCellReference(cellReference: string | null | undefined): CellAddress | null {
  if (cellReference === null || cellReference === undefined) {
    return null;
  }

  try {
    return CellAddress.Parse(cellReference);
  } catch {
    return null;
  }
}

export function buildSharedStrings(sharedStrings: SharedStringRepository): XmlElement {
  const root = XmlElement.createWithNamespace("sst", MAIN_NS)
    .attr("count", sharedStrings.Entries.length)
    .attr("uniqueCount", sharedStrings.Entries.length);

  for (const entry of sharedStrings.Entries) {
    root.child(createSharedStringItemElement(entry));
  }

  return root;
}

export function createSharedStringItemElement(entry: SharedStringEntry | null): XmlElement {
  const safeEntry = entry ?? new SharedStringEntry();
  const runs = safeEntry.Runs;

  if (runs === null || runs.length === 0) {
    return XmlElement.create("si").child(createTextElement(safeEntry.Text ?? ""));
  }

  const item = XmlElement.create("si");
  const text = safeEntry.Text ?? "";
  for (const run of runs) {
    item.child(createRichTextRunElement(text, run));
  }

  return item;
}

export function createInlineStringElement(text: string | null, runs: readonly RichTextRunValueLike[] | null): XmlElement {
  if (runs === null || runs === undefined || runs.length === 0) {
    return XmlElement.create("is").child(createTextElement(text ?? ""));
  }

  const inlineString = XmlElement.create("is");
  const safeText = text ?? "";
  for (const run of runs) {
    inlineString.child(createRichTextRunElement(safeText, run));
  }

  return inlineString;
}

interface RichTextRunValueLike {
  readonly StartIndex: number;
  readonly Length: number;
  readonly Font: FontValue;
}

function createRichTextRunElement(text: string, run: RichTextRunValueLike | null): XmlElement {
  const runElement = XmlElement.create("r");
  runElement.child(buildRichTextRunProperties(run ? run.Font : null));

  const safeText = text ?? "";
  let startIndex = run ? run.StartIndex : 0;
  let length = run ? run.Length : 0;

  if (startIndex < 0) {
    startIndex = 0;
  }

  if (length < 0) {
    length = 0;
  }

  if (startIndex > safeText.length) {
    startIndex = safeText.length;
  }

  if (startIndex + length > safeText.length) {
    length = safeText.length - startIndex;
  }

  runElement.child(createTextElement(safeText.substring(startIndex, startIndex + length)));
  return runElement;
}

export function createTextElement(value: string): XmlElement {
  const text = XmlElement.create("t").child(value);
  if (needsPreserveWhitespace(value)) {
    text.attr("xml:space", "preserve");
  }

  return text;
}

export function needsPreserveWhitespace(value: string | null): boolean {
  if (value === null || value === undefined || value.length === 0) {
    return false;
  }

  return (
    isWhitespace(value[0]) ||
    isWhitespace(value[value.length - 1]) ||
    value.includes("\n") ||
    value.includes("\r") ||
    value.includes("\t")
  );
}

function isWhitespace(character: string): boolean {
  return character === " " || character === "\t" || character === "\n" || character === "\r" || character === "\f" || character === "\v";
}

export function buildMinimalStylesheet(): XmlElement {
  return XmlElement.createWithNamespace("styleSheet", MAIN_NS)
    .child(
      XmlElement.create("fonts")
        .attr("count", 1)
        .child(
          XmlElement.create("font")
            .child(XmlElement.create("sz").attr("val", 11))
            .child(XmlElement.create("name").attr("val", "Calibri")),
        ),
    )
    .child(
      XmlElement.create("fills")
        .attr("count", 2)
        .child(XmlElement.create("fill").child(XmlElement.create("patternFill").attr("patternType", "none")))
        .child(XmlElement.create("fill").child(XmlElement.create("patternFill").attr("patternType", "gray125"))),
    )
    .child(
      XmlElement.create("borders")
        .attr("count", 1)
        .child(
          XmlElement.create("border")
            .child(XmlElement.create("left"))
            .child(XmlElement.create("right"))
            .child(XmlElement.create("top"))
            .child(XmlElement.create("bottom"))
            .child(XmlElement.create("diagonal")),
        ),
    )
    .child(
      XmlElement.create("cellStyleXfs")
        .attr("count", 1)
        .child(
          XmlElement.create("xf")
            .attr("numFmtId", 0)
            .attr("fontId", 0)
            .attr("fillId", 0)
            .attr("borderId", 0),
        ),
    )
    .child(
      XmlElement.create("cellXfs")
        .attr("count", 2)
        .child(
          XmlElement.create("xf")
            .attr("numFmtId", 0)
            .attr("fontId", 0)
            .attr("fillId", 0)
            .attr("borderId", 0)
            .attr("xfId", 0),
        )
        .child(
          XmlElement.create("xf")
            .attr("numFmtId", 14)
            .attr("fontId", 0)
            .attr("fillId", 0)
            .attr("borderId", 0)
            .attr("xfId", 0)
            .attr("applyNumberFormat", 1),
        ),
    )
    .child(
      XmlElement.create("cellStyles")
        .attr("count", 1)
        .child(
          XmlElement.create("cellStyle")
            .attr("name", "Normal")
            .attr("xfId", 0)
            .attr("builtinId", 0),
        ),
    );
}

export { buildStylesheet, stylesEqual };
