import { ChartModel } from "./Core/ChartModel";
import { WorkbookModel } from "./Core/WorkbookModel";
import { WorksheetModel } from "./Core/WorksheetModel";
import { ZipWriter } from "./Packaging/Archive";
import { XmlElement } from "./Xml/XmlWriter";
import { documentRoot, XmlNode, parseXml } from "./Xml/XmlReader";

export const DRAWING_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
export const IMAGE_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
export const DRAWING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawing+xml";
export const CHART_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
export const CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
export const CHART_EX_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
export const CHART_EX_CONTENT_TYPE = "application/vnd.ms-office.chartex+xml";
export const CHART_STYLE_CONTENT_TYPE = "application/vnd.ms-office.chartstyle+xml";
export const CHART_COLOR_STYLE_CONTENT_TYPE = "application/vnd.ms-office.chartcolorstyle+xml";
export const CHART_GRAPHIC_DATA_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";
export const CHART_STYLE_RELATIONSHIP_TYPE =
  "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
export const CHART_COLOR_STYLE_RELATIONSHIP_TYPE =
  "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
export const CHART_USER_SHAPES_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes";
export const CHART_USER_SHAPES_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml";

export const XDR_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
export const DRAWING_MAIN_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const CHART_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const RELATIONSHIP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_RELATIONSHIP_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DRAWING_2014_NS = "http://schemas.microsoft.com/office/drawing/2014/main";

const EMU_PER_COLUMN = 609600;
const EMU_PER_ROW = 190500;

export function contentTypeFromExtension(extension: string): string {
  switch (extension) {
    case "jpeg":
    case "jpg":
      return "image/jpeg";
    case "png":
      return "image/png";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    default:
      return "application/octet-stream";
  }
}

export function computePictureFileOffsets(model: WorkbookModel): number[] {
  const offsets: number[] = new Array<number>(model.Worksheets.length);
  let running = 0;
  for (let index = 0; index < model.Worksheets.length; index++) {
    offsets[index] = running;
    running += model.Worksheets[index].Pictures.length + model.Worksheets[index].ShapeImages.length;
  }

  return offsets;
}

export function computeDrawingNumbers(model: WorkbookModel): number[] {
  const numbers: number[] = new Array<number>(model.Worksheets.length);
  let drawingCounter = 0;
  for (let index = 0; index < model.Worksheets.length; index++) {
    const worksheet = model.Worksheets[index];
    if (
      worksheet.Pictures.length > 0 ||
      worksheet.Shapes.length > 0 ||
      worksheet.Charts.length > 0 ||
      worksheet.PreservedCharts.length > 0
    ) {
      drawingCounter++;
      numbers[index] = drawingCounter;
    } else {
      numbers[index] = 0;
    }
  }

  return numbers;
}

export function computeTotalDrawingCount(model: WorkbookModel): number {
  let total = 0;
  for (let index = 0; index < model.Worksheets.length; index++) {
    const worksheet = model.Worksheets[index];
    if (
      worksheet.Pictures.length > 0 ||
      worksheet.Shapes.length > 0 ||
      worksheet.Charts.length > 0 ||
      worksheet.PreservedCharts.length > 0
    ) {
      total++;
    }
  }

  return total;
}

export function collectImageExtensions(model: WorkbookModel): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (let index = 0; index < model.Worksheets.length; index++) {
    const worksheet = model.Worksheets[index];
    for (let pictureIndex = 0; pictureIndex < worksheet.Pictures.length; pictureIndex++) {
      addExtension(seen, result, worksheet.Pictures[pictureIndex].ImageExtension);
    }

    for (let imageIndex = 0; imageIndex < worksheet.ShapeImages.length; imageIndex++) {
      addExtension(seen, result, worksheet.ShapeImages[imageIndex].Extension);
    }

    for (let chartIndex = 0; chartIndex < worksheet.Charts.length; chartIndex++) {
      collectChartCompanionImageExtensions(worksheet.Charts[chartIndex], seen, result);
    }

    for (let chartIndex = 0; chartIndex < worksheet.PreservedCharts.length; chartIndex++) {
      collectChartCompanionImageExtensions(worksheet.PreservedCharts[chartIndex], seen, result);
    }
  }

  return result;
}

function addExtension(seen: Set<string>, result: string[], extension: string | null | undefined): void {
  if (extension === null || extension === undefined || extension.length === 0) {
    return;
  }

  const key = extension.toLowerCase();
  if (!seen.has(key)) {
    seen.add(key);
    result.push(extension);
  }
}

function collectChartCompanionImageExtensions(chart: ChartModel, seen: Set<string>, result: string[]): void {
  for (let index = 0; index < chart.CompanionFiles.length; index++) {
    const companion = chart.CompanionFiles[index];
    if (!equalsIgnoreCase(companion.RelationshipType, IMAGE_RELATIONSHIP_TYPE)) {
      continue;
    }

    addExtension(seen, result, getCompanionFileExtension(companion.FileName));
  }
}

function getCompanionFileExtension(target: string | null): string {
  if (target === null || target.length === 0) {
    return "";
  }

  const normalized = target.replace(/\\/g, "/");
  const slashIndex = normalized.lastIndexOf("/");
  const fileName = slashIndex >= 0 ? normalized.substring(slashIndex + 1) : normalized;
  if (fileName.length === 0) {
    return "";
  }

  const dotIndex = fileName.lastIndexOf(".");
  if (dotIndex < 0 || dotIndex === fileName.length - 1) {
    return "";
  }

  return fileName.substring(dotIndex + 1).toLowerCase();
}

function equalsIgnoreCase(left: string | null, right: string): boolean {
  return left !== null && left.toLowerCase() === right.toLowerCase();
}

export function buildDrawingDocument(
  worksheet: WorksheetModel,
  pictureFileOffset: number,
  chartFileOffset: number,
): XmlElement {
  const root = XmlElement.create("xdr:wsDr")
    .attr("xmlns:xdr", XDR_NAMESPACE)
    .attr("xmlns:a", DRAWING_MAIN_NAMESPACE)
    .attr("xmlns:r", RELATIONSHIP_NS);
  let nextGeneratedObjectId = getNextGeneratedObjectIdStart(worksheet);

  const imageCount = worksheet.Pictures.length + worksheet.ShapeImages.length;
  const rIdMap = buildRelationshipIdMap(worksheet);

  for (let index = 0; index < worksheet.Pictures.length; index++) {
    const picture = worksheet.Pictures[index];
    const rId = `rId${index + 1}`;
    const picId = nextGeneratedObjectId;
    nextGeneratedObjectId++;
    root.child(buildTwoCellAnchor(picture, rId, picId));
  }

  for (let index = 0; index < worksheet.Shapes.length; index++) {
    const shape = worksheet.Shapes[index];
    const shapeId = nextGeneratedObjectId;
    nextGeneratedObjectId++;
    root.child(buildShapeAnchor(shape, shapeId, rIdMap));
  }

  for (let index = 0; index < worksheet.Charts.length; index++) {
    const chart = worksheet.Charts[index];
    const rId = `rId${imageCount + index + 1}`;
    const chartId = nextGeneratedObjectId;
    nextGeneratedObjectId++;
    root.child(buildChartAnchor(chart, rId, chartId));
  }

  return root;
}

function getNextGeneratedObjectIdStart(worksheet: WorksheetModel): number {
  let maxExistingId = 0;
  for (let index = 0; index < worksheet.Shapes.length; index++) {
    maxExistingId = updateMaxDrawingObjectId(worksheet.Shapes[index].RawElementXml, maxExistingId);
  }

  for (let index = 0; index < worksheet.Charts.length; index++) {
    const chart = worksheet.Charts[index];
    if (chart.IsChartEx) {
      maxExistingId = updateMaxDrawingObjectId(chart.RawGraphicFrameXml, maxExistingId);
    }
  }

  for (let index = 0; index < worksheet.PreservedCharts.length; index++) {
    maxExistingId = updateMaxDrawingObjectId(worksheet.PreservedCharts[index].RawGraphicFrameXml, maxExistingId);
  }

  return maxExistingId + 1;
}

function updateMaxDrawingObjectId(rawXml: string | null, maxExistingId: number): number {
  if (rawXml === null || rawXml.length === 0) {
    return maxExistingId;
  }

  const root = tryParseElement(rawXml);
  if (root === null) {
    return maxExistingId;
  }

  const elements = collectDescendantElements(root, "cNvPr");
  for (const element of elements) {
    const idValue = attributeValue(element, "id");
    if (idValue === null) {
      continue;
    }

    const parsedId = tryParseInt(idValue);
    if (parsedId === null) {
      continue;
    }

    if (parsedId > maxExistingId) {
      maxExistingId = parsedId;
    }
  }

  return maxExistingId;
}

function buildRelationshipIdMap(worksheet: WorksheetModel): Map<string, string> {
  const map = new Map<string, string>();
  for (let index = 0; index < worksheet.Pictures.length; index++) {
    const originalRId = worksheet.Pictures[index].OriginalRId;
    if (originalRId !== null && originalRId.length > 0) {
      const newRId = `rId${index + 1}`;
      if (originalRId !== newRId) {
        map.set(originalRId, newRId);
      }
    }
  }

  for (let index = 0; index < worksheet.ShapeImages.length; index++) {
    const originalRId = worksheet.ShapeImages[index].OriginalRId;
    if (originalRId !== null && originalRId.length > 0) {
      const newRId = `rId${worksheet.Pictures.length + index + 1}`;
      if (originalRId !== newRId) {
        map.set(originalRId, newRId);
      }
    }
  }

  let chartRelationshipIndex = worksheet.Pictures.length + worksheet.ShapeImages.length;
  for (let index = 0; index < worksheet.Charts.length; index++) {
    const originalRId = worksheet.Charts[index].OriginalRId;
    if (originalRId !== null && originalRId.length > 0) {
      const newRId = `rId${chartRelationshipIndex + index + 1}`;
      if (originalRId !== newRId) {
        map.set(originalRId, newRId);
      }
    }
  }

  chartRelationshipIndex += worksheet.Charts.length;
  for (let index = 0; index < worksheet.PreservedCharts.length; index++) {
    const originalRId = worksheet.PreservedCharts[index].OriginalRId;
    if (originalRId !== null && originalRId.length > 0) {
      const newRId = `rId${chartRelationshipIndex + index + 1}`;
      if (originalRId !== newRId) {
        map.set(originalRId, newRId);
      }
    }
  }

  chartRelationshipIndex += worksheet.PreservedCharts.length;
  for (let index = 0; index < worksheet.PreservedDrawingRelationships.length; index++) {
    const originalRId = worksheet.PreservedDrawingRelationships[index].OriginalRId;
    if (originalRId !== null && originalRId.length > 0) {
      const newRId = `rId${chartRelationshipIndex + index + 1}`;
      if (originalRId !== newRId) {
        map.set(originalRId, newRId);
      }
    }
  }

  return map;
}

function buildTwoCellAnchor(picture: {
  Name: string;
  ImageExtension: string;
  UpperLeftColumn: number;
  UpperLeftColumnOffset: number;
  UpperLeftRow: number;
  UpperLeftRowOffset: number;
  LowerRightColumn: number;
  LowerRightColumnOffset: number;
  LowerRightRow: number;
  LowerRightRowOffset: number;
  ExtentCx: number;
  ExtentCy: number;
}, rId: string, picId: number): XmlElement {
  const anchor = XmlElement.create("xdr:twoCellAnchor").attr("editAs", "oneCell");
  anchor.child(
    buildFromElement(
      picture.UpperLeftColumn,
      picture.UpperLeftColumnOffset,
      picture.UpperLeftRow,
      picture.UpperLeftRowOffset,
    ),
  );
  anchor.child(
    buildToElement(
      picture.LowerRightColumn,
      picture.LowerRightColumnOffset,
      picture.LowerRightRow,
      picture.LowerRightRowOffset,
    ),
  );
  anchor.child(buildPicElement(picture, rId, picId));
  anchor.child(XmlElement.create("xdr:clientData"));
  return anchor;
}

function buildFromElement(col: number, colOff: number, row: number, rowOff: number): XmlElement {
  return XmlElement.create("xdr:from")
    .child(XmlElement.create("xdr:col").child(String(col)))
    .child(XmlElement.create("xdr:colOff").child(String(colOff)))
    .child(XmlElement.create("xdr:row").child(String(row)))
    .child(XmlElement.create("xdr:rowOff").child(String(rowOff)));
}

function buildToElement(col: number, colOff: number, row: number, rowOff: number): XmlElement {
  return XmlElement.create("xdr:to")
    .child(XmlElement.create("xdr:col").child(String(col)))
    .child(XmlElement.create("xdr:colOff").child(String(colOff)))
    .child(XmlElement.create("xdr:row").child(String(row)))
    .child(XmlElement.create("xdr:rowOff").child(String(rowOff)));
}

function buildPicElement(
  picture: {
    Name: string;
    ImageExtension: string;
    UpperLeftColumn: number;
    UpperLeftRow: number;
    LowerRightColumn: number;
    LowerRightRow: number;
    ExtentCx: number;
    ExtentCy: number;
  },
  rId: string,
  picId: number,
): XmlElement {
  let cx =
    picture.ExtentCx > 0
      ? picture.ExtentCx
      : (picture.LowerRightColumn - picture.UpperLeftColumn) * EMU_PER_COLUMN;
  let cy =
    picture.ExtentCy > 0
      ? picture.ExtentCy
      : (picture.LowerRightRow - picture.UpperLeftRow) * EMU_PER_ROW;

  if (cx <= 0) {
    cx = EMU_PER_COLUMN;
  }

  if (cy <= 0) {
    cy = EMU_PER_ROW;
  }

  const pic = XmlElement.create("xdr:pic");
  pic.child(
    XmlElement.create("xdr:nvPicPr")
      .child(
        XmlElement.create("xdr:cNvPr")
          .attr("id", picId)
          .attr(
            "name",
            picture.Name.length === 0 ? `Picture ${picId}` : picture.Name,
          ),
      )
      .child(
        XmlElement.create("xdr:cNvPicPr").child(
          XmlElement.create("a:picLocks").attr("noChangeAspect", 1),
        ),
      ),
  );

  pic.child(
    XmlElement.create("xdr:blipFill")
      .child(XmlElement.create("a:blip").attr("r:embed", rId))
      .child(XmlElement.create("a:stretch").child(XmlElement.create("a:fillRect"))),
  );

  pic.child(
    XmlElement.create("xdr:spPr")
      .child(
        XmlElement.create("a:xfrm")
          .child(XmlElement.create("a:off").attr("x", "0").attr("y", "0"))
          .child(XmlElement.create("a:ext").attr("cx", cx).attr("cy", cy)),
      )
      .child(XmlElement.create("a:prstGeom").attr("prst", "rect")),
  );

  return pic;
}

function buildShapeAnchor(shape: {
  RawElementXml: string | null;
  Name: string;
  GeometryType: string;
  RawStyleXml: string | null;
  RawTxBodyXml: string | null;
  UpperLeftColumn: number;
  UpperLeftColumnOffset: number;
  UpperLeftRow: number;
  UpperLeftRowOffset: number;
  LowerRightColumn: number;
  LowerRightColumnOffset: number;
  LowerRightRow: number;
  LowerRightRowOffset: number;
  ExtentCx: number;
  ExtentCy: number;
}, shapeId: number, pictureRIdMap: Map<string, string>): XmlElement {
  const anchor = XmlElement.create("xdr:twoCellAnchor").attr("editAs", "oneCell");
  anchor.child(
    buildFromElement(
      shape.UpperLeftColumn,
      shape.UpperLeftColumnOffset,
      shape.UpperLeftRow,
      shape.UpperLeftRowOffset,
    ),
  );
  anchor.child(
    buildToElement(
      shape.LowerRightColumn,
      shape.LowerRightColumnOffset,
      shape.LowerRightRow,
      shape.LowerRightRowOffset,
    ),
  );

  let rawElement: XmlElement | null = null;
  if (shape.RawElementXml !== null && shape.RawElementXml.length > 0) {
    const remapped = applyRIdRemap(shape.RawElementXml, pictureRIdMap);
    rawElement = tryParseMarkup(remapped);
  }

  anchor.child(rawElement !== null ? rawElement : buildSpElement(shape, shapeId));
  anchor.child(XmlElement.create("xdr:clientData"));
  return anchor;
}

function applyRIdRemap(rawXml: string, rIdMap: Map<string, string>): string {
  if (rawXml.length === 0) {
    return rawXml;
  }

  let normalizedXml = normalizePreservedDrawingMarkup(rawXml);
  if (rIdMap.size === 0) {
    return normalizedXml;
  }

  let placeholderIndex = 0;
  const placeholders: Array<[string, string]> = [];
  for (const [originalRId, newRId] of rIdMap) {
    const placeholder = `__ACF_RID_PLACEHOLDER_${placeholderIndex}__`;
    normalizedXml = normalizedXml.split(`"${originalRId}"`).join(`"${placeholder}"`);
    placeholders.push([placeholder, newRId]);
    placeholderIndex++;
  }

  for (const [placeholder, newRId] of placeholders) {
    normalizedXml = normalizedXml.split(`"${placeholder}"`).join(`"${newRId}"`);
  }

  return normalizedXml;
}

function buildSpElement(shape: {
  Name: string;
  GeometryType: string;
  RawStyleXml: string | null;
  RawTxBodyXml: string | null;
  UpperLeftColumn: number;
  UpperLeftRow: number;
  LowerRightColumn: number;
  LowerRightRow: number;
  ExtentCx: number;
  ExtentCy: number;
}, shapeId: number): XmlElement {
  let cx =
    shape.ExtentCx > 0
      ? shape.ExtentCx
      : (shape.LowerRightColumn - shape.UpperLeftColumn) * EMU_PER_COLUMN;
  let cy =
    shape.ExtentCy > 0
      ? shape.ExtentCy
      : (shape.LowerRightRow - shape.UpperLeftRow) * EMU_PER_ROW;

  if (cx <= 0) {
    cx = EMU_PER_COLUMN;
  }

  if (cy <= 0) {
    cy = EMU_PER_ROW;
  }

  const geomType = shape.GeometryType.length === 0 ? "rect" : shape.GeometryType;
  const displayName = shape.Name.length === 0 ? `Shape ${shapeId}` : shape.Name;

  const sp = XmlElement.create("xdr:sp").attr("macro", "").attr("textlink", "");

  sp.child(
    XmlElement.create("xdr:nvSpPr")
      .child(XmlElement.create("xdr:cNvPr").attr("id", shapeId).attr("name", displayName))
      .child(XmlElement.create("xdr:cNvSpPr")),
  );

  sp.child(
    XmlElement.create("xdr:spPr")
      .child(
        XmlElement.create("a:xfrm")
          .child(XmlElement.create("a:off").attr("x", "0").attr("y", "0"))
          .child(XmlElement.create("a:ext").attr("cx", cx).attr("cy", cy)),
      )
      .child(
        XmlElement.create("a:prstGeom").attr("prst", geomType).child(XmlElement.create("a:avLst")),
      ),
  );

  if (shape.RawStyleXml !== null && shape.RawStyleXml.length > 0) {
    const styleElement = tryParseMarkup(shape.RawStyleXml);
    if (styleElement !== null) {
      sp.child(styleElement);
    }
  } else {
    sp.child(buildDefaultShapeStyle());
  }

  if (shape.RawTxBodyXml !== null && shape.RawTxBodyXml.length > 0) {
    const txBodyElement = tryParseMarkup(shape.RawTxBodyXml);
    sp.child(txBodyElement !== null ? txBodyElement : buildMinimalTxBody());
  } else {
    sp.child(buildMinimalTxBody());
  }

  return sp;
}

function buildMinimalTxBody(): XmlElement {
  return XmlElement.create("xdr:txBody")
    .child(XmlElement.create("a:bodyPr"))
    .child(XmlElement.create("a:lstStyle"))
    .child(XmlElement.create("a:p"));
}

function buildDefaultShapeStyle(): XmlElement {
  return XmlElement.create("xdr:style")
    .child(
      XmlElement.create("a:lnRef")
        .attr("idx", 2)
        .child(
          XmlElement.create("a:schemeClr")
            .attr("val", "accent1")
            .child(XmlElement.create("a:shade").attr("val", "50000")),
        ),
    )
    .child(
      XmlElement.create("a:fillRef")
        .attr("idx", 1)
        .child(XmlElement.create("a:schemeClr").attr("val", "accent1")),
    )
    .child(
      XmlElement.create("a:effectRef")
        .attr("idx", 0)
        .child(XmlElement.create("a:schemeClr").attr("val", "accent1")),
    )
    .child(
      XmlElement.create("a:fontRef")
        .attr("idx", "minor")
        .child(XmlElement.create("a:schemeClr").attr("val", "lt1")),
    );
}

export function buildDrawingRelationshipsDocument(
  worksheet: WorksheetModel,
  pictureFileOffset: number,
  chartFileOffset: number,
): XmlElement {
  const relationships = XmlElement.createWithNamespace("Relationships", PACKAGE_RELATIONSHIP_NS);
  const imageCount = worksheet.Pictures.length + worksheet.ShapeImages.length;

  for (let index = 0; index < worksheet.Pictures.length; index++) {
    const picture = worksheet.Pictures[index];
    const globalPictureNumber = pictureFileOffset + index + 1;
    relationships.child(
      XmlElement.create("Relationship")
        .attr("Id", `rId${index + 1}`)
        .attr("Type", IMAGE_RELATIONSHIP_TYPE)
        .attr(
          "Target",
          `../media/image${globalPictureNumber}.${picture.ImageExtension}`,
        ),
    );
  }

  for (let index = 0; index < worksheet.ShapeImages.length; index++) {
    const shapeImage = worksheet.ShapeImages[index];
    const globalPictureNumber = pictureFileOffset + worksheet.Pictures.length + index + 1;
    relationships.child(
      XmlElement.create("Relationship")
        .attr("Id", `rId${worksheet.Pictures.length + index + 1}`)
        .attr("Type", IMAGE_RELATIONSHIP_TYPE)
        .attr("Target", `../media/image${globalPictureNumber}.${shapeImage.Extension ?? ""}`),
    );
  }

  for (let index = 0; index < worksheet.Charts.length; index++) {
    const chart = worksheet.Charts[index];
    const globalChartNumber = chartFileOffset + index + 1;
    const relType = chart.IsChartEx ? CHART_EX_RELATIONSHIP_TYPE : CHART_RELATIONSHIP_TYPE;
    const fileName = chart.IsChartEx ? `chartEx${globalChartNumber}.xml` : `chart${globalChartNumber}.xml`;
    relationships.child(
      XmlElement.create("Relationship")
        .attr("Id", `rId${imageCount + index + 1}`)
        .attr("Type", relType)
        .attr("Target", `../charts/${fileName}`),
    );
  }

  for (let index = 0; index < worksheet.PreservedCharts.length; index++) {
    const chart = worksheet.PreservedCharts[index];
    const globalChartNumber = chartFileOffset + worksheet.Charts.length + index + 1;
    const relType = chart.IsChartEx ? CHART_EX_RELATIONSHIP_TYPE : CHART_RELATIONSHIP_TYPE;
    const fileName = chart.IsChartEx ? `chartEx${globalChartNumber}.xml` : `chart${globalChartNumber}.xml`;
    relationships.child(
      XmlElement.create("Relationship")
        .attr("Id", `rId${imageCount + worksheet.Charts.length + index + 1}`)
        .attr("Type", relType)
        .attr("Target", `../charts/${fileName}`),
    );
  }

  const preservedRelationshipIndex =
    imageCount + worksheet.Charts.length + worksheet.PreservedCharts.length;
  for (let index = 0; index < worksheet.PreservedDrawingRelationships.length; index++) {
    const relationship = worksheet.PreservedDrawingRelationships[index];
    const relationshipElement = XmlElement.create("Relationship")
      .attr("Id", `rId${preservedRelationshipIndex + index + 1}`)
      .attr("Type", relationship.RelationshipType ?? "")
      .attr("Target", relationship.Target ?? "");

    if (relationship.TargetMode !== null && relationship.TargetMode.length > 0) {
      relationshipElement.attr("TargetMode", relationship.TargetMode);
    }

    relationships.child(relationshipElement);
  }

  return relationships;
}

export function writePictureMediaEntries(
  archive: ZipWriter,
  worksheet: WorksheetModel,
  pictureFileOffset: number,
): void {
  for (let index = 0; index < worksheet.Pictures.length; index++) {
    const picture = worksheet.Pictures[index];
    const globalPictureNumber = pictureFileOffset + index + 1;
    const path = `xl/media/image${globalPictureNumber}.${picture.ImageExtension}`;
    archive.addBinary(path, picture.ImageData);
  }

  for (let index = 0; index < worksheet.ShapeImages.length; index++) {
    const shapeImage = worksheet.ShapeImages[index];
    const globalPictureNumber = pictureFileOffset + worksheet.Pictures.length + index + 1;
    const path = `xl/media/image${globalPictureNumber}.${shapeImage.Extension ?? ""}`;
    archive.addBinary(path, shapeImage.ImageData ?? new Uint8Array(0));
  }
}

function buildChartAnchor(chart: ChartModel, rId: string, chartId: number): XmlElement {
  const anchor = XmlElement.create("xdr:twoCellAnchor").attr("editAs", "oneCell");
  anchor.child(
    buildFromElement(
      chart.UpperLeftColumn,
      chart.UpperLeftColumnOffset,
      chart.UpperLeftRow,
      chart.UpperLeftRowOffset,
    ),
  );
  anchor.child(
    buildToElement(
      chart.LowerRightColumn,
      chart.LowerRightColumnOffset,
      chart.LowerRightRow,
      chart.LowerRightRowOffset,
    ),
  );

  if (chart.IsChartEx && chart.RawGraphicFrameXml !== null && chart.RawGraphicFrameXml.length > 0) {
    let updatedXml = normalizePreservedDrawingMarkup(chart.RawGraphicFrameXml);
    if (chart.OriginalRId !== null && chart.OriginalRId.length > 0) {
      updatedXml = updatedXml.split(`"${chart.OriginalRId}"`).join(`"${rId}"`);
    }

    const rawElement = tryParseMarkup(updatedXml);
    if (rawElement !== null) {
      anchor.child(rawElement);
    }
  } else {
    anchor.child(buildGraphicFrameElement(chart.Name, rId, chartId));
  }

  anchor.child(XmlElement.create("xdr:clientData"));
  return anchor;
}

function buildGraphicFrameElement(displayName: string, rId: string, frameId: number): XmlElement {
  const name = displayName.length === 0 ? `Chart ${frameId}` : displayName;

  const graphicFrame = XmlElement.create("xdr:graphicFrame").attr("macro", "");

  graphicFrame.child(
    XmlElement.create("xdr:nvGraphicFramePr")
      .child(XmlElement.create("xdr:cNvPr").attr("id", frameId).attr("name", name))
      .child(XmlElement.create("xdr:cNvGraphicFramePr")),
  );

  graphicFrame.child(
    XmlElement.create("xdr:xfrm")
      .child(XmlElement.create("a:off").attr("x", "0").attr("y", "0"))
      .child(XmlElement.create("a:ext").attr("cx", "0").attr("cy", "0")),
  );

  graphicFrame.child(
    XmlElement.create("a:graphic").child(
      XmlElement.create("a:graphicData")
        .attr("uri", CHART_GRAPHIC_DATA_URI)
        .child(
          XmlElement.create("c:chart")
            .attr("xmlns:c", CHART_NAMESPACE)
            .attr("xmlns:r", RELATIONSHIP_NS)
            .attr("r:id", rId),
        ),
    ),
  );

  return graphicFrame;
}

// --- Preserved drawing markup normalization ---

interface DomAttributeValue {
  name: string;
  value: string;
}

interface DomNodeValue {
  nodeName: string;
  nodeType: number;
  textContent: string | null;
  nodeValue: string | null;
  childNodes: ArrayLike<DomNodeValue>;
  attributes?: ArrayLike<DomAttributeValue>;
}

function domNode(node: XmlNode): DomNodeValue {
  return node as unknown as DomNodeValue;
}

function domElement(node: XmlNode): Element {
  return node as unknown as Element;
}

function serializeNode(node: XmlNode): string {
  const target = node as unknown as { toString(): string };
  return target.toString();
}

function normalizePreservedDrawingMarkup(rawXml: string): string {
  const root = tryParseElement(rawXml);
  if (root === null) {
    return rawXml;
  }

  normalizeShapeTransformElements(root);
  normalizeCreationIdAttributes(root);
  return serializeNode(root);
}

function normalizeShapeTransformElements(root: XmlNode): void {
  const transformOwners = collectDescendantElements(root, "spPr");
  for (const owner of transformOwners) {
    const ownerElement = domElement(owner);
    if (ownerElement.namespaceURI !== XDR_NAMESPACE) {
      continue;
    }

    let transform: Element | null = null;
    const children = ownerElement.childNodes;
    for (let index = 0; index < children.length; index++) {
      const child = children.item(index);
      if (child === null || child.nodeType !== 1) {
        continue;
      }
      const childElement = child as unknown as Element;
      if (childElement.localName === "xfrm" && childElement.namespaceURI === XDR_NAMESPACE) {
        transform = childElement;
        break;
      }
    }

    if (transform === null) {
      continue;
    }

    const renamed = renameDomElement(transform, DRAWING_MAIN_NAMESPACE, "xfrm");
    if (renamed !== null) {
      normalizeDescendantNamespaces(renamed);
    }
  }
}

function normalizeDescendantNamespaces(element: Element): void {
  let current: Element = element;
  if (current.namespaceURI === XDR_NAMESPACE) {
    const renamed = renameDomElement(current, DRAWING_MAIN_NAMESPACE, current.localName ?? localNameOf(current.nodeName));
    if (renamed !== null) {
      current = renamed;
    }
  }

  const children = Array.from(current.childNodes);
  for (const child of children) {
    if (child.nodeType === 1) {
      normalizeDescendantNamespaces(child as unknown as Element);
    }
  }
}

function normalizeCreationIdAttributes(root: XmlNode): void {
  const elements = collectDescendantElementsIncludingSelf(root);
  for (const element of elements) {
    const node = domElement(element);
    if (node.localName !== "creationId") {
      continue;
    }

    if (node.namespaceURI !== OFFICE_DRAWING_2014_NS) {
      continue;
    }

    const idValue = node.getAttribute("id");
    if (idValue === null || idValue.trim().length === 0) {
      continue;
    }

    const normalizedGuid = tryNormalizeGuid(idValue);
    if (normalizedGuid === null) {
      continue;
    }

    node.setAttribute("id", `{${normalizedGuid}}`);
  }
}

function renameDomElement(element: Element, namespaceUri: string, localName: string): Element | null {
  const ownerDocument = element.ownerDocument;
  if (ownerDocument === null) {
    return null;
  }

  const rootNode = ownerDocument.documentElement as Element | null;
  let prefix = rootNode !== null ? rootNode.lookupPrefix(namespaceUri) : null;
  if (prefix === null || prefix.length === 0) {
    prefix = preferredNamespacePrefix(namespaceUri);
  }

  const qualifiedName = prefix !== null && prefix.length > 0 ? `${prefix}:${localName}` : localName;
  const replacement = ownerDocument.createElementNS(namespaceUri, qualifiedName);

  const attributes = element.attributes;
  for (let index = 0; index < attributes.length; index++) {
    const attributeNode = attributes.item(index);
    if (attributeNode === null) {
      continue;
    }
    replacement.setAttributeNS(attributeNode.namespaceURI, attributeNode.name, attributeNode.value);
  }

  while (element.firstChild !== null) {
    replacement.appendChild(element.firstChild);
  }

  const parentNode = element.parentNode;
  if (parentNode !== null) {
    parentNode.replaceChild(replacement, element);
  }

  return replacement;
}

function localNameOf(nodeName: string): string {
  const separator = nodeName.indexOf(":");
  return separator >= 0 ? nodeName.substring(separator + 1) : nodeName;
}

/** Returns the conventional OOXML prefix for a namespace when none is in scope. */
function preferredNamespacePrefix(namespaceUri: string): string | null {
  switch (namespaceUri) {
    case DRAWING_MAIN_NAMESPACE:
      return "a";
    case XDR_NAMESPACE:
      return "xdr";
    case RELATIONSHIP_NS:
      return "r";
    default:
      return null;
  }
}

function tryNormalizeGuid(value: string): string | null {
  let normalized = value.trim();
  if (normalized.length === 0) {
    return null;
  }

  if (
    (normalized.startsWith("{") && normalized.endsWith("}")) ||
    (normalized.startsWith("(") && normalized.endsWith(")"))
  ) {
    normalized = normalized.substring(1, normalized.length - 1);
  }

  const hyphenated = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;
  const plain = /^[0-9a-fA-F]{32}$/;
  if (!hyphenated.test(normalized) && !plain.test(normalized)) {
    return null;
  }

  const hex = normalized.replace(/-/g, "").toUpperCase();
  return (
    hex.substring(0, 4) +
    "-" +
    hex.substring(4, 6) +
    "-" +
    hex.substring(6, 8) +
    "-" +
    hex.substring(8, 10) +
    "-" +
    hex.substring(10, 16)
  );
}

// --- DOM helpers ---

function tryParseElement(markup: string): XmlNode | null {
  try {
    const document = parseXml(markup);
    return documentRoot(document);
  } catch {
    return null;
  }
}

function tryParseMarkup(markup: string): XmlElement | null {
  const root = tryParseElement(markup);
  if (root === null) {
    return null;
  }

  return convertDomToXmlElement(root);
}

function convertDomToXmlElement(node: XmlNode): XmlElement {
  const source = domNode(node);
  const element = XmlElement.create(source.nodeName);

  const attributes = source.attributes;
  if (attributes !== undefined) {
    for (let index = 0; index < attributes.length; index++) {
      const attributeNode = attributes[index];
      element.attr(attributeNode.name, attributeNode.value);
    }
  }

  const children = source.childNodes;
  for (let index = 0; index < children.length; index++) {
    const child = children[index];
    if (child.nodeType === 1) {
      element.child(convertDomToXmlElement(child as unknown as XmlNode));
    } else if (child.nodeType === 3 || child.nodeType === 4) {
      element.child(child.nodeValue ?? child.textContent ?? "");
    }
  }

  return element;
}

function collectDescendantElements(root: XmlNode, name: string): XmlNode[] {
  const result: XmlNode[] = [];
  const visit = (current: XmlNode): void => {
    const node = domNode(current);
    for (let index = 0; index < node.childNodes.length; index++) {
      const child = node.childNodes[index];
      if (child.nodeType !== 1) {
        continue;
      }
      if (localNameOf(child.nodeName) === name) {
        result.push(child as unknown as XmlNode);
      }
      visit(child as unknown as XmlNode);
    }
  };

  visit(root);
  return result;
}

function collectDescendantElementsIncludingSelf(root: XmlNode): XmlNode[] {
  const result: XmlNode[] = [];
  const visit = (current: XmlNode): void => {
    result.push(current);
    const node = domNode(current);
    for (let index = 0; index < node.childNodes.length; index++) {
      const child = node.childNodes[index];
      if (child.nodeType === 1) {
        visit(child as unknown as XmlNode);
      }
    }
  };

  visit(root);
  return result;
}

function attributeValue(node: XmlNode, name: string): string | null {
  const element = domElement(node);
  const value = element.getAttribute(name);
  return value === null ? null : value;
}

function tryParseInt(value: string): number | null {
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
