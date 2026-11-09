import { ChartCompanionFile, ChartModel } from "./Core/ChartModel";
import { ChartType } from "./ChartType";
import { DrawingRelationshipModel } from "./Core/DrawingRelationshipModel";
import { PictureModel } from "./Core/PictureModel";
import { ShapeImageModel } from "./Core/ShapeImageModel";
import { ShapeModel } from "./Core/ShapeModel";
import { SmartArtModel } from "./Core/SmartArtModel";
import { WorksheetModel } from "./Core/WorksheetModel";
import { DiagnosticSeverity } from "./enums";
import { LoadDiagnostics } from "./LoadDiagnostics";
import { LoadIssue } from "./LoadIssue";
import { LoadOptions } from "./LoadOptions";
import { WarningInfo } from "./WarningInfo";
import { ZipReader } from "./Packaging/Archive";
import {
  CHART_EX_RELATIONSHIP_TYPE,
  CHART_RELATIONSHIP_TYPE,
  CHART_STYLE_RELATIONSHIP_TYPE,
  CHART_COLOR_STYLE_RELATIONSHIP_TYPE,
  CHART_USER_SHAPES_RELATIONSHIP_TYPE,
  DRAWING_RELATIONSHIP_TYPE,
  IMAGE_RELATIONSHIP_TYPE,
} from "./XlsxWorkbookPictures";
import {
  attribute,
  descendants,
  documentRoot,
  elementChildren,
  firstElement,
  parseXml,
  textOf,
  XmlNode,
} from "./Xml/XmlReader";

const RELATIONSHIP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DIAGRAM_GRAPHIC_DATA_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const DIAGRAM_DRAWING_REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
const DSP_NAMESPACE = "http://schemas.microsoft.com/office/drawing/2008/diagram";
const CHART_EX_GRAPHIC_DATA_URI = "http://schemas.microsoft.com/office/drawing/2014/chartex";

const EMU_PER_COLUMN = 609600;
const EMU_PER_ROW = 190500;

interface ChartTarget {
  Uri: string;
  IsChartEx: boolean;
}

export function loadPictures(
  worksheetModel: WorksheetModel,
  worksheetRoot: XmlNode | null,
  archive: ZipReader,
  worksheetUri: string,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  worksheetModel.Pictures.length = 0;
  worksheetModel.ShapeImages.length = 0;
  worksheetModel.Shapes.length = 0;
  worksheetModel.Charts.length = 0;
  worksheetModel.PreservedCharts.length = 0;
  worksheetModel.PreservedDrawingRelationships.length = 0;

  const drawingUri = findDrawingUri(archive, worksheetUri);
  if (drawingUri.length === 0) {
    return;
  }

  if (!archive.hasEntry(drawingUri)) {
    addIssue(
      diagnostics,
      options,
      new IssueBuilder(
        "PIC-R001",
        DiagnosticSeverity.Recoverable,
        `Drawing part '${drawingUri}' was referenced but not found; pictures were skipped.`,
        true,
        false,
      ).withSheet(sheetName),
    );
    return;
  }

  const drawingDocument = parseXml(archive.readText(drawingUri));
  const drawingRoot = documentRoot(drawingDocument);
  if (drawingRoot === null) {
    return;
  }

  const imageTargets = loadDrawingImageTargets(archive, drawingUri);
  const chartTargets = loadDrawingChartTargets(archive, drawingUri);

  loadTwoCellAnchorPictures(worksheetModel, drawingRoot, imageTargets, archive, diagnostics, options, sheetName);
  loadOneCellAnchorPictures(worksheetModel, drawingRoot, imageTargets, archive, diagnostics, options, sheetName);
  loadShapeImages(worksheetModel, imageTargets, archive);
  loadTwoCellAnchorShapes(worksheetModel, drawingRoot);
  loadOneCellAnchorShapes(worksheetModel, drawingRoot);
  loadPreservedDrawingRelationships(worksheetModel, archive, drawingUri);
  loadTwoCellAnchorCharts(worksheetModel, drawingRoot, chartTargets, archive, diagnostics, options, sheetName);
  loadOneCellAnchorCharts(worksheetModel, drawingRoot, chartTargets, archive, diagnostics, options, sheetName);
  loadPreservedRawShapeCharts(worksheetModel, chartTargets, archive, diagnostics, options, sheetName);
  loadSmartArt(worksheetModel, drawingRoot, archive, drawingUri);
}

// --- Issue helper ---

class IssueBuilder extends LoadIssue {
  public constructor(code: string, severity: DiagnosticSeverity, message: string, repairApplied: boolean, dataLossRisk: boolean) {
    super(code, severity, message, repairApplied, dataLossRisk);
  }

  public withSheet(sheetName: string): LoadIssue {
    this.SheetName = sheetName;
    return this;
  }
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

// --- SmartArt loading ---

function loadSmartArt(
  worksheetModel: WorksheetModel,
  drawingRoot: XmlNode,
  archive: ZipReader,
  drawingUri: string,
): void {
  const drawingTargets = loadDrawingTargetsByType(archive, drawingUri, DIAGRAM_DRAWING_REL_TYPE);
  if (drawingTargets.size === 0) {
    return;
  }

  const allRels = loadAllDrawingRelTargets(archive, drawingUri);

  for (const anchor of elementChildren(drawingRoot, "twoCellAnchor")) {
    const fromEl = firstElement(anchor, "from");
    const toEl = firstElement(anchor, "to");
    const graphicFrame = firstElement(anchor, "graphicFrame");
    if (fromEl === null || toEl === null || graphicFrame === null || !isDiagramGraphicFrame(graphicFrame)) {
      continue;
    }

    const drawingPartUri = resolveSmartArtDrawingUri(graphicFrame, drawingTargets, allRels, archive);
    if (drawingPartUri.length === 0 || !archive.hasEntry(drawingPartUri)) {
      continue;
    }

    const rawDrawingXml = archive.readText(drawingPartUri);

    const model = new SmartArtModel();
    model.Name = getChartName(graphicFrame);
    model.RawDrawingXml = rawDrawingXml;
    model.UpperLeftColumn = parseAnchorInt(firstElement(fromEl, "col"));
    model.UpperLeftColumnOffset = parseAnchorLong(firstElement(fromEl, "colOff"));
    model.UpperLeftRow = parseAnchorInt(firstElement(fromEl, "row"));
    model.UpperLeftRowOffset = parseAnchorLong(firstElement(fromEl, "rowOff"));
    model.LowerRightColumn = parseAnchorInt(firstElement(toEl, "col"));
    model.LowerRightColumnOffset = parseAnchorLong(firstElement(toEl, "colOff"));
    model.LowerRightRow = parseAnchorInt(firstElement(toEl, "row"));
    model.LowerRightRowOffset = parseAnchorLong(firstElement(toEl, "rowOff"));

    worksheetModel.SmartArts.push(model);
  }
}

function isDiagramGraphicFrame(graphicFrame: XmlNode): boolean {
  const graphic = firstElement(graphicFrame, "graphic");
  const graphicData = graphic === null ? null : firstElement(graphic, "graphicData");
  return graphicData !== null && attribute(graphicData, "uri") === DIAGRAM_GRAPHIC_DATA_URI;
}

function resolveSmartArtDrawingUri(
  graphicFrame: XmlNode,
  drawingTargets: Map<string, string>,
  allRels: Map<string, string>,
  archive: ZipReader,
): string {
  const graphic = firstElement(graphicFrame, "graphic");
  const graphicData = graphic === null ? null : firstElement(graphic, "graphicData");
  if (graphicData === null) {
    return "";
  }

  let dataModelRelId: string | null = null;
  const relIdsElements = descendants(graphicData, "relIds");
  for (const relIds of relIdsElements) {
    dataModelRelId = attribute(relIds, "dm");
    break;
  }

  if (dataModelRelId !== null && dataModelRelId.length > 0) {
    const dataUri = getIgnoreCase(allRels, dataModelRelId);
    if (dataUri !== undefined && archive.hasEntry(dataUri)) {
      const dataDocument = parseXml(archive.readText(dataUri));
      const dataRoot = documentRoot(dataDocument);
      if (dataRoot !== null) {
        let drawingRelId: string | null = null;
        const extensionElements = descendants(dataRoot, "dataModelExt");
        for (const extensionElement of extensionElements) {
          if (isDspNamespace(extensionElement)) {
            drawingRelId = attribute(extensionElement, "relId");
            break;
          }
        }

        if (drawingRelId !== null && drawingRelId.length > 0) {
          const uri = getIgnoreCase(drawingTargets, drawingRelId);
          if (uri !== undefined) {
            return uri;
          }
        }
      }
    }
  }

  for (const value of drawingTargets.values()) {
    return value;
  }

  return "";
}

function isDspNamespace(node: XmlNode): boolean {
  return (node as unknown as { namespaceURI?: string | null }).namespaceURI === DSP_NAMESPACE;
}

function loadAllDrawingRelTargets(archive: ZipReader, drawingUri: string): Map<string, string> {
  const result = new Map<string, string>();
  const relsUri = getDrawingRelsUri(drawingUri);
  if (!archive.hasEntry(relsUri)) {
    return result;
  }

  const document = parseXml(archive.readText(relsUri));
  const root = documentRoot(document);
  if (root === null) {
    return result;
  }

  for (const relationship of elementChildren(root, "Relationship")) {
    const id = attribute(relationship, "Id");
    const target = attribute(relationship, "Target");
    if (id !== null && id.length > 0 && target !== null && target.length > 0) {
      result.set(id, resolvePartUri(drawingUri, target));
    }
  }

  return result;
}

// --- Picture anchors ---

function loadTwoCellAnchorPictures(
  worksheetModel: WorksheetModel,
  drawingRoot: XmlNode,
  imageTargets: Map<string, string>,
  archive: ZipReader,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  for (const anchor of elementChildren(drawingRoot, "twoCellAnchor")) {
    const pic = firstElement(anchor, "pic");
    if (pic === null) {
      continue;
    }

    const fromEl = firstElement(anchor, "from");
    const toEl = firstElement(anchor, "to");
    if (fromEl === null || toEl === null) {
      continue;
    }

    const model = parsePicture(pic, fromEl, toEl, imageTargets, archive, diagnostics, options, sheetName);
    if (model !== null) {
      worksheetModel.Pictures.push(model);
    }
  }
}

function loadOneCellAnchorPictures(
  worksheetModel: WorksheetModel,
  drawingRoot: XmlNode,
  imageTargets: Map<string, string>,
  archive: ZipReader,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  for (const anchor of elementChildren(drawingRoot, "oneCellAnchor")) {
    const pic = firstElement(anchor, "pic");
    if (pic === null) {
      continue;
    }

    const fromEl = firstElement(anchor, "from");
    if (fromEl === null) {
      continue;
    }

    const extEl = firstElement(anchor, "ext");
    const model = parseOneCellPicture(pic, fromEl, extEl, imageTargets, archive, diagnostics, options, sheetName);
    if (model !== null) {
      worksheetModel.Pictures.push(model);
    }
  }
}

function loadShapeImages(worksheetModel: WorksheetModel, imageTargets: Map<string, string>, archive: ZipReader): void {
  worksheetModel.ShapeImages.length = 0;

  const claimedRIds = new Set<string>();
  for (const picture of worksheetModel.Pictures) {
    if (picture.OriginalRId !== null && picture.OriginalRId.length > 0) {
      claimedRIds.add(picture.OriginalRId);
    }
  }

  for (const [rId, mediaUri] of imageTargets) {
    if (claimedRIds.has(rId)) {
      continue;
    }

    if (!archive.hasEntry(mediaUri)) {
      continue;
    }

    const imageData = new Uint8Array(archive.readBinary(mediaUri));
    if (imageData.length === 0) {
      continue;
    }

    const model = new ShapeImageModel();
    model.OriginalRId = rId;
    model.Extension = detectExtension(imageData);
    model.ImageData = imageData;
    worksheetModel.ShapeImages.push(model);
  }
}

function loadTwoCellAnchorShapes(worksheetModel: WorksheetModel, drawingRoot: XmlNode): void {
  for (const anchor of elementChildren(drawingRoot, "twoCellAnchor")) {
    const fromEl = firstElement(anchor, "from");
    const toEl = firstElement(anchor, "to");
    if (fromEl === null || toEl === null) {
      continue;
    }

    const sp = firstElement(anchor, "sp");
    if (sp !== null) {
      worksheetModel.Shapes.push(parseShape(sp, fromEl, toEl));
      continue;
    }

    const grpSp = firstElement(anchor, "grpSp");
    if (grpSp !== null) {
      worksheetModel.Shapes.push(parseGroupShape(grpSp, fromEl, toEl));
      continue;
    }

    const cxnSp = firstElement(anchor, "cxnSp");
    if (cxnSp !== null) {
      worksheetModel.Shapes.push(parseConnector(cxnSp, fromEl, toEl));
    }
  }
}

function loadOneCellAnchorShapes(worksheetModel: WorksheetModel, drawingRoot: XmlNode): void {
  for (const anchor of elementChildren(drawingRoot, "oneCellAnchor")) {
    const fromEl = firstElement(anchor, "from");
    if (fromEl === null) {
      continue;
    }

    const extEl = firstElement(anchor, "ext");

    const sp = firstElement(anchor, "sp");
    if (sp !== null) {
      worksheetModel.Shapes.push(parseOneCellShape(sp, fromEl, extEl));
      continue;
    }

    const grpSp = firstElement(anchor, "grpSp");
    if (grpSp !== null) {
      worksheetModel.Shapes.push(parseOneCellGroupShape(grpSp, fromEl, extEl));
      continue;
    }

    const cxnSp = firstElement(anchor, "cxnSp");
    if (cxnSp !== null) {
      worksheetModel.Shapes.push(parseOneCellConnector(cxnSp, fromEl, extEl));
    }
  }
}

function parseShape(sp: XmlNode, fromEl: XmlNode, toEl: XmlNode): ShapeModel {
  const model = new ShapeModel();
  model.Name = getShapeName(sp);
  model.DrawingObjectId = getDrawingObjectId(sp);
  model.UpperLeftColumn = parseAnchorInt(firstElement(fromEl, "col"));
  model.UpperLeftColumnOffset = parseAnchorLong(firstElement(fromEl, "colOff"));
  model.UpperLeftRow = parseAnchorInt(firstElement(fromEl, "row"));
  model.UpperLeftRowOffset = parseAnchorLong(firstElement(fromEl, "rowOff"));
  model.LowerRightColumn = parseAnchorInt(firstElement(toEl, "col"));
  model.LowerRightColumnOffset = parseAnchorLong(firstElement(toEl, "colOff"));
  model.LowerRightRow = parseAnchorInt(firstElement(toEl, "row"));
  model.LowerRightRowOffset = parseAnchorLong(firstElement(toEl, "rowOff"));
  model.GeometryType = getGeometryType(sp);
  model.RawStyleXml = getRawXml(firstElement(sp, "style"));
  model.RawTxBodyXml = getRawXml(firstElement(sp, "txBody"));
  model.RawElementXml = serializeNode(sp);
  loadShapeSpPrExtents(sp, model);
  return model;
}

function parseOneCellShape(sp: XmlNode, fromEl: XmlNode, extEl: XmlNode | null): ShapeModel {
  const fromCol = parseAnchorInt(firstElement(fromEl, "col"));
  const fromRow = parseAnchorInt(firstElement(fromEl, "row"));
  const cx = parseExtentCoordinate(extEl, "cx");
  const cy = parseExtentCoordinate(extEl, "cy");
  const colSpan = cx > 0 ? Math.trunc(cx / EMU_PER_COLUMN) + 1 : 1;
  const rowSpan = cy > 0 ? Math.trunc(cy / EMU_PER_ROW) + 1 : 1;

  const model = new ShapeModel();
  model.Name = getShapeName(sp);
  model.DrawingObjectId = getDrawingObjectId(sp);
  model.UpperLeftColumn = fromCol;
  model.UpperLeftColumnOffset = parseAnchorLong(firstElement(fromEl, "colOff"));
  model.UpperLeftRow = fromRow;
  model.UpperLeftRowOffset = parseAnchorLong(firstElement(fromEl, "rowOff"));
  model.LowerRightColumn = fromCol + colSpan;
  model.LowerRightRow = fromRow + rowSpan;
  model.ExtentCx = cx;
  model.ExtentCy = cy;
  model.GeometryType = getGeometryType(sp);
  model.RawStyleXml = getRawXml(firstElement(sp, "style"));
  model.RawTxBodyXml = getRawXml(firstElement(sp, "txBody"));
  model.RawElementXml = serializeNode(sp);
  return model;
}

function parseConnector(cxnSp: XmlNode, fromEl: XmlNode, toEl: XmlNode): ShapeModel {
  const model = new ShapeModel();
  model.Name = getConnectorName(cxnSp);
  model.DrawingObjectId = getDrawingObjectId(cxnSp);
  model.UpperLeftColumn = parseAnchorInt(firstElement(fromEl, "col"));
  model.UpperLeftColumnOffset = parseAnchorLong(firstElement(fromEl, "colOff"));
  model.UpperLeftRow = parseAnchorInt(firstElement(fromEl, "row"));
  model.UpperLeftRowOffset = parseAnchorLong(firstElement(fromEl, "rowOff"));
  model.LowerRightColumn = parseAnchorInt(firstElement(toEl, "col"));
  model.LowerRightColumnOffset = parseAnchorLong(firstElement(toEl, "colOff"));
  model.LowerRightRow = parseAnchorInt(firstElement(toEl, "row"));
  model.LowerRightRowOffset = parseAnchorLong(firstElement(toEl, "rowOff"));
  model.GeometryType = getGeometryType(cxnSp);
  model.RawElementXml = serializeNode(cxnSp);
  loadConnectorConnections(cxnSp, model);
  return model;
}

function parseOneCellConnector(cxnSp: XmlNode, fromEl: XmlNode, extEl: XmlNode | null): ShapeModel {
  const fromCol = parseAnchorInt(firstElement(fromEl, "col"));
  const fromRow = parseAnchorInt(firstElement(fromEl, "row"));
  const cx = parseExtentCoordinate(extEl, "cx");
  const cy = parseExtentCoordinate(extEl, "cy");
  const colSpan = cx > 0 ? Math.trunc(cx / EMU_PER_COLUMN) + 1 : 1;
  const rowSpan = cy > 0 ? Math.trunc(cy / EMU_PER_ROW) + 1 : 1;

  const model = new ShapeModel();
  model.Name = getConnectorName(cxnSp);
  model.DrawingObjectId = getDrawingObjectId(cxnSp);
  model.UpperLeftColumn = fromCol;
  model.UpperLeftColumnOffset = parseAnchorLong(firstElement(fromEl, "colOff"));
  model.UpperLeftRow = fromRow;
  model.UpperLeftRowOffset = parseAnchorLong(firstElement(fromEl, "rowOff"));
  model.LowerRightColumn = fromCol + colSpan;
  model.LowerRightRow = fromRow + rowSpan;
  model.ExtentCx = cx;
  model.ExtentCy = cy;
  model.GeometryType = getGeometryType(cxnSp);
  model.RawElementXml = serializeNode(cxnSp);
  loadConnectorConnections(cxnSp, model);
  return model;
}

function parseGroupShape(grpSp: XmlNode, fromEl: XmlNode, toEl: XmlNode): ShapeModel {
  const model = new ShapeModel();
  model.Name = getGroupShapeName(grpSp);
  model.DrawingObjectId = getDrawingObjectId(grpSp);
  model.UpperLeftColumn = parseAnchorInt(firstElement(fromEl, "col"));
  model.UpperLeftColumnOffset = parseAnchorLong(firstElement(fromEl, "colOff"));
  model.UpperLeftRow = parseAnchorInt(firstElement(fromEl, "row"));
  model.UpperLeftRowOffset = parseAnchorLong(firstElement(fromEl, "rowOff"));
  model.LowerRightColumn = parseAnchorInt(firstElement(toEl, "col"));
  model.LowerRightColumnOffset = parseAnchorLong(firstElement(toEl, "colOff"));
  model.LowerRightRow = parseAnchorInt(firstElement(toEl, "row"));
  model.LowerRightRowOffset = parseAnchorLong(firstElement(toEl, "rowOff"));
  model.RawElementXml = serializeNode(grpSp);
  return model;
}

function parseOneCellGroupShape(grpSp: XmlNode, fromEl: XmlNode, extEl: XmlNode | null): ShapeModel {
  const fromCol = parseAnchorInt(firstElement(fromEl, "col"));
  const fromRow = parseAnchorInt(firstElement(fromEl, "row"));
  const cx = parseExtentCoordinate(extEl, "cx");
  const cy = parseExtentCoordinate(extEl, "cy");
  const colSpan = cx > 0 ? Math.trunc(cx / EMU_PER_COLUMN) + 1 : 1;
  const rowSpan = cy > 0 ? Math.trunc(cy / EMU_PER_ROW) + 1 : 1;

  const model = new ShapeModel();
  model.Name = getGroupShapeName(grpSp);
  model.DrawingObjectId = getDrawingObjectId(grpSp);
  model.UpperLeftColumn = fromCol;
  model.UpperLeftColumnOffset = parseAnchorLong(firstElement(fromEl, "colOff"));
  model.UpperLeftRow = fromRow;
  model.UpperLeftRowOffset = parseAnchorLong(firstElement(fromEl, "rowOff"));
  model.LowerRightColumn = fromCol + colSpan;
  model.LowerRightRow = fromRow + rowSpan;
  model.ExtentCx = cx;
  model.ExtentCy = cy;
  model.RawElementXml = serializeNode(grpSp);
  return model;
}

function loadConnectorConnections(cxnSp: XmlNode, model: ShapeModel): void {
  const nvCxnSpPr = firstElement(cxnSp, "nvCxnSpPr");
  if (nvCxnSpPr === null) {
    return;
  }

  const cNvCxnSpPr = firstElement(nvCxnSpPr, "cNvCxnSpPr");
  if (cNvCxnSpPr === null) {
    return;
  }

  loadConnectorConnection(firstElement(cNvCxnSpPr, "stCxn"), true, model);
  loadConnectorConnection(firstElement(cNvCxnSpPr, "endCxn"), false, model);
}

function loadConnectorConnection(connection: XmlNode | null, isStart: boolean, model: ShapeModel): void {
  if (connection === null) {
    return;
  }

  const shapeIdValue = attribute(connection, "id");
  const shapeId = shapeIdValue === null ? -1 : (tryParseInt(shapeIdValue) ?? -1);
  const siteValue = attribute(connection, "idx");
  const site = siteValue === null ? -1 : (tryParseInt(siteValue) ?? -1);

  if (isStart) {
    model.StartConnectionShapeId = shapeId;
    model.StartConnectionSite = site;
  } else {
    model.EndConnectionShapeId = shapeId;
    model.EndConnectionSite = site;
  }
}

function getGroupShapeName(grpSp: XmlNode): string {
  const nvGrpSpPr = firstElement(grpSp, "nvGrpSpPr");
  if (nvGrpSpPr === null) {
    return "";
  }

  const cNvPr = firstElement(nvGrpSpPr, "cNvPr");
  return cNvPr === null ? "" : (attribute(cNvPr, "name") ?? "");
}

function getConnectorName(cxnSp: XmlNode): string {
  const nvCxnSpPr = firstElement(cxnSp, "nvCxnSpPr");
  if (nvCxnSpPr === null) {
    return "";
  }

  const cNvPr = firstElement(nvCxnSpPr, "cNvPr");
  return cNvPr === null ? "" : (attribute(cNvPr, "name") ?? "");
}

function getShapeName(sp: XmlNode): string {
  const nvSpPr = firstElement(sp, "nvSpPr");
  if (nvSpPr === null) {
    return "";
  }

  const cNvPr = firstElement(nvSpPr, "cNvPr");
  return cNvPr === null ? "" : (attribute(cNvPr, "name") ?? "");
}

function getDrawingObjectId(shapeElement: XmlNode): number {
  const cNvPr = getNonVisualDrawingProperties(shapeElement);
  if (cNvPr === null) {
    return -1;
  }

  const idValue = attribute(cNvPr, "id");
  if (idValue === null) {
    return -1;
  }

  return tryParseInt(idValue) ?? -1;
}

function getNonVisualDrawingProperties(shapeElement: XmlNode | null): XmlNode | null {
  if (shapeElement === null) {
    return null;
  }

  const nvSpPr = firstElement(shapeElement, "nvSpPr");
  if (nvSpPr !== null) {
    return firstElement(nvSpPr, "cNvPr");
  }

  const nvCxnSpPr = firstElement(shapeElement, "nvCxnSpPr");
  if (nvCxnSpPr !== null) {
    return firstElement(nvCxnSpPr, "cNvPr");
  }

  const nvGrpSpPr = firstElement(shapeElement, "nvGrpSpPr");
  if (nvGrpSpPr !== null) {
    return firstElement(nvGrpSpPr, "cNvPr");
  }

  return null;
}

function getGeometryType(sp: XmlNode): string {
  const spPr = firstElement(sp, "spPr");
  if (spPr === null) {
    return "rect";
  }

  const prstGeom = firstElement(spPr, "prstGeom");
  const prst = prstGeom === null ? null : attributeOrNull(prstGeom, "prst");
  return prst === null ? "rect" : prst;
}

function getRawXml(element: XmlNode | null): string | null {
  return element === null ? null : serializeNode(element);
}

function loadShapeSpPrExtents(sp: XmlNode, model: ShapeModel): void {
  const spPr = firstElement(sp, "spPr");
  if (spPr === null) {
    return;
  }

  const xfrm = firstElement(spPr, "xfrm");
  if (xfrm === null) {
    return;
  }

  const ext = firstElement(xfrm, "ext");
  if (ext === null) {
    return;
  }

  model.ExtentCx = parseExtentCoordinate(ext, "cx");
  model.ExtentCy = parseExtentCoordinate(ext, "cy");
}

function parsePicture(
  pic: XmlNode,
  fromEl: XmlNode,
  toEl: XmlNode,
  imageTargets: Map<string, string>,
  archive: ZipReader,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): PictureModel | null {
  const rId = getBlipRId(pic);
  if (rId === null || rId.length === 0) {
    return null;
  }

  const name = getPictureName(pic);
  const imageData = loadImageData(rId, imageTargets, archive, diagnostics, options, sheetName, name);
  if (imageData === null) {
    return null;
  }

  const model = new PictureModel();
  model.Name = name;
  model.UpperLeftColumn = parseAnchorInt(firstElement(fromEl, "col"));
  model.UpperLeftColumnOffset = parseAnchorLong(firstElement(fromEl, "colOff"));
  model.UpperLeftRow = parseAnchorInt(firstElement(fromEl, "row"));
  model.UpperLeftRowOffset = parseAnchorLong(firstElement(fromEl, "rowOff"));
  model.LowerRightColumn = parseAnchorInt(firstElement(toEl, "col"));
  model.LowerRightColumnOffset = parseAnchorLong(firstElement(toEl, "colOff"));
  model.LowerRightRow = parseAnchorInt(firstElement(toEl, "row"));
  model.LowerRightRowOffset = parseAnchorLong(firstElement(toEl, "rowOff"));
  model.ImageData = imageData;
  model.ImageExtension = detectExtension(imageData);
  model.OriginalRId = rId;

  loadSpPrExtents(pic, model);
  return model;
}

function parseOneCellPicture(
  pic: XmlNode,
  fromEl: XmlNode,
  extEl: XmlNode | null,
  imageTargets: Map<string, string>,
  archive: ZipReader,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): PictureModel | null {
  const rId = getBlipRId(pic);
  if (rId === null || rId.length === 0) {
    return null;
  }

  const name = getPictureName(pic);
  const imageData = loadImageData(rId, imageTargets, archive, diagnostics, options, sheetName, name);
  if (imageData === null) {
    return null;
  }

  const fromCol = parseAnchorInt(firstElement(fromEl, "col"));
  const fromRow = parseAnchorInt(firstElement(fromEl, "row"));
  const cx = parseExtentCoordinate(extEl, "cx");
  const cy = parseExtentCoordinate(extEl, "cy");
  const colSpan = cx > 0 ? Math.trunc(cx / EMU_PER_COLUMN) + 1 : 1;
  const rowSpan = cy > 0 ? Math.trunc(cy / EMU_PER_ROW) + 1 : 1;

  const model = new PictureModel();
  model.Name = name;
  model.UpperLeftColumn = fromCol;
  model.UpperLeftColumnOffset = parseAnchorLong(firstElement(fromEl, "colOff"));
  model.UpperLeftRow = fromRow;
  model.UpperLeftRowOffset = parseAnchorLong(firstElement(fromEl, "rowOff"));
  model.LowerRightColumn = fromCol + colSpan;
  model.LowerRightRow = fromRow + rowSpan;
  model.ExtentCx = cx;
  model.ExtentCy = cy;
  model.ImageData = imageData;
  model.ImageExtension = detectExtension(imageData);
  model.OriginalRId = rId;
  return model;
}

function getBlipRId(pic: XmlNode): string | null {
  const blipFill = firstElement(pic, "blipFill");
  if (blipFill === null) {
    return null;
  }

  const blip = firstElement(blipFill, "blip");
  if (blip === null) {
    return null;
  }

  return attribute(blip, "embed");
}

function getPictureName(pic: XmlNode): string {
  const nvPicPr = firstElement(pic, "nvPicPr");
  if (nvPicPr === null) {
    return "";
  }

  const cNvPr = firstElement(nvPicPr, "cNvPr");
  return cNvPr === null ? "" : (attribute(cNvPr, "name") ?? "");
}

function loadImageData(
  rId: string,
  imageTargets: Map<string, string>,
  archive: ZipReader,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  pictureName: string,
): Uint8Array | null {
  const mediaUri = getIgnoreCase(imageTargets, rId);
  if (mediaUri === undefined) {
    addIssue(
      diagnostics,
      options,
      new IssueBuilder(
        "PIC-L001",
        DiagnosticSeverity.LossyRecoverable,
        `Image relationship '${rId}' for picture '${pictureName}' could not be resolved; picture was dropped.`,
        true,
        true,
      ).withSheet(sheetName),
    );
    return null;
  }

  if (!archive.hasEntry(mediaUri)) {
    addIssue(
      diagnostics,
      options,
      new IssueBuilder(
        "PIC-L001",
        DiagnosticSeverity.LossyRecoverable,
        `Image media file '${mediaUri}' was not found in the package; picture was dropped.`,
        true,
        true,
      ).withSheet(sheetName),
    );
    return null;
  }

  return new Uint8Array(archive.readBinary(mediaUri));
}

function loadSpPrExtents(pic: XmlNode, model: PictureModel): void {
  const spPr = firstElement(pic, "spPr");
  if (spPr === null) {
    return;
  }

  const xfrm = firstElement(spPr, "xfrm");
  if (xfrm === null) {
    return;
  }

  const ext = firstElement(xfrm, "ext");
  if (ext === null) {
    return;
  }

  model.ExtentCx = parseExtentCoordinate(ext, "cx");
  model.ExtentCy = parseExtentCoordinate(ext, "cy");
}

function loadDrawingImageTargets(archive: ZipReader, drawingUri: string): Map<string, string> {
  return loadDrawingTargetsByType(archive, drawingUri, IMAGE_RELATIONSHIP_TYPE);
}

function loadPreservedDrawingRelationships(
  worksheetModel: WorksheetModel,
  archive: ZipReader,
  drawingUri: string,
): void {
  worksheetModel.PreservedDrawingRelationships.length = 0;

  const referencedRelationshipIds = getReferencedDrawingRelationshipIds(worksheetModel.Shapes);
  if (referencedRelationshipIds.size === 0) {
    return;
  }

  const relsUri = getDrawingRelsUri(drawingUri);
  if (!archive.hasEntry(relsUri)) {
    return;
  }

  const document = parseXml(archive.readText(relsUri));
  const root = documentRoot(document);
  if (root === null) {
    return;
  }

  for (const relationship of elementChildren(root, "Relationship")) {
    const id = attribute(relationship, "Id");
    const type = attribute(relationship, "Type");
    const target = attribute(relationship, "Target");
    if (id === null || id.length === 0 || type === null || type.length === 0 || target === null || target.length === 0) {
      continue;
    }

    if (!referencedRelationshipIds.has(id)) {
      continue;
    }

    if (
      equalsIgnoreCase(type, IMAGE_RELATIONSHIP_TYPE) ||
      equalsIgnoreCase(type, CHART_RELATIONSHIP_TYPE) ||
      equalsIgnoreCase(type, CHART_EX_RELATIONSHIP_TYPE)
    ) {
      continue;
    }

    const model = new DrawingRelationshipModel();
    model.OriginalRId = id;
    model.RelationshipType = type;
    model.Target = target;
    model.TargetMode = attributeOrNull(relationship, "TargetMode");
    worksheetModel.PreservedDrawingRelationships.push(model);
  }
}

function getReferencedDrawingRelationshipIds(shapes: ShapeModel[]): Set<string> {
  const referenced = new Set<string>();
  for (const shape of shapes) {
    collectReferencedDrawingRelationshipIds(shape.RawElementXml, referenced);
  }

  return referenced;
}

function collectReferencedDrawingRelationshipIds(rawXml: string | null, referenced: Set<string>): void {
  if (rawXml === null || rawXml.length === 0) {
    return;
  }

  let root: XmlNode | null = null;
  try {
    root = documentRoot(parseXml(rawXml));
  } catch {
    return;
  }

  if (root === null) {
    return;
  }

  const elements = descendantsIncludingSelf(root);
  for (const element of elements) {
    const domElement = element as unknown as {
      attributes?: ArrayLike<{ namespaceURI: string | null; localName: string; value: string }>;
    };
    const attributes = domElement.attributes;
    if (attributes === undefined) {
      continue;
    }

    for (let index = 0; index < attributes.length; index++) {
      const attributeNode = attributes[index];
      if (attributeNode.namespaceURI !== RELATIONSHIP_NS) {
        continue;
      }

      if (
        attributeNode.localName !== "id" &&
        attributeNode.localName !== "embed" &&
        attributeNode.localName !== "link"
      ) {
        continue;
      }

      if (attributeNode.value.length > 0) {
        referenced.add(attributeNode.value);
      }
    }
  }
}

// --- Chart loading ---

function loadTwoCellAnchorCharts(
  worksheetModel: WorksheetModel,
  drawingRoot: XmlNode,
  chartTargets: Map<string, ChartTarget>,
  archive: ZipReader,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  for (const anchor of elementChildren(drawingRoot, "twoCellAnchor")) {
    const fromEl = firstElement(anchor, "from");
    const toEl = firstElement(anchor, "to");
    if (fromEl === null || toEl === null) {
      continue;
    }

    const graphicFrame = firstElement(anchor, "graphicFrame");
    if (graphicFrame !== null && isChartGraphicFrame(graphicFrame)) {
      const model = parseChart(graphicFrame, null, fromEl, toEl, null, chartTargets, archive, diagnostics, options, sheetName);
      if (model !== null) {
        worksheetModel.Charts.push(model);
      }
      continue;
    }

    const alternateContent = firstElement(anchor, "AlternateContent");
    if (alternateContent !== null) {
      const choice = firstElement(alternateContent, "Choice");
      const choiceGraphicFrame = choice === null ? null : firstElement(choice, "graphicFrame");
      if (choiceGraphicFrame !== null && isChartExGraphicFrame(choiceGraphicFrame)) {
        const model = parseChart(choiceGraphicFrame, alternateContent, fromEl, toEl, null, chartTargets, archive, diagnostics, options, sheetName);
        if (model !== null) {
          worksheetModel.Charts.push(model);
        }
      }
    }
  }
}

function loadOneCellAnchorCharts(
  worksheetModel: WorksheetModel,
  drawingRoot: XmlNode,
  chartTargets: Map<string, ChartTarget>,
  archive: ZipReader,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  for (const anchor of elementChildren(drawingRoot, "oneCellAnchor")) {
    const fromEl = firstElement(anchor, "from");
    const extEl = firstElement(anchor, "ext");
    if (fromEl === null) {
      continue;
    }

    const graphicFrame = firstElement(anchor, "graphicFrame");
    if (graphicFrame !== null && isChartGraphicFrame(graphicFrame)) {
      const model = parseChart(graphicFrame, null, fromEl, null, extEl, chartTargets, archive, diagnostics, options, sheetName);
      if (model !== null) {
        worksheetModel.Charts.push(model);
      }
      continue;
    }

    const alternateContent = firstElement(anchor, "AlternateContent");
    if (alternateContent !== null) {
      const choice = firstElement(alternateContent, "Choice");
      const choiceGraphicFrame = choice === null ? null : firstElement(choice, "graphicFrame");
      if (choiceGraphicFrame !== null && isChartExGraphicFrame(choiceGraphicFrame)) {
        const model = parseChart(choiceGraphicFrame, alternateContent, fromEl, null, extEl, chartTargets, archive, diagnostics, options, sheetName);
        if (model !== null) {
          worksheetModel.Charts.push(model);
        }
      }
    }
  }
}

function loadPreservedRawShapeCharts(
  worksheetModel: WorksheetModel,
  chartTargets: Map<string, ChartTarget>,
  archive: ZipReader,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  const claimedRelationshipIds = new Set<string>();
  for (const chart of worksheetModel.Charts) {
    if (chart.OriginalRId !== null && chart.OriginalRId.length > 0) {
      claimedRelationshipIds.add(chart.OriginalRId);
    }
  }

  const unclaimedRelationshipIds: string[] = [];
  for (const relationshipId of chartTargets.keys()) {
    if (!claimedRelationshipIds.has(relationshipId)) {
      unclaimedRelationshipIds.push(relationshipId);
    }
  }

  unclaimedRelationshipIds.sort();

  for (const relationshipId of unclaimedRelationshipIds) {
    const model = loadChartForRelationship(relationshipId, "", null, chartTargets, archive, diagnostics, options, sheetName);
    if (model !== null) {
      worksheetModel.PreservedCharts.push(model);
    }
  }
}

function isChartGraphicFrame(graphicFrame: XmlNode): boolean {
  const graphic = firstElement(graphicFrame, "graphic");
  const graphicData = graphic === null ? null : firstElement(graphic, "graphicData");
  if (graphicData === null) {
    return false;
  }

  const uri = attribute(graphicData, "uri");
  return equalsIgnoreCase(uri, "http://schemas.openxmlformats.org/drawingml/2006/chart");
}

function isChartExGraphicFrame(graphicFrame: XmlNode): boolean {
  const graphic = firstElement(graphicFrame, "graphic");
  const graphicData = graphic === null ? null : firstElement(graphic, "graphicData");
  if (graphicData === null) {
    return false;
  }

  const uri = attribute(graphicData, "uri");
  return equalsIgnoreCase(uri, CHART_EX_GRAPHIC_DATA_URI);
}

function parseChart(
  graphicFrame: XmlNode,
  rawContainerElement: XmlNode | null,
  fromEl: XmlNode,
  toEl: XmlNode | null,
  extEl: XmlNode | null,
  chartTargets: Map<string, ChartTarget>,
  archive: ZipReader,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): ChartModel | null {
  const rId = getChartRelationshipId(graphicFrame);
  if (rId === null || rId.length === 0) {
    return null;
  }

  const model = loadChartForRelationship(
    rId,
    getChartName(graphicFrame),
    rawContainerElement,
    chartTargets,
    archive,
    diagnostics,
    options,
    sheetName,
  );
  if (model === null) {
    return null;
  }

  if (toEl !== null) {
    model.UpperLeftColumn = parseAnchorInt(firstElement(fromEl, "col"));
    model.UpperLeftColumnOffset = parseAnchorLong(firstElement(fromEl, "colOff"));
    model.UpperLeftRow = parseAnchorInt(firstElement(fromEl, "row"));
    model.UpperLeftRowOffset = parseAnchorLong(firstElement(fromEl, "rowOff"));
    model.LowerRightColumn = parseAnchorInt(firstElement(toEl, "col"));
    model.LowerRightColumnOffset = parseAnchorLong(firstElement(toEl, "colOff"));
    model.LowerRightRow = parseAnchorInt(firstElement(toEl, "row"));
    model.LowerRightRowOffset = parseAnchorLong(firstElement(toEl, "rowOff"));
  } else if (extEl !== null) {
    const fromCol = parseAnchorInt(firstElement(fromEl, "col"));
    const fromRow = parseAnchorInt(firstElement(fromEl, "row"));
    const cx = parseExtentCoordinate(extEl, "cx");
    const cy = parseExtentCoordinate(extEl, "cy");
    model.UpperLeftColumn = fromCol;
    model.UpperLeftColumnOffset = parseAnchorLong(firstElement(fromEl, "colOff"));
    model.UpperLeftRow = fromRow;
    model.UpperLeftRowOffset = parseAnchorLong(firstElement(fromEl, "rowOff"));
    model.LowerRightColumn = fromCol + (cx > 0 ? Math.trunc(cx / EMU_PER_COLUMN) + 1 : 1);
    model.LowerRightRow = fromRow + (cy > 0 ? Math.trunc(cy / EMU_PER_ROW) + 1 : 1);
    model.ExtentCx = cx;
    model.ExtentCy = cy;
  }

  return model;
}

function getChartRelationshipId(graphicFrame: XmlNode): string | null {
  const graphic = firstElement(graphicFrame, "graphic");
  const graphicData = graphic === null ? null : firstElement(graphic, "graphicData");
  if (graphicData === null) {
    return null;
  }

  const chartElement = firstElement(graphicData, "chart");
  if (chartElement === null) {
    return null;
  }

  return attribute(chartElement, "id");
}

function loadChartForRelationship(
  relationshipId: string,
  chartName: string,
  rawContainerElement: XmlNode | null,
  chartTargets: Map<string, ChartTarget>,
  archive: ZipReader,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): ChartModel | null {
  const chartTarget = getIgnoreCaseChart(chartTargets, relationshipId);
  if (chartTarget === undefined) {
    addIssue(
      diagnostics,
      options,
      new IssueBuilder(
        "CHT-L001",
        DiagnosticSeverity.LossyRecoverable,
        `Chart relationship '${relationshipId}' could not be resolved; chart was dropped.`,
        true,
        true,
      ).withSheet(sheetName),
    );
    return null;
  }

  const chartUri = chartTarget.Uri;
  const isChartEx = chartTarget.IsChartEx;
  if (!archive.hasEntry(chartUri)) {
    addIssue(
      diagnostics,
      options,
      new IssueBuilder(
        "CHT-R001",
        DiagnosticSeverity.Recoverable,
        `${isChartEx ? "ChartEx" : "Chart"} part '${chartUri}' was not found; chart was skipped.`,
        true,
        false,
      ).withSheet(sheetName),
    );
    return null;
  }

  const rawChartXml = archive.readText(chartUri);

  const model = new ChartModel();
  model.Name = chartName;
  model.RawChartXml = rawChartXml;
  model.ChartType = detectChartType(rawChartXml);
  model.IsChartEx = isChartEx;
  model.OriginalRId = relationshipId;
  model.RawGraphicFrameXml = rawContainerElement !== null ? serializeNode(rawContainerElement) : null;

  loadChartCompanionFiles(model, chartUri, archive, diagnostics, options, sheetName);
  return model;
}

function getChartName(graphicFrame: XmlNode): string {
  const nvGraphicFramePr = firstElement(graphicFrame, "nvGraphicFramePr");
  if (nvGraphicFramePr === null) {
    return "";
  }

  const cNvPr = firstElement(nvGraphicFramePr, "cNvPr");
  return cNvPr === null ? "" : (attribute(cNvPr, "name") ?? "");
}

function detectChartType(rawXml: string | null): ChartType {
  if (rawXml === null || rawXml.length === 0) {
    return ChartType.Unknown;
  }

  if (rawXml.indexOf("<c:barChart") >= 0) {
    return rawXml.indexOf('barDir val="bar"') >= 0 ? ChartType.Bar : ChartType.Column;
  }

  if (rawXml.indexOf("<c:bar3DChart") >= 0) {
    return rawXml.indexOf('barDir val="bar"') >= 0 ? ChartType.Bar3D : ChartType.Column3D;
  }

  if (rawXml.indexOf("<c:lineChart") >= 0) {
    return ChartType.Line;
  }

  if (rawXml.indexOf("<c:line3DChart") >= 0) {
    return ChartType.Line3D;
  }

  if (rawXml.indexOf("<c:areaChart") >= 0) {
    return ChartType.Area;
  }

  if (rawXml.indexOf("<c:area3DChart") >= 0) {
    return ChartType.Area3D;
  }

  if (rawXml.indexOf("<c:pie3DChart") >= 0) {
    return ChartType.Pie3D;
  }

  if (rawXml.indexOf("<c:pieChart") >= 0) {
    return ChartType.Pie;
  }

  if (rawXml.indexOf("<c:doughnutChart") >= 0) {
    return ChartType.Doughnut;
  }

  if (rawXml.indexOf("<c:scatterChart") >= 0) {
    return ChartType.Scatter;
  }

  if (rawXml.indexOf("<c:bubbleChart") >= 0) {
    return ChartType.Bubble;
  }

  if (rawXml.indexOf("<c:radarChart") >= 0) {
    return ChartType.Radar;
  }

  if (rawXml.indexOf("<c:stockChart") >= 0) {
    return ChartType.Stock;
  }

  if (rawXml.indexOf("<c:surface3DChart") >= 0) {
    return rawXml.indexOf('wireframe val="1"') >= 0 ? ChartType.SurfaceWireframe3D : ChartType.Surface3D;
  }

  if (rawXml.indexOf("<c:surfaceChart") >= 0) {
    return ChartType.Contour;
  }

  return ChartType.Unknown;
}

function loadChartCompanionFiles(
  model: ChartModel,
  chartUri: string,
  archive: ZipReader,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  const chartRelsUri = getRelsUri(chartUri);
  if (!archive.hasEntry(chartRelsUri)) {
    return;
  }

  const relationshipsDocument = parseXml(archive.readText(chartRelsUri));
  const root = documentRoot(relationshipsDocument);
  if (root === null) {
    return;
  }

  for (const relationship of elementChildren(root, "Relationship")) {
    const relId = attribute(relationship, "Id");
    const type = attribute(relationship, "Type");
    const target = attribute(relationship, "Target");
    if (relId === null || relId.length === 0 || target === null || target.length === 0) {
      continue;
    }

    const isXmlCompanion =
      equalsIgnoreCase(type, CHART_STYLE_RELATIONSHIP_TYPE) ||
      equalsIgnoreCase(type, CHART_COLOR_STYLE_RELATIONSHIP_TYPE) ||
      equalsIgnoreCase(type, CHART_USER_SHAPES_RELATIONSHIP_TYPE);

    if (isXmlCompanion) {
      const companionUri = resolvePartUri(chartUri, target);
      if (!archive.hasEntry(companionUri)) {
        continue;
      }

      const companion = new ChartCompanionFile();
      companion.RelationshipId = relId;
      companion.RelationshipType = type;
      companion.FileName = target;
      companion.RawContent = archive.readText(companionUri);
      model.CompanionFiles.push(companion);
    } else if (equalsIgnoreCase(type, IMAGE_RELATIONSHIP_TYPE)) {
      const imageUri = resolvePartUri(chartUri, target);
      if (!archive.hasEntry(imageUri)) {
        addIssue(
          diagnostics,
          options,
          new IssueBuilder(
            "CHT-R002",
            DiagnosticSeverity.Recoverable,
            `Chart '${model.Name}' references companion image '${imageUri}' but the image part was not found.`,
            true,
            false,
          ).withSheet(sheetName),
        );
        continue;
      }

      const imageData = new Uint8Array(archive.readBinary(imageUri));
      if (imageData.length === 0) {
        addIssue(
          diagnostics,
          options,
          new IssueBuilder(
            "CHT-R003",
            DiagnosticSeverity.Recoverable,
            `Chart '${model.Name}' references companion image '${imageUri}' but the image part was empty.`,
            true,
            false,
          ).withSheet(sheetName),
        );
        continue;
      }

      const companion = new ChartCompanionFile();
      companion.RelationshipId = relId;
      companion.RelationshipType = type;
      companion.FileName = target;
      companion.BinaryContent = imageData;
      model.CompanionFiles.push(companion);
    }
  }
}

// --- Shared helpers ---

function findDrawingUri(archive: ZipReader, worksheetUri: string): string {
  const relsUri = getWorksheetRelsUri(worksheetUri);
  if (!archive.hasEntry(relsUri)) {
    return "";
  }

  const document = parseXml(archive.readText(relsUri));
  const root = documentRoot(document);
  if (root === null) {
    return "";
  }

  for (const relationship of elementChildren(root, "Relationship")) {
    const type = attribute(relationship, "Type");
    const target = attribute(relationship, "Target");
    if (equalsIgnoreCase(type, DRAWING_RELATIONSHIP_TYPE) && target !== null && target.length > 0) {
      return resolvePartUri(worksheetUri, target);
    }
  }

  return "";
}

function loadDrawingChartTargets(archive: ZipReader, drawingUri: string): Map<string, ChartTarget> {
  const targets = new Map<string, ChartTarget>();
  const relsUri = getDrawingRelsUri(drawingUri);
  if (!archive.hasEntry(relsUri)) {
    return targets;
  }

  const document = parseXml(archive.readText(relsUri));
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

    const isStandard = equalsIgnoreCase(type, CHART_RELATIONSHIP_TYPE);
    const isChartEx = equalsIgnoreCase(type, CHART_EX_RELATIONSHIP_TYPE);
    if (!isStandard && !isChartEx) {
      continue;
    }

    targets.set(id, { Uri: resolvePartUri(drawingUri, target), IsChartEx: isChartEx });
  }

  return targets;
}

function loadDrawingTargetsByType(archive: ZipReader, drawingUri: string, relationshipType: string): Map<string, string> {
  const targets = new Map<string, string>();
  const relsUri = getDrawingRelsUri(drawingUri);
  if (!archive.hasEntry(relsUri)) {
    return targets;
  }

  const document = parseXml(archive.readText(relsUri));
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

    if (!equalsIgnoreCase(type, relationshipType)) {
      continue;
    }

    targets.set(id, resolvePartUri(drawingUri, target));
  }

  return targets;
}

function parseAnchorInt(element: XmlNode | null): number {
  if (element === null) {
    return 0;
  }

  return tryParseInt(textOf(element)) ?? 0;
}

function parseAnchorLong(element: XmlNode | null): number {
  if (element === null) {
    return 0;
  }

  return tryParseLong(textOf(element)) ?? 0;
}

function parseExtentCoordinate(element: XmlNode | null, name: string): number {
  if (element === null) {
    return 0;
  }

  const value = attribute(element, name) ?? "0";
  return tryParseLong(value) ?? 0;
}

function getWorksheetRelsUri(worksheetUri: string): string {
  return getRelsUri(worksheetUri);
}

function getDrawingRelsUri(drawingUri: string): string {
  return getRelsUri(drawingUri);
}

function getRelsUri(partUri: string): string {
  const normalized = trimStartSlashes(partUri);
  const slashIndex = normalized.lastIndexOf("/");
  const directory = slashIndex >= 0 ? normalized.substring(0, slashIndex + 1) : "";
  const fileName = slashIndex >= 0 ? normalized.substring(slashIndex + 1) : normalized;
  return `/${directory}_rels/${fileName}.rels`;
}

function trimStartSlashes(value: string): string {
  let index = 0;
  while (index < value.length && value[index] === "/") {
    index++;
  }

  return value.substring(index);
}

function resolvePartUri(sourceUri: string, target: string): string {
  if (target.startsWith("/")) {
    return target;
  }

  const resolved = new URL(target, `http://package${sourceUri}`);
  return resolved.pathname;
}

function serializeNode(node: XmlNode): string {
  return (node as unknown as { toString(): string }).toString();
}

function descendantsIncludingSelf(root: XmlNode): XmlNode[] {
  const result: XmlNode[] = [];
  const visit = (current: XmlNode): void => {
    result.push(current);
    const children = (current as unknown as { childNodes: ArrayLike<XmlNode> | null }).childNodes;
    if (children === null) {
      return;
    }

    for (let index = 0; index < children.length; index++) {
      const child = children[index] as XmlNode;
      if (localNameOf(child.nodeName) !== "#text" && localNameOf(child.nodeName) !== "#comment") {
        visit(child);
      }
    }
  };

  visit(root);
  return result;
}

function localNameOf(nodeName: string): string {
  const separator = nodeName.indexOf(":");
  return separator >= 0 ? nodeName.substring(separator + 1) : nodeName;
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

function detectExtension(data: Uint8Array | null): string {
  if (data === null || data.length < 4) {
    return "bin";
  }

  if (data[0] === 0xff && data[1] === 0xd8) {
    return "jpeg";
  }

  if (data[0] === 0x89 && data[1] === 0x50 && data[2] === 0x4e && data[3] === 0x47) {
    return "png";
  }

  if (data[0] === 0x47 && data[1] === 0x49 && data[2] === 0x46) {
    return "gif";
  }

  if (data[0] === 0x42 && data[1] === 0x4d) {
    return "bmp";
  }

  return "bin";
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

function tryParseLong(value: string): number | null {
  const trimmed = value.trim();
  if (!/^[+-]?\d+$/.test(trimmed)) {
    return null;
  }

  const parsed = Number(trimmed);
  if (!Number.isInteger(parsed) || !Number.isSafeInteger(parsed)) {
    return null;
  }

  return parsed;
}

function equalsIgnoreCase(left: string | null, right: string): boolean {
  return left !== null && left.toLowerCase() === right.toLowerCase();
}

function getIgnoreCase(map: Map<string, string>, key: string): string | undefined {
  const direct = map.get(key);
  if (direct !== undefined) {
    return direct;
  }

  const normalizedKey = key.toLowerCase();
  for (const [candidateKey, value] of map) {
    if (candidateKey.toLowerCase() === normalizedKey) {
      return value;
    }
  }

  return undefined;
}

function getIgnoreCaseChart(map: Map<string, ChartTarget>, key: string): ChartTarget | undefined {
  const direct = map.get(key);
  if (direct !== undefined) {
    return direct;
  }

  const normalizedKey = key.toLowerCase();
  for (const [candidateKey, value] of map) {
    if (candidateKey.toLowerCase() === normalizedKey) {
      return value;
    }
  }

  return undefined;
}
