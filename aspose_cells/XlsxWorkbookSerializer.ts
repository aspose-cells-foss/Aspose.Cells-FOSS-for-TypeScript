import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { ChartCompanionFile, ChartModel } from "./Core/ChartModel";
import { WorksheetModel } from "./Core/WorksheetModel";
import { WorkbookModel } from "./Core/WorkbookModel";
import { CellValueKind } from "./Core/CellValueKind";
import { SheetVisibility } from "./Core/SheetVisibility";
import { SharedStringEntry } from "./SharedStringEntry";
import { SharedStringRepository } from "./Core/SharedStringRepository";
import { RichTextRunValue } from "./Core/RichTextRunValue";
import { InvalidFileFormatException } from "./InvalidFileFormatException";
import { LoadDiagnostics } from "./LoadDiagnostics";
import { LoadOptions } from "./LoadOptions";
import { SaveFormat } from "./enums";
import { SaveOptions } from "./SaveOptions";
import { WorkbookSaveException } from "./WorkbookSaveException";
import { UnsupportedFeatureException } from "./UnsupportedFeatureException";
import { ChartXmlTemplates } from "./ChartXmlTemplates";
import { ZipReader, ZipWriter } from "./Packaging/Archive";
import {
  PACKAGE_RELATIONSHIP_NS,
  buildContentTypes,
  buildRootRelationships,
  buildSharedStrings,
  buildWorkbook,
  buildWorkbookRelationships,
  shouldPersistCell,
  writeXmlEntry,
} from "./XlsxWorkbookSerializerCommon";
import { buildCommentsDocument, computeCommentFileNumbers, computeTotalCommentCount, writeVmlDrawing } from "./XlsxWorkbookComments";
import { loadWorkbookDefinedNames } from "./XlsxWorkbookDefinedNames";
import { buildCorePropertiesDocument, buildExtendedPropertiesDocument, loadDocumentProperties } from "./XlsxDocumentProperties";
import { countExternalHyperlinks, loadWorksheetHyperlinkTargets } from "./XlsxWorkbookHyperlinks";
import {
  CHART_COLOR_STYLE_CONTENT_TYPE,
  CHART_COLOR_STYLE_RELATIONSHIP_TYPE,
  CHART_CONTENT_TYPE,
  CHART_EX_CONTENT_TYPE,
  CHART_STYLE_CONTENT_TYPE,
  CHART_STYLE_RELATIONSHIP_TYPE,
  CHART_USER_SHAPES_CONTENT_TYPE,
  CHART_USER_SHAPES_RELATIONSHIP_TYPE,
  IMAGE_RELATIONSHIP_TYPE,
  buildDrawingDocument,
  buildDrawingRelationshipsDocument,
  collectImageExtensions,
  computeDrawingNumbers,
  computePictureFileOffsets,
  computeTotalDrawingCount,
  writePictureMediaEntries,
} from "./XlsxWorkbookPictures";
import {
  buildTableDocument,
  buildWorksheetRelationshipsDocument,
  computeTableFileOffsets,
  computeTotalTableCount,
} from "./XlsxWorkbookTables";
import { buildWorksheet } from "./XlsxWorkbookWorksheetWriter";
import { loadWorksheet } from "./XlsxWorkbookWorksheetLoader";
import { buildStylesheet, loadStylesheet } from "./XlsxWorkbookStyles";
import { loadWorksheetDefinedNames } from "./XlsxWorkbookPageSetup";
import { loadWorkbookMetadata } from "./XlsxWorkbookProperties";
import {
  getRelationshipTarget,
  loadDocument,
  loadRelationships,
  loadSharedStrings,
  resolveWorksheetUri,
} from "./XlsxWorkbookArchiveHelpers";
import { XmlElement } from "./Xml/XmlWriter";
import { XmlNode, attribute, documentRoot, elementChildren, firstElement } from "./Xml/XmlReader";

/** Saves the workbook model into an XLSX package. */
export function saveWorkbook(model: WorkbookModel, options: SaveOptions): Buffer {
  if (options.SaveFormat !== SaveFormat.Xlsx) {
    throw new UnsupportedFeatureException("Only XLSX save is supported.");
  }

  if (model.Worksheets.length === 0) {
    throw new WorkbookSaveException("A workbook must contain at least one worksheet.");
  }

  const sharedStrings = new SharedStringRepository();
  const stylesheet = buildStylesheet(model);
  const corePropertiesDocument = buildCorePropertiesDocument(model);
  const extendedPropertiesDocument = buildExtendedPropertiesDocument(model);

  for (const worksheet of model.Worksheets) {
    for (const [, record] of worksheet.Cells) {
      if (!shouldPersistCell(model.DefaultStyle, record)) {
        continue;
      }

      if (record.Kind === CellValueKind.String && options.UseSharedStrings) {
        const text = record.Value;
        if (typeof text === "string") {
          if (record.RichTextRuns !== null && record.RichTextRuns.length > 0) {
            const entry = new SharedStringEntry();
            entry.Text = text;
            entry.Runs = cloneRichTextRuns(record.RichTextRuns);
            sharedStrings.Intern(entry);
          } else {
            sharedStrings.Intern(text);
          }
        }
      }
    }
  }

  const tableFileOffsets = computeTableFileOffsets(model);
  const totalTableCount = computeTotalTableCount(model);
  const pictureFileOffsets = computePictureFileOffsets(model);
  const chartFileOffsets = computeChartFileOffsets(model);
  const drawingNumbers = computeDrawingNumbers(model);
  const totalDrawingCount = computeTotalDrawingCount(model);
  const imageExtensions = collectImageExtensions(model);
  const chartPartNames = collectChartPartNames(model, chartFileOffsets);
  const chartContentTypes = collectChartContentTypes(model);
  const chartCompanionPartNames = collectChartCompanionPartNames(model, chartFileOffsets, totalDrawingCount);
  const chartCompanionContentTypes = collectChartCompanionContentTypes(model);
  const commentFileNumbers = computeCommentFileNumbers(model);
  const totalCommentCount = computeTotalCommentCount(model);

  const archive = new ZipWriter();
  let userShapesDrawingCounter = totalDrawingCount;

  const hasSharedStrings = sharedStrings.Entries.length > 0;
  const hasTheme = model.RawThemeXml !== null && model.RawThemeXml.length > 0;
  const externalLinkBaseRId =
    model.Worksheets.length + (hasSharedStrings ? 1 : 0) + 1 + (hasTheme ? 1 : 0) + 1;

  writeXmlEntry(
    archive,
    "[Content_Types].xml",
    buildContentTypes(
      model,
      hasSharedStrings,
      true,
      corePropertiesDocument !== null,
      extendedPropertiesDocument !== null,
      totalTableCount,
      totalDrawingCount,
      imageExtensions,
      chartPartNames,
      chartContentTypes,
      chartCompanionPartNames,
      chartCompanionContentTypes,
      hasTheme,
      totalCommentCount,
    ),
  );
  writeXmlEntry(
    archive,
    "_rels/.rels",
    buildRootRelationships(corePropertiesDocument !== null, extendedPropertiesDocument !== null),
  );
  writeXmlEntry(archive, "xl/workbook.xml", buildWorkbook(model, externalLinkBaseRId));
  writeXmlEntry(archive, "xl/_rels/workbook.xml.rels", buildWorkbookRelationships(model, hasSharedStrings, true));

  if (hasTheme) {
    archive.addText("xl/theme/theme1.xml", model.RawThemeXml ?? "");
  }

  for (let index = 0; index < model.ExternalLinks.length; index++) {
    const externalLink = model.ExternalLinks[index]!;
    const externalNumber = String(index + 1);
    archive.addText("xl/externalLinks/externalLink" + externalNumber + ".xml", externalLink.RawXml ?? "");
    if (externalLink.RawRelsXml !== null && externalLink.RawRelsXml.length > 0) {
      archive.addText(
        "xl/externalLinks/_rels/externalLink" + externalNumber + ".xml.rels",
        externalLink.RawRelsXml,
      );
    }
  }

  for (let index = 0; index < model.Worksheets.length; index++) {
    const worksheet = model.Worksheets[index]!;
    const tableFileOffset = tableFileOffsets[index]!;
    const pictureFileOffset = pictureFileOffsets[index]!;
    const chartFileOffset = chartFileOffsets[index]!;
    const drawingNumber = drawingNumbers[index]!;
    const commentFileNumber = commentFileNumbers[index]!;
    const hasDrawing =
      worksheet.Pictures.length > 0 ||
      worksheet.Shapes.length > 0 ||
      worksheet.Charts.length > 0 ||
      worksheet.PreservedCharts.length > 0;
    const hasComments = worksheet.Comments.length > 0;
    const externalHyperlinkCount = countExternalHyperlinks(worksheet);

    for (let tableIndex = 0; tableIndex < worksheet.ListObjects.length; tableIndex++) {
      const globalTableNumber = tableFileOffset + tableIndex + 1;
      writeXmlEntry(
        archive,
        "xl/tables/table" + globalTableNumber + ".xml",
        buildTableDocument(worksheet.ListObjects[tableIndex]!, globalTableNumber),
      );
    }

    if (hasDrawing) {
      writeXmlEntry(
        archive,
        "xl/drawings/drawing" + drawingNumber + ".xml",
        buildDrawingDocument(worksheet, pictureFileOffset, chartFileOffset),
      );

      if (worksheet.Pictures.length > 0 || worksheet.ShapeImages.length > 0) {
        writePictureMediaEntries(archive, worksheet, pictureFileOffset);
      }

      if (
        worksheet.Pictures.length > 0 ||
        worksheet.ShapeImages.length > 0 ||
        worksheet.Charts.length > 0 ||
        worksheet.PreservedCharts.length > 0 ||
        worksheet.PreservedDrawingRelationships.length > 0
      ) {
        writeXmlEntry(
          archive,
          "xl/drawings/_rels/drawing" + drawingNumber + ".xml.rels",
          buildDrawingRelationshipsDocument(worksheet, pictureFileOffset, chartFileOffset),
        );
      }

      for (let chartIndex = 0; chartIndex < worksheet.Charts.length; chartIndex++) {
        const globalChartNumber = chartFileOffset + chartIndex + 1;
        userShapesDrawingCounter = writeChartFiles(
          archive,
          worksheet.Charts[chartIndex]!,
          globalChartNumber,
          userShapesDrawingCounter,
        );
      }

      for (let chartIndex = 0; chartIndex < worksheet.PreservedCharts.length; chartIndex++) {
        const globalChartNumber = chartFileOffset + worksheet.Charts.length + chartIndex + 1;
        userShapesDrawingCounter = writeChartFiles(
          archive,
          worksheet.PreservedCharts[chartIndex]!,
          globalChartNumber,
          userShapesDrawingCounter,
        );
      }
    }

    if (hasComments) {
      writeXmlEntry(archive, "xl/comments" + commentFileNumber + ".xml", buildCommentsDocument(worksheet));
      writeVmlDrawing(archive, worksheet, commentFileNumber);
    }

    writeXmlEntry(
      archive,
      "xl/worksheets/sheet" + (index + 1) + ".xml",
      buildWorksheet(
        worksheet,
        model.DefaultStyle,
        model.Settings.DateSystem,
        sharedStrings,
        options,
        stylesheet,
        externalHyperlinkCount,
        hasDrawing,
        hasComments,
      ),
    );

    const worksheetRelationships = buildWorksheetRelationshipsDocument(
      worksheet,
      tableFileOffset,
      drawingNumber,
      commentFileNumber,
    );
    if (worksheetRelationships !== null) {
      writeXmlEntry(archive, "xl/worksheets/_rels/sheet" + (index + 1) + ".xml.rels", worksheetRelationships);
    }
  }

  if (sharedStrings.Entries.length > 0) {
    writeXmlEntry(archive, "xl/sharedStrings.xml", buildSharedStrings(sharedStrings));
  }

  writeXmlEntry(archive, "xl/styles.xml", stylesheet.Document);

  if (corePropertiesDocument !== null) {
    writeXmlEntry(archive, "docProps/core.xml", corePropertiesDocument);
  }

  if (extendedPropertiesDocument !== null) {
    writeXmlEntry(archive, "docProps/app.xml", extendedPropertiesDocument);
  }

  return archive.toBuffer();
}

/** Loads a workbook model from an XLSX package. */
export function loadWorkbook(
  data: Buffer,
  options: LoadOptions,
  diagnostics: LoadDiagnostics,
): WorkbookModel {
  let archive: ZipReader;
  try {
    archive = new ZipReader(data);
  } catch (error) {
    throw new InvalidFileFormatException("The workbook is not a valid XLSX zip package.", error);
  }

  if (!archive.hasEntry("/xl/workbook.xml")) {
    throw new InvalidFileFormatException("The package does not contain /xl/workbook.xml.");
  }

  const workbookDocument = loadDocument(archive, "/xl/workbook.xml");
  const workbookRelationships = loadRelationships(archive, "/xl/_rels/workbook.xml.rels", "/xl/workbook.xml");
  const sharedStrings = loadSharedStrings(archive, workbookRelationships, options, diagnostics);
  const stylesheet = loadStylesheet(archive, workbookRelationships, options, diagnostics);

  const workbookModel = new WorkbookModel();
  workbookModel.Worksheets.length = 0;
  workbookModel.DefaultStyle = stylesheet.DefaultCellStyle.clone();
  workbookModel.RawTableStylesXml = stylesheet.RawTableStylesXml;
  workbookModel.TableDifferentialStyles.push(...stylesheet.DifferentialFormats);
  workbookModel.RawTableDifferentialStylesXml.push(...stylesheet.RawDifferentialFormatsXml);

  const workbookRoot = documentRoot(workbookDocument);
  if (workbookRoot === null) {
    throw new InvalidFileFormatException("The workbook XML is empty.");
  }

  const worksheetDefinedNames = loadWorksheetDefinedNames(workbookRoot, diagnostics, options);

  const sheets = firstElement(workbookRoot, "sheets");
  if (sheets === null) {
    throw new InvalidFileFormatException("The workbook XML does not contain a sheets element.");
  }

  const sheetElements = elementChildren(sheets, "sheet");
  if (sheetElements.length === 0) {
    throw new InvalidFileFormatException("The workbook XML does not contain any worksheets.");
  }

  loadWorkbookMetadata(workbookRoot, workbookModel, sheetElements.length, diagnostics, options);
  loadDocumentProperties(archive, workbookModel, diagnostics, options);
  loadWorkbookDefinedNames(workbookRoot, workbookModel, sheetElements.length, diagnostics, options);
  workbookModel.RawThemeXml = loadRawTheme(archive);
  loadExternalLinks(archive, workbookRoot, workbookRelationships, workbookModel);

  workbookModel.ActiveSheetIndex =
    workbookModel.ActiveSheetIndex < sheetElements.length ? workbookModel.ActiveSheetIndex : 0;

  for (let index = 0; index < sheetElements.length; index++) {
    const sheetElement = sheetElements[index]!;
    const sheetName = (attribute(sheetElement, "name") ?? "").trim();
    const resolvedSheetName = sheetName.length === 0 ? "Sheet" + (index + 1) : sheetName;

    const relationshipId = attribute(sheetElement, "r:id");
    const worksheetUri = resolveWorksheetUri(
      index,
      relationshipId,
      workbookRelationships,
      archive,
      diagnostics,
      resolvedSheetName,
      options,
    );

    if (!archive.hasEntry(worksheetUri)) {
      throw new InvalidFileFormatException("Worksheet part '" + worksheetUri + "' was not found.");
    }

    const worksheetHyperlinkTargets = loadWorksheetHyperlinkTargets(archive, worksheetUri);
    const definedNamesState = worksheetDefinedNames.get(index) ?? null;
    const worksheetModel = loadWorksheet(
      resolvedSheetName,
      loadDocument(archive, worksheetUri),
      worksheetHyperlinkTargets,
      workbookModel.Settings.DateSystem,
      sharedStrings,
      stylesheet,
      diagnostics,
      options,
      definedNamesState,
      archive,
      worksheetUri,
    );

    const state = attribute(sheetElement, "state");
    if (state === "hidden") {
      worksheetModel.Visibility = SheetVisibility.Hidden;
    } else if (state === "veryHidden") {
      worksheetModel.Visibility = SheetVisibility.VeryHidden;
    } else {
      worksheetModel.Visibility = SheetVisibility.Visible;
    }

    workbookModel.Worksheets.push(worksheetModel);
  }

  return workbookModel;
}

function cloneRichTextRuns(runs: readonly RichTextRunValue[]): RichTextRunValue[] {
  const cloned: RichTextRunValue[] = [];
  for (const run of runs) {
    cloned.push(run.clone());
  }

  return cloned;
}

function computeChartFileOffsets(model: WorkbookModel): number[] {
  const offsets: number[] = [];
  let running = 0;
  for (const worksheet of model.Worksheets) {
    offsets.push(running);
    running += worksheet.Charts.length + worksheet.PreservedCharts.length;
  }

  return offsets;
}

function collectChartCompanionPartNames(
  model: WorkbookModel,
  chartFileOffsets: readonly number[],
  totalDrawingCount: number,
): string[] {
  const result: string[] = [];
  let userShapesCounter = totalDrawingCount;
  for (let index = 0; index < model.Worksheets.length; index++) {
    const charts = model.Worksheets[index]!.Charts;
    for (let chartIndex = 0; chartIndex < charts.length; chartIndex++) {
      const globalChartNumber = chartFileOffsets[index]! + chartIndex + 1;
      for (const companion of charts[chartIndex]!.CompanionFiles) {
        if (isImageCompanion(companion)) {
          continue;
        }

        if (hasCrossDirectoryPath(companion)) {
          userShapesCounter++;
          result.push("/xl/drawings/drawing" + userShapesCounter + ".xml");
        } else {
          result.push("/" + resolveCompanionPath(renumberCompanionFileName(companion.FileName ?? "", globalChartNumber)));
        }
      }
    }

    const preservedCharts = model.Worksheets[index]!.PreservedCharts;
    for (let chartIndex = 0; chartIndex < preservedCharts.length; chartIndex++) {
      const globalChartNumber = chartFileOffsets[index]! + charts.length + chartIndex + 1;
      for (const companion of preservedCharts[chartIndex]!.CompanionFiles) {
        if (isImageCompanion(companion)) {
          continue;
        }

        if (hasCrossDirectoryPath(companion)) {
          userShapesCounter++;
          result.push("/xl/drawings/drawing" + userShapesCounter + ".xml");
        } else {
          result.push("/" + resolveCompanionPath(renumberCompanionFileName(companion.FileName ?? "", globalChartNumber)));
        }
      }
    }
  }

  return result;
}

function collectChartCompanionContentTypes(model: WorkbookModel): string[] {
  const result: string[] = [];
  for (const worksheet of model.Worksheets) {
    for (const chart of worksheet.Charts) {
      for (const companion of chart.CompanionFiles) {
        if (!isImageCompanion(companion)) {
          result.push(getCompanionContentType(companion.RelationshipType));
        }
      }
    }

    for (const chart of worksheet.PreservedCharts) {
      for (const companion of chart.CompanionFiles) {
        if (!isImageCompanion(companion)) {
          result.push(getCompanionContentType(companion.RelationshipType));
        }
      }
    }
  }

  return result;
}

function collectChartPartNames(model: WorkbookModel, chartFileOffsets: readonly number[]): string[] {
  const result: string[] = [];
  for (let index = 0; index < model.Worksheets.length; index++) {
    const charts = model.Worksheets[index]!.Charts;
    for (let chartIndex = 0; chartIndex < charts.length; chartIndex++) {
      const globalChartNumber = String(chartFileOffsets[index]! + chartIndex + 1);
      result.push(
        charts[chartIndex]!.IsChartEx ? "/xl/charts/chartEx" + globalChartNumber + ".xml" : "/xl/charts/chart" + globalChartNumber + ".xml",
      );
    }

    const preservedCharts = model.Worksheets[index]!.PreservedCharts;
    for (let chartIndex = 0; chartIndex < preservedCharts.length; chartIndex++) {
      const globalChartNumber = String(chartFileOffsets[index]! + charts.length + chartIndex + 1);
      result.push(
        preservedCharts[chartIndex]!.IsChartEx
          ? "/xl/charts/chartEx" + globalChartNumber + ".xml"
          : "/xl/charts/chart" + globalChartNumber + ".xml",
      );
    }
  }

  return result;
}

function collectChartContentTypes(model: WorkbookModel): string[] {
  const result: string[] = [];
  for (const worksheet of model.Worksheets) {
    for (const chart of worksheet.Charts) {
      result.push(chart.IsChartEx ? CHART_EX_CONTENT_TYPE : CHART_CONTENT_TYPE);
    }

    for (const chart of worksheet.PreservedCharts) {
      result.push(chart.IsChartEx ? CHART_EX_CONTENT_TYPE : CHART_CONTENT_TYPE);
    }
  }

  return result;
}

function getCompanionContentType(relationshipType: string | null): string {
  const type = relationshipType ?? "";
  if (equalsIgnoreCase(type, CHART_STYLE_RELATIONSHIP_TYPE)) {
    return CHART_STYLE_CONTENT_TYPE;
  }

  if (equalsIgnoreCase(type, CHART_COLOR_STYLE_RELATIONSHIP_TYPE)) {
    return CHART_COLOR_STYLE_CONTENT_TYPE;
  }

  if (equalsIgnoreCase(type, CHART_USER_SHAPES_RELATIONSHIP_TYPE)) {
    return CHART_USER_SHAPES_CONTENT_TYPE;
  }

  return "application/octet-stream";
}

function writeChartFiles(
  archive: ZipWriter,
  chart: ChartModel,
  globalChartNumber: number,
  userShapesDrawingCounter: number,
): number {
  const chartFileName = (chart.IsChartEx ? "chartEx" : "chart") + String(globalChartNumber) + ".xml";
  const chartPath = "xl/charts/" + chartFileName;
  archive.addText(chartPath, buildChartXml(chart));

  if (chart.CompanionFiles.length === 0) {
    return userShapesDrawingCounter;
  }

  const companionRels = XmlElement.createWithNamespace("Relationships", PACKAGE_RELATIONSHIP_NS);
  for (const companion of chart.CompanionFiles) {
    let companionPath: string;
    let relsTarget: string;

    if (isImageCompanion(companion)) {
      const imageFileName = buildChartCompanionImageFileName(
        companion.FileName ?? "",
        globalChartNumber,
        companion.RelationshipId,
      );
      companionPath = "xl/media/" + imageFileName;
      relsTarget = "../media/" + imageFileName;
    } else if (hasCrossDirectoryPath(companion)) {
      userShapesDrawingCounter++;
      const drawingName = "drawing" + String(userShapesDrawingCounter) + ".xml";
      companionPath = "xl/drawings/" + drawingName;
      relsTarget = "../drawings/" + drawingName;
    } else {
      const newFileName = renumberCompanionFileName(companion.FileName ?? "", globalChartNumber);
      companionPath = resolveCompanionPath(newFileName);
      relsTarget = newFileName;
    }

    if (companion.BinaryContent !== null) {
      archive.addBinary(companionPath, companion.BinaryContent);
    } else {
      archive.addText(companionPath, companion.RawContent ?? "");
    }

    companionRels.child(
      XmlElement.create("Relationship")
        .attr("Id", companion.RelationshipId)
        .attr("Type", companion.RelationshipType)
        .attr("Target", relsTarget),
    );
  }

  writeXmlEntry(archive, "xl/charts/_rels/" + chartFileName + ".rels", companionRels);
  return userShapesDrawingCounter;
}

function buildChartXml(chart: ChartModel): string {
  const rawChartXml = chart.RawChartXml;
  if (rawChartXml !== null && rawChartXml.length > 0) {
    return normalizeChartXmlForSave(rawChartXml);
  }

  // Programmatically created charts carry no raw XML; generate the chart part from the model
  // exactly like Aspose.Cells_FOSS/XlsxWorkbookSerializer.cs does through ChartXmlTemplates.
  return ChartXmlTemplates.Build(chart.ChartType, chart.DataRange, chart.Name);
}

function normalizeChartXmlForSave(rawChartXml: string | null): string {
  if (rawChartXml === null || rawChartXml.length === 0) {
    return rawChartXml ?? "";
  }

  let document: DomDocument;
  try {
    const parser = new DOMParser({
      errorHandler: {
        warning: () => undefined,
        error: () => undefined,
        fatalError: () => undefined,
      },
    });
    document = parser.parseFromString(rawChartXml, "text/xml") as unknown as DomDocument;
  } catch {
    return rawChartXml;
  }

  if (document.documentElement === null) {
    return rawChartXml;
  }

  const changed = removeInvalidChartExtensionChildren(document.documentElement);
  if (!changed) {
    return rawChartXml;
  }

  const declarationEnd = rawChartXml.startsWith("<?xml") ? rawChartXml.indexOf("?>") : -1;
  const declaration = declarationEnd >= 0 ? rawChartXml.substring(0, declarationEnd + 2) : "";
  return declaration + serializeDomNode(document.documentElement);
}

function removeInvalidChartExtensionChildren(root: DomElement): boolean {
  let changed = false;
  const chart2012Namespace = "http://schemas.microsoft.com/office/drawing/2012/chart";
  const chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
  const chartExNamespace = "http://schemas.microsoft.com/office/drawing/2014/chartex";

  const candidates = collectDomDescendants(root);
  for (const element of candidates) {
    if (element.namespaceURI === chart2012Namespace && element.localName === "showLeaderLines") {
      removeDomElement(element);
      changed = true;
    }
  }

  const chartElement = findDomChild(root, chartNamespace, "chart");
  if (chartElement !== null) {
    const showLabelsOverMax = findDomChild(chartElement, chartNamespace, "showDLblsOverMax");
    if (showLabelsOverMax !== null) {
      removeDomElement(showLabelsOverMax);
      changed = true;
    }

    const chartLevelExtList = findDomChild(chartElement, chartNamespace, "extLst");
    if (chartLevelExtList !== null) {
      removeDomElement(chartLevelExtList);
      changed = true;
    }
  }

  for (const element of collectDomDescendants(root)) {
    if (element.namespaceURI !== chartNamespace || element.localName !== "surfaceChart") {
      continue;
    }

    const varyColors = findDomChild(element, chartNamespace, "varyColors");
    if (varyColors !== null) {
      removeDomElement(varyColors);
      changed = true;
    }
  }

  for (const element of collectDomDescendants(root)) {
    if (element.namespaceURI !== chartExNamespace || element.localName !== "series") {
      continue;
    }

    const valueColors = findDomChild(element, chartExNamespace, "valueColors");
    if (valueColors !== null) {
      removeDomElement(valueColors);
      changed = true;
    }

    const valueColorPositions = findDomChild(element, chartExNamespace, "valueColorPositions");
    if (valueColorPositions !== null) {
      removeDomElement(valueColorPositions);
      changed = true;
    }
  }

  for (const element of collectDomDescendants(root)) {
    if (element.localName !== "ext") {
      continue;
    }

    if (hasDomElementChildren(element)) {
      continue;
    }

    removeDomElement(element);
    changed = true;
  }

  for (const element of collectDomDescendants(root)) {
    if (element.localName !== "extLst") {
      continue;
    }

    if (hasDomElementChildren(element)) {
      continue;
    }

    removeDomElement(element);
    changed = true;
  }

  return changed;
}

function collectDomDescendants(root: DomElement): DomElement[] {
  const result: DomElement[] = [];
  const visit = (element: DomElement): void => {
    for (let index = 0; index < element.childNodes.length; index++) {
      const child = element.childNodes.item(index) as DomNode;
      if (child.nodeType === 1) {
        result.push(child as DomElement);
        visit(child as DomElement);
      }
    }
  };

  visit(root);
  return result;
}

function findDomChild(element: DomElement, namespaceUri: string, localNameValue: string): DomElement | null {
  for (let index = 0; index < element.childNodes.length; index++) {
    const child = element.childNodes.item(index) as DomNode;
    if (child.nodeType !== 1) {
      continue;
    }

    const candidate = child as DomElement;
    if (candidate.namespaceURI === namespaceUri && candidate.localName === localNameValue) {
      return candidate;
    }
  }

  return null;
}

function hasDomElementChildren(element: DomElement): boolean {
  for (let index = 0; index < element.childNodes.length; index++) {
    if ((element.childNodes.item(index) as DomNode).nodeType === 1) {
      return true;
    }
  }

  return false;
}

function removeDomElement(element: DomElement): void {
  if (element.parentNode !== null) {
    element.parentNode.removeChild(element);
  }
}

function serializeDomNode(node: DomNode): string {
  const serializer = new XMLSerializer();
  return serializer.serializeToString(node as unknown as Node);
}

function buildChartCompanionImageFileName(
  originalTarget: string,
  globalChartNumber: number,
  relationshipId: string | null,
): string {
  const extension = getCompanionFileExtension(originalTarget);
  const baseName = "chart" + String(globalChartNumber) + "-" + sanitizeRelationshipId(relationshipId);
  return extension.length === 0 ? baseName + ".bin" : baseName + "." + extension;
}

function getCompanionFileExtension(target: string): string {
  if (target.length === 0) {
    return "";
  }

  const normalized = target.replace(/\\/g, "/");
  const fileName = normalized.substring(normalized.lastIndexOf("/") + 1);
  if (fileName.length === 0) {
    return "";
  }

  const dotIndex = fileName.lastIndexOf(".");
  if (dotIndex < 0 || dotIndex === fileName.length - 1) {
    return "";
  }

  return fileName.substring(dotIndex + 1).toLowerCase();
}

function sanitizeRelationshipId(relationshipId: string | null): string {
  if (relationshipId === null || relationshipId.length === 0) {
    return "companion";
  }

  let builder = "";
  for (let index = 0; index < relationshipId.length; index++) {
    const character = relationshipId[index]!;
    builder += /[A-Za-z0-9]/.test(character) ? character : "_";
  }

  return builder;
}

function renumberCompanionFileName(originalName: string, globalNumber: number): string {
  if (originalName.includes("/")) {
    return originalName;
  }

  const suffixStart = originalName.lastIndexOf(".");
  if (suffixStart < 0) {
    return originalName;
  }

  const suffix = originalName.substring(suffixStart);
  const basePart = originalName.substring(0, suffixStart);
  const trimmed = basePart.replace(/[0-9]+$/, "");
  if (trimmed.length === 0 || trimmed === basePart) {
    return originalName;
  }

  return trimmed + String(globalNumber) + suffix;
}

function resolveCompanionPath(renumberedTarget: string): string {
  const parts = ("xl/charts/" + renumberedTarget).split("/");
  const stack: string[] = [];
  for (const part of parts) {
    if (part === "..") {
      if (stack.length > 0) {
        stack.pop();
      }
    } else if (part.length > 0 && part !== ".") {
      stack.push(part);
    }
  }

  return stack.join("/");
}

function isImageCompanion(companion: ChartCompanionFile): boolean {
  return equalsIgnoreCase(companion.RelationshipType ?? "", IMAGE_RELATIONSHIP_TYPE);
}

function hasCrossDirectoryPath(companion: ChartCompanionFile): boolean {
  return (companion.FileName ?? "").includes("/");
}

function equalsIgnoreCase(left: string, right: string): boolean {
  return left.toLowerCase() === right.toLowerCase();
}

function loadRawTheme(archive: ZipReader): string | null {
  if (!archive.hasEntry("xl/theme/theme1.xml")) {
    return null;
  }

  return archive.readText("xl/theme/theme1.xml");
}

function loadExternalLinks(
  archive: ZipReader,
  workbookRoot: XmlNode,
  workbookRelationships: ReadonlyMap<string, string>,
  workbookModel: WorkbookModel,
): void {
  const externalReferencesElement = firstElement(workbookRoot, "externalReferences");
  if (externalReferencesElement === null) {
    return;
  }

  for (const extRef of elementChildren(externalReferencesElement, "externalReference")) {
    const relationshipId = attribute(extRef, "r:id");
    if (relationshipId === null || relationshipId.length === 0) {
      continue;
    }

    const partUri = getRelationshipTarget(workbookRelationships, relationshipId);
    if (partUri === null) {
      continue;
    }

    if (!archive.hasEntry(partUri)) {
      continue;
    }

    const rawXml = archive.readText(partUri);
    const normalizedUri = partUri.startsWith("/") ? partUri.substring(1) : partUri;
    const lastSlash = normalizedUri.lastIndexOf("/");
    const directory = lastSlash >= 0 ? normalizedUri.substring(0, lastSlash + 1) : "";
    const fileName = lastSlash >= 0 ? normalizedUri.substring(lastSlash + 1) : normalizedUri;
    const relsUri = "/" + directory + "_rels/" + fileName + ".rels";

    const rawRelsXml = archive.hasEntry(relsUri) ? archive.readText(relsUri) : null;
    workbookModel.ExternalLinks.push({ RawXml: rawXml, RawRelsXml: rawRelsXml });
  }
}

// Minimal DOM shape used for chart XML normalization.
interface DomNode {
  nodeType: number;
  nodeName: string;
  namespaceURI: string | null;
  localName: string | null;
  childNodes: { length: number; item(index: number): DomNode | null };
  parentNode: { removeChild(child: DomNode): DomNode } | null;
}

interface DomElement extends DomNode {
  namespaceURI: string | null;
  localName: string | null;
}

interface DomDocument {
  documentElement: DomElement | null;
}

