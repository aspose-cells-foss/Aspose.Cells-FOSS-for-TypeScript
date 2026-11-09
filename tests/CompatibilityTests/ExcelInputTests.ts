import * as path from "path";
import {
  AutoShapeType,
  ImageType,
  TableStyleType,
  Workbook,
} from "../../aspose_cells";
import {
  XmlNode,
  attribute,
  descendants,
  documentRoot,
  elementChildren,
  firstElement,
  parseXml,
} from "../../aspose_cells/Xml/XmlReader";
import {
  AssertEx,
  TemporaryDirectory,
  ZipPackageHelper,
  getOrCreateOutputDirectory,
} from "../Shared/TestInfrastructure";
import {
  appendOutSuffix,
  childPath,
  firstQualifiedChild,
  listXlsxFiles,
  resolveRepositoryFile,
} from "./CompatSupport";

export function excelInputAutoFilterRoundtripPreservesHeaderOnlyRange(): void {
  const inputPath = resolveRepositoryFile("Input", "Autofilter.xlsx");
  const workbook = new Workbook(inputPath);
  AssertEx.equal("A2:C2", workbook.Worksheets.get(0).AutoFilter.Range);
  AssertEx.equal(0, workbook.Worksheets.get(0).AutoFilter.FilterColumns.Count);
  AssertEx.equal(0, workbook.Worksheets.get(0).AutoFilter.SortState.SortConditions.Count);

  const temp = new TemporaryDirectory("compat-input-autofilter");
  const roundtripPath = temp.getPath("autofilter-roundtrip-out.xlsx");
  workbook.Save(roundtripPath);

  const reloaded = new Workbook(roundtripPath);
  AssertEx.equal("A2:C2", reloaded.Worksheets.get(0).AutoFilter.Range);
  AssertEx.equal(0, reloaded.Worksheets.get(0).AutoFilter.FilterColumns.Count);
  AssertEx.equal(0, reloaded.Worksheets.get(0).AutoFilter.SortState.SortConditions.Count);

  const workbookXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/workbook.xml");
  const worksheetXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/worksheets/sheet1.xml");
  AssertEx.contains('<autoFilter ref="A2:C2"', worksheetXml);
  AssertEx.contains('name="_xlnm._FilterDatabase"', workbookXml);
  AssertEx.contains("$A$2:$C$2", workbookXml);
}

export function excelInputTableRoundtripPreservesTableStructure(): void {
  const inputPath = resolveRepositoryFile("Input", "table.xlsx");
  const workbook = new Workbook(inputPath);
  const sheet = workbook.Worksheets.get(0);

  AssertEx.equal(1, sheet.ListObjects.Count);
  const table = sheet.ListObjects.get(0);
  AssertEx.equal("Table5", table.DisplayName);
  AssertEx.equal(0, table.StartRow);
  AssertEx.equal(0, table.StartColumn);
  AssertEx.equal(3, table.EndRow);
  AssertEx.equal(2, table.EndColumn);
  AssertEx.true(table.ShowHeaderRow);
  AssertEx.false(table.ShowTotals);
  AssertEx.equal(TableStyleType.TableStyleMedium2, table.TableStyleType);
  AssertEx.true(table.ShowTableStyleRowStripes);
  AssertEx.false(table.ShowTableStyleFirstColumn);
  AssertEx.false(table.ShowTableStyleLastColumn);
  AssertEx.false(table.ShowTableStyleColumnStripes);
  AssertEx.equal(3, table.ListColumns.Count);
  AssertEx.equal("Product", table.ListColumns.get(0).Name);
  AssertEx.equal("Region", table.ListColumns.get(1).Name);
  AssertEx.equal("Revenue", table.ListColumns.get(2).Name);

  const temp = new TemporaryDirectory("compat-input-table");
  const roundtripPath = temp.getPath("table-roundtrip-out.xlsx");
  workbook.Save(roundtripPath);

  const reloaded = new Workbook(roundtripPath);
  const reloadedTable = reloaded.Worksheets.get(0).ListObjects.get(0);
  AssertEx.equal("Table5", reloadedTable.DisplayName);
  AssertEx.equal(0, reloadedTable.StartRow);
  AssertEx.equal(3, reloadedTable.EndRow);
  AssertEx.false(reloadedTable.ShowTotals);
  AssertEx.equal(TableStyleType.TableStyleMedium2, reloadedTable.TableStyleType);
  AssertEx.equal(3, reloadedTable.ListColumns.Count);
  AssertEx.equal("Product", reloadedTable.ListColumns.get(0).Name);
  AssertEx.equal("Revenue", reloadedTable.ListColumns.get(2).Name);

  const tableXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/tables/table1.xml");
  AssertEx.contains('ref="A1:C4"', tableXml);
  AssertEx.contains('totalsRowShown="0"', tableXml);
  AssertEx.contains('name="TableStyleMedium2"', tableXml);
}

export function excelInputPictureRoundtripPreservesPictureStructure(): void {
  const inputPath = resolveRepositoryFile("Input", path.join("compare", "17.6.3-ByCells.xlsx"));
  const workbook = new Workbook(inputPath);
  const sheet = workbook.Worksheets.get(0);

  AssertEx.true(sheet.Pictures.Count > 0, "Expected at least one picture in the input workbook.");
  const picture = sheet.Pictures.get(0);
  AssertEx.true(picture.Data !== null && picture.Data.length > 0, "Picture data must be non-empty.");
  AssertEx.equal(ImageType.Jpeg, picture.ImageType);
  AssertEx.true(picture.UpperLeftRow >= 0, "UpperLeftRow must be non-negative.");
  AssertEx.true(picture.LowerRightRow >= picture.UpperLeftRow, "LowerRightRow must be >= UpperLeftRow.");

  const originalDataLength = picture.Data.length;
  const originalUpperLeftRow = picture.UpperLeftRow;
  const originalUpperLeftColumn = picture.UpperLeftColumn;
  const originalLowerRightRow = picture.LowerRightRow;
  const originalLowerRightColumn = picture.LowerRightColumn;

  const temp = new TemporaryDirectory("compat-input-picture");
  const roundtripPath = temp.getPath("picture-roundtrip-out.xlsx");
  workbook.Save(roundtripPath);

  const reloaded = new Workbook(roundtripPath);
  const reloadedSheet = reloaded.Worksheets.get(0);
  AssertEx.equal(sheet.Pictures.Count, reloadedSheet.Pictures.Count);

  const reloadedPicture = reloadedSheet.Pictures.get(0);
  AssertEx.equal(originalDataLength, reloadedPicture.Data.length);
  AssertEx.equal(originalUpperLeftRow, reloadedPicture.UpperLeftRow);
  AssertEx.equal(originalUpperLeftColumn, reloadedPicture.UpperLeftColumn);
  AssertEx.equal(originalLowerRightRow, reloadedPicture.LowerRightRow);
  AssertEx.equal(originalLowerRightColumn, reloadedPicture.LowerRightColumn);
  AssertEx.equal(ImageType.Jpeg, reloadedPicture.ImageType);

  const drawingXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/drawings/drawing1.xml");
  AssertEx.contains("<xdr:twoCellAnchor", drawingXml);
  AssertEx.contains("<xdr:pic>", drawingXml);
  AssertEx.contains("rId1", drawingXml);
}

export function excelInputShapeRoundtripPreservesShapeStructure(): void {
  const inputPath = resolveRepositoryFile("Input", path.join("compare", "100_rows.xlsx"));
  const workbook = new Workbook(inputPath);
  const sheet = workbook.Worksheets.get(0);

  AssertEx.equal(6, sheet.Shapes.Count);

  const rect = sheet.Shapes.get(0);
  AssertEx.equal("Rectangle 1", rect.Name);
  AssertEx.equal(AutoShapeType.Rectangle, rect.AutoShapeType);
  AssertEx.true(rect.UpperLeftRow >= 0, "UpperLeftRow must be non-negative.");
  AssertEx.true(rect.LowerRightRow >= rect.UpperLeftRow, "LowerRightRow must be >= UpperLeftRow.");

  const arrow = sheet.Shapes.get(2);
  AssertEx.equal(AutoShapeType.RightArrow, arrow.AutoShapeType);

  const temp = new TemporaryDirectory("compat-input-shape");
  const roundtripPath = temp.getPath("shape-roundtrip-out.xlsx");
  workbook.Save(roundtripPath);

  const reloaded = new Workbook(roundtripPath);
  const reloadedSheet = reloaded.Worksheets.get(0);
  AssertEx.equal(6, reloadedSheet.Shapes.Count);
  AssertEx.equal("Rectangle 1", reloadedSheet.Shapes.get(0).Name);
  AssertEx.equal(AutoShapeType.Rectangle, reloadedSheet.Shapes.get(0).AutoShapeType);
  AssertEx.equal(AutoShapeType.RightArrow, reloadedSheet.Shapes.get(2).AutoShapeType);

  const drawingXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/drawings/drawing1.xml");
  AssertEx.contains("<xdr:sp ", drawingXml);
  AssertEx.contains("Rectangle 1", drawingXml);
  AssertEx.contains('prst="rect"', drawingXml);
  AssertEx.contains('prst="rightArrow"', drawingXml);
}

export function excelInputShapeXlsxRoundtripIncludesConnector(): void {
  const inputPath = resolveRepositoryFile("Input", "shape.xlsx");
  const workbook = new Workbook(inputPath);
  const sheet = workbook.Worksheets.get(0);

  AssertEx.equal(3, sheet.Shapes.Count);

  AssertEx.equal("Straight Arrow Connector 2", sheet.Shapes.get(0).Name);
  AssertEx.equal(AutoShapeType.StraightConnector, sheet.Shapes.get(0).AutoShapeType);

  AssertEx.equal("Arrow: Down 3", sheet.Shapes.get(1).Name);
  AssertEx.equal(AutoShapeType.DownArrow, sheet.Shapes.get(1).AutoShapeType);

  AssertEx.equal("Plus Sign 4", sheet.Shapes.get(2).Name);
  AssertEx.equal(AutoShapeType.MathPlus, sheet.Shapes.get(2).AutoShapeType);

  const temp = new TemporaryDirectory("compat-input-shape-xlsx");
  const roundtripPath = temp.getPath("shape-xlsx-roundtrip-out.xlsx");
  workbook.Save(roundtripPath);

  const reloaded = new Workbook(roundtripPath);
  const reloadedSheet = reloaded.Worksheets.get(0);
  AssertEx.equal(3, reloadedSheet.Shapes.Count);
  AssertEx.equal("Straight Arrow Connector 2", reloadedSheet.Shapes.get(0).Name);
  AssertEx.equal("Arrow: Down 3", reloadedSheet.Shapes.get(1).Name);
  AssertEx.equal("Plus Sign 4", reloadedSheet.Shapes.get(2).Name);

  const drawingXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/drawings/drawing1.xml");
  AssertEx.contains("<xdr:cxnSp", drawingXml);
  AssertEx.contains("<xdr:sp ", drawingXml);
  AssertEx.contains("downArrow", drawingXml);
  AssertEx.contains("mathPlus", drawingXml);
}

export function excelInput050RoundtripHasNoDanglingChartRelationships(): void {
  const inputPath = resolveRepositoryFile("Input", path.join("roundtrip", "050.xlsx"));
  const workbook = new Workbook(inputPath);
  const activeSheet = workbook.Worksheets.get(workbook.Worksheets.ActiveSheetIndex);
  AssertEx.false(activeSheet.Protection.IsProtected);

  const temp = new TemporaryDirectory("compat-input-050");
  const roundtripPath = temp.getPath("050-roundtrip-out.xlsx");
  workbook.Save(roundtripPath);

  const reloaded = new Workbook(roundtripPath);
  const reloadedActiveSheet = reloaded.Worksheets.get(reloaded.Worksheets.ActiveSheetIndex);
  AssertEx.false(reloadedActiveSheet.Protection.IsProtected);

  const drawingXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/drawings/drawing1.xml");
  const drawingRoot = documentRoot(parseXml(drawingXml));
  AssertEx.notNull(drawingRoot, "drawing1.xml must have a root element.");

  const referencedRelationshipIds = Array.from(
    new Set(
      descendants(drawingRoot, "chart")
        .map((element) => attribute(element, "id"))
        .filter((id): id is string => id !== null && id.trim().length > 0),
    ),
  );
  AssertEx.equal(3, referencedRelationshipIds.length);

  const drawingRelsXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/drawings/_rels/drawing1.xml.rels");
  const drawingRelsRoot = documentRoot(parseXml(drawingRelsXml));
  AssertEx.notNull(drawingRelsRoot, "drawing1.xml.rels must have a root element.");

  const relationshipsById = new Map<string, XmlNode>();
  for (const relationship of elementChildren(drawingRelsRoot, "Relationship")) {
    const id = attribute(relationship, "Id");
    if (id !== null && id.trim().length > 0) {
      relationshipsById.set(id, relationship);
    }
  }

  for (const relationshipId of referencedRelationshipIds) {
    const relationship = relationshipsById.get(relationshipId);
    AssertEx.notNull(relationship, "Missing drawing relationship for chart id: " + relationshipId);

    const target = attribute(relationship!, "Target") ?? "";
    AssertEx.true(target.startsWith("../charts/"), "Chart relationship target must point to xl/charts: " + target);

    const normalizedTarget = "xl/" + target.substring(3).split("\\").join("/");
    AssertEx.true(
      ZipPackageHelper.entryExists(roundtripPath, normalizedTarget),
      "Missing chart part referenced by drawing relationship: " + normalizedTarget,
    );
  }

  const worksheetXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/worksheets/sheet1.xml");
  const worksheetRoot = documentRoot(parseXml(worksheetXml));
  AssertEx.notNull(worksheetRoot, "sheet1.xml must have a root element.");
  const sheetProtection = firstElement(worksheetRoot, "sheetProtection");
  AssertEx.notNull(sheetProtection, "sheetProtection element must be preserved for 050.xlsx.");
  AssertEx.null(
    attribute(sheetProtection!, "sheet"),
    "sheetProtection@sheet must remain absent when the source worksheet is editable.",
  );
}

export function excelInput091RoundtripPreservesChartFocusedSheetView(): void {
  const inputPath = resolveRepositoryFile(
    "Input",
    path.join("roundtrip", "091-TC-LAB02_data_003_saleh_demo.xlsx"),
  );
  const workbook = new Workbook(inputPath);

  const temp = new TemporaryDirectory("compat-input-091");
  const roundtripPath = temp.getPath("091-roundtrip-out.xlsx");
  workbook.Save(roundtripPath);

  const worksheetXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/worksheets/sheet4.xml");
  const worksheetRoot = documentRoot(parseXml(worksheetXml));
  AssertEx.notNull(worksheetRoot, "sheet4.xml must have a root element.");
  const sheetViews = firstElement(worksheetRoot, "sheetViews");
  AssertEx.notNull(sheetViews, "sheet4 must preserve its sheetViews.");
  const sheetView = firstElement(sheetViews, "sheetView");
  AssertEx.notNull(sheetView, "sheet4 must preserve its sheetView.");
  AssertEx.equal("pageLayout", attribute(sheetView!, "view"));
  AssertEx.equal("X1", attribute(sheetView!, "topLeftCell"));
  AssertEx.equal("1", attribute(sheetView!, "tabSelected"));

  const selection = firstElement(sheetView, "selection");
  AssertEx.notNull(selection, "sheet4 must preserve its primary selection.");
  AssertEx.equal("AI12", attribute(selection!, "activeCell"));
  AssertEx.equal("AI12", attribute(selection!, "sqref"));
}

export function excelInput1986ChartRoundtripNormalizesDrawingMarkup(): void {
  const inputPath = resolveRepositoryFile(
    "Input",
    path.join("roundtrip", "1986e3b4-181c-46b1-ab64-6f01ad8c3cf4_chart.xlsx"),
  );
  const workbook = new Workbook(inputPath);

  const temp = new TemporaryDirectory("compat-input-1986-chart");
  const roundtripPath = temp.getPath("1986-chart-roundtrip-out.xlsx");
  workbook.Save(roundtripPath);

  const drawingXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/drawings/drawing1.xml");
  const drawingRoot = documentRoot(parseXml(drawingXml));
  AssertEx.notNull(drawingRoot, "drawing1.xml must have a root element.");
  const anchors = elementChildren(drawingRoot, "twoCellAnchor");
  AssertEx.true(anchors.length > 1, "drawing1.xml must contain at least two anchors.");

  const fallbackShapeProperties = childPath(anchors[1], "AlternateContent", "Fallback", "sp", "spPr");
  AssertEx.notNull(fallbackShapeProperties, "Fallback shape properties must be present.");
  AssertEx.notNull(
    firstQualifiedChild(fallbackShapeProperties, "a:xfrm"),
    "Fallback shape properties must use a:xfrm.",
  );
  AssertEx.null(
    firstQualifiedChild(fallbackShapeProperties, "xdr:xfrm"),
    "Fallback shape properties must not use xdr:xfrm.",
  );

  for (const creationId of descendants(drawingRoot, "creationId")) {
    const id = attribute(creationId, "id") ?? "";
    AssertEx.equal(id.toUpperCase(), id, "creationId GUIDs must be normalized to uppercase.");
  }
}

export function excelInput222RoundtripPreservesColumnStyleWithoutBorders(): void {
  const inputPath = resolveRepositoryFile("Input", path.join("roundtrip", "222.xlsx"));
  const workbook = new Workbook(inputPath);

  const temp = new TemporaryDirectory("compat-input-222");
  const roundtripPath = temp.getPath("222-roundtrip-out.xlsx");
  workbook.Save(roundtripPath);

  const worksheetXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/worksheets/sheet1.xml");
  const worksheetRoot = documentRoot(parseXml(worksheetXml));
  AssertEx.notNull(worksheetRoot, "sheet1.xml must have a root element.");
  const stylesXml = ZipPackageHelper.readEntryText(roundtripPath, "xl/styles.xml");
  const stylesRoot = documentRoot(parseXml(stylesXml));
  AssertEx.notNull(stylesRoot, "styles.xml must have a root element.");

  const cols = firstElement(worksheetRoot, "cols");
  AssertEx.notNull(cols, "sheet1 must preserve its cols element.");
  const firstColumn = firstElement(cols, "col");
  AssertEx.notNull(firstColumn, "sheet1 must preserve its first column metadata.");

  const styleIndex = attribute(firstColumn!, "style");
  AssertEx.equal("1", styleIndex);

  const cellXfs = firstElement(stylesRoot, "cellXfs");
  AssertEx.notNull(cellXfs, "styles.xml must contain cellXfs.");
  const formattingRecords = elementChildren(cellXfs, "xf");
  AssertEx.true(formattingRecords.length > 1, "styles.xml must contain at least two cell formats.");
  const columnStyle = formattingRecords[1];
  AssertEx.equal("0", attribute(columnStyle, "borderId"));
  AssertEx.null(attribute(columnStyle, "applyBorder"), "Column style must not apply borders after roundtrip.");
}

export function excelInputRoundtripSetsActiveSheetA1ToHelloWorld(): void {
  const inputDir = resolveRepositoryFile("Input", "roundtrip");
  const files = listXlsxFiles(inputDir);
  const outputDir = getOrCreateOutputDirectory("Roundtrip");

  for (const filePath of files) {
    const fileName = path.basename(filePath);
    const workbook = new Workbook(filePath);
    const sheet = workbook.Worksheets.get(workbook.Worksheets.ActiveSheetIndex);
    sheet.Cells.get("A1").PutValue("hello,world!");

    const outputPath = path.join(outputDir, appendOutSuffix(fileName));
    workbook.Save(outputPath);

    const reloaded = new Workbook(outputPath);
    const reloadedSheet = reloaded.Worksheets.get(reloaded.Worksheets.ActiveSheetIndex);
    AssertEx.equal(
      "hello,world!",
      reloadedSheet.Cells.get("A1").StringValue,
      "A1 value must be 'hello,world!' after round-trip for: " + fileName,
    );
  }
}
