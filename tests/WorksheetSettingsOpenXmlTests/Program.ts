import { Color, VisibilityType, Workbook } from "../../aspose_cells";
import {
  attribute,
  descendants,
  documentRoot,
  elementChildren,
  firstElement,
  XmlNode,
} from "../../aspose_cells/Xml/XmlReader";
import { AssertEx, TemporaryDirectory, TestCase, TestRunner } from "../Shared/TestInfrastructure";
import { WorksheetScenarioFactory } from "../Shared/ScenarioFactories";
import { AutoFilterScenarioFactory } from "../Shared/AutoFilterScenarioFactory";
import { OpenXmlPackageBuilder, OpenXmlWorkbookReader } from "../Shared/OpenXmlPackage";

const STYLES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const STYLES_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

class RowSnapshot {
  public constructor(
    public readonly Height: number | null,
    public readonly Hidden: boolean,
  ) {}
}

class ColumnSnapshot {
  public constructor(
    public readonly MinColumnIndex: number,
    public readonly MaxColumnIndex: number,
    public readonly Width: number | null,
    public readonly Hidden: boolean,
  ) {}
}

class WorksheetSnapshot {
  public constructor(
    public readonly Name: string,
    public readonly Visibility: string,
    public readonly Dimension: string,
    public readonly TabColor: string,
    public readonly ShowGridLines: boolean,
    public readonly ShowRowColumnHeaders: boolean,
    public readonly ShowZeros: boolean,
    public readonly RightToLeft: boolean,
    public readonly Zoom: number,
    public readonly IsProtected: boolean,
    public readonly ProtectObjects: boolean,
    public readonly ProtectScenarios: boolean,
    public readonly ProtectFormatCells: boolean,
    public readonly ProtectInsertRows: boolean,
    public readonly ProtectAutoFilter: boolean,
    public readonly ProtectSelectLockedCells: boolean,
    public readonly ProtectSelectUnlockedCells: boolean,
    public readonly Rows: Map<number, RowSnapshot>,
    public readonly Columns: readonly ColumnSnapshot[],
    public readonly MergeReferences: readonly string[],
    public readonly Cells: Map<string, string>,
  ) {}
}

class OpenXmlSheetDefinition {
  public constructor(
    public readonly Name: string,
    public readonly State: string | null,
    public readonly WorksheetXml: string,
  ) {}
}

function worksheetVisibilityMatchesOpenXmlSdk(): void {
  const temp = new TemporaryDirectory("worksheet-openxml-visibility");
  const path = temp.getPath("visibility.xlsx");
  const workbook = new Workbook();
  workbook.Worksheets[0].Name = "Visible";
  const hiddenIndex = workbook.Worksheets.Add("Hidden");
  const veryHiddenIndex = workbook.Worksheets.Add("VeryHidden");
  workbook.Worksheets[hiddenIndex].VisibilityType = VisibilityType.Hidden;
  workbook.Worksheets[veryHiddenIndex].VisibilityType = VisibilityType.VeryHidden;
  workbook.Save(path);

  const loaded = new Workbook(path);
  AssertEx.equal(VisibilityType.Visible, loaded.Worksheets["Visible"].VisibilityType);
  AssertEx.equal(VisibilityType.Hidden, loaded.Worksheets["Hidden"].VisibilityType);
  AssertEx.equal(VisibilityType.VeryHidden, loaded.Worksheets["VeryHidden"].VisibilityType);

  const snapshots = readWorkbookSnapshot(path);
  AssertEx.equal(3, snapshots.length);
  AssertEx.equal("Visible", snapshots[0].Name);
  AssertEx.equal("Visible", snapshots[0].Visibility);
  AssertEx.equal("Hidden", snapshots[1].Name);
  AssertEx.equal("Hidden", snapshots[1].Visibility);
  AssertEx.equal("VeryHidden", snapshots[2].Name);
  AssertEx.equal("VeryHidden", snapshots[2].Visibility);
}

function worksheetSettingsAndValuesMatchOpenXmlSdk(): void {
  const temp = new TemporaryDirectory("worksheet-openxml-settings");
  const path = temp.getPath("worksheet-settings.xlsx");
  const workbook = WorksheetScenarioFactory.CreateWorksheetSettingsWorkbook();
  workbook.Save(path);

  const loaded = new Workbook(path);
  WorksheetScenarioFactory.AssertWorksheetSettings(loaded);

  const snapshots = readWorkbookSnapshot(path);
  AssertEx.equal(2, snapshots.length);
  const snapshot = snapshots.filter((item) => item.Name === "Layout")[0];
  AssertEx.equal("Layout", snapshot.Name);
  AssertEx.equal("Hidden", snapshot.Visibility);
  AssertEx.equal("A1:C4", snapshot.Dimension);
  AssertEx.equal("Merged", snapshot.Cells.get("A1"));
  AssertEx.equal("99", snapshot.Cells.get("C4"));
  assertRow(snapshot, 1, 22.5, false);
  assertRow(snapshot, 3, null, true);
  assertColumn(snapshot, 0, 0, 18.25, false);
  assertColumn(snapshot, 2, 2, null, true);
  AssertEx.equal(1, snapshot.MergeReferences.length);
  AssertEx.equal("A1:B2", snapshot.MergeReferences[0]);

  const layoutSheet = loaded.Worksheets["Layout"];
  AssertEx.notNull(layoutSheet);
  AssertEx.equal(layoutSheet.Cells["A1"].StringValue, snapshot.Cells.get("A1"));
  AssertEx.equal(layoutSheet.Cells["C4"].StringValue, snapshot.Cells.get("C4"));
  AssertEx.equal(layoutSheet.Cells.Rows[1].Height ?? 0, snapshot.Rows.get(1)!.Height ?? 0);
  AssertEx.equal(layoutSheet.Cells.Rows[3].IsHidden, snapshot.Rows.get(3)!.Hidden);
  AssertEx.equal(layoutSheet.Cells.Columns[0].Width ?? 0, snapshot.Columns[0].Width ?? 0);
  AssertEx.equal(layoutSheet.Cells.Columns[2].IsHidden, snapshot.Columns[1].Hidden);
}

function worksheetViewAndTabColorMatchOpenXmlSdk(): void {
  const temp = new TemporaryDirectory("worksheet-openxml-view");
  const path = temp.getPath("worksheet-view.xlsx");
  const workbook = WorksheetScenarioFactory.CreateWorksheetSettingsWorkbook();
  workbook.Save(path);

  const loaded = new Workbook(path);
  WorksheetScenarioFactory.AssertWorksheetSettings(loaded);

  const snapshots = readWorkbookSnapshot(path);
  AssertEx.equal(2, snapshots.length);
  const snapshot = snapshots.filter((item) => item.Name === "Layout")[0];
  AssertEx.equal("FF224466", snapshot.TabColor);
  AssertEx.false(snapshot.ShowGridLines);
  AssertEx.false(snapshot.ShowRowColumnHeaders);
  AssertEx.false(snapshot.ShowZeros);
  AssertEx.true(snapshot.RightToLeft);
  AssertEx.equal(85, snapshot.Zoom);
}

function worksheetProtectionMatchOpenXmlSdk(): void {
  const temp = new TemporaryDirectory("worksheet-openxml-protection");
  const path = temp.getPath("worksheet-protection.xlsx");
  const workbook = WorksheetScenarioFactory.CreateWorksheetSettingsWorkbook();
  workbook.Save(path);

  const loaded = new Workbook(path);
  WorksheetScenarioFactory.AssertWorksheetSettings(loaded);

  const snapshots = readWorkbookSnapshot(path);
  AssertEx.equal(2, snapshots.length);
  const snapshot = snapshots.filter((item) => item.Name === "Layout")[0];
  AssertEx.true(snapshot.IsProtected);
  AssertEx.true(snapshot.ProtectObjects);
  AssertEx.true(snapshot.ProtectScenarios);
  AssertEx.true(snapshot.ProtectFormatCells);
  AssertEx.true(snapshot.ProtectInsertRows);
  AssertEx.true(snapshot.ProtectAutoFilter);
  AssertEx.true(snapshot.ProtectSelectLockedCells);
  AssertEx.true(snapshot.ProtectSelectUnlockedCells);
}

function autoFilterMatchesOpenXmlSdk(): void {
  const temp = new TemporaryDirectory("worksheet-openxml-autofilter");
  const path = temp.getPath("autofilter.xlsx");

  const workbook = AutoFilterScenarioFactory.CreateAutoFilterWorkbook();
  workbook.Save(path);

  const loaded = new Workbook(path);
  AutoFilterScenarioFactory.AssertAutoFilter(loaded);

  const reader = new OpenXmlWorkbookReader(path);
  const worksheetXml = reader.sheetXmlByIndex(0);
  const root = documentRoot(worksheetXml);
  const autoFilter = firstElement(root, "autoFilter");
  AssertEx.notNull(autoFilter);
  AssertEx.equal("A1:E6", (attribute(autoFilter, "ref") ?? "").toUpperCase());

  const filterColumns = elementChildren(autoFilter, "filterColumn");
  AssertEx.equal(5, filterColumns.length);
  AssertEx.equal("1", attribute(filterColumns[0], "hiddenButton") ?? "");
  AssertEx.equal(
    "Open",
    attribute(elementChildren(firstElement(filterColumns[0], "filters"), "filter")[0], "val") ?? "",
  );
  AssertEx.equal(
    "greaterThanOrEqual",
    attribute(elementChildren(firstElement(filterColumns[1], "customFilters"), "customFilter")[0], "operator") ?? "",
  );
  AssertEx.equal("3", attribute(firstElement(filterColumns[2], "colorFilter"), "dxfId") ?? "");
  AssertEx.equal("thisMonth", attribute(firstElement(filterColumns[3], "dynamicFilter"), "type") ?? "");
  AssertEx.equal("0", attribute(firstElement(filterColumns[4], "top10"), "top") ?? "");

  const sortState = firstElement(autoFilter, "sortState");
  AssertEx.equal("A2:E6", (attribute(sortState, "ref") ?? "").toUpperCase());
  const sortConditions = elementChildren(sortState, "sortCondition");
  AssertEx.equal(3, sortConditions.length);
  AssertEx.equal("value", attribute(sortConditions[0], "sortBy") ?? "");
  AssertEx.equal("High,Medium,Low", attribute(sortConditions[0], "customList") ?? "");
  AssertEx.equal("cellColor", attribute(sortConditions[1], "sortBy") ?? "");
  AssertEx.equal("4", attribute(sortConditions[1], "dxfId") ?? "");
  AssertEx.equal("icon", attribute(sortConditions[2], "sortBy") ?? "");
  AssertEx.equal("3TrafficLights1", attribute(sortConditions[2], "iconSet") ?? "");
  AssertEx.equal("2", attribute(sortConditions[2], "iconId") ?? "");
}

function adjacentColumnsAreCombinedAndReadConsistently(): void {
  const temp = new TemporaryDirectory("worksheet-openxml-columns");
  const path = temp.getPath("columns.xlsx");
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  sheet.Name = "Columns";
  sheet.Cells.Columns[0].Width = 15.5;
  sheet.Cells.Columns[1].Width = 15.5;
  sheet.Cells.Columns[3].IsHidden = true;
  sheet.Cells["A1"].PutValue("A");
  sheet.Cells["B1"].PutValue("B");
  workbook.Save(path);

  const loaded = new Workbook(path);
  AssertEx.equal(15.5, loaded.Worksheets[0].Cells.Columns[0].Width ?? 0);
  AssertEx.equal(15.5, loaded.Worksheets[0].Cells.Columns[1].Width ?? 0);
  AssertEx.true(loaded.Worksheets[0].Cells.Columns[3].IsHidden);

  const snapshot = readWorkbookSnapshot(path)[0];
  AssertEx.equal(2, snapshot.Columns.length);
  assertColumn(snapshot, 0, 1, 15.5, false);
  assertColumn(snapshot, 3, 3, null, true);
  AssertEx.equal("A", snapshot.Cells.get("A1"));
  AssertEx.equal("B", snapshot.Cells.get("B1"));
}

function mergeOnlyDimensionMatchesOpenXmlSdk(): void {
  const temp = new TemporaryDirectory("worksheet-openxml-merge-only");
  const path = temp.getPath("merge-only.xlsx");
  const workbook = new Workbook();
  workbook.Worksheets[0].Name = "MergeOnly";
  workbook.Worksheets[0].Cells.Merge(4, 5, 2, 2);
  workbook.Save(path);

  const loaded = new Workbook(path);
  AssertEx.equal(1, loaded.Worksheets[0].Cells.MergedCells.length);
  AssertEx.equal(4, loaded.Worksheets[0].Cells.MergedCells[0].StartRow);
  AssertEx.equal(5, loaded.Worksheets[0].Cells.MergedCells[0].StartColumn);

  const snapshot = readWorkbookSnapshot(path)[0];
  AssertEx.equal("F5:G6", snapshot.Dimension);
  AssertEx.equal(1, snapshot.MergeReferences.length);
  AssertEx.equal("F5:G6", snapshot.MergeReferences[0]);
}

function openXmlCreatedVisibilityLoadsInLibrary(): void {
  const temp = new TemporaryDirectory("worksheet-library-read-visibility");
  const path = temp.getPath("openxml-visibility.xlsx");
  createOpenXmlWorkbook(path, [
    new OpenXmlSheetDefinition("Visible", null, minimalWorksheetXml()),
    new OpenXmlSheetDefinition("Hidden", "hidden", minimalWorksheetXml()),
    new OpenXmlSheetDefinition("VeryHidden", "veryHidden", minimalWorksheetXml()),
  ]);

  const loaded = new Workbook(path);
  AssertEx.equal(VisibilityType.Visible, loaded.Worksheets["Visible"].VisibilityType);
  AssertEx.equal(VisibilityType.Hidden, loaded.Worksheets["Hidden"].VisibilityType);
  AssertEx.equal(VisibilityType.VeryHidden, loaded.Worksheets["VeryHidden"].VisibilityType);

  const snapshot = readWorkbookSnapshot(path);
  AssertEx.equal(3, snapshot.length);
  AssertEx.equal("Hidden", snapshot[1].Visibility);
  AssertEx.equal("VeryHidden", snapshot[2].Visibility);
}

function openXmlCreatedSettingsAndValuesLoadInLibrary(): void {
  const temp = new TemporaryDirectory("worksheet-library-read-settings");
  const path = temp.getPath("openxml-settings.xlsx");
  createOpenXmlWorkbook(path, [
    new OpenXmlSheetDefinition("Visible", null, minimalWorksheetXml()),
    new OpenXmlSheetDefinition(
      "Layout",
      "hidden",
      '<?xml version="1.0" encoding="utf-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><tabColor rgb="FF224466"/></sheetPr><dimension ref="A1:C4"/><sheetViews><sheetView workbookViewId="0" showGridLines="0" showRowColHeaders="0" showZeros="0" rightToLeft="1" zoomScale="85"/></sheetViews><cols><col min="1" max="1" width="18.25" customWidth="1"/><col min="3" max="3" hidden="1"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Merged</t></is></c></row><row r="2" ht="22.5" customHeight="1"/><row r="4" hidden="1"><c r="C4"><v>99</v></c></row></sheetData><sheetProtection sheet="1" objects="1" scenarios="1" formatCells="1" insertRows="1" autoFilter="1" selectLockedCells="1" selectUnlockedCells="1"/><mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells></worksheet>',
    ),
  ]);

  const loaded = new Workbook(path);
  WorksheetScenarioFactory.AssertWorksheetSettings(loaded);

  const snapshot = readWorkbookSnapshot(path).filter((item) => item.Name === "Layout")[0];
  AssertEx.equal("Merged", snapshot.Cells.get("A1"));
  AssertEx.equal("99", snapshot.Cells.get("C4"));
}

function openXmlCreatedSheetViewLoadsInLibrary(): void {
  const temp = new TemporaryDirectory("worksheet-library-read-view");
  const path = temp.getPath("openxml-view.xlsx");
  createOpenXmlWorkbook(path, [
    new OpenXmlSheetDefinition(
      "ViewSheet",
      null,
      '<?xml version="1.0" encoding="utf-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><tabColor rgb="FF0A141E"/></sheetPr><sheetViews><sheetView workbookViewId="0" showGridLines="0" showRowColHeaders="0" showZeros="0" rightToLeft="1" zoomScale="140"/></sheetViews><sheetData/></worksheet>',
    ),
  ]);

  const loaded = new Workbook(path);
  const sheet = loaded.Worksheets[0];
  AssertEx.equal(Color.FromArgb(255, 10, 20, 30), sheet.TabColor);
  AssertEx.false(sheet.ShowGridlines);
  AssertEx.false(sheet.ShowRowColumnHeaders);
  AssertEx.false(sheet.ShowZeros);
  AssertEx.true(sheet.RightToLeft);
  AssertEx.equal(140, sheet.Zoom);
}

function openXmlCreatedSheetProtectionLoadsInLibrary(): void {
  const temp = new TemporaryDirectory("worksheet-library-read-protection");
  const path = temp.getPath("openxml-protection.xlsx");

  createOpenXmlWorkbook(path, [
    new OpenXmlSheetDefinition(
      "ProtectedSheet",
      null,
      '<?xml version="1.0" encoding="utf-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><sheetProtection sheet="1" objects="1" scenarios="1" formatCells="1" insertRows="1" autoFilter="1" selectLockedCells="1" selectUnlockedCells="1"/></worksheet>',
    ),
  ]);

  const loaded = new Workbook(path);
  const sheet = loaded.Worksheets[0];
  AssertEx.true(sheet.Protection.IsProtected);
  AssertEx.true(sheet.Protection.Objects);
  AssertEx.true(sheet.Protection.Scenarios);
  AssertEx.true(sheet.Protection.FormatCells);
  AssertEx.true(sheet.Protection.InsertRows);
  AssertEx.true(sheet.Protection.AutoFilter);
  AssertEx.true(sheet.Protection.SelectLockedCells);
  AssertEx.true(sheet.Protection.SelectUnlockedCells);
}

function openXmlCreatedAutoFilterLoadsInLibrary(): void {
  const temp = new TemporaryDirectory("worksheet-library-read-autofilter");
  const path = temp.getPath("openxml-autofilter.xlsx");

  createOpenXmlWorkbook(
    path,
    [
      new OpenXmlSheetDefinition(
        "Filtered",
        null,
        '<?xml version="1.0" encoding="utf-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:E6"/><sheetData/><autoFilter ref="A1:E6"><filterColumn colId="0" hiddenButton="1"><filters><filter val="Open"/><filter val="Closed"/></filters></filterColumn><filterColumn colId="1"><customFilters and="1"><customFilter operator="greaterThanOrEqual" val="10"/><customFilter operator="lessThanOrEqual" val="50"/></customFilters></filterColumn><filterColumn colId="2"><colorFilter dxfId="3" cellColor="1"/></filterColumn><filterColumn colId="3"><dynamicFilter type="thisMonth" val="1" maxVal="31"/></filterColumn><filterColumn colId="4"><top10 top="0" percent="1" val="10" filterVal="2.5"/></filterColumn><sortState ref="A2:E6" caseSensitive="1" sortMethod="pinYin"><sortCondition ref="B2:B6" descending="1" sortBy="value" customList="High,Medium,Low"/><sortCondition ref="C2:C6" sortBy="cellColor" dxfId="4"/><sortCondition ref="E2:E6" sortBy="icon" iconSet="3TrafficLights1" iconId="2"/></sortState></autoFilter></worksheet>',
      ),
    ],
    [
      {
        partPath: "xl/styles.xml",
        xml: '<?xml version="1.0" encoding="utf-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="5"><dxf/><dxf/><dxf/><dxf/><dxf/></dxfs></styleSheet>',
        contentType: STYLES_CONTENT_TYPE,
        relationshipType: STYLES_RELATIONSHIP,
        relationshipId: "rId2",
      },
    ],
  );

  const loaded = new Workbook(path);
  AutoFilterScenarioFactory.AssertAutoFilter(loaded);
}

function openXmlCreatedColumnsLoadInLibrary(): void {
  const temp = new TemporaryDirectory("worksheet-library-read-columns");
  const path = temp.getPath("openxml-columns.xlsx");
  createOpenXmlWorkbook(path, [
    new OpenXmlSheetDefinition(
      "Columns",
      null,
      '<?xml version="1.0" encoding="utf-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B1"/><cols><col min="1" max="2" width="15.5" customWidth="1"/><col min="4" max="4" hidden="1"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>A</t></is></c><c r="B1" t="inlineStr"><is><t>B</t></is></c></row></sheetData></worksheet>',
    ),
  ]);

  const loaded = new Workbook(path);
  AssertEx.equal(15.5, loaded.Worksheets[0].Cells.Columns[0].Width ?? 0);
  AssertEx.equal(15.5, loaded.Worksheets[0].Cells.Columns[1].Width ?? 0);
  AssertEx.true(loaded.Worksheets[0].Cells.Columns[3].IsHidden);
  AssertEx.equal("A", loaded.Worksheets[0].Cells["A1"].StringValue);
  AssertEx.equal("B", loaded.Worksheets[0].Cells["B1"].StringValue);
}

function openXmlCreatedMergeOnlyDimensionLoadsInLibrary(): void {
  const temp = new TemporaryDirectory("worksheet-library-read-merge-only");
  const path = temp.getPath("openxml-merge-only.xlsx");
  createOpenXmlWorkbook(path, [
    new OpenXmlSheetDefinition(
      "MergeOnly",
      null,
      '<?xml version="1.0" encoding="utf-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="F5:G6"/><sheetData/><mergeCells count="1"><mergeCell ref="F5:G6"/></mergeCells></worksheet>',
    ),
  ]);

  const loaded = new Workbook(path);
  AssertEx.equal(1, loaded.Worksheets[0].Cells.MergedCells.length);
  AssertEx.equal(4, loaded.Worksheets[0].Cells.MergedCells[0].StartRow);
  AssertEx.equal(5, loaded.Worksheets[0].Cells.MergedCells[0].StartColumn);

  const snapshot = readWorkbookSnapshot(path)[0];
  AssertEx.equal("F5:G6", snapshot.Dimension);
  AssertEx.equal("F5:G6", snapshot.MergeReferences[0]);
}

function assertRow(snapshot: WorksheetSnapshot, rowIndex: number, height: number | null, hidden: boolean): void {
  const row = snapshot.Rows.get(rowIndex);
  AssertEx.notNull(row, "Expected row " + rowIndex + " to exist in Open XML snapshot.");
  AssertEx.equal(height, row!.Height);
  AssertEx.equal(hidden, row!.Hidden);
}

function assertColumn(
  snapshot: WorksheetSnapshot,
  minColumnIndex: number,
  maxColumnIndex: number,
  width: number | null,
  hidden: boolean,
): void {
  const column = snapshot.Columns.filter(
    (item) => item.MinColumnIndex === minColumnIndex && item.MaxColumnIndex === maxColumnIndex,
  )[0];
  AssertEx.notNull(
    column,
    "Expected column range " + minColumnIndex + ":" + maxColumnIndex + " to exist in Open XML snapshot.",
  );
  AssertEx.equal(width, column!.Width);
  AssertEx.equal(hidden, column!.Hidden);
}

function readWorkbookSnapshot(workbookPath: string): WorksheetSnapshot[] {
  const reader = new OpenXmlWorkbookReader(workbookPath);
  const sharedStrings = loadSharedStrings(reader);
  const snapshots: WorksheetSnapshot[] = [];

  for (const sheet of elementChildren(firstElement(reader.workbookRoot, "sheets"), "sheet")) {
    const relationshipId = attribute(sheet, "id");
    if (relationshipId === null || relationshipId.trim().length === 0) {
      continue;
    }

    const sheetName = attribute(sheet, "name") ?? "";
    snapshots.push(readWorksheetSnapshot(reader, sheet, sheetName, sharedStrings));
  }

  return snapshots;
}

function readWorksheetSnapshot(
  reader: OpenXmlWorkbookReader,
  sheetElement: XmlNode,
  sheetName: string,
  sharedStrings: readonly string[],
): WorksheetSnapshot {
  const worksheetXml = reader.sheetXmlByName(sheetName);
  const worksheetRoot = documentRoot(worksheetXml);
  if (worksheetRoot === null) {
    throw new Error("Worksheet XML is empty.");
  }

  let visibility: string;
  switch (attribute(sheetElement, "state")) {
    case "hidden":
      visibility = "Hidden";
      break;
    case "veryHidden":
      visibility = "VeryHidden";
      break;
    default:
      visibility = "Visible";
      break;
  }

  const rows = new Map<number, RowSnapshot>();
  for (const row of elementChildren(firstElement(worksheetRoot, "sheetData"), "row")) {
    const rowIndex = tryParseInt(attribute(row, "r"));
    if (rowIndex === null) {
      continue;
    }

    rows.set(rowIndex - 1, new RowSnapshot(tryParseDouble(attribute(row, "ht")), parseBool(attribute(row, "hidden"))));
  }

  const columns: ColumnSnapshot[] = [];
  for (const column of elementChildren(firstElement(worksheetRoot, "cols"), "col")) {
    const minColumnIndex = tryParseInt(attribute(column, "min"));
    const maxColumnIndex = tryParseInt(attribute(column, "max"));
    if (minColumnIndex === null || maxColumnIndex === null) {
      continue;
    }

    columns.push(
      new ColumnSnapshot(
        minColumnIndex - 1,
        maxColumnIndex - 1,
        tryParseDouble(attribute(column, "width")),
        parseBool(attribute(column, "hidden")),
      ),
    );
  }

  const merges: string[] = [];
  for (const mergeCell of elementChildren(firstElement(worksheetRoot, "mergeCells"), "mergeCell")) {
    merges.push((attribute(mergeCell, "ref") ?? "").toUpperCase());
  }

  const cells = new Map<string, string>();
  for (const cell of descendants(worksheetRoot, "c")) {
    const cellReference = (attribute(cell, "r") ?? "").toUpperCase();
    if (cellReference.trim().length === 0) {
      continue;
    }

    cells.set(cellReference, readCellValue(cell, sharedStrings));
  }

  const sheetView = firstElement(firstElement(worksheetRoot, "sheetViews"), "sheetView");
  const sheetProtection = firstElement(worksheetRoot, "sheetProtection");
  return new WorksheetSnapshot(
    sheetName,
    visibility,
    attribute(firstElement(worksheetRoot, "dimension"), "ref") ?? "",
    (attribute(firstElement(firstElement(worksheetRoot, "sheetPr"), "tabColor"), "rgb") ?? "").toUpperCase(),
    parseViewBool(attribute(sheetView, "showGridLines"), true),
    parseViewBool(attribute(sheetView, "showRowColHeaders"), true),
    parseViewBool(attribute(sheetView, "showZeros"), true),
    parseViewBool(attribute(sheetView, "rightToLeft"), false),
    tryParseInt(attribute(sheetView, "zoomScale")) ?? 100,
    parseProtectionBool(attribute(sheetProtection, "sheet"), sheetProtection !== null),
    parseProtectionBool(attribute(sheetProtection, "objects"), false),
    parseProtectionBool(attribute(sheetProtection, "scenarios"), false),
    parseProtectionBool(attribute(sheetProtection, "formatCells"), false),
    parseProtectionBool(attribute(sheetProtection, "insertRows"), false),
    parseProtectionBool(attribute(sheetProtection, "autoFilter"), false),
    parseProtectionBool(attribute(sheetProtection, "selectLockedCells"), false),
    parseProtectionBool(attribute(sheetProtection, "selectUnlockedCells"), false),
    rows,
    columns,
    merges,
    cells,
  );
}

function createOpenXmlWorkbook(
  path: string,
  sheets: readonly OpenXmlSheetDefinition[],
  extraParts: readonly {
    partPath: string;
    xml: string;
    contentType?: string;
    relationshipType?: string;
    relationshipId?: string;
  }[] = [],
): void {
  const builder = new OpenXmlPackageBuilder();
  const sheetElements: string[] = [];

  for (let index = 0; index < sheets.length; index++) {
    const sheet = sheets[index];
    const sheetNumber = index + 1;
    const stateAttribute =
      sheet.State === null || sheet.State.trim().length === 0 ? "" : ` state="${sheet.State}"`;
    sheetElements.push(
      `<sheet name="${sheet.Name}" sheetId="${sheetNumber}" r:id="rId${sheetNumber}"${stateAttribute}/>`,
    );
    builder.addWorksheet(sheet.WorksheetXml, `rId${sheetNumber}`);
  }

  builder.setWorkbookXml(
    '<?xml version="1.0" encoding="utf-8"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      "<sheets>" +
      sheetElements.join("") +
      "</sheets></workbook>",
  );

  for (const part of extraParts) {
    builder.addPart(part);
  }

  builder.save(path);
}

function minimalWorksheetXml(): string {
  return '<?xml version="1.0" encoding="utf-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>';
}

function loadSharedStrings(reader: OpenXmlWorkbookReader): string[] {
  if (!reader.hasPart("xl/sharedStrings.xml")) {
    return [];
  }

  const root = documentRoot(reader.partXml("xl/sharedStrings.xml"));
  const items: string[] = [];
  for (const item of elementChildren(root, "si")) {
    const textNode = firstElement(item, "t");
    if (textNode !== null) {
      items.push(textNode.textContent ?? "");
    } else {
      items.push(descendants(item, "t").map((text) => text.textContent ?? "").join(""));
    }
  }

  return items;
}

function readCellValue(cell: XmlNode, sharedStrings: readonly string[]): string {
  const rawValue = (firstElement(cell, "v")?.textContent ?? "") as string;
  const type = (attribute(cell, "t") ?? "").trim();
  if (type === "s") {
    const index = tryParseInt(rawValue);
    if (index !== null && index >= 0 && index < sharedStrings.length) {
      return sharedStrings[index];
    }
  }

  if (type === "inlineStr") {
    return descendants(firstElement(cell, "is"), "t")
      .map((text) => text.textContent ?? "")
      .join("");
  }

  if (type === "b") {
    return rawValue === "1" ? "TRUE" : "FALSE";
  }

  return rawValue;
}

function tryParseDouble(value: string | null): number | null {
  if (value === null) {
    return null;
  }

  const parsed = Number.parseFloat(value);
  return Number.isNaN(parsed) ? null : parsed;
}

function tryParseInt(value: string | null): number | null {
  if (value === null) {
    return null;
  }

  const parsed = Number.parseInt(value, 10);
  return Number.isNaN(parsed) ? null : parsed;
}

function parseBool(value: string | null): boolean {
  return value === "1" || (value ?? "").toLowerCase() === "true";
}

function parseViewBool(value: string | null, defaultValue: boolean): boolean {
  if (value === null || value.trim().length === 0) {
    return defaultValue;
  }

  return value === "1" || value.toLowerCase() === "true";
}

function parseProtectionBool(value: string | null, defaultValue: boolean): boolean {
  if (value === null || value.trim().length === 0) {
    return defaultValue;
  }

  return value === "1" || value.toLowerCase() === "true";
}

TestRunner.runAndExit("WorksheetSettings.OpenXmlTests", [
  new TestCase("worksheet_visibility_matches_openxml_sdk", worksheetVisibilityMatchesOpenXmlSdk),
  new TestCase("worksheet_settings_and_values_match_openxml_sdk", worksheetSettingsAndValuesMatchOpenXmlSdk),
  new TestCase("worksheet_view_and_tab_color_match_openxml_sdk", worksheetViewAndTabColorMatchOpenXmlSdk),
  new TestCase("worksheet_protection_match_openxml_sdk", worksheetProtectionMatchOpenXmlSdk),
  new TestCase("autofilter_matches_openxml_sdk", autoFilterMatchesOpenXmlSdk),
  new TestCase(
    "adjacent_columns_are_combined_and_read_consistently",
    adjacentColumnsAreCombinedAndReadConsistently,
  ),
  new TestCase("merge_only_dimension_matches_openxml_sdk", mergeOnlyDimensionMatchesOpenXmlSdk),
  new TestCase("openxml_created_visibility_loads_in_library", openXmlCreatedVisibilityLoadsInLibrary),
  new TestCase(
    "openxml_created_settings_and_values_load_in_library",
    openXmlCreatedSettingsAndValuesLoadInLibrary,
  ),
  new TestCase("openxml_created_sheet_view_loads_in_library", openXmlCreatedSheetViewLoadsInLibrary),
  new TestCase(
    "openxml_created_sheet_protection_loads_in_library",
    openXmlCreatedSheetProtectionLoadsInLibrary,
  ),
  new TestCase("openxml_created_autofilter_loads_in_library", openXmlCreatedAutoFilterLoadsInLibrary),
  new TestCase("openxml_created_columns_load_in_library", openXmlCreatedColumnsLoadInLibrary),
  new TestCase(
    "openxml_created_merge_only_dimension_loads_in_library",
    openXmlCreatedMergeOnlyDimensionLoadsInLibrary,
  ),
]);
