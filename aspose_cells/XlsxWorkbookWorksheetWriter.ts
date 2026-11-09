import { CellAddress } from "./Core/CellAddress";
import { CellRecord } from "./Core/CellRecord";
import { CellValueKind } from "./Core/CellValueKind";
import { ColumnRangeModel } from "./Core/ColumnRangeModel";
import { DateSerialConverter } from "./Core/DateSerialConverter";
import { DateSystem } from "./Core/DateSystem";
import { MergeRegion } from "./Core/MergeRegion";
import { RichTextRunValue } from "./Core/RichTextRunValue";
import { RowModel } from "./Core/RowModel";
import { SharedStringEntry } from "./SharedStringEntry";
import { SharedStringRepository } from "./Core/SharedStringRepository";
import { StyleValue } from "./Core/StyleValue";
import { WorksheetModel } from "./Core/WorksheetModel";
import { SaveOptions } from "./SaveOptions";
import { buildAutoFilterElement } from "./XlsxWorkbookAutoFilter";
import {
  MAIN_NS,
  RELATIONSHIP_NS,
  createInlineStringElement,
  createTextElement,
  shouldPersistCell,
  toRangeReference,
} from "./XlsxWorkbookSerializerCommon";
import { buildConditionalFormattingElements } from "./XlsxWorkbookConditionalFormatting";
import { buildHyperlinksElement } from "./XlsxWorkbookHyperlinks";
import {
  buildColumnBreaksElement,
  buildHeaderFooterElement,
  buildPageMarginsElement,
  buildPageSetupElement,
  buildPrintOptionsElement,
  buildRowBreaksElement,
} from "./XlsxWorkbookPageSetup";
import { buildTablePartsElement } from "./XlsxWorkbookTables";
import { buildDataValidationsElement } from "./XlsxWorkbookValidations";
import { StylesheetSaveContext } from "./XlsxWorkbookStyles";
import { buildSheetProtectionElement } from "./XlsxWorkbookWorksheetProtection";
import { buildWorksheetSheetProperties, buildWorksheetViewsElement } from "./XlsxWorkbookWorksheetViews";
import { XmlElement } from "./Xml/XmlWriter";

/** Mirrors the C# KeyValuePair<CellAddress, CellRecord> used for persisted cells. */
export interface CellAddressRecordPair {
  readonly Key: CellAddress;
  readonly Value: CellRecord;
}

export function buildWorksheet(
  worksheet: WorksheetModel,
  workbookDefaultStyle: StyleValue,
  dateSystem: DateSystem,
  sharedStrings: SharedStringRepository,
  options: SaveOptions,
  stylesheet: StylesheetSaveContext,
  externalHyperlinkCount: number,
  hasPictures: boolean,
  hasComments: boolean,
): XmlElement {
  const persistedCells = collectPersistedCells(worksheet, workbookDefaultStyle);
  const worksheetElement = XmlElement.createWithNamespace("worksheet", MAIN_NS).attr("xmlns:r", RELATIONSHIP_NS);

  const sheetProperties = buildWorksheetSheetProperties(worksheet);
  if (sheetProperties !== null) {
    worksheetElement.child(sheetProperties);
  }

  const dimensionReference = calculateDimensionReference(persistedCells, worksheet.MergeRegions);
  if (dimensionReference.length > 0) {
    worksheetElement.child(XmlElement.create("dimension").attr("ref", dimensionReference));
  }

  const sheetViews = buildWorksheetViewsElement(worksheet);
  if (sheetViews !== null) {
    worksheetElement.child(sheetViews);
  }

  const sheetFormatPr = buildSheetFormatPrElement(worksheet);
  if (sheetFormatPr !== null) {
    worksheetElement.child(sheetFormatPr);
  }

  const normalizedColumns = normalizeColumnRanges(worksheet.Columns);
  if (normalizedColumns.length > 0) {
    worksheetElement.child(XmlElement.create("cols").children(buildColumnElements(normalizedColumns, stylesheet)));
  }

  const sheetData = XmlElement.create("sheetData");
  let persistedCellIndex = 0;
  for (const rowIndex of getWorksheetRowIndexes(persistedCells, worksheet.Rows)) {
    const [rowCells, nextPersistedCellIndex] = getRowCells(persistedCells, rowIndex, persistedCellIndex);
    persistedCellIndex = nextPersistedCellIndex;
    const rowModel = worksheet.Rows.get(rowIndex) ?? null;
    if (rowCells.length === 0 && !hasRowMetadata(rowModel)) {
      continue;
    }

    sheetData.child(buildRowElement(rowIndex, rowModel, rowCells, dateSystem, sharedStrings, options, stylesheet));
  }

  worksheetElement.child(sheetData);

  const sheetProtection = buildSheetProtectionElement(worksheet);
  if (sheetProtection !== null) {
    worksheetElement.child(sheetProtection);
  }

  const autoFilter = buildAutoFilterElement(worksheet, stylesheet.DifferentialFormatCount);
  if (autoFilter !== null) {
    worksheetElement.child(autoFilter);
  }

  const mergeRegions = sortMergeRegions(worksheet.MergeRegions);
  if (mergeRegions.length > 0) {
    worksheetElement.child(
      XmlElement.create("mergeCells")
        .attr("count", mergeRegions.length)
        .children(buildMergeCellElements(mergeRegions)),
    );
  }

  worksheetElement.children(buildConditionalFormattingElements(worksheet, stylesheet));

  const dataValidations = buildDataValidationsElement(worksheet);
  if (dataValidations !== null) {
    worksheetElement.child(dataValidations);
  }

  const hyperlinks = buildHyperlinksElement(worksheet);
  if (hyperlinks !== null) {
    worksheetElement.child(hyperlinks);
  }

  const printOptions = buildPrintOptionsElement(worksheet.PageSetup);
  if (printOptions !== null) {
    worksheetElement.child(printOptions);
  }

  const pageMargins = buildPageMarginsElement(worksheet.PageSetup);
  if (pageMargins !== null) {
    worksheetElement.child(pageMargins);
  }

  const pageSetup = buildPageSetupElement(worksheet.PageSetup);
  if (pageSetup !== null) {
    worksheetElement.child(pageSetup);
  }

  const headerFooter = buildHeaderFooterElement(worksheet.PageSetup);
  if (headerFooter !== null) {
    worksheetElement.child(headerFooter);
  }

  const rowBreaks = buildRowBreaksElement(worksheet.PageSetup);
  if (rowBreaks !== null) {
    worksheetElement.child(rowBreaks);
  }

  const columnBreaks = buildColumnBreaksElement(worksheet.PageSetup);
  if (columnBreaks !== null) {
    worksheetElement.child(columnBreaks);
  }

  if (hasPictures) {
    const drawingRId = externalHyperlinkCount + worksheet.ListObjects.length + 1;
    worksheetElement.child(XmlElement.create("drawing").attr("r:id", `rId${drawingRId}`));
  }

  if (hasComments) {
    const vmlRId = externalHyperlinkCount + worksheet.ListObjects.length + (hasPictures ? 1 : 0) + 2;
    worksheetElement.child(XmlElement.create("legacyDrawing").attr("r:id", `rId${vmlRId}`));
  }

  const tableRIdStart = externalHyperlinkCount + 1;
  const tableParts = buildTablePartsElement(worksheet, tableRIdStart);
  if (tableParts !== null) {
    worksheetElement.child(tableParts);
  }

  return worksheetElement;
}

export function buildRowElement(
  rowIndex: number,
  rowModel: RowModel | null,
  rowCells: readonly CellAddressRecordPair[],
  dateSystem: DateSystem,
  sharedStrings: SharedStringRepository,
  options: SaveOptions,
  stylesheet: StylesheetSaveContext,
): XmlElement {
  const row = XmlElement.create("row").attr("r", rowIndex + 1);
  if (rowModel !== null && rowModel.Height !== null) {
    row.attr("ht", String(rowModel.Height));
    row.attr("customHeight", 1);
  }

  if (rowModel !== null && rowModel.Hidden) {
    row.attr("hidden", 1);
  }

  if (rowModel !== null && rowModel.Style !== null) {
    row.attr("s", stylesheet.GetStyleIndex(rowModel.Style));
    row.attr("customFormat", 1);
  } else if (rowModel !== null && rowModel.StyleIndex !== null && rowModel.StyleIndex >= 0) {
    row.attr("s", rowModel.StyleIndex);
    row.attr("customFormat", 1);
  }

  for (const pair of rowCells) {
    row.child(buildCell(pair.Key, pair.Value, dateSystem, sharedStrings, options, stylesheet));
  }

  return row;
}

export function buildColumnElement(column: ColumnRangeModel, stylesheet: StylesheetSaveContext): XmlElement {
  const element = XmlElement.create("col")
    .attr("min", column.MinColumnIndex + 1)
    .attr("max", column.MaxColumnIndex + 1);

  if (column.Width !== null) {
    element.attr("width", String(column.Width));
    element.attr("customWidth", 1);
  }

  if (column.Hidden) {
    element.attr("hidden", 1);
  }

  if (column.Style !== null) {
    element.attr("style", stylesheet.GetStyleIndex(column.Style));
  } else if (column.StyleIndex !== null && column.StyleIndex >= 0) {
    element.attr("style", column.StyleIndex);
  }

  return element;
}

export function hasRowMetadata(rowModel: RowModel | null): boolean {
  return rowModel !== null && (rowModel.Height !== null || rowModel.Hidden || rowModel.StyleIndex !== null);
}

export function hasColumnMetadata(column: ColumnRangeModel): boolean {
  return column.Width !== null || column.Hidden || column.StyleIndex !== null;
}

export function getWorksheetRowIndexes(
  persistedCells: readonly CellAddressRecordPair[],
  rows: ReadonlyMap<number, RowModel>,
): number[] {
  const indexes = new Set<number>();
  for (const pair of persistedCells) {
    indexes.add(pair.Key.RowIndex);
  }

  for (const rowIndex of rows.keys()) {
    indexes.add(rowIndex);
  }

  const orderedIndexes = Array.from(indexes);
  orderedIndexes.sort((left, right) => left - right);
  return orderedIndexes;
}

export function normalizeColumnRanges(columns: readonly ColumnRangeModel[]): ColumnRangeModel[] {
  const ordered: ColumnRangeModel[] = [];
  for (const column of columns) {
    if (!hasColumnMetadata(column)) {
      continue;
    }

    const clone = new ColumnRangeModel();
    clone.MinColumnIndex = column.MinColumnIndex;
    clone.MaxColumnIndex = column.MaxColumnIndex;
    clone.Width = column.Width;
    clone.Hidden = column.Hidden;
    clone.StyleIndex = column.StyleIndex;
    clone.Style = column.Style !== null ? column.Style.clone() : null;
    ordered.push(clone);
  }

  ordered.sort(compareColumnRangesByBounds);
  if (ordered.length === 0) {
    return ordered;
  }

  const normalized: ColumnRangeModel[] = [ordered[0]!];
  for (let index = 1; index < ordered.length; index++) {
    const current = ordered[index]!;
    const previous = normalized[normalized.length - 1]!;
    if (previous.MaxColumnIndex + 1 >= current.MinColumnIndex && columnMetadataEqual(previous, current)) {
      previous.MaxColumnIndex = Math.max(previous.MaxColumnIndex, current.MaxColumnIndex);
      continue;
    }

    normalized.push(current);
  }

  return normalized;
}

export function columnMetadataEqual(left: ColumnRangeModel, right: ColumnRangeModel): boolean {
  return nullableEquals(left.Width, right.Width) && left.Hidden === right.Hidden && left.StyleIndex === right.StyleIndex;
}

export function calculateDimensionReference(
  persistedCells: readonly CellAddressRecordPair[],
  mergeRegions: readonly MergeRegion[],
): string {
  const hasCells = persistedCells.length > 0;
  const hasMerges = mergeRegions.length > 0;
  if (!hasCells && !hasMerges) {
    return "";
  }

  let minRow = Number.MAX_SAFE_INTEGER;
  let minColumn = Number.MAX_SAFE_INTEGER;
  let maxRow = Number.MIN_SAFE_INTEGER;
  let maxColumn = Number.MIN_SAFE_INTEGER;

  for (const pair of persistedCells) {
    minRow = Math.min(minRow, pair.Key.RowIndex);
    minColumn = Math.min(minColumn, pair.Key.ColumnIndex);
    maxRow = Math.max(maxRow, pair.Key.RowIndex);
    maxColumn = Math.max(maxColumn, pair.Key.ColumnIndex);
  }

  for (const region of mergeRegions) {
    minRow = Math.min(minRow, region.FirstRow);
    minColumn = Math.min(minColumn, region.FirstColumn);
    maxRow = Math.max(maxRow, region.FirstRow + region.TotalRows - 1);
    maxColumn = Math.max(maxColumn, region.FirstColumn + region.TotalColumns - 1);
  }

  const firstAddress = new CellAddress(minRow, minColumn).ToString();
  const lastAddress = new CellAddress(maxRow, maxColumn).ToString();
  return firstAddress === lastAddress ? firstAddress : `${firstAddress}:${lastAddress}`;
}

export function buildCell(
  address: CellAddress,
  record: CellRecord,
  dateSystem: DateSystem,
  sharedStrings: SharedStringRepository,
  options: SaveOptions,
  stylesheet: StylesheetSaveContext,
): XmlElement {
  const cell = XmlElement.create("c").attr("r", address.ToString());
  const styleIndex = stylesheet.GetStyleIndex(record);
  if (styleIndex > 0) {
    cell.attr("s", styleIndex);
  }

  const formula = record.Formula;
  const hasFormula = formula !== null && formula.length > 0;
  if (hasFormula) {
    cell.child(XmlElement.create("f").child(formula));
  }

  const value = record.Value;
  if (value === null || value === undefined) {
    return cell;
  }

  if (record.Kind === CellValueKind.Error) {
    cell.attr("t", "e");
    cell.child(createValueElement(String(value)));
    return cell;
  }

  if (typeof value === "string") {
    if (hasFormula) {
      cell.attr("t", "str");
      cell.child(createValueElement(value));
    } else if (options.UseSharedStrings) {
      cell.attr("t", "s");
      if (record.RichTextRuns !== null && record.RichTextRuns.length > 0) {
        const entry = new SharedStringEntry();
        entry.Text = value;
        entry.Runs = cloneRichTextRuns(record.RichTextRuns);
        cell.child(createValueElement(String(sharedStrings.Intern(entry))));
      } else {
        cell.child(createValueElement(String(sharedStrings.Intern(value))));
      }
    } else {
      cell.attr("t", "inlineStr");
      cell.child(createInlineStringElement(value, record.RichTextRuns));
    }
  } else if (typeof value === "boolean") {
    cell.attr("t", "b");
    cell.child(createValueElement(value ? "1" : "0"));
  } else if (value instanceof Date) {
    cell.child(createValueElement(String(DateSerialConverter.ToSerial(value, dateSystem))));
  } else if (typeof value === "number") {
    cell.child(createValueElement(String(value)));
  } else if (hasFormula) {
    cell.attr("t", "str");
    cell.child(createValueElement(String(value)));
  } else if (options.UseSharedStrings) {
    const fallbackText = String(value);
    cell.attr("t", "s");
    cell.child(createValueElement(String(sharedStrings.Intern(fallbackText))));
  } else {
    cell.attr("t", "inlineStr");
    cell.child(XmlElement.create("is").child(createTextElement(String(value))));
  }

  return cell;
}

function createValueElement(text: string): XmlElement {
  return XmlElement.create("v").child(text);
}

function collectPersistedCells(
  worksheet: WorksheetModel,
  workbookDefaultStyle: StyleValue,
): CellAddressRecordPair[] {
  const persistedCells: CellAddressRecordPair[] = [];
  for (const pair of worksheet.Cells) {
    if (shouldPersistCell(workbookDefaultStyle, pair[1])) {
      persistedCells.push({ Key: pair[0], Value: pair[1] });
    }
  }

  persistedCells.sort(comparePersistedCells);
  return persistedCells;
}

function comparePersistedCells(left: CellAddressRecordPair, right: CellAddressRecordPair): number {
  const rowComparison = left.Key.RowIndex - right.Key.RowIndex;
  if (rowComparison !== 0) {
    return rowComparison;
  }

  return left.Key.ColumnIndex - right.Key.ColumnIndex;
}

function buildColumnElements(
  columns: readonly ColumnRangeModel[],
  stylesheet: StylesheetSaveContext,
): XmlElement[] {
  const columnElements: XmlElement[] = [];
  for (const column of columns) {
    columnElements.push(buildColumnElement(column, stylesheet));
  }

  return columnElements;
}

function getRowCells(
  persistedCells: readonly CellAddressRecordPair[],
  rowIndex: number,
  startIndex: number,
): [CellAddressRecordPair[], number] {
  const rowCells: CellAddressRecordPair[] = [];
  let index = startIndex;
  while (index < persistedCells.length && persistedCells[index]!.Key.RowIndex < rowIndex) {
    index++;
  }

  while (index < persistedCells.length) {
    const pair = persistedCells[index]!;
    if (pair.Key.RowIndex !== rowIndex) {
      break;
    }

    rowCells.push(pair);
    index++;
  }

  return [rowCells, index];
}

function sortMergeRegions(mergeRegions: readonly MergeRegion[]): MergeRegion[] {
  const sortedRegions = Array.from(mergeRegions);
  sortedRegions.sort(compareMergeRegions);
  return sortedRegions;
}

function compareMergeRegions(left: MergeRegion, right: MergeRegion): number {
  const rowComparison = left.FirstRow - right.FirstRow;
  if (rowComparison !== 0) {
    return rowComparison;
  }

  const columnComparison = left.FirstColumn - right.FirstColumn;
  if (columnComparison !== 0) {
    return columnComparison;
  }

  const rowCountComparison = left.TotalRows - right.TotalRows;
  if (rowCountComparison !== 0) {
    return rowCountComparison;
  }

  return left.TotalColumns - right.TotalColumns;
}

function buildMergeCellElements(mergeRegions: readonly MergeRegion[]): XmlElement[] {
  const mergeCellElements: XmlElement[] = [];
  for (const region of mergeRegions) {
    mergeCellElements.push(XmlElement.create("mergeCell").attr("ref", toRangeReference(region)));
  }

  return mergeCellElements;
}

function buildSheetFormatPrElement(worksheet: WorksheetModel): XmlElement | null {
  if (
    worksheet.BaseColumnWidth === null &&
    worksheet.DefaultColumnWidth === null &&
    worksheet.DefaultRowHeight === null &&
    worksheet.CustomHeight === null
  ) {
    return null;
  }

  const element = XmlElement.create("sheetFormatPr");
  if (worksheet.BaseColumnWidth !== null) {
    element.attr("baseColWidth", worksheet.BaseColumnWidth);
  }

  if (worksheet.DefaultColumnWidth !== null) {
    element.attr("defaultColWidth", formatAtMost4Decimals(worksheet.DefaultColumnWidth));
  }

  // defaultRowHeight is required by the OOXML schema whenever sheetFormatPr is present.
  const defaultRowHeight = worksheet.DefaultRowHeight ?? 15;
  element.attr("defaultRowHeight", formatAtMost4Decimals(defaultRowHeight));

  if (worksheet.CustomHeight === true) {
    element.attr("customHeight", 1);
  }

  return element;
}

function cloneRichTextRuns(runs: readonly RichTextRunValue[] | null): RichTextRunValue[] | null {
  if (runs === null || runs.length === 0) {
    return null;
  }

  const cloned: RichTextRunValue[] = [];
  for (const run of runs) {
    cloned.push(run.clone());
  }

  return cloned;
}

function compareColumnRangesByBounds(left: ColumnRangeModel, right: ColumnRangeModel): number {
  const minComparison = left.MinColumnIndex - right.MinColumnIndex;
  return minComparison !== 0 ? minComparison : left.MaxColumnIndex - right.MaxColumnIndex;
}

function nullableEquals(left: number | null, right: number | null): boolean {
  return left === right;
}

/** Mirrors the invariant "0.####" number formatting used by the .NET writer. */
function formatAtMost4Decimals(value: number): string {
  if (!Number.isFinite(value)) {
    return "0";
  }

  return String(Number(value.toFixed(4)));
}
