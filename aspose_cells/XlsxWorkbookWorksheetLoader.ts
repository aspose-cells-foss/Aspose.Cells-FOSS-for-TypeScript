import { CellAddress } from "./Core/CellAddress";
import { CellRecord } from "./Core/CellRecord";
import { CellValueKind } from "./Core/CellValueKind";
import { ColorValue } from "./Core/ColorValue";
import { ColumnRangeModel } from "./Core/ColumnRangeModel";
import { DateSerialConverter } from "./Core/DateSerialConverter";
import { DateSystem } from "./Core/DateSystem";
import { FontSchemeType } from "./Core/FontSchemeType";
import { FontUnderlineType } from "./FontUnderlineType";
import { FontValue } from "./Core/FontValue";
import { MergeRegion } from "./Core/MergeRegion";
import { RichTextRunValue } from "./Core/RichTextRunValue";
import { RowModel } from "./Core/RowModel";
import { SharedStringEntry } from "./SharedStringEntry";
import { StyleValue } from "./Core/StyleValue";
import { StylesheetLoadContext } from "./StylesheetLoadContext";
import { WorksheetModel } from "./Core/WorksheetModel";
import { DiagnosticSeverity } from "./enums";
import { InvalidFileFormatException } from "./InvalidFileFormatException";
import { LoadDiagnostics } from "./LoadDiagnostics";
import { LoadIssue } from "./LoadIssue";
import { LoadOptions } from "./LoadOptions";
import { WarningInfo } from "./WarningInfo";
import { ZipReader } from "./Packaging/Archive";
import { shouldPersistCell, tryParseCellReference, tryParseMergeReference } from "./XlsxWorkbookSerializerCommon";
import { XmlNode, attribute, documentRoot, elementChildren, firstElement, localName, textOf } from "./Xml/XmlReader";
import { loadAutoFilter } from "./XlsxWorkbookAutoFilter";
import { loadComments } from "./XlsxWorkbookCommentsLoader";
import { loadConditionalFormattings } from "./XlsxWorkbookConditionalFormatting";
import { loadHyperlinks } from "./XlsxWorkbookHyperlinks";
import { applyWorksheetDefinedNames, loadWorksheetPageSetup } from "./XlsxWorkbookPageSetup";
import type { WorksheetDefinedNamesState } from "./XlsxWorkbookPageSetup";
import { loadPictures } from "./XlsxWorkbookPicturesLoader";
import { loadPrinterSettings } from "./XlsxWorkbookPrinterSettings";
import { loadTables } from "./XlsxWorkbookTables";
import { loadValidations } from "./XlsxWorkbookValidations";
import { loadWorksheetProtection } from "./XlsxWorkbookWorksheetProtection";
import { loadWorksheetViewSettings } from "./XlsxWorkbookWorksheetViews";

interface CellReadResult {
  readonly address: CellAddress;
  readonly record: CellRecord;
}

interface CellValueReadResult {
  readonly value: unknown | null;
  readonly kind: CellValueKind;
  readonly richTextRuns: RichTextRunValue[] | null;
}

export function loadWorksheet(
  sheetName: string,
  worksheetDocument: XmlNode,
  worksheetHyperlinkTargets: ReadonlyMap<string, string>,
  dateSystem: DateSystem,
  sharedStrings: readonly SharedStringEntry[],
  stylesheet: StylesheetLoadContext,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  definedNamesState: WorksheetDefinedNamesState | null,
  archive: ZipReader,
  worksheetUri: string,
): WorksheetModel {
  const worksheetModel = new WorksheetModel(sheetName);
  const worksheetRoot = documentRoot(worksheetDocument);
  if (worksheetRoot === null) {
    return worksheetModel;
  }

  loadWorksheetViewSettings(worksheetModel, worksheetRoot, diagnostics, options, sheetName);
  loadSheetFormatPr(worksheetModel, worksheetRoot);
  loadWorksheetProtection(worksheetModel, worksheetRoot, diagnostics, options, sheetName);
  loadAutoFilter(worksheetModel, worksheetRoot, stylesheet, diagnostics, options, sheetName);
  loadColumns(worksheetModel, worksheetRoot, stylesheet, diagnostics, options, sheetName);
  loadWorksheetPageSetup(worksheetModel, worksheetRoot, diagnostics, options, sheetName);
  loadPrinterSettings(worksheetModel.PageSetup, worksheetRoot, archive, worksheetUri, diagnostics, options, sheetName);
  loadHyperlinks(worksheetModel, worksheetRoot, worksheetHyperlinkTargets, diagnostics, options, sheetName);
  loadConditionalFormattings(worksheetModel, worksheetRoot, stylesheet, diagnostics, options, sheetName);
  loadValidations(worksheetModel, worksheetRoot, diagnostics, options, sheetName);
  loadTables(worksheetModel, worksheetRoot, archive, worksheetUri, diagnostics, options, sheetName);
  loadPictures(worksheetModel, worksheetRoot, archive, worksheetUri, diagnostics, options, sheetName);
  loadComments(worksheetModel, archive, worksheetUri, diagnostics, options, sheetName);
  applyWorksheetDefinedNames(worksheetModel, definedNamesState);

  const sheetData = firstElement(worksheetRoot, "sheetData");
  if (sheetData === null) {
    const issue = new LoadIssue(
      "ACF-WS-001",
      DiagnosticSeverity.Recoverable,
      "Worksheet sheetData is missing; an empty sheet was synthesized.",
      true,
      false,
    );
    issue.SheetName = sheetName;
    addIssue(diagnostics, options, issue);
    return worksheetModel;
  }

  const rowElements = elementChildren(sheetData, "row");
  const seenRows = new Set<number>();
  const seenCells = new Set<string>();
  let previousRowIndex = -1;
  let rowsOutOfOrderReported = false;

  for (const rowElement of rowElements) {
    const resolvedRow = tryResolveRowIndex(rowElement, diagnostics, options, sheetName);
    if (!resolvedRow.found) {
      continue;
    }

    const rowIndex = resolvedRow.rowIndex;
    if (seenRows.has(rowIndex)) {
      throw new InvalidFileFormatException(`Duplicate row index '${rowIndex + 1}' was found in worksheet '${sheetName}'.`);
    }
    seenRows.add(rowIndex);

    if (previousRowIndex > rowIndex && !rowsOutOfOrderReported) {
      const issue = new LoadIssue(
        "WS-R002",
        DiagnosticSeverity.Recoverable,
        "Worksheet rows were out of order and were normalized during load.",
        true,
        false,
      );
      issue.SheetName = sheetName;
      addIssue(diagnostics, options, issue);
      rowsOutOfOrderReported = true;
    }

    previousRowIndex = rowIndex;
    applyRowMetadata(worksheetModel, rowElement, rowIndex, stylesheet, diagnostics, options, sheetName);

    let previousColumnIndex = -1;
    let cellsOutOfOrderReported = false;
    for (const cellElement of elementChildren(rowElement, "c")) {
      const cellRead = tryReadCellRecord(
        cellElement,
        rowIndex,
        dateSystem,
        sharedStrings,
        stylesheet,
        diagnostics,
        options,
        sheetName,
      );
      if (cellRead === null) {
        continue;
      }

      if (previousColumnIndex > cellRead.address.ColumnIndex && !cellsOutOfOrderReported) {
        const issue = new LoadIssue(
          "WS-R003",
          DiagnosticSeverity.Recoverable,
          "Worksheet cells were out of order within a row and were normalized during load.",
          true,
          false,
        );
        issue.SheetName = sheetName;
        issue.RowIndex = rowIndex;
        addIssue(diagnostics, options, issue);
        cellsOutOfOrderReported = true;
      }

      previousColumnIndex = cellRead.address.ColumnIndex;
      const cellKey = cellRead.address.ToString();
      if (seenCells.has(cellKey)) {
        throw new InvalidFileFormatException(
          `Duplicate cell reference '${cellRead.address.ToString()}' was found in worksheet '${sheetName}'.`,
        );
      }
      seenCells.add(cellKey);

      if (shouldPersistCell(stylesheet.DefaultCellStyle, cellRead.record)) {
        worksheetModel.Cells.set(cellRead.address, cellRead.record);
      }
    }
  }

  loadMergeRegions(worksheetModel, worksheetRoot, diagnostics, options, sheetName);

  if (
    firstElement(worksheetRoot, "dimension") === null &&
    (worksheetModel.Cells.size > 0 || worksheetModel.MergeRegions.length > 0)
  ) {
    const issue = new LoadIssue(
      "WS-R001",
      DiagnosticSeverity.Recoverable,
      "Worksheet dimension was missing and was recalculated during load.",
      true,
      false,
    );
    issue.SheetName = sheetName;
    addIssue(diagnostics, options, issue);
  }

  return worksheetModel;
}

function loadSheetFormatPr(worksheetModel: WorksheetModel, worksheetRoot: XmlNode): void {
  const sheetFormatPr = firstElement(worksheetRoot, "sheetFormatPr");
  if (sheetFormatPr === null) {
    return;
  }

  worksheetModel.BaseColumnWidth = parseInvariantInt(attribute(sheetFormatPr, "baseColWidth"));
  worksheetModel.DefaultColumnWidth = parseInvariantDouble(attribute(sheetFormatPr, "defaultColWidth"));
  worksheetModel.DefaultRowHeight = parseInvariantDouble(attribute(sheetFormatPr, "defaultRowHeight"));
  if (attribute(sheetFormatPr, "customHeight") !== null) {
    worksheetModel.CustomHeight = parseBoolAttribute(attribute(sheetFormatPr, "customHeight"));
  }
}

function loadColumns(
  worksheetModel: WorksheetModel,
  worksheetRoot: XmlNode,
  stylesheet: StylesheetLoadContext,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  const columns: ColumnRangeModel[] = [];
  const colsElement = firstElement(worksheetRoot, "cols");
  const columnElements = colsElement === null ? [] : elementChildren(colsElement, "col");
  for (const columnElement of columnElements) {
    const min = parseInvariantInt(attribute(columnElement, "min"));
    const max = parseInvariantInt(attribute(columnElement, "max"));
    if (min === null || max === null || min <= 0 || max <= 0 || min > max) {
      throw new InvalidFileFormatException(
        `Worksheet column metadata in '${sheetName}' contains an invalid min/max span.`,
      );
    }

    let styleIndex = parseInvariantInt(attribute(columnElement, "style"));
    if (styleIndex !== null && (styleIndex < 0 || styleIndex >= stylesheet.CellFormats.length)) {
      if (options.StrictMode) {
        throw new InvalidFileFormatException(`The column style index '${styleIndex}' is invalid.`);
      }

      const issue = new LoadIssue(
        "COL-L001",
        DiagnosticSeverity.Warning,
        `Column style index '${styleIndex}' is invalid and was dropped.`,
        false,
        true,
      );
      issue.SheetName = sheetName;
      addIssue(diagnostics, options, issue);
      styleIndex = null;
    }

    const column = new ColumnRangeModel();
    column.MinColumnIndex = min - 1;
    column.MaxColumnIndex = max - 1;
    column.Width = parseInvariantDouble(attribute(columnElement, "width"));
    column.Hidden = parseBoolAttribute(attribute(columnElement, "hidden"));
    column.StyleIndex = styleIndex;
    column.Style = resolveMetadataStyle(styleIndex, stylesheet);
    columns.push(column);
  }

  for (const column of normalizeLoadedColumns(columns, diagnostics, options, sheetName)) {
    worksheetModel.Columns.push(column);
  }
}

function normalizeLoadedColumns(
  columns: readonly ColumnRangeModel[],
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): ColumnRangeModel[] {
  const ordered = Array.from(columns);
  ordered.sort(compareColumnRangesByBounds);
  if (ordered.length === 0) {
    return ordered;
  }

  const normalized: ColumnRangeModel[] = [ordered[0]!];
  let mergeReported = false;
  for (let index = 1; index < ordered.length; index++) {
    const current = ordered[index]!;
    const previous = normalized[normalized.length - 1]!;
    if (current.MinColumnIndex <= previous.MaxColumnIndex + 1 && columnRangesCompatible(previous, current)) {
      if (
        (current.MinColumnIndex <= previous.MaxColumnIndex || current.MaxColumnIndex <= previous.MaxColumnIndex) &&
        !mergeReported
      ) {
        const issue = new LoadIssue(
          "COL-R001",
          DiagnosticSeverity.Recoverable,
          "Overlapping compatible column metadata was normalized during load.",
          true,
          false,
        );
        issue.SheetName = sheetName;
        addIssue(diagnostics, options, issue);
        mergeReported = true;
      }

      previous.MaxColumnIndex = Math.max(previous.MaxColumnIndex, current.MaxColumnIndex);
      continue;
    }

    normalized.push(current);
  }

  return normalized;
}

function compareColumnRangesByBounds(left: ColumnRangeModel, right: ColumnRangeModel): number {
  const minComparison = left.MinColumnIndex - right.MinColumnIndex;
  if (minComparison !== 0) {
    return minComparison;
  }

  return left.MaxColumnIndex - right.MaxColumnIndex;
}

function columnRangesCompatible(left: ColumnRangeModel, right: ColumnRangeModel): boolean {
  return nullableEquals(left.Width, right.Width) && left.Hidden === right.Hidden && left.StyleIndex === right.StyleIndex;
}

function tryResolveRowIndex(
  rowElement: XmlNode,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): { found: boolean; rowIndex: number } {
  const rowIndexAttribute = parseInvariantInt(attribute(rowElement, "r"));
  if (rowIndexAttribute !== null) {
    if (rowIndexAttribute <= 0) {
      throw new InvalidFileFormatException(`Worksheet row index '${rowIndexAttribute}' is invalid.`);
    }

    return { found: true, rowIndex: rowIndexAttribute - 1 };
  }

  for (const cellElement of elementChildren(rowElement, "c")) {
    const cellReference = attribute(cellElement, "r") ?? "";
    const address = tryParseCellReference(cellReference);
    if (address !== null) {
      const issue = new LoadIssue(
        "WS-R004",
        DiagnosticSeverity.Recoverable,
        "A worksheet row index was missing and was inferred from contained cells.",
        true,
        false,
      );
      issue.SheetName = sheetName;
      issue.RowIndex = address.RowIndex;
      addIssue(diagnostics, options, issue);
      return { found: true, rowIndex: address.RowIndex };
    }
  }

  const issue = new LoadIssue(
    "ROW-F001",
    DiagnosticSeverity.Warning,
    "A worksheet row without an index and without parseable cells was skipped.",
  );
  issue.SheetName = sheetName;
  addIssue(diagnostics, options, issue);
  return { found: false, rowIndex: -1 };
}

function applyRowMetadata(
  worksheetModel: WorksheetModel,
  rowElement: XmlNode,
  rowIndex: number,
  stylesheet: StylesheetLoadContext,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  let styleIndex = parseInvariantInt(attribute(rowElement, "s"));
  if (styleIndex !== null && (styleIndex < 0 || styleIndex >= stylesheet.CellFormats.length)) {
    if (options.StrictMode) {
      throw new InvalidFileFormatException(`The row style index '${styleIndex}' is invalid.`);
    }

    const issue = new LoadIssue(
      "ROW-L001",
      DiagnosticSeverity.Warning,
      `Row style index '${styleIndex}' is invalid and was dropped.`,
      false,
      true,
    );
    issue.SheetName = sheetName;
    issue.RowIndex = rowIndex;
    addIssue(diagnostics, options, issue);
    styleIndex = null;
  }

  const rowModel = new RowModel();
  rowModel.Height = parseInvariantDouble(attribute(rowElement, "ht"));
  rowModel.Hidden = parseBoolAttribute(attribute(rowElement, "hidden"));
  rowModel.CustomHeight = parseBoolAttribute(attribute(rowElement, "customHeight"));
  rowModel.StyleIndex = styleIndex;
  rowModel.Style = resolveMetadataStyle(styleIndex, stylesheet);

  if (rowModel.Height !== null && !parseBoolAttribute(attribute(rowElement, "customHeight"))) {
    const issue = new LoadIssue(
      "ROW-R002",
      DiagnosticSeverity.Recoverable,
      "Row height metadata was missing customHeight and was normalized during load.",
      true,
      false,
    );
    issue.SheetName = sheetName;
    issue.RowIndex = rowIndex;
    addIssue(diagnostics, options, issue);
  }

  if (rowModel.Height !== null || rowModel.Hidden || rowModel.StyleIndex !== null) {
    worksheetModel.Rows.set(rowIndex, rowModel);
  }
}

function tryReadCellRecord(
  cellElement: XmlNode,
  rowIndex: number,
  dateSystem: DateSystem,
  sharedStrings: readonly SharedStringEntry[],
  stylesheet: StylesheetLoadContext,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): CellReadResult | null {
  const record = new CellRecord();
  record.Style = stylesheet.DefaultCellStyle.clone();
  record.IsExplicitlyStored = true;

  const cellReference = attribute(cellElement, "r");
  if (cellReference === null || cellReference.trim().length === 0) {
    const issue = new LoadIssue(
      "CELL-F001",
      DiagnosticSeverity.Warning,
      "A cell without a reference was skipped.",
    );
    issue.SheetName = sheetName;
    issue.RowIndex = rowIndex;
    addIssue(diagnostics, options, issue);
    return null;
  }

  let address: CellAddress;
  try {
    address = CellAddress.Parse(cellReference);
  } catch {
    if (options.StrictMode) {
      throw new InvalidFileFormatException(`The cell reference '${cellReference}' is invalid.`);
    }

    const issue = new LoadIssue(
      "CELL-F001",
      DiagnosticSeverity.Warning,
      `Cell reference '${cellReference}' is invalid and was skipped.`,
    );
    issue.SheetName = sheetName;
    issue.CellRef = cellReference;
    issue.RowIndex = rowIndex;
    addIssue(diagnostics, options, issue);
    return null;
  }

  const styleIndex = parseInvariantInt(attribute(cellElement, "s"));
  const isDateStyle = styleIndex !== null && stylesheet.DateStyleIndexes.has(styleIndex);
  if (styleIndex !== null) {
    if (styleIndex >= 0 && styleIndex < stylesheet.CellFormats.length) {
      record.Style = stylesheet.CellFormats[styleIndex]!.clone();
    } else if (options.StrictMode) {
      throw new InvalidFileFormatException(`The style index '${styleIndex}' is invalid.`);
    } else {
      const issue = new LoadIssue(
        "STYLE-F001",
        DiagnosticSeverity.Warning,
        `Cell style index '${styleIndex}' is invalid and style 0 was used instead.`,
      );
      issue.SheetName = sheetName;
      issue.CellRef = cellReference;
      issue.RowIndex = rowIndex;
      addIssue(diagnostics, options, issue);
    }
  }

  const formulaText = normalizeFormula(textOfElement(cellElement, "f"));
  if (formulaText.length > 0) {
    record.Formula = formulaText;
    record.Kind = CellValueKind.Formula;
  }

  const cellType = attribute(cellElement, "t");
  const valueElement = firstElement(cellElement, "v");
  const rawValue = valueElement === null ? null : textOf(valueElement);
  const valueResult = tryReadCellValue(
    cellElement,
    cellType,
    rawValue,
    isDateStyle,
    dateSystem,
    sharedStrings,
    diagnostics,
    options,
    sheetName,
    cellReference,
  );
  if (valueResult !== null) {
    record.Value = valueResult.value;
    if (
      record.Formula === null ||
      record.Formula.length === 0 ||
      valueResult.kind === CellValueKind.Error
    ) {
      record.Kind = valueResult.kind;
    }

    if (valueResult.richTextRuns !== null && valueResult.richTextRuns.length > 0) {
      record.RichTextRuns = valueResult.richTextRuns;
    }
  }

  return { address, record };
}

function resolveMetadataStyle(styleIndex: number | null, stylesheet: StylesheetLoadContext): StyleValue | null {
  if (styleIndex === null) {
    return null;
  }

  if (styleIndex < 0 || styleIndex >= stylesheet.CellFormats.length) {
    return null;
  }

  return stylesheet.CellFormats[styleIndex]!.clone();
}

function loadMergeRegions(
  worksheetModel: WorksheetModel,
  worksheetRoot: XmlNode,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  const mergeCellsElement = firstElement(worksheetRoot, "mergeCells");
  const mergeElements = mergeCellsElement === null ? [] : elementChildren(mergeCellsElement, "mergeCell");
  for (const mergeElement of mergeElements) {
    const mergeReference = attribute(mergeElement, "ref") ?? "";
    const region = tryParseMergeReference(mergeReference);
    if (region === null) {
      throw new InvalidFileFormatException(`The merge reference '${mergeReference}' is invalid.`);
    }

    if (containsOverlappingMergeRegion(worksheetModel.MergeRegions, region)) {
      if (options.StrictMode) {
        throw new InvalidFileFormatException(
          `The merge reference '${mergeReference}' overlaps an existing merged range.`,
        );
      }

      const issue = new LoadIssue(
        "MRG-L001",
        DiagnosticSeverity.LossyRecoverable,
        `Overlapping merge range '${mergeReference}' was dropped during load.`,
        true,
        true,
      );
      issue.SheetName = sheetName;
      addIssue(diagnostics, options, issue);
      continue;
    }

    worksheetModel.MergeRegions.push(region);
  }

  worksheetModel.MergeRegions.sort((left: MergeRegion, right: MergeRegion): number => {
    const rowComparison = left.FirstRow - right.FirstRow;
    if (rowComparison !== 0) {
      return rowComparison;
    }

    const columnComparison = left.FirstColumn - right.FirstColumn;
    if (columnComparison !== 0) {
      return columnComparison;
    }

    const rowCountComparison = left.TotalRows - right.TotalRows;
    return rowCountComparison !== 0 ? rowCountComparison : left.TotalColumns - right.TotalColumns;
  });
}

function containsOverlappingMergeRegion(mergeRegions: readonly MergeRegion[], candidate: MergeRegion): boolean {
  for (const region of mergeRegions) {
    if (mergeRegionsOverlap(region, candidate)) {
      return true;
    }
  }

  return false;
}

function mergeRegionsOverlap(left: MergeRegion, right: MergeRegion): boolean {
  const leftLastRow = left.FirstRow + left.TotalRows - 1;
  const leftLastColumn = left.FirstColumn + left.TotalColumns - 1;
  const rightLastRow = right.FirstRow + right.TotalRows - 1;
  const rightLastColumn = right.FirstColumn + right.TotalColumns - 1;

  return (
    left.FirstRow <= rightLastRow &&
    right.FirstRow <= leftLastRow &&
    left.FirstColumn <= rightLastColumn &&
    right.FirstColumn <= leftLastColumn
  );
}

function tryReadCellValue(
  cellElement: XmlNode,
  cellType: string | null,
  rawValue: string | null,
  isDateStyle: boolean,
  dateSystem: DateSystem,
  sharedStrings: readonly SharedStringEntry[],
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  cellReference: string,
): CellValueReadResult | null {
  if (cellType === "inlineStr") {
    const inlineString = readSharedStringEntry(firstElement(cellElement, "is"));
    return { value: inlineString.Text, kind: CellValueKind.String, richTextRuns: inlineString.Runs };
  }

  if (cellType === "s") {
    const sharedStringIndex = parseInvariantInt(rawValue);
    if (sharedStringIndex !== null && sharedStringIndex >= 0 && sharedStringIndex < sharedStrings.length) {
      const entry = sharedStrings[sharedStringIndex]!;
      return { value: entry.Text, kind: CellValueKind.String, richTextRuns: entry.Runs };
    }

    const issue = new LoadIssue(
      "SST-L001",
      DiagnosticSeverity.LossyRecoverable,
      "The cell points to an invalid shared string index.",
      false,
      true,
    );
    issue.SheetName = sheetName;
    issue.CellRef = cellReference;
    addIssue(diagnostics, options, issue);

    return { value: "", kind: CellValueKind.String, richTextRuns: null };
  }

  if (cellType === "b") {
    const normalized = rawValue ?? "";
    const booleanValue = normalized === "1" || normalized.toLowerCase() === "true";
    return { value: booleanValue, kind: CellValueKind.Boolean, richTextRuns: null };
  }

  if (cellType === "d") {
    const dateTime = tryParseRoundtripDate(rawValue);
    if (dateTime !== null) {
      return { value: dateTime, kind: CellValueKind.DateTime, richTextRuns: null };
    }

    return null;
  }

  if (cellType === "str") {
    return { value: rawValue ?? "", kind: CellValueKind.String, richTextRuns: null };
  }

  if (cellType === "e") {
    return { value: rawValue ?? "", kind: CellValueKind.Error, richTextRuns: null };
  }

  if (rawValue === null || rawValue.length === 0) {
    return null;
  }

  if (isDateStyle) {
    const serial = parseInvariantDouble(rawValue);
    if (serial !== null) {
      return {
        value: DateSerialConverter.FromSerial(serial, dateSystem),
        kind: CellValueKind.DateTime,
        richTextRuns: null,
      };
    }

    const issue = new LoadIssue(
      "CELL-R002",
      DiagnosticSeverity.Recoverable,
      "A formula or numeric cell contained an invalid cached date serial and the cached value was cleared.",
      true,
      false,
    );
    issue.SheetName = sheetName;
    issue.CellRef = cellReference;
    addIssue(diagnostics, options, issue);
    return null;
  }

  const numberValue = tryParseNumber(rawValue);
  if (numberValue !== null) {
    return { value: numberValue, kind: CellValueKind.Number, richTextRuns: null };
  }

  return null;
}

function readSharedStringEntry(inlineStringElement: XmlNode | null): SharedStringEntry {
  const entry = new SharedStringEntry();
  if (inlineStringElement === null) {
    entry.Text = "";
    return entry;
  }

  const runs = readRichTextRuns(inlineStringElement);
  entry.Text = readInlineString(inlineStringElement);
  if (runs.length > 0) {
    entry.Runs = runs;
  }

  return entry;
}

function readRichTextRuns(inlineStringElement: XmlNode | null): RichTextRunValue[] {
  const runs: RichTextRunValue[] = [];
  if (inlineStringElement === null) {
    return runs;
  }

  let currentIndex = 0;
  for (const runElement of elementChildren(inlineStringElement, "r")) {
    const text = readInlineString(runElement);
    if (text.length === 0) {
      continue;
    }

    const run = new RichTextRunValue();
    run.StartIndex = currentIndex;
    run.Length = text.length;
    run.Font = readRichTextFontValue(firstElement(runElement, "rPr"));
    runs.push(run);
    currentIndex += text.length;
  }

  return runs;
}

function readRichTextFontValue(runProperties: XmlNode | null): FontValue {
  if (runProperties === null) {
    return new FontValue();
  }

  const font = new FontValue();
  font.Name = attribute(firstElement(runProperties, "rFont"), "val") ?? "Calibri";
  font.Size = parseInvariantDouble(attribute(firstElement(runProperties, "sz"), "val")) ?? 11;
  font.Bold = firstElement(runProperties, "b") !== null;
  font.Italic = firstElement(runProperties, "i") !== null;
  font.Underline = parseRichTextUnderlineType(firstElement(runProperties, "u"));
  font.StrikeThrough = firstElement(runProperties, "strike") !== null;
  font.Color = readRichTextColorValue(firstElement(runProperties, "color"));
  font.Family = parseInvariantInt(attribute(firstElement(runProperties, "family"), "val"));
  font.Scheme = parseRichTextFontScheme(attribute(firstElement(runProperties, "scheme"), "val"));
  return font;
}

function parseRichTextUnderlineType(underlineElement: XmlNode | null): FontUnderlineType {
  if (underlineElement === null) {
    return FontUnderlineType.None;
  }

  const value = attribute(underlineElement, "val");
  if (value === null || value.length === 0) {
    return FontUnderlineType.Single;
  }

  switch (value) {
    case "double":
      return FontUnderlineType.Double;
    case "singleAccounting":
      return FontUnderlineType.Accounting;
    case "doubleAccounting":
      return FontUnderlineType.DoubleAccounting;
    default:
      return FontUnderlineType.Single;
  }
}

function parseRichTextFontScheme(value: string | null): FontSchemeType {
  switch (value) {
    case "major":
      return FontSchemeType.Major;
    case "minor":
      return FontSchemeType.Minor;
    default:
      return FontSchemeType.None;
  }
}

function readRichTextColorValue(colorElement: XmlNode | null): ColorValue {
  if (colorElement === null) {
    return new ColorValue();
  }

  const themeAttribute = attribute(colorElement, "theme");
  if (themeAttribute !== null && themeAttribute.length > 0) {
    const themeIndex = parseInvariantInt(themeAttribute);
    if (themeIndex !== null) {
      const tint = parseInvariantDouble(attribute(colorElement, "tint"));
      return new ColorValue(themeIndex, tint);
    }
  }

  const indexedAttribute = attribute(colorElement, "indexed");
  if (indexedAttribute !== null && indexedAttribute.length > 0) {
    const indexedValue = parseInvariantInt(indexedAttribute);
    if (indexedValue !== null) {
      return new ColorValue(indexedValue);
    }
  }

  const rgb = attribute(colorElement, "rgb");
  if (rgb === null || rgb.trim().length === 0) {
    return new ColorValue();
  }

  let rgbValue = rgb.trim().toUpperCase();
  if (rgbValue.length === 6) {
    rgbValue = "FF" + rgbValue;
  }

  if (rgbValue.length !== 8 || !/^[0-9A-F]{8}$/.test(rgbValue)) {
    return new ColorValue();
  }

  const a = parseInt(rgbValue.substring(0, 2), 16);
  const r = parseInt(rgbValue.substring(2, 2 + 2), 16);
  const g = parseInt(rgbValue.substring(4, 4 + 2), 16);
  const b = parseInt(rgbValue.substring(6, 6 + 2), 16);
  return new ColorValue(a, r, g, b);
}

function readInlineString(inlineStringElement: XmlNode | null): string {
  if (inlineStringElement === null) {
    return "";
  }

  const textNodes: string[] = [];
  for (const element of descendantsAndSelf(inlineStringElement)) {
    if (localName(element.nodeName) === "t") {
      textNodes.push(element.textContent ?? "");
    }
  }

  return textNodes.length > 0 ? textNodes.join("") : inlineStringElement.textContent ?? "";
}

function descendantsAndSelf(node: XmlNode): XmlNode[] {
  const result: XmlNode[] = [];
  const visit = (current: XmlNode): void => {
    result.push(current);
    if (!current.childNodes) {
      return;
    }

    for (let index = 0; index < current.childNodes.length; index++) {
      visit(current.childNodes[index]!);
    }
  };

  visit(node);
  return result;
}

function normalizeFormula(value: string): string {
  if (value.trim().length === 0) {
    return "";
  }

  const normalizedValue = value.trim();
  return normalizedValue.startsWith("=") ? normalizedValue.substring(1) : normalizedValue;
}

function textOfElement(element: XmlNode, name: string): string {
  const child = firstElement(element, name);
  return child === null ? "" : textOf(child);
}

function tryParseRoundtripDate(value: string | null): Date | null {
  if (value === null || value.trim().length === 0) {
    return null;
  }

  const date = new Date(value.trim());
  return Number.isNaN(date.getTime()) ? null : date;
}

const INT_MIN = -2147483648;
const INT_MAX = 2147483647;

function tryParseNumber(rawValue: string): number | null {
  if (!rawValue.includes("e") && !rawValue.includes("E")) {
    const decimalValue = parseDecimalInvariant(rawValue);
    if (decimalValue !== null) {
      if (
        !rawValue.includes(".") &&
        !rawValue.includes(",") &&
        Number.isInteger(decimalValue) &&
        decimalValue >= INT_MIN &&
        decimalValue <= INT_MAX
      ) {
        return decimalValue;
      }

      return decimalValue;
    }
  }

  const doubleValue = parseInvariantDouble(rawValue);
  if (doubleValue !== null) {
    if (Math.abs(doubleValue % 1) < Number.EPSILON && doubleValue >= INT_MIN && doubleValue <= INT_MAX) {
      return Math.trunc(doubleValue);
    }

    return doubleValue;
  }

  return null;
}

function parseDecimalInvariant(text: string): number | null {
  const cleaned = text.trim().replace(/,/g, "");
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(cleaned)) {
    return null;
  }

  const value = Number(cleaned);
  return Number.isFinite(value) ? value : null;
}

function parseInvariantDouble(text: string | null): number | null {
  if (text === null) {
    return null;
  }

  const trimmed = text.trim();
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(trimmed)) {
    return null;
  }

  const value = Number(trimmed);
  return Number.isFinite(value) ? value : null;
}

function parseInvariantInt(text: string | null): number | null {
  if (text === null) {
    return null;
  }

  const trimmed = text.trim();
  if (!/^[+-]?\d+$/.test(trimmed)) {
    return null;
  }

  const value = Number(trimmed);
  return Number.isSafeInteger(value) ? value : null;
}

function parseBoolAttribute(value: string | null): boolean {
  if (value === null) {
    return false;
  }

  return value === "1" || value.toLowerCase() === "true";
}

function nullableEquals(left: number | null, right: number | null): boolean {
  return left === right;
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
