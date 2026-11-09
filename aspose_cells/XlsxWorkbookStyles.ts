import { AlignmentValue } from "./Core/AlignmentValue";
import { BorderSideValue } from "./Core/BorderSideValue";
import { BorderStyle } from "./Core/BorderStyle";
import { BordersValue } from "./Core/BordersValue";
import { CellAddress } from "./Core/CellAddress";
import { CellRecord } from "./Core/CellRecord";
import { CellValueKind } from "./Core/CellValueKind";
import { ColorValue } from "./Core/ColorValue";
import { FillPatternKind } from "./Core/FillPatternKind";
import { FontSchemeType } from "./Core/FontSchemeType";
import { FontUnderlineType } from "./FontUnderlineType";
import { FontValue } from "./Core/FontValue";
import { FormatConditionModel } from "./Core/FormatConditionModel";
import { NumberFormatValue } from "./Core/NumberFormatValue";
import { ProtectionValue } from "./Core/ProtectionValue";
import { StyleValue } from "./Core/StyleValue";
import { StylesheetLoadContext } from "./StylesheetLoadContext";
import { WorkbookModel } from "./Core/WorkbookModel";
import { WorksheetModel } from "./Core/WorksheetModel";
import { LoadDiagnostics } from "./LoadDiagnostics";
import { LoadOptions } from "./LoadOptions";
import { resolveStylesUri } from "./XlsxWorkbookArchiveHelpers";
import { ZipReader } from "./Packaging/Archive";
import {
  CellFormatValue,
  FillValue,
  formatDoubleR,
  formatNumber4,
  toArgbHex,
} from "./XlsxWorkbookStylesValueHelpers";
import {
  CustomNumberFormat,
  buildStylesheetDocument,
  parseHorizontalAlignment,
  parseOptionalBoolAttribute,
  parseVerticalAlignment,
  readBordersValues,
  readDifferentialStyleValues,
  readFillValues,
  readFontValues,
} from "./XlsxWorkbookStylesXml";
import {
  XmlElement,
  escapeXmlAttribute as escapeAttribute,
  escapeXmlText as escapeText,
} from "./Xml/XmlWriter";
import {
  XmlNode,
  attribute,
  attributeInt,
  documentRoot,
  elementChildren,
  firstElement,
  localName,
  parseXml,
} from "./Xml/XmlReader";

const BUILT_IN_DATE_FORMATS: ReadonlySet<number> = new Set<number>([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);

/** Builds the stylesheet document and the style index maps used by the worksheet writer. */
export function buildStylesheet(model: WorkbookModel): StylesheetSaveContext {
  const defaultStyle = model.DefaultStyle.clone();

  const fonts: FontValue[] = [defaultStyle.Font.clone()];
  const fontIds = new Map<string, number>();
  fontIds.set(getFontKey(fonts[0]), 0);

  const fills: FillValue[] = [toFillValue(defaultStyle), createGray125Fill()];
  const fillIds = new Map<string, number>();
  fillIds.set(getFillKey(fills[0]), 0);
  fillIds.set(getFillKey(fills[1]), 1);

  const borders: BordersValue[] = [defaultStyle.Borders.clone()];
  const borderIds = new Map<string, number>();
  borderIds.set(getBordersKey(borders[0]), 0);

  const customNumberFormatIds = new Map<string, number>();
  const customNumberFormats: CustomNumberFormat[] = [];
  let nextCustomNumberFormatId = 164;
  const defaultNumberFormatId = resolveNumberFormatId(
    defaultStyle.NumberFormat,
    customNumberFormatIds,
    customNumberFormats,
    nextCustomNumberFormatId,
  );
  nextCustomNumberFormatId = defaultNumberFormatId.nextId;

  const normalCellFormat = new CellFormatValue();
  normalCellFormat.NumFmtId = defaultNumberFormatId.id;
  normalCellFormat.FontId = 0;
  normalCellFormat.FillId = 0;
  normalCellFormat.BorderId = 0;
  normalCellFormat.Alignment = defaultStyle.Alignment.clone();
  normalCellFormat.Protection = defaultStyle.Protection.clone();

  const cellFormats: CellFormatValue[] = [normalCellFormat];
  const styleIndices = new Map<string, number>();
  styleIndices.set(getStyleKey(defaultStyle), 0);
  const differentialFormats: StyleValue[] = [];
  const differentialStyleIndices = new Map<string, number>();

  for (const worksheet of model.Worksheets) {
    const metadataResult = addWorksheetMetadataStyles(
      worksheet,
      styleIndices,
      cellFormats,
      fonts,
      fontIds,
      fills,
      fillIds,
      borders,
      borderIds,
      customNumberFormatIds,
      customNumberFormats,
      nextCustomNumberFormatId,
    );
    nextCustomNumberFormatId = metadataResult;

    const persistedEntries = collectPersistedCellEntries(worksheet, model.DefaultStyle);
    for (const entry of persistedEntries) {
      const style = getStyleForSerialization(entry[1]);
      const styleKey = getStyleKey(style);
      if (styleIndices.has(styleKey)) {
        continue;
      }

      const fontId = intern(fontIds, fonts, style.Font, getFontKey);
      const fillId = intern(fillIds, fills, toFillValue(style), getFillKey);
      const borderId = intern(borderIds, borders, style.Borders, getBordersKey);
      const numFmtResult = resolveNumberFormatId(
        style.NumberFormat,
        customNumberFormatIds,
        customNumberFormats,
        nextCustomNumberFormatId,
      );
      nextCustomNumberFormatId = numFmtResult.nextId;

      styleIndices.set(styleKey, cellFormats.length);
      const cellFormat = new CellFormatValue();
      cellFormat.NumFmtId = numFmtResult.id;
      cellFormat.FontId = fontId;
      cellFormat.FillId = fillId;
      cellFormat.BorderId = borderId;
      cellFormat.Alignment = style.Alignment.clone();
      cellFormat.Protection = style.Protection.clone();
      cellFormats.push(cellFormat);
    }

    for (const formatting of worksheet.ConditionalFormattings) {
      for (const condition of formatting.Conditions) {
        const style = condition.Style;
        if (stylesEqual(style, StyleValue.Default)) {
          continue;
        }

        const styleKey = getStyleKey(style);
        if (differentialStyleIndices.has(styleKey)) {
          continue;
        }

        differentialStyleIndices.set(styleKey, differentialFormats.length);
        differentialFormats.push(style.clone());
      }
    }
  }

  const hasStyles =
    styleIndices.size > 1 ||
    !stylesEqual(defaultStyle, StyleValue.Default) ||
    differentialFormats.length > 0;

  return new StylesheetSaveContext(
    buildStylesheetDocument(
      fonts,
      fills,
      borders,
      normalCellFormat,
      cellFormats,
      customNumberFormats,
      differentialFormats,
    ),
    styleIndices,
    differentialStyleIndices,
    differentialFormats.length,
    hasStyles,
  );
}

function addWorksheetMetadataStyles(
  worksheet: WorksheetModel,
  styleIndices: Map<string, number>,
  cellFormats: CellFormatValue[],
  fonts: FontValue[],
  fontIds: Map<string, number>,
  fills: FillValue[],
  fillIds: Map<string, number>,
  borders: BordersValue[],
  borderIds: Map<string, number>,
  customNumberFormatIds: Map<string, number>,
  customNumberFormats: CustomNumberFormat[],
  nextCustomNumberFormatId: number,
): number {
  for (const column of worksheet.Columns) {
    nextCustomNumberFormatId = addMetadataStyle(
      column.Style,
      styleIndices,
      cellFormats,
      fonts,
      fontIds,
      fills,
      fillIds,
      borders,
      borderIds,
      customNumberFormatIds,
      customNumberFormats,
      nextCustomNumberFormatId,
    );
  }

  for (const row of worksheet.Rows.values()) {
    nextCustomNumberFormatId = addMetadataStyle(
      row.Style,
      styleIndices,
      cellFormats,
      fonts,
      fontIds,
      fills,
      fillIds,
      borders,
      borderIds,
      customNumberFormatIds,
      customNumberFormats,
      nextCustomNumberFormatId,
    );
  }

  return nextCustomNumberFormatId;
}

function addMetadataStyle(
  style: StyleValue | null,
  styleIndices: Map<string, number>,
  cellFormats: CellFormatValue[],
  fonts: FontValue[],
  fontIds: Map<string, number>,
  fills: FillValue[],
  fillIds: Map<string, number>,
  borders: BordersValue[],
  borderIds: Map<string, number>,
  customNumberFormatIds: Map<string, number>,
  customNumberFormats: CustomNumberFormat[],
  nextCustomNumberFormatId: number,
): number {
  if (style === null) {
    return nextCustomNumberFormatId;
  }

  const styleKey = getStyleKey(style);
  if (styleIndices.has(styleKey)) {
    return nextCustomNumberFormatId;
  }

  const fontId = intern(fontIds, fonts, style.Font, getFontKey);
  const fillId = intern(fillIds, fills, toFillValue(style), getFillKey);
  const borderId = intern(borderIds, borders, style.Borders, getBordersKey);
  const numFmtResult = resolveNumberFormatId(
    style.NumberFormat,
    customNumberFormatIds,
    customNumberFormats,
    nextCustomNumberFormatId,
  );

  styleIndices.set(styleKey, cellFormats.length);
  const cellFormat = new CellFormatValue();
  cellFormat.NumFmtId = numFmtResult.id;
  cellFormat.FontId = fontId;
  cellFormat.FillId = fillId;
  cellFormat.BorderId = borderId;
  cellFormat.Alignment = style.Alignment.clone();
  cellFormat.Protection = style.Protection.clone();
  cellFormats.push(cellFormat);

  return numFmtResult.nextId;
}

/** Returns the effective style for a cell record, defaulting to the normal style. */
export function getStyleForSerialization(record: CellRecord): StyleValue {
  const style = getEffectiveStyle(record);
  if (record.Kind !== CellValueKind.DateTime || isDateNumberFormat(style.NumberFormat)) {
    return style;
  }

  const serializedStyle = style.clone();
  const numberFormat = new NumberFormatValue();
  numberFormat.Number = 14;
  numberFormat.Custom = null;
  serializedStyle.NumberFormat = numberFormat;
  return serializedStyle;
}

function collectPersistedCellEntries(
  worksheet: WorksheetModel,
  workbookDefaultStyle: StyleValue,
): Array<[CellAddress, CellRecord]> {
  const persistedEntries: Array<[CellAddress, CellRecord]> = [];
  worksheet.Cells.forEach((record, address) => {
    if (shouldPersistCell(workbookDefaultStyle, record)) {
      persistedEntries.push([address, record]);
    }
  });

  persistedEntries.sort(compareCellEntries);
  return persistedEntries;
}

function compareCellEntries(left: [CellAddress, CellRecord], right: [CellAddress, CellRecord]): number {
  const rowComparison = left[0].RowIndex - right[0].RowIndex;
  if (rowComparison !== 0) {
    return rowComparison;
  }

  return left[0].ColumnIndex - right[0].ColumnIndex;
}

function shouldPersistCell(workbookDefaultStyle: StyleValue, record: CellRecord): boolean {
  return (
    record.IsExplicitlyStored ||
    (record.Formula !== null && record.Formula.length > 0) ||
    record.Value !== null ||
    record.Kind !== CellValueKind.Blank ||
    !stylesEqual(record.Style, workbookDefaultStyle)
  );
}

function intern<T>(
  indices: Map<string, number>,
  values: T[],
  value: T,
  getKey: (candidate: T) => string,
): number {
  const key = getKey(value);
  const existing = indices.get(key);
  if (existing !== undefined) {
    return existing;
  }

  const id = values.length;
  values.push(value);
  indices.set(key, id);
  return id;
}

/** Returns the canonical key used to deduplicate styles. */
export function getStyleKey(style: StyleValue): string {
  return (
    getFontKey(style.Font) +
    "|" +
    getFillKey(toFillValue(style)) +
    "|" +
    getBordersKey(style.Borders) +
    "|" +
    String(style.NumberFormat.Number) +
    "|" +
    (style.NumberFormat.Custom ?? "") +
    "|" +
    String(style.Alignment.Horizontal | 0) +
    "|" +
    String(style.Alignment.Vertical | 0) +
    "|" +
    (style.Alignment.WrapText ? "1" : "0") +
    "|" +
    String(style.Alignment.IndentLevel) +
    "|" +
    String(style.Alignment.TextRotation) +
    "|" +
    (style.Alignment.ShrinkToFit ? "1" : "0") +
    "|" +
    String(style.Alignment.ReadingOrder) +
    "|" +
    String(style.Alignment.RelativeIndent) +
    "|" +
    (style.Protection.IsLocked ? "1" : "0") +
    "|" +
    (style.Protection.IsHidden ? "1" : "0")
  );
}

function getFontKey(font: FontValue): string {
  return (
    (font.Name ?? "") +
    "|" +
    formatNumber4(font.Size) +
    "|" +
    (font.Bold ? "1" : "0") +
    "|" +
    (font.Italic ? "1" : "0") +
    "|" +
    String(font.Underline | 0) +
    "|" +
    (font.StrikeThrough ? "1" : "0") +
    "|" +
    getColorKey(font.Color)
  );
}

function getFillKey(fill: FillValue): string {
  return (
    String(fill.Pattern | 0) +
    "|" +
    getColorKey(fill.ForegroundColor) +
    "|" +
    getColorKey(fill.BackgroundColor)
  );
}

function getBordersKey(borders: BordersValue): string {
  return (
    getBorderSideKey(borders.Left) +
    "|" +
    getBorderSideKey(borders.Right) +
    "|" +
    getBorderSideKey(borders.Top) +
    "|" +
    getBorderSideKey(borders.Bottom) +
    "|" +
    getBorderSideKey(borders.Diagonal) +
    "|" +
    (borders.DiagonalUp ? "1" : "0") +
    "|" +
    (borders.DiagonalDown ? "1" : "0")
  );
}

function getBorderSideKey(border: BorderSideValue): string {
  return String(border.Style | 0) + "|" + getColorKey(border.Color);
}

function getColorKey(color: ColorValue): string {
  if (color.ThemeIndex !== null) {
    return "T" + String(color.ThemeIndex) + (color.Tint !== null ? "_" + formatDoubleR(color.Tint) : "");
  }

  if (color.Indexed !== null) {
    return "I" + String(color.Indexed);
  }

  return toArgbHex(color);
}

/** Compares two style values using the canonical style key. */
export function stylesEqual(left: StyleValue, right: StyleValue): boolean {
  return getStyleKey(left) === getStyleKey(right);
}

function getEffectiveStyle(record: CellRecord): StyleValue {
  if (record.Style === null || record.Style === undefined) {
    return StyleValue.Default.clone();
  }

  return record.Style;
}

function toFillValue(style: StyleValue): FillValue {
  const fill = new FillValue();
  fill.Pattern = style.Pattern;
  fill.ForegroundColor = style.ForegroundColor;
  fill.BackgroundColor = style.BackgroundColor;
  return fill;
}

function createGray125Fill(): FillValue {
  const fill = new FillValue();
  fill.Pattern = FillPatternKind.Gray125;
  return fill;
}

interface NumberFormatResolution {
  readonly id: number;
  readonly nextId: number;
}

function resolveNumberFormatId(
  numberFormat: NumberFormatValue,
  customNumberFormatIds: Map<string, number>,
  customNumberFormats: CustomNumberFormat[],
  nextCustomNumberFormatId: number,
): NumberFormatResolution {
  const customFormat = numberFormat.Custom;
  if (customFormat !== null && customFormat.length > 0) {
    const existingId = customNumberFormatIds.get(customFormat);
    if (existingId !== undefined) {
      return { id: existingId, nextId: nextCustomNumberFormatId };
    }

    const customId = nextCustomNumberFormatId;
    nextCustomNumberFormatId++;
    customNumberFormatIds.set(customFormat, customId);
    customNumberFormats.push([customId, customFormat]);
    return { id: customId, nextId: nextCustomNumberFormatId };
  }

  return { id: numberFormat.Number >= 0 ? numberFormat.Number : 0, nextId: nextCustomNumberFormatId };
}

function isDateNumberFormat(numberFormat: NumberFormatValue): boolean {
  if (BUILT_IN_DATE_FORMATS.has(numberFormat.Number)) {
    return true;
  }

  const customFormat = numberFormat.Custom;
  return customFormat !== null && customFormat.length > 0 && looksLikeDateFormat(customFormat);
}

function looksLikeDateFormat(formatCode: string): boolean {
  if (formatCode.trim().length === 0) {
    return false;
  }

  const lower = formatCode.toLowerCase();
  if (lower.includes("[$-f400]") || lower.includes("[$-f800]")) {
    return true;
  }

  let builder = "";
  let inQuote = false;
  let inBracket = false;

  for (let index = 0; index < formatCode.length; index++) {
    const character = formatCode[index];
    if (character === '"') {
      inQuote = !inQuote;
      continue;
    }

    if (inQuote) {
      continue;
    }

    if (character === "[") {
      inBracket = true;
      continue;
    }

    if (character === "]" && inBracket) {
      inBracket = false;
      continue;
    }

    if (inBracket) {
      builder += character.toLowerCase();
      continue;
    }

    if (character === "\\" || character === "_") {
      index++;
      continue;
    }

    if (character === "*") {
      continue;
    }

    builder += character.toLowerCase();
  }

  const normalized = builder;
  return (
    normalized.includes("yy") ||
    normalized.includes("aaa") ||
    (normalized.includes("m") && formatCode.includes("\u6708")) ||
    (normalized.includes("d") && formatCode.includes("\u65e5")) ||
    normalized.includes("yyyy") ||
    normalized.includes("dd") ||
    normalized.includes("ddd") ||
    normalized.includes("mm") ||
    normalized.includes("mmm") ||
    normalized.includes("m/") ||
    normalized.includes("/m") ||
    normalized.includes("h") ||
    normalized.includes("am/pm") ||
    normalized.includes("a/p") ||
    normalized.includes("ss")
  );
}

/** Builds the rich text run properties element (rPr) for a font, omitting an empty font color. */
export function buildRichTextRunProperties(font: FontValue | null): XmlElement {
  const safeFont = font ?? new FontValue();
  const properties = XmlElement.create("rPr");
  if (safeFont.Bold) {
    properties.child(XmlElement.create("b"));
  }

  if (safeFont.Italic) {
    properties.child(XmlElement.create("i"));
  }

  if (safeFont.StrikeThrough) {
    properties.child(XmlElement.create("strike"));
  }

  if (safeFont.Underline !== FontUnderlineType.None) {
    properties.child(buildRichTextUnderlineElement(safeFont.Underline));
  }

  properties.child(XmlElement.create("sz").attr("val", formatNumber4(safeFont.Size)));

  const colorElement = buildRichTextColorElement(safeFont.Color);
  if (colorElement !== null) {
    properties.child(colorElement);
  }

  properties.child(XmlElement.create("rFont").attr("val", safeFont.Name ?? "Calibri"));

  if (safeFont.Family !== null) {
    properties.child(XmlElement.create("family").attr("val", safeFont.Family));
  }

  if (safeFont.Scheme !== FontSchemeType.None) {
    properties.child(
      XmlElement.create("scheme").attr("val", safeFont.Scheme === FontSchemeType.Major ? "major" : "minor"),
    );
  }

  return properties;
}

function buildRichTextUnderlineElement(underlineType: FontUnderlineType): XmlElement {
  switch (underlineType) {
    case FontUnderlineType.Double:
      return XmlElement.create("u").attr("val", "double");
    case FontUnderlineType.Accounting:
      return XmlElement.create("u").attr("val", "singleAccounting");
    case FontUnderlineType.DoubleAccounting:
      return XmlElement.create("u").attr("val", "doubleAccounting");
    default:
      return XmlElement.create("u");
  }
}

function buildRichTextColorElement(color: ColorValue): XmlElement | null {
  if (color.ThemeIndex !== null) {
    const element = XmlElement.create("color").attr("theme", color.ThemeIndex);
    if (color.Tint !== null) {
      element.attr("tint", formatDoubleR(color.Tint));
    }

    return element;
  }

  if (color.Indexed !== null) {
    return XmlElement.create("color").attr("indexed", color.Indexed);
  }

  if (color.A === 0 && color.R === 0 && color.G === 0 && color.B === 0) {
    return null;
  }

  return XmlElement.create("color").attr("rgb", toArgbHex(color));
}

/** Encapsulates the built stylesheet document together with the style index maps. */
export class StylesheetSaveContext {
  private readonly _styleIndices: ReadonlyMap<string, number>;
  private readonly _differentialStyleIndices: ReadonlyMap<string, number>;
  public readonly Document: XmlElement;
  public readonly DifferentialFormatCount: number;
  public readonly HasStyles: boolean;

  public constructor(
    document: XmlElement,
    styleIndices: ReadonlyMap<string, number>,
    differentialStyleIndices: ReadonlyMap<string, number>,
    differentialFormatCount: number,
    hasStyles: boolean,
  ) {
    this.Document = document;
    this._styleIndices = styleIndices;
    this._differentialStyleIndices = differentialStyleIndices;
    this.DifferentialFormatCount = differentialFormatCount;
    this.HasStyles = hasStyles;
  }

  public GetStyleIndex(record: CellRecord): number;
  public GetStyleIndex(style: StyleValue | null): number;
  public GetStyleIndex(recordOrStyle: CellRecord | StyleValue | null): number {
    if (recordOrStyle === null || recordOrStyle === undefined) {
      return 0;
    }

    const style = StylesheetSaveContext.isCellRecord(recordOrStyle)
      ? getStyleForSerialization(recordOrStyle)
      : recordOrStyle;
    const index = this._styleIndices.get(getStyleKey(style));
    return index === undefined ? 0 : index;
  }

  public GetDifferentialStyleIndex(condition: FormatConditionModel): number | null {
    if (stylesEqual(condition.Style, StyleValue.Default)) {
      return null;
    }

    const index = this._differentialStyleIndices.get(getStyleKey(condition.Style));
    return index === undefined ? null : index;
  }

  private static isCellRecord(value: CellRecord | StyleValue): value is CellRecord {
    return (value as CellRecord).Kind !== undefined;
  }
}

const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/** Loads the stylesheet part into a StylesheetLoadContext. */
export function loadStylesheet(
  archive: ZipReader,
  workbookRelationships: ReadonlyMap<string, string>,
  options: LoadOptions,
  diagnostics: LoadDiagnostics,
): StylesheetLoadContext {
  void options;
  void diagnostics;

  const stylesUri = resolveStylesUri(workbookRelationships);
  if (!archive.hasEntry(stylesUri)) {
    return new StylesheetLoadContext();
  }

  const root = documentRoot(parseXml(archive.readText(stylesUri)));
  if (root === null) {
    return new StylesheetLoadContext();
  }

  const customFormats = new Map<number, string>();
  for (const numFmt of elementChildren(firstElement(root, "numFmts"), "numFmt")) {
    const id = attributeInt(numFmt, "numFmtId");
    const code = attribute(numFmt, "formatCode");
    if (id !== null && code !== null && code.length > 0) {
      customFormats.set(id, code);
    }
  }

  const fonts = readFontValues(root);
  if (fonts.length === 0) {
    fonts.push(StyleValue.Default.Font.clone());
  }

  const fills = readFillValues(root);
  if (fills.length === 0) {
    const noneFill = new FillValue();
    noneFill.Pattern = FillPatternKind.None;
    const grayFill = new FillValue();
    grayFill.Pattern = FillPatternKind.Gray125;
    fills.push(noneFill, grayFill);
  }

  const borders = readBordersValues(root);
  if (borders.length === 0) {
    borders.push(StyleValue.Default.Borders.clone());
  }

  const context = new StylesheetLoadContext();
  context.CellFormats.length = 0;
  context.DifferentialFormats.push(...readDifferentialStyleValues(root));

  const tableStyles = firstElement(root, "tableStyles");
  context.RawTableStylesXml = tableStyles === null ? null : serializeRawElement(tableStyles, MAIN_NS);
  for (const differentialFormat of elementChildren(firstElement(root, "dxfs"), "dxf")) {
    context.RawDifferentialFormatsXml.push(serializeRawElement(differentialFormat, MAIN_NS));
  }

  const cellXfs = elementChildren(firstElement(root, "cellXfs"), "xf");
  if (cellXfs.length === 0) {
    context.CellFormats.push(StyleValue.Default.clone());
    context.DefaultCellStyle = context.CellFormats[0]!.clone();
    return context;
  }

  for (let index = 0; index < cellXfs.length; index++) {
    const xf = cellXfs[index]!;
    const style = StyleValue.Default.clone();

    const fontId = attributeInt(xf, "fontId");
    if (fontId !== null && fontId >= 0 && fontId < fonts.length) {
      style.Font = fonts[fontId]!.clone();
    }

    const fillId = attributeInt(xf, "fillId");
    if (fillId !== null && fillId >= 0 && fillId < fills.length) {
      const fill = fills[fillId]!;
      style.Pattern = fill.Pattern;
      style.ForegroundColor = fill.ForegroundColor;
      style.BackgroundColor = fill.BackgroundColor;
    }

    const borderId = attributeInt(xf, "borderId");
    if (borderId !== null && borderId >= 0 && borderId < borders.length) {
      style.Borders = borders[borderId]!.clone();
    }

    const numFmtId = attributeInt(xf, "numFmtId") ?? 0;
    style.NumberFormat = new NumberFormatValue();
    style.NumberFormat.Number = numFmtId;
    style.NumberFormat.Custom = customFormats.get(numFmtId) ?? null;

    const alignment = firstElement(xf, "alignment");
    if (alignment !== null) {
      style.Alignment = new AlignmentValue();
      style.Alignment.Horizontal = parseHorizontalAlignment(attribute(alignment, "horizontal"));
      style.Alignment.Vertical = parseVerticalAlignment(attribute(alignment, "vertical"));
      style.Alignment.WrapText = parseOptionalBoolAttribute(alignment, "wrapText") ?? false;
      style.Alignment.IndentLevel = normalizeIndentLevel(attributeInt(alignment, "indent"));
      style.Alignment.TextRotation = normalizeTextRotation(attributeInt(alignment, "textRotation"));
      style.Alignment.ShrinkToFit = parseOptionalBoolAttribute(alignment, "shrinkToFit") ?? false;
      style.Alignment.ReadingOrder = normalizeReadingOrder(attributeInt(alignment, "readingOrder"));
      style.Alignment.RelativeIndent = attributeInt(alignment, "relativeIndent") ?? 0;
    }

    const protection = firstElement(xf, "protection");
    if (protection !== null) {
      style.Protection = new ProtectionValue();
      style.Protection.IsLocked = parseOptionalBoolAttribute(protection, "locked") ?? true;
      style.Protection.IsHidden = parseOptionalBoolAttribute(protection, "hidden") ?? false;
    }

    context.CellFormats.push(style);
    if (isDateNumberFormat(style.NumberFormat)) {
      context.DateStyleIndexes.add(index);
    }
  }

  if (context.CellFormats.length === 0) {
    context.CellFormats.push(StyleValue.Default.clone());
  }

  context.DefaultCellStyle = context.CellFormats[0]!.clone();
  return context;
}

function normalizeIndentLevel(value: number | null): number {
  if (value === null || value < 0 || value > 250) {
    return 0;
  }

  return value;
}

function normalizeTextRotation(value: number | null): number {
  if (value === null) {
    return 0;
  }

  if (value === 255) {
    return 255;
  }

  if (value < 0 || value > 180) {
    return 0;
  }

  return value;
}

function normalizeReadingOrder(value: number | null): number {
  if (value === null || value < 0 || value > 2) {
    return 0;
  }

  return value;
}

/** Serializes a parsed element without formatting, mirroring XElement.ToString(DisableFormatting). */
function serializeRawElement(node: XmlNode, inheritedNamespace: string): string {
  const element = node as unknown as {
    nodeName: string;
    attributes?: ArrayLike<XmlNode>;
    childNodes?: ArrayLike<XmlNode>;
  };

  const name = element.nodeName;
  let attributes = "";
  const attributeNodes = element.attributes;
  let hasDefaultNamespace = false;
  if (attributeNodes !== undefined) {
    for (let index = 0; index < attributeNodes.length; index++) {
      const item = attributeNodes[index]!;
      if (item.nodeName === "xmlns") {
        hasDefaultNamespace = true;
      }

      attributes += " " + item.nodeName + '="' + escapeAttribute(item.textContent ?? "") + '"';
    }
  }

  if (!hasDefaultNamespace) {
    attributes = ' xmlns="' + escapeAttribute(inheritedNamespace) + '"' + attributes;
  }

  const children = element.childNodes;
  if (children === undefined || children.length === 0) {
    return "<" + name + attributes + " />";
  }

  let inner = "";
  for (let index = 0; index < children.length; index++) {
    const child = children[index]!;
    if (localName(child.nodeName) === "#text" || localName(child.nodeName) === "#cdata-section") {
      inner += escapeText(child.textContent ?? "");
    } else {
      inner += serializeRawElement(child, inheritedNamespace);
    }
  }

  return "<" + name + attributes + ">" + inner + "</" + name + ">";
}
