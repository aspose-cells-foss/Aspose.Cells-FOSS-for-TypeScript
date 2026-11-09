import {
  BorderStyleType,
  CellValueType,
  FillPattern,
  FontUnderlineType,
  HorizontalAlignmentType,
  Style,
  VerticalAlignmentType,
  Workbook,
} from "../../aspose_cells";
import {
  attribute,
  attributeBool,
  descendants,
  documentRoot,
  elementChildren,
  firstElement,
  textOf,
  XmlNode,
} from "../../aspose_cells/Xml/XmlReader";
import { OpenXmlWorkbookReader } from "../Shared/OpenXmlPackage";
import {
  CellSnapshot,
  CaseInsensitiveStringMap,
  OpenXmlBorderDescriptor,
  OpenXmlBorderSideDescriptor,
  OpenXmlCellDescriptor,
  OpenXmlDateSystem,
  OpenXmlFillDescriptor,
  OpenXmlFontDescriptor,
  OpenXmlReadContext,
  OpenXmlStylesContext,
  StyleSnapshot,
} from "./ComparisonModels";
import {
  BuiltInDateFormats,
  buildCellKey,
  getOrDefault,
  looksLikeDateFormat,
  normalizeColor,
  normalizeDouble,
  normalizeFormulaText,
  normalizeOpenXmlColor,
  normalizeValue,
  parseBoolAttribute,
  parseBorderStyle,
  parseDoubleValue,
  parseFillPattern,
  parseHorizontalAlignment,
  parseIntAttribute,
  parseOptionalBoolAttribute,
  parseVerticalAlignment,
  tryNormalizeCellReference,
  tryParseNumber,
} from "./ComparisonValueHelpers";

const TicksPerDay = 864000000000;

/** Reads every cell value snapshot from a workbook using the Open XML reader. */
export function readCellDataWithOpenXmlSdk(workbookPath: string): CaseInsensitiveStringMap<CellSnapshot> {
  return readWithOpenXmlSdk(workbookPath, (context) => {
    const map = new CaseInsensitiveStringMap<CellSnapshot>();
    for (const descriptor of enumerateOpenXmlCells(context)) {
      map.set(
        buildCellKey(descriptor.SheetName, descriptor.CellReference),
        createOpenXmlCellSnapshot(context, descriptor),
      );
    }

    return map;
  });
}

/** Reads every cell style snapshot from a workbook using the Open XML reader. */
export function readCellStylesWithOpenXmlSdk(workbookPath: string): CaseInsensitiveStringMap<StyleSnapshot> {
  return readWithOpenXmlSdk(workbookPath, (context) => {
    const map = new CaseInsensitiveStringMap<StyleSnapshot>();
    for (const descriptor of enumerateOpenXmlCells(context)) {
      const snapshot = createOpenXmlStyleSnapshot(descriptor.Cell, context.Styles).CopyWithLocation(
        descriptor.SheetName,
        descriptor.CellReference,
      );
      map.set(buildCellKey(descriptor.SheetName, descriptor.CellReference), snapshot);
    }

    return map;
  });
}

/** Creates a style snapshot from a library Style facade. */
export function createLibraryStyleSnapshot(
  sheetName: string,
  cellReference: string,
  style: Style,
): StyleSnapshot {
  return new StyleSnapshot(
    sheetName,
    cellReference,
    style.Font.Name,
    normalizeDouble(style.Font.Size),
    style.Font.IsBold,
    style.Font.IsItalic,
    style.Font.Underline !== FontUnderlineType.None,
    style.Font.IsStrikeout,
    normalizeColor(style.Font.Color),
    FillPattern[style.Pattern],
    normalizeColor(style.ForegroundColor),
    normalizeColor(style.BackgroundColor),
    BorderStyleType[style.Borders.Left.LineStyle],
    normalizeColor(style.Borders.Left.Color),
    BorderStyleType[style.Borders.Right.LineStyle],
    normalizeColor(style.Borders.Right.Color),
    BorderStyleType[style.Borders.Top.LineStyle],
    normalizeColor(style.Borders.Top.Color),
    BorderStyleType[style.Borders.Bottom.LineStyle],
    normalizeColor(style.Borders.Bottom.Color),
    BorderStyleType[style.Borders.Diagonal.LineStyle],
    normalizeColor(style.Borders.Diagonal.Color),
    style.Borders.DiagonalUp,
    style.Borders.DiagonalDown,
    HorizontalAlignmentType[style.HorizontalAlignment],
    VerticalAlignmentType[style.VerticalAlignment],
    style.WrapText,
    style.IndentLevel,
    style.TextRotation,
    style.ShrinkToFit,
    style.ReadingOrder,
    style.RelativeIndent,
    style.Number,
    style.Custom ?? "",
    style.IsLocked,
    style.IsHidden,
  );
}

function readWithOpenXmlSdk<T>(workbookPath: string, read: (context: OpenXmlReadContext) => T): T {
  const reader = new OpenXmlWorkbookReader(workbookPath);
  return read(createOpenXmlReadContext(reader));
}

function createOpenXmlReadContext(reader: OpenXmlWorkbookReader): OpenXmlReadContext {
  const sharedStrings = loadSharedStrings(reader);
  const workbookProperties = firstElement(reader.workbookRoot, "workbookPr");
  const dateSystem = attributeBool(workbookProperties, "date1904")
    ? OpenXmlDateSystem.Mac1904
    : OpenXmlDateSystem.Windows1900;
  const styles = loadOpenXmlStylesContext(reader);
  return new OpenXmlReadContext(reader, sharedStrings, dateSystem, styles);
}

function* enumerateOpenXmlCells(context: OpenXmlReadContext): Iterable<OpenXmlCellDescriptor> {
  const sheets = elementChildren(firstElement(context.Reader.workbookRoot, "sheets"), "sheet");
  for (const sheet of sheets) {
    const relationshipId = attribute(sheet, "id");
    if (relationshipId === null) {
      continue;
    }

    const sheetName = attribute(sheet, "name") ?? "<unknown>";
    let partPath: string;
    try {
      partPath = context.Reader.sheetPartPathByName(sheetName);
    } catch {
      continue;
    }

    const worksheetRoot = documentRoot(context.Reader.partXml(partPath));
    if (worksheetRoot === null) {
      continue;
    }

    for (const cell of descendants(worksheetRoot, "c")) {
      const cellReference = tryNormalizeCellReference(attribute(cell, "r"));
      if (cellReference === null) {
        continue;
      }

      yield new OpenXmlCellDescriptor(sheetName, cellReference, cell);
    }
  }
}

function createOpenXmlCellSnapshot(context: OpenXmlReadContext, descriptor: OpenXmlCellDescriptor): CellSnapshot {
  const formula = normalizeFormulaText(textOf(firstElement(descriptor.Cell, "f")));
  const readValue = readOpenXmlValue(
    descriptor.Cell,
    context.SharedStrings,
    context.Styles.DateStyleIndexes,
    context.DateSystem,
  );
  const effectiveType = formula.length === 0 ? readValue.kind : "Formula";
  return new CellSnapshot(
    descriptor.SheetName,
    descriptor.CellReference,
    effectiveType,
    normalizeValue(readValue.value),
    formula,
  );
}

function createOpenXmlStyleSnapshot(cell: XmlNode, styles: OpenXmlStylesContext): StyleSnapshot {
  let styleIndex = parseIntAttribute(attribute(cell, "s")) ?? 0;
  if (styleIndex < 0 || styleIndex >= styles.CellFormats.length) {
    styleIndex = 0;
  }

  return styles.CellFormats[styleIndex];
}

function readOpenXmlValue(
  cell: XmlNode,
  sharedStrings: readonly string[],
  dateStyleIndexes: ReadonlySet<number>,
  dateSystem: OpenXmlDateSystem,
): { value: unknown; kind: string } {
  const cellType = attribute(cell, "t");
  const rawValue = textOf(firstElement(cell, "v"));
  const styleIndex = parseIntAttribute(attribute(cell, "s"));
  const isDateStyle = styleIndex !== null && dateStyleIndexes.has(styleIndex);

  if (cellType === "inlineStr") {
    return { value: readInlineString(firstElement(cell, "is")), kind: "String" };
  }

  if (cellType === "s") {
    const index = parseIntAttribute(rawValue);
    if (index !== null && index >= 0 && index < sharedStrings.length) {
      return { value: sharedStrings[index], kind: "String" };
    }

    return { value: "", kind: "String" };
  }

  if (cellType === "b") {
    return {
      value: rawValue === "1" || rawValue.toLowerCase() === "true",
      kind: "Boolean",
    };
  }

  if (cellType === "d") {
    const parsed = new Date(rawValue);
    if (!Number.isNaN(parsed.getTime())) {
      return { value: parsed, kind: "DateTime" };
    }

    return { value: null, kind: "Blank" };
  }

  if (cellType === "str" || cellType === "e") {
    return { value: rawValue, kind: "String" };
  }

  if (rawValue.length === 0) {
    return { value: null, kind: "Blank" };
  }

  if (isDateStyle) {
    const serial = Number(rawValue.trim());
    if (!Number.isNaN(serial)) {
      return { value: convertSerialToDateTime(serial, dateSystem), kind: "DateTime" };
    }

    return { value: null, kind: "Blank" };
  }

  const numericValue = tryParseNumber(rawValue);
  if (numericValue !== null) {
    return numericValue;
  }

  return { value: null, kind: "Blank" };
}

function readInlineString(inlineString: XmlNode | null): string {
  if (inlineString === null) {
    return "";
  }

  const texts = descendants(inlineString, "t").map((text) => text.textContent ?? "");
  return texts.length === 0 ? textOf(inlineString) : texts.join("");
}

function loadSharedStrings(reader: OpenXmlWorkbookReader): string[] {
  if (!reader.hasPart("xl/sharedStrings.xml")) {
    return [];
  }

  const root = documentRoot(reader.partXml("xl/sharedStrings.xml"));
  if (root === null) {
    return [];
  }

  return elementChildren(root, "si").map((item) => {
    const texts = descendants(item, "t").map((text) => text.textContent ?? "");
    return texts.length === 0 ? textOf(item) : texts.join("");
  });
}

function loadOpenXmlStylesContext(reader: OpenXmlWorkbookReader): OpenXmlStylesContext {
  if (!reader.hasPart("xl/styles.xml")) {
    return OpenXmlStylesContext.Default;
  }

  const root = documentRoot(reader.partXml("xl/styles.xml"));
  if (root === null) {
    return OpenXmlStylesContext.Default;
  }

  const customFormats = new Map<number, string>();
  for (const numFmt of elementChildren(firstElement(root, "numFmts"), "numFmt")) {
    const id = parseIntAttribute(attribute(numFmt, "numFmtId"));
    const formatCode = attribute(numFmt, "formatCode");
    if (id !== null && formatCode !== null && formatCode.length > 0) {
      customFormats.set(id, formatCode);
    }
  }

  const fonts = elementChildren(firstElement(root, "fonts"), "font").map((font) => readOpenXmlFont(font));
  if (fonts.length === 0) {
    fonts.push(OpenXmlFontDescriptor.Default);
  }

  const fills = elementChildren(firstElement(root, "fills"), "fill").map((fill) => readOpenXmlFill(fill));
  if (fills.length === 0) {
    fills.push(OpenXmlFillDescriptor.Default);
  }

  const borders = elementChildren(firstElement(root, "borders"), "border").map((border) =>
    readOpenXmlBorder(border),
  );
  if (borders.length === 0) {
    borders.push(OpenXmlBorderDescriptor.Default);
  }

  const cellFormats = elementChildren(firstElement(root, "cellXfs"), "xf");
  if (cellFormats.length === 0) {
    return new OpenXmlStylesContext([StyleSnapshot.Default], new Set<number>());
  }

  const styleSnapshots: StyleSnapshot[] = [];
  const dateStyleIndexes = new Set<number>();
  for (let index = 0; index < cellFormats.length; index++) {
    const cellFormat = cellFormats[index];
    const fontId = parseIntAttribute(attribute(cellFormat, "fontId")) ?? 0;
    const fillId = parseIntAttribute(attribute(cellFormat, "fillId")) ?? 0;
    const borderId = parseIntAttribute(attribute(cellFormat, "borderId")) ?? 0;
    const numberFormatId = parseIntAttribute(attribute(cellFormat, "numFmtId")) ?? 0;
    const font = getOrDefault(fonts, fontId, OpenXmlFontDescriptor.Default);
    const fill = getOrDefault(fills, fillId, OpenXmlFillDescriptor.Default);
    const border = getOrDefault(borders, borderId, OpenXmlBorderDescriptor.Default);
    const alignmentElement = firstElement(cellFormat, "alignment");
    const protectionElement = firstElement(cellFormat, "protection");
    const customFormatCode = customFormats.get(numberFormatId) ?? "";

    styleSnapshots.push(
      new StyleSnapshot(
        "",
        "",
        font.Name,
        font.Size,
        font.Bold,
        font.Italic,
        font.Underline,
        font.StrikeThrough,
        font.Color,
        fill.Pattern,
        fill.ForegroundColor,
        fill.BackgroundColor,
        border.Left.Style,
        border.Left.Color,
        border.Right.Style,
        border.Right.Color,
        border.Top.Style,
        border.Top.Color,
        border.Bottom.Style,
        border.Bottom.Color,
        border.Diagonal.Style,
        border.Diagonal.Color,
        border.DiagonalUp,
        border.DiagonalDown,
        parseHorizontalAlignment(attribute(alignmentElement, "horizontal")),
        parseVerticalAlignment(attribute(alignmentElement, "vertical")),
        parseBoolAttribute(attribute(alignmentElement, "wrapText")),
        parseIntAttribute(attribute(alignmentElement, "indent")) ?? 0,
        parseIntAttribute(attribute(alignmentElement, "textRotation")) ?? 0,
        parseBoolAttribute(attribute(alignmentElement, "shrinkToFit")),
        parseIntAttribute(attribute(alignmentElement, "readingOrder")) ?? 0,
        parseIntAttribute(attribute(alignmentElement, "relativeIndent")) ?? 0,
        numberFormatId,
        customFormatCode,
        parseOptionalBoolAttribute(attribute(protectionElement, "locked")) ?? true,
        parseBoolAttribute(attribute(protectionElement, "hidden")),
      ),
    );

    if (
      BuiltInDateFormats.has(numberFormatId) ||
      (customFormatCode.length > 0 && looksLikeDateFormat(customFormatCode))
    ) {
      dateStyleIndexes.add(index);
    }
  }

  return new OpenXmlStylesContext(styleSnapshots, dateStyleIndexes);
}

function readOpenXmlFont(font: XmlNode): OpenXmlFontDescriptor {
  return new OpenXmlFontDescriptor(
    attribute(firstElement(font, "name"), "val") ?? "Calibri",
    normalizeDouble(parseDoubleValue(attribute(firstElement(font, "sz"), "val"), 11)),
    firstElement(font, "b") !== null,
    firstElement(font, "i") !== null,
    firstElement(font, "u") !== null,
    firstElement(font, "strike") !== null,
    normalizeOpenXmlColor(firstElement(font, "color")),
  );
}

function readOpenXmlFill(fill: XmlNode): OpenXmlFillDescriptor {
  const patternFill = firstElement(fill, "patternFill");
  if (patternFill === null) {
    return OpenXmlFillDescriptor.Default;
  }

  return new OpenXmlFillDescriptor(
    parseFillPattern(attribute(patternFill, "patternType")),
    normalizeOpenXmlColor(firstElement(patternFill, "fgColor")),
    normalizeOpenXmlColor(firstElement(patternFill, "bgColor")),
  );
}

function readOpenXmlBorder(border: XmlNode): OpenXmlBorderDescriptor {
  return new OpenXmlBorderDescriptor(
    readOpenXmlBorderSide(firstElement(border, "left")),
    readOpenXmlBorderSide(firstElement(border, "right")),
    readOpenXmlBorderSide(firstElement(border, "top")),
    readOpenXmlBorderSide(firstElement(border, "bottom")),
    readOpenXmlBorderSide(firstElement(border, "diagonal")),
    parseBoolAttribute(attribute(border, "diagonalUp")),
    parseBoolAttribute(attribute(border, "diagonalDown")),
  );
}

function readOpenXmlBorderSide(side: XmlNode | null): OpenXmlBorderSideDescriptor {
  if (side === null) {
    return OpenXmlBorderSideDescriptor.Default;
  }

  return new OpenXmlBorderSideDescriptor(
    parseBorderStyle(attribute(side, "style")),
    normalizeOpenXmlColor(firstElement(side, "color")),
  );
}

function convertSerialToDateTime(serial: number, dateSystem: OpenXmlDateSystem): Date {
  let adjusted = serial;
  if (dateSystem === OpenXmlDateSystem.Windows1900 && adjusted >= 60) {
    adjusted -= 1;
  }

  const epochMs =
    dateSystem === OpenXmlDateSystem.Mac1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
  const scaled = adjusted * TicksPerDay;
  const ticks = scaled >= 0 ? Math.floor(scaled + 0.5) : -Math.floor(-scaled + 0.5);
  return new Date(epochMs + ticks / 10000);
}
