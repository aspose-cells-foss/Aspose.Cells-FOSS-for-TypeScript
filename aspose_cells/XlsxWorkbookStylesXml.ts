import { AlignmentValue } from "./Core/AlignmentValue";
import { BorderSideValue } from "./Core/BorderSideValue";
import { BordersValue } from "./Core/BordersValue";
import { ColorValue } from "./Core/ColorValue";
import { FillPatternKind } from "./Core/FillPatternKind";
import { FontSchemeType } from "./Core/FontSchemeType";
import { FontUnderlineType } from "./FontUnderlineType";
import { FontValue } from "./Core/FontValue";
import { HorizontalAlignment } from "./Core/HorizontalAlignment";
import { NumberFormatValue } from "./Core/NumberFormatValue";
import { ProtectionValue } from "./Core/ProtectionValue";
import { StyleValue } from "./Core/StyleValue";
import { VerticalAlignment } from "./Core/VerticalAlignment";
import { XmlNode, attribute, elementChildren, firstElement } from "./Xml/XmlReader";
import { XmlElement } from "./Xml/XmlWriter";
import { buildStyleAlignment } from "./XlsxStyleAlignmentWriter";
import {
  CellFormatValue,
  FillValue,
  bordersEqual,
  fontEquals,
  formatDoubleR,
  formatNumber4,
  getBorderStyleName,
  isEmptyColor,
  parseBorderStyle,
  toArgbHex,
} from "./XlsxWorkbookStylesValueHelpers";

const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/** Custom number format entry mirroring a .NET KeyValuePair<int, string>. */
export type CustomNumberFormat = readonly [id: number, code: string];

/** Builds the styleSheet element tree written to xl/styles.xml. */
export function buildStylesheetDocument(
  fonts: readonly FontValue[],
  fills: readonly FillValue[],
  borders: readonly BordersValue[],
  normalCellFormat: CellFormatValue,
  cellFormats: readonly CellFormatValue[],
  customNumberFormats: readonly CustomNumberFormat[],
  differentialFormats: readonly StyleValue[],
): XmlElement {
  const stylesheet = XmlElement.createWithNamespace("styleSheet", MAIN_NS);

  if (customNumberFormats.length > 0) {
    const numFmtElements = customNumberFormats.map((pair) =>
      XmlElement.create("numFmt")
        .attr("numFmtId", pair[0])
        .attr("formatCode", pair[1]),
    );

    stylesheet.child(
      XmlElement.create("numFmts")
        .attr("count", customNumberFormats.length)
        .children(numFmtElements),
    );
  }

  stylesheet.child(
    XmlElement.create("fonts")
      .attr("count", fonts.length)
      .children(buildFontElements(fonts)),
  );
  stylesheet.child(
    XmlElement.create("fills")
      .attr("count", fills.length)
      .children(buildFillElements(fills)),
  );
  stylesheet.child(
    XmlElement.create("borders")
      .attr("count", borders.length)
      .children(buildBorderElements(borders)),
  );
  stylesheet.child(
    XmlElement.create("cellStyleXfs")
      .attr("count", 1)
      .child(buildCellStyleFormatElement(normalCellFormat)),
  );
  stylesheet.child(
    XmlElement.create("cellXfs")
      .attr("count", cellFormats.length)
      .children(buildCellFormatElements(cellFormats)),
  );
  stylesheet.child(
    XmlElement.create("cellStyles")
      .attr("count", 1)
      .child(
        XmlElement.create("cellStyle")
          .attr("name", "Normal")
          .attr("xfId", 0)
          .attr("builtinId", 0),
      ),
  );

  if (differentialFormats.length > 0) {
    stylesheet.child(
      XmlElement.create("dxfs")
        .attr("count", differentialFormats.length)
        .children(buildDifferentialFormatElements(differentialFormats)),
    );
  }

  return stylesheet;
}

export function readFontValues(root: XmlNode): FontValue[] {
  const fonts: FontValue[] = [];
  for (const font of elementChildren(firstElement(root, "fonts"), "font")) {
    fonts.push(readFontValue(font));
  }

  return fonts;
}

export function readFillValues(root: XmlNode): FillValue[] {
  const fills: FillValue[] = [];
  for (const fill of elementChildren(firstElement(root, "fills"), "fill")) {
    fills.push(readFillValue(fill));
  }

  return fills;
}

export function readBordersValues(root: XmlNode): BordersValue[] {
  const borders: BordersValue[] = [];
  for (const border of elementChildren(firstElement(root, "borders"), "border")) {
    borders.push(readBordersValue(border));
  }

  return borders;
}

export function readDifferentialStyleValues(root: XmlNode): StyleValue[] {
  const styles: StyleValue[] = [];
  for (const dxf of elementChildren(firstElement(root, "dxfs"), "dxf")) {
    styles.push(readDifferentialStyleValue(dxf));
  }

  return styles;
}

export function buildFontElements(fonts: readonly FontValue[]): XmlElement[] {
  const elements: XmlElement[] = [];
  for (const font of fonts) {
    elements.push(buildFontElement(font));
  }

  return elements;
}

export function buildFillElements(fills: readonly FillValue[]): XmlElement[] {
  const elements: XmlElement[] = [];
  for (const fill of fills) {
    elements.push(buildFillElement(fill));
  }

  return elements;
}

export function buildBorderElements(borders: readonly BordersValue[]): XmlElement[] {
  const elements: XmlElement[] = [];
  for (const border of borders) {
    elements.push(buildBorderElement(border));
  }

  return elements;
}

export function buildCellFormatElements(cellFormats: readonly CellFormatValue[]): XmlElement[] {
  const elements: XmlElement[] = [];
  for (const cellFormat of cellFormats) {
    elements.push(buildCellFormatElement(cellFormat, true));
  }

  return elements;
}

export function buildDifferentialFormatElements(differentialFormats: readonly StyleValue[]): XmlElement[] {
  const elements: XmlElement[] = [];
  for (const style of differentialFormats) {
    elements.push(buildDifferentialFormatElement(style));
  }

  return elements;
}

export function parseHorizontalAlignment(value: string | null | undefined): HorizontalAlignment {
  switch ((value ?? "").toLowerCase()) {
    case "left":
      return HorizontalAlignment.Left;
    case "center":
      return HorizontalAlignment.Center;
    case "right":
      return HorizontalAlignment.Right;
    case "fill":
      return HorizontalAlignment.Fill;
    case "justify":
      return HorizontalAlignment.Justify;
    case "centercontinuous":
      return HorizontalAlignment.CenterContinuous;
    case "distributed":
      return HorizontalAlignment.Distributed;
    default:
      return HorizontalAlignment.General;
  }
}

export function parseVerticalAlignment(value: string | null | undefined): VerticalAlignment {
  switch ((value ?? "").toLowerCase()) {
    case "center":
      return VerticalAlignment.Center;
    case "top":
      return VerticalAlignment.Top;
    case "justify":
      return VerticalAlignment.Justify;
    case "distributed":
      return VerticalAlignment.Distributed;
    default:
      return VerticalAlignment.Bottom;
  }
}

export function parseOptionalBoolAttribute(node: XmlNode | null | undefined, name: string): boolean | null {
  const value = attribute(node, name);
  if (value === null) {
    return null;
  }

  return parseBoolValue(value);
}

function buildCellStyleFormatElement(cellFormat: CellFormatValue): XmlElement {
  return buildCellFormatElement(cellFormat, false);
}

function readDifferentialStyleValue(dxf: XmlNode): StyleValue {
  const style = StyleValue.Default.clone();
  const font = firstElement(dxf, "font");
  if (font !== null) {
    style.Font = readFontValue(font);
  }

  const fill = firstElement(dxf, "fill");
  if (fill !== null) {
    const fillValue = readFillValue(fill);
    style.Pattern = fillValue.Pattern;
    const patternFill = firstElement(fill, "patternFill");
    if (
      style.Pattern === FillPatternKind.None &&
      patternFill !== null &&
      attribute(patternFill, "patternType") === null &&
      firstElement(patternFill, "bgColor") !== null
    ) {
      style.Pattern = FillPatternKind.Solid;
    }

    // Differential formats store the visible solid fill color in bgColor, reversed from regular cell fills.
    style.ForegroundColor = fillValue.BackgroundColor;
    style.BackgroundColor = fillValue.ForegroundColor;
  }

  const border = firstElement(dxf, "border");
  if (border !== null) {
    style.Borders = readBordersValue(border);
  }

  const numFmt = firstElement(dxf, "numFmt");
  if (numFmt !== null) {
    const numberFormat = new NumberFormatValue();
    numberFormat.Number = parseIntAttribute(numFmt, "numFmtId") ?? 0;
    numberFormat.Custom = attribute(numFmt, "formatCode");
    style.NumberFormat = numberFormat;
  }

  const alignment = firstElement(dxf, "alignment");
  if (alignment !== null) {
    const alignmentValue = new AlignmentValue();
    alignmentValue.Horizontal = parseHorizontalAlignment(attribute(alignment, "horizontal"));
    alignmentValue.Vertical = parseVerticalAlignment(attribute(alignment, "vertical"));
    alignmentValue.WrapText = parseBoolAttribute(alignment, "wrapText");
    alignmentValue.IndentLevel = normalizeIndentLevel(parseIntAttribute(alignment, "indent"));
    alignmentValue.TextRotation = normalizeTextRotation(parseIntAttribute(alignment, "textRotation"));
    alignmentValue.ShrinkToFit = parseBoolAttribute(alignment, "shrinkToFit");
    alignmentValue.ReadingOrder = normalizeReadingOrder(parseIntAttribute(alignment, "readingOrder"));
    alignmentValue.RelativeIndent = parseIntAttribute(alignment, "relativeIndent") ?? 0;
    style.Alignment = alignmentValue;
  }

  const protection = firstElement(dxf, "protection");
  if (protection !== null) {
    const protectionValue = new ProtectionValue();
    protectionValue.IsLocked = parseOptionalBoolAttribute(protection, "locked") ?? true;
    protectionValue.IsHidden = parseBoolAttribute(protection, "hidden");
    style.Protection = protectionValue;
  }

  return style;
}

function readFontValue(font: XmlNode): FontValue {
  const result = new FontValue();
  result.Name = attribute(firstElement(font, "name"), "val") ?? "Calibri";
  result.Size = parseDoubleAttribute(firstElement(font, "sz"), "val") ?? 11;
  result.Bold = firstElement(font, "b") !== null;

  const italic = firstElement(font, "i");
  result.Italic = italic !== null && parseOptionalBoolAttribute(italic, "val") !== false;

  result.Underline = parseUnderlineType(firstElement(font, "u"));
  result.StrikeThrough = firstElement(font, "strike") !== null;
  result.Color = readColorValue(firstElement(font, "color"));
  result.Family = parseIntAttribute(firstElement(font, "family"), "val");
  result.Scheme = parseFontSchemeType(attribute(firstElement(font, "scheme"), "val"));
  return result;
}

function parseFontSchemeType(value: string | null | undefined): FontSchemeType {
  switch (value ?? "") {
    case "major":
      return FontSchemeType.Major;
    case "minor":
      return FontSchemeType.Minor;
    default:
      return FontSchemeType.None;
  }
}

function readFillValue(fill: XmlNode): FillValue {
  const patternFill = firstElement(fill, "patternFill");
  if (patternFill === null) {
    return new FillValue();
  }

  const patternType = (attribute(patternFill, "patternType") ?? "none").toLowerCase();
  const result = new FillValue();
  result.Pattern = parseFillPatternKind(patternType);
  result.ForegroundColor = readColorValue(firstElement(patternFill, "fgColor"));
  result.BackgroundColor = readColorValue(firstElement(patternFill, "bgColor"));
  return result;
}

function parseFillPatternKind(patternType: string): FillPatternKind {
  switch (patternType) {
    case "solid":
      return FillPatternKind.Solid;
    case "mediumgray":
      return FillPatternKind.MediumGray;
    case "darkgray":
      return FillPatternKind.DarkGray;
    case "lightgray":
      return FillPatternKind.LightGray;
    case "gray125":
      return FillPatternKind.Gray125;
    case "gray0625":
      return FillPatternKind.Gray0625;
    case "darkhorizontal":
      return FillPatternKind.DarkHorizontal;
    case "darkvertical":
      return FillPatternKind.DarkVertical;
    case "darkdown":
      return FillPatternKind.DarkDown;
    case "darkup":
      return FillPatternKind.DarkUp;
    case "darkgrid":
      return FillPatternKind.DarkGrid;
    case "darktrellis":
      return FillPatternKind.DarkTrellis;
    case "lighthorizontal":
      return FillPatternKind.LightHorizontal;
    case "lightvertical":
      return FillPatternKind.LightVertical;
    case "lightdown":
      return FillPatternKind.LightDown;
    case "lightup":
      return FillPatternKind.LightUp;
    case "lightgrid":
      return FillPatternKind.LightGrid;
    case "lighttrellis":
      return FillPatternKind.LightTrellis;
    default:
      return FillPatternKind.None;
  }
}

function readBordersValue(border: XmlNode): BordersValue {
  const result = new BordersValue();
  result.Left = readBorderSideValue(firstElement(border, "left"));
  result.Right = readBorderSideValue(firstElement(border, "right"));
  result.Top = readBorderSideValue(firstElement(border, "top"));
  result.Bottom = readBorderSideValue(firstElement(border, "bottom"));
  result.Diagonal = readBorderSideValue(firstElement(border, "diagonal"));
  result.DiagonalUp = parseBoolAttribute(border, "diagonalUp");
  result.DiagonalDown = parseBoolAttribute(border, "diagonalDown");
  return result;
}

function buildFontElement(font: FontValue): XmlElement {
  const element = XmlElement.create("font");
  if (font.Bold) {
    element.child(XmlElement.create("b"));
  }

  if (font.Italic) {
    element.child(XmlElement.create("i"));
  }

  if (font.StrikeThrough) {
    element.child(XmlElement.create("strike"));
  }

  if (font.Underline !== FontUnderlineType.None) {
    element.child(buildUnderlineElement(font.Underline));
  }

  element.child(XmlElement.create("sz").attr("val", formatNumber4(font.Size)));

  const colorElement = buildColorElement("color", font.Color);
  if (colorElement !== null) {
    element.child(colorElement);
  }

  element.child(XmlElement.create("name").attr("val", font.Name));

  if (font.Family !== null) {
    element.child(XmlElement.create("family").attr("val", font.Family));
  }

  if (font.Scheme !== FontSchemeType.None) {
    element.child(
      XmlElement.create("scheme").attr("val", font.Scheme === FontSchemeType.Major ? "major" : "minor"),
    );
  }

  return element;
}

function parseUnderlineType(underlineElement: XmlNode | null): FontUnderlineType {
  if (underlineElement === null) {
    return FontUnderlineType.None;
  }

  const value = attribute(underlineElement, "val");
  if (value === null || value.length === 0) {
    return FontUnderlineType.Single;
  }

  switch (value) {
    case "single":
      return FontUnderlineType.Single;
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

function buildUnderlineElement(underlineType: FontUnderlineType): XmlElement {
  switch (underlineType) {
    case FontUnderlineType.Double:
      return XmlElement.create("u").attr("val", "double");
    case FontUnderlineType.Accounting:
      return XmlElement.create("u").attr("val", "singleAccounting");
    case FontUnderlineType.DoubleAccounting:
      return XmlElement.create("u").attr("val", "doubleAccounting");
    case FontUnderlineType.Single:
      return XmlElement.create("u");
    default:
      return XmlElement.create("u");
  }
}

function buildFillElement(fill: FillValue): XmlElement {
  const patternFill = XmlElement.create("patternFill");
  patternFill.attr("patternType", getFillPatternName(fill.Pattern));

  const foregroundColor = buildColorElement("fgColor", fill.ForegroundColor);
  if (foregroundColor !== null) {
    patternFill.child(foregroundColor);
  }

  const backgroundColor = buildColorElement("bgColor", fill.BackgroundColor);
  if (backgroundColor !== null) {
    patternFill.child(backgroundColor);
  }

  return XmlElement.create("fill").child(patternFill);
}

function getFillPatternName(pattern: FillPatternKind): string {
  switch (pattern) {
    case FillPatternKind.Solid:
      return "solid";
    case FillPatternKind.MediumGray:
      return "mediumGray";
    case FillPatternKind.DarkGray:
      return "darkGray";
    case FillPatternKind.LightGray:
      return "lightGray";
    case FillPatternKind.Gray125:
      return "gray125";
    case FillPatternKind.Gray0625:
      return "gray0625";
    case FillPatternKind.DarkHorizontal:
      return "darkHorizontal";
    case FillPatternKind.DarkVertical:
      return "darkVertical";
    case FillPatternKind.DarkDown:
      return "darkDown";
    case FillPatternKind.DarkUp:
      return "darkUp";
    case FillPatternKind.DarkGrid:
      return "darkGrid";
    case FillPatternKind.DarkTrellis:
      return "darkTrellis";
    case FillPatternKind.LightHorizontal:
      return "lightHorizontal";
    case FillPatternKind.LightVertical:
      return "lightVertical";
    case FillPatternKind.LightDown:
      return "lightDown";
    case FillPatternKind.LightUp:
      return "lightUp";
    case FillPatternKind.LightGrid:
      return "lightGrid";
    case FillPatternKind.LightTrellis:
      return "lightTrellis";
    default:
      return "none";
  }
}

function buildBorderElement(borders: BordersValue): XmlElement {
  const element = XmlElement.create("border").children([
    buildBorderSideElement("left", borders.Left),
    buildBorderSideElement("right", borders.Right),
    buildBorderSideElement("top", borders.Top),
    buildBorderSideElement("bottom", borders.Bottom),
    buildBorderSideElement("diagonal", borders.Diagonal),
  ]);

  if (borders.DiagonalUp) {
    element.attr("diagonalUp", 1);
  }

  if (borders.DiagonalDown) {
    element.attr("diagonalDown", 1);
  }

  return element;
}

function buildCellFormatElement(cellFormat: CellFormatValue, includeXfId: boolean): XmlElement {
  const element = XmlElement.create("xf")
    .attr("numFmtId", cellFormat.NumFmtId)
    .attr("fontId", cellFormat.FontId)
    .attr("fillId", cellFormat.FillId)
    .attr("borderId", cellFormat.BorderId);

  if (includeXfId) {
    element.attr("xfId", 0);
  }

  if (cellFormat.NumFmtId !== 0) {
    element.attr("applyNumberFormat", 1);
  }

  if (cellFormat.FontId !== 0) {
    element.attr("applyFont", 1);
  }

  if (cellFormat.FillId !== 0) {
    element.attr("applyFill", 1);
  }

  if (cellFormat.BorderId !== 0) {
    element.attr("applyBorder", 1);
  }

  const alignmentElement = buildStyleAlignment(cellFormat.Alignment);
  if (alignmentElement !== null) {
    element.attr("applyAlignment", 1);
    element.child(alignmentElement);
  }

  const protectionElement = buildProtectionElement(cellFormat.Protection);
  if (protectionElement !== null) {
    element.attr("applyProtection", 1);
    element.child(protectionElement);
  }

  return element;
}

function buildDifferentialFormatElement(style: StyleValue): XmlElement {
  const element = XmlElement.create("dxf");
  if (!fontEquals(style.Font, StyleValue.Default.Font)) {
    element.child(buildFontElement(style.Font));
  }

  if (
    style.Pattern !== FillPatternKind.None ||
    !isEmptyColor(style.ForegroundColor) ||
    !isEmptyColor(style.BackgroundColor)
  ) {
    const fill = new FillValue();
    // Differential formats store the visible solid fill color in bgColor, reversed from regular cell fills.
    fill.Pattern = style.Pattern;
    fill.ForegroundColor = style.BackgroundColor;
    fill.BackgroundColor = style.ForegroundColor;
    element.child(buildFillElement(fill));
  }

  if (!bordersEqual(style.Borders, StyleValue.Default.Borders)) {
    element.child(buildBorderElement(style.Borders));
  }

  if (style.NumberFormat.Number !== 0 || (style.NumberFormat.Custom !== null && style.NumberFormat.Custom.length > 0)) {
    const numFmtElement = XmlElement.create("numFmt").attr(
      "numFmtId",
      style.NumberFormat.Number >= 0 ? style.NumberFormat.Number : 0,
    );
    if (style.NumberFormat.Custom !== null && style.NumberFormat.Custom.length > 0) {
      numFmtElement.attr("formatCode", style.NumberFormat.Custom);
    }

    element.child(numFmtElement);
  }

  const alignmentElement = buildStyleAlignment(style.Alignment);
  if (alignmentElement !== null) {
    element.child(alignmentElement);
  }

  const protectionElement = buildProtectionElement(style.Protection);
  if (protectionElement !== null) {
    element.child(protectionElement);
  }

  return element;
}

function readBorderSideValue(side: XmlNode | null): BorderSideValue {
  if (side === null) {
    return new BorderSideValue();
  }

  const result = new BorderSideValue();
  result.Style = parseBorderStyle(attribute(side, "style"));
  result.Color = readColorValue(firstElement(side, "color"));
  return result;
}

function readColorValue(colorElement: XmlNode | null): ColorValue {
  if (colorElement === null) {
    return new ColorValue();
  }

  const themeIndex = parseIntAttribute(colorElement, "theme");
  if (themeIndex !== null) {
    let tint: number | null = null;
    const tintValue = parseDoubleAttribute(colorElement, "tint");
    if (tintValue !== null) {
      tint = tintValue;
    }

    return new ColorValue(themeIndex, tint);
  }

  const indexedValue = parseIntAttribute(colorElement, "indexed");
  if (indexedValue !== null) {
    return new ColorValue(indexedValue);
  }

  const rgb = attribute(colorElement, "rgb");
  if (rgb === null || rgb.trim().length === 0) {
    return new ColorValue();
  }

  let rgbValue = rgb.trim().toUpperCase();
  if (rgbValue.length === 6) {
    rgbValue = "FF" + rgbValue;
  }

  if (rgbValue.length !== 8) {
    return new ColorValue();
  }

  const a = parseHexByte(rgbValue.substring(0, 2));
  const r = parseHexByte(rgbValue.substring(2, 4));
  const g = parseHexByte(rgbValue.substring(4, 6));
  const b = parseHexByte(rgbValue.substring(6, 8));
  if (a === null || r === null || g === null || b === null) {
    return new ColorValue();
  }

  return new ColorValue(a, r, g, b);
}

function buildColorElement(elementName: string, color: ColorValue): XmlElement | null {
  if (isEmptyColor(color)) {
    return null;
  }

  if (color.ThemeIndex !== null) {
    const element = XmlElement.create(elementName).attr("theme", color.ThemeIndex);
    if (color.Tint !== null) {
      element.attr("tint", formatDoubleR(color.Tint));
    }

    return element;
  }

  if (color.Indexed !== null) {
    return XmlElement.create(elementName).attr("indexed", color.Indexed);
  }

  return XmlElement.create(elementName).attr("rgb", toArgbHex(color));
}

function buildBorderSideElement(sideName: string, side: BorderSideValue): XmlElement {
  const element = XmlElement.create(sideName);
  const styleName = getBorderStyleName(side.Style);
  if (styleName.length > 0) {
    element.attr("style", styleName);
  }

  const colorElement = buildColorElement("color", side.Color);
  if (colorElement !== null) {
    element.child(colorElement);
  }

  return element;
}

function buildProtectionElement(protection: ProtectionValue): XmlElement | null {
  const element = XmlElement.create("protection");
  let hasValue = false;
  if (!protection.IsLocked) {
    element.attr("locked", 0);
    hasValue = true;
  }

  if (protection.IsHidden) {
    element.attr("hidden", 1);
    hasValue = true;
  }

  return hasValue ? element : null;
}

function parseHexByte(value: string): number | null {
  if (!/^[0-9A-F]+$/i.test(value)) {
    return null;
  }

  const parsed = parseInt(value, 16);
  return Number.isNaN(parsed) ? null : parsed;
}

function parseBoolAttribute(node: XmlNode | null | undefined, name: string): boolean {
  const value = attribute(node, name);
  if (value === null) {
    return false;
  }

  return parseBoolValue(value);
}

function parseBoolValue(value: string): boolean {
  return value === "1" || value.toLowerCase() === "true";
}

function parseIntAttribute(node: XmlNode | null | undefined, name: string): number | null {
  const value = attribute(node, name);
  if (value === null) {
    return null;
  }

  const trimmed = value.trim();
  if (!/^[+-]?\d+$/.test(trimmed)) {
    return null;
  }

  return parseInt(trimmed, 10);
}

function parseDoubleAttribute(node: XmlNode | null | undefined, name: string): number | null {
  const value = attribute(node, name);
  if (value === null) {
    return null;
  }

  const trimmed = value.trim();
  if (trimmed.length === 0) {
    return null;
  }

  const parsed = Number(trimmed);
  return Number.isFinite(parsed) ? parsed : null;
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
