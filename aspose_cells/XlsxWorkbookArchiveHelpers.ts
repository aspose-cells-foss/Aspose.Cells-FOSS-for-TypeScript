import { ColorValue } from "./Core/ColorValue";
import { FontSchemeType } from "./Core/FontSchemeType";
import { FontUnderlineType } from "./FontUnderlineType";
import { FontValue } from "./Core/FontValue";
import { RichTextRunValue } from "./Core/RichTextRunValue";
import { SharedStringEntry } from "./SharedStringEntry";
import { DiagnosticSeverity } from "./enums";
import { LoadDiagnostics } from "./LoadDiagnostics";
import { LoadIssue } from "./LoadIssue";
import { LoadOptions } from "./LoadOptions";
import { WarningInfo } from "./WarningInfo";
import { ZipReader } from "./Packaging/Archive";
import {
  BUILT_IN_DATE_FORMATS,
  SHARED_STRINGS_RELATIONSHIP_TYPE,
  STYLES_RELATIONSHIP_TYPE,
  WORKSHEET_RELATIONSHIP_TYPE,
} from "./XlsxWorkbookSerializerCommon";
import {
  XmlNode,
  attribute,
  descendants,
  documentRoot,
  elementChildren,
  firstElement,
  localName,
  parseXml,
  textOf,
} from "./Xml/XmlReader";

/** Resolves a package part URI against a source part URI, mirroring System.Uri resolution. */
export function resolvePartUri(sourceUri: string, target: string): string {
  if (target.startsWith("/")) {
    return target;
  }

  const source = sourceUri.startsWith("/") ? sourceUri : "/" + sourceUri;
  const baseSegments = source.split("/");
  baseSegments.pop();
  const targetSegments = target.replace(/\\/g, "/").split("/");
  for (const segment of targetSegments) {
    if (segment === "" || segment === ".") {
      continue;
    }

    if (segment === "..") {
      if (baseSegments.length > 0) {
        baseSegments.pop();
      }

      continue;
    }

    baseSegments.push(segment);
  }

  const absolute = "/" + baseSegments.join("/");
  return absolute.replace(/\/{2,}/g, "/");
}

/** Returns true when a relationship target matches the expected part URI. */
export function relationshipPointsTo(candidate: string, expectedUri: string): boolean {
  return candidate.toLowerCase() === expectedUri.toLowerCase();
}

/** Classifies a relationship target URI into a known relationship relationship kind. */
export function getRelationshipTargetKind(candidate: string): string {
  const lower = candidate.toLowerCase();
  if (lower.endsWith("/sharedstrings.xml")) {
    return SHARED_STRINGS_RELATIONSHIP_TYPE;
  }

  if (lower.endsWith("/styles.xml")) {
    return STYLES_RELATIONSHIP_TYPE;
  }

  if (lower.includes("/worksheets/")) {
    return WORKSHEET_RELATIONSHIP_TYPE;
  }

  return "";
}

/** Case-insensitive relationship lookup, mirroring Dictionary(StringComparer.OrdinalIgnoreCase). */
export function getRelationshipTarget(
  relationships: ReadonlyMap<string, string>,
  relationshipId: string,
): string | null {
  const direct = relationships.get(relationshipId);
  if (direct !== undefined) {
    return direct;
  }

  const lower = relationshipId.toLowerCase();
  for (const [key, value] of relationships) {
    if (key.toLowerCase() === lower) {
      return value;
    }
  }

  return null;
}

/** Parses a relationships part into a relationship-id to resolved-part-uri map. */
export function loadRelationships(
  archive: ZipReader,
  relationshipsUri: string,
  sourceUri: string,
): Map<string, string> {
  const relationships = new Map<string, string>();
  const text = archive.tryReadText(relationshipsUri);
  if (text === null) {
    return relationships;
  }

  const document = parseXml(text);
  const root = documentRoot(document);
  if (root === null) {
    return relationships;
  }

  for (const relationship of elementChildren(root, "Relationship")) {
    const id = attribute(relationship, "Id");
    const target = attribute(relationship, "Target");
    if (id === null || target === null || id.length === 0 || target.length === 0) {
      continue;
    }

    relationships.set(id, resolvePartUri(sourceUri, target));
  }

  return relationships;
}

/** Parses a part into an XML node, throwing when the part is missing or malformed. */
export function loadDocument(archive: ZipReader, partUri: string): XmlNode {
  return parseXml(archive.readText(partUri));
}

/** Finds the first relationship target that points at the expected part URI. */
export function findRelationshipTarget(
  workbookRelationships: ReadonlyMap<string, string>,
  expectedPath: string,
): string {
  for (const target of workbookRelationships.values()) {
    if (relationshipPointsTo(target, expectedPath)) {
      return target;
    }
  }

  return "";
}

/** Finds the first relationship whose target kind matches the requested relationship type. */
export function findRelationshipTargetByType(
  workbookRelationships: ReadonlyMap<string, string>,
  relationshipType: string,
): string {
  for (const target of workbookRelationships.values()) {
    if (getRelationshipTargetKind(target) === relationshipType) {
      return target;
    }
  }

  return "";
}

/** Resolves the shared strings part URI from workbook relationships using the same fallbacks. */
export function resolveSharedStringsUri(workbookRelationships: ReadonlyMap<string, string>): string {
  let sharedStringsUri = findRelationshipTarget(workbookRelationships, "/xl/sharedStrings.xml");
  if (sharedStringsUri.length === 0) {
    sharedStringsUri =
      workbookRelationships.size === 0
        ? "/xl/sharedStrings.xml"
        : findRelationshipTargetByType(workbookRelationships, SHARED_STRINGS_RELATIONSHIP_TYPE);
  }

  if (sharedStringsUri.length === 0) {
    sharedStringsUri = "/xl/sharedStrings.xml";
  }

  return sharedStringsUri;
}

/** Resolves the styles part URI from workbook relationships using the same fallbacks. */
export function resolveStylesUri(workbookRelationships: ReadonlyMap<string, string>): string {
  let stylesUri = findRelationshipTarget(workbookRelationships, "/xl/styles.xml");
  if (stylesUri.length === 0) {
    stylesUri =
      workbookRelationships.size === 0
        ? "/xl/styles.xml"
        : findRelationshipTargetByType(workbookRelationships, STYLES_RELATIONSHIP_TYPE);
  }

  if (stylesUri.length === 0) {
    stylesUri = "/xl/styles.xml";
  }

  return stylesUri;
}

/** Loads the shared string table entries. */
export function loadSharedStrings(
  archive: ZipReader,
  workbookRelationships: ReadonlyMap<string, string>,
  options: LoadOptions,
  diagnostics: LoadDiagnostics,
): SharedStringEntry[] {
  void options;
  void diagnostics;

  const sharedStringsUri = resolveSharedStringsUri(workbookRelationships);
  if (!archive.hasEntry(sharedStringsUri)) {
    return [];
  }

  const document = loadDocument(archive, sharedStringsUri);
  const root = documentRoot(document);
  if (root === null) {
    return [];
  }

  const items: SharedStringEntry[] = [];
  for (const item of elementChildren(root, "si")) {
    items.push(readSharedStringEntry(item));
  }

  return items;
}

/** Loads the indexes of cell formats whose number format is a date format. */
export function loadDateStyleIndexes(
  archive: ZipReader,
  workbookRelationships: ReadonlyMap<string, string>,
  options: LoadOptions,
  diagnostics: LoadDiagnostics,
): Set<number> {
  void options;
  void diagnostics;

  const stylesUri = resolveStylesUri(workbookRelationships);
  const dateStyleIndexes = new Set<number>();
  if (!archive.hasEntry(stylesUri)) {
    return dateStyleIndexes;
  }

  const root = documentRoot(loadDocument(archive, stylesUri));
  if (root === null) {
    return dateStyleIndexes;
  }

  const customFormats = new Map<number, string>();
  for (const numFmt of elementChildren(firstElementOrNull(root, "numFmts"), "numFmt")) {
    const id = parseIntAttribute(numFmt, "numFmtId");
    const code = attribute(numFmt, "formatCode");
    if (id !== null && code !== null && code.length > 0) {
      customFormats.set(id, code);
    }
  }

  const cellXfs = elementChildren(firstElementOrNull(root, "cellXfs"), "xf");
  for (let index = 0; index < cellXfs.length; index++) {
    const numFmtId = parseIntAttribute(cellXfs[index]!, "numFmtId");
    if (numFmtId === null) {
      continue;
    }

    if (BUILT_IN_DATE_FORMATS.has(numFmtId)) {
      dateStyleIndexes.add(index);
      continue;
    }

    const code = customFormats.get(numFmtId);
    if (code !== undefined && looksLikeDateFormat(code)) {
      dateStyleIndexes.add(index);
    }
  }

  return dateStyleIndexes;
}

/** Parses a boolean attribute with the same "1"/"true" convention as the .NET loader. */
export function parseBoolAttribute(node: XmlNode | null | undefined, name: string): boolean {
  const value = attribute(node, name);
  return value === "1" || (value !== null && value.toLowerCase() === "true");
}

/** Parses an optional Boolean attribute, returning null when the attribute is missing. */
export function parseOptionalBoolAttribute(node: XmlNode | null | undefined, name: string): boolean | null {
  const value = attribute(node, name);
  if (value === null) {
    return null;
  }

  const normalized = value.trim();
  if (normalized.length === 0) {
    return null;
  }

  return normalized === "1" || normalized.toLowerCase() === "true";
}

/** Parses an integer attribute, returning null when the value is missing or invalid. */
export function parseIntAttribute(node: XmlNode | null | undefined, name: string): number | null {
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

/** Parses a floating-point attribute, returning null when the value is missing or invalid. */
export function parseDoubleAttribute(node: XmlNode | null | undefined, name: string): number | null {
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

/** Normalizes a formula by trimming and removing a leading '='. */
export function normalizeFormula(value: string | null | undefined): string {
  if (value === null || value === undefined || value.trim().length === 0) {
    return "";
  }

  const normalizedValue = value.trim();
  return normalizedValue.startsWith("=") ? normalizedValue.substring(1) : normalizedValue;
}

/** Reads the concatenated text of all main-namespace 't' elements, falling back to node text. */
export function readInlineString(inlineStringElement: XmlNode | null | undefined): string {
  if (inlineStringElement === null || inlineStringElement === undefined) {
    return "";
  }

  const textNodes: string[] = [];
  const elements = [
    ...(localName(inlineStringElement.nodeName) === "t" ? [inlineStringElement] : []),
    ...descendants(inlineStringElement, "t"),
  ];
  for (const element of elements) {
    textNodes.push(textOf(element));
  }

  return textNodes.length > 0 ? textNodes.join("") : textOf(inlineStringElement);
}

/** Reads a shared string entry from a shared string or inline string element. */
export function readSharedStringEntry(inlineStringElement: XmlNode | null | undefined): SharedStringEntry {
  const entry = new SharedStringEntry();
  if (inlineStringElement === null || inlineStringElement === undefined) {
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

/** Reads rich text runs from the immediate 'r' child elements. */
export function readRichTextRuns(inlineStringElement: XmlNode | null | undefined): RichTextRunValue[] {
  const runs: RichTextRunValue[] = [];
  if (inlineStringElement === null || inlineStringElement === undefined) {
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
    run.Font = readRichTextFontValue(firstElementOrNull(runElement, "rPr"));
    runs.push(run);
    currentIndex += text.length;
  }

  return runs;
}

/** Reads the rich text run font properties. */
export function readRichTextFontValue(runProperties: XmlNode | null | undefined): FontValue {
  const font = new FontValue();
  if (runProperties === null || runProperties === undefined) {
    return font;
  }

  const rFont = firstElementOrNull(runProperties, "rFont");
  font.Name = attribute(rFont, "val") ?? "Calibri";
  font.Size = parseDoubleAttribute(firstElementOrNull(runProperties, "sz"), "val") ?? 11;
  font.Bold = firstElementOrNull(runProperties, "b") !== null;
  font.Italic = firstElementOrNull(runProperties, "i") !== null;
  font.Underline = parseRichTextUnderlineType(firstElementOrNull(runProperties, "u"));
  font.StrikeThrough = firstElementOrNull(runProperties, "strike") !== null;
  font.Color = readRichTextColorValue(firstElementOrNull(runProperties, "color"));
  font.Family = parseIntAttribute(firstElementOrNull(runProperties, "family"), "val");
  font.Scheme = parseRichTextFontScheme(firstElementOrNull(runProperties, "scheme"));
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

function parseRichTextFontScheme(schemeElement: XmlNode | null): FontSchemeType {
  const value = attribute(schemeElement, "val");
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
    const themeIndex = parseIntStrict(themeAttribute);
    if (themeIndex !== null) {
      const tintAttribute = attribute(colorElement, "tint");
      const tint = tintAttribute === null || tintAttribute.length === 0 ? null : parseDoubleStrict(tintAttribute);
      return new ColorValue(themeIndex, tint);
    }
  }

  const indexedAttribute = attribute(colorElement, "indexed");
  if (indexedAttribute !== null && indexedAttribute.length > 0) {
    const indexed = parseIntStrict(indexedAttribute);
    if (indexed !== null) {
      return new ColorValue(indexed);
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

  if (rgbValue.length !== 8) {
    return new ColorValue();
  }

  const a = parseInt(rgbValue.substring(0, 2), 16);
  const r = parseInt(rgbValue.substring(2, 4), 16);
  const g = parseInt(rgbValue.substring(4, 6), 16);
  const b = parseInt(rgbValue.substring(6, 8), 16);
  if (![a, r, g, b].every((value) => value >= 0 && value <= 255)) {
    return new ColorValue();
  }

  return new ColorValue(a, r, g, b);
}

function parseIntStrict(value: string): number | null {
  const trimmed = value.trim();
  return /^[+-]?\d+$/.test(trimmed) ? parseInt(trimmed, 10) : null;
}

function parseDoubleStrict(value: string): number | null {
  const parsed = Number(value.trim());
  return Number.isFinite(parsed) ? parsed : null;
}

/** Parses a numeric cell value, preferring integral values within the Int32 range. */
export function tryParseNumber(rawValue: string): number | null {
  if (rawValue.length === 0) {
    return null;
  }

  if (!rawValue.includes("e") && !rawValue.includes("E")) {
    const withoutGrouping = rawValue.replace(/,/g, "");
    if (isDecimalNumber(withoutGrouping)) {
      const decimalValue = Number(withoutGrouping);
      if (Number.isFinite(decimalValue)) {
        return decimalValue;
      }
    }
  }

  const doubleValue = rawValue.trim().length === 0 ? Number.NaN : Number(rawValue);
  return Number.isFinite(doubleValue) ? doubleValue : null;
}

function isDecimalNumber(value: string): boolean {
  return /^[+-]?\d+(\.\d+)?$/.test(value.trim());
}

/** Adds a load issue to diagnostics and forwards it to the warning callback. */
export function addIssue(diagnostics: LoadDiagnostics, options: LoadOptions, issue: LoadIssue): void {
  diagnostics.Add(issue);
  if (options.WarningCallback !== null && options.WarningCallback !== undefined) {
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

/** Resolves a worksheet part URI, falling back to the conventional part name with a diagnostic. */
export function resolveWorksheetUri(
  sheetIndex: number,
  relationshipId: string | null,
  workbookRelationships: ReadonlyMap<string, string>,
  archive: ZipReader,
  diagnostics: LoadDiagnostics,
  sheetName: string,
  options: LoadOptions,
): string {
  if (relationshipId !== null && relationshipId.trim().length > 0) {
    const target = getRelationshipTarget(workbookRelationships, relationshipId);
    if (target !== null) {
      return target;
    }
  }

  const fallbackUri = "/xl/worksheets/sheet" + (sheetIndex + 1) + ".xml";
  if (archive.hasEntry(fallbackUri)) {
    const issue = new LoadIssue(
      "PKG-R001",
      DiagnosticSeverity.Recoverable,
      "Worksheet relationship metadata was incomplete; the worksheet part was resolved by convention.",
      true,
    );
    issue.SheetName = sheetName;
    issue.PartUri = fallbackUri;
    addIssue(diagnostics, options, issue);
  }

  return fallbackUri;
}

/** Returns true when a number format code represents a date or time format. */
export function looksLikeDateFormat(formatCode: string | null | undefined): boolean {
  if (formatCode === null || formatCode === undefined || formatCode.trim().length === 0) {
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
    const character = formatCode[index]!;
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
    (normalized.includes("m") && formatCode.includes("月")) ||
    (normalized.includes("d") && formatCode.includes("日")) ||
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

function firstElementOrNull(node: XmlNode | null | undefined, name: string): XmlNode | null {
  return firstElement(node, name);
}
