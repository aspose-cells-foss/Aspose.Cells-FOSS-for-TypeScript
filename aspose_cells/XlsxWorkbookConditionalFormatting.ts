import { CellAddress } from "./Core/CellAddress";
import { CellArea } from "./CellArea";
import { ColorValue } from "./Core/ColorValue";
import { ConditionalFormattingModel } from "./Core/ConditionalFormattingModel";
import { FormatConditionModel } from "./Core/FormatConditionModel";
import { FormatConditionType } from "./FormatConditionType";
import { OperatorType } from "./OperatorType";
import { StyleValue } from "./Core/StyleValue";
import { StylesheetLoadContext } from "./StylesheetLoadContext";
import { WorksheetModel } from "./Core/WorksheetModel";
import { DiagnosticSeverity } from "./enums";
import { InvalidFileFormatException } from "./InvalidFileFormatException";
import { LoadDiagnostics } from "./LoadDiagnostics";
import { LoadIssue } from "./LoadIssue";
import { LoadOptions } from "./LoadOptions";
import { WarningInfo } from "./WarningInfo";
import { XmlElement } from "./Xml/XmlWriter";
import { attribute, elementChildren, firstElement, textOf, XmlNode } from "./Xml/XmlReader";
import { tryParseMergeReference } from "./XlsxWorkbookSerializerCommon";
import { applyDataBarExtension } from "./XlsxDataBarExtensionReader";
import { isEmptyColor, toArgbHex } from "./XlsxWorkbookStylesValueHelpers";
import type { StylesheetSaveContext } from "./XlsxWorkbookStyles";

/** Builds the worksheet "conditionalFormatting" elements, mirroring XlsxWorkbookConditionalFormatting.BuildConditionalFormattingElements. */
export function buildConditionalFormattingElements(
  worksheet: WorksheetModel,
  stylesheet: StylesheetSaveContext,
): XmlElement[] {
  const ordered = getOrderedConditionalFormattings(worksheet.ConditionalFormattings);
  const elements: XmlElement[] = [];

  for (const formatting of ordered) {
    if (formatting.Areas.length === 0 || formatting.Conditions.length === 0) {
      continue;
    }

    const element = XmlElement.create("conditionalFormatting").attr("sqref", buildSqref(formatting.Areas));

    for (let conditionIndex = 0; conditionIndex < formatting.Conditions.length; conditionIndex++) {
      element.child(buildConditionRule(formatting, formatting.Conditions[conditionIndex], conditionIndex, stylesheet));
    }

    elements.push(element);
  }

  return elements;
}

/** Loads conditional formatting from a worksheet root element. */
export function loadConditionalFormattings(
  worksheetModel: WorksheetModel,
  worksheetRoot: XmlNode,
  stylesheet: StylesheetLoadContext,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  worksheetModel.ConditionalFormattings.length = 0;

  for (const formattingElement of elementChildren(worksheetRoot, "conditionalFormatting")) {
    const formatting = new ConditionalFormattingModel();
    const sqref = attribute(formattingElement, "sqref");
    if (!tryLoadAreas(formatting, sqref, diagnostics, options, sheetName)) {
      continue;
    }

    for (const ruleElement of elementChildren(formattingElement, "cfRule")) {
      if (!tryLoadCondition(formatting, ruleElement, stylesheet, diagnostics, options, sheetName, sqref)) {
        continue;
      }

      applyDataBarExtension(worksheetRoot, ruleElement, formatting.Conditions[formatting.Conditions.length - 1]);
    }

    if (formatting.Conditions.length === 0) {
      addIssue(
        diagnostics,
        options,
        issueWithLocation(
          "CF-L003",
          DiagnosticSeverity.LossyRecoverable,
          "A conditional formatting entry without any supported rules was dropped during load.",
          sheetName,
          sqref,
        ),
      );
      continue;
    }

    worksheetModel.ConditionalFormattings.push(formatting);
  }
}

function buildConditionRule(
  formatting: ConditionalFormattingModel,
  condition: FormatConditionModel,
  conditionIndex: number,
  stylesheet: StylesheetSaveContext,
): XmlElement {
  const rule = XmlElement.create("cfRule")
    .attr("type", toRuleTypeName(condition))
    .attr("priority", condition.Priority > 0 ? condition.Priority : conditionIndex + 1);

  const differentialStyleIndex = stylesheet.GetDifferentialStyleIndex(condition);
  if (differentialStyleIndex !== null) {
    rule.attr("dxfId", differentialStyleIndex);
  }

  if (condition.StopIfTrue) {
    rule.boolAttr("stopIfTrue", true);
  }

  switch (condition.Type) {
    case FormatConditionType.CellValue: {
      const operatorName = toOperatorName(condition.Operator);
      if (!isEmptyString(operatorName)) {
        rule.attr("operator", operatorName);
      }

      addFormulaElement(rule, condition.Formula1);
      addFormulaElement(rule, condition.Formula2);
      break;
    }
    case FormatConditionType.Expression:
      addFormulaElement(rule, condition.Formula1);
      break;
    case FormatConditionType.ContainsText:
    case FormatConditionType.NotContainsText:
    case FormatConditionType.BeginsWith:
    case FormatConditionType.EndsWith:
      if (condition.Formula1 !== null && condition.Formula1.length > 0) {
        rule.attr("text", condition.Formula1);
        addFormulaElement(rule, buildTextRuleFormula(condition.Type, condition.Formula1, getAnchorCellReference(formatting)));
      }
      break;
    case FormatConditionType.TimePeriod:
      if (!isEmptyString(condition.TimePeriod)) {
        rule.attr("timePeriod", condition.TimePeriod);
      }
      break;
    case FormatConditionType.Top10:
    case FormatConditionType.Bottom10:
      rule.boolAttr("bottom", condition.Type === FormatConditionType.Bottom10 || !condition.Top);
      if (condition.Percent) {
        rule.boolAttr("percent", true);
      }

      if (condition.Rank > 0) {
        rule.attr("rank", condition.Rank);
      }
      break;
    case FormatConditionType.AboveAverage:
    case FormatConditionType.BelowAverage:
      if (condition.Type === FormatConditionType.BelowAverage || !condition.Above) {
        rule.attr("aboveAverage", 0);
      }

      if (condition.StandardDeviation > 0) {
        rule.attr("stdDev", condition.StandardDeviation);
      }
      break;
    case FormatConditionType.ColorScale:
      rule.child(buildColorScaleElement(condition));
      break;
    case FormatConditionType.DataBar:
      rule.child(buildDataBarElement(condition));
      break;
    case FormatConditionType.IconSet:
      rule.child(buildIconSetElement(condition));
      break;
  }

  return rule;
}

function tryLoadAreas(
  formatting: ConditionalFormattingModel,
  sqref: string | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): boolean {
  if (sqref === null || sqref.trim().length === 0) {
    if (options.StrictMode) {
      throw new InvalidFileFormatException("A conditional formatting entry is missing sqref.");
    }

    addIssue(
      diagnostics,
      options,
      issueWithLocation(
        "CF-L003",
        DiagnosticSeverity.LossyRecoverable,
        "A conditional formatting entry without any valid areas was dropped during load.",
        sheetName,
        null,
      ),
    );
    return false;
  }

  const tokens = sqref.split(/[ \t\r\n]+/).filter((token) => token.length > 0);
  let invalidAreaReported = false;
  for (let index = 0; index < tokens.length; index++) {
    const token = tokens[index];
    const region = tryParseMergeReference(token);
    if (region === null) {
      if (options.StrictMode) {
        throw new InvalidFileFormatException("The conditional formatting sqref '" + token + "' is invalid.");
      }

      if (!invalidAreaReported) {
        addIssue(
          diagnostics,
          options,
          issueWithLocation(
            "CF-L001",
            DiagnosticSeverity.LossyRecoverable,
            "One or more conditional formatting areas were invalid and were dropped during load.",
            sheetName,
            token,
          ),
        );
        invalidAreaReported = true;
      }
      continue;
    }

    formatting.Areas.push(new CellArea(region.FirstRow, region.FirstColumn, region.TotalRows, region.TotalColumns));
  }

  formatting.Areas.sort(compareAreas);
  if (formatting.Areas.length === 0) {
    addIssue(
      diagnostics,
      options,
      issueWithLocation(
        "CF-L003",
        DiagnosticSeverity.LossyRecoverable,
        "A conditional formatting entry without any valid areas was dropped during load.",
        sheetName,
        sqref,
      ),
    );
    return false;
  }

  return true;
}

function tryLoadCondition(
  formatting: ConditionalFormattingModel,
  ruleElement: XmlNode,
  stylesheet: StylesheetLoadContext,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  cellRef: string | null,
): boolean {
  const type = tryParseRuleType(ruleElement, options, diagnostics, sheetName, cellRef);
  if (type === null) {
    return false;
  }

  const condition = new FormatConditionModel();
  condition.Type = type;
  condition.Operator = parseOperatorType(attribute(ruleElement, "operator"));
  condition.StopIfTrue = parseBoolAttribute(ruleElement, "stopIfTrue");
  condition.Priority = parsePriority(ruleElement, formatting, diagnostics, options, sheetName, cellRef);
  condition.Style = resolveDifferentialStyle(ruleElement, stylesheet, diagnostics, options, sheetName, cellRef);
  applyDefaults(condition);

  const formulas = elementChildren(ruleElement, "formula");
  switch (type) {
    case FormatConditionType.Expression:
      condition.Formula1 = formulas.length > 0 ? normalizeFormula(textOf(formulas[0])) : null;
      break;
    case FormatConditionType.CellValue:
      condition.Formula1 = formulas.length > 0 ? normalizeFormula(textOf(formulas[0])) : null;
      condition.Formula2 = formulas.length > 1 ? normalizeFormula(textOf(formulas[1])) : null;
      break;
    case FormatConditionType.ContainsText:
    case FormatConditionType.NotContainsText:
    case FormatConditionType.BeginsWith:
    case FormatConditionType.EndsWith:
      condition.Formula1 = normalizeText(attribute(ruleElement, "text"));
      break;
    case FormatConditionType.TimePeriod:
      condition.TimePeriod = normalizeText(attribute(ruleElement, "timePeriod"));
      break;
    case FormatConditionType.Top10:
    case FormatConditionType.Bottom10:
      condition.Top = type === FormatConditionType.Top10;
      condition.Percent = parseBoolAttribute(ruleElement, "percent");
      condition.Rank = parseNonNegativeIntAttribute(ruleElement, "rank");
      break;
    case FormatConditionType.AboveAverage:
    case FormatConditionType.BelowAverage:
      condition.Above = type === FormatConditionType.AboveAverage;
      condition.StandardDeviation = parseNonNegativeIntAttribute(ruleElement, "stdDev");
      break;
    case FormatConditionType.ColorScale:
      loadColorScale(condition, firstElement(ruleElement, "colorScale"));
      break;
    case FormatConditionType.DataBar:
      loadDataBar(condition, firstElement(ruleElement, "dataBar"));
      break;
    case FormatConditionType.IconSet:
      loadIconSet(condition, firstElement(ruleElement, "iconSet"));
      break;
    case FormatConditionType.DuplicateValues:
      condition.Duplicate = true;
      break;
    case FormatConditionType.UniqueValues:
      condition.Duplicate = false;
      break;
  }

  formatting.Conditions.push(condition);
  return true;
}

function tryParseRuleType(
  ruleElement: XmlNode,
  options: LoadOptions,
  diagnostics: LoadDiagnostics,
  sheetName: string,
  cellRef: string | null,
): FormatConditionType | null {
  const typeText = normalizeToken(attribute(ruleElement, "type"));
  switch (typeText) {
    case "cellis":
      return FormatConditionType.CellValue;
    case "expression":
      return FormatConditionType.Expression;
    case "containstext":
      return FormatConditionType.ContainsText;
    case "notcontainstext":
      return FormatConditionType.NotContainsText;
    case "beginswith":
      return FormatConditionType.BeginsWith;
    case "endswith":
      return FormatConditionType.EndsWith;
    case "timeperiod":
      return FormatConditionType.TimePeriod;
    case "duplicatevalues":
      return FormatConditionType.DuplicateValues;
    case "uniquevalues":
      return FormatConditionType.UniqueValues;
    case "top10":
      return parseBoolAttribute(ruleElement, "bottom") ? FormatConditionType.Bottom10 : FormatConditionType.Top10;
    case "aboveaverage":
      return attribute(ruleElement, "aboveAverage") !== null && !parseBoolAttribute(ruleElement, "aboveAverage")
        ? FormatConditionType.BelowAverage
        : FormatConditionType.AboveAverage;
    case "colorscale":
      return FormatConditionType.ColorScale;
    case "databar":
      return FormatConditionType.DataBar;
    case "iconset":
      return FormatConditionType.IconSet;
    default:
      if (options.StrictMode) {
        throw new InvalidFileFormatException("The conditional formatting rule type '" + (typeText ?? "") + "' is not supported.");
      }

      addIssue(
        diagnostics,
        options,
        issueWithLocation(
          "CF-L002",
          DiagnosticSeverity.LossyRecoverable,
          "An unsupported conditional formatting rule was dropped during load.",
          sheetName,
          cellRef,
        ),
      );
      return null;
  }
}

function applyDefaults(condition: FormatConditionModel): void {
  switch (condition.Type) {
    case FormatConditionType.DuplicateValues:
      condition.Duplicate = true;
      break;
    case FormatConditionType.UniqueValues:
      condition.Duplicate = false;
      break;
    case FormatConditionType.Top10:
      condition.Top = true;
      if (condition.Rank === 0) {
        condition.Rank = 10;
      }
      break;
    case FormatConditionType.Bottom10:
      condition.Top = false;
      if (condition.Rank === 0) {
        condition.Rank = 10;
      }
      break;
    case FormatConditionType.AboveAverage:
      condition.Above = true;
      break;
    case FormatConditionType.BelowAverage:
      condition.Above = false;
      break;
    case FormatConditionType.ColorScale:
      if (condition.ColorScaleCount === 0) {
        condition.ColorScaleCount = 2;
      }
      break;
    case FormatConditionType.DataBar:
      if (isEmptyColor(condition.BarColor)) {
        condition.BarColor = new ColorValue(255, 99, 142, 198);
      }
      break;
    case FormatConditionType.IconSet:
      if (isEmptyString(condition.IconSetType)) {
        condition.IconSetType = "3TrafficLights1";
      }
      break;
  }
}

function parsePriority(
  ruleElement: XmlNode,
  formatting: ConditionalFormattingModel,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  cellRef: string | null,
): number {
  const parsed = parseIntAttribute(ruleElement, "priority");
  if (parsed !== null && parsed > 0) {
    return parsed;
  }

  if (attribute(ruleElement, "priority") !== null) {
    addIssue(
      diagnostics,
      options,
      issueWithLocation(
        "CF-R001",
        DiagnosticSeverity.Recoverable,
        "A conditional formatting priority was invalid and was normalized during load.",
        sheetName,
        cellRef,
        false,
      ),
    );
  }

  return formatting.Conditions.length + 1;
}

function resolveDifferentialStyle(
  ruleElement: XmlNode,
  stylesheet: StylesheetLoadContext,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  cellRef: string | null,
): StyleValue {
  const dxfId = parseIntAttribute(ruleElement, "dxfId");
  if (dxfId === null) {
    return StyleValue.Default.clone();
  }

  if (dxfId >= 0 && dxfId < stylesheet.DifferentialFormats.length) {
    return stylesheet.DifferentialFormats[dxfId].clone();
  }

  if (options.StrictMode) {
    throw new InvalidFileFormatException("The conditional formatting dxfId '" + dxfId.toString() + "' is invalid.");
  }

  addIssue(
    diagnostics,
    options,
    issueWithLocation(
      "CF-R002",
      DiagnosticSeverity.Recoverable,
      "A conditional formatting differential style reference was invalid and the rule style was reset to default during load.",
      sheetName,
      cellRef,
      false,
    ),
  );
  return StyleValue.Default.clone();
}

function loadColorScale(condition: FormatConditionModel, colorScaleElement: XmlNode | null): void {
  if (colorScaleElement === null) {
    return;
  }

  const cfvoCount = elementChildren(colorScaleElement, "cfvo").length;
  condition.ColorScaleCount = cfvoCount >= 3 ? 3 : 2;
  const colors = elementChildren(colorScaleElement, "color");
  if (colors.length > 0) {
    condition.MinColor = readRgbColor(colors[0]);
  }

  if (condition.ColorScaleCount === 3 && colors.length > 2) {
    condition.MidColor = readRgbColor(colors[1]);
    condition.MaxColor = readRgbColor(colors[2]);
    return;
  }

  if (colors.length > 1) {
    condition.MaxColor = readRgbColor(colors[1]);
  }
}

function loadDataBar(condition: FormatConditionModel, dataBarElement: XmlNode | null): void {
  if (dataBarElement === null) {
    return;
  }

  const color = firstElement(dataBarElement, "color");
  if (color !== null) {
    condition.BarColor = readRgbColor(color);
  }
}

function loadIconSet(condition: FormatConditionModel, iconSetElement: XmlNode | null): void {
  if (iconSetElement === null) {
    return;
  }

  condition.IconSetType = normalizeText(attribute(iconSetElement, "iconSet")) ?? "3TrafficLights1";
  condition.ReverseIcons = parseBoolAttribute(iconSetElement, "reverse");
  condition.ShowIconOnly = attribute(iconSetElement, "showValue") !== null && !parseBoolAttribute(iconSetElement, "showValue");
}

function buildColorScaleElement(condition: FormatConditionModel): XmlElement {
  const element = XmlElement.create("colorScale").child(XmlElement.create("cfvo").attr("type", "min"));

  if (condition.ColorScaleCount === 3) {
    element.child(XmlElement.create("cfvo").attr("type", "percentile").attr("val", 50));
  }

  element.child(XmlElement.create("cfvo").attr("type", "max"));
  element.child(buildColorElement(condition.MinColor, new ColorValue(255, 248, 105, 107)));
  if (condition.ColorScaleCount === 3) {
    element.child(buildColorElement(condition.MidColor, new ColorValue(255, 255, 235, 132)));
  }

  element.child(buildColorElement(condition.MaxColor, new ColorValue(255, 99, 190, 123)));
  return element;
}

function buildDataBarElement(condition: FormatConditionModel): XmlElement {
  return XmlElement.create("dataBar")
    .child(XmlElement.create("cfvo").attr("type", "min"))
    .child(XmlElement.create("cfvo").attr("type", "max"))
    .child(buildColorElement(condition.BarColor, new ColorValue(255, 99, 142, 198)));
}

function buildIconSetElement(condition: FormatConditionModel): XmlElement {
  const iconSetType =
    condition.IconSetType !== null && condition.IconSetType.length > 0 ? condition.IconSetType : "3TrafficLights1";
  const element = XmlElement.create("iconSet").attr("iconSet", iconSetType);
  if (condition.ReverseIcons) {
    element.boolAttr("reverse", true);
  }

  if (condition.ShowIconOnly) {
    element.boolAttr("showValue", false);
  }

  const iconCount = getIconCount(iconSetType);
  for (let index = 0; index < iconCount; index++) {
    element.child(XmlElement.create("cfvo").attr("type", "percent").attr("val", Math.trunc((100 * index) / iconCount)));
  }

  return element;
}

function buildColorElement(actual: ColorValue, fallback: ColorValue): XmlElement {
  const color = isEmptyColor(actual) ? fallback : actual;
  return XmlElement.create("color").attr("rgb", toArgbHex(color));
}

function addFormulaElement(parent: XmlElement, formula: string | null): void {
  if (!isEmptyString(formula)) {
    parent.child(XmlElement.create("formula").child(formula));
  }
}

function buildTextRuleFormula(type: FormatConditionType, text: string, firstCell: string): string {
  const escapedText = text.replace(/"/g, '""');
  switch (type) {
    case FormatConditionType.ContainsText:
      return 'NOT(ISERROR(SEARCH("' + escapedText + '",' + firstCell + ")))";
    case FormatConditionType.NotContainsText:
      return 'ISERROR(SEARCH("' + escapedText + '",' + firstCell + "))";
    case FormatConditionType.BeginsWith:
      return "LEFT(" + firstCell + ',LEN("' + escapedText + '"))="' + escapedText + '"';
    case FormatConditionType.EndsWith:
      return "RIGHT(" + firstCell + ',LEN("' + escapedText + '"))="' + escapedText + '"';
    default:
      return "";
  }
}

function getAnchorCellReference(formatting: ConditionalFormattingModel): string {
  if (formatting.Areas.length === 0) {
    return "A1";
  }

  const area = formatting.Areas[0];
  return new CellAddress(area.FirstRow, area.FirstColumn).ToString();
}

function getOrderedConditionalFormattings(collections: ConditionalFormattingModel[]): ConditionalFormattingModel[] {
  const ordered: ConditionalFormattingModel[] = [];
  for (const collection of collections) {
    if (collection.Areas.length === 0 || collection.Conditions.length === 0) {
      continue;
    }

    collection.Areas.sort(compareAreas);
    ordered.push(collection);
  }

  ordered.sort((left, right) => compareAreas(left.Areas[0], right.Areas[0]));
  return ordered;
}

function compareAreas(left: CellArea, right: CellArea): number {
  if (left.StartRow !== right.StartRow) {
    return left.StartRow - right.StartRow;
  }

  if (left.StartColumn !== right.StartColumn) {
    return left.StartColumn - right.StartColumn;
  }

  if (left.TotalRows !== right.TotalRows) {
    return left.TotalRows - right.TotalRows;
  }

  return left.TotalColumns - right.TotalColumns;
}

function buildSqref(areas: CellArea[]): string {
  return areas.map(toAreaReference).join(" ");
}

function toAreaReference(area: CellArea): string {
  const first = new CellAddress(area.FirstRow, area.FirstColumn).ToString();
  if (area.TotalRows === 1 && area.TotalColumns === 1) {
    return first;
  }

  const last = new CellAddress(area.FirstRow + area.TotalRows - 1, area.FirstColumn + area.TotalColumns - 1).ToString();
  return first + ":" + last;
}

function toRuleTypeName(condition: FormatConditionModel): string {
  switch (condition.Type) {
    case FormatConditionType.Expression:
      return "expression";
    case FormatConditionType.ContainsText:
      return "containsText";
    case FormatConditionType.NotContainsText:
      return "notContainsText";
    case FormatConditionType.BeginsWith:
      return "beginsWith";
    case FormatConditionType.EndsWith:
      return "endsWith";
    case FormatConditionType.TimePeriod:
      return "timePeriod";
    case FormatConditionType.DuplicateValues:
      return "duplicateValues";
    case FormatConditionType.UniqueValues:
      return "uniqueValues";
    case FormatConditionType.Top10:
    case FormatConditionType.Bottom10:
      return "top10";
    case FormatConditionType.AboveAverage:
    case FormatConditionType.BelowAverage:
      return "aboveAverage";
    case FormatConditionType.ColorScale:
      return "colorScale";
    case FormatConditionType.DataBar:
      return "dataBar";
    case FormatConditionType.IconSet:
      return "iconSet";
    default:
      return "cellIs";
  }
}

function toOperatorName(type: OperatorType): string | null {
  switch (type) {
    case OperatorType.Between:
      return "between";
    case OperatorType.Equal:
      return "equal";
    case OperatorType.GreaterThan:
      return "greaterThan";
    case OperatorType.GreaterOrEqual:
      return "greaterThanOrEqual";
    case OperatorType.LessThan:
      return "lessThan";
    case OperatorType.LessOrEqual:
      return "lessThanOrEqual";
    case OperatorType.NotBetween:
      return "notBetween";
    case OperatorType.NotEqual:
      return "notEqual";
    default:
      return null;
  }
}

function parseOperatorType(text: string | null): OperatorType {
  switch (normalizeToken(text)) {
    case "between":
      return OperatorType.Between;
    case "equal":
      return OperatorType.Equal;
    case "greaterthan":
      return OperatorType.GreaterThan;
    case "greaterthanorequal":
      return OperatorType.GreaterOrEqual;
    case "lessthan":
      return OperatorType.LessThan;
    case "lessthanorequal":
      return OperatorType.LessOrEqual;
    case "notbetween":
      return OperatorType.NotBetween;
    case "notequal":
      return OperatorType.NotEqual;
    default:
      return OperatorType.None;
  }
}

function parseNonNegativeIntAttribute(node: XmlNode, name: string): number {
  const parsed = parseIntAttribute(node, name);
  return parsed !== null && parsed >= 0 ? parsed : 0;
}

function getIconCount(iconSetType: string): number {
  if (iconSetType.length === 0) {
    return 3;
  }

  if (iconSetType[0] === "4") {
    return 4;
  }

  if (iconSetType[0] === "5") {
    return 5;
  }

  return 3;
}

function normalizeToken(value: string | null): string | null {
  if (value === null) {
    return null;
  }

  const trimmed = value.trim();
  return trimmed.length === 0 ? null : trimmed.toLowerCase();
}

function normalizeFormula(value: string | null): string | null {
  if (value === null) {
    return null;
  }

  const trimmed = value.trim();
  if (trimmed.length === 0) {
    return null;
  }

  return trimmed[0] === "=" ? trimmed.substring(1) : trimmed;
}

function normalizeText(value: string | null): string | null {
  if (value === null) {
    return null;
  }

  const trimmed = value.trim();
  return trimmed.length === 0 ? null : trimmed;
}

function readRgbColor(element: XmlNode): ColorValue {
  const rgb = normalizeText(attribute(element, "rgb"));
  if (rgb === null || rgb.length !== 8) {
    return new ColorValue();
  }

  return new ColorValue(
    Number.parseInt(rgb.substring(0, 2), 16),
    Number.parseInt(rgb.substring(2, 4), 16),
    Number.parseInt(rgb.substring(4, 6), 16),
    Number.parseInt(rgb.substring(6, 8), 16),
  );
}

function isEmptyString(value: string | null | undefined): boolean {
  return value === null || value === undefined || value.length === 0;
}

function parseIntAttribute(node: XmlNode, name: string): number | null {
  const value = attribute(node, name);
  if (value === null || !/^\s*[+-]?\d+\s*$/.test(value)) {
    return null;
  }

  return Number.parseInt(value, 10);
}

function parseBoolAttribute(node: XmlNode, name: string): boolean {
  const value = attribute(node, name);
  if (value === null) {
    return false;
  }

  return value === "1" || value.toLowerCase() === "true";
}

function issueWithLocation(
  code: string,
  severity: DiagnosticSeverity,
  message: string,
  sheetName: string,
  cellRef: string | null,
  dataLossRisk = true,
): LoadIssue {
  const issue = new LoadIssue(code, severity, message, true, dataLossRisk);
  issue.SheetName = sheetName;
  issue.CellRef = cellRef;
  return issue;
}

function addIssue(diagnostics: LoadDiagnostics, options: LoadOptions, issue: LoadIssue): void {
  diagnostics.Add(issue);
  const callback = options.WarningCallback;
  if (callback !== null) {
    const warning = new WarningInfo();
    warning.Code = issue.Code;
    warning.Severity = issue.Severity;
    warning.Message = issue.Message;
    warning.DataLossRisk = issue.DataLossRisk;
    warning.PartUri = issue.PartUri;
    warning.SheetName = issue.SheetName;
    warning.CellRef = issue.CellRef;
    warning.RowIndex = issue.RowIndex;
    callback.Warning(warning);
  }
}
