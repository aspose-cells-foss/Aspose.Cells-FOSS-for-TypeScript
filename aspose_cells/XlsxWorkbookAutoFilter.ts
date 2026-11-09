import { AutoFilterColorFilterModel } from "./Core/AutoFilterColorFilterModel";
import { AutoFilterCustomFilterModel } from "./Core/AutoFilterCustomFilterModel";
import { AutoFilterDynamicFilterModel } from "./Core/AutoFilterDynamicFilterModel";
import { AutoFilterSortConditionModel } from "./Core/AutoFilterSortConditionModel";
import { AutoFilterSortStateModel } from "./Core/AutoFilterSortStateModel";
import { AutoFilterTop10Model } from "./Core/AutoFilterTop10Model";
import { FilterColumnModel } from "./Core/FilterColumnModel";
import { StylesheetLoadContext } from "./StylesheetLoadContext";
import { WorksheetModel } from "./Core/WorksheetModel";
import { DiagnosticSeverity } from "./enums";
import { InvalidFileFormatException } from "./InvalidFileFormatException";
import { LoadDiagnostics } from "./LoadDiagnostics";
import { LoadIssue } from "./LoadIssue";
import { LoadOptions } from "./LoadOptions";
import { WarningInfo } from "./WarningInfo";
import { XmlElement } from "./Xml/XmlWriter";
import { attribute, elementChildren, firstElement, XmlNode } from "./Xml/XmlReader";
import { toRangeReference, tryParseMergeReference } from "./XlsxWorkbookSerializerCommon";
import { formatDoubleR } from "./XlsxWorkbookStylesValueHelpers";

/** Builds the worksheet "autoFilter" element, mirroring XlsxWorkbookAutoFilter.BuildAutoFilterElement. */
export function buildAutoFilterElement(worksheet: WorksheetModel, differentialFormatCount: number): XmlElement | null {
  const autoFilter = worksheet.AutoFilter;
  if (!autoFilter.HasStoredState() || autoFilter.Range === null || autoFilter.Range.length === 0) {
    return null;
  }

  const element = XmlElement.create("autoFilter").attr("ref", autoFilter.Range);
  const orderedColumns = getOrderedFilterColumns(autoFilter.FilterColumns);
  for (const filterColumn of orderedColumns) {
    if (!filterColumn.HasStoredState()) {
      continue;
    }

    const filterColumnElement = buildFilterColumnElement(filterColumn, differentialFormatCount);
    if (filterColumnElement !== null) {
      element.child(filterColumnElement);
    }
  }

  const sortState = buildSortStateElement(autoFilter.SortState, differentialFormatCount);
  if (sortState !== null) {
    element.child(sortState);
  }

  return element;
}

/** Loads the worksheet autoFilter metadata from a worksheet root element. */
export function loadAutoFilter(
  worksheetModel: WorksheetModel,
  worksheetRoot: XmlNode,
  stylesheet: StylesheetLoadContext,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  worksheetModel.AutoFilter.Clear();
  const autoFilterElement = firstElement(worksheetRoot, "autoFilter");
  if (autoFilterElement === null) {
    return;
  }

  const reference = attribute(autoFilterElement, "ref");
  const normalizedRange = tryNormalizeRange(reference);
  if (normalizedRange === null) {
    if (options.StrictMode) {
      throw new InvalidFileFormatException("The worksheet autoFilter ref '" + (reference ?? "") + "' is invalid.");
    }

    addIssue(
      diagnostics,
      options,
      issueWithLocation(
        "WS-L010",
        DiagnosticSeverity.LossyRecoverable,
        "Worksheet autoFilter metadata had an invalid range and was dropped during load.",
        sheetName,
        reference,
      ),
    );
    return;
  }

  worksheetModel.AutoFilter.Range = normalizedRange;

  const seenColumnIndexes = new Set<number>();
  for (const filterColumnElement of elementChildren(autoFilterElement, "filterColumn")) {
    const column = loadFilterColumn(filterColumnElement, stylesheet, diagnostics, options, sheetName, normalizedRange);
    if (column === null) {
      continue;
    }

    if (!seenColumnIndexes.add(column.ColumnIndex)) {
      if (options.StrictMode) {
        throw new InvalidFileFormatException("The worksheet autoFilter contains duplicate filter columns.");
      }

      addIssue(
        diagnostics,
        options,
        issueWithLocation(
          "WS-L011",
          DiagnosticSeverity.LossyRecoverable,
          "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.",
          sheetName,
          normalizedRange,
        ),
      );
      continue;
    }

    worksheetModel.AutoFilter.FilterColumns.push(column);
  }

  worksheetModel.AutoFilter.FilterColumns.sort(compareFilterColumns);

  const sortStateElement = firstElement(autoFilterElement, "sortState");
  if (sortStateElement !== null) {
    loadSortState(worksheetModel.AutoFilter.SortState, sortStateElement, stylesheet, diagnostics, options, sheetName, normalizedRange);
  }
}

function buildFilterColumnElement(model: FilterColumnModel, differentialFormatCount: number): XmlElement | null {
  const customFilters = buildCustomFiltersElement(model);
  const filters = buildFiltersElement(model.Filters);
  const colorFilter = buildColorFilterElement(model.ColorFilter, differentialFormatCount);
  const dynamicFilter = buildDynamicFilterElement(model.DynamicFilter);
  const top10 = buildTop10Element(model.Top10);
  if (
    !model.HiddenButton &&
    filters === null &&
    customFilters === null &&
    colorFilter === null &&
    dynamicFilter === null &&
    top10 === null
  ) {
    return null;
  }

  const element = XmlElement.create("filterColumn").attr("colId", model.ColumnIndex);
  if (model.HiddenButton) {
    element.boolAttr("hiddenButton", true);
  }

  if (colorFilter !== null) {
    element.child(colorFilter);
  }

  if (customFilters !== null) {
    element.child(customFilters);
  }

  if (dynamicFilter !== null) {
    element.child(dynamicFilter);
  }

  if (filters !== null) {
    element.child(filters);
  }

  if (top10 !== null) {
    element.child(top10);
  }

  return element;
}

function buildFiltersElement(filters: string[]): XmlElement | null {
  if (filters.length === 0) {
    return null;
  }

  const element = XmlElement.create("filters");
  for (const filter of filters) {
    element.child(XmlElement.create("filter").attr("val", filter));
  }

  return element;
}

function buildCustomFiltersElement(model: FilterColumnModel): XmlElement | null {
  if (model.CustomFilters.length === 0) {
    return null;
  }

  const element = XmlElement.create("customFilters");
  if (model.CustomFiltersAnd) {
    element.boolAttr("and", true);
  }

  for (const filter of model.CustomFilters) {
    const customFilter = XmlElement.create("customFilter").attr("val", filter.Value);
    if (!isEmptyString(filter.Operator)) {
      customFilter.attr("operator", filter.Operator);
    }

    element.child(customFilter);
  }

  return element;
}

function buildColorFilterElement(model: AutoFilterColorFilterModel, differentialFormatCount: number): XmlElement | null {
  if (!model.Enabled || !hasValidDifferentialStyleId(model.DifferentialStyleId, differentialFormatCount)) {
    return null;
  }

  const element = XmlElement.create("colorFilter").attr("dxfId", model.DifferentialStyleId);
  if (model.CellColor) {
    element.boolAttr("cellColor", true);
  }

  return element;
}

function buildDynamicFilterElement(model: AutoFilterDynamicFilterModel): XmlElement | null {
  if (!model.Enabled) {
    return null;
  }

  const element = XmlElement.create("dynamicFilter");
  if (!isEmptyString(model.Type)) {
    element.attr("type", model.Type);
  }

  if (model.Value !== null) {
    element.attr("val", formatDoubleR(model.Value));
  }

  if (model.MaxValue !== null) {
    element.attr("maxVal", formatDoubleR(model.MaxValue));
  }

  return element;
}

function buildTop10Element(model: AutoFilterTop10Model): XmlElement | null {
  if (!model.Enabled) {
    return null;
  }

  const element = XmlElement.create("top10");
  if (!model.Top) {
    element.attr("top", 0);
  }

  if (model.Percent) {
    element.boolAttr("percent", true);
  }

  if (model.Value !== null) {
    element.attr("val", formatDoubleR(model.Value));
  }

  if (model.FilterValue !== null) {
    element.attr("filterVal", formatDoubleR(model.FilterValue));
  }

  return element;
}

function buildSortStateElement(model: AutoFilterSortStateModel, differentialFormatCount: number): XmlElement | null {
  if (!model.HasStoredState() || model.Ref === null || model.Ref.length === 0) {
    return null;
  }

  const element = XmlElement.create("sortState").attr("ref", model.Ref);
  if (model.ColumnSort) {
    element.boolAttr("columnSort", true);
  }

  if (model.CaseSensitive) {
    element.boolAttr("caseSensitive", true);
  }

  if (!isEmptyString(model.SortMethod)) {
    element.attr("sortMethod", model.SortMethod);
  }

  for (const condition of model.Conditions) {
    if (isEmptyString(condition.Ref)) {
      continue;
    }

    const conditionElement = XmlElement.create("sortCondition").attr("ref", condition.Ref);
    const sortBy = normalizeSortBy(condition.SortBy);
    if (condition.Descending) {
      conditionElement.boolAttr("descending", true);
    }

    if (sortBy.length > 0) {
      conditionElement.attr("sortBy", sortBy);
    }

    if (allowsCustomList(sortBy) && !isEmptyString(condition.CustomList)) {
      conditionElement.attr("customList", condition.CustomList);
    }

    if (allowsDifferentialStyle(sortBy) && hasValidDifferentialStyleId(condition.DifferentialStyleId, differentialFormatCount)) {
      conditionElement.attr("dxfId", condition.DifferentialStyleId);
    }

    if (allowsIconSortCondition(condition, sortBy)) {
      conditionElement.attr("iconSet", condition.IconSet);
      conditionElement.attr("iconId", condition.IconId);
    }

    element.child(conditionElement);
  }

  return element;
}

function getOrderedFilterColumns(columns: FilterColumnModel[]): FilterColumnModel[] {
  return columns.slice().sort(compareFilterColumns);
}

function compareFilterColumns(left: FilterColumnModel, right: FilterColumnModel): number {
  return left.ColumnIndex - right.ColumnIndex;
}

function loadFilterColumn(
  filterColumnElement: XmlNode,
  stylesheet: StylesheetLoadContext,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  autoFilterRange: string,
): FilterColumnModel | null {
  const columnIndex = parseIntAttribute(filterColumnElement, "colId");
  if (columnIndex === null || columnIndex < 0) {
    if (options.StrictMode) {
      throw new InvalidFileFormatException("The worksheet autoFilter contains an invalid filter column id.");
    }

    addIssue(
      diagnostics,
      options,
      issueWithLocation(
        "WS-L011",
        DiagnosticSeverity.LossyRecoverable,
        "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.",
        sheetName,
        autoFilterRange,
      ),
    );
    return null;
  }

  const model = new FilterColumnModel();
  model.ColumnIndex = columnIndex;
  model.HiddenButton = parseFilterColumnHiddenButton(filterColumnElement, diagnostics, options, sheetName, autoFilterRange);
  loadSimpleFilters(model, filterColumnElement, diagnostics, options, sheetName, autoFilterRange);
  loadCustomFilters(model, filterColumnElement, diagnostics, options, sheetName, autoFilterRange);
  loadColorFilter(model.ColorFilter, firstElement(filterColumnElement, "colorFilter"), stylesheet, diagnostics, options, sheetName, autoFilterRange);
  loadDynamicFilter(model.DynamicFilter, firstElement(filterColumnElement, "dynamicFilter"), diagnostics, options, sheetName, autoFilterRange);
  loadTop10(model.Top10, firstElement(filterColumnElement, "top10"), diagnostics, options, sheetName, autoFilterRange);
  return model;
}

function parseFilterColumnHiddenButton(
  filterColumnElement: XmlNode,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  autoFilterRange: string,
): boolean {
  const hiddenButton = tryParseBooleanAttribute(filterColumnElement, "hiddenButton", false);
  if (hiddenButton.success) {
    return hiddenButton.value;
  }

  if (attribute(filterColumnElement, "hiddenButton") !== null) {
    return warnAndReturnFalse(diagnostics, options, sheetName, autoFilterRange);
  }

  const showButton = tryParseBooleanAttribute(filterColumnElement, "showButton", true);
  if (showButton.success) {
    return !showButton.value;
  }

  if (attribute(filterColumnElement, "showButton") !== null) {
    return warnAndReturnFalse(diagnostics, options, sheetName, autoFilterRange);
  }

  return false;
}

function warnAndReturnFalse(
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  autoFilterRange: string,
): boolean {
  if (options.StrictMode) {
    throw new InvalidFileFormatException("The worksheet autoFilter contains an invalid filter column attribute.");
  }

  addIssue(
    diagnostics,
    options,
    issueWithLocation(
      "WS-L011",
      DiagnosticSeverity.LossyRecoverable,
      "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.",
      sheetName,
      autoFilterRange,
    ),
  );
  return false;
}

function loadSimpleFilters(
  model: FilterColumnModel,
  filterColumnElement: XmlNode,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  autoFilterRange: string,
): void {
  const filtersElement = firstElement(filterColumnElement, "filters");
  if (filtersElement === null) {
    return;
  }

  if (elementChildren(filtersElement, "dateGroupItem").length > 0) {
    if (options.StrictMode) {
      throw new InvalidFileFormatException("The worksheet autoFilter contains unsupported date group filters.");
    }

    addIssue(
      diagnostics,
      options,
      issueWithLocation(
        "WS-L011",
        DiagnosticSeverity.LossyRecoverable,
        "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.",
        sheetName,
        autoFilterRange,
      ),
    );
  }

  for (const filterElement of elementChildren(filtersElement, "filter")) {
    const value = attribute(filterElement, "val");
    if (value === null) {
      if (options.StrictMode) {
        throw new InvalidFileFormatException("The worksheet autoFilter contains a filter without a value.");
      }

      addIssue(
        diagnostics,
        options,
        issueWithLocation(
          "WS-L011",
          DiagnosticSeverity.LossyRecoverable,
          "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.",
          sheetName,
          autoFilterRange,
        ),
      );
      continue;
    }

    model.Filters.push(value);
  }
}

function loadCustomFilters(
  model: FilterColumnModel,
  filterColumnElement: XmlNode,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  autoFilterRange: string,
): void {
  const customFiltersElement = firstElement(filterColumnElement, "customFilters");
  if (customFiltersElement === null) {
    return;
  }

  const andValue = tryParseBooleanAttribute(customFiltersElement, "and", false);
  if (andValue.success) {
    model.CustomFiltersAnd = andValue.value;
  } else if (attribute(customFiltersElement, "and") !== null) {
    if (options.StrictMode) {
      throw new InvalidFileFormatException("The worksheet autoFilter contains an invalid custom filter 'and' attribute.");
    }

    addIssue(
      diagnostics,
      options,
      issueWithLocation(
        "WS-L011",
        DiagnosticSeverity.LossyRecoverable,
        "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.",
        sheetName,
        autoFilterRange,
      ),
    );
  }

  for (const customFilterElement of elementChildren(customFiltersElement, "customFilter")) {
    const value = attribute(customFilterElement, "val");
    if (value === null) {
      if (options.StrictMode) {
        throw new InvalidFileFormatException("The worksheet autoFilter contains a custom filter without a value.");
      }

      addIssue(
        diagnostics,
        options,
        issueWithLocation(
          "WS-L011",
          DiagnosticSeverity.LossyRecoverable,
          "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.",
          sheetName,
          autoFilterRange,
        ),
      );
      continue;
    }

    const operatorName = tryParseFilterOperator(attribute(customFilterElement, "operator"));
    if (operatorName === null) {
      if (options.StrictMode) {
        throw new InvalidFileFormatException("The worksheet autoFilter contains an invalid custom filter operator.");
      }

      addIssue(
        diagnostics,
        options,
        issueWithLocation(
          "WS-L011",
          DiagnosticSeverity.LossyRecoverable,
          "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.",
          sheetName,
          autoFilterRange,
        ),
      );
      continue;
    }

    const customFilter = new AutoFilterCustomFilterModel();
    customFilter.Operator = operatorName;
    customFilter.Value = value;
    model.CustomFilters.push(customFilter);
  }
}

function loadColorFilter(
  model: AutoFilterColorFilterModel,
  colorFilterElement: XmlNode | null,
  stylesheet: StylesheetLoadContext,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  autoFilterRange: string,
): void {
  model.Clear();
  if (colorFilterElement === null) {
    return;
  }

  model.Enabled = true;
  const dxfId = parseIntAttribute(colorFilterElement, "dxfId");
  if (attribute(colorFilterElement, "dxfId") !== null) {
    if (hasValidDifferentialStyleId(dxfId, stylesheet.DifferentialFormats.length)) {
      model.DifferentialStyleId = dxfId;
    } else {
      if (options.StrictMode) {
        throw new InvalidFileFormatException("The worksheet autoFilter contains an invalid color filter dxfId.");
      }

      addIssue(
        diagnostics,
        options,
        issueWithLocation(
          "WS-L011",
          DiagnosticSeverity.LossyRecoverable,
          "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.",
          sheetName,
          autoFilterRange,
        ),
      );
      model.Clear();
      return;
    }
  }

  const cellColor = tryParseBooleanAttribute(colorFilterElement, "cellColor", false);
  if (cellColor.success) {
    model.CellColor = cellColor.value;
  } else if (attribute(colorFilterElement, "cellColor") !== null) {
    if (options.StrictMode) {
      throw new InvalidFileFormatException("The worksheet autoFilter contains an invalid color filter cellColor attribute.");
    }

    addIssue(
      diagnostics,
      options,
      issueWithLocation(
        "WS-L011",
        DiagnosticSeverity.LossyRecoverable,
        "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.",
        sheetName,
        autoFilterRange,
      ),
    );
  }
}

function loadDynamicFilter(
  model: AutoFilterDynamicFilterModel,
  dynamicFilterElement: XmlNode | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  autoFilterRange: string,
): void {
  model.Clear();
  if (dynamicFilterElement === null) {
    return;
  }

  model.Enabled = true;
  model.Type = normalizeOptionalText(attribute(dynamicFilterElement, "type"));
  model.Value = parseOptionalDouble(dynamicFilterElement, "val", diagnostics, options, sheetName, autoFilterRange, "dynamic filter value");
  model.MaxValue = parseOptionalDouble(dynamicFilterElement, "maxVal", diagnostics, options, sheetName, autoFilterRange, "dynamic filter max value");
}

function loadTop10(
  model: AutoFilterTop10Model,
  top10Element: XmlNode | null,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  autoFilterRange: string,
): void {
  model.Clear();
  if (top10Element === null) {
    return;
  }

  model.Enabled = true;
  const topValue = tryParseBooleanAttribute(top10Element, "top", true);
  if (topValue.success) {
    model.Top = topValue.value;
  } else if (attribute(top10Element, "top") !== null) {
    if (options.StrictMode) {
      throw new InvalidFileFormatException("The worksheet autoFilter contains an invalid top10 top attribute.");
    }

    addIssue(
      diagnostics,
      options,
      issueWithLocation(
        "WS-L011",
        DiagnosticSeverity.LossyRecoverable,
        "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.",
        sheetName,
        autoFilterRange,
      ),
    );
  }

  const percentValue = tryParseBooleanAttribute(top10Element, "percent", false);
  if (percentValue.success) {
    model.Percent = percentValue.value;
  } else if (attribute(top10Element, "percent") !== null) {
    if (options.StrictMode) {
      throw new InvalidFileFormatException("The worksheet autoFilter contains an invalid top10 percent attribute.");
    }

    addIssue(
      diagnostics,
      options,
      issueWithLocation(
        "WS-L011",
        DiagnosticSeverity.LossyRecoverable,
        "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.",
        sheetName,
        autoFilterRange,
      ),
    );
  }

  model.Value = parseOptionalDouble(top10Element, "val", diagnostics, options, sheetName, autoFilterRange, "top10 value");
  model.FilterValue = parseOptionalDouble(top10Element, "filterVal", diagnostics, options, sheetName, autoFilterRange, "top10 filter value");
}

function loadSortState(
  model: AutoFilterSortStateModel,
  sortStateElement: XmlNode,
  stylesheet: StylesheetLoadContext,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  autoFilterRange: string,
): void {
  model.Clear();
  const reference = attribute(sortStateElement, "ref");
  const normalizedRange = tryNormalizeRange(reference);
  if (normalizedRange === null) {
    if (options.StrictMode) {
      throw new InvalidFileFormatException("The worksheet autoFilter sortState ref is invalid.");
    }

    addIssue(
      diagnostics,
      options,
      issueWithLocation(
        "WS-L012",
        DiagnosticSeverity.LossyRecoverable,
        "Worksheet autoFilter sort state was invalid and was dropped during load.",
        sheetName,
        autoFilterRange,
      ),
    );
    return;
  }

  model.Ref = normalizedRange;
  const columnSort = tryParseBooleanAttribute(sortStateElement, "columnSort", false);
  if (columnSort.success) {
    model.ColumnSort = columnSort.value;
  } else if (attribute(sortStateElement, "columnSort") !== null) {
    warnSortState(diagnostics, options, sheetName, autoFilterRange);
  }

  const caseSensitive = tryParseBooleanAttribute(sortStateElement, "caseSensitive", false);
  if (caseSensitive.success) {
    model.CaseSensitive = caseSensitive.value;
  } else if (attribute(sortStateElement, "caseSensitive") !== null) {
    warnSortState(diagnostics, options, sheetName, autoFilterRange);
  }

  model.SortMethod = normalizeOptionalText(attribute(sortStateElement, "sortMethod"));
  for (const sortConditionElement of elementChildren(sortStateElement, "sortCondition")) {
    const conditionReference = attribute(sortConditionElement, "ref");
    const normalizedConditionRef = tryNormalizeRange(conditionReference);
    if (normalizedConditionRef === null) {
      if (options.StrictMode) {
        throw new InvalidFileFormatException("The worksheet autoFilter sortCondition ref is invalid.");
      }

      warnSortState(diagnostics, options, sheetName, autoFilterRange);
      continue;
    }

    const condition = new AutoFilterSortConditionModel();
    condition.Ref = normalizedConditionRef;
    const descending = tryParseBooleanAttribute(sortConditionElement, "descending", false);
    if (descending.success) {
      condition.Descending = descending.value;
    } else if (attribute(sortConditionElement, "descending") !== null) {
      warnSortState(diagnostics, options, sheetName, autoFilterRange);
    }

    const sortBy = normalizeSortBy(attribute(sortConditionElement, "sortBy"));
    if (attribute(sortConditionElement, "sortBy") !== null && sortBy.length === 0) {
      warnSortState(diagnostics, options, sheetName, autoFilterRange);
    }

    condition.SortBy = sortBy;

    const customList = normalizeOptionalText(attribute(sortConditionElement, "customList"));
    if (allowsCustomList(sortBy)) {
      condition.CustomList = customList.length > 0 ? customList : "";
    } else if (customList.length > 0) {
      warnSortState(diagnostics, options, sheetName, autoFilterRange);
    }

    const iconSet = normalizeOptionalText(attribute(sortConditionElement, "iconSet"));
    if (allowsIconSort(sortBy, iconSet)) {
      condition.IconSet = iconSet.length > 0 ? iconSet : "";
    } else if (iconSet.length > 0) {
      warnSortState(diagnostics, options, sheetName, autoFilterRange);
    }

    const dxfId = parseIntAttribute(sortConditionElement, "dxfId");
    if (attribute(sortConditionElement, "dxfId") !== null) {
      if (allowsDifferentialStyle(sortBy) && hasValidDifferentialStyleId(dxfId, stylesheet.DifferentialFormats.length)) {
        condition.DifferentialStyleId = dxfId;
      } else {
        if (options.StrictMode) {
          throw new InvalidFileFormatException("The worksheet autoFilter sortCondition dxfId is invalid.");
        }

        warnSortState(diagnostics, options, sheetName, autoFilterRange);
      }
    }

    const iconId = parseIntAttribute(sortConditionElement, "iconId");
    if (attribute(sortConditionElement, "iconId") !== null) {
      if (allowsIconSort(sortBy, iconSet) && iconId !== null && iconId >= 0) {
        condition.IconId = iconId;
      } else {
        if (options.StrictMode) {
          throw new InvalidFileFormatException("The worksheet autoFilter sortCondition iconId is invalid.");
        }

        warnSortState(diagnostics, options, sheetName, autoFilterRange);
      }
    }

    model.Conditions.push(condition);
  }
}

function warnSortState(
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  autoFilterRange: string,
): void {
  addIssue(
    diagnostics,
    options,
    issueWithLocation(
      "WS-L012",
      DiagnosticSeverity.LossyRecoverable,
      "Worksheet autoFilter sort state was invalid and was dropped during load.",
      sheetName,
      autoFilterRange,
    ),
  );
}

function tryParseBooleanAttribute(
  node: XmlNode,
  name: string,
  defaultValue: boolean,
): { success: boolean; value: boolean } {
  const value = attribute(node, name);
  if (value === null) {
    return { success: true, value: defaultValue };
  }

  const rawValue = value.trim();
  if (rawValue === "1" || rawValue.toLowerCase() === "true") {
    return { success: true, value: true };
  }

  if (rawValue === "0" || rawValue.toLowerCase() === "false") {
    return { success: true, value: false };
  }

  return { success: false, value: defaultValue };
}

function parseOptionalDouble(
  node: XmlNode,
  name: string,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  autoFilterRange: string,
  attributeLabel: string,
): number | null {
  const value = attribute(node, name);
  if (value === null) {
    return null;
  }

  if (value.trim().length === 0) {
    return invalidDouble(diagnostics, options, sheetName, autoFilterRange, attributeLabel);
  }

  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    return invalidDouble(diagnostics, options, sheetName, autoFilterRange, attributeLabel);
  }

  return parsed;
}

function invalidDouble(
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
  autoFilterRange: string,
  attributeLabel: string,
): number | null {
  if (options.StrictMode) {
    throw new InvalidFileFormatException("The worksheet autoFilter contains an invalid " + attributeLabel + ".");
  }

  addIssue(
    diagnostics,
    options,
    issueWithLocation(
      "WS-L011",
      DiagnosticSeverity.LossyRecoverable,
      "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.",
      sheetName,
      autoFilterRange,
    ),
  );
  return null;
}

function hasValidDifferentialStyleId(
  differentialStyleId: number | null,
  differentialFormatCount: number,
): differentialStyleId is number {
  return differentialStyleId !== null && differentialStyleId >= 0 && differentialStyleId < differentialFormatCount;
}

function normalizeSortBy(sortBy: string | null): string {
  if (sortBy === null || sortBy.trim().length === 0) {
    return "";
  }

  const normalized = sortBy.trim();
  const lower = normalized.toLowerCase();
  if (lower === "value") {
    return "value";
  }

  if (lower === "cellcolor") {
    return "cellColor";
  }

  if (lower === "fontcolor") {
    return "fontColor";
  }

  if (lower === "icon") {
    return "icon";
  }

  return "";
}

function allowsCustomList(sortBy: string): boolean {
  return sortBy.length === 0 || sortBy === "value";
}

function allowsDifferentialStyle(sortBy: string): boolean {
  return sortBy === "cellColor" || sortBy === "fontColor";
}

function allowsIconSortCondition(condition: AutoFilterSortConditionModel, sortBy: string): boolean {
  return (
    sortBy === "icon" &&
    !isEmptyString(condition.IconSet) &&
    condition.IconId !== null &&
    condition.IconId >= 0
  );
}

function allowsIconSort(sortBy: string, iconSet: string): boolean {
  return sortBy === "icon" && iconSet.length > 0;
}

function tryNormalizeRange(value: string | null): string | null {
  if (value === null || value.trim().length === 0) {
    return null;
  }

  const region = tryParseMergeReference(value.trim());
  return region === null ? null : toRangeReference(region);
}

function normalizeOptionalText(value: string | null): string {
  if (value === null || value.trim().length === 0) {
    return "";
  }

  return value.trim();
}

function tryParseFilterOperator(value: string | null): string | null {
  if (value === null || value.trim().length === 0) {
    return "";
  }

  switch (value.trim()) {
    case "lessThan":
      return "lessThan";
    case "lessThanOrEqual":
      return "lessThanOrEqual";
    case "notEqual":
      return "notEqual";
    case "greaterThanOrEqual":
      return "greaterThanOrEqual";
    case "greaterThan":
      return "greaterThan";
    case "equal":
      return "";
    default:
      return null;
  }
}

function parseIntAttribute(node: XmlNode, name: string): number | null {
  const value = attribute(node, name);
  if (value === null || !/^\s*[+-]?\d+\s*$/.test(value)) {
    return null;
  }

  return Number.parseInt(value, 10);
}

function isEmptyString(value: string | null | undefined): boolean {
  return value === null || value === undefined || value.length === 0;
}

function issueWithLocation(
  code: string,
  severity: DiagnosticSeverity,
  message: string,
  sheetName: string,
  cellRef: string | null,
): LoadIssue {
  const issue = new LoadIssue(code, severity, message, true, true);
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
