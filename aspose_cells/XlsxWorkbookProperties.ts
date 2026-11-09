import { DateSystem } from "./Core/DateSystem";
import { WorkbookModel } from "./Core/WorkbookModel";
import { DiagnosticSeverity } from "./enums";
import { CellsException } from "./CellsException";
import { LoadDiagnostics } from "./LoadDiagnostics";
import { LoadIssue } from "./LoadIssue";
import { LoadOptions } from "./LoadOptions";
import {
  normalizeCalculationMode,
  normalizeReferenceMode,
  normalizeShowObjects,
  normalizeUpdateLinks,
  normalizeVisibility,
} from "./WorkbookPropertySupport";
import { addIssue, parseIntAttribute } from "./XlsxWorkbookArchiveHelpers";
import { XmlElement } from "./Xml/XmlWriter";
import { XmlNode, attribute, firstElement } from "./Xml/XmlReader";

const WORKBOOK_PART_URI = "/xl/workbook.xml";

/** Builds the workbookPr element, or null when no workbook property state is stored. */
export function buildWorkbookPropertiesElement(model: WorkbookModel): XmlElement | null {
  const element = XmlElement.create("workbookPr");
  let hasState = false;
  if (model.Settings.DateSystem === DateSystem.Mac1904) {
    element.attr("date1904", 1);
    hasState = true;
  }

  const properties = model.Properties;
  if (properties.CodeName.length > 0) {
    element.attr("codeName", properties.CodeName);
    hasState = true;
  }

  if (properties.ShowObjects.length > 0 && properties.ShowObjects.toLowerCase() !== "all") {
    element.attr("showObjects", properties.ShowObjects);
    hasState = true;
  }

  if (properties.FilterPrivacy) {
    element.attr("filterPrivacy", 1);
    hasState = true;
  }

  if (!properties.ShowBorderUnselectedTables) {
    element.attr("showBorderUnselectedTables", 0);
    hasState = true;
  }

  if (!properties.ShowInkAnnotation) {
    element.attr("showInkAnnotation", 0);
    hasState = true;
  }

  if (properties.BackupFile) {
    element.attr("backupFile", 1);
    hasState = true;
  }

  if (!properties.SaveExternalLinkValues) {
    element.attr("saveExternalLinkValues", 0);
    hasState = true;
  }

  if (properties.UpdateLinks.length > 0 && properties.UpdateLinks.toLowerCase() !== "userset") {
    element.attr("updateLinks", properties.UpdateLinks);
    hasState = true;
  }

  if (properties.HidePivotFieldList) {
    element.attr("hidePivotFieldList", 1);
    hasState = true;
  }

  if (properties.DefaultThemeVersion !== null) {
    element.attr("defaultThemeVersion", properties.DefaultThemeVersion);
    hasState = true;
  }

  return hasState ? element : null;
}

/** Builds the workbookProtection element, or null when no protection state is stored. */
export function buildWorkbookProtectionElement(model: WorkbookModel): XmlElement | null {
  const protection = model.Properties.Protection;
  if (!protection.HasStoredState()) {
    return null;
  }

  const element = XmlElement.create("workbookProtection");
  if (protection.LockStructure) {
    element.attr("lockStructure", 1);
  }

  if (protection.LockWindows) {
    element.attr("lockWindows", 1);
  }

  if (protection.LockRevision) {
    element.attr("lockRevision", 1);
  }

  if (protection.WorkbookPassword.length > 0) {
    element.attr("workbookPassword", protection.WorkbookPassword);
  }

  if (protection.RevisionsPassword.length > 0) {
    element.attr("revisionsPassword", protection.RevisionsPassword);
  }

  return element;
}

/** Builds the bookViews element, or null when the workbook view has no stored state. */
export function buildBookViewsElement(model: WorkbookModel): XmlElement | null {
  const view = model.Properties.View;
  if (!view.HasStoredState(model.ActiveSheetIndex)) {
    return null;
  }

  const workbookView = XmlElement.create("workbookView");
  if (view.XWindow !== null) {
    workbookView.attr("xWindow", view.XWindow);
  }

  if (view.YWindow !== null) {
    workbookView.attr("yWindow", view.YWindow);
  }

  if (view.WindowWidth !== null) {
    workbookView.attr("windowWidth", view.WindowWidth);
  }

  if (view.WindowHeight !== null) {
    workbookView.attr("windowHeight", view.WindowHeight);
  }

  if (model.ActiveSheetIndex > 0 && model.ActiveSheetIndex < model.Worksheets.length) {
    workbookView.attr("activeTab", model.ActiveSheetIndex);
  }

  if (view.FirstSheet !== null) {
    let firstSheet = view.FirstSheet;
    if (model.Worksheets.length > 0 && firstSheet >= model.Worksheets.length) {
      firstSheet = model.Worksheets.length - 1;
    }

    if (firstSheet < 0) {
      firstSheet = 0;
    }

    workbookView.attr("firstSheet", firstSheet);
  }

  if (view.ShowHorizontalScroll !== null) {
    workbookView.attr("showHorizontalScroll", view.ShowHorizontalScroll ? 1 : 0);
  }

  if (view.ShowVerticalScroll !== null) {
    workbookView.attr("showVerticalScroll", view.ShowVerticalScroll ? 1 : 0);
  }

  if (view.ShowSheetTabs !== null) {
    workbookView.attr("showSheetTabs", view.ShowSheetTabs ? 1 : 0);
  }

  if (view.TabRatio !== null) {
    workbookView.attr("tabRatio", view.TabRatio);
  }

  if (view.Visibility.length > 0 && view.Visibility.toLowerCase() !== "visible") {
    workbookView.attr("visibility", view.Visibility);
  }

  if (view.Minimized) {
    workbookView.attr("minimized", 1);
  }

  if (!view.AutoFilterDateGrouping) {
    workbookView.attr("autoFilterDateGrouping", 0);
  }

  return XmlElement.create("bookViews").child(workbookView);
}

/** Builds the calcPr element, or null when no calculation state is stored. */
export function buildCalculationPropertiesElement(model: WorkbookModel): XmlElement | null {
  const calculation = model.Properties.Calculation;
  if (!calculation.HasStoredState()) {
    return null;
  }

  const element = XmlElement.create("calcPr");
  if (calculation.CalculationId !== null) {
    element.attr("calcId", calculation.CalculationId);
  }

  if (calculation.CalculationMode.length > 0 && calculation.CalculationMode.toLowerCase() !== "auto") {
    element.attr("calcMode", calculation.CalculationMode);
  }

  if (calculation.FullCalculationOnLoad) {
    element.attr("fullCalcOnLoad", 1);
  }

  if (calculation.ReferenceMode.length > 0 && calculation.ReferenceMode.toLowerCase() !== "a1") {
    element.attr("refMode", calculation.ReferenceMode);
  }

  if (calculation.Iterate) {
    element.attr("iterate", 1);
  }

  if (calculation.IterateCount !== null) {
    element.attr("iterateCount", calculation.IterateCount);
  }

  if (calculation.IterateDelta !== null) {
    element.attr("iterateDelta", formatIterateDelta(calculation.IterateDelta));
  }

  if (calculation.FullPrecision !== null) {
    element.attr("fullPrecision", calculation.FullPrecision ? 1 : 0);
  }

  if (calculation.CalculationCompleted !== null) {
    element.attr("calcCompleted", calculation.CalculationCompleted ? 1 : 0);
  }

  if (calculation.CalculationOnSave !== null) {
    element.attr("calcOnSave", calculation.CalculationOnSave ? 1 : 0);
  }

  if (calculation.ConcurrentCalculation !== null) {
    element.attr("concurrentCalc", calculation.ConcurrentCalculation ? 1 : 0);
  }

  if (calculation.ForceFullCalculation) {
    element.attr("forceFullCalc", 1);
  }

  return element;
}

/** Loads workbookPr, workbookProtection, workbookView and calcPr metadata. */
export function loadWorkbookMetadata(
  workbookRoot: XmlNode,
  workbookModel: WorkbookModel,
  sheetCount: number,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
): void {
  const workbookProperties = workbookModel.Properties;
  const workbookPr = firstElement(workbookRoot, "workbookPr");
  if (workbookPr === null) {
    workbookModel.Settings.DateSystem = DateSystem.Windows1900;
  } else {
    workbookModel.Settings.DateSystem = readBoolAttribute(
      workbookPr,
      "date1904",
      diagnostics,
      options,
      false,
      WORKBOOK_PART_URI,
    )
      ? DateSystem.Mac1904
      : DateSystem.Windows1900;
    workbookProperties.CodeName = readStringAttribute(workbookPr, "codeName");
    workbookProperties.ShowObjects = readChoiceAttribute(
      workbookPr,
      "showObjects",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
      normalizeShowObjects,
    );
    workbookProperties.FilterPrivacy = readBoolAttribute(
      workbookPr,
      "filterPrivacy",
      diagnostics,
      options,
      false,
      WORKBOOK_PART_URI,
    );
    workbookProperties.ShowBorderUnselectedTables = readBoolAttribute(
      workbookPr,
      "showBorderUnselectedTables",
      diagnostics,
      options,
      true,
      WORKBOOK_PART_URI,
    );
    workbookProperties.ShowInkAnnotation = readBoolAttribute(
      workbookPr,
      "showInkAnnotation",
      diagnostics,
      options,
      true,
      WORKBOOK_PART_URI,
    );
    workbookProperties.BackupFile = readBoolAttribute(
      workbookPr,
      "backupFile",
      diagnostics,
      options,
      false,
      WORKBOOK_PART_URI,
    );
    workbookProperties.SaveExternalLinkValues = readBoolAttribute(
      workbookPr,
      "saveExternalLinkValues",
      diagnostics,
      options,
      true,
      WORKBOOK_PART_URI,
    );
    workbookProperties.UpdateLinks = readChoiceAttribute(
      workbookPr,
      "updateLinks",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
      normalizeUpdateLinks,
    );
    workbookProperties.HidePivotFieldList = readBoolAttribute(
      workbookPr,
      "hidePivotFieldList",
      diagnostics,
      options,
      false,
      WORKBOOK_PART_URI,
    );
    workbookProperties.DefaultThemeVersion = readNonNegativeIntAttribute(
      workbookPr,
      "defaultThemeVersion",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
    );
  }

  const protection = firstElement(workbookRoot, "workbookProtection");
  if (protection !== null) {
    workbookProperties.Protection.LockStructure = readBoolAttribute(
      protection,
      "lockStructure",
      diagnostics,
      options,
      false,
      WORKBOOK_PART_URI,
    );
    workbookProperties.Protection.LockWindows = readBoolAttribute(
      protection,
      "lockWindows",
      diagnostics,
      options,
      false,
      WORKBOOK_PART_URI,
    );
    workbookProperties.Protection.LockRevision = readBoolAttribute(
      protection,
      "lockRevision",
      diagnostics,
      options,
      false,
      WORKBOOK_PART_URI,
    );
    workbookProperties.Protection.WorkbookPassword = readStringAttribute(protection, "workbookPassword");
    workbookProperties.Protection.RevisionsPassword = readStringAttribute(protection, "revisionsPassword");
  }

  const workbookView = firstElement(firstElement(workbookRoot, "bookViews"), "workbookView");
  if (workbookView !== null) {
    workbookProperties.View.XWindow = readNonNegativeIntAttribute(
      workbookView,
      "xWindow",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
    );
    workbookProperties.View.YWindow = readNonNegativeIntAttribute(
      workbookView,
      "yWindow",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
    );
    workbookProperties.View.WindowWidth = readNonNegativeIntAttribute(
      workbookView,
      "windowWidth",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
    );
    workbookProperties.View.WindowHeight = readNonNegativeIntAttribute(
      workbookView,
      "windowHeight",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
    );
    let firstSheet = readNonNegativeIntAttribute(
      workbookView,
      "firstSheet",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
    );
    if (firstSheet !== null) {
      if (firstSheet >= sheetCount) {
        addWorkbookMetadataIssue(
          diagnostics,
          options,
          WORKBOOK_PART_URI,
          "Workbook firstSheet exceeded the worksheet count and was clamped.",
        );
        firstSheet = sheetCount > 0 ? sheetCount - 1 : 0;
      }

      workbookProperties.View.FirstSheet = firstSheet;
    }

    workbookProperties.View.ShowHorizontalScroll = readNullableBoolAttribute(
      workbookView,
      "showHorizontalScroll",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
    );
    workbookProperties.View.ShowVerticalScroll = readNullableBoolAttribute(
      workbookView,
      "showVerticalScroll",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
    );
    workbookProperties.View.ShowSheetTabs = readNullableBoolAttribute(
      workbookView,
      "showSheetTabs",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
    );
    const tabRatio = readNonNegativeIntAttribute(
      workbookView,
      "tabRatio",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
    );
    if (tabRatio !== null) {
      if (tabRatio > 1000) {
        addWorkbookMetadataIssue(
          diagnostics,
          options,
          WORKBOOK_PART_URI,
          "Workbook tabRatio was out of range and was ignored.",
        );
      } else {
        workbookProperties.View.TabRatio = tabRatio;
      }
    }

    workbookProperties.View.Visibility = readChoiceAttribute(
      workbookView,
      "visibility",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
      normalizeVisibility,
    );
    workbookProperties.View.Minimized = readBoolAttribute(
      workbookView,
      "minimized",
      diagnostics,
      options,
      false,
      WORKBOOK_PART_URI,
    );
    workbookProperties.View.AutoFilterDateGrouping = readBoolAttribute(
      workbookView,
      "autoFilterDateGrouping",
      diagnostics,
      options,
      true,
      WORKBOOK_PART_URI,
    );

    const activeTab = readNonNegativeIntAttribute(
      workbookView,
      "activeTab",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
    );
    if (activeTab !== null) {
      if (activeTab >= sheetCount) {
        addWorkbookMetadataIssue(
          diagnostics,
          options,
          WORKBOOK_PART_URI,
          "Workbook activeTab exceeded the worksheet count and was ignored.",
        );
      } else {
        workbookModel.ActiveSheetIndex = activeTab;
      }
    }
  }

  const calcPr = firstElement(workbookRoot, "calcPr");
  if (calcPr !== null) {
    const calculation = workbookProperties.Calculation;
    calculation.CalculationId = readNonNegativeIntAttribute(
      calcPr,
      "calcId",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
    );
    calculation.CalculationMode = readChoiceAttribute(
      calcPr,
      "calcMode",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
      normalizeCalculationMode,
    );
    calculation.FullCalculationOnLoad = readBoolAttribute(
      calcPr,
      "fullCalcOnLoad",
      diagnostics,
      options,
      false,
      WORKBOOK_PART_URI,
    );
    calculation.ReferenceMode = readChoiceAttribute(
      calcPr,
      "refMode",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
      normalizeReferenceMode,
    );
    calculation.Iterate = readBoolAttribute(
      calcPr,
      "iterate",
      diagnostics,
      options,
      false,
      WORKBOOK_PART_URI,
    );
    calculation.IterateCount = readNonNegativeIntAttribute(
      calcPr,
      "iterateCount",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
    );
    calculation.IterateDelta = readNonNegativeDoubleAttribute(
      calcPr,
      "iterateDelta",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
    );
    calculation.FullPrecision = readNullableBoolAttribute(
      calcPr,
      "fullPrecision",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
    );
    calculation.CalculationCompleted = readNullableBoolAttribute(
      calcPr,
      "calcCompleted",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
    );
    calculation.CalculationOnSave = readNullableBoolAttribute(
      calcPr,
      "calcOnSave",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
    );
    calculation.ConcurrentCalculation = readNullableBoolAttribute(
      calcPr,
      "concurrentCalc",
      diagnostics,
      options,
      WORKBOOK_PART_URI,
    );
    calculation.ForceFullCalculation = readBoolAttribute(
      calcPr,
      "forceFullCalc",
      diagnostics,
      options,
      false,
      WORKBOOK_PART_URI,
    );
  }
}

function readStringAttribute(element: XmlNode, name: string): string {
  const value = attribute(element, name);
  return value === null ? "" : value.trim();
}

function readChoiceAttribute(
  element: XmlNode,
  attributeName: string,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  partUri: string,
  normalizer: (value: string | null) => string,
): string {
  const value = attribute(element, attributeName);
  if (value === null) {
    return "";
  }

  try {
    return normalizer(value);
  } catch (error) {
    if (error instanceof CellsException) {
      addWorkbookMetadataIssue(
        diagnostics,
        options,
        partUri,
        "Workbook metadata attribute '" + attributeName + "' had an invalid value and was ignored.",
      );
      return "";
    }

    throw error;
  }
}

function readBoolAttribute(
  element: XmlNode,
  attributeName: string,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  defaultValue: boolean,
  partUri: string,
): boolean {
  const value = readNullableBoolAttribute(element, attributeName, diagnostics, options, partUri);
  return value === null ? defaultValue : value;
}

function readNullableBoolAttribute(
  element: XmlNode,
  attributeName: string,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  partUri: string,
): boolean | null {
  const rawValue = attribute(element, attributeName);
  if (rawValue === null) {
    return null;
  }

  if (rawValue === "1" || rawValue.toLowerCase() === "true") {
    return true;
  }

  if (rawValue === "0" || rawValue.toLowerCase() === "false") {
    return false;
  }

  addWorkbookMetadataIssue(
    diagnostics,
    options,
    partUri,
    "Workbook metadata attribute '" + attributeName + "' had an invalid Boolean value and was ignored.",
  );
  return null;
}

function readNonNegativeIntAttribute(
  element: XmlNode,
  attributeName: string,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  partUri: string,
): number | null {
  const value = attribute(element, attributeName);
  if (value === null) {
    return null;
  }

  const parsed = parseIntAttribute(element, attributeName);
  if (parsed !== null && parsed >= 0) {
    return parsed;
  }

  addWorkbookMetadataIssue(
    diagnostics,
    options,
    partUri,
    "Workbook metadata attribute '" + attributeName + "' had an invalid integer value and was ignored.",
  );
  return null;
}

function readNonNegativeDoubleAttribute(
  element: XmlNode,
  attributeName: string,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  partUri: string,
): number | null {
  const value = attribute(element, attributeName);
  if (value === null) {
    return null;
  }

  const trimmed = value.trim();
  const parsed = trimmed.length === 0 ? Number.NaN : Number(trimmed);
  if (Number.isFinite(parsed) && parsed >= 0) {
    return parsed;
  }

  addWorkbookMetadataIssue(
    diagnostics,
    options,
    partUri,
    "Workbook metadata attribute '" + attributeName + "' had an invalid numeric value and was ignored.",
  );
  return null;
}

function addWorkbookMetadataIssue(
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  partUri: string,
  message: string,
): void {
  const issue = new LoadIssue("WB-L003", DiagnosticSeverity.Warning, message);
  issue.PartUri = partUri;
  addIssue(diagnostics, options, issue);
}

/** Formats iterateDelta using the invariant "0.################" format. */
function formatIterateDelta(value: number): string {
  let text = value.toFixed(16);
  if (text.includes(".")) {
    text = text.replace(/0+$/, "").replace(/\.$/, "");
  }

  return text === "-0" ? "0" : text;
}
