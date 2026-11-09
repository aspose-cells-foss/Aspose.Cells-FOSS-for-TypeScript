import * as fs from "fs";
import * as path from "path";
import { Cell, CellValueType, Workbook } from "../../aspose_cells";
import {
  resolveRepositoryRoot,
  TestCase,
  TestRunner,
  TemporaryDirectory,
} from "../Shared/TestInfrastructure";
import {
  ConditionalFormattingScenarioFactory,
  DefinedNameScenarioFactory,
  HyperlinkScenarioFactory,
  PageSetupScenarioFactory,
  StyleScenarioFactory,
  ValidationScenarioFactory,
  WorkbookScenarioFactory,
  WorksheetScenarioFactory,
} from "../Shared/ScenarioFactories";
import { AutoFilterScenarioFactory } from "../Shared/AutoFilterScenarioFactory";
import { WorkbookMetadataScenarioFactory } from "../Shared/WorkbookMetadataScenarioFactory";
import {
  CaseInsensitiveStringMap,
  CellSnapshot,
  SnapshotMismatch,
  StyleSnapshot,
} from "./ComparisonModels";
import {
  buildCellKey,
  normalizeFormulaText,
  normalizeValue,
  tryNormalizeCellReference,
} from "./ComparisonValueHelpers";
import {
  createLibraryStyleSnapshot,
  readCellDataWithOpenXmlSdk,
  readCellStylesWithOpenXmlSdk,
} from "./OpenXmlComparisonSupport";

const SupportedExtensions = new Set<string>([".xlsx", ".xlsm", ".xltx", ".xltm"]);

const CompareInputRoot = path.join(resolveRepositoryRoot(), "Input", "compare");

function main(): number {
  const dataTemp = new TemporaryDirectory("compare-openxml-data");
  const styleTemp = new TemporaryDirectory("compare-openxml-styles");

  const dataLogPath = dataTemp.getPath("mismatches.log");
  const styleLogPath = styleTemp.getPath("mismatches.log");

  const dataCases: TestCase[] = [
    new TestCase("mixed_cell_workbook_data_matches_openxml_sdk", () => {
      const filePath = dataTemp.getPath("mixed.xlsx");
      WorkbookScenarioFactory.CreateMixedCellWorkbook().Save(filePath);
      assertNoDataMismatches(filePath, dataLogPath);
    }),
    new TestCase("worksheet_settings_workbook_data_matches_openxml_sdk", () => {
      const filePath = dataTemp.getPath("worksheet-settings.xlsx");
      WorksheetScenarioFactory.CreateWorksheetSettingsWorkbook().Save(filePath);
      assertNoDataMismatches(filePath, dataLogPath);
    }),
    new TestCase("page_setup_workbook_data_matches_openxml_sdk", () => {
      const filePath = dataTemp.getPath("page-setup.xlsx");
      PageSetupScenarioFactory.CreatePageSetupWorkbook().Save(filePath);
      assertNoDataMismatches(filePath, dataLogPath);
    }),
    new TestCase("hyperlinks_workbook_data_matches_openxml_sdk", () => {
      const filePath = dataTemp.getPath("hyperlinks.xlsx");
      HyperlinkScenarioFactory.CreateHyperlinkWorkbook().Save(filePath);
      assertNoDataMismatches(filePath, dataLogPath);
    }),
    new TestCase("styled_workbook_data_matches_openxml_sdk", () => {
      const filePath = dataTemp.getPath("styled.xlsx");
      StyleScenarioFactory.CreateStyledWorkbook().Save(filePath);
      assertNoDataMismatches(filePath, dataLogPath);
    }),
    new TestCase("validation_workbook_data_matches_openxml_sdk", () => {
      const filePath = dataTemp.getPath("validation.xlsx");
      ValidationScenarioFactory.CreateValidationWorkbook().Save(filePath);
      assertNoDataMismatches(filePath, dataLogPath);
    }),
    new TestCase("conditional_formatting_workbook_data_matches_openxml_sdk", () => {
      const filePath = dataTemp.getPath("conditional-formatting.xlsx");
      ConditionalFormattingScenarioFactory.CreateConditionalFormattingWorkbook().Save(filePath);
      assertNoDataMismatches(filePath, dataLogPath);
    }),
    new TestCase("defined_names_workbook_data_matches_openxml_sdk", () => {
      const filePath = dataTemp.getPath("defined-names.xlsx");
      DefinedNameScenarioFactory.CreateDefinedNamesWorkbook().Save(filePath);
      assertNoDataMismatches(filePath, dataLogPath);
    }),
    new TestCase("autofilter_workbook_data_matches_openxml_sdk", () => {
      const filePath = dataTemp.getPath("autofilter.xlsx");
      AutoFilterScenarioFactory.CreateAutoFilterWorkbook().Save(filePath);
      assertNoDataMismatches(filePath, dataLogPath);
    }),
    new TestCase("workbook_metadata_workbook_data_matches_openxml_sdk", () => {
      const filePath = dataTemp.getPath("workbook-metadata.xlsx");
      WorkbookMetadataScenarioFactory.CreateWorkbookMetadataWorkbook().Save(filePath);
      assertNoDataMismatches(filePath, dataLogPath);
    }),
  ];

  const styleCases: TestCase[] = [
    new TestCase("mixed_cell_workbook_styles_match_openxml_sdk", () => {
      const filePath = styleTemp.getPath("mixed.xlsx");
      WorkbookScenarioFactory.CreateMixedCellWorkbook().Save(filePath);
      assertNoStyleMismatches(filePath, styleLogPath);
    }),
    new TestCase("worksheet_settings_workbook_styles_match_openxml_sdk", () => {
      const filePath = styleTemp.getPath("worksheet-settings.xlsx");
      WorksheetScenarioFactory.CreateWorksheetSettingsWorkbook().Save(filePath);
      assertNoStyleMismatches(filePath, styleLogPath);
    }),
    new TestCase("page_setup_workbook_styles_match_openxml_sdk", () => {
      const filePath = styleTemp.getPath("page-setup.xlsx");
      PageSetupScenarioFactory.CreatePageSetupWorkbook().Save(filePath);
      assertNoStyleMismatches(filePath, styleLogPath);
    }),
    new TestCase("hyperlinks_workbook_styles_match_openxml_sdk", () => {
      const filePath = styleTemp.getPath("hyperlinks.xlsx");
      HyperlinkScenarioFactory.CreateHyperlinkWorkbook().Save(filePath);
      assertNoStyleMismatches(filePath, styleLogPath);
    }),
    new TestCase("styled_workbook_styles_match_openxml_sdk", () => {
      const filePath = styleTemp.getPath("styled.xlsx");
      StyleScenarioFactory.CreateStyledWorkbook().Save(filePath);
      assertNoStyleMismatches(filePath, styleLogPath);
    }),
    new TestCase("validation_workbook_styles_match_openxml_sdk", () => {
      const filePath = styleTemp.getPath("validation.xlsx");
      ValidationScenarioFactory.CreateValidationWorkbook().Save(filePath);
      assertNoStyleMismatches(filePath, styleLogPath);
    }),
    new TestCase("conditional_formatting_workbook_styles_match_openxml_sdk", () => {
      const filePath = styleTemp.getPath("conditional-formatting.xlsx");
      ConditionalFormattingScenarioFactory.CreateConditionalFormattingWorkbook().Save(filePath);
      assertNoStyleMismatches(filePath, styleLogPath);
    }),
    new TestCase("defined_names_workbook_styles_match_openxml_sdk", () => {
      const filePath = styleTemp.getPath("defined-names.xlsx");
      DefinedNameScenarioFactory.CreateDefinedNamesWorkbook().Save(filePath);
      assertNoStyleMismatches(filePath, styleLogPath);
    }),
    new TestCase("autofilter_workbook_styles_match_openxml_sdk", () => {
      const filePath = styleTemp.getPath("autofilter.xlsx");
      AutoFilterScenarioFactory.CreateAutoFilterWorkbook().Save(filePath);
      assertNoStyleMismatches(filePath, styleLogPath);
    }),
    new TestCase("workbook_metadata_workbook_styles_match_openxml_sdk", () => {
      const filePath = styleTemp.getPath("workbook-metadata.xlsx");
      WorkbookMetadataScenarioFactory.CreateWorkbookMetadataWorkbook().Save(filePath);
      assertNoStyleMismatches(filePath, styleLogPath);
    }),
  ];

  for (const filePath of discoverCompareFiles(CompareInputRoot)) {
    const capturedPath = filePath;
    const caseName = path.basename(filePath, path.extname(filePath)).toLowerCase().replace(/ /g, "_");
    dataCases.push(
      new TestCase(`${caseName}_data_matches_openxml_sdk`, () => {
        assertNoDataMismatches(capturedPath, dataLogPath);
      }),
    );
    styleCases.push(
      new TestCase(`${caseName}_styles_match_openxml_sdk`, () => {
        assertNoStyleMismatches(capturedPath, styleLogPath);
      }),
    );
  }

  const dataResult = TestRunner.run("CompareOpenXml.DataTests", dataCases);
  const styleResult = TestRunner.run("CompareOpenXml.StyleTests", styleCases);

  if (fs.existsSync(dataLogPath)) {
    console.log(`Data mismatch log: ${dataLogPath}`);
  }

  if (fs.existsSync(styleLogPath)) {
    console.log(`Style mismatch log: ${styleLogPath}`);
  }

  return dataResult | styleResult;
}

function discoverCompareFiles(root: string): string[] {
  if (!fs.existsSync(root)) {
    return [];
  }

  const results: string[] = [];
  const visit = (directory: string): void => {
    for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
      const fullPath = path.join(directory, entry.name);
      if (entry.isDirectory()) {
        visit(fullPath);
        continue;
      }

      if (!entry.isFile()) {
        continue;
      }

      if (!SupportedExtensions.has(path.extname(fullPath).toLowerCase())) {
        continue;
      }

      if (path.basename(fullPath).startsWith("~$")) {
        continue;
      }

      results.push(fullPath);
    }
  };

  visit(root);
  results.sort((left, right) => {
    const a = left.toLowerCase();
    const b = right.toLowerCase();
    return a < b ? -1 : a > b ? 1 : 0;
  });
  return results;
}

function assertNoDataMismatches(workbookPath: string, logPath: string): void {
  const mismatches = compareData(workbookPath);
  if (mismatches.length === 0) {
    return;
  }

  appendMismatchLog(logPath, workbookPath, mismatches, "cell mismatch(es)");

  const lines: string[] = [`${mismatches.length} data mismatch(es) in '${path.basename(workbookPath)}':`];
  for (const mismatch of mismatches) {
    lines.push(`  Cell: ${mismatch.CellKey}`);
    lines.push(`    Library: ${mismatch.LibrarySnapshot}`);
    lines.push(`    OpenXml: ${mismatch.OpenXmlSnapshot}`);
  }

  throw new Error(lines.join("\n"));
}

function assertNoStyleMismatches(workbookPath: string, logPath: string): void {
  const mismatches = compareStyles(workbookPath);
  if (mismatches.length === 0) {
    return;
  }

  appendMismatchLog(logPath, workbookPath, mismatches, "style mismatch(es)");

  const lines: string[] = [`${mismatches.length} style mismatch(es) in '${path.basename(workbookPath)}':`];
  for (const mismatch of mismatches) {
    lines.push(`  Cell: ${mismatch.CellKey}`);
    lines.push(`    Library: ${mismatch.LibrarySnapshot}`);
    lines.push(`    OpenXml: ${mismatch.OpenXmlSnapshot}`);
  }

  throw new Error(lines.join("\n"));
}

function appendMismatchLog(
  logPath: string,
  workbookPath: string,
  mismatches: readonly SnapshotMismatch[],
  mismatchLabel: string,
): void {
  let text =
    `Workbook: ${path.basename(workbookPath)}\n` +
    `Path: ${workbookPath}\n` +
    `Result: mismatched (${mismatches.length} ${mismatchLabel})\n`;

  for (const mismatch of mismatches) {
    text += `Cell: ${mismatch.CellKey}\n`;
    text += `  Aspose.Cells_FOSS: ${mismatch.LibrarySnapshot}\n`;
    text += `  Open XML SDK:     ${mismatch.OpenXmlSnapshot}\n`;
  }

  text += "\n";
  fs.appendFileSync(logPath, text);
}

function compareData(workbookPath: string): SnapshotMismatch[] {
  const libraryCells = readCellDataWithLibrary(workbookPath);
  const openXmlCells = readCellDataWithOpenXmlSdk(workbookPath);
  return compareCellMaps(libraryCells, openXmlCells);
}

function compareStyles(workbookPath: string): SnapshotMismatch[] {
  const libraryStyles = readCellStylesWithLibrary(workbookPath);
  const openXmlStyles = readCellStylesWithOpenXmlSdk(workbookPath);
  return compareStyleMaps(libraryStyles, openXmlStyles);
}

function readCellDataWithLibrary(workbookPath: string): CaseInsensitiveStringMap<CellSnapshot> {
  const workbook = new Workbook(workbookPath);
  return enumerateLibraryCells(
    workbook,
    (sheetName, cellReference, cell) =>
      new CellSnapshot(
        sheetName,
        cellReference,
        normalizeCellType(cell.Type),
        normalizeValue(cell.Value),
        normalizeFormulaText(cell.Formula),
      ),
  );
}

function readCellStylesWithLibrary(workbookPath: string): CaseInsensitiveStringMap<StyleSnapshot> {
  const workbook = new Workbook(workbookPath);
  return enumerateLibraryCells(workbook, (sheetName, cellReference, cell) =>
    createLibraryStyleSnapshot(sheetName, cellReference, cell.GetStyle()),
  );
}

function enumerateLibraryCells<T>(
  workbook: Workbook,
  snapshotFactory: (sheetName: string, cellReference: string, cell: Cell) => T,
): CaseInsensitiveStringMap<T> {
  const map = new CaseInsensitiveStringMap<T>();

  for (const worksheetModel of workbook.Model.Worksheets) {
    const sheetName = worksheetModel.Name ?? "<unknown>";
    const sheet = workbook.Worksheets.get(sheetName);

    for (const address of worksheetModel.Cells.keys()) {
      const addressText = address.ToString();
      const cellReference = tryNormalizeCellReference(addressText);
      if (cellReference === null) {
        continue;
      }

      const cell = sheet.Cells.get(cellReference);
      map.set(buildCellKey(sheetName, cellReference), snapshotFactory(sheetName, cellReference, cell));
    }
  }

  return map;
}

function compareCellMaps(
  libraryCells: CaseInsensitiveStringMap<CellSnapshot>,
  openXmlCells: CaseInsensitiveStringMap<CellSnapshot>,
): SnapshotMismatch[] {
  const allKeys = unionKeys(libraryCells.keys(), openXmlCells.keys());

  const mismatches: SnapshotMismatch[] = [];
  for (const key of allKeys) {
    const librarySnapshot = libraryCells.get(key);
    const openXmlSnapshot = openXmlCells.get(key);

    if (librarySnapshot === undefined || openXmlSnapshot === undefined) {
      if (isEffectivelyBlank(librarySnapshot) && isEffectivelyBlank(openXmlSnapshot)) {
        continue;
      }

      mismatches.push(
        new SnapshotMismatch(
          key,
          librarySnapshot?.ToString() ?? "<missing>",
          openXmlSnapshot?.ToString() ?? "<missing>",
        ),
      );
      continue;
    }

    if (
      librarySnapshot.CellType !== openXmlSnapshot.CellType ||
      librarySnapshot.Value !== openXmlSnapshot.Value ||
      librarySnapshot.Formula !== openXmlSnapshot.Formula
    ) {
      mismatches.push(new SnapshotMismatch(key, librarySnapshot.ToString(), openXmlSnapshot.ToString()));
    }
  }

  return mismatches;
}

function compareStyleMaps(
  libraryStyles: CaseInsensitiveStringMap<StyleSnapshot>,
  openXmlStyles: CaseInsensitiveStringMap<StyleSnapshot>,
): SnapshotMismatch[] {
  const allKeys = unionKeys(libraryStyles.keys(), openXmlStyles.keys());

  const mismatches: SnapshotMismatch[] = [];
  for (const key of allKeys) {
    const librarySnapshot = libraryStyles.get(key);
    const openXmlSnapshot = openXmlStyles.get(key);

    if (librarySnapshot === undefined || openXmlSnapshot === undefined) {
      if (isEffectivelyDefaultStyle(librarySnapshot) && isEffectivelyDefaultStyle(openXmlSnapshot)) {
        continue;
      }

      mismatches.push(
        new SnapshotMismatch(
          key,
          librarySnapshot?.ToString() ?? "<missing>",
          openXmlSnapshot?.ToString() ?? "<missing>",
        ),
      );
      continue;
    }

    if (!librarySnapshot.Equals(openXmlSnapshot)) {
      mismatches.push(new SnapshotMismatch(key, librarySnapshot.ToString(), openXmlSnapshot.ToString()));
    }
  }

  return mismatches;
}

function unionKeys(leftKeys: readonly string[], rightKeys: readonly string[]): string[] {
  const seen = new Set<string>();
  const keys: string[] = [];
  for (const key of leftKeys) {
    const normalized = key.toLowerCase();
    if (!seen.has(normalized)) {
      seen.add(normalized);
      keys.push(key);
    }
  }

  for (const key of rightKeys) {
    const normalized = key.toLowerCase();
    if (!seen.has(normalized)) {
      seen.add(normalized);
      keys.push(key);
    }
  }

  keys.sort((left, right) => {
    const a = left.toLowerCase();
    const b = right.toLowerCase();
    return a < b ? -1 : a > b ? 1 : 0;
  });
  return keys;
}

function isEffectivelyBlank(snapshot: CellSnapshot | undefined): boolean {
  return (
    snapshot === undefined ||
    (snapshot.CellType === "Blank" &&
      snapshot.Formula.length === 0 &&
      snapshot.Value === "<null>")
  );
}

function isEffectivelyDefaultStyle(snapshot: StyleSnapshot | undefined): boolean {
  return snapshot === undefined || snapshot.Equals(StyleSnapshot.Default);
}

function normalizeCellType(type: CellValueType): string {
  switch (type) {
    case CellValueType.IsString:
      return "String";
    case CellValueType.IsNumeric:
      return "Number";
    case CellValueType.IsBool:
      return "Boolean";
    case CellValueType.IsDateTime:
      return "DateTime";
    case CellValueType.IsError:
      return "String";
    case CellValueType.IsUnknown:
      return "Formula";
    default:
      return "Blank";
  }
}

process.exitCode = main();
