import { VisibilityType, Workbook } from "../../aspose_cells";
import { AssertEx } from "./TestInfrastructure";

/** Scenario factory for workbook metadata, document properties, and view settings. */
export class WorkbookMetadataScenarioFactory {
  public static CreateWorkbookMetadataWorkbook(): Workbook {
    const workbook = new Workbook();
    workbook.Settings.Date1904 = true;

    workbook.Worksheets[0].Name = "Summary";
    const dataIndex = workbook.Worksheets.Add("Data");
    const archiveIndex = workbook.Worksheets.Add("Archive");
    workbook.Worksheets[archiveIndex].VisibilityType = VisibilityType.Hidden;
    workbook.Worksheets[dataIndex].Cells.get(0, 0).PutValue("Ready");

    workbook.Properties.CodeName = "WorkbookCode";
    workbook.Properties.ShowObjects = "placeholders";
    workbook.Properties.FilterPrivacy = true;
    workbook.Properties.ShowBorderUnselectedTables = false;
    workbook.Properties.ShowInkAnnotation = false;
    workbook.Properties.BackupFile = true;
    workbook.Properties.SaveExternalLinkValues = false;
    workbook.Properties.UpdateLinks = "never";
    workbook.Properties.HidePivotFieldList = true;
    workbook.Properties.DefaultThemeVersion = 166925;

    workbook.Properties.Protection.LockStructure = true;
    workbook.Properties.Protection.LockWindows = true;
    workbook.Properties.Protection.WorkbookPassword = "ABCD";
    workbook.Properties.Protection.RevisionsPassword = "CDEF";

    workbook.Properties.View.XWindow = 120;
    workbook.Properties.View.YWindow = 240;
    workbook.Properties.View.WindowWidth = 16000;
    workbook.Properties.View.WindowHeight = 9000;
    workbook.Properties.View.ActiveTab = 1;
    workbook.Properties.View.FirstSheet = 1;
    workbook.Properties.View.ShowHorizontalScroll = false;
    workbook.Properties.View.ShowVerticalScroll = false;
    workbook.Properties.View.ShowSheetTabs = false;
    workbook.Properties.View.TabRatio = 700;
    workbook.Properties.View.Visibility = "visible";
    workbook.Properties.View.Minimized = false;
    workbook.Properties.View.AutoFilterDateGrouping = false;

    workbook.Properties.Calculation.CalculationId = 191029;
    workbook.Properties.Calculation.CalculationMode = "manual";
    workbook.Properties.Calculation.FullCalculationOnLoad = true;
    workbook.Properties.Calculation.ReferenceMode = "R1C1";
    workbook.Properties.Calculation.Iterate = true;
    workbook.Properties.Calculation.IterateCount = 9;
    workbook.Properties.Calculation.IterateDelta = 0.00125;
    workbook.Properties.Calculation.FullPrecision = false;
    workbook.Properties.Calculation.CalculationCompleted = false;
    workbook.Properties.Calculation.CalculationOnSave = false;
    workbook.Properties.Calculation.ConcurrentCalculation = false;
    workbook.Properties.Calculation.ForceFullCalculation = true;

    workbook.DocumentProperties.Title = "Quarterly Summary";
    workbook.DocumentProperties.Subject = "Operations";
    workbook.DocumentProperties.Author = "Automation";
    workbook.DocumentProperties.Keywords = "finance;ops";
    workbook.DocumentProperties.Comments = "Generated during tests.";
    workbook.DocumentProperties.Category = "Reports";
    workbook.DocumentProperties.Company = "Aspose Cells FOSS";
    workbook.DocumentProperties.Manager = "Release";
    workbook.DocumentProperties.Core.LastModifiedBy = "Verifier";
    workbook.DocumentProperties.Core.Revision = "7";
    workbook.DocumentProperties.Core.ContentStatus = "Draft";
    workbook.DocumentProperties.Core.Created = new Date(Date.UTC(2024, 0, 2, 3, 4, 5));
    workbook.DocumentProperties.Core.Modified = new Date(Date.UTC(2024, 5, 7, 8, 9, 10));
    workbook.DocumentProperties.Extended.Application = "Aspose.Cells_FOSS Tests";
    workbook.DocumentProperties.Extended.AppVersion = "0.2";
    workbook.DocumentProperties.Extended.DocSecurity = 2;
    workbook.DocumentProperties.Extended.HyperlinkBase = "https://example.com/base/";
    workbook.DocumentProperties.Extended.ScaleCrop = true;
    workbook.DocumentProperties.Extended.LinksUpToDate = true;
    workbook.DocumentProperties.Extended.SharedDoc = true;

    return workbook;
  }

  public static AssertWorkbookMetadata(workbook: Workbook): void {
    AssertEx.true(workbook.Settings.Date1904);
    AssertEx.equal("WorkbookCode", workbook.Properties.CodeName);
    AssertEx.equal("placeholders", workbook.Properties.ShowObjects);
    AssertEx.true(workbook.Properties.FilterPrivacy);
    AssertEx.false(workbook.Properties.ShowBorderUnselectedTables);
    AssertEx.false(workbook.Properties.ShowInkAnnotation);
    AssertEx.true(workbook.Properties.BackupFile);
    AssertEx.false(workbook.Properties.SaveExternalLinkValues);
    AssertEx.equal("never", workbook.Properties.UpdateLinks);
    AssertEx.true(workbook.Properties.HidePivotFieldList);
    AssertEx.equal(166925, workbook.Properties.DefaultThemeVersion ?? 0);

    AssertEx.true(workbook.Properties.Protection.IsProtected);
    AssertEx.true(workbook.Properties.Protection.LockStructure);
    AssertEx.true(workbook.Properties.Protection.LockWindows);
    AssertEx.equal("ABCD", workbook.Properties.Protection.WorkbookPassword);
    AssertEx.equal("CDEF", workbook.Properties.Protection.RevisionsPassword);

    AssertEx.equal(120, workbook.Properties.View.XWindow);
    AssertEx.equal(240, workbook.Properties.View.YWindow);
    AssertEx.equal(16000, workbook.Properties.View.WindowWidth);
    AssertEx.equal(9000, workbook.Properties.View.WindowHeight);
    AssertEx.equal(1, workbook.Properties.View.ActiveTab);
    AssertEx.equal(1, workbook.Properties.View.FirstSheet);
    AssertEx.false(workbook.Properties.View.ShowHorizontalScroll);
    AssertEx.false(workbook.Properties.View.ShowVerticalScroll);
    AssertEx.false(workbook.Properties.View.ShowSheetTabs);
    AssertEx.equal(700, workbook.Properties.View.TabRatio);
    AssertEx.equal("visible", workbook.Properties.View.Visibility);
    AssertEx.false(workbook.Properties.View.Minimized);
    AssertEx.false(workbook.Properties.View.AutoFilterDateGrouping);

    AssertEx.equal(191029, workbook.Properties.Calculation.CalculationId ?? 0);
    AssertEx.equal("manual", workbook.Properties.Calculation.CalculationMode);
    AssertEx.true(workbook.Properties.Calculation.FullCalculationOnLoad);
    AssertEx.equal("R1C1", workbook.Properties.Calculation.ReferenceMode);
    AssertEx.true(workbook.Properties.Calculation.Iterate);
    AssertEx.equal(9, workbook.Properties.Calculation.IterateCount);
    AssertEx.equal(0.00125, workbook.Properties.Calculation.IterateDelta);
    AssertEx.false(workbook.Properties.Calculation.FullPrecision);
    AssertEx.false(workbook.Properties.Calculation.CalculationCompleted);
    AssertEx.false(workbook.Properties.Calculation.CalculationOnSave);
    AssertEx.false(workbook.Properties.Calculation.ConcurrentCalculation);
    AssertEx.true(workbook.Properties.Calculation.ForceFullCalculation);

    AssertEx.equal("Quarterly Summary", workbook.DocumentProperties.Title);
    AssertEx.equal("Operations", workbook.DocumentProperties.Subject);
    AssertEx.equal("Automation", workbook.DocumentProperties.Author);
    AssertEx.equal("finance;ops", workbook.DocumentProperties.Keywords);
    AssertEx.equal("Generated during tests.", workbook.DocumentProperties.Comments);
    AssertEx.equal("Reports", workbook.DocumentProperties.Category);
    AssertEx.equal("Aspose Cells FOSS", workbook.DocumentProperties.Company);
    AssertEx.equal("Release", workbook.DocumentProperties.Manager);
    AssertEx.equal("Verifier", workbook.DocumentProperties.Core.LastModifiedBy);
    AssertEx.equal("7", workbook.DocumentProperties.Core.Revision);
    AssertEx.equal("Draft", workbook.DocumentProperties.Core.ContentStatus);
    AssertEx.equal(
      new Date(Date.UTC(2024, 0, 2, 3, 4, 5)),
      workbook.DocumentProperties.Core.Created ?? new Date(0),
    );
    AssertEx.equal(
      new Date(Date.UTC(2024, 5, 7, 8, 9, 10)),
      workbook.DocumentProperties.Core.Modified ?? new Date(0),
    );
    AssertEx.equal("Aspose.Cells_FOSS Tests", workbook.DocumentProperties.Extended.Application);
    AssertEx.equal("0.2", workbook.DocumentProperties.Extended.AppVersion);
    AssertEx.equal(2, workbook.DocumentProperties.Extended.DocSecurity);
    AssertEx.equal("https://example.com/base/", workbook.DocumentProperties.Extended.HyperlinkBase);
    AssertEx.true(workbook.DocumentProperties.Extended.ScaleCrop);
    AssertEx.true(workbook.DocumentProperties.Extended.LinksUpToDate);
    AssertEx.true(workbook.DocumentProperties.Extended.SharedDoc);
  }
}
