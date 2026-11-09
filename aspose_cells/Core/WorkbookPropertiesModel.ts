import { CalculationPropertiesModel } from "./CalculationPropertiesModel";
import { WorkbookProtectionModel } from "./WorkbookProtectionModel";
import { WorkbookViewModel } from "./WorkbookViewModel";

export class WorkbookPropertiesModel {
  public CodeName: string = "";
  public ShowObjects: string = "";
  public FilterPrivacy: boolean = false;
  public ShowBorderUnselectedTables: boolean = true;
  public ShowInkAnnotation: boolean = true;
  public BackupFile: boolean = false;
  public SaveExternalLinkValues: boolean = true;
  public UpdateLinks: string = "";
  public HidePivotFieldList: boolean = false;
  public DefaultThemeVersion: number | null = null;
  public readonly Protection: WorkbookProtectionModel;
  public readonly View: WorkbookViewModel;
  public readonly Calculation: CalculationPropertiesModel;

  public constructor() {
    this.Protection = new WorkbookProtectionModel();
    this.View = new WorkbookViewModel();
    this.Calculation = new CalculationPropertiesModel();
  }

  public CopyFrom(source: WorkbookPropertiesModel): void {
    this.CodeName = source.CodeName;
    this.ShowObjects = source.ShowObjects;
    this.FilterPrivacy = source.FilterPrivacy;
    this.ShowBorderUnselectedTables = source.ShowBorderUnselectedTables;
    this.ShowInkAnnotation = source.ShowInkAnnotation;
    this.BackupFile = source.BackupFile;
    this.SaveExternalLinkValues = source.SaveExternalLinkValues;
    this.UpdateLinks = source.UpdateLinks;
    this.HidePivotFieldList = source.HidePivotFieldList;
    this.DefaultThemeVersion = source.DefaultThemeVersion;
    this.Protection.CopyFrom(source.Protection);
    this.View.CopyFrom(source.View);
    this.Calculation.CopyFrom(source.Calculation);
  }

  public HasWorkbookPropertiesState(): boolean {
    return (
      (this.CodeName != null && this.CodeName.length > 0) ||
      (this.ShowObjects != null && this.ShowObjects.length > 0) ||
      this.FilterPrivacy ||
      !this.ShowBorderUnselectedTables ||
      !this.ShowInkAnnotation ||
      this.BackupFile ||
      !this.SaveExternalLinkValues ||
      (this.UpdateLinks != null && this.UpdateLinks.length > 0) ||
      this.HidePivotFieldList ||
      this.DefaultThemeVersion !== null
    );
  }
}
