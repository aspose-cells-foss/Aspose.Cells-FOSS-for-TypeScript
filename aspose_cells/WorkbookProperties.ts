import { CalculationProperties } from "./CalculationProperties";
import { CellsException } from "./CellsException";
import { WorkbookProtection } from "./WorkbookProtection";
import { WorkbookView } from "./WorkbookView";
import { normalizeShowObjects, normalizeUpdateLinks } from "./WorkbookPropertySupport";
import { WorkbookModel } from "./Core/WorkbookModel";

/** Represents workbook metadata and view settings. */
export class WorkbookProperties {
  private readonly workbookModel: WorkbookModel;
  public readonly Protection: WorkbookProtection;
  public readonly View: WorkbookView;
  public readonly Calculation: CalculationProperties;

  public constructor(workbookModel: WorkbookModel) {
    this.workbookModel = workbookModel;
    this.Protection = new WorkbookProtection(workbookModel.Properties.Protection);
    this.View = new WorkbookView(workbookModel);
    this.Calculation = new CalculationProperties(workbookModel.Properties.Calculation);
  }

  private get model(): WorkbookModel["Properties"] {
    return this.workbookModel.Properties;
  }

  public get CodeName(): string {
    return this.model.CodeName;
  }

  public set CodeName(value: string) {
    this.model.CodeName = value ?? "";
  }

  public get ShowObjects(): string {
    return this.model.ShowObjects === null || this.model.ShowObjects.length === 0
      ? "all"
      : this.model.ShowObjects;
  }

  public set ShowObjects(value: string) {
    this.model.ShowObjects = normalizeShowObjects(value);
  }

  public get FilterPrivacy(): boolean {
    return this.model.FilterPrivacy;
  }

  public set FilterPrivacy(value: boolean) {
    this.model.FilterPrivacy = value;
  }

  public get ShowBorderUnselectedTables(): boolean {
    return this.model.ShowBorderUnselectedTables;
  }

  public set ShowBorderUnselectedTables(value: boolean) {
    this.model.ShowBorderUnselectedTables = value;
  }

  public get ShowInkAnnotation(): boolean {
    return this.model.ShowInkAnnotation;
  }

  public set ShowInkAnnotation(value: boolean) {
    this.model.ShowInkAnnotation = value;
  }

  public get BackupFile(): boolean {
    return this.model.BackupFile;
  }

  public set BackupFile(value: boolean) {
    this.model.BackupFile = value;
  }

  public get SaveExternalLinkValues(): boolean {
    return this.model.SaveExternalLinkValues;
  }

  public set SaveExternalLinkValues(value: boolean) {
    this.model.SaveExternalLinkValues = value;
  }

  public get UpdateLinks(): string {
    return this.model.UpdateLinks === null || this.model.UpdateLinks.length === 0
      ? "userSet"
      : this.model.UpdateLinks;
  }

  public set UpdateLinks(value: string) {
    this.model.UpdateLinks = normalizeUpdateLinks(value);
  }

  public get HidePivotFieldList(): boolean {
    return this.model.HidePivotFieldList;
  }

  public set HidePivotFieldList(value: boolean) {
    this.model.HidePivotFieldList = value;
  }

  public get DefaultThemeVersion(): number | null {
    return this.model.DefaultThemeVersion;
  }

  public set DefaultThemeVersion(value: number | null) {
    if (value !== null && value !== undefined && value < 0) {
      throw new CellsException("DefaultThemeVersion must be non-negative.");
    }

    this.model.DefaultThemeVersion = value;
  }
}
