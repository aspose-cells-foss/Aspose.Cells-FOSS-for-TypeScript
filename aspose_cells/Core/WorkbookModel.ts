import { DiagnosticBag } from "./DiagnosticBag";
import { DocumentPropertiesModel } from "./DocumentPropertiesModel";
import { DefinedNameModel } from "./DefinedNameModel";
import { ExternalLinkModel } from "./ExternalLinkModel";
import { SharedStringRepository } from "./SharedStringRepository";
import { StyleRepository } from "./StyleRepository";
import { StyleValue } from "./StyleValue";
import { WorkbookPropertiesModel } from "./WorkbookPropertiesModel";
import { WorkbookSettingsModel } from "./WorkbookSettingsModel";
import { WorksheetModel } from "./WorksheetModel";

export class WorkbookModel {
  public readonly Worksheets: WorksheetModel[];
  public readonly Settings: WorkbookSettingsModel;
  public readonly Properties: WorkbookPropertiesModel;
  public readonly DocumentProperties: DocumentPropertiesModel;
  public readonly Diagnostics: DiagnosticBag;
  public readonly Styles: StyleRepository;
  public readonly SharedStrings: SharedStringRepository;
  public DefaultStyle: StyleValue;
  public ActiveSheetIndex: number;
  public readonly DefinedNames: DefinedNameModel[];
  public RawThemeXml: string | null = null;
  public RawTableStylesXml: string | null = null;
  public readonly TableDifferentialStyles: StyleValue[];
  public readonly RawTableDifferentialStylesXml: string[];
  public readonly ExternalLinks: ExternalLinkModel[];

  public constructor() {
    this.Settings = new WorkbookSettingsModel();
    this.Properties = new WorkbookPropertiesModel();
    this.DocumentProperties = new DocumentPropertiesModel();
    this.Diagnostics = new DiagnosticBag();
    this.Styles = new StyleRepository();
    this.SharedStrings = new SharedStringRepository();
    this.DefaultStyle = StyleValue.Default.clone();
    this.Worksheets = [new WorksheetModel("Sheet1")];
    this.DefinedNames = [];
    this.ActiveSheetIndex = 0;
    this.TableDifferentialStyles = [];
    this.RawTableDifferentialStylesXml = [];
    this.ExternalLinks = [];
  }
}
