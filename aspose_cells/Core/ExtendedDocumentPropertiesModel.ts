export class ExtendedDocumentPropertiesModel {
  public Application: string = "";
  public AppVersion: string = "";
  public Company: string = "";
  public Manager: string = "";
  public DocSecurity: number | null = null;
  public HyperlinkBase: string = "";
  public ScaleCrop: boolean | null = null;
  public LinksUpToDate: boolean | null = null;
  public SharedDoc: boolean | null = null;

  public CopyFrom(source: ExtendedDocumentPropertiesModel): void {
    this.Application = source.Application;
    this.AppVersion = source.AppVersion;
    this.Company = source.Company;
    this.Manager = source.Manager;
    this.DocSecurity = source.DocSecurity;
    this.HyperlinkBase = source.HyperlinkBase;
    this.ScaleCrop = source.ScaleCrop;
    this.LinksUpToDate = source.LinksUpToDate;
    this.SharedDoc = source.SharedDoc;
  }

  public HasStoredState(): boolean {
    return (
      (this.Application != null && this.Application.length > 0) ||
      (this.AppVersion != null && this.AppVersion.length > 0) ||
      (this.Company != null && this.Company.length > 0) ||
      (this.Manager != null && this.Manager.length > 0) ||
      this.DocSecurity !== null ||
      (this.HyperlinkBase != null && this.HyperlinkBase.length > 0) ||
      this.ScaleCrop !== null ||
      this.LinksUpToDate !== null ||
      this.SharedDoc !== null
    );
  }
}
