import { CellsException } from "./CellsException";
import { ExtendedDocumentPropertiesModel } from "./Core/ExtendedDocumentPropertiesModel";

/** Represents extended document properties. */
export class ExtendedDocumentProperties {
  private readonly model: ExtendedDocumentPropertiesModel;

  public constructor(model: ExtendedDocumentPropertiesModel) {
    this.model = model;
  }

  public get Application(): string {
    return this.model.Application;
  }

  public set Application(value: string) {
    this.model.Application = value ?? "";
  }

  public get AppVersion(): string {
    return this.model.AppVersion;
  }

  public set AppVersion(value: string) {
    this.model.AppVersion = value ?? "";
  }

  public get Company(): string {
    return this.model.Company;
  }

  public set Company(value: string) {
    this.model.Company = value ?? "";
  }

  public get Manager(): string {
    return this.model.Manager;
  }

  public set Manager(value: string) {
    this.model.Manager = value ?? "";
  }

  public get DocSecurity(): number {
    return this.model.DocSecurity ?? 0;
  }

  public set DocSecurity(value: number) {
    if (value < 0) {
      throw new CellsException("DocSecurity must be non-negative.");
    }

    this.model.DocSecurity = value;
  }

  public get HyperlinkBase(): string {
    return this.model.HyperlinkBase;
  }

  public set HyperlinkBase(value: string) {
    this.model.HyperlinkBase = value ?? "";
  }

  public get ScaleCrop(): boolean {
    return this.model.ScaleCrop ?? false;
  }

  public set ScaleCrop(value: boolean) {
    this.model.ScaleCrop = value;
  }

  public get LinksUpToDate(): boolean {
    return this.model.LinksUpToDate ?? false;
  }

  public set LinksUpToDate(value: boolean) {
    this.model.LinksUpToDate = value;
  }

  public get SharedDoc(): boolean {
    return this.model.SharedDoc ?? false;
  }

  public set SharedDoc(value: boolean) {
    this.model.SharedDoc = value;
  }
}
