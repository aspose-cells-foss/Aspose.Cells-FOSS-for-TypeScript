import { CoreDocumentProperties } from "./CoreDocumentProperties";
import { ExtendedDocumentProperties } from "./ExtendedDocumentProperties";
import { DocumentPropertiesModel } from "./Core/DocumentPropertiesModel";

/** Represents document properties. */
export class DocumentProperties {
  private readonly model: DocumentPropertiesModel;
  public readonly Core: CoreDocumentProperties;
  public readonly Extended: ExtendedDocumentProperties;

  public constructor(model: DocumentPropertiesModel) {
    this.model = model;
    this.Core = new CoreDocumentProperties(model.Core);
    this.Extended = new ExtendedDocumentProperties(model.Extended);
  }

  public get Title(): string {
    return this.model.Core.Title;
  }

  public set Title(value: string) {
    this.model.Core.Title = value ?? "";
  }

  public get Subject(): string {
    return this.model.Core.Subject;
  }

  public set Subject(value: string) {
    this.model.Core.Subject = value ?? "";
  }

  public get Author(): string {
    return this.model.Core.Creator;
  }

  public set Author(value: string) {
    this.model.Core.Creator = value ?? "";
  }

  public get Keywords(): string {
    return this.model.Core.Keywords;
  }

  public set Keywords(value: string) {
    this.model.Core.Keywords = value ?? "";
  }

  public get Comments(): string {
    return this.model.Core.Description;
  }

  public set Comments(value: string) {
    this.model.Core.Description = value ?? "";
  }

  public get Category(): string {
    return this.model.Core.Category;
  }

  public set Category(value: string) {
    this.model.Core.Category = value ?? "";
  }

  public get Company(): string {
    return this.model.Extended.Company;
  }

  public set Company(value: string) {
    this.model.Extended.Company = value ?? "";
  }

  public get Manager(): string {
    return this.model.Extended.Manager;
  }

  public set Manager(value: string) {
    this.model.Extended.Manager = value ?? "";
  }
}
