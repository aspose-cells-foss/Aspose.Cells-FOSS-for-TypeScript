import { CoreDocumentPropertiesModel } from "./Core/CoreDocumentPropertiesModel";

/** Represents core document properties. */
export class CoreDocumentProperties {
  private readonly model: CoreDocumentPropertiesModel;

  public constructor(model: CoreDocumentPropertiesModel) {
    this.model = model;
  }

  public get Title(): string {
    return this.model.Title;
  }

  public set Title(value: string) {
    this.model.Title = value ?? "";
  }

  public get Subject(): string {
    return this.model.Subject;
  }

  public set Subject(value: string) {
    this.model.Subject = value ?? "";
  }

  public get Creator(): string {
    return this.model.Creator;
  }

  public set Creator(value: string) {
    this.model.Creator = value ?? "";
  }

  public get Keywords(): string {
    return this.model.Keywords;
  }

  public set Keywords(value: string) {
    this.model.Keywords = value ?? "";
  }

  public get Description(): string {
    return this.model.Description;
  }

  public set Description(value: string) {
    this.model.Description = value ?? "";
  }

  public get LastModifiedBy(): string {
    return this.model.LastModifiedBy;
  }

  public set LastModifiedBy(value: string) {
    this.model.LastModifiedBy = value ?? "";
  }

  public get Revision(): string {
    return this.model.Revision;
  }

  public set Revision(value: string) {
    this.model.Revision = value ?? "";
  }

  public get Category(): string {
    return this.model.Category;
  }

  public set Category(value: string) {
    this.model.Category = value ?? "";
  }

  public get ContentStatus(): string {
    return this.model.ContentStatus;
  }

  public set ContentStatus(value: string) {
    this.model.ContentStatus = value ?? "";
  }

  public get Created(): Date | null {
    return this.model.Created;
  }

  public set Created(value: Date | null) {
    this.model.Created = value;
  }

  public get Modified(): Date | null {
    return this.model.Modified;
  }

  public set Modified(value: Date | null) {
    this.model.Modified = value;
  }
}
