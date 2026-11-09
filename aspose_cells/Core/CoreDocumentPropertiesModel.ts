export class CoreDocumentPropertiesModel {
  public Title: string = "";
  public Subject: string = "";
  public Creator: string = "";
  public Keywords: string = "";
  public Description: string = "";
  public LastModifiedBy: string = "";
  public Revision: string = "";
  public Category: string = "";
  public ContentStatus: string = "";
  public Created: Date | null = null;
  public Modified: Date | null = null;

  public CopyFrom(source: CoreDocumentPropertiesModel): void {
    this.Title = source.Title;
    this.Subject = source.Subject;
    this.Creator = source.Creator;
    this.Keywords = source.Keywords;
    this.Description = source.Description;
    this.LastModifiedBy = source.LastModifiedBy;
    this.Revision = source.Revision;
    this.Category = source.Category;
    this.ContentStatus = source.ContentStatus;
    this.Created = source.Created;
    this.Modified = source.Modified;
  }

  public HasStoredState(): boolean {
    return (
      (this.Title != null && this.Title.length > 0) ||
      (this.Subject != null && this.Subject.length > 0) ||
      (this.Creator != null && this.Creator.length > 0) ||
      (this.Keywords != null && this.Keywords.length > 0) ||
      (this.Description != null && this.Description.length > 0) ||
      (this.LastModifiedBy != null && this.LastModifiedBy.length > 0) ||
      (this.Revision != null && this.Revision.length > 0) ||
      (this.Category != null && this.Category.length > 0) ||
      (this.ContentStatus != null && this.ContentStatus.length > 0) ||
      this.Created !== null ||
      this.Modified !== null
    );
  }
}
