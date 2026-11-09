import { CoreDocumentPropertiesModel } from "./CoreDocumentPropertiesModel";
import { ExtendedDocumentPropertiesModel } from "./ExtendedDocumentPropertiesModel";

export class DocumentPropertiesModel {
  public readonly Core: CoreDocumentPropertiesModel;
  public readonly Extended: ExtendedDocumentPropertiesModel;

  public constructor() {
    this.Core = new CoreDocumentPropertiesModel();
    this.Extended = new ExtendedDocumentPropertiesModel();
  }

  public CopyFrom(source: DocumentPropertiesModel): void {
    this.Core.CopyFrom(source.Core);
    this.Extended.CopyFrom(source.Extended);
  }

  public HasStoredState(): boolean {
    return this.Core.HasStoredState() || this.Extended.HasStoredState();
  }
}
