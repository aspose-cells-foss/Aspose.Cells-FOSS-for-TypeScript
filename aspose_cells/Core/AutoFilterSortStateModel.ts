import { AutoFilterSortConditionModel } from "./AutoFilterSortConditionModel";

export class AutoFilterSortStateModel {
  public ColumnSort: boolean;
  public CaseSensitive: boolean;
  public SortMethod: string;
  public Ref: string;
  public readonly Conditions: AutoFilterSortConditionModel[];

  public constructor() {
    this.Ref = "";
    this.SortMethod = "";
    this.Conditions = [];
    this.ColumnSort = false;
    this.CaseSensitive = false;
  }

  public Clear(): void {
    this.ColumnSort = false;
    this.CaseSensitive = false;
    this.SortMethod = "";
    this.Ref = "";
    this.Conditions.length = 0;
  }

  public HasStoredState(): boolean {
    return (
      (this.Ref != null && this.Ref.length > 0) ||
      this.ColumnSort ||
      this.CaseSensitive ||
      (this.SortMethod != null && this.SortMethod.length > 0) ||
      this.Conditions.length > 0
    );
  }
}
