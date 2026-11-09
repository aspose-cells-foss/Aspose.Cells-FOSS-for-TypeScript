import { AutoFilterSortStateModel } from "./AutoFilterSortStateModel";
import { FilterColumnModel } from "./FilterColumnModel";

export class AutoFilterModel {
  public Range: string;
  public readonly FilterColumns: FilterColumnModel[];
  public readonly SortState: AutoFilterSortStateModel;

  public constructor() {
    this.Range = "";
    this.FilterColumns = [];
    this.SortState = new AutoFilterSortStateModel();
  }

  public Clear(): void {
    this.Range = "";
    this.FilterColumns.length = 0;
    this.SortState.Clear();
  }

  public HasStoredState(): boolean {
    return this.Range != null && this.Range.length > 0;
  }
}
