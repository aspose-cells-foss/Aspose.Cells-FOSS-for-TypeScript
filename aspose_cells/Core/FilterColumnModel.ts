import { AutoFilterColorFilterModel } from "./AutoFilterColorFilterModel";
import { AutoFilterCustomFilterModel } from "./AutoFilterCustomFilterModel";
import { AutoFilterDynamicFilterModel } from "./AutoFilterDynamicFilterModel";
import { AutoFilterTop10Model } from "./AutoFilterTop10Model";

export class FilterColumnModel {
  public ColumnIndex: number = 0;
  public HiddenButton: boolean = false;
  public readonly Filters: string[];
  public readonly CustomFilters: AutoFilterCustomFilterModel[];
  public CustomFiltersAnd: boolean = false;
  public readonly ColorFilter: AutoFilterColorFilterModel;
  public readonly DynamicFilter: AutoFilterDynamicFilterModel;
  public readonly Top10: AutoFilterTop10Model;

  public constructor() {
    this.Filters = [];
    this.CustomFilters = [];
    this.ColorFilter = new AutoFilterColorFilterModel();
    this.DynamicFilter = new AutoFilterDynamicFilterModel();
    this.Top10 = new AutoFilterTop10Model();
  }

  public ClearCriteria(): void {
    this.HiddenButton = false;
    this.Filters.length = 0;
    this.CustomFilters.length = 0;
    this.CustomFiltersAnd = false;
    this.ColorFilter.Clear();
    this.DynamicFilter.Clear();
    this.Top10.Clear();
  }

  public HasStoredState(): boolean {
    return (
      this.HiddenButton ||
      this.Filters.length > 0 ||
      this.CustomFilters.length > 0 ||
      this.ColorFilter.Enabled ||
      this.DynamicFilter.Enabled ||
      this.Top10.Enabled
    );
  }
}
