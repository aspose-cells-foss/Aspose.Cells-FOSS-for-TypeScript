import { ListColumnModel } from "./ListColumnModel";

export class ListObjectModel {
  public DisplayName: string;
  public Name: string;
  public Comment: string;
  public StartRow: number = 0;
  public StartColumn: number = 0;
  public EndRow: number = 0;
  public EndColumn: number = 0;
  public ShowHeaderRow: boolean;
  public ShowTotals: boolean = false;
  public HasAutoFilter: boolean = false;
  public TableStyleName: string;
  public ShowFirstColumn: boolean = false;
  public ShowLastColumn: boolean = false;
  public ShowRowStripes: boolean;
  public ShowColumnStripes: boolean = false;
  public readonly Columns: ListColumnModel[];

  public constructor() {
    this.Columns = [];
    this.DisplayName = "";
    this.Name = "";
    this.Comment = "";
    this.TableStyleName = "";
    this.ShowHeaderRow = true;
    this.ShowRowStripes = true;
  }
}
