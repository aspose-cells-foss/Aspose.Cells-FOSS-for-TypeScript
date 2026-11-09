export class ListColumnModel {
  public readonly Id: number;
  public Name: string;
  public TotalsRowFunction: string;
  public TotalsRowLabel: string;
  public TotalsRowFormula: string;

  public constructor(id: number, name: string) {
    this.Id = id;
    this.Name = name;
    this.TotalsRowFunction = "none";
    this.TotalsRowLabel = "";
    this.TotalsRowFormula = "";
  }
}
