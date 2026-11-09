export class WorksheetViewModel {
  public ShowGridLines: boolean = true;
  public ShowRowColumnHeaders: boolean = true;
  public ShowZeros: boolean = true;
  public RightToLeft: boolean = false;
  public ZoomScale: number = 100;
  public ViewType: string | null = null;
  public TopLeftCell: string | null = null;
  public TabSelected: boolean = false;
  public SelectionActiveCell: string | null = null;
  public SelectionSqref: string | null = null;
}
