export class WorkbookViewModel {
  public XWindow: number | null = null;
  public YWindow: number | null = null;
  public WindowWidth: number | null = null;
  public WindowHeight: number | null = null;
  public FirstSheet: number | null = null;
  public ShowHorizontalScroll: boolean | null = null;
  public ShowVerticalScroll: boolean | null = null;
  public ShowSheetTabs: boolean | null = null;
  public TabRatio: number | null = null;
  public Visibility: string = "";
  public Minimized: boolean = false;
  public AutoFilterDateGrouping: boolean = true;

  public CopyFrom(source: WorkbookViewModel): void {
    this.XWindow = source.XWindow;
    this.YWindow = source.YWindow;
    this.WindowWidth = source.WindowWidth;
    this.WindowHeight = source.WindowHeight;
    this.FirstSheet = source.FirstSheet;
    this.ShowHorizontalScroll = source.ShowHorizontalScroll;
    this.ShowVerticalScroll = source.ShowVerticalScroll;
    this.ShowSheetTabs = source.ShowSheetTabs;
    this.TabRatio = source.TabRatio;
    this.Visibility = source.Visibility;
    this.Minimized = source.Minimized;
    this.AutoFilterDateGrouping = source.AutoFilterDateGrouping;
  }

  public HasStoredState(activeSheetIndex: number): boolean {
    return (
      activeSheetIndex > 0 ||
      this.XWindow !== null ||
      this.YWindow !== null ||
      this.WindowWidth !== null ||
      this.WindowHeight !== null ||
      this.FirstSheet !== null ||
      this.ShowHorizontalScroll !== null ||
      this.ShowVerticalScroll !== null ||
      this.ShowSheetTabs !== null ||
      this.TabRatio !== null ||
      (this.Visibility != null && this.Visibility.length > 0) ||
      this.Minimized ||
      !this.AutoFilterDateGrouping
    );
  }
}
