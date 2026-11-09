import { CellsException } from "./CellsException";
import { normalizeVisibility } from "./WorkbookPropertySupport";
import { WorkbookModel } from "./Core/WorkbookModel";

/** Represents the workbook window view settings. */
export class WorkbookView {
  private readonly workbookModel: WorkbookModel;

  public constructor(workbookModel: WorkbookModel) {
    this.workbookModel = workbookModel;
  }

  private get model(): WorkbookModel["Properties"]["View"] {
    return this.workbookModel.Properties.View;
  }

  public get XWindow(): number {
    return this.model.XWindow ?? 0;
  }

  public set XWindow(value: number) {
    this.model.XWindow = value;
  }

  public get YWindow(): number {
    return this.model.YWindow ?? 0;
  }

  public set YWindow(value: number) {
    this.model.YWindow = value;
  }

  public get WindowWidth(): number {
    return this.model.WindowWidth ?? 0;
  }

  public set WindowWidth(value: number) {
    if (value < 0) {
      throw new CellsException("WindowWidth must be non-negative.");
    }

    this.model.WindowWidth = value;
  }

  public get WindowHeight(): number {
    return this.model.WindowHeight ?? 0;
  }

  public set WindowHeight(value: number) {
    if (value < 0) {
      throw new CellsException("WindowHeight must be non-negative.");
    }

    this.model.WindowHeight = value;
  }

  public get ActiveTab(): number {
    return this.workbookModel.ActiveSheetIndex;
  }

  public set ActiveTab(value: number) {
    if (value < 0 || value >= this.workbookModel.Worksheets.length) {
      throw new CellsException("ActiveTab must refer to an existing worksheet.");
    }

    this.workbookModel.ActiveSheetIndex = value;
  }

  public get FirstSheet(): number {
    return this.model.FirstSheet ?? 0;
  }

  public set FirstSheet(value: number) {
    if (value < 0) {
      throw new CellsException("FirstSheet must be non-negative.");
    }

    this.model.FirstSheet = value;
  }

  public get ShowHorizontalScroll(): boolean {
    return this.model.ShowHorizontalScroll === null || this.model.ShowHorizontalScroll;
  }

  public set ShowHorizontalScroll(value: boolean) {
    this.model.ShowHorizontalScroll = value;
  }

  public get ShowVerticalScroll(): boolean {
    return this.model.ShowVerticalScroll === null || this.model.ShowVerticalScroll;
  }

  public set ShowVerticalScroll(value: boolean) {
    this.model.ShowVerticalScroll = value;
  }

  public get ShowSheetTabs(): boolean {
    return this.model.ShowSheetTabs === null || this.model.ShowSheetTabs;
  }

  public set ShowSheetTabs(value: boolean) {
    this.model.ShowSheetTabs = value;
  }

  public get TabRatio(): number {
    return this.model.TabRatio ?? 600;
  }

  public set TabRatio(value: number) {
    if (value < 0 || value > 1000) {
      throw new CellsException("TabRatio must be between 0 and 1000.");
    }

    this.model.TabRatio = value;
  }

  public get Visibility(): string {
    return this.model.Visibility === null || this.model.Visibility.length === 0
      ? "visible"
      : this.model.Visibility;
  }

  public set Visibility(value: string) {
    this.model.Visibility = normalizeVisibility(value);
  }

  public get Minimized(): boolean {
    return this.model.Minimized;
  }

  public set Minimized(value: boolean) {
    this.model.Minimized = value;
  }

  public get AutoFilterDateGrouping(): boolean {
    return this.model.AutoFilterDateGrouping;
  }

  public set AutoFilterDateGrouping(value: boolean) {
    this.model.AutoFilterDateGrouping = value;
  }
}
