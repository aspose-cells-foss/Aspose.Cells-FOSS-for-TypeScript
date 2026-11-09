import { CellsException } from "./CellsException";
import { PageOrientationType, PaperSizeType } from "./enums";
import { PageSetupModel } from "./Core/PageSetupModel";
import { PageOrientation } from "./Core/PageOrientation";

/** Represents worksheet print and page-layout settings. */
export class PageSetup {
  private static readonly CentimetersPerInch = 2.54;
  private readonly model: PageSetupModel;

  public constructor(model: PageSetupModel) {
    this.model = model;
  }

  /** Gets the underlying core model. */
  public get Model(): PageSetupModel {
    return this.model;
  }

  public get PaperSize(): PaperSizeType {
    return this.model.PaperSize as PaperSizeType;
  }

  public set PaperSize(value: PaperSizeType) {
    this.model.PaperSize = value;
  }

  public get Orientation(): PageOrientationType {
    switch (this.model.Orientation) {
      case PageOrientation.Portrait:
        return PageOrientationType.Portrait;
      case PageOrientation.Landscape:
        return PageOrientationType.Landscape;
      default:
        return PageOrientationType.Default;
    }
  }

  public set Orientation(value: PageOrientationType) {
    switch (value) {
      case PageOrientationType.Portrait:
        this.model.Orientation = PageOrientation.Portrait;
        break;
      case PageOrientationType.Landscape:
        this.model.Orientation = PageOrientation.Landscape;
        break;
      default:
        this.model.Orientation = PageOrientation.Default;
        break;
    }
  }

  public get FirstPageNumber(): number | null {
    return this.model.FirstPageNumber;
  }

  public set FirstPageNumber(value: number | null) {
    if (value !== null && value !== undefined && value <= 0) {
      throw new CellsException("FirstPageNumber must be positive.");
    }

    this.model.FirstPageNumber = value;
  }

  public get Scale(): number | null {
    return this.model.Scale;
  }

  public set Scale(value: number | null) {
    if (value !== null && value !== undefined && (value < 10 || value > 400)) {
      throw new CellsException("Scale must be between 10 and 400.");
    }

    this.model.Scale = value;
    if (value !== null && value !== undefined) {
      this.model.FitToPage = false;
    }
  }

  public get FitToPagesWide(): number | null {
    return this.model.FitToWidth;
  }

  public set FitToPagesWide(value: number | null) {
    if (value !== null && value !== undefined && value < 0) {
      throw new CellsException("FitToPagesWide must be zero or greater.");
    }

    this.model.FitToWidth = value;
    if (value !== null && value !== undefined) {
      this.model.FitToPage = true;
    }
  }

  public get FitToPagesTall(): number | null {
    return this.model.FitToHeight;
  }

  public set FitToPagesTall(value: number | null) {
    if (value !== null && value !== undefined && value < 0) {
      throw new CellsException("FitToPagesTall must be zero or greater.");
    }

    this.model.FitToHeight = value;
    if (value !== null && value !== undefined) {
      this.model.FitToPage = true;
    }
  }

  public get PrintArea(): string {
    return this.model.PrintArea ?? "";
  }

  public set PrintArea(value: string | null) {
    this.model.PrintArea = PageSetup.normalizeText(value);
  }

  public get PrintTitleRows(): string {
    return this.model.PrintTitleRows ?? "";
  }

  public set PrintTitleRows(value: string | null) {
    this.model.PrintTitleRows = PageSetup.normalizeText(value);
  }

  public get PrintTitleColumns(): string {
    return this.model.PrintTitleColumns ?? "";
  }

  public set PrintTitleColumns(value: string | null) {
    this.model.PrintTitleColumns = PageSetup.normalizeText(value);
  }

  public get LeftMargin(): number {
    return this.model.Margins.Left * PageSetup.CentimetersPerInch;
  }

  public set LeftMargin(value: number) {
    this.model.Margins.Left = PageSetup.validateMargin(value / PageSetup.CentimetersPerInch, "LeftMargin");
  }

  public get RightMargin(): number {
    return this.model.Margins.Right * PageSetup.CentimetersPerInch;
  }

  public set RightMargin(value: number) {
    this.model.Margins.Right = PageSetup.validateMargin(value / PageSetup.CentimetersPerInch, "RightMargin");
  }

  public get TopMargin(): number {
    return this.model.Margins.Top * PageSetup.CentimetersPerInch;
  }

  public set TopMargin(value: number) {
    this.model.Margins.Top = PageSetup.validateMargin(value / PageSetup.CentimetersPerInch, "TopMargin");
  }

  public get BottomMargin(): number {
    return this.model.Margins.Bottom * PageSetup.CentimetersPerInch;
  }

  public set BottomMargin(value: number) {
    this.model.Margins.Bottom = PageSetup.validateMargin(value / PageSetup.CentimetersPerInch, "BottomMargin");
  }

  public get HeaderMargin(): number {
    return this.model.Margins.Header * PageSetup.CentimetersPerInch;
  }

  public set HeaderMargin(value: number) {
    this.model.Margins.Header = PageSetup.validateMargin(value / PageSetup.CentimetersPerInch, "HeaderMargin");
  }

  public get FooterMargin(): number {
    return this.model.Margins.Footer * PageSetup.CentimetersPerInch;
  }

  public set FooterMargin(value: number) {
    this.model.Margins.Footer = PageSetup.validateMargin(value / PageSetup.CentimetersPerInch, "FooterMargin");
  }

  public get LeftMarginInch(): number {
    return this.model.Margins.Left;
  }

  public set LeftMarginInch(value: number) {
    this.model.Margins.Left = PageSetup.validateMargin(value, "LeftMarginInch");
  }

  public get RightMarginInch(): number {
    return this.model.Margins.Right;
  }

  public set RightMarginInch(value: number) {
    this.model.Margins.Right = PageSetup.validateMargin(value, "RightMarginInch");
  }

  public get TopMarginInch(): number {
    return this.model.Margins.Top;
  }

  public set TopMarginInch(value: number) {
    this.model.Margins.Top = PageSetup.validateMargin(value, "TopMarginInch");
  }

  public get BottomMarginInch(): number {
    return this.model.Margins.Bottom;
  }

  public set BottomMarginInch(value: number) {
    this.model.Margins.Bottom = PageSetup.validateMargin(value, "BottomMarginInch");
  }

  public get HeaderMarginInch(): number {
    return this.model.Margins.Header;
  }

  public set HeaderMarginInch(value: number) {
    this.model.Margins.Header = PageSetup.validateMargin(value, "HeaderMarginInch");
  }

  public get FooterMarginInch(): number {
    return this.model.Margins.Footer;
  }

  public set FooterMarginInch(value: number) {
    this.model.Margins.Footer = PageSetup.validateMargin(value, "FooterMarginInch");
  }

  public get LeftHeader(): string {
    return this.model.HeaderFooter.LeftHeader ?? "";
  }

  public set LeftHeader(value: string | null) {
    this.model.HeaderFooter.LeftHeader = PageSetup.normalizeText(value);
  }

  public get CenterHeader(): string {
    return this.model.HeaderFooter.CenterHeader ?? "";
  }

  public set CenterHeader(value: string | null) {
    this.model.HeaderFooter.CenterHeader = PageSetup.normalizeText(value);
  }

  public get RightHeader(): string {
    return this.model.HeaderFooter.RightHeader ?? "";
  }

  public set RightHeader(value: string | null) {
    this.model.HeaderFooter.RightHeader = PageSetup.normalizeText(value);
  }

  public get LeftFooter(): string {
    return this.model.HeaderFooter.LeftFooter ?? "";
  }

  public set LeftFooter(value: string | null) {
    this.model.HeaderFooter.LeftFooter = PageSetup.normalizeText(value);
  }

  public get CenterFooter(): string {
    return this.model.HeaderFooter.CenterFooter ?? "";
  }

  public set CenterFooter(value: string | null) {
    this.model.HeaderFooter.CenterFooter = PageSetup.normalizeText(value);
  }

  public get RightFooter(): string {
    return this.model.HeaderFooter.RightFooter ?? "";
  }

  public set RightFooter(value: string | null) {
    this.model.HeaderFooter.RightFooter = PageSetup.normalizeText(value);
  }

  public get PrintGridlines(): boolean {
    return this.model.PrintOptions.GridLines;
  }

  public set PrintGridlines(value: boolean) {
    this.model.PrintOptions.GridLines = value;
  }

  public get PrintHeadings(): boolean {
    return this.model.PrintOptions.Headings;
  }

  public set PrintHeadings(value: boolean) {
    this.model.PrintOptions.Headings = value;
  }

  public get CenterHorizontally(): boolean {
    return this.model.PrintOptions.HorizontalCentered;
  }

  public set CenterHorizontally(value: boolean) {
    this.model.PrintOptions.HorizontalCentered = value;
  }

  public get CenterVertically(): boolean {
    return this.model.PrintOptions.VerticalCentered;
  }

  public set CenterVertically(value: boolean) {
    this.model.PrintOptions.VerticalCentered = value;
  }

  public get HorizontalPageBreaks(): readonly number[] {
    return [...this.model.HorizontalPageBreaks].sort((left, right) => left - right);
  }

  public get VerticalPageBreaks(): readonly number[] {
    return [...this.model.VerticalPageBreaks].sort((left, right) => left - right);
  }

  public AddHorizontalPageBreak(rowIndex: number): void {
    if (rowIndex < 0) {
      throw new CellsException("Horizontal page break row index must be non-negative.");
    }

    if (!this.model.HorizontalPageBreaks.includes(rowIndex)) {
      this.model.HorizontalPageBreaks.push(rowIndex);
    }
  }

  public AddVerticalPageBreak(columnIndex: number): void {
    if (columnIndex < 0) {
      throw new CellsException("Vertical page break column index must be non-negative.");
    }

    if (!this.model.VerticalPageBreaks.includes(columnIndex)) {
      this.model.VerticalPageBreaks.push(columnIndex);
    }
  }

  public ClearHorizontalPageBreaks(): void {
    this.model.HorizontalPageBreaks.length = 0;
  }

  public ClearVerticalPageBreaks(): void {
    this.model.VerticalPageBreaks.length = 0;
  }

  private static validateMargin(value: number, propertyName: string): number {
    if (value < 0) {
      throw new CellsException(`${propertyName} must be zero or greater.`);
    }

    return value;
  }

  private static normalizeText(value: string | null): string | null {
    if (value === null || value === undefined || value.trim().length === 0) {
      return null;
    }

    return value.trim();
  }
}
