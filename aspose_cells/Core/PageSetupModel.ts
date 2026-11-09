import { HeaderFooterModel } from "./HeaderFooterModel";
import { PageMarginsModel } from "./PageMarginsModel";
import { PageOrientation } from "./PageOrientation";
import { PrintOptionsModel } from "./PrintOptionsModel";

export class PageSetupModel {
  public readonly Margins: PageMarginsModel = new PageMarginsModel();
  public readonly PrintOptions: PrintOptionsModel = new PrintOptionsModel();
  public readonly HeaderFooter: HeaderFooterModel = new HeaderFooterModel();
  public PaperSize: number = 0;
  public Orientation: PageOrientation = PageOrientation.Default;
  public FirstPageNumber: number | null = null;
  public Scale: number | null = null;
  public FitToPage: boolean = false;
  public FitToWidth: number | null = null;
  public FitToHeight: number | null = null;
  public PageOrderOverThenDown: boolean = false;
  public PrintArea: string | null = null;
  public PrintTitleRows: string | null = null;
  public PrintTitleColumns: string | null = null;
  public readonly HorizontalPageBreaks: number[] = [];
  public readonly VerticalPageBreaks: number[] = [];
}
