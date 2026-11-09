import { LegacyCells, LegacyCell } from "./cell";
import { LegacyChartCollection } from "./chartCollection";
import type {
  CellValue,
  Style,
  DataValidation,
  Comment,
  Hyperlink,
  ConditionalFormatRule,
  ShapeInfo,
  ChartInfo,
  ImageInfo,
} from "./types";
import { cellRef, parseCellRef, parseRange } from "./util";

export class LegacyWorksheet {
  private _name: string;
  private _index: number;
  private _cells = new LegacyCells();
  private _dataValidations: DataValidation[] = [];
  private _comments: Comment[] = [];
  private _hyperlinks: { cell: string; link: Hyperlink }[] = [];
  private _autoFilter?: {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
  };
  private _conditionalFormats: {
    range: string;
    rules: ConditionalFormatRule[];
  }[] = [];
  private _mergedCells: string[] = [];
  private _columnWidths: Map<number, number> = new Map();
  private _rowHeights: Map<number, number> = new Map();
  private _hiddenRows: Set<number> = new Set<number>();
  private _shapes: ShapeInfo[] = [];
  private _images: ImageInfo[] = [];
  private _charts: LegacyChartCollection;
  private _defaultColumnWidth?: number;
  private _defaultRowHeight?: number;
  _drawingIndex: number = 0;

  constructor(name: string, index: number) {
    this._name = name;
    this._index = index;
    this._charts = new LegacyChartCollection(index, this);
  }

  get name(): string {
    return this._name;
  }

  set name(name: string) {
    this._name = name;
  }

  get index(): number {
    return this._index;
  }

  get cells() {
    return this._cells;
  }

  getCell(row: number, col: number): LegacyCell | undefined {
    return this._cells.get(cellRef(row, col));
  }

  getCellByRef(ref: string): LegacyCell | undefined {
    return this._cells.get(ref);
  }

  getCell2(key: string): LegacyCell {
    if (/^[A-Z]+\d+$/.test(key)) {
      const { row, col } = parseCellRef(key);
      return this._cells.getOrCreate(row, col);
    }
    throw new Error(`Invalid cell reference: ${key}`);
  }

  putValue(key: string, value: CellValue): LegacyCell {
    const { row, col } = parseCellRef(key);
    const cell = this._cells.getOrCreate(row, col);
    cell.putValue(value);
    return cell;
  }

  get value(): CellValue | undefined {
    return undefined;
  }

  set value(v: CellValue) {
    this.putValue("A1", v);
  }

  addDataValidation(validation: DataValidation, range: string) {
    this._dataValidations.push({ ...validation, formula1: range });
  }

  get dataValidations() {
    return this._dataValidations;
  }

  addComment(comment: Comment) {
    this._comments.push(comment);
  }

  get comments() {
    return this._comments;
  }

  addHyperlink(cellRef: string, hyperlink: Hyperlink) {
    this._hyperlinks.push({ cell: cellRef, link: hyperlink });
  }

  get hyperlinks() {
    return this._hyperlinks;
  }

  setAutoFilter(range: string) {
    const r = parseRange(range);
    this._autoFilter = {
      startRow: r.startRow,
      startCol: r.startCol,
      endRow: r.endRow,
      endCol: r.endCol,
    };
  }

  get autoFilter() {
    return this._autoFilter;
  }

  removeAutoFilter() {
    this._autoFilter = undefined;
  }

  addConditionalFormat(range: string, rules: ConditionalFormatRule[]) {
    this._conditionalFormats.push({ range, rules });
  }

  get conditionalFormats() {
    return this._conditionalFormats;
  }

  mergeCells(range: string) {
    this._mergedCells.push(range);
  }

  unmergeCells(range: string) {
    const idx = this._mergedCells.indexOf(range);
    if (idx >= 0) this._mergedCells.splice(idx, 1);
  }

  get mergedCells() {
    return this._mergedCells;
  }

  setColumnWidth(column: number, width: number) {
    this._columnWidths.set(column, width);
  }

  getColumnWidth(column: number): number | undefined {
    return this._columnWidths.get(column);
  }

  setRowHeight(row: number, height: number) {
    this._rowHeights.set(row, height);
  }

  getRowHeight(row: number): number | undefined {
    return this._rowHeights.get(row);
  }

  setRowHidden(row: number, hidden: boolean) {
    if (hidden) {
      this._hiddenRows.add(row);
    } else {
      this._hiddenRows.delete(row);
    }
  }

  isRowHidden(row: number): boolean {
    return this._hiddenRows.has(row);
  }

  set defaultColumnWidth(width: number | undefined) {
    this._defaultColumnWidth = width;
  }

  get defaultColumnWidth(): number | undefined {
    return this._defaultColumnWidth;
  }

  set defaultRowHeight(height: number | undefined) {
    this._defaultRowHeight = height;
  }

  get defaultRowHeight(): number | undefined {
    return this._defaultRowHeight;
  }

  setCellStyle(cellRef: string, style: any) {
    const { row, col } = parseCellRef(cellRef);
    const cell = this._cells.getOrCreate(row, col);
    cell.setStyle(style);
  }

  addShape(shape: ShapeInfo) {
    this._shapes.push(shape);
  }

  get shapes(): ShapeInfo[] {
    return this._shapes;
  }

  get charts(): LegacyChartCollection {
    return this._charts;
  }

  addImage(image: ImageInfo) {
    this._images.push(image);
  }

  get images(): ImageInfo[] {
    return this._images;
  }

  toXml(drawingIndex: number = 0): string {
    const cols: string[] = [];
    if (this._columnWidths.size > 0) {
      const sortedCols = Array.from(this._columnWidths.entries()).sort((a, b) => a[0] - b[0]);
      
      let groupStart = 0;
      let groupWidth = sortedCols[0][1];
      
      for (let i = 1; i <= sortedCols.length; i++) {
        const currentWidth = i < sortedCols.length ? sortedCols[i][1] : -1;
        const prevWidth = sortedCols[i - 1][1];
        
        if (currentWidth !== prevWidth || i === sortedCols.length) {
          const minCol = sortedCols[groupStart][0] + 1;
          const maxCol = sortedCols[i - 1][0] + 1;
          const excelWidth = Math.round((prevWidth / 7) * 100) / 100;
          
          if (minCol === maxCol) {
            cols.push(`<col min="${minCol}" max="${maxCol}" width="${excelWidth}" customWidth="1"/>`);
          } else {
            cols.push(`<col min="${minCol}" max="${maxCol}" width="${excelWidth}" customWidth="1"/>`);
          }
          
          if (i < sortedCols.length) {
            groupStart = i;
            groupWidth = currentWidth;
          }
        }
      }
    }

    const rows = new Map<number, LegacyCell[]>();
    let minRow = Infinity,
      maxRow = -Infinity,
      minCol = Infinity,
      maxCol = -Infinity;
    for (const cell of this._cells) {
      const rowCells = rows.get(cell.row) || [];
      rowCells.push(cell);
      rows.set(cell.row, rowCells);
      if (cell.row < minRow) minRow = cell.row;
      if (cell.row > maxRow) maxRow = cell.row;
      if (cell.col < minCol) minCol = cell.col;
      if (cell.col > maxCol) maxCol = cell.col;
    }

    for (const rowNum of this._rowHeights.keys()) {
      if (rowNum < minRow) minRow = rowNum;
      if (rowNum > maxRow) maxRow = rowNum;
    }

    if (rows.size === 0) {
      minRow = 0;
      maxRow = 0;
      minCol = 0;
      maxCol = 0;
    }

    const dimension =
      maxRow >= 0
        ? `ref="${cellRef(minRow, minCol)}:${cellRef(maxRow, maxCol)}" `
        : "";

    let autoFilter = "";
    if (this._autoFilter) {
      const { startRow, startCol, endRow, endCol } = this._autoFilter;
      autoFilter = `<autoFilter ref="${cellRef(startRow, startCol)}:${cellRef(endRow, endCol)}"/>`;
    }

    const colsStr =
      this._columnWidths.size > 0 ? "<cols>" + cols.join("") + "</cols>" : "";

    const xmlnsAttrs = `xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{CA87E96B-18A1-4CCB-99B8-D994280CBAA2}"`;
    const sheetViews = `<sheetViews><sheetView workbookViewId="0"/></sheetViews>`;
    const sheetFormatPr = `<sheetFormatPr defaultRowHeight="${this._defaultRowHeight || 12.5}" defaultColWidth="${this._defaultColumnWidth || 8}" x14ac:dyDescent="0.25"/>`;
    let sheetDataWithAttrs = "";
    for (let rowNum = minRow; rowNum <= maxRow; rowNum++) {
      const rowCells = rows.get(rowNum) || [];
      const span = `${minCol + 1}:${maxCol + 1}`;
      const rowHeight = this._rowHeights.get(rowNum);
      const heightAttr = rowHeight ? ` ht="${rowHeight}" customHeight="1"` : "";
      const hiddenAttr = this._hiddenRows.has(rowNum) ? ' hidden="1"' : "";
      sheetDataWithAttrs += `<row r="${rowNum + 1}" spans="${span}" x14ac:dyDescent="0.25"${heightAttr}${hiddenAttr}>`;
      const sortedCells = rowCells.sort((a, b) => a.col - b.col);
      for (const cell of sortedCells) {
        sheetDataWithAttrs += cell.toXml();
      }
      sheetDataWithAttrs += "</row>";
    }

    if (minRow === Infinity) {
      sheetDataWithAttrs = "";
    }

    let drawingRel = "";
    if (drawingIndex > 0) {
      drawingRel = `<drawing r:id="rId1"/>`;
    }

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet ${xmlnsAttrs}><dimension ${dimension}/>${sheetViews}${sheetFormatPr}${colsStr}<sheetData>${sheetDataWithAttrs}</sheetData>${autoFilter}<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>${drawingRel}</worksheet>`;
  }

  getXml(drawingIndex: number = 0): string {
    return this.toXml(drawingIndex);
  }
}

import { AutoFilter } from "./autofilter";
import { Cells } from "./Cells";
import { ChartCollection } from "./chartCollection";
import { CellsException } from "./CellsException";
import { Color } from "./Color";
import { ColorValue } from "./Core/ColorValue";
import { CommentCollection } from "./CommentCollection";
import { ConditionalFormattingCollection } from "./ConditionalFormattingCollection";
import { createIndexerProxy } from "./IndexerProxy";
import { HyperlinkCollection } from "./HyperlinkCollection";
import { ListObjectCollection } from "./ListObjectCollection";
import { PageSetup } from "./PageSetup";
import { PictureCollection } from "./PictureCollection";
import { ShapeCollection } from "./ShapeCollection";
import { SheetVisibility } from "./Core/SheetVisibility";
import { ValidationCollection } from "./ValidationCollection";
import { VisibilityType } from "./enums";
import { Workbook } from "./workbook";
import { WorksheetModel } from "./Core/WorksheetModel";
import { WorksheetProtection } from "./WorksheetProtection";

/** Encapsulates a single worksheet and its supported worksheet features. */
export class Worksheet {
  private readonly workbook: Workbook;
  private readonly model: WorksheetModel;
  private readonly cells: Cells;
  private readonly hyperlinks: HyperlinkCollection;
  private readonly validations: ValidationCollection;
  private readonly conditionalFormattings: ConditionalFormattingCollection;
  private readonly pageSetup: PageSetup;
  private readonly protection: WorksheetProtection;
  private readonly autoFilter: AutoFilter;
  private readonly listObjects: ListObjectCollection;
  private readonly pictures: PictureCollection;
  private readonly shapes: ShapeCollection;
  private readonly charts: ChartCollection;
  private readonly comments: CommentCollection;

  public constructor(workbook: Workbook, model: WorksheetModel) {
    this.workbook = workbook;
    this.model = model;

    const cells = new Cells(this);
    this.cells = createIndexerProxy(cells, (key) =>
      typeof key === "string" ? cells.getByName(key) : undefined,
    );

    const hyperlinks = new HyperlinkCollection(model.Hyperlinks);
    this.hyperlinks = createIndexerProxy(hyperlinks, (key) =>
      typeof key === "number" ? hyperlinks.get(key) : undefined,
    );

    const validations = new ValidationCollection(model.Validations);
    this.validations = createIndexerProxy(validations, (key) =>
      typeof key === "number" ? validations.get(key) : undefined,
    );

    const conditionalFormattings = new ConditionalFormattingCollection(model.ConditionalFormattings);
    this.conditionalFormattings = createIndexerProxy(conditionalFormattings, (key) =>
      typeof key === "number" ? conditionalFormattings.get(key) : undefined,
    );

    this.pageSetup = new PageSetup(model.PageSetup);
    this.protection = new WorksheetProtection(model.Protection);
    this.autoFilter = new AutoFilter(model.AutoFilter);

    const listObjects = new ListObjectCollection(model.ListObjects);
    this.listObjects = createIndexerProxy(listObjects, (key) => {
      if (typeof key === "number") {
        return listObjects.get(key);
      }

      return listObjects.getByName(key);
    });

    const pictures = new PictureCollection(model.Pictures);
    this.pictures = createIndexerProxy(pictures, (key) =>
      typeof key === "number" ? pictures.get(key) : undefined,
    );

    const shapes = new ShapeCollection(model.Shapes);
    this.shapes = createIndexerProxy(shapes, (key) =>
      typeof key === "number" ? shapes.get(key) : undefined,
    );

    const charts = new ChartCollection(model.Charts);
    this.charts = createIndexerProxy(charts, (key) =>
      typeof key === "number" ? charts.get(key) : undefined,
    );

    const comments = new CommentCollection(model.Comments);
    this.comments = createIndexerProxy(comments, (key) =>
      typeof key === "number" ? comments.get(key) : comments.get(key),
    );
  }

  /** Gets the underlying core model. */
  public get Model(): WorksheetModel {
    return this.model;
  }

  /** Gets the owning workbook. */
  public get Workbook(): Workbook {
    return this.workbook;
  }

  /** Gets or sets the worksheet name. */
  public get Name(): string {
    return this.model.Name;
  }

  public set Name(value: string) {
    if (value === null || value === undefined || value.trim().length === 0) {
      throw new CellsException("Worksheet name must be non-empty.");
    }

    this.workbook.EnsureUniqueSheetName(value, this.model);
    this.model.Name = value;
  }

  /** Gets or sets the worksheet visibility state. */
  public get VisibilityType(): VisibilityType {
    switch (this.model.Visibility) {
      case SheetVisibility.Hidden:
        return VisibilityType.Hidden;
      case SheetVisibility.VeryHidden:
        return VisibilityType.VeryHidden;
      default:
        return VisibilityType.Visible;
    }
  }

  public set VisibilityType(value: VisibilityType) {
    switch (value) {
      case VisibilityType.Hidden:
        this.model.Visibility = SheetVisibility.Hidden;
        break;
      case VisibilityType.VeryHidden:
        this.model.Visibility = SheetVisibility.VeryHidden;
        break;
      default:
        this.model.Visibility = SheetVisibility.Visible;
        break;
    }
  }

  /** Gets or sets the worksheet tab color. */
  public get TabColor(): Color {
    if (this.model.TabColor !== null && this.model.TabColor !== undefined) {
      return Color.FromCore(this.model.TabColor);
    }

    return Color.Empty;
  }

  public set TabColor(value: Color) {
    if (value === null || value === undefined || value.Equals(Color.Empty)) {
      this.model.TabColor = null;
      return;
    }

    this.model.TabColor = value.ToCore() as unknown as ColorValue;
  }

  /** Gets or sets whether gridlines are shown. */
  public get ShowGridlines(): boolean {
    return this.model.View.ShowGridLines;
  }

  public set ShowGridlines(value: boolean) {
    this.model.View.ShowGridLines = value;
  }

  /** Gets or sets whether row and column headers are shown. */
  public get ShowRowColumnHeaders(): boolean {
    return this.model.View.ShowRowColumnHeaders;
  }

  public set ShowRowColumnHeaders(value: boolean) {
    this.model.View.ShowRowColumnHeaders = value;
  }

  /** Gets or sets whether zero values are shown. */
  public get ShowZeros(): boolean {
    return this.model.View.ShowZeros;
  }

  public set ShowZeros(value: boolean) {
    this.model.View.ShowZeros = value;
  }

  /** Gets or sets whether the worksheet view is right-to-left. */
  public get RightToLeft(): boolean {
    return this.model.View.RightToLeft;
  }

  public set RightToLeft(value: boolean) {
    this.model.View.RightToLeft = value;
  }

  /** Gets or sets the worksheet zoom percentage. */
  public get Zoom(): number {
    return this.model.View.ZoomScale;
  }

  public set Zoom(value: number) {
    if (value < 10 || value > 400) {
      throw new CellsException("Zoom must be between 10 and 400.");
    }

    this.model.View.ZoomScale = value;
  }

  public get Cells(): Cells {
    return this.cells;
  }

  public get Hyperlinks(): HyperlinkCollection {
    return this.hyperlinks;
  }

  public get Validations(): ValidationCollection {
    return this.validations;
  }

  public get ConditionalFormattings(): ConditionalFormattingCollection {
    return this.conditionalFormattings;
  }

  public get PageSetup(): PageSetup {
    return this.pageSetup;
  }

  public get Protection(): WorksheetProtection {
    return this.protection;
  }

  public get AutoFilter(): AutoFilter {
    return this.autoFilter;
  }

  public get ListObjects(): ListObjectCollection {
    return this.listObjects;
  }

  public get Pictures(): PictureCollection {
    return this.pictures;
  }

  public get Shapes(): ShapeCollection {
    return this.shapes;
  }

  public get Charts(): ChartCollection {
    return this.charts;
  }

  public get Comments(): CommentCollection {
    return this.comments;
  }

  /** Marks the worksheet as protected using the current protection settings. */
  public Protect(): void {
    this.model.Protection.IsProtected = true;
  }

  /** Clears worksheet protection and resets supported protection flags. */
  public Unprotect(): void {
    this.protection.Reset();
  }
}
