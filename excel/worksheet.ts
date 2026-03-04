import { Cells, Cell } from "./cell";
import type {
  CellValue,
  Style,
  DataValidation,
  Comment,
  Hyperlink,
  ConditionalFormatRule,
} from "./types";
import { cellRef, parseCellRef, parseRange } from "./util";

export class Worksheet {
  private _name: string;
  private _index: number;
  private _cells = new Cells();
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

  constructor(name: string, index: number) {
    this._name = name;
    this._index = index;
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

  getCell(row: number, col: number): Cell | undefined {
    return this._cells.get(cellRef(row, col));
  }

  getCellByRef(ref: string): Cell | undefined {
    return this._cells.get(ref);
  }

  getCell2(key: string): Cell {
    if (/^[A-Z]+\d+$/.test(key)) {
      const { row, col } = parseCellRef(key);
      return this._cells.getOrCreate(row, col);
    }
    throw new Error(`Invalid cell reference: ${key}`);
  }

  putValue(key: string, value: CellValue): Cell {
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

  toXml(): string {
    const cols: string[] = [];
    for (const [col, width] of this._columnWidths) {
      cols.push(
        `<col min="${col + 1}" max="${col + 1}" width="${width}" customWidth="1"/>`,
      );
    }

    const rows = new Map<number, Cell[]>();
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

    const colsStr = cols.length > 0 ? "<cols>" + cols.join("") + "</cols>" : "";

    const xmlnsAttrs = `xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{CA87E96B-18A1-4CCB-99B8-D994280CBAA2}"`;
    const sheetViews = `<sheetViews><sheetView tabSelected="1" workbookViewId="0" topLeftCell="A1" /></sheetViews>`;
    const sheetFormatPr = `<sheetFormatPr defaultRowHeight="12.75" />`;
    let sheetDataWithAttrs = "";
    for (const [rowNum, rowCells] of rows) {
      const height = this._rowHeights.get(rowNum);
      const ht = height !== undefined ? height : 12.75;
      const span = `${minCol + 1}:${maxCol + 1}`;
      sheetDataWithAttrs += `<row r="${rowNum + 1}" spans="${span}" ht="${ht}">`;
      const sortedCells = rowCells.sort((a, b) => a.col - b.col);
      for (const cell of sortedCells) {
        sheetDataWithAttrs += cell.toXml();
      }
      sheetDataWithAttrs += "</row>";
    }
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet ${xmlnsAttrs}><dimension ${dimension}/>${sheetViews}${sheetFormatPr}${colsStr}<sheetData>${sheetDataWithAttrs}</sheetData>${autoFilter}<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5" /></worksheet>`;
  }

  getXml(): string {
    return this.toXml();
  }
}
