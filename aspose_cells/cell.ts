import type { CellValue, Style as LegacyStyleType } from "./types";
import { cellRef as ref, escapeXml } from "./util";

function dateToExcelSerial(date: Date): number {
  const excelEpoch = new Date(1899, 11, 30);
  const diff = date.getTime() - excelEpoch.getTime();
  return Math.floor(diff / (1000 * 60 * 60 * 24));
}

export class LegacyCell {
  static sharedStringMap: Map<string, number> = new Map();
  static setSharedStrings(strings: string[]) {
    LegacyCell.sharedStringMap = new Map();
    strings.forEach((s, i) => LegacyCell.sharedStringMap.set(s, i));
  }
  private _row: number;
  private _col: number;
  private _value: CellValue = null;
  private _formula?: string;
  private _style?: LegacyStyleType;
  private _styleIndex?: number;
  private _hyperlink?: string;

  constructor(row: number, col: number, value?: CellValue) {
    this._row = row;
    this._col = col;
    if (value !== undefined) this._value = value;
  }

  get row() {
    return this._row;
  }

  get col() {
    return this._col;
  }

  get ref() {
    return ref(this._row, this._col);
  }

  get value(): CellValue {
    return this._value;
  }

  set value(v: CellValue) {
    this._value = v;
  }

  putValue(value: CellValue) {
    this._value = value;
    return this;
  }

  get formula(): string | undefined {
    return this._formula;
  }

  set formula(f: string | undefined) {
    this._formula = f;
  }

  setFormula(formula: string) {
    this._formula = formula;
    return this;
  }

  get style(): LegacyStyleType | undefined {
    return this._style;
  }

  set style(s: LegacyStyleType | undefined) {
    this._style = s;
  }

  setStyle(style: LegacyStyleType) {
    this._style = style;
    return this;
  }

  setStyleIndex(index: number) {
    this._styleIndex = index;
  }

  get styleIndex(): number | undefined {
    return this._styleIndex;
  }

  get hyperlink(): string | undefined {
    return this._hyperlink;
  }

  setHyperlink(url: string) {
    this._hyperlink = url;
    return this;
  }

  toXml(): string {
    const coords = ref(this._row, this._col);
    let attrs = `r="${coords}"`;
    let valueXml = "";
    let typeAttr = "";

    if (this._value !== null && this._value !== undefined) {
      if (typeof this._value === "string") {
        const idx = LegacyCell.sharedStringMap.get(this._value);
        if (idx !== undefined) {
          typeAttr = 't="s"';
          valueXml = `<v>${idx}</v>`;
        } else {
          typeAttr = 't="inlineStr"';
          valueXml = `<is><t>${escapeXml(this._value)}</t></is>`;
        }
      } else if (typeof this._value === "number") {
        valueXml = `<v>${this._value}</v>`;
      } else if (typeof this._value === "boolean") {
        valueXml = `<v>${this._value ? 1 : 0}</v>`;
      } else if (this._value instanceof Date) {
        const serial = dateToExcelSerial(this._value);
        valueXml = `<v>${serial}</v>`;
      }
    }

    if (this._formula) {
      valueXml = `<f>${escapeXml(this._formula)}</f>${valueXml}`;
    }

    const styleAttr =
      this._styleIndex !== undefined && this._styleIndex !== 0
        ? ` s="${this._styleIndex}"`
        : "";

    let attrsStr = " " + attrs;
    if (typeAttr) attrsStr += " " + typeAttr;
    if (styleAttr) attrsStr += styleAttr;

    return `<c${attrsStr}>${valueXml}</c>`;
  }
}

export class LegacyCells {
  private cells: Map<string, LegacyCell> = new Map();

  get(key: string): LegacyCell | undefined {
    return this.cells.get(key);
  }

  set(key: string, cell: LegacyCell) {
    this.cells.set(key, cell);
    return this;
  }

  getOrCreate(row: number, col: number): LegacyCell {
    const key = ref(row, col);
    let cell = this.cells.get(key);
    if (!cell) {
      cell = new LegacyCell(row, col);
      this.cells.set(key, cell);
    }
    return cell;
  }

  get count() {
    return this.cells.size;
  }

  *[Symbol.iterator]() {
    for (const cell of this.cells.values()) {
      yield cell;
    }
  }

  values() {
    return this.cells.values();
  }

  toXml(): string {
    const rows = new Map<number, LegacyCell[]>();

    for (const cell of this.cells.values()) {
      const rowCells = rows.get(cell.row) || [];
      rowCells.push(cell);
      rows.set(cell.row, rowCells);
    }

    let xml = "";
    for (const [rowNum, rowCells] of rows) {
      xml += `<row r="${rowNum + 1}">`;
      const sortedCells = rowCells.sort((a, b) => a.col - b.col);
      for (const cell of sortedCells) {
        xml += cell.toXml();
      }
      xml += "</row>";
    }

    return xml;
  }
}

import { CellAddress } from "./Core/CellAddress";
import { findCellRecord } from "./CellStore";
import { RichTextRunValue } from "./Core/RichTextRunValue";
import { CellRecord } from "./Core/CellRecord";
import { CellValueKind } from "./Core/CellValueKind";
import { CellValueType, FontUnderlineType } from "./enums";
import { Color } from "./Color";
import { CellsException } from "./CellsException";
import { FontSetting } from "./FontSetting";
import { Style } from "./style";
import { StyleFlag } from "./StyleFlag";
import { DisplayText } from "./DisplayTextFormatter";
import { getOrCreateCellRecord } from "./CellStore";
import { Worksheet } from "./worksheet";

/** Represents a single worksheet cell and exposes value, formula, and style operations. */
export class Cell {
  private readonly worksheet: Worksheet;
  private readonly address: CellAddress;

  public constructor(worksheet: Worksheet, address: CellAddress) {
    this.worksheet = worksheet;
    this.address = address;
  }

  /** Gets or sets the logical cell value. */
  public get Value(): unknown {
    const record = this.tryGetRecord();
    return record === undefined ? null : record.Value;
  }

  public set Value(value: unknown) {
    this.assignValue(value);
  }

  /** Gets a stable string representation without style-based display formatting. */
  public get StringValue(): string {
    const record = this.tryGetRecord();
    return DisplayText.FormatStringValue(record === undefined ? null : record.Value);
  }

  /** Gets the display text generated from the value, style, and workbook culture. */
  public get DisplayStringValue(): string {
    const record = this.tryGetRecord();
    const workbook = this.worksheet.Workbook;
    const style = record === undefined ? workbook.Model.DefaultStyle : record.Style;
    return DisplayText.FormatDisplayValue(
      record === undefined ? null : record.Value,
      style,
      workbook.Model.Settings.DisplayCulture,
      workbook.Model.Settings.DateSystem,
    );
  }

  /** Gets or sets the cell formula. */
  public get Formula(): string {
    const record = this.tryGetRecord();
    if (record === undefined || record.Formula === null || record.Formula.length === 0) {
      return "";
    }

    return `=${record.Formula}`;
  }

  public set Formula(value: string) {
    const record = this.getOrCreateRecord();
    record.Formula = Cell.normalizeFormula(value);
    if (record.Formula.length === 0) {
      if (record.Value === null || record.Value === undefined) {
        record.Kind = CellValueKind.Blank;
      }

      return;
    }

    record.Kind = CellValueKind.Formula;
  }

  /** Gets the current logical cell value type. */
  public get Type(): CellValueType {
    const record = this.tryGetRecord();
    if (record === undefined) {
      return CellValueType.IsNull;
    }

    if (record.Formula !== null && record.Formula.length > 0) {
      return CellValueType.IsUnknown;
    }

    switch (record.Kind) {
      case CellValueKind.String:
        return CellValueType.IsString;
      case CellValueKind.Number:
        return CellValueType.IsNumeric;
      case CellValueKind.Boolean:
        return CellValueType.IsBool;
      case CellValueKind.DateTime:
        return CellValueType.IsDateTime;
      case CellValueKind.Error:
        return CellValueType.IsError;
      case CellValueKind.Formula:
        return CellValueType.IsUnknown;
      default:
        return CellValueType.IsNull;
    }
  }

  /** Returns a rich-text character range within the cell text. */
  public Characters(startIndex: number, length: number): FontSetting {
    if (startIndex < 0) {
      throw new RangeError("startIndex");
    }

    if (length < 0) {
      throw new RangeError("length");
    }

    const text = this.getCurrentTextValue();
    if (startIndex + length > text.length) {
      throw new RangeError("length");
    }

    return this.createFontSetting(startIndex, length);
  }

  /** Returns all rich-text character ranges in the cell text. */
  public GetCharacters(): FontSetting[] {
    const record = this.tryGetRecord();
    if (record === undefined || record.RichTextRuns === null || record.RichTextRuns.length === 0) {
      return [];
    }

    return record.RichTextRuns.map((run) => FontSetting.FromCore(run));
  }

  /** Indicates whether the cell string contains rich-text formatting runs. */
  public IsRichText(): boolean {
    const record = this.tryGetRecord();
    return record !== undefined && record.RichTextRuns !== null && record.RichTextRuns.length > 0;
  }

  /** Sets rich-text formatting runs for the current cell text. */
  public SetCharacters(characters: FontSetting[]): void {
    if (characters === null || characters === undefined) {
      throw new Error("characters");
    }

    const text = this.getCurrentTextValue();
    const record = this.getOrCreateRecord();
    if (characters.length === 0) {
      record.RichTextRuns = null;
      return;
    }

    const normalized = Cell.normalizeFontSettings(characters, text.length);
    record.RichTextRuns = normalized.map((setting) => setting.ToCore());
  }

  public PutValue(value: string): void;
  public PutValue(value: number): void;
  public PutValue(value: boolean): void;
  public PutValue(value: Date): void;
  public PutValue(value: string, isConverted: boolean): void;
  public PutValue(value: string, isConverted: boolean, setStyle: boolean): void;
  public PutValue(value: unknown): void;
  public PutValue(value: unknown, isConverted?: boolean, setStyle?: boolean): void {
    if (typeof value === "string" && isConverted !== undefined) {
      if (setStyle !== undefined) {
        // The setStyle flag is retained for Aspose-compatible signature shape.
      }

      if (!isConverted) {
        this.setScalar(value, CellValueKind.String);
        return;
      }

      this.putConvertedString(value);
      return;
    }

    this.assignValue(value);
  }

  /** Gets a detached copy of the cell style. */
  public GetStyle(checkBorders: boolean = false): Style {
    if (checkBorders) {
      // Retained for Aspose-compatible overload shape.
    }

    const record = this.tryGetRecord();
    const style = record === undefined ? this.worksheet.Workbook.Model.DefaultStyle : record.Style;
    return Style.FromCore(style);
  }

  /** Replaces the cell style with the supplied style object. */
  public SetStyle(style: Style): void;
  public SetStyle(style: Style, explicitFlag: boolean): void;
  public SetStyle(style: Style, flag: StyleFlag): void;
  public SetStyle(style: Style, flagOrExplicit?: boolean | StyleFlag): void {
    if (style === null || style === undefined) {
      throw new Error("style");
    }

    if (flagOrExplicit === undefined || typeof flagOrExplicit === "boolean") {
      this.getOrCreateRecord().Style = style.ToCore();
      return;
    }

    this.applyStyleWithFlags(style, flagOrExplicit);
  }

  public toString(): string {
    return this.StringValue;
  }

  private applyStyleWithFlags(style: Style, flag: StyleFlag): void {
    if (flag.All) {
      this.SetStyle(style);
      return;
    }

    const target = this.GetStyle();

    if (flag.NumberFormat) {
      target.Number = style.Number;
      target.Custom = style.Custom;
    }

    if (flag.CellShading) {
      target.Pattern = style.Pattern;
      target.ForegroundColor = style.ForegroundColor;
      target.BackgroundColor = style.BackgroundColor;
    }

    if (flag.Font) {
      target.Font.Name = style.Font.Name;
      target.Font.Size = style.Font.Size;
      target.Font.IsBold = style.Font.IsBold;
      target.Font.IsItalic = style.Font.IsItalic;
      target.Font.Underline = style.Font.Underline;
      target.Font.IsStrikeout = style.Font.IsStrikeout;
      target.Font.Color = style.Font.Color;
    } else {
      if (flag.FontName) {
        target.Font.Name = style.Font.Name;
      }

      if (flag.FontSize) {
        target.Font.Size = style.Font.Size;
      }

      if (flag.FontBold) {
        target.Font.IsBold = style.Font.IsBold;
      }

      if (flag.FontItalic) {
        target.Font.IsItalic = style.Font.IsItalic;
      }

      if (flag.FontUnderline) {
        target.Font.Underline = style.Font.Underline;
      }

      if (flag.FontStrike) {
        target.Font.IsStrikeout = style.Font.IsStrikeout;
      }

      if (flag.FontColor) {
        target.Font.Color = style.Font.Color;
      }
    }

    if (flag.Borders) {
      target.Borders.Left = style.Borders.Left.clone();
      target.Borders.Right = style.Borders.Right.clone();
      target.Borders.Top = style.Borders.Top.clone();
      target.Borders.Bottom = style.Borders.Bottom.clone();
      target.Borders.Diagonal = style.Borders.Diagonal.clone();
      target.Borders.DiagonalUp = style.Borders.DiagonalUp;
      target.Borders.DiagonalDown = style.Borders.DiagonalDown;
    } else {
      if (flag.LeftBorder) {
        target.Borders.Left = style.Borders.Left.clone();
      }

      if (flag.RightBorder) {
        target.Borders.Right = style.Borders.Right.clone();
      }

      if (flag.TopBorder) {
        target.Borders.Top = style.Borders.Top.clone();
      }

      if (flag.BottomBorder) {
        target.Borders.Bottom = style.Borders.Bottom.clone();
      }

      if (flag.DiagonalUpBorder) {
        target.Borders.Diagonal = style.Borders.Diagonal.clone();
        target.Borders.DiagonalUp = style.Borders.DiagonalUp;
      }

      if (flag.DiagonalDownBorder) {
        target.Borders.Diagonal = style.Borders.Diagonal.clone();
        target.Borders.DiagonalDown = style.Borders.DiagonalDown;
      }
    }

    if (flag.Alignments) {
      target.HorizontalAlignment = style.HorizontalAlignment;
      target.VerticalAlignment = style.VerticalAlignment;
      target.WrapText = style.WrapText;
      target.IndentLevel = style.IndentLevel;
      target.TextRotation = style.TextRotation;
      target.ReadingOrder = style.ReadingOrder;
      target.RelativeIndent = style.RelativeIndent;
      target.ShrinkToFit = style.ShrinkToFit;
    } else {
      if (flag.HorizontalAlignment) {
        target.HorizontalAlignment = style.HorizontalAlignment;
      }

      if (flag.Indent) {
        target.IndentLevel = style.IndentLevel;
      }

      if (flag.VerticalAlignment) {
        target.VerticalAlignment = style.VerticalAlignment;
      }

      if (flag.WrapText) {
        target.WrapText = style.WrapText;
      }

      if (flag.Rotation) {
        target.TextRotation = style.TextRotation;
      }

      if (flag.TextDirection) {
        target.ReadingOrder = style.ReadingOrder;
      }

      if (flag.ShrinkToFit) {
        target.ShrinkToFit = style.ShrinkToFit;
      }
    }

    if (flag.Locked) {
      target.IsLocked = style.IsLocked;
    }

    if (flag.Hidden) {
      target.IsHidden = style.IsHidden;
    }

    if (flag.QuotePrefix) {
      target.QuotePrefix = style.QuotePrefix;
    }

    this.SetStyle(target);
  }

  private putConvertedString(value: string): void {
    if (/^[+-]?\d+$/.test(value.trim())) {
      this.setScalar(Number.parseInt(value.trim(), 10), CellValueKind.Number);
      return;
    }

    if (/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(value.trim())) {
      this.setScalar(Number.parseFloat(value.trim()), CellValueKind.Number);
      return;
    }

    const lower = value.trim().toLowerCase();
    if (lower === "true" || lower === "false") {
      this.setScalar(lower === "true", CellValueKind.Boolean);
      return;
    }

    const parsedDate = new Date(value);
    if (!Number.isNaN(parsedDate.getTime())) {
      this.setScalar(parsedDate, CellValueKind.DateTime);
      return;
    }

    this.setScalar(value, CellValueKind.String);
  }

  private assignValue(value: unknown): void {
    if (value === null || value === undefined) {
      this.clearValue();
      return;
    }

    if (typeof value === "string") {
      this.setScalar(value, CellValueKind.String);
      return;
    }

    if (typeof value === "boolean") {
      this.setScalar(value, CellValueKind.Boolean);
      return;
    }

    if (value instanceof Date) {
      this.setScalar(value, CellValueKind.DateTime);
      return;
    }

    if (typeof value === "number" || typeof value === "bigint") {
      this.setScalar(Number(value), CellValueKind.Number);
      return;
    }

    this.setScalar(String(value), CellValueKind.String);
  }

  private clearValue(): void {
    const record = this.getOrCreateRecord();
    record.Value = null;
    record.Formula = null;
    record.Kind = CellValueKind.Blank;
    record.RichTextRuns = null;
  }

  private tryGetRecord() {
    return findCellRecord(this.worksheet.Model, this.address.RowIndex, this.address.ColumnIndex);
  }

  private getOrCreateRecord() {
    return getOrCreateCellRecord(
      this.worksheet.Model,
      this.address.RowIndex,
      this.address.ColumnIndex,
      this.worksheet.Workbook.Model.DefaultStyle,
    );
  }

  private setScalar(value: unknown, kind: CellValueKind): void {
    const record = this.getOrCreateRecord();
    record.Value = value;
    record.Kind = kind;
    record.Formula = null;
    record.RichTextRuns = null;
  }

  private getCurrentTextValue(): string {
    const record = this.tryGetRecord();
    if (record === undefined || record.Value === null || record.Value === undefined) {
      return "";
    }

    if (typeof record.Value !== "string") {
      throw new CellsException("Rich text can only be applied to string cells.");
    }

    return record.Value;
  }

  private createFontSetting(startIndex: number, length: number): FontSetting {
    const setting = new FontSetting(startIndex, length);
    const record = this.tryGetRecord();
    const run = Cell.findRichTextRun(record, startIndex, length);
    if (run !== null) {
      return FontSetting.FromCore(run);
    }

    const style = record === undefined
      ? this.worksheet.Workbook.Model.DefaultStyle
      : record.Style;
    setting.Font.Name = style.Font.Name;
    setting.Font.Size = style.Font.Size;
    setting.Font.IsBold = style.Font.Bold;
    setting.Font.IsItalic = style.Font.Italic;
    setting.Font.Underline = style.Font.Underline as unknown as FontUnderlineType;
    setting.Font.IsStrikeout = style.Font.StrikeThrough;
    setting.Font.Color = Color.FromCore(style.Font.Color);
    return setting;
  }

  private static findRichTextRun(
    record: CellRecord | undefined,
    startIndex: number,
    length: number,
  ): RichTextRunValue | null {
    if (record === undefined || record.RichTextRuns === null) {
      return null;
    }

    for (const run of record.RichTextRuns) {
      if (run.StartIndex === startIndex && run.Length === length) {
        return run;
      }
    }

    return null;
  }

  private static normalizeFontSettings(characters: FontSetting[], textLength: number): FontSetting[] {
    const normalized: FontSetting[] = [];
    for (const setting of characters) {
      if (setting === null || setting === undefined) {
        throw new Error("characters");
      }

      if (setting.StartIndex < 0 || setting.Length < 0 || setting.StartIndex + setting.Length > textLength) {
        throw new RangeError("characters");
      }

      normalized.push(setting.clone());
    }

    normalized.sort((left, right) => {
      if (left.StartIndex !== right.StartIndex) {
        return left.StartIndex - right.StartIndex;
      }

      return left.Length - right.Length;
    });

    let previousEnd = -1;
    for (let index = 0; index < normalized.length; index++) {
      const setting = normalized[index];
      if (index > 0 && setting.StartIndex < previousEnd) {
        throw new CellsException("Rich text runs must not overlap.");
      }

      previousEnd = setting.StartIndex + setting.Length;
    }

    return normalized;
  }

  private static normalizeFormula(value: string | null | undefined): string {
    if (value === null || value === undefined || value.trim().length === 0) {
      return "";
    }

    return value[0] === "=" ? value.substring(1) : value;
  }
}
