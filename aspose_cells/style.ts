import type { Style as StyleType, Font as FontType, Fill, Border as BorderType, Alignment, Protection } from "./types"

export class LegacyStyle implements StyleType {
  font?: FontType
  fill?: Fill
  border?: BorderType
  alignment?: Alignment
  numberFormat?: string
  protection?: Protection

  constructor() {}

  getFont(): FontType {
    if (!this.font) {
      this.font = {}
    }
    return this.font
  }

  setFont(font: FontType) {
    this.font = font
    return this
  }

  getFill(): Fill {
    if (!this.fill) {
      this.fill = { patternType: "none" }
    }
    return this.fill
  }

  setFill(fill: Fill) {
    this.fill = fill
    return this
  }

  getBorder(): BorderType {
    if (!this.border) {
      this.border = {}
    }
    return this.border
  }

  setBorder(border: BorderType) {
    this.border = border
    return this
  }

  getAlignment(): Alignment {
    if (!this.alignment) {
      this.alignment = {}
    }
    return this.alignment
  }

  setAlignment(alignment: Alignment) {
    this.alignment = alignment
    return this
  }

  setNumberFormat(format: string) {
    this.numberFormat = format
    return this
  }

  getNumberFormat(): string | undefined {
    return this.numberFormat
  }

  setProtection(protection: Protection) {
    this.protection = protection
    return this
  }

  isBold(): boolean {
    return this.font?.bold ?? false
  }

  setBold(bold: boolean) {
    this.getFont().bold = bold
    return this
  }

  isItalic(): boolean {
    return this.font?.italic ?? false
  }

  setItalic(italic: boolean) {
    this.getFont().italic = italic
    return this
  }

  getFontName(): string {
    return this.font?.name ?? "Calibri"
  }

  setFontName(name: string) {
    this.getFont().name = name
    return this
  }

  getFontSize(): number {
    return this.font?.size ?? 11
  }

  setFontSize(size: number) {
    this.getFont().size = size
    return this
  }

  getFontColor(): string | undefined {
    return this.font?.color
  }

  setFontColor(color: string) {
    this.getFont().color = color
    return this
  }

  getForegroundColor(): string | undefined {
    return this.fill?.fgColor
  }

  setForegroundColor(color: string) {
    this.getFill().patternType = "solid"
    this.getFill().fgColor = color
    return this
  }

  getBackgroundColor(): string | undefined {
    return this.fill?.bgColor
  }

  setBackgroundColor(color: string) {
    this.getFill().bgColor = color
    return this
  }

  getHorizontalAlignment(): string {
    return this.alignment?.horizontal ?? "general"
  }

  setHorizontalAlignment(alignment: Alignment["horizontal"]) {
    this.getAlignment().horizontal = alignment
    return this
  }

  getVerticalAlignment(): string {
    return this.alignment?.vertical ?? "bottom"
  }

  setVerticalAlignment(alignment: Alignment["vertical"]) {
    this.getAlignment().vertical = alignment
    return this
  }

  isWrapText(): boolean {
    return this.alignment?.wrapText ?? false
  }

  setWrapText(wrapText: boolean) {
    this.getAlignment().wrapText = wrapText
    return this
  }

  isLocked(): boolean {
    return this.protection?.locked ?? true
  }

  setLocked(locked: boolean) {
    this.getProtection().locked = locked
    return this
  }

  isHidden(): boolean {
    return this.protection?.hidden ?? false
  }

  setHidden(hidden: boolean) {
    this.getProtection().hidden = hidden
    return this
  }

  private getProtection(): Protection {
    if (!this.protection) {
      this.protection = {}
    }
    return this.protection
  }

  toXml(): string {
    let xml = ""

    if (this.font) {
      xml += "<font>"
      if (this.font.bold) xml += "<b/>"
      if (this.font.italic) xml += "<i/>"
      if (this.font.underline) xml += "<u/>"
      if (this.font.strikeout) xml += "<strike/>"
      if (this.font.size) xml += `<sz val="${this.font.size}"/>`
      if (this.font.name) xml += `<name val="${this.font.name}"/>`
      if (this.font.color) xml += `<color rgb="${this.font.color}"/>`
      xml += "</font>"
    }

    if (this.fill) {
      xml += `<fill><patternFill patternType="${this.fill.patternType || "none"}">`
      if (this.fill.fgColor) xml += `<fgColor rgb="${this.fill.fgColor}"/>`
      if (this.fill.bgColor) xml += `<bgColor rgb="${this.fill.bgColor}"/>`
      xml += "</patternFill></fill>"
    }

    if (this.border) {
      xml += "<border>"
      if (this.border.left)
        xml += `<left style="${this.border.left.style || "thin"}"><color rgb="${this.border.left.color || "000000"}"/></left>`
      else xml += "<left/>"
      if (this.border.right)
        xml += `<right style="${this.border.right.style || "thin"}"><color rgb="${this.border.right.color || "000000"}"/></right>`
      else xml += "<right/>"
      if (this.border.top)
        xml += `<top style="${this.border.top.style || "thin"}"><color rgb="${this.border.top.color || "000000"}"/></top>`
      else xml += "<top/>"
      if (this.border.bottom)
        xml += `<bottom style="${this.border.bottom.style || "thin"}"><color rgb="${this.border.bottom.color || "000000"}"/></bottom>`
      else xml += "<bottom/>"
      xml += "</border>"
    }

    if (this.alignment) {
      xml += `<alignment`
      if (this.alignment.horizontal) xml += ` horizontal="${this.alignment.horizontal}"`
      if (this.alignment.vertical) xml += ` vertical="${this.alignment.vertical}"`
      if (this.alignment.wrapText) xml += ` wrapText="1"`
      xml += "/>"
    }

    if (this.numberFormat) {
      xml += `<numFmt formatCode="${this.numberFormat}" count="1"/>`
    }

    return xml
  }
}

import { Border } from "./Border";
import { Borders } from "./Borders";
import { Color } from "./Color";
import { CellsException } from "./CellsException";
import { Font } from "./Font";
import { NumberFormat } from "./NumberFormat";
import { FillPattern, HorizontalAlignmentType, VerticalAlignmentType } from "./enums";
import { AlignmentValue } from "./Core/AlignmentValue";
import { BorderSideValue } from "./Core/BorderSideValue";
import { BordersValue } from "./Core/BordersValue";
import { FillPatternKind } from "./Core/FillPatternKind";
import { FontValue } from "./Core/FontValue";
import { NumberFormatValue } from "./Core/NumberFormatValue";
import { ProtectionValue } from "./Core/ProtectionValue";
import { StyleValue } from "./Core/StyleValue";
import { BorderStyle as CoreBorderStyle } from "./Core/BorderStyle";
import { FontUnderlineType as CoreFontUnderlineType } from "./FontUnderlineType";
import { HorizontalAlignment as CoreHorizontalAlignment } from "./Core/HorizontalAlignment";
import { VerticalAlignment as CoreVerticalAlignment } from "./Core/VerticalAlignment";
import { toCoreColor } from "./ColorConvert";

/** Represents a mutable cell style facade that can be applied to one or more cells. */
export class Style {
  private indentLevel: number = 0;
  private textRotation: number = 0;
  private readingOrder: number = 0;

  public Font: Font = new Font();
  public Borders: Borders = new Borders();
  public Pattern: FillPattern = FillPattern.None;
  public ForegroundColor: Color = Color.Empty;
  public BackgroundColor: Color = Color.Empty;
  public Number: number = 0;
  public Custom: string | null = null;
  public HorizontalAlignment: HorizontalAlignmentType = HorizontalAlignmentType.General;
  public VerticalAlignment: VerticalAlignmentType = VerticalAlignmentType.Bottom;
  public WrapText: boolean = false;
  public ShrinkToFit: boolean = false;
  public RelativeIndent: number = 0;
  public IsLocked: boolean = true;
  public IsHidden: boolean = false;
  public QuotePrefix: boolean = false;

  /** Copies data from another style object. */
  public Copy(source: Style): void {
    if (source === null || source === undefined) {
      throw new Error("source");
    }

    this.Font = source.Font.clone();
    this.Borders = source.Borders.clone();
    this.Pattern = source.Pattern;
    this.ForegroundColor = source.ForegroundColor;
    this.BackgroundColor = source.BackgroundColor;
    this.Number = source.Number;
    this.Custom = source.Custom;
    this.HorizontalAlignment = source.HorizontalAlignment;
    this.VerticalAlignment = source.VerticalAlignment;
    this.WrapText = source.WrapText;
    this.IndentLevel = source.IndentLevel;
    this.TextRotation = source.TextRotation;
    this.ShrinkToFit = source.ShrinkToFit;
    this.ReadingOrder = source.ReadingOrder;
    this.RelativeIndent = source.RelativeIndent;
    this.IsLocked = source.IsLocked;
    this.IsHidden = source.IsHidden;
    this.QuotePrefix = source.QuotePrefix;
  }

  /** Gets or sets the resolved number format string. */
  public get NumberFormat(): string {
    return NumberFormat.ResolveFormatCode(this.Number, this.Custom);
  }

  public set NumberFormat(value: string) {
    const builtInId = NumberFormat.GetBuiltInFormatId(value);
    if (builtInId !== null) {
      this.Number = builtInId;
      this.Custom = null;
      return;
    }

    this.Number = 0;
    this.Custom = value === null || value === undefined || value.trim().length === 0 ? null : value.trim();
  }

  /** Gets or sets the indentation level. */
  public get IndentLevel(): number {
    return this.indentLevel;
  }

  public set IndentLevel(value: number) {
    if (value < 0 || value > 250) {
      throw new CellsException("IndentLevel must be between 0 and 250.");
    }

    this.indentLevel = value;
  }

  /** Gets or sets the text rotation. */
  public get TextRotation(): number {
    return this.textRotation;
  }

  public set TextRotation(value: number) {
    if (value !== 255 && (value < 0 || value > 180)) {
      throw new CellsException("TextRotation must be between 0 and 180, or 255 for vertical text.");
    }

    this.textRotation = value;
  }

  /** Gets or sets the reading order. */
  public get ReadingOrder(): number {
    return this.readingOrder;
  }

  public set ReadingOrder(value: number) {
    if (value < 0 || value > 2) {
      throw new CellsException("ReadingOrder must be 0, 1, or 2.");
    }

    this.readingOrder = value;
  }

  /** Creates a copy of the current style. */
  public clone(): Style {
    const clone = new Style();
    clone.Copy(this);
    return clone;
  }

  /** Determines whether two style instances are equal. */
  public Equals(other: Style | null | undefined): boolean {
    if (other === null || other === undefined) {
      return false;
    }

    return (
      this.Font.Equals(other.Font) &&
      this.bordersEqual(this.Borders, other.Borders) &&
      this.Pattern === other.Pattern &&
      this.ForegroundColor.Equals(other.ForegroundColor) &&
      this.BackgroundColor.Equals(other.BackgroundColor) &&
      this.Number === other.Number &&
      this.Custom === other.Custom &&
      this.HorizontalAlignment === other.HorizontalAlignment &&
      this.VerticalAlignment === other.VerticalAlignment &&
      this.WrapText === other.WrapText &&
      this.IndentLevel === other.IndentLevel &&
      this.TextRotation === other.TextRotation &&
      this.ShrinkToFit === other.ShrinkToFit &&
      this.ReadingOrder === other.ReadingOrder &&
      this.RelativeIndent === other.RelativeIndent &&
      this.IsLocked === other.IsLocked &&
      this.IsHidden === other.IsHidden &&
      this.QuotePrefix === other.QuotePrefix
    );
  }

  /** Serves as a hash function for a style object. */
  public GetHashCode(): number {
    let hash = 17;
    hash = (hash * 31 + this.fontHashCode(this.Font)) | 0;
    hash = (hash * 31 + this.bordersHashCode(this.Borders)) | 0;
    hash = (hash * 31 + this.Pattern) | 0;
    hash = (hash * 31 + this.ForegroundColor.ToArgb()) | 0;
    hash = (hash * 31 + this.BackgroundColor.ToArgb()) | 0;
    hash = (hash * 31 + this.Number) | 0;
    hash = (hash * 31 + (this.Custom === null ? 0 : this.hashString(this.Custom))) | 0;
    hash = (hash * 31 + this.HorizontalAlignment) | 0;
    hash = (hash * 31 + this.VerticalAlignment) | 0;
    hash = (hash * 31 + (this.WrapText ? 1 : 0)) | 0;
    hash = (hash * 31 + this.IndentLevel) | 0;
    hash = (hash * 31 + this.TextRotation) | 0;
    hash = (hash * 31 + (this.ShrinkToFit ? 1 : 0)) | 0;
    hash = (hash * 31 + this.ReadingOrder) | 0;
    hash = (hash * 31 + this.RelativeIndent) | 0;
    hash = (hash * 31 + (this.IsLocked ? 1 : 0)) | 0;
    hash = (hash * 31 + (this.IsHidden ? 1 : 0)) | 0;
    hash = (hash * 31 + (this.QuotePrefix ? 1 : 0)) | 0;
    return hash;
  }

  public ToCore(): StyleValue {
    const value = new StyleValue();
    value.Font = this.Font.ToCore();
    value.Pattern = this.Pattern as unknown as FillPatternKind;
    value.ForegroundColor = toCoreColor(this.ForegroundColor);
    value.BackgroundColor = toCoreColor(this.BackgroundColor);

    const borders = new BordersValue();
    borders.Left = this.toCoreBorder(this.Borders.Left);
    borders.Right = this.toCoreBorder(this.Borders.Right);
    borders.Top = this.toCoreBorder(this.Borders.Top);
    borders.Bottom = this.toCoreBorder(this.Borders.Bottom);
    borders.Diagonal = this.toCoreBorder(this.Borders.Diagonal);
    borders.DiagonalUp = this.Borders.DiagonalUp;
    borders.DiagonalDown = this.Borders.DiagonalDown;
    value.Borders = borders;

    const alignment = new AlignmentValue();
    alignment.Horizontal = this.HorizontalAlignment as unknown as CoreHorizontalAlignment;
    alignment.Vertical = this.VerticalAlignment as unknown as CoreVerticalAlignment;
    alignment.WrapText = this.WrapText;
    alignment.IndentLevel = this.IndentLevel;
    alignment.TextRotation = this.TextRotation;
    alignment.ShrinkToFit = this.ShrinkToFit;
    alignment.ReadingOrder = this.ReadingOrder;
    alignment.RelativeIndent = this.RelativeIndent;
    value.Alignment = alignment;

    const protection = new ProtectionValue();
    protection.IsLocked = this.IsLocked;
    protection.IsHidden = this.IsHidden;
    value.Protection = protection;

    const numberFormat = new NumberFormatValue();
    numberFormat.Number = this.Number;
    numberFormat.Custom = this.Custom;
    value.NumberFormat = numberFormat;

    value.QuotePrefix = this.QuotePrefix;
    return value;
  }

  public static FromCore(value: StyleValue | null | undefined): Style {
    const style = new Style();
    if (value === null || value === undefined) {
      return style;
    }

    style.Font = Font.FromCore(value.Font);
    style.Pattern = value.Pattern as unknown as FillPattern;
    style.ForegroundColor = Color.FromCore(value.ForegroundColor);
    style.BackgroundColor = Color.FromCore(value.BackgroundColor);

    const borders = new Borders();
    borders.Left = Border.FromCore(value.Borders.Left);
    borders.Right = Border.FromCore(value.Borders.Right);
    borders.Top = Border.FromCore(value.Borders.Top);
    borders.Bottom = Border.FromCore(value.Borders.Bottom);
    borders.Diagonal = Border.FromCore(value.Borders.Diagonal);
    borders.DiagonalUp = value.Borders.DiagonalUp;
    borders.DiagonalDown = value.Borders.DiagonalDown;
    style.Borders = borders;

    style.HorizontalAlignment = value.Alignment.Horizontal as unknown as HorizontalAlignmentType;
    style.VerticalAlignment = value.Alignment.Vertical as unknown as VerticalAlignmentType;
    style.WrapText = value.Alignment.WrapText;
    style.IndentLevel = value.Alignment.IndentLevel;
    style.TextRotation = value.Alignment.TextRotation;
    style.ShrinkToFit = value.Alignment.ShrinkToFit;
    style.ReadingOrder = value.Alignment.ReadingOrder;
    style.RelativeIndent = value.Alignment.RelativeIndent;

    style.IsLocked = value.Protection.IsLocked;
    style.IsHidden = value.Protection.IsHidden;
    style.Number = value.NumberFormat.Number;
    style.Custom = value.NumberFormat.Custom;
    style.QuotePrefix = value.QuotePrefix;
    return style;
  }

  private toCoreBorder(border: Border): BorderSideValue {
    const value = new BorderSideValue();
    value.Style = border.LineStyle as unknown as CoreBorderStyle;
    value.Color = toCoreColor(border.Color);
    return value;
  }

  private bordersEqual(left: Borders, right: Borders): boolean {
    return (
      this.borderEqual(left.Left, right.Left) &&
      this.borderEqual(left.Right, right.Right) &&
      this.borderEqual(left.Top, right.Top) &&
      this.borderEqual(left.Bottom, right.Bottom) &&
      this.borderEqual(left.Diagonal, right.Diagonal) &&
      left.DiagonalUp === right.DiagonalUp &&
      left.DiagonalDown === right.DiagonalDown
    );
  }

  private borderEqual(left: Border, right: Border): boolean {
    return left.LineStyle === right.LineStyle && left.Color.Equals(right.Color);
  }

  private fontHashCode(font: Font): number {
    let hash = 17;
    hash = (hash * 31 + this.hashString(font.Name)) | 0;
    hash = (hash * 31 + font.Size) | 0;
    hash = (hash * 31 + (font.IsBold ? 1 : 0)) | 0;
    hash = (hash * 31 + (font.IsItalic ? 1 : 0)) | 0;
    hash = (hash * 31 + font.Underline) | 0;
    hash = (hash * 31 + (font.IsStrikeout ? 1 : 0)) | 0;
    hash = (hash * 31 + font.Color.ToArgb()) | 0;
    return hash;
  }

  private bordersHashCode(borders: Borders): number {
    let hash = 17;
    hash = (hash * 31 + borders.Left.LineStyle) | 0;
    hash = (hash * 31 + borders.Left.Color.ToArgb()) | 0;
    hash = (hash * 31 + borders.Right.LineStyle) | 0;
    hash = (hash * 31 + borders.Right.Color.ToArgb()) | 0;
    hash = (hash * 31 + borders.Top.LineStyle) | 0;
    hash = (hash * 31 + borders.Top.Color.ToArgb()) | 0;
    hash = (hash * 31 + borders.Bottom.LineStyle) | 0;
    hash = (hash * 31 + borders.Bottom.Color.ToArgb()) | 0;
    hash = (hash * 31 + borders.Diagonal.LineStyle) | 0;
    hash = (hash * 31 + borders.Diagonal.Color.ToArgb()) | 0;
    hash = (hash * 31 + (borders.DiagonalUp ? 1 : 0)) | 0;
    hash = (hash * 31 + (borders.DiagonalDown ? 1 : 0)) | 0;
    return hash;
  }

  private hashString(text: string | null): number {
    if (text === null) {
      return 0;
    }

    let hash = 0;
    for (let index = 0; index < text.length; index++) {
      hash = (Math.imul(hash, 31) + text.charCodeAt(index)) | 0;
    }

    return hash;
  }
}
