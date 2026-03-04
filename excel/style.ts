import type { Style as StyleType, Font, Fill, Border, Alignment, Protection } from "./types"

export class Style implements StyleType {
  font?: Font
  fill?: Fill
  border?: Border
  alignment?: Alignment
  numberFormat?: string
  protection?: Protection

  constructor() {}

  getFont(): Font {
    if (!this.font) {
      this.font = {}
    }
    return this.font
  }

  setFont(font: Font) {
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

  getBorder(): Border {
    if (!this.border) {
      this.border = {}
    }
    return this.border
  }

  setBorder(border: Border) {
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
