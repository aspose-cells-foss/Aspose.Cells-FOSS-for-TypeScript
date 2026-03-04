import type { ConditionalFormatRule, ColorScaleRule, IconSetRule, DataBarRule, CellRange } from "./types"
import { parseRange, cellRef, escapeXml } from "./util"

export class ConditionalFormat {
  private _ranges: CellRange[] = []
  private _rules: ConditionalFormatRule[] = []

  addArea(range: string) {
    this._ranges.push(parseRange(range))
  }

  getRanges(): CellRange[] {
    return this._ranges
  }

  addRule(rule: ConditionalFormatRule) {
    this._rules.push(rule)
  }

  getRules(): ConditionalFormatRule[] {
    return this._rules
  }

  toXml(): string {
    const sqref = this._ranges.map((r) => `${cellRef(r.startRow, r.startCol)}:${cellRef(r.endRow, r.endCol)}`).join(" ")

    let xml = `<conditionalFormatting sqref="${sqref}">`

    for (const rule of this._rules) {
      xml += "<cfRule"
      if (rule.type) xml += ` type="${rule.type}"`
      if (rule.operator) xml += ` operator="${rule.operator}"`
      if (rule.formula1) xml += ` dxfId="0"`
      xml += ">"

      if (rule.formula1) {
        xml += `<formula>${escapeXml(rule.formula1)}</formula>`
      }
      if (rule.formula2) {
        xml += `<formula>${escapeXml(rule.formula2)}</formula>`
      }

      xml += "</cfRule>"
    }

    xml += "</conditionalFormatting>"
    return xml
  }
}

export class ConditionalFormatCollection {
  private _formats: ConditionalFormat[] = []

  add(): ConditionalFormat {
    const format = new ConditionalFormat()
    this._formats.push(format)
    return format
  }

  get(index: number): ConditionalFormat | undefined {
    return this._formats[index]
  }

  remove(index: number) {
    this._formats.splice(index, 1)
  }

  clear() {
    this._formats = []
  }

  get count(): number {
    return this._formats.length
  }

  toXml(): string {
    return this._formats.map((f) => f.toXml()).join("\n")
  }
}
