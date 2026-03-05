import type { FilterColumn } from "./types"
import { cellRef, escapeXml } from "./util"

export class AutoFilter {
  private _range: string
  private _columns: FilterColumn[] = []

  constructor(range: string) {
    this._range = range
  }

  get range(): string {
    return this._range
  }

  set range(value: string) {
    this._range = value
  }

  get columns(): FilterColumn[] {
    return this._columns
  }

  addFilterColumn(col: number, filters: (string | number)[], blank?: boolean) {
    this._columns.push({ col, filters, blank })
  }

  removeFilterColumn(col: number) {
    const idx = this._columns.findIndex((c) => c.col === col)
    if (idx >= 0) {
      this._columns.splice(idx, 1)
    }
  }

  clear() {
    this._columns = []
  }

  toXml(): string {
    let xml = `<autoFilter ref="${this._range}">`
    for (const col of this._columns) {
      xml += `<filterColumn colId="${col.col}">`
      if (col.blank) {
        xml += "<blank/>"
      }
      for (const filter of col.filters) {
        xml += `<filters><filter val="${escapeXml(String(filter))}"/></filters>`
      }
      xml += "</filterColumn>"
    }
    xml += "</autoFilter>"
    return xml
  }
}
