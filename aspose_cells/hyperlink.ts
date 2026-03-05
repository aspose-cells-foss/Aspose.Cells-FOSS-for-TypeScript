import type { Hyperlink as HyperlinkType } from "./types"
import { cellRef, escapeXml, generateUuid } from "./util"

export class Hyperlink {
  private _row: number
  private _col: number
  private _address?: string
  private _display?: string
  private _tooltip?: string
  private _type: HyperlinkType["type"] = "url"
  private _target?: string

  constructor(row: number, col: number, address: string, display?: string, tooltip?: string) {
    this._row = row
    this._col = col
    this._address = address

    if (address.startsWith("http://") || address.startsWith("https://") || address.startsWith("mailto:")) {
      this._type = "url"
    } else if (address.startsWith("#")) {
      this._type = "internal"
      this._target = address.substring(1)
    } else {
      this._type = "file"
    }

    this._display = display ?? address
    this._tooltip = tooltip
  }

  get row(): number {
    return this._row
  }

  get col(): number {
    return this._col
  }

  get ref(): string {
    return cellRef(this._row, this._col)
  }

  get address(): string | undefined {
    return this._address
  }

  set address(value: string) {
    this._address = value
  }

  get display(): string | undefined {
    return this._display
  }

  set display(value: string) {
    this._display = value
  }

  get tooltip(): string | undefined {
    return this._tooltip
  }

  set tooltip(value: string) {
    this._tooltip = value
  }

  get type(): HyperlinkType["type"] {
    return this._type
  }

  get target(): string | undefined {
    return this._target
  }

  toXml(): string {
    const rId = generateUuid()
    return `<hyperlink r:id="r${rId}" ref="${this.ref}" ${this.address ? `address="${escapeXml(this.address)}"` : ""} ${this.display ? `display="${escapeXml(this.display)}"` : ""} ${this.tooltip ? `tooltip="${escapeXml(this.tooltip)}"` : ""}/>`
  }
}

export class HyperlinkCollection {
  private _hyperlinks: Hyperlink[] = []

  add(row: number, col: number, address: string, display?: string, tooltip?: string): Hyperlink {
    const hyperlink = new Hyperlink(row, col, address, display, tooltip)
    this._hyperlinks.push(hyperlink)
    return hyperlink
  }

  get(index: number): Hyperlink | undefined {
    return this._hyperlinks[index]
  }

  getByCell(row: number, col: number): Hyperlink | undefined {
    return this._hyperlinks.find((h) => h.row === row && h.col === col)
  }

  remove(index: number) {
    this._hyperlinks.splice(index, 1)
  }

  clear() {
    this._hyperlinks = []
  }

  get count(): number {
    return this._hyperlinks.length
  }

  toXml(): string {
    return this._hyperlinks.map((h) => h.toXml()).join("\n")
  }
}
