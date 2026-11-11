import type { Hyperlink as HyperlinkType } from "./types"
import { cellRef, escapeXml, generateUuid } from "./util"

export class LegacyHyperlink {
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

export class LegacyHyperlinkCollection {
  private _hyperlinks: LegacyHyperlink[] = []

  add(row: number, col: number, address: string, display?: string, tooltip?: string): LegacyHyperlink {
    const hyperlink = new LegacyHyperlink(row, col, address, display, tooltip)
    this._hyperlinks.push(hyperlink)
    return hyperlink
  }

  get(index: number): LegacyHyperlink | undefined {
    return this._hyperlinks[index]
  }

  getByCell(row: number, col: number): LegacyHyperlink | undefined {
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

import { TargetModeType } from "./enums";
import { HyperlinkModel } from "./Core/HyperlinkModel";

// Implemented by serializer module.
/** Represents a worksheet hyperlink. */
export class Hyperlink {
  private readonly onDelete: ((model: HyperlinkModel) => void) | undefined;

  public constructor(
    private readonly model: HyperlinkModel,
    onDelete?: (model: HyperlinkModel) => void,
    private readonly owner?: HyperlinkModel[],
  ) {
    this.onDelete = onDelete;
  }

  public get Model(): HyperlinkModel {
    return this.model;
  }

  /** Mirrors Hyperlink.BelongsTo: the facade's backing collection identity. */
  public BelongsTo(owner: HyperlinkModel[]): boolean {
    return this.owner === owner;
  }

  public get Address(): string {
    if (this.model.Address !== null && this.model.Address !== undefined && this.model.Address.length > 0) {
      return this.model.Address;
    }

    return this.model.SubAddress ?? "";
  }

  public set Address(value: string) {
    Hyperlink.assignAddress(this.model, value);
  }

  public get SubAddress(): string {
    return this.model.SubAddress ?? "";
  }

  public set SubAddress(value: string) {
    this.model.SubAddress = value ?? "";
  }

  public get ScreenTip(): string {
    return this.model.ScreenTip ?? "";
  }

  public set ScreenTip(value: string) {
    this.model.ScreenTip = Hyperlink.normalizeText(value);
  }

  public get TextToDisplay(): string {
    if (this.model.TextToDisplay !== null && this.model.TextToDisplay !== undefined && this.model.TextToDisplay.length > 0) {
      return this.model.TextToDisplay;
    }

    return this.Address;
  }

  public set TextToDisplay(value: string) {
    this.model.TextToDisplay = Hyperlink.normalizeText(value);
  }

  public get LinkType(): TargetModeType {
    if (this.model.SubAddress !== null && this.model.SubAddress !== undefined && this.model.SubAddress.length > 0) {
      return TargetModeType.CellReference;
    }

    const address = this.model.Address;
    if (address === null || address === undefined || address.length === 0) {
      return TargetModeType.External;
    }

    if (address.toLowerCase().startsWith("mailto:")) {
      return TargetModeType.Email;
    }

    if (address.startsWith("\\") || address.indexOf(":\\") > 0) {
      return TargetModeType.FilePath;
    }

    return TargetModeType.External;
  }

  public get Area(): string {
    const first = `${Hyperlink.columnLetter(this.model.FirstColumn)}${this.model.FirstRow + 1}`;
    if (this.model.TotalRows === 1 && this.model.TotalColumns === 1) {
      return first;
    }

    const last = `${Hyperlink.columnLetter(this.model.FirstColumn + this.model.TotalColumns - 1)}${this.model.FirstRow + this.model.TotalRows}`;
    return `${first}:${last}`;
  }

  public Delete(): void {
    if (this.onDelete !== undefined) {
      this.onDelete(this.model);
    }
  }

  private static normalizeText(value: string | null | undefined): string {
    return value === null || value === undefined || value.length === 0 ? "" : value;
  }

  private static assignAddress(model: HyperlinkModel, value: string | null | undefined): void {
    if (value === null || value === undefined || value.trim().length === 0) {
      model.Address = null;
      model.SubAddress = null;
      return;
    }

    const normalized = value.trim();
    if (normalized.startsWith("#")) {
      model.Address = null;
      model.SubAddress = normalized.substring(1);
      return;
    }

    if (normalized.indexOf("!") >= 0) {
      model.Address = null;
      model.SubAddress = normalized;
      return;
    }

    model.Address = normalized;
    model.SubAddress = null;
  }

  private static columnLetter(columnIndex: number): string {
    let value = columnIndex + 1;
    let letters = "";
    while (value > 0) {
      const remainder = (value - 1) % 26;
      letters = String.fromCharCode(65 + remainder) + letters;
      value = Math.floor((value - 1) / 26);
    }

    return letters;
  }
}
