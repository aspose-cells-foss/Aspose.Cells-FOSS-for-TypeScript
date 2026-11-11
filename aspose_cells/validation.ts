import type { DataValidation as DataValidationType, CellRange } from "./types";
import { parseRange, cellRef, escapeXml } from "./util";

export class LegacyValidation {
  private _type: DataValidationType["type"] = "any";
  private _operator?: DataValidationType["operator"];
  private _formula1?: string;
  private _formula2?: string;
  private _allowBlank = false;
  private _showDropDown = false;
  private _showInputMessage = false;
  private _inputTitle?: string;
  private _inputMessage?: string;
  private _errorTitle?: string;
  private _errorMessage?: string;
  private _errorStyle?: DataValidationType["errorStyle"];
  private _ranges: CellRange[] = [];

  get type(): DataValidationType["type"] {
    return this._type;
  }

  set type(value: DataValidationType["type"]) {
    this._type = value;
  }

  get operator(): DataValidationType["operator"] | undefined {
    return this._operator;
  }

  set operator(value: DataValidationType["operator"]) {
    this._operator = value;
  }

  get formula1(): string | undefined {
    return this._formula1;
  }

  set formula1(value: string) {
    this._formula1 = value;
  }

  get formula2(): string | undefined {
    return this._formula2;
  }

  set formula2(value: string) {
    this._formula2 = value;
  }

  get allowBlank(): boolean {
    return this._allowBlank;
  }

  set allowBlank(value: boolean) {
    this._allowBlank = value;
  }

  get showDropDown(): boolean {
    return this._showDropDown;
  }

  set showDropDown(value: boolean) {
    this._showDropDown = value;
  }

  get showInputMessage(): boolean {
    return this._showInputMessage;
  }

  set showInputMessage(value: boolean) {
    this._showInputMessage = value;
  }

  get inputTitle(): string | undefined {
    return this._inputTitle;
  }

  set inputTitle(value: string) {
    this._inputTitle = value;
  }

  get inputMessage(): string | undefined {
    return this._inputMessage;
  }

  set inputMessage(value: string) {
    this._inputMessage = value;
  }

  get errorTitle(): string | undefined {
    return this._errorTitle;
  }

  set errorTitle(value: string) {
    this._errorTitle = value;
  }

  get errorMessage(): string | undefined {
    return this._errorMessage;
  }

  set errorMessage(value: string) {
    this._errorMessage = value;
  }

  get errorStyle(): DataValidationType["errorStyle"] | undefined {
    return this._errorStyle;
  }

  set errorStyle(value: DataValidationType["errorStyle"]) {
    this._errorStyle = value;
  }

  addArea(range: string) {
    this._ranges.push(parseRange(range));
  }

  getRanges(): CellRange[] {
    return this._ranges;
  }

  toXml(): string {
    let typeStr: string = this._type;
    if (this._operator && this._type !== "list") {
      typeStr = `${this._type}:${this._operator}`;
    }

    let xml = `<dataValidation type="${typeStr}"`;
    if (this._allowBlank) xml += ` allowBlank="1"`;
    if (this._showDropDown) xml += ` showDropDown="1"`;
    if (this._showInputMessage) xml += ` showInputMessage="1"`;
    if (this._inputTitle) xml += ` inputTitle="${escapeXml(this._inputTitle)}"`;
    if (this._inputMessage)
      xml += ` inputMessage="${escapeXml(this._inputMessage)}"`;
    if (this._errorTitle) xml += ` errorTitle="${escapeXml(this._errorTitle)}"`;
    if (this._errorMessage)
      xml += ` errorMessage="${escapeXml(this._errorMessage)}"`;
    if (this._errorStyle) xml += ` errorStyle="${this._errorStyle}"`;

    for (const range of this._ranges) {
      xml += ` sqref="${cellRef(range.startRow, range.startCol)}:${cellRef(range.endRow, range.endCol)}"`;
    }

    xml += ">";

    if (this._formula1) {
      xml += `<formula1>${escapeXml(this._formula1)}</formula1>`;
    }
    if (this._formula2) {
      xml += `<formula2>${escapeXml(this._formula2)}</formula2>`;
    }

    xml += "</dataValidation>";
    return xml;
  }
}

export class LegacyValidationCollection {
  private _validations: LegacyValidation[] = [];

  add(): LegacyValidation {
    const validation = new LegacyValidation();
    this._validations.push(validation);
    return validation;
  }

  get(index: number): LegacyValidation | undefined {
    return this._validations[index];
  }

  get count(): number {
    return this._validations.length;
  }

  remove(index: number) {
    this._validations.splice(index, 1);
  }

  clear() {
    this._validations = [];
  }

  toXml(): string {
    return this._validations.map((v) => v.toXml()).join("\n");
  }
}

import { CellArea } from "./CellArea";
import { toApiCellArea } from "./CellAreaConvert";
import { ValidationAlertType, ValidationType, OperatorType } from "./enums";
import { ValidationModel } from "./Core/ValidationModel";
import { ValidationCollection } from "./ValidationCollection";

// Implemented by serializer module.
/** Represents a data validation rule. */
export class Validation {
  public constructor(
    private readonly owner: ValidationModel[],
    private readonly model: ValidationModel,
  ) {}

  public get Model(): ValidationModel {
    return this.model;
  }

  public get Areas(): readonly CellArea[] {
    return this.model.Areas.map((area) => toApiCellArea(area));
  }

  public get Type(): ValidationType {
    return this.model.Type as unknown as ValidationType;
  }

  public set Type(value: ValidationType) {
    this.model.Type = value as unknown as import("./ValidationType").ValidationType;
  }

  public get AlertStyle(): ValidationAlertType {
    return this.model.AlertStyle as unknown as ValidationAlertType;
  }

  public set AlertStyle(value: ValidationAlertType) {
    this.model.AlertStyle = value as unknown as import("./ValidationAlertType").ValidationAlertType;
  }

  public get Operator(): OperatorType {
    return this.model.Operator as unknown as OperatorType;
  }

  public set Operator(value: OperatorType) {
    this.model.Operator = value as unknown as import("./OperatorType").OperatorType;
  }

  public get Formula1(): string {
    return this.model.Formula1 ?? "";
  }

  public set Formula1(value: string) {
    this.model.Formula1 = Validation.normalizeFormula(value);
  }

  public get Formula2(): string {
    return this.model.Formula2 ?? "";
  }

  public set Formula2(value: string) {
    this.model.Formula2 = Validation.normalizeFormula(value);
  }

  public get IgnoreBlank(): boolean {
    return this.model.IgnoreBlank;
  }

  public set IgnoreBlank(value: boolean) {
    this.model.IgnoreBlank = value;
  }

  public get InCellDropDown(): boolean {
    return this.model.InCellDropDown;
  }

  public set InCellDropDown(value: boolean) {
    this.model.InCellDropDown = value;
  }

  public get InputTitle(): string {
    return this.model.InputTitle ?? "";
  }

  public set InputTitle(value: string) {
    this.model.InputTitle = Validation.normalizeText(value);
  }

  public get InputMessage(): string {
    return this.model.InputMessage ?? "";
  }

  public set InputMessage(value: string) {
    this.model.InputMessage = Validation.normalizeText(value);
  }

  public get ErrorTitle(): string {
    return this.model.ErrorTitle ?? "";
  }

  public set ErrorTitle(value: string) {
    this.model.ErrorTitle = Validation.normalizeText(value);
  }

  public get ErrorMessage(): string {
    return this.model.ErrorMessage ?? "";
  }

  public set ErrorMessage(value: string) {
    this.model.ErrorMessage = Validation.normalizeText(value);
  }

  public get ShowInput(): boolean {
    return this.model.ShowInput;
  }

  public set ShowInput(value: boolean) {
    this.model.ShowInput = value;
  }

  public get ShowError(): boolean {
    return this.model.ShowError;
  }

  public set ShowError(value: boolean) {
    this.model.ShowError = value;
  }

  public AddArea(area: CellArea): void {
    ValidationCollection.AddAreaToValidation(this.owner, this.model, area);
  }

  public RemoveArea(area: CellArea): void {
    ValidationCollection.RemoveAreaFromValidation(this.owner, this.model, area);
  }

  private static normalizeFormula(value: string | null | undefined): string | null {
    if (value === null || value === undefined) {
      return null;
    }

    const trimmed = value.trim();
    if (trimmed.length === 0) {
      return null;
    }

    return trimmed.startsWith("=") ? trimmed.substring(1) : trimmed;
  }

  private static normalizeText(value: string | null | undefined): string | null {
    return value === null || value === undefined || value.length === 0 ? null : value;
  }
}
