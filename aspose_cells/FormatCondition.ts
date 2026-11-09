import { Color } from "./Color";
import { CellsException } from "./CellsException";
import { Style } from "./style";
import { FormatConditionCollection } from "./FormatConditionCollection";
import { ConditionalFormattingModel } from "./Core/ConditionalFormattingModel";
import { FormatConditionModel } from "./Core/FormatConditionModel";
import { StyleValue } from "./Core/StyleValue";
import { ColorValue } from "./Core/ColorValue";

function toCoreColor(value: Color): ColorValue {
  return new ColorValue(value.A, value.R, value.G, value.B);
}

/** Represents a format condition. Mirrors Aspose.Cells_FOSS.FormatCondition. */
export class FormatCondition {
  private readonly owner: ConditionalFormattingModel[];
  private readonly collection: ConditionalFormattingModel;
  private readonly model: FormatConditionModel;

  public constructor(
    owner: ConditionalFormattingModel[],
    collection: ConditionalFormattingModel,
    model: FormatConditionModel,
  ) {
    this.owner = owner;
    this.collection = collection;
    this.model = model;
  }

  public get Model(): FormatConditionModel {
    return this.model;
  }

  public get Type() {
    return this.model.Type;
  }

  public set Type(value) {
    this.model.Type = value;
  }

  public get Operator() {
    return this.model.Operator;
  }

  public set Operator(value) {
    this.model.Operator = value;
  }

  public get Formula1(): string {
    return this.model.Formula1 ?? "";
  }

  public set Formula1(value: string) {
    this.model.Formula1 = normalizeFormula(value);
  }

  public get Formula2(): string {
    return this.model.Formula2 ?? "";
  }

  public set Formula2(value: string) {
    this.model.Formula2 = normalizeFormula(value);
  }

  public get Formula(): string {
    return this.model.Formula1 ?? "";
  }

  public set Formula(value: string) {
    this.model.Formula1 = normalizeFormula(value);
  }

  public get TimePeriod(): string {
    return this.model.TimePeriod ?? "";
  }

  public set TimePeriod(value: string) {
    this.model.TimePeriod = normalizeText(value);
  }

  public get Duplicate(): boolean {
    return this.model.Duplicate;
  }

  public set Duplicate(value: boolean) {
    this.model.Duplicate = value;
  }

  public get Top(): boolean {
    return this.model.Top;
  }

  public set Top(value: boolean) {
    this.model.Top = value;
  }

  public get Percent(): boolean {
    return this.model.Percent;
  }

  public set Percent(value: boolean) {
    this.model.Percent = value;
  }

  public get Rank(): number {
    return this.model.Rank;
  }

  public set Rank(value: number) {
    if (value < 0) {
      throw new CellsException("Conditional formatting rank must be zero or greater.");
    }

    this.model.Rank = value;
  }

  public get Above(): boolean {
    return this.model.Above;
  }

  public set Above(value: boolean) {
    this.model.Above = value;
  }

  public get StandardDeviation(): number {
    return this.model.StandardDeviation;
  }

  public set StandardDeviation(value: number) {
    if (value < 0) {
      throw new CellsException("Conditional formatting standard deviation must be zero or greater.");
    }

    this.model.StandardDeviation = value;
  }

  public get ColorScaleCount(): number {
    return this.model.ColorScaleCount;
  }

  public set ColorScaleCount(value: number) {
    if (value !== 2 && value !== 3) {
      throw new CellsException("ColorScaleCount must be 2 or 3.");
    }

    this.model.ColorScaleCount = value;
  }

  public get MinColor(): Color {
    return Color.FromCore(this.model.MinColor);
  }

  public set MinColor(value: Color) {
    this.model.MinColor = toCoreColor(value);
  }

  public get MidColor(): Color {
    return Color.FromCore(this.model.MidColor);
  }

  public set MidColor(value: Color) {
    this.model.MidColor = toCoreColor(value);
  }

  public get MaxColor(): Color {
    return Color.FromCore(this.model.MaxColor);
  }

  public set MaxColor(value: Color) {
    this.model.MaxColor = toCoreColor(value);
  }

  public get BarColor(): Color {
    return Color.FromCore(this.model.BarColor);
  }

  public set BarColor(value: Color) {
    this.model.BarColor = toCoreColor(value);
  }

  public get NegativeBarColor(): Color {
    return Color.FromCore(this.model.NegativeBarColor);
  }

  public set NegativeBarColor(value: Color) {
    this.model.NegativeBarColor = toCoreColor(value);
  }

  public get ShowBorder(): boolean {
    return this.model.ShowBorder;
  }

  public set ShowBorder(value: boolean) {
    this.model.ShowBorder = value;
  }

  public get Direction(): string {
    return this.model.Direction ?? "";
  }

  public set Direction(value: string) {
    this.model.Direction = normalizeText(value);
  }

  public get BarLength(): string {
    return this.model.BarLength ?? "";
  }

  public set BarLength(value: string) {
    this.model.BarLength = normalizeText(value);
  }

  public get IconSetType(): string {
    return this.model.IconSetType ?? "";
  }

  public set IconSetType(value: string) {
    this.model.IconSetType = normalizeText(value);
  }

  public get ReverseIcons(): boolean {
    return this.model.ReverseIcons;
  }

  public set ReverseIcons(value: boolean) {
    this.model.ReverseIcons = value;
  }

  public get ShowIconOnly(): boolean {
    return this.model.ShowIconOnly;
  }

  public set ShowIconOnly(value: boolean) {
    this.model.ShowIconOnly = value;
  }

  public get Priority(): number {
    return this.model.Priority;
  }

  public set Priority(value: number) {
    if (value <= 0) {
      throw new CellsException("Conditional formatting priority must be greater than zero.");
    }

    this.model.Priority = value;
  }

  public get StopIfTrue(): boolean {
    return this.model.StopIfTrue;
  }

  public set StopIfTrue(value: boolean) {
    this.model.StopIfTrue = value;
  }

  public get Style(): Style {
    return Style.FromCore(this.model.Style).clone();
  }

  public set Style(value: Style) {
    this.model.Style = value === null || value === undefined ? StyleValue.Default.clone() : value.ToCore();
  }

  public Remove(): void {
    FormatConditionCollection.removeCondition(this.owner, this.collection, this.model);
  }
}

export function normalizeFormula(value: string | null | undefined): string | null {
  if (value === null || value === undefined) {
    return null;
  }

  const trimmed = value.trim();
  if (trimmed.length === 0) {
    return null;
  }

  return trimmed[0] === "=" ? trimmed.substring(1) : trimmed;
}

export function normalizeText(value: string | null | undefined): string | null {
  if (value === null || value === undefined) {
    return null;
  }

  const trimmed = value.trim();
  return trimmed.length === 0 ? null : trimmed;
}
