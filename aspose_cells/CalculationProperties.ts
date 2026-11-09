import { CellsException } from "./CellsException";
import { normalizeCalculationMode, normalizeReferenceMode } from "./WorkbookPropertySupport";
import { CalculationPropertiesModel } from "./Core/CalculationPropertiesModel";

/** Represents workbook calculation settings. */
export class CalculationProperties {
  private readonly model: CalculationPropertiesModel;

  public constructor(model: CalculationPropertiesModel) {
    this.model = model;
  }

  public get CalculationId(): number | null {
    return this.model.CalculationId;
  }

  public set CalculationId(value: number | null) {
    if (value !== null && value !== undefined && value < 0) {
      throw new CellsException("CalculationId must be non-negative.");
    }

    this.model.CalculationId = value;
  }

  public get CalculationMode(): string {
    return this.model.CalculationMode === null || this.model.CalculationMode.length === 0
      ? "auto"
      : this.model.CalculationMode;
  }

  public set CalculationMode(value: string) {
    this.model.CalculationMode = normalizeCalculationMode(value);
  }

  public get FullCalculationOnLoad(): boolean {
    return this.model.FullCalculationOnLoad;
  }

  public set FullCalculationOnLoad(value: boolean) {
    this.model.FullCalculationOnLoad = value;
  }

  public get ReferenceMode(): string {
    return this.model.ReferenceMode === null || this.model.ReferenceMode.length === 0
      ? "A1"
      : this.model.ReferenceMode;
  }

  public set ReferenceMode(value: string) {
    this.model.ReferenceMode = normalizeReferenceMode(value);
  }

  public get Iterate(): boolean {
    return this.model.Iterate;
  }

  public set Iterate(value: boolean) {
    this.model.Iterate = value;
  }

  public get IterateCount(): number {
    return this.model.IterateCount ?? 100;
  }

  public set IterateCount(value: number) {
    if (value < 0) {
      throw new CellsException("IterateCount must be non-negative.");
    }

    this.model.IterateCount = value;
  }

  public get IterateDelta(): number {
    return this.model.IterateDelta ?? 0.001;
  }

  public set IterateDelta(value: number) {
    if (value < 0) {
      throw new CellsException("IterateDelta must be non-negative.");
    }

    this.model.IterateDelta = value;
  }

  public get FullPrecision(): boolean {
    return this.model.FullPrecision === null || this.model.FullPrecision;
  }

  public set FullPrecision(value: boolean) {
    this.model.FullPrecision = value;
  }

  public get CalculationCompleted(): boolean {
    return this.model.CalculationCompleted === null || this.model.CalculationCompleted;
  }

  public set CalculationCompleted(value: boolean) {
    this.model.CalculationCompleted = value;
  }

  public get CalculationOnSave(): boolean {
    return this.model.CalculationOnSave === null || this.model.CalculationOnSave;
  }

  public set CalculationOnSave(value: boolean) {
    this.model.CalculationOnSave = value;
  }

  public get ConcurrentCalculation(): boolean {
    return this.model.ConcurrentCalculation === null || this.model.ConcurrentCalculation;
  }

  public set ConcurrentCalculation(value: boolean) {
    this.model.ConcurrentCalculation = value;
  }

  public get ForceFullCalculation(): boolean {
    return this.model.ForceFullCalculation;
  }

  public set ForceFullCalculation(value: boolean) {
    this.model.ForceFullCalculation = value;
  }
}
