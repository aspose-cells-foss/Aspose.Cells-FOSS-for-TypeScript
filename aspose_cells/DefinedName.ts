import { CellsException } from "./CellsException";
import { normalizeDefinedName, normalizeDefinedNameComment, normalizeDefinedNameFormula } from "./DefinedNameSupport";
import { DefinedNameModel } from "./Core/DefinedNameModel";

/** Represents a defined name. */
export class DefinedName {
  private readonly workbook: IWorkbookDefinedNameHost;
  private readonly model: DefinedNameModel;

  public constructor(workbook: IWorkbookDefinedNameHost, model: DefinedNameModel) {
    this.workbook = workbook;
    this.model = model;
  }

  public get Name(): string {
    return this.model.Name;
  }

  public set Name(value: string) {
    const normalized = normalizeDefinedName(value);
    this.workbook.EnsureUniqueDefinedName(this.model, normalized, this.model.LocalSheetIndex);
    this.model.Name = normalized;
  }

  public get Formula(): string {
    return this.model.Formula;
  }

  public set Formula(value: string) {
    this.model.Formula = normalizeDefinedNameFormula(value);
  }

  public get LocalSheetIndex(): number | null {
    return this.model.LocalSheetIndex;
  }

  public set LocalSheetIndex(value: number | null) {
    this.workbook.EnsureValidDefinedNameScope(value);
    this.workbook.EnsureUniqueDefinedName(this.model, this.model.Name, value);
    this.model.LocalSheetIndex = value;
  }

  public get Hidden(): boolean {
    return this.model.Hidden;
  }

  public set Hidden(value: boolean) {
    this.model.Hidden = value;
  }

  public get Comment(): string {
    return this.model.Comment;
  }

  public set Comment(value: string) {
    this.model.Comment = normalizeDefinedNameComment(value);
  }
}

/** Host contract used by DefinedName to validate against the owning workbook. */
export interface IWorkbookDefinedNameHost {
  EnsureValidDefinedNameScope(localSheetIndex: number | null): void;
  EnsureUniqueDefinedName(
    currentDefinedName: DefinedNameModel | null,
    name: string,
    localSheetIndex: number | null,
  ): void;
}

/** Throws when a defined name index is out of range. */
export function validateDefinedNameIndex(index: number, count: number): void {
  if (index < 0 || index >= count) {
    throw new CellsException("Defined name index was out of range.");
  }
}
