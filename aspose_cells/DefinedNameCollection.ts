import { DefinedName, IWorkbookDefinedNameHost, validateDefinedNameIndex } from "./DefinedName";
import { normalizeDefinedName, normalizeDefinedNameFormula } from "./DefinedNameSupport";
import { DefinedNameModel } from "./Core/DefinedNameModel";

/** Represents a collection of defined names. */
export class DefinedNameCollection {
  [key: string]: any;

  private readonly workbook: DefinedNameCollectionHost;
  private readonly models: DefinedNameModel[];

  public constructor(workbook: DefinedNameCollectionHost) {
    this.workbook = workbook;
    this.models = workbook.Model.DefinedNames;
  }

  /** Gets the number of items. */
  public get Count(): number {
    return this.models.length;
  }

  /** Gets the element at the specified zero-based index. */
  public get(index: number): DefinedName {
    validateDefinedNameIndex(index, this.models.length);
    return new DefinedName(this.workbook, this.models[index]);
  }

  /** Alias of get. */
  public at(index: number): DefinedName {
    return this.get(index);
  }

  /** Adds a defined name. */
  public Add(name: string, formula: string): number;
  public Add(name: string, formula: string, localSheetIndex: number | null): number;
  public Add(name: string, formula: string, localSheetIndex: number | null = null): number {
    const normalizedName = normalizeDefinedName(name);
    const normalizedFormula = normalizeDefinedNameFormula(formula);
    this.workbook.EnsureValidDefinedNameScope(localSheetIndex);
    this.workbook.EnsureUniqueDefinedName(null, normalizedName, localSheetIndex);

    const model = new DefinedNameModel();
    model.Name = normalizedName;
    model.Formula = normalizedFormula;
    model.LocalSheetIndex = localSheetIndex;
    this.models.push(model);
    return this.models.length - 1;
  }

  /** Removes the item at the specified index. */
  public RemoveAt(index: number): void {
    validateDefinedNameIndex(index, this.models.length);
    this.models.splice(index, 1);
  }

  public *[Symbol.iterator](): IterableIterator<DefinedName> {
    for (let index = 0; index < this.models.length; index++) {
      yield this.get(index);
    }
  }
}

/** Host contract required by DefinedNameCollection. */
export interface DefinedNameCollectionHost extends IWorkbookDefinedNameHost {
  readonly Model: { readonly DefinedNames: DefinedNameModel[] };
}
