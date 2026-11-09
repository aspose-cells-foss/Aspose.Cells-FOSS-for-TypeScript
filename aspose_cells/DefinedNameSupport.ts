import { CellsException } from "./CellsException";

const PrintAreaDefinedName = "_xlnm.Print_Area";
const PrintTitlesDefinedName = "_xlnm.Print_Titles";
const FilterDatabaseDefinedName = "_xlnm._FilterDatabase";

export function isReservedDefinedName(name: string): boolean {
  return (
    name.toLowerCase() === PrintAreaDefinedName.toLowerCase() ||
    name.toLowerCase() === PrintTitlesDefinedName.toLowerCase() ||
    name.toLowerCase() === FilterDatabaseDefinedName.toLowerCase()
  );
}

export function normalizeDefinedName(name: string | null): string {
  const normalized = (name ?? "").trim();
  if (normalized.length === 0) {
    throw new CellsException("Defined name must be non-empty.");
  }

  if (isReservedDefinedName(normalized)) {
    throw new CellsException("Built-in print defined names must be managed through PageSetup.");
  }

  return normalized;
}

export function normalizeDefinedNameFormula(formula: string | null): string {
  let normalized = (formula ?? "").trim();
  if (normalized.startsWith("=")) {
    normalized = normalized.substring(1).trim();
  }

  if (normalized.length === 0) {
    throw new CellsException("Defined name formula must be non-empty.");
  }

  return normalized;
}

export function normalizeDefinedNameComment(comment: string | null): string {
  return (comment ?? "").trim();
}

export function sameDefinedNameScope(left: number | null, right: number | null): boolean {
  if (left === null && right === null) {
    return true;
  }

  return left !== null && right !== null && left === right;
}
