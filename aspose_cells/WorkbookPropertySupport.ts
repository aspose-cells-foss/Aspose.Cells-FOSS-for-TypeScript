import { CellsException } from "./CellsException";

/** Normalization helpers shared by workbook property facades. */
export function normalizeChoice(
  value: string | null,
  propertyName: string,
  allowed: readonly string[],
): string {
  if (value === null || value === undefined || value.trim().length === 0) {
    return "";
  }

  const trimmed = value.trim();
  for (const candidate of allowed) {
    if (candidate.toLowerCase() === trimmed.toLowerCase()) {
      return candidate;
    }
  }

  throw new CellsException(`Unsupported ${propertyName} value '${value}'.`);
}

export function normalizeShowObjects(value: string | null): string {
  return normalizeChoice(value, "showObjects", ["all", "placeholders", "none"]);
}

export function normalizeUpdateLinks(value: string | null): string {
  return normalizeChoice(value, "updateLinks", ["userSet", "never", "always"]);
}

export function normalizeVisibility(value: string | null): string {
  return normalizeChoice(value, "visibility", ["visible", "hidden", "veryHidden"]);
}

export function normalizeCalculationMode(value: string | null): string {
  return normalizeChoice(value, "calcMode", ["auto", "manual", "autoNoTable"]);
}

export function normalizeReferenceMode(value: string | null): string {
  return normalizeChoice(value, "refMode", ["A1", "R1C1"]);
}
