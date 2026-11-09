import { CellAddress } from "./Core/CellAddress";
import { CellRecord } from "./Core/CellRecord";
import { StyleValue } from "./Core/StyleValue";
import { WorksheetModel } from "./Core/WorksheetModel";

/**
 * Locates a cell record by coordinates.
 *
 * The core model stores records in a Map keyed by CellAddress instances, so lookups
 * must compare coordinates instead of relying on object identity.
 */
export function findCellRecord(
  model: WorksheetModel,
  row: number,
  column: number,
): CellRecord | undefined {
  for (const [address, record] of model.Cells) {
    if (address.RowIndex === row && address.ColumnIndex === column) {
      return record;
    }
  }

  return undefined;
}

/** Returns the existing record or allocates one using the workbook default style. */
export function getOrCreateCellRecord(
  model: WorksheetModel,
  row: number,
  column: number,
  defaultStyle: StyleValue,
): CellRecord {
  const existing = findCellRecord(model, row, column);
  if (existing !== undefined) {
    return existing;
  }

  const record = new CellRecord();
  record.Style = defaultStyle.clone();
  model.Cells.set(new CellAddress(row, column), record);
  return record;
}
