import { CellArea } from "./CellArea";

/** Converts a core cell area into the public facade type. */
export function toApiCellArea(area: CellArea): CellArea {
  return new CellArea(area.FirstRow, area.FirstColumn, area.TotalRows, area.TotalColumns);
}

/** Converts a public facade cell area into the core model type. */
export function toCoreCellArea(area: CellArea): CellArea {
  return new CellArea(area.FirstRow, area.FirstColumn, area.TotalRows, area.TotalColumns);
}
