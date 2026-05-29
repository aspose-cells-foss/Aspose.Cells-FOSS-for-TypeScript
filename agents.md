# AGENTS.md

## Purpose

This file defines how agents should produce and modify code for `Aspose.Cells FOSS for TypeScript`.

The project is a pure-TypeScript library for creating, loading, editing, and saving Excel `.xlsx` workbooks without Microsoft Excel. It mirrors a subset of the Aspose.Cells API surface.

## Source Layout

```
aspose_cells/
  workbook.ts          — Workbook class, load/save entry points, XLSX package assembly
  worksheet.ts         — Worksheet class (cells, validations, comments, hyperlinks, shapes, etc.)
  worksheetCollection.ts — collection of worksheets, style repository
  cell.ts              — Cell class, value, formula, style index
  style.ts             — Style class (font, fill, border, alignment, protection)
  types.ts             — all type definitions, interfaces, enums
  autofilter.ts        — AutoFilter
  chart.ts / chartCollection.ts / chartLoader.ts — chart support
  comment.ts           — comment support
  conditionalformat.ts — conditional formatting
  hyperlink.ts         — hyperlink support
  validation.ts        — data validation
  util.ts              — helpers (cellRef, parseCellRef, parseRange, zip I/O, XML utilities)
  impDrawing.ts        — drawing import (shapes, images from OPC)
  expDrawing.ts        — drawing export (shapes, images to OPC)
  dataRelationship.ts  — OPC relationships generation
  html/                — HTML import/export subsystem
    htmlDocument.ts    — HTML document model
    htmlExporter.ts    — workbook → HTML export
    htmlReader.ts      — HTML → workbook import
    htmlTable.ts       — table-level HTML parsing
    htmlWriter.ts      — table-level HTML writing
    chartRenderer.ts   — chart rendering for HTML
    index.ts           — barrel export
examples/              — runnable example scripts
index.ts               — top-level entry point (usage demo)
```

## Coding Rules

- one production class per `.ts` file
- keep each `.ts` file under 1000 lines
- use `CellValue` union type (`string | number | boolean | Date | null`)
- use explicit `Map<K, V>` and `Set<T>` for collections (not plain objects)
- prefer `for` loops over array methods for performance-critical paths
- do **not** use third-party spreadsheet libraries (no ExcelJS, no SheetJS)
- XML handling via `@xmldom/xmldom` only
- zip handling via `@zip.js/zip.js` for reading, `adm-zip` or manual writer for writing
- async load/save methods (`Workbook.load()`, `workbook.save()`)
- keep public API compatible with Aspose.Cells naming conventions where possible

## Implementation Workflow

For every feature or bug fix:

1. Read the relevant source files under `aspose_cells/` to understand current behavior.
2. Read the matching type definitions in `types.ts`.
3. Check `examples/` for usage patterns.
4. Read `aspose_cells/html/` files if the change touches HTML import/export.
5. Implement load, save, and round-trip behavior together.
6. Verify with a test script (run `npx tsx <test_script.ts>`).

## Feature Areas

### Workbook and OPC Package

Files:
- `workbook.ts`
- `worksheetCollection.ts`
- `dataRelationship.ts`
- `util.ts`

Responsible for:
- `Workbook` class (constructor, `load()`, `save()`)
- `WorksheetCollection`
- OPC package part registration (`[Content_Types].xml`, `_rels/.rels`)
- workbook relationships (`xl/_rels/workbook.xml.rels`)
- shared strings (`xl/sharedStrings.xml`)
- styles (`xl/styles.xml`)
- core properties, app properties, theme
- chart XML generation for chart shapes

### Worksheet Grid and Metadata

Files:
- `worksheet.ts`
- `autofilter.ts`

Responsible for:
- worksheet XML generation (`xl/worksheets/sheetN.xml`)
- row/column metadata (widths, heights, hidden rows)
- merge regions
- auto-filter
- sheet protection
- default column width / row height

### Cell Values and Addressing

Files:
- `cell.ts`
- `util.ts` (cellRef, parseCellRef, parseRange)

Responsible for:
- A1-style and zero-based addressing
- `Cell` class (value, formula, style)
- `putValue()`, `setFormula()`, `setStyle()`
- inline string vs shared string logic
- date serial conversion (`dateToExcelSerial`)
- formula persistence

### Styles

Files:
- `style.ts`
- `types.ts` — `Style`, `Font`, `Fill`, `Border`, `Alignment`, `Protection` interfaces

Responsible for:
- `Style` class with fluent setters
- style index allocation in `WorksheetCollection`
- font, fill, border, alignment, number format, protection
- cell style lookup by index during load

### Data Validation

Files:
- `validation.ts`

Responsible for:
- `DataValidation` and `DataValidationCollection`
- load/save of `<dataValidations>` in worksheet XML

### Conditional Formatting

Files:
- `conditionalformat.ts`

Responsible for:
- `ConditionalFormat` and `ConditionalFormatCollection`
- load/save of `<conditionalFormatting>` in worksheet XML
- color scale, icon set, data bar rules

### Hyperlinks and Comments

Files:
- `hyperlink.ts`
- `comment.ts`

Responsible for:
- `Hyperlink` and `HyperlinkCollection` — load/save of `<hyperlinks>`
- `Comment` and `CommentCollection` — load/save of `<comments>`

### Drawings, Shapes, Pictures

Files:
- `impDrawing.ts`
- `expDrawing.ts`
- `dataRelationship.ts`

Responsible for:
- shape/picture/chart XML import from `xl/drawings/drawingN.xml`
- shape/picture/chart XML export
- drawing relationships (`xl/drawings/_rels/drawingN.xml.rels`)
- position calculation (EMU-based anchor computation)

### Charts

Files:
- `chart.ts`
- `chartCollection.ts`
- `chartLoader.ts`
- `workbook.ts` — chart XML generation methods

Responsible for:
- `ChartInfo` type (`types.ts`)
- chart types: bar, column, line, pie, doughnut, area, scatter, radar
- stacked variants
- chart title, series, legend, axes
- chart XML generation for export
- chart loading from existing XLSX (`xl/charts/chartN.xml`)

### HTML Import/Export

Files:
- `html/` directory

Responsible for:
- `Workbook.load()` for `.html` files
- `workbook.save()` with `SaveFormat.HTML`
- workbook-to-HTML frameset export
- HTML-to-workbook parsing
- chart rendering in HTML
- CSS-based styling for HTML output
- image extraction/embedding for HTML

## Definition Of Done

A feature is complete only when all of the following are true:

- public API is consistent with `aspose_cells/index.ts` exports
- XLSX load/save round-trips without data loss for the feature
- HTML import/export is consistent where applicable
- generated XLSX opens correctly in Microsoft Excel
- no silent data loss on malformed input
- test script in `examples/` or a standalone `.ts` script demonstrates the feature

## Anti-Patterns

Do not:

- add dependencies on ExcelJS, SheetJS, or similar spreadsheet libraries
- use `any` type when a proper union or interface exists in `types.ts`
- mix `CellValue` domain (string|number|boolean|Date|null) with raw XML strings
- hardcode sheet count limits or column letter arithmetic past ZZ
- silently truncate or drop data during load without a console warning
- add unsupported Excel features as half-implemented placeholders
- use sync I/O where async is available (load/save paths)

## Build & Test

```bash
npm install
npx tsc --noEmit          # type-check
npx tsx index.ts           # quick smoke test
npx tsx examples/<name>.ts # run an example
```
