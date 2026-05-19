# Aspose.Cells FOSS for TypeScript

Open-source TypeScript library for creating, reading, modifying, and converting
Excel files programmatically -- without requiring Microsoft Excel.

## Features

- Read and write XLSX files
- HTML export and import
- Cell styling (fonts, fills, borders, number formats)
- Charts (bar, line, pie, scatter) and shapes
- Data validation and conditional formatting
- Hyperlinks, comments, auto-filter
- Merged cells, column/row width control
- Multiple worksheet support

## Installation

```bash
npm install @aspose/cells
```

Or clone and use directly:

```bash
git clone https://github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-TypeScript.git
cd Aspose.Cells-FOSS-for-TypeScript
npm install
```

## Quick Start

```typescript
import { Workbook } from "./aspose_cells";

const workbook = new Workbook();
const sheet = workbook.worksheets.get(0)!;
sheet.putValue("A1", "Hello, Aspose.Cells!");
workbook.save("output.xlsx");
```

## Documentation

- [Product Page](https://products.aspose.org/cells/)

## License

MIT -- see [LICENSE](LICENSE) file.
