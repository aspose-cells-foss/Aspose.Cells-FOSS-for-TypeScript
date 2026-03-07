import { Workbook } from "../workbook";
import { Worksheet } from "../worksheet";
import { HtmlTable, HtmlSaveOptions } from "./htmlTable";

export class HtmlWriter {
  static fromWorkbook(workbook: Workbook, options?: HtmlSaveOptions): string {
    const tables: string[] = [];

    for (const worksheet of workbook.worksheets) {
      const table = HtmlWriter.fromWorksheet(worksheet, options);
      tables.push(table);
    }

    return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Excel Export</title>
</head>
<body>
${tables.join("\n\n")}
</body>
</html>`;
  }

  static fromWorksheet(
    worksheet: Worksheet,
    options?: HtmlSaveOptions,
  ): string {
    const table = new HtmlTable();

    const cells = Array.from(worksheet.cells);
    if (cells.length === 0) {
      return "<table></table>";
    }

    const maxRow = Math.max(...cells.map((c) => c.row));
    const maxCol = Math.max(...cells.map((c) => c.col));

    for (let row = 0; row <= maxRow; row++) {
      const rowData: string[] = [];
      for (let col = 0; col <= maxCol; col++) {
        const cell = worksheet.getCell(row, col);
        rowData.push(
          cell?.value !== null && cell?.value !== undefined
            ? String(cell.value)
            : "",
        );
      }
      table.addRow(rowData);
    }

    return table.toHtml(options);
  }
}

export function workbookToHtml(
  workbook: Workbook,
  options?: HtmlSaveOptions,
): string {
  return HtmlWriter.fromWorkbook(workbook, options);
}

export function worksheetToHtml(
  worksheet: Worksheet,
  options?: HtmlSaveOptions,
): string {
  return HtmlWriter.fromWorksheet(worksheet, options);
}

export function tableToHtml(
  table: HtmlTable,
  options?: HtmlSaveOptions,
): string {
  return table.toHtml(options);
}

export function workbookToHtmlFull(workbook: Workbook): string {
  const sheets = workbook.worksheets;
  let html = `<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<meta name="ProgId" content="Excel.Sheet"/>
<meta name="Generator" content="Aspose.Cells FOSS for TypeScript"/>
<style>
<!--table
 {mso-displayed-decimal-separator:"\\.";
 mso-displayed-thousand-separator:"\\,";}
@page
 {
 margin:1in 0.75in 1in 0.75in;
 mso-header-margin:0.5in;
 mso-footer-margin:0.5in;
 mso-page-orientation:Portrait;
 }
tr
 {mso-height-source:auto;
 mso-ruby-visibility:none;}
col
 {mso-width-source:auto;
 mso-ruby-visibility:none;}
br
 {mso-data-placement:same-cell;}
ruby
 {ruby-align:start;}
.style0
 {
 mso-number-format:General;
 text-align:general;
 vertical-align:middle;
 white-space:nowrap;
 background:auto;
 mso-pattern:auto;
 color:#000000;
 font-size:10pt;
 font-weight:400;
 font-style:normal;
 font-family:Arial,sans-serif;
 mso-protection:locked visible;
 mso-style-name:Normal;
 mso-style-id:0;}
td
 {mso-style-parent:style0;
 mso-number-format:General;
 text-align:general;
 vertical-align:middle;
 white-space:nowrap;
 background:auto;
 mso-pattern:auto;
 color:#000000;
 font-size:10pt;
 font-weight:400;
 font-style:normal;
 font-family:Arial,sans-serif;
 mso-protection:locked visible;
 mso-ignore:padding;
 }
-->
</style>
</head>
<body link='blue' vlink='purple' >
`;

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    html += worksheetToHtml(sheet);
  }

  html += `
</body>

</html>`;
  return html;
}
