import { Workbook } from "../workbook";
import { Worksheet } from "../worksheet";
import { Cell } from "../cell";
import type { CellValue } from "../types";
import { cellRef } from "../util";

export { Workbook, Worksheet, Cell };
export type { CellValue };

export interface HtmlParseOptions {
  trimWhitespace?: boolean;
  useFirstRowAsHeader?: boolean;
}

export interface HtmlSaveOptions {
  includeStyles?: boolean;
  borderCollapse?: boolean;
}

export class HtmlDocument {
  private _html: string = "";
  private _tables: HtmlTable[] = [];

  static async load(filePath: string): Promise<HtmlDocument> {
    const { readFile } = await import("fs/promises");
    const content = await readFile(filePath, "utf-8");
    return HtmlDocument.parse(content);
  }

  static parse(html: string): HtmlDocument {
    const doc = new HtmlDocument();
    doc._html = html;
    doc._tables = HtmlTable.parseAll(html);
    return doc;
  }

  get tables(): HtmlTable[] {
    return this._tables;
  }

  get html(): string {
    return this._html;
  }

  toWorkbook(options?: HtmlParseOptions): Workbook {
    const workbook = new Workbook();
    workbook.worksheets.length = 0;

    for (const table of this._tables) {
      const worksheet = workbook.addWorksheet(
        table.caption || `Sheet${workbook.worksheetCount}`,
      );
      table.toWorksheet(worksheet, options);
    }

    if (this._tables.length === 0) {
      workbook.addWorksheet("Sheet1");
    }

    return workbook;
  }

  toExcel(filePath: string, options?: HtmlParseOptions) {
    const workbook = this.toWorkbook(options);
    return workbook.save(filePath);
  }
}

export class HtmlTable {
  private _caption?: string;
  private _headers: string[] = [];
  private _rows: string[][] = [];

  static parseAll(html: string): HtmlTable[] {
    const tables: HtmlTable[] = [];
    const tableRegex = /<table[^>]*>([\s\S]*?)<\/table>/gi;
    let match;

    while ((match = tableRegex.exec(html)) !== null) {
      const table = HtmlTable.parse(match[1]);
      tables.push(table);
    }

    return tables;
  }

  static parse(html: string): HtmlTable {
    const table = new HtmlTable();

    const captionMatch = /<caption[^>]*>([\s\S]*?)<\/caption>/i.exec(html);
    if (captionMatch) {
      table._caption = captionMatch[1].replace(/<[^>]+>/g, "").trim();
    }

    const theadMatch = /<thead[^>]*>([\s\S]*?)<\/thead>/i.exec(html);
    if (theadMatch) {
      const thRegex = /<th[^>]*>([\s\S]*?)<\/th>/gi;
      let thMatch;
      while ((thMatch = thRegex.exec(theadMatch[1])) !== null) {
        table._headers.push(HtmlTable.extractText(thMatch[1]));
      }
    }

    const tbodyMatch =
      html.match(/<tbody[^>]*>([\s\S]*?)<\/tbody>/i) ||
      html.match(/<table[^>]*>([\s\S]*?)<\/table>/i);
    const tbody = tbodyMatch ? tbodyMatch[1] : html;

    const trRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
    let trMatch;
    while ((trMatch = trRegex.exec(tbody)) !== null) {
      const cells: string[] = [];
      const tdRegex = /<t[hd][^>]*>([\s\S]*?)<\/t[hd]>/gi;
      let tdMatch;
      while ((tdMatch = tdRegex.exec(trMatch[1])) !== null) {
        cells.push(HtmlTable.extractText(tdMatch[1]));
      }
      if (cells.length > 0) {
        table._rows.push(cells);
      }
    }

    return table;
  }

  private static extractText(html: string): string {
    return html
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<[^>]+>/g, "")
      .replace(/&nbsp;/g, " ")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .trim();
  }

  get caption(): string | undefined {
    return this._caption;
  }

  get headers(): string[] {
    return this._headers;
  }

  get rows(): string[][] {
    return this._rows;
  }

  addRow(row: string[]) {
    this._rows.push(row);
  }

  get rowCount(): number {
    return this._rows.length;
  }

  get columnCount(): number {
    return Math.max(this._headers.length, ...this._rows.map((r) => r.length));
  }

  toWorksheet(worksheet: Worksheet, options?: HtmlParseOptions) {
    const opts = {
      trimWhitespace: true,
      useFirstRowAsHeader: false,
      ...options,
    };

    let startRow = 0;

    if (opts.useFirstRowAsHeader && this._headers.length > 0) {
      for (let col = 0; col < this._headers.length; col++) {
        const value = opts.trimWhitespace
          ? this._headers[col].trim()
          : this._headers[col];
        worksheet.putValue(cellRef(0, col), value);
      }
      startRow = 1;
    }

    for (let row = 0; row < this._rows.length; row++) {
      const dataRow = this._rows[row];
      for (let col = 0; col < dataRow.length; col++) {
        const value = opts.trimWhitespace ? dataRow[col].trim() : dataRow[col];
        const cellValue = this.parseValue(value);
        worksheet.putValue(cellRef(row + startRow, col), cellValue);
      }
    }
  }

  private parseValue(value: string): CellValue {
    if (value === "") return null;

    const num = parseFloat(value);
    if (!isNaN(num) && value.trim() === String(num)) {
      return num;
    }

    if (value.toLowerCase() === "true") return true;
    if (value.toLowerCase() === "false") return false;

    const dateMatch = value.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (dateMatch) {
      return new Date(
        parseInt(dateMatch[1]),
        parseInt(dateMatch[2]) - 1,
        parseInt(dateMatch[3]),
      );
    }

    return value;
  }

  toHtml(options?: HtmlSaveOptions): string {
    const opts = { includeStyles: true, borderCollapse: true, ...options };
    const styles: string[] = [];

    if (opts.borderCollapse) {
      styles.push("border-collapse: collapse");
    }

    let html = `<table`;
    if (opts.includeStyles && styles.length > 0) {
      html += ` style="${styles.join("; ")}"`;
    }
    html += ">\n";

    if (this._caption) {
      html += `  <caption>${this._caption}</caption>\n`;
    }

    if (this._headers.length > 0) {
      html += "  <thead>\n    <tr>\n";
      for (const header of this._headers) {
        html += `      <th>${this.escapeHtml(header)}</th>\n`;
      }
      html += "    </tr>\n  </thead>\n";
    }

    html += "  <tbody>\n";
    for (const row of this._rows) {
      html += "    <tr>\n";
      for (const cell of row) {
        html += `      <td>${this.escapeHtml(cell)}</td>\n`;
      }
      html += "    </tr>\n";
    }
    html += "  </tbody>\n</table>";

    return html;
  }

  private escapeHtml(str: string): string {
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }
}

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

export function tableToHtml(
  table: HtmlTable,
  options?: HtmlSaveOptions,
): string {
  return table.toHtml(options);
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
