import { HtmlReader } from "./htmlReader";
import { Worksheet } from "../worksheet";
import { cellRef } from "../util";
import type { CellValue } from "../types";
import type { HtmlParseOptions } from "./htmlDocument";

export interface HtmlSaveOptions {
  includeStyles?: boolean;
  borderCollapse?: boolean;
}

export class HtmlTable {
  private _caption?: string;
  private _headers: string[] = [];
  private _rows: string[][] = [];

  constructor(caption?: string) {
    this._caption = caption;
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

  addHeader(header: string): void {
    this._headers.push(header);
  }

  addRow(row: string[]): void {
    this._rows.push(row);
  }

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
        const cellValue = HtmlTable.parseValue(value);
        worksheet.putValue(cellRef(row + startRow, col), cellValue);
      }
    }
  }

  private static parseValue(value: string): CellValue {
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

  static parse(html: string): HtmlTable {
    const table = new HtmlTable();
    const reader = new HtmlReader(html);

    let inThead = false;
    let inTbody = false;
    let inTr = false;
    let inTh = false;
    let inTd = false;
    let inCaption = false;
    let currentRow: string[] = [];
    let currentCell = "";
    let captionText = "";

    while (reader.read()) {
      const tag = reader.tagName;
      const isStart = reader.isStartTag;
      const isEnd = reader.isEndTag;

      if (tag === "caption" && isStart) {
        inCaption = true;
        captionText = "";
      } else if (tag === "caption" && isEnd && inCaption) {
        inCaption = false;
        table._caption = captionText.trim();
      } else if (inCaption && (isStart || isEnd)) {
        captionText += " ";
      } else if (inCaption && reader.text) {
        captionText += reader.text;
      }

      if (tag === "thead" && isStart) {
        inThead = true;
      } else if (tag === "thead" && isEnd) {
        inThead = false;
      }

      if (tag === "tbody" && isStart) {
        inTbody = true;
      } else if (tag === "tbody" && isEnd) {
        inTbody = false;
      }

      if (tag === "tr" && isStart) {
        inTr = true;
        currentRow = [];
      } else if (tag === "tr" && isEnd && inTr) {
        inTr = false;
        if (inThead && currentRow.length > 0) {
          table._headers = currentRow;
        } else if (currentRow.length > 0) {
          table._rows.push(currentRow);
        }
      }

      if (tag === "th" && isStart) {
        if (inTr) {
          if (inThead) {
            inTh = true;
          } else {
            inTd = true;
          }
        }
        currentCell = "";
      } else if (tag === "th" && isEnd && inTh) {
        inTh = false;
        if (inTr) {
          currentRow.push(HtmlTable.extractText(currentCell));
        }
      } else if (tag === "td" && isStart) {
        if (inTr) {
          inTd = true;
        }
        currentCell = "";
      } else if (tag === "td" && isEnd && inTd) {
        inTd = false;
        if (inTr) {
          currentRow.push(HtmlTable.extractText(currentCell));
        }
      } else if ((inTd || inTh) && reader.text) {
        currentCell += reader.text;
      } else if ((inTd || inTh) && isStart && reader.tagName === "br") {
        currentCell += "\n";
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
      html += `  <caption>${this.escapeHtml(this._caption)}</caption>\n`;
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
