import { HtmlReader } from "./htmlReader";
import { Worksheet } from "../worksheet";
import { cellRef } from "../util";
import type {
  CellValue,
  Style,
  Font,
  Fill,
  Border,
  BorderLine,
  Alignment,
} from "../types";
import type { HtmlParseOptions } from "./htmlDocument";

export interface HtmlSaveOptions {
  includeStyles?: boolean;
  borderCollapse?: boolean;
}

export interface CellStyle {
  align?: string;
  valign?: string;
  bgcolor?: string;
  color?: string;
  font?: {
    name?: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
  };
  border?: {
    left?: { style?: string; color?: string };
    right?: { style?: string; color?: string };
    top?: { style?: string; color?: string };
    bottom?: { style?: string; color?: string };
  };
}

export class HtmlTable {
  private _caption?: string;
  private _headers: string[] = [];
  private _rows: string[][] = [];
  private _colWidths: number[] = [];
  private _rowHeights: number[] = [];
  private _cellStyles: Map<string, CellStyle> = new Map();
  private _cssClasses: Map<string, CellStyle> = new Map();
  private _cellClasses: Map<string, string> = new Map();

  constructor(caption?: string) {
    this._caption = caption;
  }

  setCssClasses(classes: Map<string, CellStyle>) {
    this._cssClasses = classes;
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

  get colWidths(): number[] {
    return this._colWidths;
  }

  get rowHeights(): number[] {
    return this._rowHeights;
  }

  getCellStyle(row: number, col: number): CellStyle | undefined {
    return this._cellStyles.get(`${row},${col}`);
  }

  addHeader(header: string): void {
    this._headers.push(header);
  }

  addRow(row: string[]): void {
    this._rows.push(row);
  }

  static parseAll(
    html: string,
    cssClasses?: Map<string, CellStyle>,
  ): HtmlTable[] {
    const tables: HtmlTable[] = [];
    const tableRegex = /<table[^>]*>([\s\S]*?)<\/table>/gi;
    let match;

    while ((match = tableRegex.exec(html)) !== null) {
      const table = HtmlTable.parse(match[1], cssClasses);
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

    for (let col = 0; col < this._colWidths.length; col++) {
      worksheet.setColumnWidth(col, this._colWidths[col]);
    }

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
      if (row < this._rowHeights.length) {
        worksheet.setRowHeight(row + startRow, this._rowHeights[row]);
      }

      const dataRow = this._rows[row];
      for (let col = 0; col < dataRow.length; col++) {
        const rawValue = opts.trimWhitespace
          ? dataRow[col].trim()
          : dataRow[col];
        const cellValue = HtmlTable.parseValue(rawValue);
        const cellStyle = this.getCellStyle(row, col);

        // Write the cell only if it has a non‑null value or a style applied
        if (cellValue !== null || cellStyle) {
          const cellRefStr = cellRef(row + startRow, col);
          worksheet.putValue(cellRefStr, cellValue);
        }

        if (cellStyle) {
          const style: any = {};

          if (cellStyle.align || cellStyle.valign) {
            style.alignment = {};
            if (cellStyle.align) style.alignment.horizontal = cellStyle.align;
            if (cellStyle.valign) style.alignment.vertical = cellStyle.valign;
          }

          if (cellStyle.bgcolor) {
            style.fill = {
              patternType: "solid",
              fgColor: cellStyle.bgcolor,
            };
          }

          if (
            cellStyle.font?.bold ||
            cellStyle.font?.italic ||
            cellStyle.color
          ) {
            style.font = {};
            if (cellStyle.font?.bold) style.font.bold = true;
            if (cellStyle.font?.italic) style.font.italic = true;
            if (cellStyle.color) style.font.color = cellStyle.color;
          }

          if (cellStyle.border) {
            style.border = {};
            if (cellStyle.border.left)
              style.border.left = cellStyle.border.left;
            if (cellStyle.border.right)
              style.border.right = cellStyle.border.right;
            if (cellStyle.border.top) style.border.top = cellStyle.border.top;
            if (cellStyle.border.bottom)
              style.border.bottom = cellStyle.border.bottom;
          }

          if (Object.keys(style).length > 0) {
            worksheet.setCellStyle(cellRef(row + startRow, col), style);
          }
        }
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

  static parse(html: string, cssClasses?: Map<string, CellStyle>): HtmlTable {
    const table = new HtmlTable();
    if (cssClasses) {
      table._cssClasses = cssClasses;
    }
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
    let currentRowIndex = 0;
    let currentColIndex = 0;
    // Flag to skip rows that are hidden via CSS (display:none)
    let skipCurrentRow = false;

    const parseStyle = (styleStr: string): CellStyle => {
      const style: CellStyle = {};
      if (!styleStr) return style;

      const parts = styleStr.split(";");
      for (const part of parts) {
        const [key, ...valueParts] = part.split(":");
        if (!key || valueParts.length === 0) continue;
        const value = valueParts.join(":").trim();

        if (key.trim() === "background-color" || key.trim() === "background") {
          style.bgcolor = value;
        } else if (key.trim() === "color") {
          style.color = value;
        } else if (key.trim() === "font-weight" && value === "bold") {
          style.font = { ...style.font, bold: true };
        } else if (key.trim() === "font-style" && value === "italic") {
          style.font = { ...style.font, italic: true };
        } else if (key.trim() === "text-align") {
          style.align = value;
        } else if (key.trim() === "vertical-align") {
          style.valign = value;
        } else if (key.trim() === "border-left") {
          style.border = {
            ...style.border,
            left: HtmlTable.parseBorderSide(value),
          };
        } else if (key.trim() === "border-right") {
          style.border = {
            ...style.border,
            right: HtmlTable.parseBorderSide(value),
          };
        } else if (key.trim() === "border-top") {
          style.border = {
            ...style.border,
            top: HtmlTable.parseBorderSide(value),
          };
        } else if (key.trim() === "border-bottom") {
          style.border = {
            ...style.border,
            bottom: HtmlTable.parseBorderSide(value),
          };
        }
      }
      return style;
    };

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

      if (tag === "col" && isStart) {
        const width = reader.getAttribute("width");
        const span = reader.getAttribute("span");
        const style = reader.getAttribute("style") || "";
        let colWidth = width ? parseFloat(width) : 64;
        const styleWidthMatch = style.match(/width:\s*(\d+(?:\.\d+)?)(pt|px)/);
        if (styleWidthMatch) {
          colWidth = parseFloat(styleWidthMatch[1]);
          if (styleWidthMatch[2] === "pt") {
            colWidth = Math.round(colWidth * 4 / 3); // pt to px at 96dpi
          }
        }
        const spanCount = span ? parseInt(span, 10) : 1;
        for (let i = 0; i < spanCount; i++) {
          table._colWidths.push(colWidth);
        }
      }

      if (tag === "tr" && isStart) {
        inTr = true;
        currentRow = [];
        currentColIndex = 0;
        const height = reader.getAttribute("height");
        const style = reader.getAttribute("style") || "";
        // Detect hidden rows via display:none and skip processing them
        skipCurrentRow = /display\s*:\s*none/i.test(style);
        let rowHeight = height ? parseFloat(height) : 16;
        const styleHeightMatch = style.match(
          /height:\s*(\d+(?:\.\d+)?)(pt|px)/,
        );
        if (styleHeightMatch) {
          rowHeight = parseFloat(styleHeightMatch[1]);
          if (styleHeightMatch[2] === "pt") {
            rowHeight = rowHeight * 1.33;
          }
        }
        if (!skipCurrentRow) {
          table._rowHeights.push(rowHeight);
        }
      } else if (tag === "tr" && isEnd && inTr) {
        inTr = false;
        currentRowIndex++;
        if (!skipCurrentRow) {
          if (inThead && currentRow.length > 0) {
            table._headers = currentRow;
          } else if (currentRow.length > 0) {
            table._rows.push(currentRow);
          }
        }
      }

      if ((tag === "th" || tag === "td") && isStart) {
        if (inTr) {
          // Skip processing cells in hidden rows
          if (!skipCurrentRow) {
            inTd = true;
            currentCell = "";

            const styleStr = reader.getAttribute("style") || "";
            const align = reader.getAttribute("align");
            const valign = reader.getAttribute("valign");
            const bgcolor = reader.getAttribute("bgcolor");
            const cssClass = reader.getAttribute("class");

            let cellStyle: CellStyle = parseStyle(styleStr);

            if (cssClass && table._cssClasses.has(cssClass)) {
              const cssStyle = table._cssClasses.get(cssClass)!;
              cellStyle = { ...cssStyle, ...cellStyle };
            }

            if (align) cellStyle.align = align;
            if (valign) cellStyle.valign = valign;
            if (bgcolor) cellStyle.bgcolor = bgcolor;

            // Record style only for visible rows
            if (Object.keys(cellStyle).length > 0) {
              table._cellStyles.set(
                `${table._rows.length},${currentColIndex}`,
                cellStyle,
              );
            }
          } else {
            // Still need to advance column index for hidden rows
            inTd = true;
            currentCell = "";
          }
        }
        currentColIndex++;
      } else if ((tag === "th" || tag === "td") && isEnd && inTd) {
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

  private static parseBorderSide(value: string): {
    style?: string;
    color?: string;
  } {
    const result: { style?: string; color?: string } = {};
    const parts = value.split(" ");
    for (const part of parts) {
      if (part === "solid" || part === "dashed" || part === "dotted") {
        result.style = part;
      } else if (part.startsWith("#") || part.startsWith("rgb")) {
        result.color = part;
      }
    }
    return result;
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
