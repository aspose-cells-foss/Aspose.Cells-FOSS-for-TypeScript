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
  numberFormat?: string;
}

export interface PictureInfo {
  row: number;
  col: number;
  name: string;
  src: string;
  width: number;
  height: number;
  marginLeft: number;
  marginTop: number;
  zIndex: number;
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
  private _pendingColspan = 1;
  private _pendingRowspans: { cols: number; count: number; height: number }[] = [];
  private _currentRowHeight = 16;
  private _rowHidden: boolean[] = [];
  private _rowColspans: number[][] = [];
  private _rowXlRowspans: number[] = [];
  private _rowIsXlRow: boolean[] = [];
  private _cellXlRowspans: Map<string, number> = new Map();
  private _cellXlRowspanCells: Map<string, boolean> = new Map();
  private _cellXlRowspanColspan: Map<string, number> = new Map();
  private _pictures: PictureInfo[] = [];

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

  get hiddenRows(): boolean[] {
    return this._rowHidden;
  }

  get pictures(): PictureInfo[] {
    return this._pictures;
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

    for (let i = 0; i < this._rowHeights.length; i++) {
      worksheet.setRowHeight(i + startRow, this._rowHeights[i]);
    }

    for (let i = 0; i < this._rowHidden.length; i++) {
      if (this._rowHidden[i]) {
        worksheet.setRowHidden(i + startRow, true);
      }
    }

    const totalRows = Math.max(this._rows.length, this._rowHeights.length);

    const xlRowGroupStarts: number[] = [];
    for (let r = 0; r < totalRows; r++) {
      const xlRows = this._rowXlRowspans[r] || 1;
      if (xlRows > 1 && !xlRowGroupStarts.includes(r)) {
        xlRowGroupStarts.push(r);
      }
    }

    const xlRowGroupMap: Map<number, number> = new Map();
    for (let i = 0; i < xlRowGroupStarts.length; i++) {
      const groupStart = xlRowGroupStarts[i];
      const xlRows = this._rowXlRowspans[groupStart] || 1;
      for (let j = 0; j < xlRows; j++) {
        xlRowGroupMap.set(groupStart + j, groupStart);
      }
    }

    for (let row = 0; row < totalRows; row++) {
      const dataRow = this._rows[row] || [];
      const colspans = this._rowColspans[row] || [];
      const xlRows = this._rowXlRowspans[row] || 1;
      const groupStart = xlRowGroupMap.get(row) ?? -1;
      const isSpacerRow = groupStart >= 0 && row > groupStart;
      const isXlRow = groupStart >= 0 && row === groupStart;

      if (isXlRow) {
        const xlCellStyle = this.getCellStyle(groupStart, 0);
        const xlColspan = this._cellXlRowspanColspan.get(`${groupStart},0`) || 1;
        const xlStartCol = 0;
        for (let c = 0; c < this._colWidths.length; c++) {
          const xlCellsVal = this._cellXlRowspanCells.get(`${groupStart},${c}`);
          if (xlCellsVal && c >= xlStartCol && c < xlStartCol + xlColspan) continue;
          const cellRefStr = cellRef(row + startRow, c);
          const style: any = {};
          if (xlCellStyle) {
            if (xlCellStyle.align || xlCellStyle.valign) {
              style.alignment = {};
              if (xlCellStyle.align) style.alignment.horizontal = xlCellStyle.align;
              if (xlCellStyle.valign) style.alignment.vertical = xlCellStyle.valign;
            }
            if (xlCellStyle.bgcolor) {
              style.fill = { patternType: "solid", fgColor: xlCellStyle.bgcolor };
            }
            if (xlCellStyle.font) {
              style.font = {};
              if (xlCellStyle.font.bold) style.font.bold = true;
              if (xlCellStyle.font.italic) style.font.italic = true;
              if (xlCellStyle.color) style.font.color = xlCellStyle.color;
            }
            if (xlCellStyle.border) {
              style.border = {};
              if (xlCellStyle.border.left) style.border.left = xlCellStyle.border.left;
              if (xlCellStyle.border.right) style.border.right = xlCellStyle.border.right;
              if (xlCellStyle.border.top) style.border.top = xlCellStyle.border.top;
              if (xlCellStyle.border.bottom) style.border.bottom = xlCellStyle.border.bottom;
            }
          }
          worksheet.setCellStyle(cellRefStr, style);
        }
        continue;
      }

      let col = 0;
      for (let cellIdx = 0; cellIdx < dataRow.length; cellIdx++) {
        const colspan = colspans[cellIdx] || 1;
        const xlCellsVal = this._cellXlRowspanCells.get(`${groupStart},${col}`);
        const xlRowspan = this._cellXlRowspans.get(`${groupStart},${col}`) || 1;
        const isSpacerCell = isSpacerRow && xlCellsVal === true;
        const isCoveredByXlColspan = isSpacerRow && xlRowspan > 1 && col < this._colWidths.length;

        if (!isSpacerCell && !isCoveredByXlColspan) {
          const cellStyle = this.getCellStyle(row, cellIdx);
          const cellRefStr = cellRef(row + startRow, col);
          const rawValue = opts.trimWhitespace
            ? dataRow[cellIdx].trim()
            : dataRow[cellIdx];
          const cellValue = HtmlTable.parseValue(rawValue);

          if (cellValue !== null) {
            worksheet.putValue(cellRefStr, cellValue);
          }

          const style: any = {};

          if (cellStyle?.align || cellStyle?.valign) {
            style.alignment = {};
            if (cellStyle?.align) style.alignment.horizontal = cellStyle.align;
            if (cellStyle?.valign) style.alignment.vertical = cellStyle.valign;
          }

          if (cellStyle?.bgcolor) {
            style.fill = {
              patternType: "solid",
              fgColor: cellStyle.bgcolor,
            };
          }

          if (
            cellStyle?.font?.bold ||
            cellStyle?.font?.italic ||
            cellStyle?.color
          ) {
            style.font = {};
            if (cellStyle?.font?.bold) style.font.bold = true;
            if (cellStyle?.font?.italic) style.font.italic = true;
            if (cellStyle?.color) style.font.color = cellStyle.color;
          }

          if (cellStyle?.border) {
            style.border = {};
            if (cellStyle.border.left)
              style.border.left = cellStyle.border.left;
            if (cellStyle.border.right)
              style.border.right = cellStyle.border.right;
            if (cellStyle.border.top) style.border.top = cellStyle.border.top;
            if (cellStyle.border.bottom)
              style.border.bottom = cellStyle.border.bottom;
          }

          if (cellStyle?.numberFormat) {
            style.numberFormat = cellStyle.numberFormat;
          }

          worksheet.setCellStyle(cellRefStr, style);
        }

        if (!isSpacerCell && !isCoveredByXlColspan) {
          for (let padCol = col + 1; padCol < col + colspan; padCol++) {
            if (padCol >= this._colWidths.length) break;
            const padXlCellsVal = this._cellXlRowspanCells.get(`${groupStart},${padCol}`);
            const padXlRowspan = this._cellXlRowspans.get(`${groupStart},${padCol}`) || 1;
            const isPadCovered = isSpacerRow && (padXlCellsVal === true || padXlRowspan > 1);
            if (padCol !== col && !isPadCovered) {
              const padCellRefStr = cellRef(row + startRow, padCol);
              const padStyle = this.getCellStyle(row, cellIdx);
              const style: any = {};
              if (padStyle?.align || padStyle?.valign) {
                style.alignment = {};
                if (padStyle?.align) style.alignment.horizontal = padStyle.align;
                if (padStyle?.valign) style.alignment.vertical = padStyle.valign;
              }
              if (padStyle?.bgcolor) {
                style.fill = {
                  patternType: "solid",
                  fgColor: padStyle.bgcolor,
                };
              }
              if (
                padStyle?.font?.bold ||
                padStyle?.font?.italic ||
                padStyle?.color
                ) {
                  style.font = {};
                  if (padStyle.font?.bold) style.font.bold = true;
                  if (padStyle.font?.italic) style.font.italic = true;
                  if (padStyle.color) style.font.color = padStyle.color;
                }
                if (padStyle?.border) {
                  style.border = {};
                  if (padStyle.border.left)
                    style.border.left = padStyle.border.left;
                  if (padStyle.border.right)
                    style.border.right = padStyle.border.right;
                  if (padStyle.border.top) style.border.top = padStyle.border.top;
                  if (padStyle.border.bottom)
                    style.border.bottom = padStyle.border.bottom;
                }
              worksheet.setCellStyle(padCellRefStr, style);
            }
          }
        }

        col += colspan;
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
    let currentRowColspans: number[] = [];
    let currentCell = "";
    let currentCellColspan = 1;
    let currentCellXlRowspan = 1;
    let currentCellIsSpacer = false;
    let captionText = "";
    let currentRowIndex = 0;
    let currentColIndex = 0;
    let currentVisualCol = 0;
    let tdVisualCol = 0;
    let skipCurrentRow = false;
    let rowMaxXlRowspan = 1;
    let rowHtmlHeightPx = 16;
    let currentRowEffectiveXlRowspan = 1;

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
        currentRowColspans = [];
        currentColIndex = 0;
        currentVisualCol = 0;
        rowMaxXlRowspan = 1;
        const height = reader.getAttribute("height");
        const style = reader.getAttribute("style") || "";
        skipCurrentRow = /display\s*:\s*none/i.test(style);
        let rowHeightPt = 12;
        if (height) {
          const h = parseFloat(height);
          if (h > 0) rowHeightPt = h * 72 / 96;
        }
        const styleHeightMatch = style.match(/height\s*:\s*(\d+(?:\.\d+)?)\s*pt/i);
        if (styleHeightMatch) {
          rowHeightPt = parseFloat(styleHeightMatch[1]);
        }
        const xlMatch = style.match(/mso-xlrowspan:\s*(\d+)/i);
        if (xlMatch) {
          const xlVal = parseInt(xlMatch[1], 10);
          if (xlVal > rowMaxXlRowspan) {
            rowMaxXlRowspan = xlVal;
          }
        }
        rowHtmlHeightPx = rowHeightPt * 96 / 72;
      } else if (tag === "tr" && isEnd && inTr) {
        inTr = false;
        table._rowHidden.push(skipCurrentRow);
        if (inThead && currentRow.length > 0) {
          table._headers = currentRow;
        }
        const rowHeightPt = rowHtmlHeightPx * 72 / 96;
        const hasExplicitXlRowspan = rowMaxXlRowspan > 1;
        const xlRowspanFromHeight = Math.round(rowHeightPt / 12);
        const effectiveXlRowspan = hasExplicitXlRowspan 
          ? rowMaxXlRowspan 
          : (currentRow.length === 0 && xlRowspanFromHeight > 1 ? xlRowspanFromHeight - 1 : 1);
        currentRowEffectiveXlRowspan = effectiveXlRowspan;
        const effectiveRowHeight = hasExplicitXlRowspan && xlRowspanFromHeight > 1 
          ? rowHeightPt / xlRowspanFromHeight 
          : rowHeightPt;
        table._rowHeights.push(effectiveRowHeight);
        table._currentRowHeight = effectiveRowHeight;

        const isXlRow = effectiveXlRowspan > 1;
        if (currentRow.length > 0) {
          table._rows.push(currentRow);
          table._rowColspans.push([...currentRowColspans]);
          table._rowXlRowspans.push(effectiveXlRowspan);
        } else {
          table._rowColspans.push([]);
          table._rowXlRowspans.push(isXlRow ? effectiveXlRowspan : 1);
        }

        if (isXlRow) {
          for (let i = 1; i < effectiveXlRowspan; i++) {
            const spacerRow: string[] = [];
            const spacerColspans: number[] = [];
            for (let c = 0; c < table._colWidths.length; c++) {
              spacerRow.push("");
              spacerColspans.push(1);
            }
            table._rows.push(spacerRow);
            table._rowColspans.push(spacerColspans);
            table._rowHeights.push(effectiveRowHeight);
            table._rowHidden.push(false);
            table._rowXlRowspans.push(1);
          }
        }
        currentRowColspans = [];
        rowMaxXlRowspan = 1;
      }

      if ((tag === "th" || tag === "td") && isStart) {
        const colspan = parseInt(reader.getAttribute("colspan") || "1", 10);
        const remainingCols = table._colWidths.length - currentColIndex;
        const effectiveColspan = Math.min(colspan, remainingCols > 0 ? remainingCols : 1);
        currentCellColspan = effectiveColspan;
        currentCellXlRowspan = 1;
        currentCellIsSpacer = false;
        const styleStr = reader.getAttribute("style") || "";
        const isSpacerCell = styleStr.includes("mso-ignore:colspan");
        if (isSpacerCell) {
          currentCellIsSpacer = true;
        }
        const xlMatch = styleStr.match(/mso-xlrowspan:\s*(\d+)/i);
        if (xlMatch) {
          const xlVal = parseInt(xlMatch[1], 10);
          currentCellXlRowspan = xlVal;
          if (currentCellXlRowspan > rowMaxXlRowspan) {
            rowMaxXlRowspan = currentCellXlRowspan;
          }
        } else if (rowMaxXlRowspan > 1) {
          currentCellXlRowspan = rowMaxXlRowspan;
        }
        if (inTr && currentCellXlRowspan > 1) {
          const xlStartRow = table._rows.length;
          const xlStartCol = currentColIndex;
          table._cellXlRowspans.set(`${xlStartRow},${xlStartCol}`, currentCellXlRowspan);
          table._cellXlRowspanColspan.set(`${xlStartRow},${xlStartCol}`, effectiveColspan);
          const isSpacerCellWithColspan = isSpacerCell && effectiveColspan > 1;
          const spanCols = isSpacerCell ? 1 : effectiveColspan;
          for (let r = xlStartRow; r < xlStartRow + currentCellXlRowspan; r++) {
            for (let c = xlStartCol; c < xlStartCol + spanCols; c++) {
              table._cellXlRowspanCells.set(`${r},${c}`, true);
            }
          }
        }
        if (inTr) {
          inTd = true;
          currentCell = "";

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

          if (Object.keys(cellStyle).length > 0) {
            table._cellStyles.set(
              `${table._rows.length},${currentColIndex}`,
              cellStyle,
            );
          }
        }
        tdVisualCol = currentVisualCol;
        if (isSpacerCell) {
          currentVisualCol += effectiveColspan;
        } else {
          currentColIndex++;
          currentVisualCol++;
        }
      } else if ((tag === "th" || tag === "td") && isEnd && inTd) {
        inTd = false;
        if (inTr) {
          const cellText = currentCellIsSpacer ? "" : HtmlTable.extractText(currentCell);
          currentRow.push(cellText);
          currentRowColspans.push(currentCellColspan);
        }
        currentCellColspan = 1;
        currentCellXlRowspan = 1;
        currentCellIsSpacer = false;
      } else if ((inTd || inTh) && reader.text) {
        currentCell += reader.text;
      } else if ((inTd || inTh) && isStart && reader.tagName === "br") {
        currentCell += "\n";
      } else if ((inTd || inTh) && isStart && reader.tagName === "span") {
        const style = reader.getAttribute("style") || "";
        if (style.includes("mso-ignore:vglayout")) {
          const spanFull = reader.getFullElement();
          const imgMatch = spanFull.match(/<img[^>]+>/i);
          if (imgMatch) {
            const imgTag = imgMatch[0];
            const srcMatch = imgTag.match(/src=["']([^"']+)["']/i);
            const nameMatch = imgTag.match(/name=["']([^"']+)["']/i);
            const widthMatch = imgTag.match(/width=["']?(\d+)/i);
            const heightMatch = imgTag.match(/height=["']?(\d+)/i);
            const marginLeftMatch = style.match(/margin-left:\s*(-?\d+)px/i);
            const marginTopMatch = style.match(/margin-top:\s*(-?\d+)px/i);
            const zIndexMatch = style.match(/z-index:\s*(\d+)/i);

            const src = srcMatch ? srcMatch[1] : "";
            const name = nameMatch ? nameMatch[1] : "";
            const width = widthMatch ? parseInt(widthMatch[1]) : 64;
            const height = heightMatch ? parseInt(heightMatch[1]) : 64;
            const marginLeft = marginLeftMatch ? parseInt(marginLeftMatch[1]) : 0;
            const marginTop = marginTopMatch ? parseInt(marginTopMatch[1]) : 0;
            const zIndex = zIndexMatch ? parseInt(zIndexMatch[1]) : 0;

            table._pictures.push({
              row: table._rows.length,
              col: tdVisualCol,
              name,
              src,
              width,
              height,
              marginLeft,
              marginTop,
              zIndex,
            });

          }
        }
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

  static mergeTables(tables: HtmlTable[]): HtmlTable {
    const merged = new HtmlTable();
    merged._caption = tables[0]?._caption;
    merged._headers = tables[0]?._headers || [];
    merged._colWidths = tables[0]?._colWidths || [];

    for (const table of tables) {
      for (let i = 0; i < table._rows.length; i++) {
        merged._rows.push(table._rows[i]);
      }
      for (let i = 0; i < table._rowHeights.length; i++) {
        merged._rowHeights.push(table._rowHeights[i]);
      }
      for (let i = 0; i < table._rowHidden.length; i++) {
        merged._rowHidden.push(table._rowHidden[i]);
      }
      for (let i = 0; i < table._rowColspans.length; i++) {
        merged._rowColspans.push(table._rowColspans[i]);
      }
      for (let i = 0; i < table._rowXlRowspans.length; i++) {
        merged._rowXlRowspans.push(table._rowXlRowspans[i]);
      }
    }

    for (const table of tables) {
      for (const [key, val] of table._cellStyles) {
        merged._cellStyles.set(key, val);
      }
      for (const [key, val] of table._cellXlRowspans) {
        merged._cellXlRowspans.set(key, val);
      }
      for (const [key, val] of table._cellXlRowspanCells) {
        merged._cellXlRowspanCells.set(key, val);
      }
    }

    return merged;
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
