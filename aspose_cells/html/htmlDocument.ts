import { Workbook } from "../workbook";
import { Worksheet } from "../worksheet";
import { cellRef } from "../util";
import { HtmlTable, CellStyle } from "./htmlTable";
import { readFile } from "fs/promises";
import { dirname, join } from "path";
import type { CellValue } from "../types";

export interface HtmlParseOptions {
  trimWhitespace?: boolean;
  useFirstRowAsHeader?: boolean;
}

export class HtmlDocument {
  private _html: string = "";
  private _tables: HtmlTable[] = [];
  private _sheetNames: string[] = [];
  private _sheetFiles: string[] = [];
  private _cssClasses: Map<string, CellStyle> = new Map();

  static async load(filePath: string): Promise<HtmlDocument> {
    const content = await readFile(filePath, "utf-8");
    const doc = new HtmlDocument();
    doc._html = content;

    const dir = dirname(filePath);

    const tables = HtmlTable.parseAll(content);

    if (tables.length > 0) {
      doc._tables = tables;
    } else {
      const sheetFiles = doc.parseSheetReferences(content);
      for (const sheetFile of sheetFiles) {
        const sheetPath = join(dir, sheetFile);
        try {
          const sheetContent = await readFile(sheetPath, "utf-8");

          const stylesheetPath = sheetPath.replace(
            /sheet\d+\.htm$/,
            "stylesheet.css",
          );
          try {
            const cssContent = await readFile(stylesheetPath, "utf-8");
            doc._cssClasses = doc.parseStylesheet(cssContent);
          } catch (e) {}

          const sheetTables = HtmlTable.parseAll(sheetContent, doc._cssClasses);
          doc._tables.push(...sheetTables);
        } catch (e) {
          console.log("Could not load sheet file:", sheetFile);
        }
      }

      if (doc._tables.length === 0) {
        doc._tables = tables;
      }
    }

    return doc;
  }

  private parseStylesheet(css: string): Map<string, CellStyle> {
    const classes = new Map<string, CellStyle>();

    const classRegex = /\.(\w+)\s*\{([^}]+)\}/g;
    let match;

    while ((match = classRegex.exec(css)) !== null) {
      const className = match[1];
      const styles = match[2];
      const cellStyle: CellStyle = {};

      const parseStyleValue = (prop: string, value: string) => {
        if (
          prop === "background" ||
          prop === "background-color" ||
          prop === "mso-pattern"
        ) {
          if (value !== "auto" && value !== "none") {
            const colorMatch = value.match(/#[0-9A-Fa-f]{6}/);
            if (colorMatch) {
              cellStyle.bgcolor = colorMatch[0];
            }
          }
        } else if (prop === "color") {
          cellStyle.color = value;
        } else if (prop === "font-weight" && value === "bold") {
          cellStyle.font = { ...cellStyle.font, bold: true };
        } else if (prop === "font-style" && value === "italic") {
          cellStyle.font = { ...cellStyle.font, italic: true };
        } else if (prop === "font-size") {
          const sizeMatch = value.match(/(\d+)/);
          if (sizeMatch) {
            cellStyle.font = {
              ...cellStyle.font,
              size: parseInt(sizeMatch[1]),
            };
          }
        } else if (prop === "font-family") {
          cellStyle.font = {
            ...cellStyle.font,
            name: value.split(",")[0].trim(),
          };
        } else if (prop === "text-align") {
          cellStyle.align = value;
        } else if (prop === "vertical-align") {
          cellStyle.valign = value;
        } else if (prop === "border-top") {
          cellStyle.border = cellStyle.border || {};
          cellStyle.border.top = this.parseBorderValue(styles, "border-top");
        } else if (prop === "border-bottom") {
          cellStyle.border = cellStyle.border || {};
          cellStyle.border.bottom = this.parseBorderValue(
            styles,
            "border-bottom",
          );
        } else if (prop === "border-left") {
          cellStyle.border = cellStyle.border || {};
          cellStyle.border.left = this.parseBorderValue(styles, "border-left");
        } else if (prop === "border-right") {
          cellStyle.border = cellStyle.border || {};
          cellStyle.border.right = this.parseBorderValue(
            styles,
            "border-right",
          );
        }
      };

      const props = styles.split(";").filter((p) => p.trim());
      for (const prop of props) {
        const [key, ...valueParts] = prop.split(":");
        if (key && valueParts.length > 0) {
          parseStyleValue(key.trim(), valueParts.join(":").trim());
        }
      }

      if (Object.keys(cellStyle).length > 0) {
        classes.set(className, cellStyle);
      }
    }

    return classes;
  }

  private parseBorderValue(
    styles: string,
    side: string,
  ): { style?: string; color?: string } {
    const result: { style?: string; color?: string } = {};
    const borderRegex = new RegExp(`${side}:\\s*([^;]+)`, "i");
    const match = styles.match(borderRegex);
    if (match) {
      const parts = match[1].split(" ");
      for (const part of parts) {
        if (
          part === "solid" ||
          part === "dashed" ||
          part === "dotted" ||
          part === "thick"
        ) {
          result.style = part;
        } else if (
          part.startsWith("#") ||
          part === "windowtext" ||
          part === "black"
        ) {
          result.color =
            part === "windowtext" || part === "black" ? "#000000" : part;
        }
      }
    }
    return result;
  }

  private parseSheetReferences(html: string): string[] {
    const files: string[] = [];

    const linkMatches = html.match(
      /<link[^>]*href="([^"]*sheet\d+\.htm)"[^>]*>/gi,
    );
    if (linkMatches) {
      for (const match of linkMatches) {
        const hrefMatch = match.match(/href="([^"]*sheet\d+\.htm)"/i);
        if (hrefMatch) {
          files.push(hrefMatch[1]);
        }
      }
    }

    const scriptMatches = html.match(/c_rgszSh\[(\d+)\]="([^"]+)"/g);
    if (scriptMatches) {
      for (const match of scriptMatches) {
        const nameMatch = match.match(/c_rgszSh\[\d+\]="([^"]+)"/);
        if (nameMatch) {
          this._sheetNames.push(nameMatch[1]);
        }
      }
    }

    if (files.length === 0) {
      const fileListMatch = html.match(
        /<link[^>]*href="([^"]*filelist\.xml)"[^>]*>/i,
      );
      if (fileListMatch) {
        console.log("Found filelist reference:", fileListMatch[1]);
      }
    }

    return files;
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
    (workbook as any)._sheets.worksheets.length = 0;

    for (let i = 0; i < this._tables.length; i++) {
      const table = this._tables[i];
      const sheetName = this._sheetNames[i] || `Sheet${i + 1}`;
      const worksheet = workbook.addWorksheet(sheetName);
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
