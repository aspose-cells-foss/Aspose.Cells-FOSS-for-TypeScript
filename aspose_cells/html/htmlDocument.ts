import { Workbook } from "../workbook";
import { Worksheet } from "../worksheet";
import { cellRef } from "../util";
import { HtmlTable } from "./htmlTable";
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
          const sheetTables = HtmlTable.parseAll(sheetContent);
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
