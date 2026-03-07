import { Workbook } from "../workbook";
import { Worksheet } from "../worksheet";
import { cellRef } from "../util";
import { HtmlTable } from "./htmlTable";
import type { CellValue } from "../types";

export interface HtmlParseOptions {
  trimWhitespace?: boolean;
  useFirstRowAsHeader?: boolean;
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
    (workbook as any)._sheets.worksheets.length = 0;

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
