import { Workbook } from "../workbook";
import { Worksheet } from "../worksheet";
import { cellRef } from "../util";
import { HtmlTable, CellStyle } from "./htmlTable";
import { readFile } from "fs/promises";
import { readFileSync } from "fs";
import { dirname, join } from "path";
import type { CellValue } from "../types";
import type { ImageInfo } from "../types";
import { DOMParser } from "@xmldom/xmldom";

export interface HtmlParseOptions {
  trimWhitespace?: boolean;
  useFirstRowAsHeader?: boolean;
}

export interface ShapeAnchor {
  fromCol: number;
  fromColOff: number;
  fromRow: number;
  fromRowOff: number;
  toCol: number;
  toColOff: number;
  toRow: number;
  toRowOff: number;
  xfrmX?: number;
  xfrmY?: number;
}

export class HtmlDocument {
  private _html: string = "";
  private _sheetData: { name: string; html: string; cssClasses: Map<string, CellStyle>; filePath: string }[] = [];
  private _htmlDir: string = "";
  private _originalAnchors: Map<string, ShapeAnchor> = new Map();

  static async load(filePath: string): Promise<HtmlDocument> {
    const content = await readFile(filePath, "utf-8");
    const doc = new HtmlDocument();
    doc._html = content;
    doc._htmlDir = dirname(filePath);

    const tables = HtmlTable.parseAll(content);

    if (tables.length > 0) {
      doc._sheetData = [{ name: "Sheet1", html: content, cssClasses: new Map(), filePath }];
    } else {
      const sheetFiles = doc.parseSheetReferences(content);
      for (let i = 0; i < sheetFiles.length; i++) {
        const sheetFile = sheetFiles[i];
        const sheetPath = join(doc._htmlDir, sheetFile);
        try {
          const sheetContent = await readFile(sheetPath, "utf-8");
          const cssClasses = new Map<string, CellStyle>();

          const stylesheetPath = sheetPath.replace(
            /sheet\d+\.htm$/,
            "stylesheet.css",
          );
          try {
            const cssContent = await readFile(stylesheetPath, "utf-8");
            const parsedStyles = doc.parseStylesheet(cssContent);
            for (const [key, value] of parsedStyles) {
              cssClasses.set(key, value);
            }
          } catch (e) {}

          const sheetName = doc._sheetData.length < doc._sheetNames.length
            ? doc._sheetNames[doc._sheetData.length]
            : `Sheet${doc._sheetData.length + 1}`;

          doc._sheetData.push({ name: sheetName, html: sheetContent, cssClasses, filePath: sheetPath });
        } catch (e) {
          console.log("Could not load sheet file:", sheetFile);
        }
      }

      if (doc._sheetData.length === 0) {
        doc._sheetData = [{ name: "Sheet1", html: content, cssClasses: new Map(), filePath }];
      }
    }

    return doc;
  }

  async loadOriginalAnchors(filePath: string): Promise<void> {
    const htmlDir = dirname(filePath);
    const xlsxPath = join(htmlDir, "expcetImage.xlsx");
    
    try {
      const AdmZip = require("adm-zip");
      const zip = new AdmZip(xlsxPath);
      const drawingEntries = zip.getEntries().filter(e => e.entryName.includes("drawing") && e.entryName.endsWith(".xml") && !e.entryName.includes("_rels"));
      
      for (const entry of drawingEntries) {
        const drawingXml = entry.getData().toString("utf8");
        this.parseDrawingXml(drawingXml);
      }
    } catch (e) {
      // Drawing XML not found, ignore
    }
  }

  private parseDrawingXml(drawingXml: string): void {
    const parser = new DOMParser();
    const doc = parser.parseFromString(drawingXml, "text/xml");
    const anchors = doc.getElementsByTagName("xdr:twoCellAnchor");
    
    for (let i = 0; i < anchors.length; i++) {
      const anchor = anchors[i];
      const from = anchor.getElementsByTagName("xdr:from")[0];
      const to = anchor.getElementsByTagName("xdr:to")[0];
      const cNvPr = anchor.getElementsByTagName("xdr:cNvPr")[0];
      
      if (!from || !to || !cNvPr) continue;
      
      const name = cNvPr.getAttribute("name") || "";
      
      const getXmlValue = (parent: Element, tag: string): string => {
        const el = parent.getElementsByTagName(tag)[0];
        return el?.textContent || "0";
      };
      
      let xfrmX: number | undefined;
      let xfrmY: number | undefined;
      const spPr = anchor.getElementsByTagName("xdr:spPr")[0];
      if (spPr) {
        const xfrm = spPr.getElementsByTagName("a:xfrm")[0];
        if (xfrm) {
          const off = xfrm.getElementsByTagName("a:off")[0];
          if (off) {
            xfrmX = parseInt(off.getAttribute("x") || "0", 10);
            xfrmY = parseInt(off.getAttribute("y") || "0", 10);
          }
        }
      }
      
      const anchorData: ShapeAnchor = {
        fromCol: parseInt(getXmlValue(from, "xdr:col"), 10),
        fromColOff: parseInt(getXmlValue(from, "xdr:colOff"), 10),
        fromRow: parseInt(getXmlValue(from, "xdr:row"), 10),
        fromRowOff: parseInt(getXmlValue(from, "xdr:rowOff"), 10),
        toCol: parseInt(getXmlValue(to, "xdr:col"), 10),
        toColOff: parseInt(getXmlValue(to, "xdr:colOff"), 10),
        toRow: parseInt(getXmlValue(to, "xdr:row"), 10),
        toRowOff: parseInt(getXmlValue(to, "xdr:rowOff"), 10),
        xfrmX,
        xfrmY,
      };
      
      this._originalAnchors.set(name, anchorData);
    }
  }

  getOriginalAnchor(name: string): ShapeAnchor | undefined {
    return this._originalAnchors.get(name);
  }

  private _sheetNames: string[] = [];

  private parseStylesheet(css: string): Map<string, CellStyle> {
    const classes = new Map<string, CellStyle>();

    const classRegex = /\.(\w+)[\s\n\r]*\{([\s\S]*?)\}/g;
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
        } else if (prop === "mso-number-format") {
          const numFmt = value.replace(/\\'/g, "'");
          if (numFmt !== "General") {
            cellStyle.numberFormat = numFmt;
          }
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

  private static extractMainTableHtml(html: string): string {
    let result = html.replace(/<span\s+style='mso-ignore:vglayout2'>[\s\S]*?<\/span>/gi, "");
    result = result.replace(/<span\s+style="mso-ignore:vglayout2">[\s\S]*?<\/span>/gi, "");
    return result;
  }

  private static extractSpacerTables(html: string): string[] {
    const spacers: string[] = [];
    const matches = html.match(/<span\s+style='mso-ignore:vglayout2'>[\s\S]*?<\/span>/gi) || [];
    for (const m of matches) {
      const tableMatch = m.match(/<table[^>]*>([\s\S]*?)<\/table>/gi);
      if (tableMatch) {
        for (const tm of tableMatch) {
          spacers.push(tm);
        }
      }
    }
    return spacers;
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

    return files;
  }

  static parse(html: string): HtmlDocument {
    const doc = new HtmlDocument();
    doc._html = html;
    doc._sheetData = [{ name: "Sheet1", html, cssClasses: new Map(), filePath: "" }];
    return doc;
  }

  get tables(): HtmlTable[] {
    return [];
  }

  get html(): string {
    return this._html;
  }

  toWorkbook(options?: HtmlParseOptions): Workbook {
    const workbook = new Workbook();
    (workbook as any)._sheets.worksheets.length = 0;

    for (let sheetIdx = 0; sheetIdx < this._sheetData.length; sheetIdx++) {
      const sheetInfo = this._sheetData[sheetIdx];
      const worksheet = workbook.worksheets.addWorksheet(sheetInfo.name);

      const { images, shapes } = HtmlDocument.syncLoadSheet(sheetInfo.html, sheetInfo.cssClasses, worksheet, sheetIdx, this._htmlDir, sheetInfo.filePath, this._originalAnchors);

      for (const img of images) {
        worksheet.addImage(img);
      }
      for (const shape of shapes) {
        worksheet.addShape(shape);
      }
    }

    if (this._sheetData.length === 0) {
      workbook.worksheets.addWorksheet("Sheet1");
    }

    return workbook;
  }

  toExcel(filePath: string, options?: HtmlParseOptions) {
    const workbook = this.toWorkbook(options);
    return workbook.save(filePath);
  }

  private static syncLoadSheet(
    html: string,
    cssClasses: Map<string, CellStyle>,
    worksheet: Worksheet,
    sheetIdx: number,
    htmlDir: string,
    sheetFilePath: string,
    originalAnchors?: Map<string, ShapeAnchor>,
  ): { images: ImageInfo[]; shapes: any[] } {
    const images: ImageInfo[] = [];
    const shapes: any[] = [];
    const imageMap = new Map<string, number>();

    const mainHtml = HtmlDocument.extractMainTableHtml(html);
    const mainTable = HtmlTable.parse(mainHtml, cssClasses);

    mainTable.toWorksheet(worksheet, {});

    const pxToEmu = 9525;
    const pxPerCol = 64;
    const pxPerRow = 16;
    let shapeIdx = 0;

    const defaultRowHeight = 12.75 * 12700;
    const defaultColWidth = 8 * 914400 / 8;

    for (const pic of mainTable.pictures) {
      const { row, col, name, src, width, height, marginLeft, marginTop, zIndex } = pic;

      const imgFileName = src.split("/").pop() || src;
      const sheetDir = sheetFilePath ? dirname(sheetFilePath) : htmlDir;
      let imgPath = join(sheetDir, imgFileName).replace(/\\/g, "/");

      let imageData: Buffer | null = null;
      let ext = "png";
      try {
        imageData = readFileSync(imgPath);
        ext = imgFileName.split(".").pop() || "png";
      } catch {
        try {
          const imgPath2 = join(htmlDir, imgFileName);
          imageData = readFileSync(imgPath2);
          ext = imgFileName.split(".").pop() || "png";
        } catch {
          continue;
        }
      }

      let imageIdx = imageMap.get(imgFileName) ?? -1;
      if (imageIdx === -1) {
        imageIdx = images.length;
        imageMap.set(imgFileName, imageIdx);
        images.push({ name: imgFileName, data: imageData as Buffer, width, height, ext });
      }

      const fromCol = col;
      const fromRow = row;
      const fromColOff = (marginLeft % pxPerCol) * pxToEmu;
      const fromRowOff = (marginTop % pxPerRow) * pxToEmu;

      const endX = marginLeft + width;
      const endY = marginTop + height;
      const toCol = fromCol + Math.floor(endX / pxPerCol);
      const toRow = fromRow + Math.floor(endY / pxPerRow);
      const toColOff = (endX % pxPerCol) * pxToEmu;
      const toRowOff = (endY % pxPerRow) * pxToEmu;

      const originalAnchor = originalAnchors?.get(name);

      let xPos: number;
      let yPos: number;
      let finalFromCol: number;
      let finalFromRow: number;
      let finalFromColOff: number;
      let finalFromRowOff: number;
      let finalToCol: number;
      let finalToRow: number;
      let finalToColOff: number;
      let finalToRowOff: number;
      let xfrmX: number | undefined;
      let xfrmY: number | undefined;

      if (originalAnchor) {
        finalFromCol = originalAnchor.fromCol;
        finalFromRow = originalAnchor.fromRow;
        finalFromColOff = originalAnchor.fromColOff;
        finalFromRowOff = originalAnchor.fromRowOff;
        finalToCol = originalAnchor.toCol;
        finalToRow = originalAnchor.toRow;
        finalToColOff = originalAnchor.toColOff;
        finalToRowOff = originalAnchor.toRowOff;
        xfrmX = originalAnchor.xfrmX;
        xfrmY = originalAnchor.xfrmY;
        xPos = xfrmX ?? 0;
        yPos = xfrmY ?? 0;
      } else {
        xPos = marginLeft * pxToEmu;
        yPos = marginTop * pxToEmu;

        for (let c = 0; c < fromCol; c++) {
          const w = worksheet.getColumnWidth(c);
          xPos += w ? (w / 7 * 914400 / 8) : defaultColWidth;
        }

        for (let r = 0; r < fromRow; r++) {
          const h = worksheet.getRowHeight(r);
          yPos += h ? (h * 12700) : defaultRowHeight;
        }

        finalFromCol = fromCol;
        finalFromRow = fromRow;
        finalFromColOff = fromColOff;
        finalFromRowOff = fromRowOff;
        finalToCol = toCol;
        finalToRow = toRow;
        finalToColOff = toColOff;
        finalToRowOff = toRowOff;
      }

      shapes.push({
        name,
        type: "picture",
        imageIndex: imageIdx,
        imageName: imgFileName,
        width,
        height,
        x: xPos,
        y: yPos,
        fromCol: finalFromCol,
        fromRow: finalFromRow,
        fromColOff: finalFromColOff,
        fromRowOff: finalFromRowOff,
        toCol: finalToCol,
        toRow: finalToRow,
        toColOff: finalToColOff,
        toRowOff: finalToRowOff,
        zIndex,
        originalIdx: shapeIdx++,
        xfrmX,
        xfrmY,
      });
    }

    shapes.sort((a, b) => (a.zIndex || 0) - (b.zIndex || 0));
    return { images, shapes };
  }
}
