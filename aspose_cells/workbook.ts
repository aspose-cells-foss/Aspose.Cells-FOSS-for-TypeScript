import { Worksheet } from "./worksheet";
import { WorksheetCollection } from "./worksheetCollection";
import { Cell } from "./cell";
import { DOMParser } from "@xmldom/xmldom";
import { CellValue, Style, SaveFormat } from "./types";
import {
  openZip as openZipUtil,
  readZipEntry as readZipEntryUtil,
  escapeXml,
  parseRange,
  cellRef,
  createZipWriter,
  addZipEntry,
  finalizeZip,
} from "./util";
import { writeFile, mkdir } from "fs/promises";
import { existsSync } from "fs";
import { dirname, join } from "path";
import {
  HtmlDocument,
  HtmlExporter,
  worksheetToHtml,
  workbookToHtmlFull,
} from "./html/index";

export class Workbook {
  private _sheets: WorksheetCollection;
  private _sharedStrings: string[] = [];
  private _protected = false;
  private _password?: string;

  constructor() {
    this._sheets = new WorksheetCollection();
  }

  static async load(filePath: string, password?: string): Promise<Workbook> {
    const workbook = new Workbook();
    workbook._password = password;

    const ext = filePath.toLowerCase().split(".").pop();
    if (ext === "html" || ext === "htm") {
      const htmlDoc = await HtmlDocument.load(filePath);
      const wb = htmlDoc.toWorkbook();
      workbook._sheets = wb._sheets;
      return workbook;
    }

    await workbook.loadFile(filePath);
    return workbook;
  }

  get worksheets(): Worksheet[] {
    return this._sheets.worksheets;
  }

  get worksheetCount(): number {
    return this._sheets.worksheetCount;
  }

  get worksheet(): Worksheet {
    return this._sheets.worksheet;
  }

  get styles(): Map<number, Style> {
    return this._sheets.styles;
  }

  get sharedStrings(): string[] {
    return this._sharedStrings;
  }

  get isProtected(): boolean {
    return this._protected;
  }

  getNumFmt(numFmtId: string | null): string {
    return this._sheets.getNumFmt(numFmtId);
  }

  addWorksheet(name?: string): Worksheet {
    return this._sheets.addWorksheet(name);
  }

  removeWorksheet(index: number): boolean {
    return this._sheets.removeWorksheet(index);
  }

  moveWorksheet(fromIndex: number, toIndex: number): boolean {
    return this._sheets.moveWorksheet(fromIndex, toIndex);
  }

  private async loadFile(filePath: string) {
    const zip = await openZipUtil(filePath);
    try {
      await this.loadSharedStrings(zip);
      await this.loadStyles(zip);
      await this.loadSheets(zip);
    } finally {
      await zip.close();
    }
  }

  private async loadSharedStrings(zip: any) {
    const content = await readZipEntryUtil(zip, "xl/sharedStrings.xml");
    if (!content) return;
    const parser = new DOMParser();
    const doc = parser.parseFromString(content, "text/xml");
    const strings = doc.getElementsByTagName("si");
    for (let i = 0; i < strings.length; i++) {
      const si = strings[i];
      const texts = si.getElementsByTagName("t");
      if (texts.length > 0) {
        this._sharedStrings.push(texts[0].textContent ?? "");
      } else {
        this._sharedStrings.push("");
      }
    }
  }

  private async loadStyles(zip: any) {
    const content = await readZipEntryUtil(zip, "xl/styles.xml");
    if (!content) return;
    const parser = new DOMParser();
    const doc = parser.parseFromString(content, "text/xml");

    const fonts = doc.getElementsByTagName("fonts")[0];
    const fontList = fonts?.getElementsByTagName("font") || [];

    const cellXfs = doc.getElementsByTagName("cellXfs")[0];
    if (!cellXfs) return;
    const xfs = cellXfs.getElementsByTagName("xf");

    for (let i = 0; i < xfs.length; i++) {
      const xf = xfs[i];
      const numFmtId = xf.getAttribute("numFmtId");
      const numFmt = this.getNumFmt(numFmtId);

      let font: any = { name: "Calibri", size: 11 };
      const fontId = xf.getAttribute("fontId");
      if (fontId !== null && fontList[parseInt(fontId)]) {
        const fontEl = fontList[parseInt(fontId)];
        const nameEl = fontEl.getElementsByTagName("name")[0];
        const szEl = fontEl.getElementsByTagName("sz")[0];
        const bEl = fontEl.getElementsByTagName("b")[0];
        const iEl = fontEl.getElementsByTagName("i")[0];
        const colorEl = fontEl.getElementsByTagName("color")[0];

        font = {
          name: nameEl?.getAttribute("val") || "Calibri",
          size: szEl ? parseInt(szEl.getAttribute("val") || "11") : 11,
          bold: bEl !== undefined,
          italic: iEl !== undefined,
          color: colorEl?.getAttribute("rgb") || colorEl?.getAttribute("theme"),
        };
      }

      const fills = doc.getElementsByTagName("fills")[0];
      const fillList = fills?.getElementsByTagName("fill") || [];
      let fill: any = {};
      const fillId = xf.getAttribute("fillId");
      if (fillId !== null && fillList[parseInt(fillId)]) {
        const fillEl = fillList[parseInt(fillId)];
        const patternFill = fillEl.getElementsByTagName("patternFill")[0];
        if (patternFill) {
          const patternType = patternFill.getAttribute("patternType");
          const fgColorEl = patternFill.getElementsByTagName("fgColor")[0];
          const bgColorEl = patternFill.getElementsByTagName("bgColor")[0];
          fill = {
            patternType: patternType || "none",
            fgColor: fgColorEl?.getAttribute("rgb"),
            bgColor: bgColorEl?.getAttribute("rgb"),
          };
        }
      }

      const borders = doc.getElementsByTagName("borders")[0];
      const borderList = borders?.getElementsByTagName("border") || [];
      let border: any = {};
      const borderId = xf.getAttribute("borderId");
      if (borderId !== null && borderList[parseInt(borderId)]) {
        const borderEl = borderList[parseInt(borderId)];
        const left = borderEl.getElementsByTagName("left")[0];
        const right = borderEl.getElementsByTagName("right")[0];
        const top = borderEl.getElementsByTagName("top")[0];
        const bottom = borderEl.getElementsByTagName("bottom")[0];

        const parseBorderSide = (el: Element | null): any => {
          if (!el) return {};
          return {
            style: el.getAttribute("style") || undefined,
            color: el.getElementsByTagName("color")[0]?.getAttribute("rgb"),
          };
        };

        border = {
          left: parseBorderSide(left),
          right: parseBorderSide(right),
          top: parseBorderSide(top),
          bottom: parseBorderSide(bottom),
        };
      }

      const alignment = xf.getElementsByTagName("alignment")[0];

      let textAlign = "general";
      let verticalAlign = "bottom";
      let wrapText = false;
      if (alignment) {
        textAlign = alignment.getAttribute("horizontal") || "general";
        verticalAlign = alignment.getAttribute("vertical") || "bottom";
        wrapText = alignment.getAttribute("wrapText") === "1";
      }

      this._sheets.setStyle(i, {
        font,
        fill,
        border,
        numberFormat: numFmt,
        alignment: {
          horizontal: textAlign as any,
          vertical: verticalAlign as any,
          wrapText,
        },
      });
    }
  }

  private async loadSheets(zip: any) {
    const content = await readZipEntryUtil(zip, "xl/workbook.xml");
    if (!content) return;
    const parser = new DOMParser();
    const doc = parser.parseFromString(content, "text/xml");
    const sheets = doc.getElementsByTagName("sheet");
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      const name = sheet.getAttribute("name") ?? `Sheet${i + 1}`;
      const sheetPath = sheet.getAttribute("r:id");
      if (!sheetPath) continue;

      const relsContent = await readZipEntryUtil(
        zip,
        "xl/_rels/workbook.xml.rels",
      );
      let target = "";
      if (relsContent) {
        const relParser = new DOMParser();
        const relDoc = relParser.parseFromString(relsContent, "text/xml");
        const rels = relDoc.getElementsByTagName("Relationship");
        for (let j = 0; j < rels.length; j++) {
          const rel = rels[j];
          if (rel.getAttribute("Id") === sheetPath) {
            target = rel.getAttribute("Target") ?? "";
            break;
          }
        }
      }

      const worksheet = new Worksheet(name, i);
      await this.loadWorksheetData(zip, worksheet, target);
      this._sheets.worksheets.push(worksheet);
    }
  }

  private async loadWorksheetData(
    zip: any,
    worksheet: Worksheet,
    path: string,
  ) {
    const content = await readZipEntryUtil(zip, `xl/${path}`);
    if (!content) return;

    const parser = new DOMParser();
    const doc = parser.parseFromString(content, "text/xml");

    const cols = doc.getElementsByTagName("cols")[0];
    if (cols) {
      const colElements = cols.getElementsByTagName("col");
      for (let i = 0; i < colElements.length; i++) {
        const col = colElements[i];
        const min = parseInt(col.getAttribute("min") ?? "1", 10) - 1;
        const max = parseInt(col.getAttribute("max") ?? "1", 10);
        const width = parseFloat(col.getAttribute("width") ?? "8.43");

        const htmlWidth = Math.round(width * 8);

        for (let c = min; c < max; c++) {
          worksheet.setColumnWidth(c, htmlWidth);
        }
      }
    }

    const rows = doc.getElementsByTagName("row");

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const rowNum = parseInt(row.getAttribute("r") ?? "0", 10) - 1;
      const rowHeight = row.getAttribute("ht");
      if (rowHeight) {
        worksheet.setRowHeight(rowNum, parseFloat(rowHeight));
      }
      const cells = row.getElementsByTagName("c");

      for (let j = 0; j < cells.length; j++) {
        const cell = cells[j];
        const ref = cell.getAttribute("r") ?? "";
        const { startRow: r, startCol: c } = parseRange(ref);
        const type = cell.getAttribute("t");
        const valueEl = cell.getElementsByTagName("v")[0];
        const inlineStr = cell.getElementsByTagName("is")[0];

        let value: CellValue = null;

        if (valueEl && valueEl.textContent) {
          if (type === "s") {
            const idx = parseInt(valueEl.textContent, 10);
            value = this._sharedStrings[idx] ?? "";
          } else if (type === "b") {
            value = valueEl.textContent === "1";
          } else if (type === "e") {
            value = valueEl.textContent;
          } else if (type === "str") {
            value = valueEl.textContent;
          } else {
            const num = parseFloat(valueEl.textContent);
            value = isNaN(num) ? valueEl.textContent : num;
          }
        } else if (inlineStr) {
          const t = inlineStr.getElementsByTagName("t")[0];
          value = t?.textContent ?? "";
        }

        const cellRef = worksheet.putValue(ref, value);
        const styleIdx = cell.getAttribute("s");
        if (styleIdx) {
          const styleIndex = parseInt(styleIdx, 10);
          const style = this._sheets.styles.get(styleIndex);
          if (style) {
            cellRef.setStyleIndex(styleIndex);
          }
        }
      }
    }
  }

  protect(protect: boolean, password?: string) {
    this._protected = protect;
    this._password = password;
  }

  async save(
    filePath: string,
    options?: { password?: string; saveFormat?: SaveFormat },
  ) {
    const format = options?.saveFormat || this.detectFormat(filePath);

    if (format === SaveFormat.HTML) {
      const dir = dirname(filePath);
      const fileName =
        filePath
          .replace(/\.[^.]+$/, "")
          .split(/[/\\]/)
          .pop() || "sheet";
      const htmlDir = join(dir, fileName + "_files");

      if (!existsSync(dir)) {
        await mkdir(dir, { recursive: true });
      }
      if (!existsSync(htmlDir)) {
        await mkdir(htmlDir, { recursive: true });
      }

      const exporter = new HtmlExporter(this);
      await exporter.saveAsHtmlFrameset(filePath, htmlDir, fileName);
      return;
    }

    const zip = await createZipWriter();
    await this.writeToZip(zip, options);
    const data = await finalizeZip(zip);
    await writeFile(filePath, Buffer.from(data));
  }

  private detectFormat(filePath: string): SaveFormat {
    const ext = filePath.toLowerCase().split(".").pop();
    if (ext === "html" || ext === "htm") return SaveFormat.HTML;
    if (ext === "csv") return SaveFormat.CSV;
    if (ext === "json") return SaveFormat.JSON;
    if (ext === "md" || ext === "markdown") return SaveFormat.MARKDOWN;
    return SaveFormat.XLSX;
  }

  toHtml(): string {
    return workbookToHtmlFull(this);
  }

  private async writeToZip(
    zip: any,
    options?: { password?: string; saveFormat?: SaveFormat },
  ) {
    this.generateSharedStrings();

    await addZipEntry(zip, "[Content_Types].xml", this.generateContentTypes());
    await addZipEntry(zip, "_rels/.rels", this.generateRels());
    await addZipEntry(zip, "docProps/core.xml", this.generateCore());
    await addZipEntry(zip, "docProps/app.xml", this.generateApp());
    await addZipEntry(zip, "xl/theme/theme1.xml", this.generateTheme());
    await addZipEntry(zip, "xl/workbook.xml", this.generateWorkbook());
    await addZipEntry(
      zip,
      "xl/_rels/workbook.xml.rels",
      this.generateWorkbookRels(),
    );
    await addZipEntry(zip, "xl/styles.xml", this.generateStyles());

    const sharedStringsXml = this.generateSharedStrings();
    if (sharedStringsXml) {
      await addZipEntry(zip, "xl/sharedStrings.xml", sharedStringsXml);
    }

    for (let i = 0; i < this._sheets.worksheets.length; i++) {
      const sheet = this._sheets.worksheets[i];
      await addZipEntry(zip, `xl/worksheets/sheet${i + 1}.xml`, sheet.getXml());
    }
  }

  private generateRels(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml" /><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml" /><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" /></Relationships>`;
  }

  private generateCore(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title /><dc:subject /><dc:creator /><cp:keywords /><dc:description /><cp:lastModifiedBy /><cp:category /></cp:coreProperties>`;
  }
  private generateApp(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><AppVersion>14.0300</AppVersion><DocSecurity>0</DocSecurity><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts><Template /><Manager /><Company /></Properties>`;
  }

  private generateTheme(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000" /></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF" /></a:lt1><a:dk2><a:srgbClr val="44546A" /></a:dk2><a:lt2><a:srgbClr val="E7E6E6" /></a:lt2><a:accent1><a:srgbClr val="4472C4" /></a:accent1><a:accent2><a:srgbClr val="ED7D31" /></a:accent2><a:accent3><a:srgbClr val="A5A5A5" /></a:accent3><a:accent4><a:srgbClr val="FFC000" /></a:accent4><a:accent5><a:srgbClr val="5B9BD5" /></a:accent5><a:accent6><a:srgbClr val="70AD47" /></a:accent6><a:hlink><a:srgbClr val="0563C1" /></a:hlink><a:folHlink><a:srgbClr val="954F72" /></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204" /><a:ea typeface="" /><a:cs typeface="" /><a:font script="Jpan" typeface="游ゴシック Light" /><a:font script="Hang" typeface="맑은 고딕" /><a:font script="Hans" typeface="等线 Light" /><a:font script="Hant" typeface="新細明體" /><a:font script="Arab" typeface="Times New Roman" /><a:font script="Hebr" typeface="Times New Roman" /><a:font script="Thai" typeface="Tahoma" /><a:font script="Ethi" typeface="Nyala" /><a:font script="Beng" typeface="Vrinda" /><a:font script="Gujr" typeface="Shruti" /><a:font script="Khmr" typeface="MoolBoran" /><a:font script="Knda" typeface="Tunga" /><a:font script="Guru" typeface="Raavi" /><a:font script="Cans" typeface="Euphemia" /><a:font script="Cher" typeface="Plantagenet Cherokee" /><a:font script="Yiii" typeface="Microsoft Yi Baiti" /><a:font script="Tibt" typeface="Microsoft Himalaya" /><a:font script="Thaa" typeface="MV Boli" /><a:font script="Deva" typeface="Mangal" /><a:font script="Telu" typeface="Gautami" /><a:font script="Taml" typeface="Latha" /><a:font script="Syrc" typeface="Estrangelo Edessa" /><a:font script="Orya" typeface="Kalinga" /><a:font script="Mlym" typeface="Kartika" /><a:font script="Laoo" typeface="DokChampa" /><a:font script="Sinh" typeface="Iskoola Pota" /><a:font script="Mong" typeface="Mongolian Baiti" /><a:font script="Viet" typeface="Times New Roman" /><a:font script="Uigh" typeface="Microsoft Uighur" /><a:font script="Geor" typeface="Sylfaen" /><a:font script="Armn" typeface="Arial" /><a:font script="Bugi" typeface="Leelawadee UI" /><a:font script="Bopo" typeface="Microsoft JhengHei" /><a:font script="Java" typeface="Javanese Text" /><a:font script="Lisu" typeface="Segoe UI" /><a:font script="Mymr" typeface="Myanmar Text" /><a:font script="Nkoo" typeface="Ebrima" /><a:font script="Olck" typeface="Nirmala UI" /><a:font script="Osma" typeface="Ebrima" /><a:font script="Phag" typeface="Phagspa" /><a:font script="Syrn" typeface="Estrangelo Edessa" /><a:font script="Syrj" typeface="Estrangelo Edessa" /><a:font script="Syre" typeface="Estrangelo Edessa" /><a:font script="Sora" typeface="Nirmala UI" /><a:font script="Tale" typeface="Microsoft Tai Le" /><a:font script="Talu" typeface="Microsoft New Tai Lue" /><a:font script="Tfng" typeface="Ebrima" /></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204" /><a:ea typeface="" /><a:cs typeface="" /><a:font script="Jpan" typeface="游ゴシック" /><a:font script="Hang" typeface="맑은 고딕" /><a:font script="Hans" typeface="等线" /><a:font script="Hant" typeface="新細明體" /><a:font script="Arab" typeface="Arial" /><a:font script="Hebr" typeface="Arial" /><a:font script="Thai" typeface="Tahoma" /><a:font script="Ethi" typeface="Nyala" /><a:font script="Beng" typeface="Vrinda" /><a:font script="Gujr" typeface="Shruti" /><a:font script="Khmr" typeface="DaunPenh" /><a:font script="Knda" typeface="Tunga" /><a:font script="Guru" typeface="Raavi" /><a:font script="Cans" typeface="Euphemia" /><a:font script="Cher" typeface="Plantagenet Cherokee" /><a:font script="Yiii" typeface="Microsoft Yi Baiti" /><a:font script="Tibt" typeface="Microsoft Himalaya" /><a:font script="Thaa" typeface="MV Boli" /><a:font script="Deva" typeface="Mangal" /><a:font script="Telu" typeface="Gautami" /><a:font script="Taml" typeface="Latha" /><a:font script="Syrc" typeface="Estrangelo Edessa" /><a:font script="Orya" typeface="Kalinga" /><a:font script="Mlym" typeface="Kartika" /><a:font script="Laoo" typeface="DokChampa" /><a:font script="Sinh" typeface="Iskoola Pota" /><a:font script="Mong" typeface="Mongolian Baiti" /><a:font script="Viet" typeface="Arial" /><a:font script="Uigh" typeface="Microsoft Uighur" /><a:font script="Geor" typeface="Sylfaen" /><a:font script="Armn" typeface="Arial" /><a:font script="Bugi" typeface="Leelawadee UI" /><a:font script="Bopo" typeface="Microsoft JhengHei" /><a:font script="Java" typeface="Javanese Text" /><a:font script="Lisu" typeface="Segoe UI" /><a:font script="Mymr" typeface="Myanmar Text" /><a:font script="Nkoo" typeface="Ebrima" /><a:font script="Olck" typeface="Nirmala UI" /><a:font script="Osma" typeface="Ebrima" /><a:font script="Phag" typeface="Phagspa" /><a:font script="Syrn" typeface="Estrangelo Edessa" /><a:font script="Syrj" typeface="Estrangelo Edessa" /><a:font script="Syre" typeface="Estrangelo Edessa" /><a:font script="Sora" typeface="Nirmala UI" /><a:font script="Tale" typeface="Microsoft Tai Le" /><a:font script="Talu" typeface="Microsoft New Tai Lue" /><a:font script="Tfng" typeface="Ebrima" /></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr" /></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000" /><a:satMod val="105000" /><a:tint val="67000" /></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000" /><a:satMod val="103000" /><a:tint val="73000" /></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000" /><a:satMod val="109000" /><a:tint val="81000" /></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0" /></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000" /><a:lumMod val="102000" /><a:tint val="94000" /></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000" /><a:lumMod val="100000" /><a:shade val="100000" /></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000" /><a:satMod val="120000" /><a:shade val="78000" /></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0" /></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr" /></a:solidFill><a:prstDash val="solid" /><a:miter lim="800000" /></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr" /></a:solidFill><a:prstDash val="solid" /><a:miter lim="800000" /></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr" /></a:solidFill><a:prstDash val="solid" /><a:miter lim="800000" /></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst /></a:effectStyle><a:effectStyle><a:effectLst /></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000" /></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr" /></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000" /><a:satMod val="170000" /></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000" /><a:satMod val="150000" /><a:shade val="98000" /><a:lumMod val="102000" /></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000" /><a:satMod val="130000" /><a:shade val="90000" /><a:lumMod val="103000" /></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000" /><a:satMod val="120000" /></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0" /></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults /><a:extraClrSchemeLst /><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" /></a:ext></a:extLst></a:theme>`;
  }

  private generateContentTypes(): string {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" /><Default Extension="xml" ContentType="application/xml" /><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" /><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml" /><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" /><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml" /><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />`;

    // Add each worksheet override
    for (let i = 0; i < this._sheets.worksheets.length; i++) {
      xml += `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />`;
    }

    // Add shared strings part if needed
    if (this._sharedStrings.length > 0) {
      xml += `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" />`;
    }

    xml += "</Types>";
    return xml;
  }

  private generateWorkbook(): string {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"><fileVersion appName="xl" lastEdited="7" lowestEdited="4" rupBuild="29725"/><workbookPr/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x15"><x15ac:absPath url="E:\\FOSS\\aspose.cells-foss-for-typescript\\" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/></mc:Choice></mc:AlternateContent><xr:revisionPtr revIDLastSave="0" documentId="13_ncr:1_{43418319-D520-46C2-81F6-417174F36FCE}" xr6:coauthVersionLast="47" xr6:coauthVersionMax="47" xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}"/><bookViews><workbookView xWindow="-110" yWindow="-110" windowWidth="25820" windowHeight="15500" activeTab="1" xr2:uid="{00000000-000D-0000-FFFF-FFFF00000000}"/></bookViews><sheets>`;

    for (let i = 0; i < this._sheets.worksheets.length; i++) {
      const sheet = this._sheets.worksheets[i];
      const rid = i + 1;
      xml += `<sheet name="${escapeXml(sheet.name)}" sheetId="${i + 1}" r:id="rId${rid}" />`;
    }

    xml += `</sheets><calcPr calcId="0"/></workbook>`;
    return xml;
  }

  private generateWorkbookRels(): string {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;

    let rid = 1;

    for (let i = 0; i < this._sheets.worksheets.length; i++) {
      xml += `<Relationship Id="rId${rid++}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml" />`;
    }

    xml += `<Relationship Id="rId${rid++}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />`;

    if (this._sharedStrings.length > 0) {
      xml += `<Relationship Id="rId${rid++}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml" />`;
    }

    xml += `<Relationship Id="rId${rid++}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml" /></Relationships>`;
    return xml;
  }

  private generateStyles(): string {
    const styles = this._sheets.styles;

    const fonts = new Map<number, any>();
    const fills = new Map<number, any>();
    const borders = new Map<number, any>();
    const numFmts = new Map<number, string>();

    let fontId = 0;
    let fillId = 0;
    let borderId = 0;
    let numFmtId = 164;

    for (const [idx, style] of styles) {
      if (style.font) {
        if (!fonts.has(fontId)) {
          fonts.set(fontId++, style.font);
        }
      }
      if (style.fill?.patternType && style.fill.patternType !== "none") {
        if (!fills.has(fillId)) {
          fills.set(fillId++, style.fill);
        }
      }
      if (style.border) {
        if (!borders.has(borderId)) {
          borders.set(borderId++, style.border);
        }
      }
      if (
        style.numberFormat &&
        ![
          "General",
          "0",
          "0.00",
          "#,##0",
          "#,##0.00",
          "0%",
          "0.00%",
          "0.00E+00",
        ].includes(style.numberFormat)
      ) {
        numFmts.set(numFmtId++, style.numberFormat);
      }
    }

    let fontsXml = `<fonts count="${Math.max(1, fonts.size)}">`;
    fontsXml += `<font><sz val="10" /><color theme="1" /><name val="Arial" /><family val="2" /></font>`;
    for (const [id, font] of fonts) {
      const size = font.size || 10;
      const name = font.name || "Arial";
      const color = font.color ? ` rgb="${font.color}"` : ' theme="1"';
      const bold = font.bold ? "<b />" : "";
      const italic = font.italic ? "<i />" : "";
      fontsXml += `<font><sz val="${size}" /><color${color} /><name val="${name}" /><family val="2" />${bold}${italic}</font>`;
    }
    fontsXml += "</fonts>";

    let fillsXml = `<fills count="${Math.max(2, fills.size + 2)}">`;
    fillsXml += `<fill><patternFill patternType="none" /></fill>`;
    fillsXml += `<fill><patternFill patternType="gray125" /></fill>`;
    for (const [id, fill] of fills) {
      const fgColor = fill.fgColor ? ` rgb="${fill.fgColor}"` : "";
      const patternType = fill.patternType || "solid";
      fillsXml += `<fill><patternFill patternType="${patternType}"><fgColor${fgColor} /></patternFill></fill>`;
    }
    fillsXml += "</fills>";

    let bordersXml = `<borders count="${Math.max(1, borders.size + 1)}">`;
    bordersXml += `<border><left /><right /><top /><bottom /><diagonal /></border>`;
    for (const [id, border] of borders) {
      const parseSide = (side: any) => {
        if (!side?.style) return "<left />";
        const style =
          side.style === "thick"
            ? 'style="thick"'
            : side.style === "medium"
              ? 'style="medium"'
              : 'style="thin"';
        const color = side.color ? ` rgb="${side.color}"` : ' auto="1"';
        return `<left ${style}><color${color} /></left>`;
      };
      bordersXml += `<border>${parseSide(border.left)}${parseSide(border.right)}${parseSide(border.top)}${parseSide(border.bottom)}<diagonal /></border>`;
    }
    bordersXml += "</borders>";

    let numFmtsXml = "";
    if (numFmts.size > 0) {
      numFmtsXml = `<numFmts count="${numFmts.size}">`;
      for (const [id, fmt] of numFmts) {
        const code = fmt.replace(/\\/g, "\\\\");
        numFmtsXml += `<numFmt numFmtId="${id}" formatCode="${code}" />`;
      }
      numFmtsXml += "</numFmts>";
    }

    let cellXfsXml = `<cellXfs count="${styles.size || 1}">`;
    for (const [idx, style] of styles) {
      const fontRef = style.font ? 1 : 0;
      const fillRef =
        style.fill?.patternType && style.fill.patternType !== "none" ? 2 : 0;
      const borderRef = style.border ? 1 : 0;

      let applyAlignment = "";
      let alignmentXml = "";
      if (style.alignment) {
        const h = style.alignment.horizontal
          ? ` horizontal="${style.alignment.horizontal}"`
          : "";
        const v = style.alignment.vertical
          ? ` vertical="${style.alignment.vertical}"`
          : "";
        const wrap = style.alignment.wrapText ? ' wrapText="1"' : "";
        if (h || v || wrap) {
          applyAlignment = ' applyAlignment="1"';
          alignmentXml = `<alignment${h}${v}${wrap} />`;
        }
      }

      cellXfsXml += `<xf numFmtId="0" fontId="${fontRef}" fillId="${fillRef}" borderId="${borderRef}" xfId="0"${applyAlignment}>${alignmentXml}</xf>`;
    }
    if (styles.size === 0) {
      cellXfsXml += `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />`;
    }
    cellXfsXml += "</cellXfs>";

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">${numFmtsXml}${fontsXml}${fillsXml}${bordersXml}<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>${cellXfsXml}<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext><ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>`;
  }

  private generateSharedStrings(): string {
    const strings = new Set<string>();
    for (const sheet of this._sheets.worksheets) {
      for (const cell of sheet.cells) {
        if (typeof cell.value === "string") {
          strings.add(cell.value);
        }
      }
    }

    const uniqueStrings = Array.from(strings);
    // Populate the workbook's shared strings array for later use (e.g., relations)
    this._sharedStrings = uniqueStrings;
    // Update Cell static map for string lookup during XML generation
    // @ts-ignore – using static method defined in Cell
    (Cell as any).setSharedStrings(uniqueStrings);

    if (uniqueStrings.length === 0) return "";
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${uniqueStrings.length}" uniqueCount="${uniqueStrings.length}">`;

    for (const str of uniqueStrings) {
      xml += `<si><t>${escapeXml(str)}</t></si>`;
    }

    xml += "</sst>";
    return xml;
  }

  toCsv(): string {
    const lines: string[] = [];
    const sheet = this._sheets.worksheets[0];
    if (!sheet) return "";

    const maxRow = Math.max(...Array.from(sheet.cells).map((c) => c.row), 0);
    const maxCol = Math.max(...Array.from(sheet.cells).map((c) => c.col), 0);

    for (let r = 0; r <= maxRow; r++) {
      const row: string[] = [];
      for (let c = 0; c <= maxCol; c++) {
        const cell = sheet.getCell(r, c);
        const value = cell?.value;
        if (value === null || value === undefined) {
          row.push("");
        } else if (typeof value === "string") {
          row.push(`"${value.replace(/"/g, '""')}"`);
        } else if (typeof value === "boolean") {
          row.push(value ? "TRUE" : "FALSE");
        } else if (value instanceof Date) {
          row.push(value.toISOString());
        } else {
          row.push(String(value));
        }
      }
      lines.push(row.join(","));
    }

    return lines.join("\n");
  }

  toJson(): string {
    const data: { sheetName: string; cells: { [key: string]: CellValue } }[] =
      [];

    for (const sheet of this._sheets.worksheets) {
      const sheetData: { [key: string]: CellValue } = {};
      for (const cell of sheet.cells) {
        sheetData[cell.ref] = cell.value;
      }
      data.push({ sheetName: sheet.name, cells: sheetData });
    }

    return JSON.stringify(data, null, 2);
  }

  toMarkdown(): string {
    const lines: string[] = [];
    const sheet = this._sheets.worksheets[0];
    if (!sheet) return "";

    const maxRow = Math.max(...Array.from(sheet.cells).map((c) => c.row), 0);
    const maxCol = Math.max(...Array.from(sheet.cells).map((c) => c.col), 0);

    for (let r = 0; r <= maxRow; r++) {
      const row: string[] = [];
      for (let c = 0; c <= maxCol; c++) {
        const cell = sheet.getCell(r, c);
        const value = cell?.value ?? "";
        row.push(String(value));
      }
      lines.push("| " + row.join(" | ") + " |");
    }

    return lines.join("\n");
  }
}
