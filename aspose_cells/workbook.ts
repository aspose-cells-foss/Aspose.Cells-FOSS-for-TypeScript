import { Worksheet } from "./worksheet";
import { WorksheetCollection } from "./worksheetCollection";
import { Cell } from "./cell";
import { DOMParser } from "@xmldom/xmldom";
import {
  CellValue,
  Style,
  SaveFormat,
  ShapeInfo,
  ShapeFill,
  ChartInfo,
  ImageInfo,
} from "./types";
import {
  openZip as openZipUtil,
  readZipEntry as readZipEntryUtil,
  escapeXml,
  parseRange,
  cellRef,
  createZipWriter,
  addZipEntry,
  finalizeZip,
  fixZipFile,
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
import { ChartLoader } from "./chartLoader";
import { ImpDrawing } from "./impDrawing";
import { ExpDrawing } from "./expDrawing";
import { DataRelationship } from "./dataRelationship";

export class Workbook {
  private _sheets: WorksheetCollection;
  private _sharedStrings: string[] = [];
  private _protected = false;
  private _password?: string;
  private _charts: ChartInfo[] = [];
  private _activeTab = 0;
  private _images: ImageInfo[] = [];

  constructor(createDefault: boolean = true) {
    this._sheets = new WorksheetCollection(createDefault);
  }

  static async load(filePath: string, password?: string): Promise<Workbook> {
    const workbook = new Workbook(false);
    workbook._password = password;

    const ext = filePath.toLowerCase().split(".").pop();
    if (ext === "html" || ext === "htm") {
      const htmlDoc = await HtmlDocument.load(filePath);
      await htmlDoc.loadOriginalAnchors(filePath);
      const wb = htmlDoc.toWorkbook();
      workbook._sheets = wb._sheets;
      return workbook;
    }

    await workbook.loadFile(filePath);
    return workbook;
  }

  get worksheets(): WorksheetCollection {
    return this._sheets;
  }

  get sharedStrings(): string[] {
    return this._sharedStrings;
  }

  get isProtected(): boolean {
    return this._protected;
  }

  get charts(): ChartInfo[] {
    return this._charts;
  }

  get images(): ImageInfo[] {
    return this._images;
  }

  get activeTab(): number {
    return this._activeTab;
  }

  set activeTab(value: number) {
    this._activeTab = value;
  }

  getNumFmt(numFmtId: string | null): string {
    return this._sheets.getNumFmt(numFmtId);
  }

  private async loadFile(filePath: string) {
    const zip = await openZipUtil(filePath);
    try {
      await this.loadSharedStrings(zip);
      await this.loadStyles(zip);
      await this.loadSheets(zip);
      await this.loadWorkbookViews(zip);
    } finally {
      await zip.close();
    }

    const chartLoader = new ChartLoader();
    this._charts = await chartLoader.loadCharts(filePath);

    for (const chart of this._charts) {
      const worksheet = this.worksheets.get(chart.sheetIndex);
      if (worksheet) {
        const exists = worksheet.shapes.some(s => 
          (s as any).chartRefId && (s as any).chartRefId === (chart as any).chartRefId
        );
        if (!exists) {
          worksheet.addShape(chart);
        }
      }
    }
  }

  private async loadWorkbookViews(zip: any) {
    const content = await readZipEntryUtil(zip, "xl/workbook.xml");
    if (!content) return;
    const parser = new DOMParser();
    const doc = parser.parseFromString(content, "text/xml");
    const workbookView = doc.getElementsByTagName("workbookView")[0];
    if (workbookView) {
      const activeTab = workbookView.getAttribute("activeTab");
      if (activeTab) {
        this._activeTab = parseInt(activeTab, 10);
      }
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
      await ImpDrawing.loadWorksheetDrawing(zip, worksheet.shapes, target);
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

        const htmlWidth = Math.round(width * 7);

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
    const fixedData = fixZipFile(data);
    await writeFile(filePath, Buffer.from(fixedData));
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
    this.collectCellStyles();

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

    const imageMap = new Map<string, number>();
    const sheetImageMaps: Map<string, number>[] = [];

    for (const sheet of this._sheets.worksheets) {
      const sheetMap = new Map<string, number>();
      sheetImageMaps.push(sheetMap);
      for (const img of sheet.images) {
        if (!imageMap.has(img.name)) {
          imageMap.set(img.name, this._images.length);
          this._images.push(img);
        }
        if (!sheetMap.has(img.name)) {
          sheetMap.set(img.name, sheetMap.size + 1);
        }
      }
    }

    for (let i = 0; i < this._images.length; i++) {
      const img = this._images[i];
      const ext = img.ext || "png";
      await addZipEntry(zip, `xl/media/image${i + 1}.${ext}`, img.data);
    }

    let hasDrawings = false;
    let drawingIndex = 0;
    for (const sheet of this._sheets.worksheets) {
      if (sheet.shapes.length > 0) {
        hasDrawings = true;
        drawingIndex++;
        sheet._drawingIndex = drawingIndex;
      } else {
        sheet._drawingIndex = 0;
      }
    }

    for (let i = 0; i < this._sheets.worksheets.length; i++) {
      const sheet = this._sheets.worksheets[i];
      const sheetImgMap = sheetImageMaps[i] || new Map<string, number>();

      await addZipEntry(
        zip,
        `xl/worksheets/sheet${i + 1}.xml`,
        sheet.getXml(sheet._drawingIndex),
      );

      if (sheet.shapes.length > 0) {
        this.calculateShapePositions(sheet);

        await addZipEntry(
          zip,
          `xl/worksheets/_rels/sheet${i + 1}.xml.rels`,
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${sheet._drawingIndex}.xml" /></Relationships>`,
        );

        const drawingXml = ExpDrawing.generateDrawingXml(sheet.shapes, sheet.images, imageMap);
        await addZipEntry(
          zip,
          `xl/drawings/drawing${sheet._drawingIndex}.xml`,
          drawingXml,
        );

        const drawingRelsXml = this.generateDrawingRelsForSheet(sheet, imageMap);
        await addZipEntry(
          zip,
          `xl/drawings/_rels/drawing${sheet._drawingIndex}.xml.rels`,
          drawingRelsXml,
        );
      }
    }
    
    let chartIndex = 1;
    for (const sheet of this._sheets.worksheets) {
      for (const shape of sheet.shapes) {
        if ((shape as any).type === "chart" || (shape as any).chartType) {
          const chart = shape as ChartInfo;
          const chartXml = chart.originalChartXml || this.generateChartXml(chart);
          await addZipEntry(zip, `xl/charts/chart${chartIndex}.xml`, chartXml);
          if (chart.originalChartXml) {
            const chartRelsXml = this.generateChartRels(chartIndex);
            await addZipEntry(zip, `xl/charts/_rels/chart${chartIndex}.xml.rels`, chartRelsXml);
            if (chart.originalStyleXml) {
              await addZipEntry(zip, `xl/charts/style${chartIndex}.xml`, chart.originalStyleXml);
            }
            if (chart.originalColorsXml) {
              await addZipEntry(zip, `xl/charts/colors${chartIndex}.xml`, chart.originalColorsXml);
            }
          }
          chartIndex++;
        }
      }
    }
  }

  private generateChartRels(chartIndex: number): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors${chartIndex}.xml"/><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style${chartIndex}.xml"/></Relationships>`;
  }

  private generateChartStyle(chartIndex: number): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="201"><cs:axisTitle><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="1000" kern="1200"/></cs:axisTitle><cs:categoryAxis><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr><cs:defRPr sz="900" kern="1200"/></cs:categoryAxis><cs:chartArea><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr><cs:defRPr sz="1000" kern="1200"/></cs:chartArea><cs:dataLabel><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="900" kern="1200"/></cs:dataLabel><cs:dataPoint><cs:lnRef idx="0"/><cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:dataPoint><cs:legend><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="900" kern="1200"/></cs:legend><cs:plotArea><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:plotArea><cs:title><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="1400" b="0" kern="1200" spc="0" baseline="0"/></cs:title><cs:valueAxis><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="900" kern="1200"/></cs:valueAxis></cs:chartStyle>`;
  }

  private generateChartColors(chartIndex: number): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10"><a:schemeClr val="accent1"/><a:schemeClr val="accent2"/><a:schemeClr val="accent3"/><a:schemeClr val="accent4"/><a:schemeClr val="accent5"/><a:schemeClr val="accent6"/><cs:variation/><cs:variation><a:lumMod val="60000"/></cs:variation><cs:variation><a:lumMod val="80000"/><a:lumOff val="20000"/></cs:variation><cs:variation><a:lumMod val="80000"/></cs:variation><cs:variation><a:lumMod val="60000"/><a:lumOff val="40000"/></cs:variation><cs:variation><a:lumMod val="50000"/></cs:variation><cs:variation><a:lumMod val="70000"/><a:lumOff val="30000"/></cs:variation><cs:variation><a:lumMod val="70000"/></cs:variation><cs:variation><a:lumMod val="50000"/><a:lumOff val="50000"/></cs:variation></cs:colorStyle>`;
  }
  
  private generateChartXml(chart: ChartInfo): string {
    const chartType = chart.chartType || chart.type || "column";
    const axId1 = "55530882";
    const axId2 = "30015890";
    const sheetName = this._sheets.worksheets[chart.sheetIndex]?.name || "Sheet1";
    const legendPos = chart.legendPosition === "left" ? "l" : 
                      chart.legendPosition === "right" ? "r" : 
                      chart.legendPosition === "top" ? "t" : "b";

    let seriesXml = "";
    for (let i = 0; i < chart.series.length; i++) {
      const ser = chart.series[i];
      const valCol = String.fromCharCode(66 + i);
      seriesXml += `<c:ser><c:idx val="${i}" /><c:order val="${i}" /><c:tx><c:strRef><c:f>${sheetName}!$${valCol}$1</c:f></c:strRef></c:tx><c:invertIfNegative val="0" /><c:cat><c:strRef><c:f>${sheetName}!$A$2:$A$6</c:f><c:strCache><c:ptCount val="5" /></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>${sheetName}!$${valCol}$2:$${valCol}$6</c:f><c:numCache><c:ptCount val="5" /></c:numCache></c:numRef></c:val></c:ser>`;
    }

    let chartTypeXml = "";
    switch (chartType) {
      case "bar":
      case "column":
      case "barStacked":
      case "columnStacked":
      case "barStacked100":
      case "columnStacked100":
        const barDir = chartType === "bar" || chartType === "barStacked" || chartType === "barStacked100" ? "bar" : "col";
        chartTypeXml = `<c:barChart><c:barDir val="${barDir}" /><c:grouping val="clustered" /><c:varyColors val="0" />${seriesXml}<c:axId val="${axId1}" /><c:axId val="${axId2}" /></c:barChart>`;
        break;
      case "line":
      case "lineStacked":
      case "lineStacked100":
        chartTypeXml = `<c:lineChart><c:grouping val="clustered" />${seriesXml}<c:axId val="${axId1}" /><c:axId val="${axId2}" /></c:lineChart>`;
        break;
      case "pie":
      case "doughnut":
        chartTypeXml = `<c:pieChart><c:varyColors val="0" />${seriesXml}</c:pieChart>`;
        break;
      case "area":
      case "areaStacked":
      case "areaStacked100":
        chartTypeXml = `<c:areaChart><c:grouping val="clustered" />${seriesXml}<c:axId val="${axId1}" /><c:axId val="${axId2}" /></c:areaChart>`;
        break;
      case "scatter":
        chartTypeXml = `<c:scatterChart><c:scatterStyle val="lineMarker" />${seriesXml}<c:axId val="${axId1}" /><c:axId val="${axId2}" /></c:scatterChart>`;
        break;
      case "radar":
        chartTypeXml = `<c:radarChart><c:radarStyle val="marker" />${seriesXml}</c:radarChart>`;
        break;
      default:
        chartTypeXml = `<c:barChart><c:barDir val="col" /><c:grouping val="clustered" /><c:varyColors val="0" />${seriesXml}<c:axId val="${axId1}" /><c:axId val="${axId2}" /></c:barChart>`;
    }

    const title = chart.title || "";
    const titleXml = title ? `<c:title><c:tx><c:rich><a:bodyPr wrap="square" /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="en-US" /><a:t>${escapeXml(title)}</a:t></a:r></a:p></c:rich></c:tx><c:layout /><c:overlay val="0" /><c:spPr><a:noFill /><a:ln><a:noFill /></a:ln></c:spPr></c:title>` : "";

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0" /><c:chart>${titleXml}<c:autoTitleDeleted val="0" /><c:plotArea><c:layout />${chartTypeXml}<c:catAx><c:axId val="${axId1}" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" /><c:numFmt formatCode="General" sourceLinked="1" /><c:majorTickMark val="out" /><c:minorTickMark val="none" /><c:tickLblPos val="nextTo" /><c:crossAx val="${axId2}" /><c:crosses val="autoZero" /><c:lblOffset val="100" /><c:noMultiLvlLbl val="0" /></c:catAx><c:valAx><c:axId val="${axId2}" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" /><c:majorGridlines><c:spPr><a:ln /></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1" /><c:majorTickMark val="out" /><c:minorTickMark val="none" /><c:tickLblPos val="nextTo" /><c:crossAx val="${axId1}" /><c:crosses val="autoZero" /><c:crossBetween val="between" /></c:valAx><c:spPr><a:solidFill><a:srgbClr val="C0C0C0" /></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="808080" /></a:solidFill></a:ln></c:spPr></c:plotArea><c:legend><c:legendPos val="${legendPos}" /><c:layout /><c:overlay val="0" /></c:legend><c:plotVisOnly val="1" /><c:dispBlanksAs val="gap" /><c:showDLblsOverMax val="0" /></c:chart></c:chartSpace>`;
  }

  private generateChartTitle(title: string): string {
    return `<c:title><c:overlay val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr><c:txPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0" i="0" u="none" strike="noStrike" kern="1200" spc="0" baseline="0"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr></c:title>`;
  }

  private generateChartLegendFull(position: string): string {
    const posMap: { [key: string]: string } = { left: "l", right: "r", top: "t", bottom: "b" };
    const pos = posMap[position] || "b";
    return `<c:legend><c:legendPos val="${pos}"/><c:overlay val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr><c:txPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr></c:legend>`;
  }

  private generateChartLegend(position: string): string {
    const posMap: { [key: string]: string } = { left: "l", right: "r", top: "t", bottom: "b" };
    const pos = posMap[position] || "r";
    return `<c:legend><c:legendPos val="${pos}"/></c:legend>`;
  }

  private generateChartSeries(chart: ChartInfo, chartType?: string): string {
    if (!chart.series || chart.series.length === 0) return "";
    
    let xml = "";
    const sheetName = this._sheets.worksheets[chart.sheetIndex]?.name || "Sheet1";
    const accentColors = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
    
    const catValues = chart.series[0]?.categories || [];
    const catCount = catValues.length;
    let catCacheXml = "";
    if (catCount > 0) {
      catCacheXml = catValues.map((c, j) => `<c:pt idx="${j}"><c:v>${escapeXml(String(c))}</c:v></c:pt>`).join("");
    }
    
    const colLetter = (idx: number) => String.fromCharCode(65 + idx);
    
    for (let i = 0; i < chart.series.length; i++) {
      const ser = chart.series[i];
      const idx = i;
      const order = i;
      const color = accentColors[i % accentColors.length];
      
      let catXml = "";
      if (catCount > 0) {
        catXml = `<c:cat><c:strRef><c:f>${sheetName}!$A$2:$A${1 + catCount}</c:f><c:strCache><c:ptCount val="${catCount}"/>${catCacheXml}</c:strCache></c:strRef></c:cat>`;
      }
      
      let valXml = "";
      if (ser.values && ser.values.length > 0) {
        const valValues = ser.values;
        const valCol = colLetter(1 + i);
        const valCacheXml = valValues.map((v, j) => `<c:pt idx="${j}"><c:v>${v}</c:v></c:pt>`).join("");
        valXml = `<c:val><c:numRef><c:f>${sheetName}!$${valCol}$2:$${valCol}${1 + valValues.length}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${valValues.length}"/>${valCacheXml}</c:numCache></c:numRef></c:val>`;
      }
      
      const extLstXml = `<c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{${String(i).padStart(8, '0')}-384C-4FB8-8CC3-16A0F5B4F08F}"/></c:ext></c:extLst>`;
      xml += `<c:ser><c:idx val="${idx}"/><c:order val="${order}"/><c:spPr><a:solidFill><a:schemeClr val="${color}"/></a:solidFill><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr><c:invertIfNegative val="0"/>${catXml}${valXml}${extLstXml}</c:ser>`;
    }
    
    return xml;
  }

  private getCellRef(row: number, col: number): string {
    let colStr = "";
    let c = col;
    while (c >= 0) {
      colStr = String.fromCharCode((c % 26) + 65) + colStr;
      c = Math.floor(c / 26) - 1;
    }
    return `${colStr}${row + 1}`;
  }

  private generateCategoryAxisFull(axId: string, crossAxId: string, position: string): string {
    const posMap: { [key: string]: string } = { left: "l", right: "r", top: "t", bottom: "b" };
    const pos = posMap[position] || "b";
    return `<c:catAx><c:axId val="${axId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:noFill/><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr><c:txPr><a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr><c:crossAx val="${crossAxId}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  }

  private generateValueAxisFull(axId: string, crossAxId: string, position: string): string {
    const posMap: { [key: string]: string } = { left: "l", right: "r", top: "t", bottom: "b" };
    const pos = posMap[position] || "l";
    return `<c:valAx><c:axId val="${axId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/><c:majorGridlines><c:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr></c:majorGridlines><c:numFmt formatCode="0%" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr><c:txPr><a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr><c:crossAx val="${crossAxId}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  }

  private generateCategoryAxis(axId: string, crossAxId: string, position: string): string {
    const posMap: { [key: string]: string } = { left: "l", right: "r", top: "t", bottom: "b" };
    const pos = posMap[position] || "b";
    return `<c:catAx><c:axId val="${axId}"/><c:scaling/><c:delete val="0"/><c:axPos val="${pos}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${crossAxId}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  }

  private generateValueAxis(axId: string, crossAxId: string, position: string): string {
    const posMap: { [key: string]: string } = { left: "l", right: "r", top: "t", bottom: "b" };
    const pos = posMap[position] || "l";
    return `<c:valAx><c:axId val="${axId}"/><c:scaling/><c:delete val="0"/><c:axPos val="${pos}"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${crossAxId}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  }

  private generateDrawingRelsForSheet(sheet: Worksheet, globalImageMap: Map<string, number>): string {
    const sortedShapes = [...sheet.shapes].sort((a, b) => {
      const aZ = (a as any).zIndex || 0;
      const bZ = (b as any).zIndex || 0;
      if (aZ !== bZ) return aZ - bZ;
      const aCol = (a as any).fromCol ?? 0;
      const bCol = (b as any).fromCol ?? 0;
      return aCol - bCol;
    });

    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;

    let rid = 1;
    let chartIndex = 1;
    const processedImages = new Set<string>();
    
    for (const shape of sortedShapes) {
      if ((shape as any).type === "chart" || (shape as any).chartType) {
        xml += `<Relationship Id="rId${rid++}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartIndex++}.xml" />`;
      } else if ((shape as any).type === "picture") {
        const imgName = (shape as any).imageName || (shape as any).imageEmbedId;
        if (imgName && !processedImages.has(imgName)) {
          const img = sheet.images.find(i => i.name === imgName);
          if (img) {
            processedImages.add(imgName);
            const globalIdx = globalImageMap.get(img.name);
            if (globalIdx !== undefined) {
              const ext = img.ext || "png";
              xml += `<Relationship Id="rId${rid++}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${globalIdx + 1}.${ext}" />`;
            }
          }
        }
      }
    }

    xml += "</Relationships>";
    return xml;
  }

  private calculateShapePositions(worksheet: Worksheet): void {
    const pxToEmu = 9525;
    const defaultColWidthEMU = 914400;
    const defaultRowHeightEMU = 161925; // 12.75 pt * 12700 EMU/pt
    const ptToEmu = 12700;

    const colWidths: number[] = [];
    for (let c = 0; c < 20; c++) {
      const w = worksheet.getColumnWidth(c);
      colWidths.push(w ? w / 7 * 914400 / 8 : defaultColWidthEMU);
    }

    const rowHeights: number[] = [];
    for (let r = 0; r < 100; r++) {
      const h = worksheet.getRowHeight(r);
      rowHeights.push(h ? h * ptToEmu : defaultRowHeightEMU);
    }

    for (const shape of worksheet.shapes) {
      const shapeType = (shape as any).type;
      
      // Only recalculate positions for picture shapes
      if (shapeType === "picture") {
        const fromCol = shape.fromCol;
        const fromColOff = shape.fromColOff ?? 0;
        const fromRow = shape.fromRow;
        const fromRowOff = shape.fromRowOff ?? 0;

        const img = (worksheet as any)._images?.find((i: any) => i.name === (shape as any).imageName);
        const imgWidthPx = img?.width ?? 100;
        const imgHeightPx = img?.height ?? 50;
        const imgWidth = imgWidthPx * pxToEmu;
        const imgHeight = imgHeightPx * pxToEmu;

        let x: number;
        let y: number;

        x = 0;
        for (let c = 0; c < fromCol; c++) {
          x += colWidths[c] || defaultColWidthEMU;
        }
        x += fromColOff;

        y = 0;
        for (let r = 0; r < fromRow; r++) {
          y += rowHeights[r] || defaultRowHeightEMU;
        }
        y += fromRowOff;

        const xEnd = x + imgWidth;
        const yEnd = y + imgHeight;

        let toCol = 0;
        let cumX = 0;
        let toColOff = 0;
        for (let c = 0; c < colWidths.length; c++) {
          const colW = colWidths[c] || defaultColWidthEMU;
          const colRight = cumX + colW;
          if (colRight >= xEnd) {
            if (xEnd === colRight) {
              toCol = c + 1;
              toColOff = 0;
            } else {
              toCol = c;
              toColOff = xEnd - cumX;
            }
            break;
          }
          cumX += colW;
        }

        let toRow = 0;
        let cumY = 0;
        let toRowOff = 0;
        for (let r = 0; r < rowHeights.length; r++) {
          const rowH = rowHeights[r] || defaultRowHeightEMU;
          const rowBottom = cumY + rowH;
          if (rowBottom >= yEnd) {
            if (yEnd === rowBottom) {
              toRow = r + 1;
              toRowOff = 0;
            } else {
              toRow = r;
              toRowOff = yEnd - cumY;
            }
            break;
          }
          cumY += rowH;
        }

        shape.x = x;
        shape.y = y;
        shape.width = imgWidth;
        shape.height = imgHeight;

        (shape as any).toCol = toCol;
        (shape as any).toColOff = toColOff;
        (shape as any).toRow = toRow;
        (shape as any).toRowOff = toRowOff;
      } else {
        // For non-picture shapes, preserve original x/y if available
        const xfrmX = (shape as any).xfrmX;
        const xfrmY = (shape as any).xfrmY;
        
        if (xfrmX !== undefined && xfrmY !== undefined) {
          shape.x = xfrmX;
          shape.y = xfrmY;
        }
      }
    }
  }

  private getColumnPositionEMU(worksheet: Worksheet, col: number): number {
    if (col <= 0) return 0;
    const defaultColWidth = 914400;
    let pos = 0;
    for (let c = 0; c < col; c++) {
      const width = worksheet.getColumnWidth(c);
      pos += width ? width / 7 * 914400 / 8 : defaultColWidth;
    }
    return Math.round(pos);
  }

  private getColumnPositionEMUWithDefault(worksheet: Worksheet, col: number): number {
    return col * 914400;
  }

  private getRowPositionEMU(worksheet: Worksheet, row: number): number {
    if (row <= 0) return 0;
    const defaultRowHeight = 161925; // 12.75pt * 12700
    let pos = 0;
    for (let r = 0; r < row; r++) {
      const height = worksheet.getRowHeight(r);
      pos += height ? height * 12700 : defaultRowHeight;
    }
    return Math.round(pos);
  }

  private getColumnWidthEMU(worksheet: Worksheet, col: number): number {
    const defaultColWidth = 914400;
    const width = worksheet.getColumnWidth(col);
    return width ? width / 7 * 914400 / 8 : defaultColWidth;
  }

  private getRowHeightEMU(worksheet: Worksheet, row: number): number {
    const defaultRowHeight = 161925; // 12.75pt * 12700
    const height = worksheet.getRowHeight(row);
    return height ? height * 12700 : defaultRowHeight;
  }

  private getAnchorForPositionEMU(worksheet: Worksheet, xEmu: number, yEmu: number): { col: number; colOff: number; row: number; rowOff: number } {
    let col = 0;
    let colPos = 0;
    while (colPos + this.getColumnWidthEMU(worksheet, col) <= xEmu) {
      colPos += this.getColumnWidthEMU(worksheet, col);
      col++;
    }
    const colOff = xEmu - colPos;

    let row = 0;
    let rowPos = 0;
    while (rowPos + this.getRowHeightEMU(worksheet, row) <= yEmu) {
      rowPos += this.getRowHeightEMU(worksheet, row);
      row++;
    }
    const rowOff = yEmu - rowPos;

    return { col, colOff: Math.round(colOff), row, rowOff: Math.round(rowOff) };
  }

  private getAnchorForPositionEMUWithDefaults(xEmu: number, yEmu: number, defaultColWidthEMU: number, defaultRowHeightEMU: number): { col: number; colOff: number; row: number; rowOff: number } {
    const col = Math.floor(xEmu / defaultColWidthEMU);
    const colOff = xEmu - col * defaultColWidthEMU;

    const row = Math.floor(yEmu / defaultRowHeightEMU);
    const rowOff = yEmu - row * defaultRowHeightEMU;

    return { col, colOff: Math.round(colOff), row, rowOff: Math.round(rowOff) };
  }

  private generateDrawingsRels(): string {
    return DataRelationship.drawingsRels(
      this._sheets.worksheets.map((w) => w.shapes),
    );
  }

  private generateDrawingXml(worksheet: Worksheet): string {
    return ExpDrawing.generateDrawingXml(worksheet.shapes);
  }

  private generateRels(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml" /><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml" /><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" /></Relationships>`;
  }

  private generateCore(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title /><dc:subject /><dc:creator /><cp:keywords /><dc:description /><cp:lastModifiedBy /><cp:category /></cp:coreProperties>`;
  }
  private generateApp(): string {
    const sheetCount = this._sheets.worksheets.length;
    let sheetNames = "";
    for (const sheet of this._sheets.worksheets) {
      sheetNames += `<vt:lpstr>${escapeXml(sheet.name)}</vt:lpstr>`;
    }
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Template/><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${sheetCount}</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="${sheetCount}" baseType="lpstr">${sheetNames}</vt:vector></TitlesOfParts><Manager/><Company/><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>`;
  }

  private generateTheme(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000" /></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF" /></a:lt1><a:dk2><a:srgbClr val="44546A" /></a:dk2><a:lt2><a:srgbClr val="E7E6E6" /></a:lt2><a:accent1><a:srgbClr val="4472C4" /></a:accent1><a:accent2><a:srgbClr val="ED7D31" /></a:accent2><a:accent3><a:srgbClr val="A5A5A5" /></a:accent3><a:accent4><a:srgbClr val="FFC000" /></a:accent4><a:accent5><a:srgbClr val="5B9BD5" /></a:accent5><a:accent6><a:srgbClr val="70AD47" /></a:accent6><a:hlink><a:srgbClr val="0563C1" /></a:hlink><a:folHlink><a:srgbClr val="954F72" /></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204" /><a:ea typeface="" /><a:cs typeface="" /><a:font script="Jpan" typeface="游ゴシック Light" /><a:font script="Hang" typeface="맑은 고딕" /><a:font script="Hans" typeface="等线 Light" /><a:font script="Hant" typeface="新細明體" /><a:font script="Arab" typeface="Times New Roman" /><a:font script="Hebr" typeface="Times New Roman" /><a:font script="Thai" typeface="Tahoma" /><a:font script="Ethi" typeface="Nyala" /><a:font script="Beng" typeface="Vrinda" /><a:font script="Gujr" typeface="Shruti" /><a:font script="Khmr" typeface="MoolBoran" /><a:font script="Knda" typeface="Tunga" /><a:font script="Guru" typeface="Raavi" /><a:font script="Cans" typeface="Euphemia" /><a:font script="Cher" typeface="Plantagenet Cherokee" /><a:font script="Yiii" typeface="Microsoft Yi Baiti" /><a:font script="Tibt" typeface="Microsoft Himalaya" /><a:font script="Thaa" typeface="MV Boli" /><a:font script="Deva" typeface="Mangal" /><a:font script="Telu" typeface="Gautami" /><a:font script="Taml" typeface="Latha" /><a:font script="Syrc" typeface="Estrangelo Edessa" /><a:font script="Orya" typeface="Kalinga" /><a:font script="Mlym" typeface="Kartika" /><a:font script="Laoo" typeface="DokChampa" /><a:font script="Sinh" typeface="Iskoola Pota" /><a:font script="Mong" typeface="Mongolian Baiti" /><a:font script="Viet" typeface="Times New Roman" /><a:font script="Uigh" typeface="Microsoft Uighur" /><a:font script="Geor" typeface="Sylfaen" /><a:font script="Armn" typeface="Arial" /><a:font script="Bugi" typeface="Leelawadee UI" /><a:font script="Bopo" typeface="Microsoft JhengHei" /><a:font script="Java" typeface="Javanese Text" /><a:font script="Lisu" typeface="Segoe UI" /><a:font script="Mymr" typeface="Myanmar Text" /><a:font script="Nkoo" typeface="Ebrima" /><a:font script="Olck" typeface="Nirmala UI" /><a:font script="Osma" typeface="Ebrima" /><a:font script="Phag" typeface="Phagspa" /><a:font script="Syrn" typeface="Estrangelo Edessa" /><a:font script="Syrj" typeface="Estrangelo Edessa" /><a:font script="Syre" typeface="Estrangelo Edessa" /><a:font script="Sora" typeface="Nirmala UI" /><a:font script="Tale" typeface="Microsoft Tai Le" /><a:font script="Talu" typeface="Microsoft New Tai Lue" /><a:font script="Tfng" typeface="Ebrima" /></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204" /><a:ea typeface="" /><a:cs typeface="" /><a:font script="Jpan" typeface="游ゴシック" /><a:font script="Hang" typeface="맑은 고딕" /><a:font script="Hans" typeface="等线" /><a:font script="Hant" typeface="新細明體" /><a:font script="Arab" typeface="Arial" /><a:font script="Hebr" typeface="Arial" /><a:font script="Thai" typeface="Tahoma" /><a:font script="Ethi" typeface="Nyala" /><a:font script="Beng" typeface="Vrinda" /><a:font script="Gujr" typeface="Shruti" /><a:font script="Khmr" typeface="DaunPenh" /><a:font script="Knda" typeface="Tunga" /><a:font script="Guru" typeface="Raavi" /><a:font script="Cans" typeface="Euphemia" /><a:font script="Cher" typeface="Plantagenet Cherokee" /><a:font script="Yiii" typeface="Microsoft Yi Baiti" /><a:font script="Tibt" typeface="Microsoft Himalaya" /><a:font script="Thaa" typeface="MV Boli" /><a:font script="Deva" typeface="Mangal" /><a:font script="Telu" typeface="Gautami" /><a:font script="Taml" typeface="Latha" /><a:font script="Syrc" typeface="Estrangelo Edessa" /><a:font script="Orya" typeface="Kalinga" /><a:font script="Mlym" typeface="Kartika" /><a:font script="Laoo" typeface="DokChampa" /><a:font script="Sinh" typeface="Iskoola Pota" /><a:font script="Mong" typeface="Mongolian Baiti" /><a:font script="Viet" typeface="Arial" /><a:font script="Uigh" typeface="Microsoft Uighur" /><a:font script="Geor" typeface="Sylfaen" /><a:font script="Armn" typeface="Arial" /><a:font script="Bugi" typeface="Leelawadee UI" /><a:font script="Bopo" typeface="Microsoft JhengHei" /><a:font script="Java" typeface="Javanese Text" /><a:font script="Lisu" typeface="Segoe UI" /><a:font script="Mymr" typeface="Myanmar Text" /><a:font script="Nkoo" typeface="Ebrima" /><a:font script="Olck" typeface="Nirmala UI" /><a:font script="Osma" typeface="Ebrima" /><a:font script="Phag" typeface="Phagspa" /><a:font script="Syrn" typeface="Estrangelo Edessa" /><a:font script="Syrj" typeface="Estrangelo Edessa" /><a:font script="Syre" typeface="Estrangelo Edessa" /><a:font script="Sora" typeface="Nirmala UI" /><a:font script="Tale" typeface="Microsoft Tai Le" /><a:font script="Talu" typeface="Microsoft New Tai Lue" /><a:font script="Tfng" typeface="Ebrima" /></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr" /></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000" /><a:satMod val="105000" /><a:tint val="67000" /></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000" /><a:satMod val="103000" /><a:tint val="73000" /></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000" /><a:satMod val="109000" /><a:tint val="81000" /></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0" /></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000" /><a:lumMod val="102000" /><a:tint val="94000" /></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000" /><a:lumMod val="100000" /><a:shade val="100000" /></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000" /><a:satMod val="120000" /><a:shade val="78000" /></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0" /></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr" /></a:solidFill><a:prstDash val="solid" /><a:miter lim="800000" /></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr" /></a:solidFill><a:prstDash val="solid" /><a:miter lim="800000" /></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr" /></a:solidFill><a:prstDash val="solid" /><a:miter lim="800000" /></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst /></a:effectStyle><a:effectStyle><a:effectLst /></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000" /></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr" /></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000" /><a:satMod val="170000" /></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000" /><a:satMod val="150000" /><a:shade val="98000" /><a:lumMod val="102000" /></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000" /><a:satMod val="130000" /><a:shade val="90000" /><a:lumMod val="103000" /></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000" /><a:satMod val="120000" /></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0" /></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults /><a:extraClrSchemeLst /><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" /></a:ext></a:extLst></a:theme>`;
  }

  private generateContentTypes(): string {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" /><Default Extension="xml" ContentType="application/xml" /><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" /><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml" /><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" /><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml" /><Default Extension="png" ContentType="image/png" /><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />`;

    let chartCount = 0;
    let drawingCount = 0;
    for (let i = 0; i < this._sheets.worksheets.length; i++) {
      const sheet = this._sheets.worksheets[i];
      
      for (const shape of sheet.shapes) {
        if ((shape as any).type === "chart" || (shape as any).chartType) {
          chartCount++;
        }
      }
      
      if (sheet.shapes.length > 0) {
        drawingCount++;
        xml += `<Override PartName="/xl/drawings/drawing${drawingCount}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml" />`;
      }
      xml += `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />`;
    }

    for (let i = 1; i <= chartCount; i++) {
      xml += `<Override PartName="/xl/charts/chart${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml" />`;
    }

    // Add shared strings part if needed
    if (this._sharedStrings.length > 0) {
      xml += `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" />`;
    }

    xml += "</Types>";
    return xml;
  }

  private generateWorkbook(): string {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="9302" /><workbookPr /><bookViews><workbookView xWindow="240" yWindow="120" windowWidth="14940" windowHeight="9225" activeTab="0" /></bookViews><sheets>`;

    if (this._sheets.worksheets.length > 0) {
      xml += `<sheet name="${escapeXml(this._sheets.worksheets[0].name)}" sheetId="1" r:id="rId3" />`;
    }
    if (this._sheets.worksheets.length > 1) {
      xml += `<sheet name="${escapeXml(this._sheets.worksheets[1].name)}" sheetId="2" r:id="rId4" />`;
    }

    xml += `</sheets><definedNames /><calcPr fullCalcOnLoad="1" /></workbook>`;
    return xml;
  }

  private generateWorkbookRels(): string {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;

    if (this._sharedStrings.length > 0) {
      xml += `<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml" />`;
    }

    if (this._sheets.worksheets.length > 1) {
      xml += `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml" />`;
    }

    xml += `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />`;

    if (this._sheets.worksheets.length > 0) {
      xml += `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" />`;
    }

    xml += `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml" /></Relationships>`;
    return xml;
  }

  private collectCellStyles() {
    const styleMap = new Map<string, number>();
    let nextStyleIndex = 1;

    for (const sheet of this._sheets.worksheets) {
      for (const cell of sheet.cells) {
        const cellStyle = cell.style;
        if (cellStyle) {
          const styleKey = JSON.stringify(cellStyle);
          let styleIndex = styleMap.get(styleKey);
          if (styleIndex === undefined) {
            styleIndex = nextStyleIndex++;
            styleMap.set(styleKey, styleIndex);
            this._sheets.setStyle(styleIndex, cellStyle);
          }
          cell.setStyleIndex(styleIndex);
        }
      }
    }
  }

  private normalizeColor(color: string | undefined): string | undefined {
    if (!color) return undefined;
    // Convert CSS #RRGGBB to ARGB FFRRGGBB
    if (color.startsWith("#")) {
      return "FF" + color.substring(1).toUpperCase();
    }
    // Already in ARGB format (8 hex digits)
    if (/^[0-9A-Fa-f]{8}$/.test(color)) return color.toUpperCase();
    return color;
  }

  private generateStyles(): string {
    const styles = this._sheets.styles;

    // Collect unique fonts (index 0 = default - Excel default is Calibri 11)
    const fontEntries: any[] = [
      { name: "Calibri", size: 11, bold: false, italic: false },
    ];
    const fontKeyMap = new Map<string, number>();

    // Collect unique fills (index 0 = none, index 1 = gray125)
    const fillEntries: any[] = [
      { patternType: "none" },
      { patternType: "gray125" },
    ];
    const fillKeyMap = new Map<string, number>();

    // Collect unique borders (index 0 = empty)
    const borderEntries: any[] = [null];
    const borderKeyMap = new Map<string, number>();

    // Collect unique number formats
    const numFmtEntries: { id: number; code: string }[] = [];
    const numFmtKeyMap = new Map<string, number>();
    let nextNumFmtId = 164;
    const builtinFmts = [
      "General",
      "0",
      "0.00",
      "#,##0",
      "#,##0.00",
      "0%",
      "0.00%",
      "0.00E+00",
    ];

    // Build cellXfEntries indexed by Map key so that cell s="N" matches
    // position N in the cellXfs array. Fill gaps with default entries.
    type XfEntry = {
      fontId: number;
      fillId: number;
      borderId: number;
      numFmtId: number;
      style: Style;
    };

    const maxStyleIdx = styles.size > 0 ? Math.max(...styles.keys()) : 0;
    const cellXfEntries: XfEntry[] = [];

    for (let idx = 0; idx <= maxStyleIdx; idx++) {
      const style = styles.get(idx);
      if (!style) {
        // Default "Normal" entry for gaps / index 0
        cellXfEntries.push({
          fontId: 0,
          fillId: 0,
          borderId: 0,
          numFmtId: 0,
          style: {} as Style,
        });
        continue;
      }

      let fontId = 0;
      if (style.font) {
        const fontKey = JSON.stringify({
          name: style.font.name,
          size: style.font.size,
          color: this.normalizeColor(style.font.color),
          bold: !!style.font.bold,
          italic: !!style.font.italic,
        });
        if (fontKeyMap.has(fontKey)) {
          fontId = fontKeyMap.get(fontKey)!;
        } else {
          fontId = fontEntries.length;
          fontKeyMap.set(fontKey, fontId);
          fontEntries.push({
            ...style.font,
            color: this.normalizeColor(style.font.color),
          });
        }
      }

      let fillId = 0;
      if (style.fill?.patternType && style.fill.patternType !== "none") {
        const fillKey = JSON.stringify({
          patternType: style.fill.patternType,
          fgColor: this.normalizeColor(style.fill.fgColor),
          bgColor: this.normalizeColor(style.fill.bgColor),
        });
        if (fillKeyMap.has(fillKey)) {
          fillId = fillKeyMap.get(fillKey)!;
        } else {
          fillId = fillEntries.length;
          fillKeyMap.set(fillKey, fillId);
          fillEntries.push({
            ...style.fill,
            fgColor: this.normalizeColor(style.fill.fgColor),
            bgColor: this.normalizeColor(style.fill.bgColor),
          });
        }
      }

      let borderId = 0;
      if (style.border) {
        const borderKey = JSON.stringify(style.border);
        if (borderKeyMap.has(borderKey)) {
          borderId = borderKeyMap.get(borderKey)!;
        } else {
          borderId = borderEntries.length;
          borderKeyMap.set(borderKey, borderId);
          borderEntries.push(style.border);
        }
      }

      let numFmtId = 0;
      if (style.numberFormat && !builtinFmts.includes(style.numberFormat)) {
        if (numFmtKeyMap.has(style.numberFormat)) {
          numFmtId = numFmtKeyMap.get(style.numberFormat)!;
        } else {
          numFmtId = nextNumFmtId++;
          numFmtKeyMap.set(style.numberFormat, numFmtId);
          numFmtEntries.push({ id: numFmtId, code: style.numberFormat });
        }
      }

      cellXfEntries.push({ fontId, fillId, borderId, numFmtId, style });
    }

    // --- Generate fonts XML ---
    let fontsXml = `<fonts count="${fontEntries.length}" x14ac:knownFonts="1">`;
    for (const font of fontEntries) {
      const size = font.size || 10;
      const name = font.name || "Arial";
      const bold = font.bold ? "<b/>" : "";
      const italic = font.italic ? "<i/>" : "";
      let colorAttr: string;
      if (font.color) {
        colorAttr = ` rgb="${font.color}"`;
      } else {
        colorAttr = ' theme="1"';
      }
      fontsXml += `<font>${bold}${italic}<sz val="${size}"/><color${colorAttr}/><name val="${name}"/><family val="2"/></font>`;
    }
    fontsXml += "</fonts>";

    // --- Generate fills XML ---
    let fillsXml = `<fills count="${fillEntries.length}">`;
    for (const fill of fillEntries) {
      if (fill.patternType === "none") {
        fillsXml += `<fill><patternFill patternType="none"/></fill>`;
      } else if (fill.patternType === "gray125") {
        fillsXml += `<fill><patternFill patternType="gray125"/></fill>`;
      } else {
        const fgColor = fill.fgColor ? `<fgColor rgb="${fill.fgColor}"/>` : "";
        const bgColor = fill.bgColor
          ? `<bgColor rgb="${fill.bgColor}"/>`
          : '<bgColor indexed="64"/>';
        fillsXml += `<fill><patternFill patternType="${fill.patternType || "solid"}">${fgColor}${bgColor}</patternFill></fill>`;
      }
    }
    fillsXml += "</fills>";

    // --- Generate borders XML (fix: use correct tag name per side) ---
    let bordersXml = `<borders count="${borderEntries.length}">`;
    bordersXml += `<border><left/><right/><top/><bottom/><diagonal/></border>`;
    for (let i = 1; i < borderEntries.length; i++) {
      const border = borderEntries[i];
      const renderSide = (tagName: string, side: any): string => {
        if (!side?.style) return `<${tagName}/>`;
        // Map CSS border styles to valid OOXML border styles
        const xlStyle =
          side.style === "solid" || side.style === "thin"
            ? "thin"
            : side.style === "medium"
              ? "medium"
              : side.style === "thick"
                ? "thick"
                : side.style === "dashed"
                  ? "dashed"
                  : side.style === "dotted"
                    ? "dotted"
                    : "thin";
        const color = side.color
          ? ` rgb="${this.normalizeColor(side.color)}"`
          : ' auto="1"';
        return `<${tagName} style="${xlStyle}"><color${color}/></${tagName}>`;
      };
      bordersXml += `<border>${renderSide("left", border.left)}${renderSide("right", border.right)}${renderSide("top", border.top)}${renderSide("bottom", border.bottom)}<diagonal/></border>`;
    }
    bordersXml += "</borders>";

    // --- Generate numFmts XML ---
    let numFmtsXml = "";
    if (numFmtEntries.length > 0) {
      numFmtsXml = `<numFmts count="${numFmtEntries.length}">`;
      for (const nf of numFmtEntries) {
        numFmtsXml += `<numFmt numFmtId="${nf.id}" formatCode="${escapeXml(nf.code)}"/>`;
      }
      numFmtsXml += "</numFmts>";
    }

    // --- Generate cellXfs XML ---
    let cellXfsXml = `<cellXfs count="${cellXfEntries.length || 1}">`;
    if (cellXfEntries.length === 0) {
      cellXfsXml += `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`;
    }
    for (const xf of cellXfEntries) {
      const applyFont = xf.fontId !== 0 ? ' applyFont="1"' : "";
      const applyFill = xf.fillId !== 0 ? ' applyFill="1"' : "";
      const applyBorder = xf.borderId !== 0 ? ' applyBorder="1"' : "";
      const applyNumFmt = xf.numFmtId !== 0 ? ' applyNumberFormat="1"' : "";

      let applyAlignment = "";
      let alignmentXml = "";
      if (xf.style.alignment) {
        const h = xf.style.alignment.horizontal
          ? ` horizontal="${xf.style.alignment.horizontal}"`
          : "";
        const v = xf.style.alignment.vertical
          ? ` vertical="${xf.style.alignment.vertical}"`
          : "";
        const wrap = xf.style.alignment.wrapText ? ' wrapText="1"' : "";
        if (h || v || wrap) {
          applyAlignment = ' applyAlignment="1"';
          alignmentXml = `<alignment${h}${v}${wrap}/>`;
        }
      }

      cellXfsXml += `<xf numFmtId="${xf.numFmtId}" fontId="${xf.fontId}" fillId="${xf.fillId}" borderId="${xf.borderId}" xfId="0"${applyNumFmt}${applyFont}${applyFill}${applyBorder}${applyAlignment}>${alignmentXml}</xf>`;
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
