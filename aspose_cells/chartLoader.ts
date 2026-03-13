import { ChartInfo, ChartSeries } from "./types";
import {
  openZip as openZipUtil,
  readZipEntry as readZipEntryUtil,
} from "./util";
import { DOMParser } from "@xmldom/xmldom";

export class ChartLoader {
  private charts: ChartInfo[] = [];

  async loadCharts(filePath: string): Promise<ChartInfo[]> {
    const zip = await openZipUtil(filePath);
    try {
      await this.loadChartsFromZip(zip);
    } finally {
      await zip.close();
    }
    return this.charts;
  }

  private async loadChartsFromZip(zip: any) {
    const workbookContent = await readZipEntryUtil(zip, "xl/workbook.xml");
    if (!workbookContent) return;

    const parser = new DOMParser();
    const workbookDoc = parser.parseFromString(workbookContent, "text/xml");

    const sheets = workbookDoc.getElementsByTagName("sheet");
    for (let sheetIndex = 0; sheetIndex < sheets.length; sheetIndex++) {
      const sheet = sheets[sheetIndex];
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

      const sheetRelsPath =
        "xl/" + target.replace("worksheets/", "worksheets/_rels/") + ".rels";
      const sheetRelsContent = await readZipEntryUtil(zip, sheetRelsPath);
      if (!sheetRelsContent) continue;

      const sheetRelsParser = new DOMParser();
      const sheetRelsDoc = sheetRelsParser.parseFromString(
        sheetRelsContent,
        "text/xml",
      );
      const sheetRels = sheetRelsDoc.getElementsByTagName("Relationship");

      let drawingTarget = "";
      for (let i = 0; i < sheetRels.length; i++) {
        const rel = sheetRels[i];
        const type = rel.getAttribute("Type") || "";
        if (type.includes("drawing")) {
          drawingTarget = rel.getAttribute("Target") || "";
          break;
        }
      }

      if (!drawingTarget) continue;

      let drawingPath: string;
      if (drawingTarget.startsWith("../")) {
        drawingPath = "xl/" + drawingTarget.substring(3);
      } else if (target.includes("/")) {
        drawingPath = "xl/" + target.replace(/[^/]+$/, "") + drawingTarget;
      } else {
        drawingPath = "xl/drawings/" + drawingTarget;
      }

      const drawingContent = await readZipEntryUtil(zip, drawingPath);
      if (!drawingContent) continue;

      const drawingDoc = parser.parseFromString(drawingContent, "text/xml");
      const graphicFrames = drawingDoc.getElementsByTagName("xdr:graphicFrame");

      for (let i = 0; i < graphicFrames.length; i++) {
        const frame = graphicFrames[i];
        const graphicData = frame.getElementsByTagName("a:graphicData")[0];
        if (!graphicData) continue;

        const uri = graphicData.getAttribute("uri");
        if (!uri || !uri.includes("chart")) continue;

        const chartRel = graphicData.getElementsByTagName("c:chart")[0];
        if (!chartRel) continue;

        const rId = chartRel.getAttributeNS(
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "id",
        );
        if (!rId) continue;

        const drawingRelsPath =
          drawingPath.replace("xl/drawings/", "xl/drawings/_rels/") + ".rels";
        const drawingRelsContent = await readZipEntryUtil(zip, drawingRelsPath);
        if (!drawingRelsContent) continue;

        const drawingRelsParser = new DOMParser();
        const drawingRelsDoc = drawingRelsParser.parseFromString(
          drawingRelsContent,
          "text/xml",
        );
        const drawingRels = drawingRelsDoc.getElementsByTagName("Relationship");

        let chartTarget = "";
        for (let j = 0; j < drawingRels.length; j++) {
          const rel = drawingRels[j];
          if (rel.getAttribute("Id") === rId) {
            chartTarget = rel.getAttribute("Target") || "";
            break;
          }
        }

        if (!chartTarget) continue;

        let chartPath: string;
        if (chartTarget.startsWith("../")) {
          chartPath = "xl/" + chartTarget.substring(3);
        } else {
          chartPath = drawingPath.replace(/[^/]+$/, "") + chartTarget;
        }

        const chartContent = await readZipEntryUtil(zip, chartPath);
        if (!chartContent) continue;

        const chartDoc = parser.parseFromString(chartContent, "text/xml");

        const twoCellAnchor = frame.parentNode as Element;
        let fromCol = 0,
          fromRow = 0,
          toCol = 5,
          toRow = 15;
        let fromColOff = 0,
          fromRowOff = 0,
          toColOff = 0,
          toRowOff = 0;

        if (twoCellAnchor && twoCellAnchor.nodeName === "xdr:twoCellAnchor") {
          const from = twoCellAnchor.getElementsByTagName("xdr:from")[0];
          const to = twoCellAnchor.getElementsByTagName("xdr:to")[0];
          if (from) {
            fromCol = parseInt(this.getXmlValue(from, "xdr:col") || "0", 10);
            fromColOff = parseInt(
              this.getXmlValue(from, "xdr:colOff") || "0",
              10,
            );
            fromRow = parseInt(this.getXmlValue(from, "xdr:row") || "0", 10);
            fromRowOff = parseInt(
              this.getXmlValue(from, "xdr:rowOff") || "0",
              10,
            );
          }
          if (to) {
            toCol = parseInt(this.getXmlValue(to, "xdr:col") || "5", 10);
            toColOff = parseInt(this.getXmlValue(to, "xdr:colOff") || "0", 10);
            toRow = parseInt(this.getXmlValue(to, "xdr:row") || "15", 10);
            toRowOff = parseInt(this.getXmlValue(to, "xdr:rowOff") || "0", 10);
          }
        }

        this.parseChart(
          chartDoc,
          sheetIndex,
          fromCol,
          fromRow,
          toCol,
          toRow,
          fromColOff,
          fromRowOff,
          toColOff,
          toRowOff,
        );
      }
    }
  }

  private getXmlValue(parent: Element, tagName: string): string | null {
    const el = parent.getElementsByTagName(tagName)[0];
    return el?.textContent || null;
  }

  private parseChart(
    chartDoc: Document,
    sheetIndex: number,
    fromCol: number,
    fromRow: number,
    toCol: number,
    toRow: number,
    fromColOff: number,
    fromRowOff: number,
    toColOff: number,
    toRowOff: number,
  ) {
    const chartSpace = chartDoc.getElementsByTagName("c:chartSpace")[0];
    if (!chartSpace) return;

    const chart = chartSpace.getElementsByTagName("c:chart")[0];
    if (!chart) return;

    const titleEl = chart.getElementsByTagName("c:title")[0];
    const title = titleEl ? titleEl.getElementsByTagName("c:tx")[0] : null;
    let chartTitle = "";
    if (title) {
      const richText = title.getElementsByTagName("a:r")[0];
      if (richText) {
        const t = richText.getElementsByTagName("a:t")[0];
        chartTitle = t?.textContent || "";
      }
    }

    const plotArea = chart.getElementsByTagName("c:plotArea")[0];
    if (!plotArea) return;

    let plotAreaLayout:
      | { x: number; y: number; width: number; height: number }
      | undefined;
    const layoutEl = plotArea.getElementsByTagName("c:layout")[0];
    if (layoutEl) {
      const manualLayout = layoutEl.getElementsByTagName("c:manualLayout")[0];
      if (manualLayout) {
        const xEl = manualLayout.getElementsByTagName("c:x")[0];
        const yEl = manualLayout.getElementsByTagName("c:y")[0];
        const wEl = manualLayout.getElementsByTagName("c:w")[0];
        const hEl = manualLayout.getElementsByTagName("c:h")[0];
        if (xEl && yEl && wEl && hEl) {
          plotAreaLayout = {
            x: parseFloat(xEl.getAttribute("val") || "0"),
            y: parseFloat(yEl.getAttribute("val") || "0"),
            width: parseFloat(wEl.getAttribute("val") || "0"),
            height: parseFloat(hEl.getAttribute("val") || "0"),
          };
        }
      }
    }

    const chartTypes =
      plotArea.getElementsByTagName("c:barChart")[0] ||
      plotArea.getElementsByTagName("c:bar3DChart")[0] ||
      plotArea.getElementsByTagName("c:lineChart")[0] ||
      plotArea.getElementsByTagName("c:line3DChart")[0] ||
      plotArea.getElementsByTagName("c:pieChart")[0] ||
      plotArea.getElementsByTagName("c:pie3DChart")[0] ||
      plotArea.getElementsByTagName("c:areaChart")[0] ||
      plotArea.getElementsByTagName("c:area3DChart")[0] ||
      plotArea.getElementsByTagName("c:surfaceChart")[0] ||
      plotArea.getElementsByTagName("c:surface3DChart")[0] ||
      plotArea.getElementsByTagName("c:scatterChart")[0] ||
      plotArea.getElementsByTagName("c:doughnutChart")[0] ||
      plotArea.getElementsByTagName("c:doughnut3DChart")[0] ||
      plotArea.getElementsByTagName("c:radarChart")[0] ||
      plotArea.getElementsByTagName("c:radar3DChart")[0];

    let chartType = "column";
    if (plotArea.getElementsByTagName("c:barChart")[0]) {
      const barDir = plotArea
        .getElementsByTagName("c:barChart")[0]
        .getElementsByTagName("c:barDir")[0];
      chartType = barDir?.getAttribute("val") === "bar" ? "bar" : "column";
    } else if (plotArea.getElementsByTagName("c:bar3DChart")[0]) {
      const barDir = plotArea
        .getElementsByTagName("c:bar3DChart")[0]
        .getElementsByTagName("c:barDir")[0];
      chartType = barDir?.getAttribute("val") === "bar" ? "bar" : "column";
    } else if (plotArea.getElementsByTagName("c:lineChart")[0])
      chartType = "line";
    else if (plotArea.getElementsByTagName("c:line3DChart")[0])
      chartType = "line";
    else if (plotArea.getElementsByTagName("c:pieChart")[0]) chartType = "pie";
    else if (plotArea.getElementsByTagName("c:pie3DChart")[0])
      chartType = "pie";
    else if (plotArea.getElementsByTagName("c:areaChart")[0])
      chartType = "area";
    else if (plotArea.getElementsByTagName("c:area3DChart")[0])
      chartType = "area";
    else if (plotArea.getElementsByTagName("c:surfaceChart")[0])
      chartType = "area";
    else if (plotArea.getElementsByTagName("c:surface3DChart")[0])
      chartType = "area";
    else if (plotArea.getElementsByTagName("c:scatterChart")[0])
      chartType = "scatter";
    else if (plotArea.getElementsByTagName("c:doughnutChart")[0])
      chartType = "doughnut";
    else if (plotArea.getElementsByTagName("c:doughnut3DChart")[0])
      chartType = "doughnut";
    else if (plotArea.getElementsByTagName("c:radarChart")[0])
      chartType = "radar";
    else if (plotArea.getElementsByTagName("c:radar3DChart")[0])
      chartType = "radar";
    else if (plotArea.getElementsByTagName("c:radar3DChart")[0])
      chartType = "radar";

    const barDir = chartTypes?.getElementsByTagName("c:barDir")[0];
    if (barDir?.getAttribute("val") === "bar") {
      chartType = "bar";
    }

    const serList = chartTypes ? chartTypes.getElementsByTagName("c:ser") : [];
    const series: Array<{
      name?: string;
      values: number[];
      categories: string[];
    }> = [];

    for (let i = 0; i < serList.length; i++) {
      const ser = serList[i];

      const txEl = ser.getElementsByTagName("c:tx")[0];
      let seriesName = `Series ${i + 1}`;
      if (txEl) {
        const rEl = txEl.getElementsByTagName("a:r")[0];
        if (rEl) {
          const tEl = rEl.getElementsByTagName("a:t")[0];
          seriesName = tEl?.textContent || seriesName;
        }
      }

      const catEl = ser.getElementsByTagName("c:cat")[0];
      const categories: string[] = [];
      if (catEl) {
        const ptList = catEl.getElementsByTagName("c:pt");
        for (let j = 0; j < ptList.length; j++) {
          const pt = ptList[j];
          const vEl = pt.getElementsByTagName("c:v")[0];
          categories.push(vEl?.textContent || "");
        }
        if (categories.length === 0) {
          const strCache = catEl.getElementsByTagName("c:strCache")[0];
          const cachePtList = strCache
            ? strCache.getElementsByTagName("c:pt")
            : [];
          for (let j = 0; j < cachePtList.length; j++) {
            const pt = cachePtList[j];
            const vEl = pt.getElementsByTagName("c:v")[0];
            categories.push(vEl?.textContent || "");
          }
        }
      }

      let valEl = ser.getElementsByTagName("c:val")[0];
      if (!valEl) {
        valEl = ser.getElementsByTagName("c:yVal")[0];
      }
      const values: number[] = [];
      if (valEl) {
        const ptList = valEl.getElementsByTagName("c:pt");
        for (let j = 0; j < ptList.length; j++) {
          const pt = ptList[j];
          const vEl = pt.getElementsByTagName("c:v")[0];
          const val = parseFloat(vEl?.textContent || "0");
          values.push(isNaN(val) ? 0 : val);
        }
        if (values.length === 0) {
          const numCache = valEl.getElementsByTagName("c:numCache")[0];
          const cachePtList = numCache
            ? numCache.getElementsByTagName("c:pt")
            : [];
          for (let j = 0; j < cachePtList.length; j++) {
            const pt = cachePtList[j];
            const vEl = pt.getElementsByTagName("c:v")[0];
            const val = parseFloat(vEl?.textContent || "0");
            values.push(isNaN(val) ? 0 : val);
          }
        }
      }

      series.push({ name: seriesName, values, categories });
    }

    let legendPosition: "left" | "right" | "top" | "bottom" = "right";
    const legendEl = chart.getElementsByTagName("c:legend")[0];
    if (legendEl) {
      const legendPosEl = legendEl.getElementsByTagName("c:legendPos")[0];
      if (legendPosEl) {
        const pos = legendPosEl.getAttribute("val");
        if (pos === "l") legendPosition = "left";
        else if (pos === "r") legendPosition = "right";
        else if (pos === "t") legendPosition = "top";
        else if (pos === "b") legendPosition = "bottom";
      }
    }

    let categoryAxisPosition: "left" | "right" | "top" | "bottom" = "bottom";
    let valueAxisPosition: "left" | "right" | "top" | "bottom" = "left";

    const catAx = plotArea.getElementsByTagName("c:catAx")[0];
    if (catAx) {
      const catAxPos = catAx.getElementsByTagName("c:axPos")[0];
      if (catAxPos) {
        const pos = catAxPos.getAttribute("val");
        if (pos === "l") categoryAxisPosition = "left";
        else if (pos === "r") categoryAxisPosition = "right";
        else if (pos === "t") categoryAxisPosition = "top";
        else if (pos === "b") categoryAxisPosition = "bottom";
      }
    }

    const valAx = plotArea.getElementsByTagName("c:valAx")[0];
    if (valAx) {
      const valAxPos = valAx.getElementsByTagName("c:axPos")[0];
      if (valAxPos) {
        const pos = valAxPos.getAttribute("val");
        if (pos === "l") valueAxisPosition = "left";
        else if (pos === "r") valueAxisPosition = "right";
        else if (pos === "t") valueAxisPosition = "top";
        else if (pos === "b") valueAxisPosition = "bottom";
      }
    }

    this.charts.push({
      name: chartTitle || `Chart ${this.charts.length + 1}`,
      type: chartType,
      chartType: chartType as any,
      title: chartTitle,
      series,
      legendPosition,
      categoryAxisPosition,
      valueAxisPosition,
      plotArea: plotAreaLayout,
      fromRow,
      fromCol,
      toRow,
      toCol,
      fromColOff: fromColOff || 0,
      fromRowOff: fromRowOff || 0,
      toColOff: toColOff || 0,
      toRowOff: toRowOff || 0,
      sheetIndex,
    });
  }
}
