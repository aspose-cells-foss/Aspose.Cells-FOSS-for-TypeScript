import { ChartInfo, ChartSeries } from "./types";
import { escapeXml, cellRef } from "./util";

export class ExpChart {
  static generateChartXml(chart: ChartInfo, sheetName: string): string {
    const chartType = chart.chartType;
    const chartXml = ExpChart.getChartTypeXml(chart, chartType, sheetName);
    
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="0"/><c:lang val="en-GB"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent>${chartXml}</c:chartSpace>`;
  }

  private static getChartTypeXml(chart: ChartInfo, chartType: string, sheetName: string): string {
    const chartXml = `<c:chart><c:autoTitleDeleted val="0"/><c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.14596295425768249"/><c:y val="9.3633301268804384E-2"/><c:w val="0.65217490200241113"/><c:h val="0.7528117422011873"/></c:manualLayout></c:layout>${ExpChart.getChartDataXml(chart, chartType, sheetName)}</c:plotArea>${ExpChart.getLegendXml(chart.legendPosition)}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/></c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="3175"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="solid"/></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800" b="0" i="0" u="none" strike="noStrike" baseline="0"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/></a:defRPr></a:pPr><a:endParaRPr lang="zh-CN"/></a:p></c:txPr><c:printSettings><c:headerFooter alignWithMargins="0"/><c:pageMargins b="1" l="0.75" r="0.75" t="1" header="0.5" footer="0.5"/><c:pageSetup/></c:printSettings>`;
    return chartXml;
  }

  private static getChartDataXml(chart: ChartInfo, chartType: string, sheetName: string): string {
    const seriesXml = chart.series.map((s, idx) => {
      return ExpChart.getSeriesXml(s, idx, chartType, sheetName);
    }).join("");

    switch (chartType) {
      case "bar":
      case "barStacked":
      case "barStacked100":
      case "column":
      case "columnStacked":
      case "columnStacked100":
        return `<c:barChart><c:barDir val="col"/><c:grouping val="${ExpChart.getGrouping(chartType)}"/><c:varyColors val="0"/>${seriesXml}<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls><c:gapWidth val="150"/><c:axId val="107911040"/><c:axId val="107912576"/></c:barChart>${ExpChart.getCatAxXml()}${ExpChart.getValAxXml(chartType)}<c:spPr><a:solidFill><a:srgbClr val="C0C0C0"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="808080"/></a:solidFill><a:prstDash val="solid"/></a:ln></c:spPr>`;
      case "line":
      case "lineStacked":
      case "lineStacked100":
        return `<c:lineChart><c:grouping val="${ExpChart.getGrouping(chartType)}"/>${seriesXml}<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls><c:axId val="107911040"/><c:axId val="107912576"/></c:lineChart>${ExpChart.getCatAxXml()}${ExpChart.getValAxXml(chartType)}<c:spPr><a:solidFill><a:srgbClr val="C0C0C0"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="808080"/></a:solidFill><a:prstDash val="solid"/></a:ln></c:spPr>`;
      case "pie":
      case "doughnut":
        return `<c:pieChart><c:varyColors val="0"/>${seriesXml}<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls></c:pieChart>`;
      case "area":
      case "areaStacked":
      case "areaStacked100":
        return `<c:areaChart><c:grouping val="${ExpChart.getGrouping(chartType)}"/>${seriesXml}<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls><c:axId val="107911040"/><c:axId val="107912576"/></c:areaChart>${ExpChart.getCatAxXml()}${ExpChart.getValAxXml(chartType)}<c:spPr><a:solidFill><a:srgbClr val="C0C0C0"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="808080"/></a:solidFill><a:prstDash val="solid"/></a:ln></c:spPr>`;
      case "scatter":
        return `<c:scatterChart><c:scatterStyle val="lineMarker"/>${seriesXml}</c:scatterChart>`;
      case "radar":
        return `<c:radarChart><c:radarStyle val="marker"/>${seriesXml}</c:radarChart>`;
      default:
        return `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${seriesXml}<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls><c:gapWidth val="150"/><c:axId val="107911040"/><c:axId val="107912576"/></c:barChart>${ExpChart.getCatAxXml()}${ExpChart.getValAxXml(chartType)}<c:spPr><a:solidFill><a:srgbClr val="C0C0C0"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="808080"/></a:solidFill><a:prstDash val="solid"/></a:ln></c:spPr>`;
    }
  }

  private static getGrouping(chartType: string): string {
    switch (chartType) {
      case "barStacked":
      case "columnStacked":
      case "lineStacked":
      case "areaStacked":
        return "stacked";
      case "barStacked100":
      case "columnStacked100":
      case "lineStacked100":
      case "areaStacked100":
        return "percentStacked";
      default:
        return "clustered";
    }
  }

  private static getSeriesXml(series: ChartSeries, index: number, chartType: string, sheetName: string): string {
    const row = 2 + index; // Data starts from row 2
    const catRange = series.categories ? ExpChart.getCategoriesRange(series.categories, sheetName, 2) : "";
    const valRange = ExpChart.getValuesRange(series.values, sheetName, row);
    
    const formatCode = "0%";
    
    const catXml = series.categories ? `<c:cat><c:strRef><c:f>${catRange}</c:f><c:strCache><c:ptCount val="${series.categories.length}"/>${series.categories.map((c, i) => `<c:pt idx="${i}"><c:v>${escapeXml(c)}</c:v></c:pt>`).join("")}</c:strCache></c:strRef></c:cat>` : "";
    
    const valXml = `<c:val><c:numRef><c:f>${valRange}</c:f><c:numCache><c:formatCode>${formatCode}</c:formatCode><c:ptCount val="${series.values.length}"/>${series.values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("")}</c:numCache></c:numRef></c:val>`;

    const colors = ["9999FF", "993366", "C99000", "E95658", "F2B01C", "638ACA", "47C6BF", "DA4F45"];
    const color = colors[index % colors.length];
    const uniqueId = `{0000000${index}-7F15-4AAD-B402-0D917DAD514E}`;

    return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="solid"/></a:ln></c:spPr><c:invertIfNegative val="0"/>${catXml}${valXml}<c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="${uniqueId}"/></c:ext></c:extLst></c:ser>`;
  }

  private static getCategoriesRange(categories: string[], sheetName: string, row: number): string {
    const startCol = 0;
    const endCol = categories.length - 1;
    return `${sheetName}!$${cellRef(1, startCol).replace(/(\d+)/, "$$$1")}:$${cellRef(1, endCol).replace(/(\d+)/, "$$$1")}`;
  }

  private static getValuesRange(values: number[], sheetName: string, row: number): string {
    const startCol = 0;
    const endCol = values.length - 1;
    return `${sheetName}!$${cellRef(row, startCol).replace(/(\d+)/, "$$$1")}:$${cellRef(row, endCol).replace(/(\d+)/, "$$$1")}`;
  }

  private static getLegendXml(position?: string): string {
    if (!position || position === "none") {
      return "<c:legend><c:legendEntry><c:idx val=\"0\"/><c:delete val=\"1\"/></c:legendEntry></c:legend>";
    }
    const posMap: Record<string, string> = {
      "left": "l",
      "right": "r", 
      "top": "t",
      "bottom": "b"
    };
    const pos = posMap[position || "right"] || "r";
    return `<c:legend><c:legendPos val="${pos}"/><c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.83229939874593417"/><c:y val="0.39325986532897839"/><c:w val="0.14285735948624242"/><c:h val="0.15355861408083918"/></c:manualLayout></c:layout><c:overlay val="0"/><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="3175"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="solid"/></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="735" b="0" i="0" u="none" strike="noStrike" baseline="0"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/></a:defRPr></a:pPr><a:endParaRPr lang="zh-CN"/></a:p></c:txPr></c:legend>`;
  }

  private static getCatAxXml(): string {
    return `<c:catAx><c:axId val="107911040"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln w="3175"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="solid"/></a:ln></c:spPr><c:txPr><a:bodyPr rot="0" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800" b="0" i="0" u="none" strike="noStrike" baseline="0"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr><c:crossAx val="107912576"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:tickLblSkip val="1"/><c:tickMarkSkip val="1"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  }

  private static getValAxXml(chartType: string): string {
    const formatCode = chartType === "bar" || chartType === "column" || chartType === "line" || chartType === "area" ? "0%" : "General";
    return `<c:valAx><c:axId val="107912576"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines><c:spPr><a:ln w="3175"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="solid"/></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="${formatCode}" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln w="3175"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="solid"/></a:ln></c:spPr><c:txPr><a:bodyPr rot="0" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800" b="0" i="0" u="none" strike="noStrike" baseline="0"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr><c:crossAx val="107911040"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  }

  static generateChartStyleXml(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:style xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1400" b="1" u="none" strike="no" schemeClr="tx1"/></a:r></a:p></a:rich></c:tx><c:layout/><c:overlay val="0"/></c:title><c:legend><c:legendPos val="r"/><c:layout/><c:overlay val="0"/><c:legendEntry><c:idx val="0"/><c:delete val="0"/></c:legendEntry></c:legend><c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:x val="0"/><c:y val="0"/><c:w val="1"/><c:h val="1"/></c:manualLayout></c:layout><c:barChart><c:gapWidth val="150"/><c:seriesLines><c:spPr><a:ln w="9525" cap="round" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:prstDash val="solid"/></a:ln></c:spPr></c:seriesLines></c:barChart></c:plotArea></c:style>`;
  }

  static generateColorsXml(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:colorMapping xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:bgClr><a:schemeClr val="bg1"/></c:bgClr><c:txClr><a:schemeClr val="tx1"/></c:txClr><c:accent1><a:schemeClr val="accent1"/></c:accent1><c:accent2><a:schemeClr val="accent2"/></c:accent2><c:accent3><a:schemeClr val="accent3"/></c:accent3><c:accent4><a:schemeClr val="accent4"/></c:accent4><c:accent5><a:schemeClr val="accent5"/></c:accent5><c:accent6><a:schemeClr val="accent6"/></c:accent6><c:hardClr><a:schemeClr val="accent1"/></c:hardClr></c:colorMapping>`;
  }
}
