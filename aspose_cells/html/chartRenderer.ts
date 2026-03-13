import type { ChartInfo, ChartSeries } from "../types";

const EMU_PER_PIXEL = 914400 / 96;

interface ChartArea {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface PlotArea {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface SeriesInfo {
  name: string;
  color: string;
  values: number[];
}

interface LegendInfo {
  position: "right" | "left" | "top" | "bottom";
  series: SeriesInfo[];
}

interface AxisInfo {
  position: "left" | "right" | "top" | "bottom";
  labels: string[];
  labelPositions: number[];
  tickPositions: number[];
  title?: string;
  minValue?: number;
  maxValue?: number;
  isPercentage?: boolean;
}

export class ChartRenderer {
  private colors = [
    "#9999FF",
    "#993366",
    "#C0C0C0",
    "#FF0000",
    "#00FF00",
    "#0000FF",
  ];

  renderChart(chart: ChartInfo, worksheet: any): string {
    const {
      fromCol,
      fromRow,
      toCol,
      toRow,
      fromColOff = 0,
      fromRowOff = 0,
      toColOff = 0,
      toRowOff = 0,
    } = chart;

    const defaultWidth = 48; // Excel default column width approximation
    const defaultHeight = 12.75;

    let left = 0;
    for (let col = 0; col < fromCol; col++) {
      left += worksheet.getColumnWidth(col) || defaultWidth;
    }
    left += fromColOff / EMU_PER_PIXEL;

    let top = 0;
    for (let row = 0; row < fromRow; row++) {
      top += worksheet.getRowHeight(row) || defaultHeight;
    }
    top += fromRowOff / EMU_PER_PIXEL;

    let right = 0;
    for (let col = 0; col < toCol; col++) {
      right += worksheet.getColumnWidth(col) || defaultWidth;
    }
    right += toColOff / EMU_PER_PIXEL;

    let bottom = 0;
    for (let row = 0; row < toRow; row++) {
      bottom += worksheet.getRowHeight(row) || defaultHeight;
    }
    bottom += toRowOff / EMU_PER_PIXEL;

    const width = right - left;
    const height = bottom - top;

    const chartSvg = this.generateChartSvg(chart, width, height);

    return `<span style='mso-ignore:vglayout;position:absolute;z-index:1;margin-left:${left}px;margin-top:${top}px;width:${width}px;height:${height}px'>${chartSvg}</span>`;
  }

  private escapeHtml(str: string): string {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  private generateChartSvg(
    chart: ChartInfo,
    width: number,
    height: number,
  ): string {
    const {
      chartType,
      title,
      series,
      legendPosition = "bottom",
      categoryAxisPosition = "bottom",
      valueAxisPosition = "left",
      plotArea: chartPlotArea,
    } = chart;

    const chartArea: ChartArea = {
      x: 0,
      y: 0,
      width: width,
      height: height,
    };

    const titleHeight = title ? 30 : 0;

    const leftMarginPct = 0.15;
    const rightMarginPct = 0.12;
    const topMarginPct = 0.08;
    const bottomMarginPct = 0.12;
    const legendWidthPct = 0.15;
    const legendHeightPct = 0.12;

    const leftMargin = width * leftMarginPct;
    const rightMargin = width * rightMarginPct;
    const topMargin = height * topMarginPct;
    const bottomMargin = height * bottomMarginPct;
    const legendWidth = width * legendWidthPct;
    const legendHeight = height * legendHeightPct;

    let adjustedRightMargin = rightMargin;
    let adjustedBottomMargin = bottomMargin;
    if (legendPosition === "right") {
      adjustedRightMargin = rightMargin + legendWidth;
    } else if (legendPosition === "bottom") {
      adjustedBottomMargin = bottomMargin + legendHeight;
    }

    let plotLeft: number,
      plotTop: number,
      plotWidth: number,
      plotHeight: number;

    // Use Excel manualLayout if available, otherwise use proportional defaults
    let plotBottom: number;
    const isPieChart = chartType === "pie" || chartType === "doughnut";

    if (chartPlotArea && !isPieChart) {
      plotLeft = chartPlotArea.x * width;
      plotTop = chartPlotArea.y * height;
      plotWidth = chartPlotArea.width * width;
      plotHeight = chartPlotArea.height * height;
      plotBottom = plotTop + plotHeight;
    } else {
      plotWidth = width - leftMargin - adjustedRightMargin;
      plotLeft = leftMargin;
      plotTop = title ? topMargin + titleHeight : topMargin;
      plotBottom = height - adjustedBottomMargin;
      plotHeight = plotBottom - plotTop;
    }

    const plotArea: PlotArea = {
      x: plotLeft,
      y: plotTop,
      width: plotWidth,
      height: plotHeight,
    };

    const seriesInfo: SeriesInfo[] = series.map((s, i) => ({
      name: s.name || `Series ${i + 1}`,
      color: this.colors[i % this.colors.length],
      values: s.values,
    }));

    const legend: LegendInfo = {
      position: legendPosition,
      series: seriesInfo,
    };

    const categories = series[0]?.categories || [];
    const rawMaxValue = Math.max(...series.flatMap((s) => s.values), 1);
    const isPercentage = rawMaxValue <= 1;
    const maxValue = isPercentage ? 0.25 : rawMaxValue;

    const xAxis: AxisInfo = {
      position: "bottom",
      labels: categories,
      labelPositions: categories.map((_, i) => {
        const categorySpace = plotWidth / categories.length;
        return plotLeft + i * categorySpace + categorySpace / 2;
      }),
      tickPositions: categories.map((_, i) => {
        const categorySpace = plotWidth / categories.length;
        return plotLeft + i * categorySpace + categorySpace / 2;
      }),
    };

    const numGridLines = 5;
    const yAxis: AxisInfo = {
      position: "left",
      labels: [],
      labelPositions: [],
      tickPositions: [],
      minValue: 0,
      maxValue: maxValue,
      isPercentage: isPercentage,
    };
    for (let i = 0; i <= numGridLines; i++) {
      const ratio = i / numGridLines;
      const value = isPercentage ? ratio * 25 : ratio * maxValue;
      yAxis.labels.push(
        isPercentage ? `${value.toFixed(0)}%` : value.toFixed(0),
      );
      const y = plotBottom - ratio * plotHeight;
      yAxis.labelPositions.push(y);
      yAxis.tickPositions.push(y);
    }

    let chartContent = "";

    if (chartType === "pie" || chartType === "doughnut") {
      chartContent = this.generatePieChart(
        series,
        plotWidth,
        plotHeight,
        plotLeft,
        plotTop,
        chartType === "doughnut",
      );
    } else if (chartType === "bar" || chartType === "column") {
      chartContent = this.generateBarChart(
        series,
        plotWidth,
        plotHeight,
        plotLeft,
        plotTop,
        plotBottom,
        xAxis,
        yAxis,
        chartType === "bar",
      );
    } else if (chartType === "line") {
      chartContent = this.generateLineChart(
        series,
        plotWidth,
        plotHeight,
        plotLeft,
        plotTop,
        plotBottom,
        xAxis,
        yAxis,
      );
    } else if (chartType === "area") {
      chartContent = this.generateAreaChart(
        series,
        plotWidth,
        plotHeight,
        plotLeft,
        plotTop,
        plotBottom,
        xAxis,
        yAxis,
      );
    } else if (chartType === "scatter") {
      chartContent = this.generateScatterChart(
        series,
        plotWidth,
        plotHeight,
        plotLeft,
        plotTop,
        plotBottom,
        xAxis,
        yAxis,
      );
    } else {
      chartContent = this.generateBarChart(
        series,
        plotWidth,
        plotHeight,
        plotLeft,
        plotTop,
        plotBottom,
        xAxis,
        yAxis,
        false,
      );
    }

    const titleSvg = title
      ? `<text x="${width / 2}" y="20" text-anchor="middle" font-size="14" font-family="Arial" font-weight="bold">${this.escapeHtml(title)}</text>`
      : "";

    const legendSvg = this.generateLegend(
      series,
      width,
      height,
      legend,
      plotLeft,
      plotTop,
      plotWidth,
      plotBottom,
    );

    const showPlotArea = !["pie", "doughnut", "area"].includes(chartType);

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="${width}pt" height="${height}pt">
 <g id="SFixTitle" />
 <g id="SContent">
  <g transform="scale(1.33333)">
   <rect width="100%" height="100%" fill="white" />
   <g>
    <g>
     <path d="M0,0 L${width},0 L${width},${height} L0,${height} Z " fill="#FFFFFF" />
     ${
       showPlotArea
         ? `<path d="M${plotLeft},${plotTop} L${plotLeft + plotWidth},${plotTop} L${plotLeft + plotWidth},${plotBottom} L${plotLeft},${plotBottom} Z " fill="#C0C0C0" />
     <path d="M${plotLeft},${plotTop} L${plotLeft + plotWidth},${plotTop} L${plotLeft + plotWidth},${plotBottom} L${plotLeft},${plotBottom} Z " class="p1025_0" fill="none" />`
         : ""
     }
     ${chartContent}
     ${
       showPlotArea
         ? `<path d="M${plotLeft},${plotBottom} L${plotLeft + plotWidth},${plotBottom} " class="p1025_1" fill="none" />
     <path d="M${plotLeft},${plotBottom} L${plotLeft},${plotBottom + 2} " class="p1025_1" fill="none" />
     <path d="M${plotLeft},${plotTop} L${plotLeft},${plotBottom} " class="p1025_1" fill="none" />`
         : ""
     }
     ${titleSvg}
     ${legendSvg}
     <path d="M0,0 L${width},0 L${width},${height} L0,${height} Z " class="p1025_1" fill="none" />
    </g>
   </g>
  </g>
 </g>
 <defs />
 <defs>
  <style type="text/css">
   <![CDATA[
   
.f1025_0
{
font-family:Arial;
font-size:8px;
fill:#000000;
}

.f1025_1
{
font-family:Arial;
font-size:7.349999905px;
fill:#000000;
}

.p1025_0
{
stroke:#808080;
stroke-width:1px;
stroke-linecap:butt;
stroke-linejoin:round;
}

.p1025_1
{
stroke:#000000;
stroke-width:0.25px;
stroke-linecap:butt;
stroke-linejoin:round;
}

.p1025_2
{
stroke:#000000;
stroke-width:1px;
stroke-linecap:butt;
stroke-linejoin:round;
}

   ]]></style>
 </defs>
</svg>`;
  }

  private generatePieChart(
    series: ChartSeries[],
    plotWidth: number,
    plotHeight: number,
    plotLeft: number,
    plotTop: number,
    isDoughnut: boolean,
  ): string {
    const allValues = series.flatMap((s) => s.values);
    const total = allValues.reduce((a, b) => a + b, 0);
    const radius = Math.min(plotWidth, plotHeight) / 2 - 10;
    const cx = plotLeft + plotWidth / 2;
    const cy = plotTop + plotHeight / 2;
    const innerRadius = isDoughnut ? radius * 0.5 : 0;

    let paths = "";
    let startAngle = 0;

    const categories = series[0]?.categories || [];

    for (let i = 0; i < allValues.length; i++) {
      const value = allValues[i];
      const angle = (value / total) * 2 * Math.PI;
      const endAngle = startAngle + angle;

      const x1 = cx + radius * Math.cos(startAngle);
      const y1 = cy + radius * Math.sin(startAngle);
      const x2 = cx + radius * Math.cos(endAngle);
      const y2 = cy + radius * Math.sin(endAngle);

      const ix1 = cx + innerRadius * Math.cos(startAngle);
      const iy1 = cy + innerRadius * Math.sin(startAngle);
      const ix2 = cx + innerRadius * Math.cos(endAngle);
      const iy2 = cy + innerRadius * Math.sin(endAngle);

      const largeArc = angle > Math.PI ? 1 : 0;

      const labelRadius = radius * 0.75;
      const labelX = cx + labelRadius * Math.cos(startAngle + angle / 2);
      const labelY = cy + labelRadius * Math.sin(startAngle + angle / 2);

      let d: string;
      if (isDoughnut) {
        d = `M ${ix1} ${iy1} L ${x1} ${y1} A ${radius} ${radius} 0 ${largeArc} 1 ${x2} ${y2} L ${ix2} ${iy2} A ${innerRadius} ${innerRadius} 0 ${largeArc} 0 ${ix1} ${iy1}`;
      } else {
        d = `M ${cx} ${cy} L ${x1} ${y1} A ${radius} ${radius} 0 ${largeArc} 1 ${x2} ${y2} Z`;
      }

      paths += `<path d="${d}" fill="${this.colors[i % this.colors.length]}" stroke="white" stroke-width="1"/>`;
      paths += `<text x="${labelX}" y="${labelY}" text-anchor="middle" font-size="11" font-family="Arial" fill="white">${((value / total) * 100).toFixed(0)}%</text>`;

      startAngle = endAngle;
    }

    return paths;
  }

  private generateBarChart(
    series: ChartSeries[],
    plotWidth: number,
    plotHeight: number,
    plotLeft: number,
    plotTop: number,
    plotBottom: number,
    xAxis: AxisInfo,
    yAxis: AxisInfo,
    isHorizontal: boolean,
  ): string {
    const categories = series[0]?.categories || [];
    const rawMaxValue = Math.max(...series.flatMap((s) => s.values), 1);
    const isPercentage = rawMaxValue <= 1;
    const maxValue = isPercentage ? 0.25 : rawMaxValue;
    const barCount = categories.length;
    const categorySpace = plotWidth / barCount;
    const barWidth = categorySpace * 0.35;
    const barGap = categorySpace * 0.1;

    let paths = "";
    let clipDefs = "";
    let barPaths = "";

    clipDefs = `<defs>
       <clipPath id="CLIP1025_0">
        <path d="M${plotLeft - 5},${plotTop - 5} L${plotLeft + plotWidth + 5},${plotTop - 5} L${plotLeft + plotWidth + 5},${plotBottom + 5} L${plotLeft - 5},${plotBottom + 5} Z " />
       </clipPath>
      </defs>`;

    for (let i = 0; i < categories.length; i++) {
      for (let j = 0; j < series.length; j++) {
        const value = series[j].values[i] || 0;
        const barHeight = (value / maxValue) * plotHeight;

        let x: number, y: number, w: number, h: number;
        if (isHorizontal) {
          x = plotLeft;
          y = plotTop + i * categorySpace + j * (barWidth + barGap);
          w = barHeight;
          h = barWidth;
        } else {
          x = plotLeft + i * categorySpace + j * (barWidth + barGap);
          y = plotBottom - barHeight;
          w = barWidth;
          h = barHeight;
        }

        barPaths += `<path d="M${x},${y} L${x + w},${y} L${x + w},${y + h} L${x},${y + h} Z " fill="${this.colors[j % this.colors.length]}" />`;
        barPaths += `<path d="M${x},${y} L${x + w},${y} L${x + w},${y + h} L${x},${y + h} Z " class="p1025_2" fill="none" />`;
      }
    }

    if (!isHorizontal) {
      for (let i = 0; i < categories.length; i++) {
        const xPos = plotLeft + i * categorySpace + categorySpace / 2;
        const yPos = plotBottom + 10;
        paths += `<g><g><g transform="matrix(1,0,0,1,${xPos},${yPos})"><text x="0" y="-1.6953125" class="f1025_0">${this.escapeHtml(categories[i])}</text></g></g></g>`;
      }
    }

    if (isHorizontal) {
      for (let i = 0; i < categories.length; i++) {
        const yPos = plotTop + i * categorySpace + categorySpace / 2 + 4;
        paths += `<g><g><g transform="matrix(1,0,0,1,${plotLeft - 15},${yPos})"><text x="0" y="-1.6953125" class="f1025_0">${this.escapeHtml(categories[i])}</text></g></g></g>`;
      }
    }

    const numGridLines = 5;
    for (let i = 0; i <= numGridLines; i++) {
      const y = plotBottom - (i / numGridLines) * plotHeight;
      paths += `<path d="M${plotLeft},${y} L${plotLeft + plotWidth},${y} " class="p1025_1" fill="none" />`;

      if (!isHorizontal) {
        const labelValue = isPercentage
          ? ((i / numGridLines) * 25).toFixed(0) + "%"
          : ((i / numGridLines) * maxValue).toFixed(0);
        const labelX = plotLeft - 15;
        const labelY = y + 3;
        paths += `<g><g><g transform="matrix(1,0,0,1,${labelX},${labelY})"><text x="0" y="-1.6953125" class="f1025_0">${labelValue}</text></g></g></g>`;
      } else {
        const labelValue = isPercentage
          ? ((i / numGridLines) * 25).toFixed(0) + "%"
          : ((i / numGridLines) * maxValue).toFixed(0);
        const labelX = plotLeft + (i / numGridLines) * plotWidth;
        const labelY = plotBottom + 10;
        paths += `<g><g><g transform="matrix(1,0,0,1,${labelX},${labelY})"><text x="0" y="-1.6953125" class="f1025_0">${labelValue}</text></g></g></g>`;
      }
    }

    return (
      clipDefs + `<g clip-path="url(#CLIP1025_0)">` + barPaths + `</g>` + paths
    );
  }

  private generateLineChart(
    series: ChartSeries[],
    width: number,
    height: number,
    plotLeft: number,
    plotTop: number,
    plotBottom: number,
    xAxis: AxisInfo,
    yAxis: AxisInfo,
  ): string {
    const categories = series[0]?.categories || [];
    const rawMaxValue = Math.max(...series.flatMap((s) => s.values), 1);
    const isPercentage = rawMaxValue <= 1;
    const maxValue = isPercentage ? 0.25 : rawMaxValue;
    const minValue = 0;
    const valueRange = maxValue - minValue || 1;

    let paths = "";
    let dots = "";

    const plotHeight = plotBottom - plotTop;

    for (let j = 0; j < series.length; j++) {
      const values = series[j].values;
      let pathD = "";

      for (let i = 0; i < values.length; i++) {
        const x = xAxis.labelPositions[i];
        const y =
          plotBottom - ((values[i] - minValue) / valueRange) * plotHeight;

        if (i === 0) {
          pathD = `M ${x} ${y}`;
        } else {
          pathD += ` L ${x} ${y}`;
        }

        dots += `<circle cx="${x}" cy="${y}" r="4" fill="${this.colors[j % this.colors.length]}"/>`;
      }

      paths += `<path d="${pathD}" fill="none" stroke="${this.colors[j % this.colors.length]}" stroke-width="2"/>`;
    }

    for (let i = 0; i < categories.length; i++) {
      const x = 40 + (i / (categories.length - 1 || 1)) * (width - 50);
      paths += `<text x="${x}" y="${height - 5}" text-anchor="middle" font-size="10" font-family="Arial">${this.escapeHtml(categories[i])}</text>`;
    }

    return paths + dots;
  }

  private generateAreaChart(
    series: ChartSeries[],
    width: number,
    height: number,
    plotLeft: number,
    plotTop: number,
    plotBottom: number,
    xAxis: AxisInfo,
    yAxis: AxisInfo,
  ): string {
    const categories = series[0]?.categories || [];
    const maxValue = Math.max(...series.flatMap((s) => s.values), 1);
    const minValue = Math.min(...series.flatMap((s) => s.values), 0);
    const valueRange = maxValue - minValue || 1;

    let paths = "";

    for (let j = 0; j < series.length; j++) {
      const values = series[j].values;
      let pathD = "";

      for (let i = 0; i < values.length; i++) {
        const x = 40 + (i / (values.length - 1 || 1)) * (width - 50);
        const y =
          height - 20 - ((values[i] - minValue) / valueRange) * (height - 40);

        if (i === 0) {
          pathD = `M ${x} ${height - 20} L ${x} ${y}`;
        } else {
          pathD += ` L ${x} ${y}`;
        }
      }

      pathD += ` L ${40 + ((values.length - 1) / (values.length - 1 || 1)) * (width - 50)} ${height - 20} Z`;

      paths += `<path d="${pathD}" fill="${this.colors[j % this.colors.length]}" fill-opacity="0.5" stroke="${this.colors[j % this.colors.length]}" stroke-width="2"/>`;
    }

    for (let i = 0; i < categories.length; i++) {
      const x = 40 + (i / (categories.length - 1 || 1)) * (width - 50);
      paths += `<text x="${x}" y="${height - 5}" text-anchor="middle" font-size="10" font-family="Arial">${this.escapeHtml(categories[i])}</text>`;
    }

    return paths;
  }

  private generateScatterChart(
    series: ChartSeries[],
    width: number,
    height: number,
    plotLeft: number,
    plotTop: number,
    plotBottom: number,
    xAxis: AxisInfo,
    yAxis: AxisInfo,
  ): string {
    const allValues = series.flatMap((s) => s.values);
    const maxValue = Math.max(...allValues, 1);
    const minValue = Math.min(...allValues, 0);
    const valueRange = maxValue - minValue || 1;

    const categories = series[0]?.categories || [];
    const maxCat = categories.length || 1;

    let paths = "";
    let dots = "";

    const symbols = ["circle", "square", "triangle", "diamond"];
    const symbolSizes = [5, 6, 5, 5];

    for (let j = 0; j < series.length; j++) {
      const values = series[j].values;
      const symbol = symbols[j % symbols.length];
      const size = symbolSizes[j % symbolSizes.length];

      for (let i = 0; i < values.length; i++) {
        const x = 40 + (i / (maxCat - 1 || 1)) * (width - 50);
        const y =
          height - 20 - ((values[i] - minValue) / valueRange) * (height - 40);

        if (symbol === "circle") {
          dots += `<circle cx="${x}" cy="${y}" r="${size}" fill="${this.colors[j % this.colors.length]}"/>`;
        } else if (symbol === "square") {
          dots += `<rect x="${x - size / 2}" y="${y - size / 2}" width="${size}" height="${size}" fill="${this.colors[j % this.colors.length]}"/>`;
        } else if (symbol === "triangle") {
          const h = size * 1.5;
          dots += `<polygon points="${x},${y - h / 2} ${x - size / 2},${y + h / 2} ${x + size / 2},${y + h / 2}" fill="${this.colors[j % this.colors.length]}"/>`;
        } else {
          const s = size / 2;
          dots += `<polygon points="${x},${y - s} ${x + s},${y} ${x},${y + s} ${x - s},${y}" fill="${this.colors[j % this.colors.length]}"/>`;
        }
      }

      const pathValues = values;
      if (pathValues.length > 1) {
        let pathD = "";
        for (let i = 0; i < pathValues.length; i++) {
          const x = 40 + (i / (maxCat - 1 || 1)) * (width - 50);
          const y =
            height -
            20 -
            ((pathValues[i] - minValue) / valueRange) * (height - 40);
          pathD += i === 0 ? `M ${x} ${y}` : ` L ${x} ${y}`;
        }
        paths += `<path d="${pathD}" fill="none" stroke="${this.colors[j % this.colors.length]}" stroke-width="1"/>`;
      }
    }

    for (let i = 0; i < categories.length; i++) {
      const x = 40 + (i / (categories.length - 1 || 1)) * (width - 50);
      paths += `<text x="${x}" y="${height - 5}" text-anchor="middle" font-size="10" font-family="Arial">${this.escapeHtml(categories[i])}</text>`;
    }

    return paths + dots;
  }

  private generateLegend(
    series: ChartSeries[],
    width: number,
    height: number,
    legend: LegendInfo,
    plotLeft: number,
    plotTop: number,
    plotWidth: number,
    plotBottom: number,
  ): string {
    const titleHeight = 0;
    const topMargin = 10;
    const bottomMargin = 30;
    const legendGap = 10;

    let legendX: number, legendY: number;
    const legendBoxWidth = 50;
    const legendBoxHeight = series.length * 18 + 15;
    const legendItemWidth = legendBoxWidth - 20;

    switch (legend.position) {
      case "right":
        legendX = plotLeft + plotWidth + legendGap;
        legendY = plotTop + 10;
        if (legendX + legendBoxWidth > width) {
          legendX = width - legendBoxWidth - legendGap;
        }
        break;
      case "left":
        legendX = 10;
        legendY = plotTop + 10;
        break;
      case "top":
        legendX = plotLeft + plotWidth / 2 - legendBoxWidth / 2;
        legendY = titleHeight + topMargin;
        break;
      case "bottom":
      default:
        legendX = plotLeft + plotWidth / 2 - legendBoxWidth / 2;
        legendY = plotBottom + legendGap;
        if (legendY + legendBoxHeight > height) {
          legendY = height - legendBoxHeight - legendGap;
        }
        break;
    }

    let legendItems = "";

    const legendBoxX = legendX;
    const legendBoxY = legendY;

    legendItems += `<path d="M${legendBoxX},${legendBoxY} L${legendBoxX + legendBoxWidth},${legendBoxY} L${legendBoxX + legendBoxWidth},${legendBoxY + legendBoxHeight} L${legendBoxX},${legendBoxY + legendBoxHeight} Z " fill="#FFFFFF" />`;
    legendItems += `<path d="M${legendBoxX},${legendBoxY} L${legendBoxX + legendBoxWidth},${legendBoxY} L${legendBoxX + legendBoxWidth},${legendBoxY + legendBoxHeight} L${legendBoxX},${legendBoxY + legendBoxHeight} Z " class="p1025_1" fill="none" />`;

    for (let i = 0; i < series.length; i++) {
      const x = legendBoxX + 5;
      const y = legendBoxY + 10 + i * 20;
      legendItems += `<rect x="${x}" y="${y}" width="12" height="12" fill="${this.colors[i % this.colors.length]}"/>`;
      legendItems += `<rect x="${x}" y="${y}" width="12" height="12" class="p1025_2" fill="none" />`;
      legendItems += `<text x="${x + 16}" y="${y + 10}" font-size="7.35" font-family="Arial">${this.escapeHtml(series[i].name || `Series ${i + 1}`)}</text>`;
    }

    return `<g>${legendItems}</g>`;
  }
}
