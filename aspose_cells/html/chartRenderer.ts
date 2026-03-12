import type { ChartInfo, ChartSeries } from "../types";

const EMU_PER_PIXEL = 914400 / 96;

export class ChartRenderer {
  private colors = [
    "#4472C4",
    "#ED7D31",
    "#A5A5A5",
    "#FFC000",
    "#5B9BD5",
    "#70AD47",
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

    const defaultWidth = 64;
    const defaultHeight = 16;

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
    const { chartType, title, series } = chart;

    const titleHeight = title ? 30 : 0;
    const legendHeight = 30;
    const plotHeight = height - titleHeight - legendHeight;
    const plotWidth = width - 40;
    const plotLeft = 40;
    const plotTop = titleHeight;

    let chartContent = "";

    if (chartType === "pie" || chartType === "doughnut") {
      chartContent = this.generatePieChart(
        series,
        plotWidth,
        plotHeight,
        chartType === "doughnut",
      );
    } else if (chartType === "bar" || chartType === "column") {
      chartContent = this.generateBarChart(
        series,
        plotWidth,
        plotHeight,
        chartType === "bar",
      );
    } else if (chartType === "line") {
      chartContent = this.generateLineChart(series, plotWidth, plotHeight);
    } else if (chartType === "area") {
      chartContent = this.generateAreaChart(series, plotWidth, plotHeight);
    } else if (chartType === "scatter") {
      chartContent = this.generateScatterChart(series, plotWidth, plotHeight);
    } else {
      chartContent = this.generateBarChart(
        series,
        plotWidth,
        plotHeight,
        false,
      );
    }

    const titleSvg = title
      ? `<text x="${width / 2}" y="20" text-anchor="middle" font-size="14" font-family="Arial" font-weight="bold">${this.escapeHtml(title)}</text>`
      : "";

    const legendSvg = this.generateLegend(series, width, height, legendHeight);

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
  <rect width="100%" height="100%" fill="white"/>
  ${titleSvg}
  <g transform="translate(${plotLeft}, ${plotTop})">
    ${chartContent}
  </g>
  ${legendSvg}
</svg>`;
  }

  private generatePieChart(
    series: ChartSeries[],
    width: number,
    height: number,
    isDoughnut: boolean,
  ): string {
    const allValues = series.flatMap((s) => s.values);
    const total = allValues.reduce((a, b) => a + b, 0);
    const radius = Math.min(width, height) / 2 - 10;
    const cx = width / 2;
    const cy = height / 2;
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
    width: number,
    height: number,
    isHorizontal: boolean,
  ): string {
    const categories = series[0]?.categories || [];
    const maxValue = Math.max(...series.flatMap((s) => s.values), 1);
    const barCount = categories.length;
    const barWidth = (width - 40) / barCount - 4;
    const barGroupWidth = barWidth / series.length;

    let paths = "";

    for (let i = 0; i < categories.length; i++) {
      for (let j = 0; j < series.length; j++) {
        const value = series[j].values[i] || 0;
        const barHeight = (value / maxValue) * (height - 40);

        let x: number, y: number, w: number, h: number;
        if (isHorizontal) {
          x = 40;
          y = 20 + i * (barWidth + 4) + j * barGroupWidth;
          w = barHeight;
          h = barGroupWidth - 2;
        } else {
          x = 40 + i * (barWidth + 4) + j * barGroupWidth;
          y = height - 20 - barHeight;
          w = barGroupWidth - 2;
          h = barHeight;
        }

        paths += `<rect x="${x}" y="${y}" width="${w}" height="${h}" fill="${this.colors[j % this.colors.length]}" stroke="white" stroke-width="1"/>`;
      }

      if (!isHorizontal) {
        paths += `<text x="${40 + i * (barWidth + 4) + barWidth / 2}" y="${height - 5}" text-anchor="middle" font-size="10" font-family="Arial">${this.escapeHtml(categories[i])}</text>`;
      }
    }

    if (isHorizontal) {
      for (let i = 0; i < categories.length; i++) {
        paths += `<text x="35" y="${20 + i * (barWidth + 4) + barWidth / 2 + 4}" text-anchor="end" font-size="10" font-family="Arial">${this.escapeHtml(categories[i])}</text>`;
      }
    }

    return paths;
  }

  private generateLineChart(
    series: ChartSeries[],
    width: number,
    height: number,
  ): string {
    const categories = series[0]?.categories || [];
    const maxValue = Math.max(...series.flatMap((s) => s.values), 1);
    const minValue = Math.min(...series.flatMap((s) => s.values), 0);
    const valueRange = maxValue - minValue || 1;

    let paths = "";
    let dots = "";

    for (let j = 0; j < series.length; j++) {
      const values = series[j].values;
      let pathD = "";

      for (let i = 0; i < values.length; i++) {
        const x = 40 + (i / (values.length - 1 || 1)) * (width - 50);
        const y =
          height - 20 - ((values[i] - minValue) / valueRange) * (height - 40);

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
    legendHeight: number,
  ): string {
    const legendY = height - legendHeight + 10;
    let legendItems = "";

    for (let i = 0; i < series.length; i++) {
      const x = 10 + i * (width / series.length);
      legendItems += `<rect x="${x}" y="${legendY}" width="12" height="12" fill="${this.colors[i % this.colors.length]}"/>`;
      legendItems += `<text x="${x + 16}" y="${legendY + 10}" font-size="11" font-family="Arial">${this.escapeHtml(series[i].name || `Series ${i + 1}`)}</text>`;
    }

    return `<g>${legendItems}</g>`;
  }
}
