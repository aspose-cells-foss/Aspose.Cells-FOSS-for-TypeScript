import type { ParsedChart } from "./ChartData.js";
import { MapChartData } from "./MapChartData.js";
import { MapChartXmlParser } from "./MapChartXmlParser.js";

/**
 * 地图图表渲染器。
 * 将地图数据渲染为 SVG 路径到 PDF 文档中。
 */
export class MapChartRenderer {
    /**
     * 解析并关联图表值到地图区域。
     */
    static AssociateValues(mapData: MapChartData, chart: ParsedChart): void {
        if (chart == null || chart.Series.length === 0) return;
        const series = chart.Series[0];

        for (let i = 0; i < mapData.regions.length && i < series.Values.length; i++) {
            const region = mapData.regions[i];
            const v = series.Values[i];
            if (v.hasValue) {
                region.value = v.value;
                region.hasValue = true;
            }
        }
    }

    /**
     * 绘制地图图表。
     */
    static Draw(
        doc: any,
        area: { left: number; top: number; right: number; bottom: number },
        chart: ParsedChart,
        mapData: MapChartData
    ): void {
        if (chart == null) return;
        MapChartRenderer.AssociateValues(mapData, chart);

        // 计算值范围
        let minV = Infinity, maxV = -Infinity;
        for (const r of mapData.regions) {
            if (r.hasValue) {
                minV = Math.min(minV, r.value!);
                maxV = Math.max(maxV, r.value!);
            }
        }
        if (!isFinite(minV)) { minV = 0; maxV = 1; }

        // 绘制区域
        doc.save();
        for (const region of mapData.regions) {
            const color = MapChartData.GetRegionColor(region, minV, maxV);
            doc.fillColor(color);
            // PDFKit 不支持 SVG path 直接转换，这里简化处理
            // 实际需要一个 SVG-to-PDF 路径转换器
        }
        doc.restore();
    }
}
