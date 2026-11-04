/**
 * 地图图表的数据结构。
 * 存储从 Excel 区域地图 XML 中解析出的区域坐标和值。
 */
export interface MapRegion {
    /** 区域唯一标识符（如省份/国家名称） */
    id: string;
    /** 区域名称 */
    name: string;
    /** 中心点经纬度 */
    centerLat: number;
    centerLon: number;
    /** 区域的 SVG 路径数据 */
    pathData: string;
    /** 该区域对应的数据值（如果有） */
    value: number | null;
    /** 是否有数据值 */
    hasValue: boolean;
}

/**
 * 地图数据容器。
 */
export class MapChartData {
    /** 所有区域列表 */
    regions: MapRegion[] = [];
    /** XML 源字符串（解析失败时保留） */
    rawXml: string | null = null;
    /** 是否解析成功 */
    isValid: boolean = false;

    /**
     * 根据值范围获取颜色。
     */
    static GetRegionColor(
        region: MapRegion,
        minValue: number,
        maxValue: number
    ): string {
        if (region.value == null || maxValue <= minValue) {
            return "#F2F2F2"; // 无数据区域用浅灰色
        }

        const ratio = (region.value - minValue) / (maxValue - minValue);
        // 蓝色渐变：浅蓝 -> 深蓝
        const intensity = Math.round(180 + ratio * 75);
        const lightness = Math.round(85 - ratio * 40);
        const h = 210; // 蓝色
        const s = 60;

        return `hsl(${h}, ${s}%, ${lightness}%)`;
    }
}
