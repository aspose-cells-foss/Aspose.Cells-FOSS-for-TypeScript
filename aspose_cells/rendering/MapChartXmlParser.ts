import { MapChartData, type MapRegion } from "./MapChartData.js";

/**
 * 解析 Excel 地图图表的 XML 数据。
 * 将 Excel 存储的 SVG 路径数据转换为 MapChartData。
 */
export class MapChartXmlParser {
    /**
     * 从原始 XML 字符串解析地图数据。
     */
    static Parse(xml: string | null): MapChartData {
        const data = new MapChartData();
        data.rawXml = xml;

        if (!xml) {
            data.isValid = false;
            return data;
        }

        try {
            data.regions = MapChartXmlParser.ExtractRegions(xml);
            data.isValid = data.regions.length > 0;
        } catch {
            data.isValid = false;
        }

        return data;
    }

    /**
     * 从 XML 中提取所有区域的路径和属性。
     */
    private static ExtractRegions(xml: string): MapRegion[] {
        const regions: MapRegion[] = [];

        // 简单解析：查找 <path> 或 <Region> 元素
        const pathRegex = /<path[^>]*d="([^"]+)"[^>]*>/gi;
        const idRegex = /id="([^"]+)"/i;
        const nameRegex = /name="([^"]+)"/i;

        let match: RegExpExecArray | null;
        while ((match = pathRegex.exec(xml)) !== null) {
            const pathData = match[1];
            const segment = match[0];

            const idMatch = idRegex.exec(segment);
            const nameMatch = nameRegex.exec(segment);

            regions.push({
                id: idMatch ? idMatch[1] : `region_${regions.length}`,
                name: nameMatch ? nameMatch[1] : idMatch ? idMatch[1] : "",
                centerLat: 0,
                centerLon: 0,
                pathData,
                value: null,
                hasValue: false
            });
        }

        // 如果没找到 path，尝试查找 Region 元素
        if (regions.length === 0) {
            const regionRegex = /<Region[^>]*>/gi;
            while ((match = regionRegex.exec(xml)) !== null) {
                const segment = match[0];
                regions.push({
                    id: MapChartXmlParser.Attr(segment, "id") || `region_${regions.length}`,
                    name: MapChartXmlParser.Attr(segment, "name") || "",
                    centerLat: 0,
                    centerLon: 0,
                    pathData: "",
                    value: null,
                    hasValue: false
                });
            }
        }

        return regions;
    }

    private static Attr(tag: string, name: string): string | null {
        const re = new RegExp(`${name}="([^"]*)"`, "i");
        const m = re.exec(tag);
        return m ? m[1] : null;
    }
}
