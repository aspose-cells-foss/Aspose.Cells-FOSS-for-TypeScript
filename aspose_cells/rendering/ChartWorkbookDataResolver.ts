/**
 * 解析图表工作簿中的数据引用。
 * 将 Excel 单元格引用解析为实际值，填充到 ParsedChart 中。
 */
export class ChartWorkbookDataResolver {
    /**
     * 从二维单元格数据中提取图表值。
     * @param cellData 二维单元格值数组
     * @param headerRow 是否第一行是标题
     */
    static ExtractSeriesFromCellGrid(
        cellData: Array<Array<{ value: number | string | null; hasValue: boolean }>>,
        headerRow: boolean = true
    ): {
        categories: string[];
        series: Array<{ name: string; values: Array<{ hasValue: boolean; value: number }> }>;
    } {
        if (cellData.length === 0) {
            return { categories: [], series: [] };
        }

        const dataRows = headerRow ? cellData.slice(1) : cellData;
        const categories: string[] = [];

        // 第一列作为分类
        for (const row of dataRows) {
            if (row.length > 0 && row[0].hasValue && typeof row[0].value === "string") {
                categories.push(row[0].value as string);
            } else {
                categories.push("");
            }
        }

        // 其余列作为序列
        const series: Array<{ name: string; values: Array<{ hasValue: boolean; value: number }> }> = [];
        const startCol = 1;

        for (let col = startCol; col < (headerRow ? cellData[0].length : cellData[0].length); col++) {
            const seriesName = headerRow && cellData[0][col]?.hasValue
                ? String(cellData[0][col].value)
                : `Series ${col}`;

            const values: Array<{ hasValue: boolean; value: number }> = [];
            for (let rowIdx = 0; rowIdx < dataRows.length; rowIdx++) {
                const cell = dataRows[rowIdx][col];
                if (cell && cell.hasValue && typeof cell.value === "number") {
                    values.push({ hasValue: true, value: cell.value });
                } else {
                    values.push({ hasValue: false, value: 0 });
                }
            }
            series.push({ name: seriesName, values });
        }

        return { categories, series };
    }
}
