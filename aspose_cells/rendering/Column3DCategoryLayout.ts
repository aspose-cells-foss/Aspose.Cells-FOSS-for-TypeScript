import type { ParsedChart } from "./ChartData.js";

/**
 * 3D 分类轴槽信息。
 */
export interface CategorySlot {
    center: number;
    label: string;
}

/**
 * 构建 3D 图表的分类轴槽信息。
 */
export class Column3DCategoryLayout {
    readonly SlotCount: number;
    readonly Centers: number[] = [];
    readonly Labels: string[] = [];

    constructor(chart: ParsedChart, count: number, _dateSystem: number) {
        this.SlotCount = count;
        for (let i = 0; i < count; i++) {
            this.Centers.push((i + 0.5) / count);
            this.Labels.push(i < chart.Categories.length ? chart.Categories[i] : "");
        }
    }
}
