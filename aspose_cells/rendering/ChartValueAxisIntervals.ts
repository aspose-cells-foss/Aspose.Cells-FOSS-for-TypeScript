/**
 * 计算数值轴的主刻度和次刻度间隔。
 * 使用基于 1-2-5 序列的标准数值轴刻度算法。
 */
export class ChartValueAxisIntervals {
    /**
     * 计算主要刻度间隔。
     */
    static ComputeMajorUnit(
        min: number,
        max: number,
        targetSteps: number = 5
    ): number {
        if (!isFinite(min) || !isFinite(max) || max <= min) {
            return 0;
        }
        const raw = (max - min) / targetSteps;
        if (raw <= 0) return 0;
        return ChartValueAxisIntervals.RoundToNice(raw);
    }

    /**
     * 根据主刻度计算次刻度间隔（默认为主刻度的 1/5）。
     */
    static ComputeMinorUnit(major: number): number {
        if (major <= 0) return 0;
        return major / 5;
    }

    /**
     * 将数值四舍五入到"漂亮"的刻度值（1、2、5 的倍数）。
     */
    static RoundToNice(value: number): number {
        if (value <= 0) return 0;
        const exp = Math.pow(10, Math.floor(Math.log10(value)));
        const f = value / exp;
        let niceF: number;
        if (f <= 1.5) {
            niceF = 1;
        } else if (f <= 3) {
            niceF = 2;
        } else if (f <= 7) {
            niceF = 5;
        } else {
            niceF = 10;
        }
        return niceF * exp;
    }
}
