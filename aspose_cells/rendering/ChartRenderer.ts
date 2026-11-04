import type { Rect, ParsedChart } from "./ChartData.js";
import { ChartKind, ChartValue, ChartSeries, type TextAlign, type VAlign, ToPdfKitColor } from "./ChartData.js";

type PdfDoc = any;
import { ChartValueAxisIntervals } from "./ChartValueAxisIntervals.js";
import { ChartValueAxisLineRenderer } from "./ChartValueAxisLineRenderer.js";
import { ChartCategoryLabels } from "./ChartCategoryLabels.js";
import { ChartStackingMath } from "./ChartStackingMath.js";
import { ChartBarBorder } from "./ChartBarBorder.js";
import { ChartExplicitFormatting } from "./ChartExplicitFormatting.js";
import { ChartPatternRenderer } from "./ChartPatternRenderer.js";
import { ChartSoftEdgeRenderer } from "./ChartSoftEdgeRenderer.js";
import { ChartTextStyleResolver } from "./ChartTextStyleResolver.js";
import { ChartDateAxisLayout } from "./ChartDateAxisLayout.js";
import { ChartDateAxisMath } from "./ChartDateAxisMath.js";
import { ChartSeriesPointBuilder } from "./ChartSeriesPointBuilder.js";
import { ChartSeriesStrokeRenderer } from "./ChartSeriesStrokeRenderer.js";
import { LineChartAutomaticLayout } from "./LineChartAutomaticLayout.js";
import { LineChartCategoryAxisRenderer } from "./LineChartCategoryAxisRenderer.js";
import { LineChartDecorationLayout } from "./LineChartDecorationLayout.js";
import { ScatterChartAxisMath } from "./ScatterChartAxisMath.js";
import { ScatterChartPlotLayout } from "./ScatterChartPlotLayout.js";
import { ScatterSmoothPathBuilder } from "./ScatterSmoothPathBuilder.js";
import { BubbleChartMath } from "./BubbleChartMath.js";
import { DoughnutChartRenderer } from "./DoughnutChartRenderer.js";
import { Pie3DChartRenderer } from "./Pie3DChartRenderer.js";
import { Area3DChartRenderer } from "./Area3DChartRenderer.js";
import { Line3DChartRenderer } from "./Line3DChartRenderer.js";
import { Bar3DChartRenderer } from "./Bar3DChartRenderer.js";
import { Column3DChartRenderer } from "./Column3DChartRenderer.js";

/**
 * 核心图表渲染器。
 * 将 ParsedChart 渲染为 PDF 内容。
 * 调度所有子渲染器处理不同图表类型。
 */
export class ChartRenderer {
    private labelSizePt: number = 9;
    private _suppressCategoryAxisLabels: boolean = false;
    private _chartCulture: string | null = null;

    constructor(
        private readonly _doc: PdfDoc,
        private readonly _dateSystem: number = 1900
    ) { }

    // ========== 主入口 ==========

    Draw(
        area: Rect,
        chart: ParsedChart,
        suppressCategoryAxisLabels: boolean = false
    ): void {
        if (!chart) return;
        this._chartCulture = chart.Culture;
        this.labelSizePt = chart.TextSizePt ?? 9;
        this._suppressCategoryAxisLabels = suppressCategoryAxisLabels;

        this.DrawFrame(area, chart);

        // 饼图
        if (chart.Kind === ChartKind.Pie) {
            this.DrawPie(area, chart);
            return;
        }

        // 无序列，只画标题
        if (chart.Series.length === 0) {
            this.DrawTitle(area, chart, (area.top + area.bottom) / 2);
            return;
        }

        let plot = this.ResolvePlotRect(area, chart);

        // 散点图
        if (chart.Kind === ChartKind.Scatter) {
            this.DrawScatter(area, plot, chart);
            return;
        }

        // 雷达图
        if (chart.Kind === ChartKind.Radar) {
            this.DrawRadar(area, plot, chart);
            return;
        }

        // 计算值轴
        const { min, max, major } = this.ComputeValueAxis(chart);

        // 标题
        let titleCenterY = area.top + (plot.top - area.top) * 0.5;
        if (chart.HasDataTable) titleCenterY = area.top + area.height * 0.08;
        else if (chart.Kind === ChartKind.Column && chart.IsBar3D) titleCenterY = area.top + area.height * 0.08;

        this.DrawTitle(area, chart, titleCenterY);

        // 3D 图表
        if (chart.Kind === ChartKind.Bar && chart.IsBar3D) {
            new Bar3DChartRenderer().Draw(this._doc, area, plot, chart, min, max, major);
            this.DrawLegend(area, plot, chart);
            return;
        }
        if (chart.Kind === ChartKind.Column && chart.IsBar3D) {
            new Column3DChartRenderer().Draw(this._doc, area, chart, min, max, major);
            this.DrawLegend(area, plot, chart);
            return;
        }
        if (chart.Kind === ChartKind.Area && chart.IsArea3D) {
            new Area3DChartRenderer().Draw(this._doc, area, plot, chart, min, max, major);
            this.DrawLegend(area, plot, chart);
            return;
        }
        if (chart.Kind === ChartKind.Line && chart.IsLine3D) {
            new Line3DChartRenderer().Draw(this._doc, null as any, plot, chart, min, max, major);
            this.DrawLegend(area, plot, chart);
            return;
        }

        // 常规 2D 图表
        this.DrawValueAxis(area, plot, chart, min, max, major);
        this.DrawCategoryAxis(area, plot, chart);

        if (chart.Kind === ChartKind.Column) {
            this.DrawColumns(plot, chart, min, max);
        } else if (chart.Kind === ChartKind.Bar) {
            this.DrawBars(plot, chart, min, max);
        } else {
            this.DrawSeries(plot, chart, min, max);
        }

        this.DrawAxisTitles(area, plot, chart);
        this.DrawPlotAreaBorder(area, plot, chart);
        this.DrawDataTable(area, plot, chart);
        this.DrawLegend(area, plot, chart);
    }

    // ========== 框架 ==========

    DrawFrame(area: Rect, chart: ParsedChart): void {
        if (chart == null) return;
        const doc = this._doc;
        doc.save();

        // 背景填充
        if (chart.HasChartAreaFill) {
            doc.fillColor(chart.ChartAreaFillColor.toPdfKitColor());
            doc.rect(area.left, area.top, area.right - area.left, area.bottom - area.top);
            doc.fill();
        } else {
            doc.fillColor("#FFFFFF");
            doc.rect(area.left, area.top, area.right - area.left, area.bottom - area.top);
            doc.fill();
        }

        // 边框
        const borderColor = chart.HasChartAreaBorder
            ? chart.ChartAreaBorderColor.toPdfKitColor()
            : chart.RoundedCorners ? "#888888" : "#D9D9D9";
        const borderWidth = Math.max(0.5, chart.ChartAreaBorderWidthPt || 1);
        doc.lineWidth(borderWidth);
        doc.strokeColor(borderColor);

        if (chart.RoundedCorners) {
            doc.roundedRect(area.left, area.top, area.right - area.left, area.bottom - area.top, 10);
        } else {
            doc.rect(area.left, area.top, area.right - area.left, area.bottom - area.top);
        }
        doc.stroke();

        doc.restore();
    }

    // ========== 绘图区域 ==========

    private ResolvePlotRect(area: Rect, chart: ParsedChart): Rect {
        if (chart.Kind === ChartKind.Scatter) {
            return ScatterChartPlotLayout.Resolve(area, chart);
        }

        if (chart.Kind === ChartKind.Radar) {
            return {
                left: area.left + area.width * 0.11,
                top: area.top + area.height * 0.1,
                right: area.right - area.width * 0.15,
                bottom: area.bottom - area.height * 0.12
            };
        }

        if (chart.HasManualPlot) {
            const w = area.right - area.left;
            const h = area.bottom - area.top;
            return {
                left: area.left + chart.PlotX * w,
                top: area.top + chart.PlotY * h,
                right: area.left + (chart.PlotX + chart.PlotW) * w,
                bottom: area.top + (chart.PlotY + chart.PlotH) * h
            };
        }

        const axisLabelMaxWidth = this.MeasureMaxValueAxisLabelWidth(chart);
        const leftPadding = Math.max((area.right - area.left) * 0.06, axisLabelMaxWidth + 14);
        const hasTitle = chart.Title != null;
        const topPadding = (area.right - area.left) * 0 + (hasTitle ? area.height * 0.2 : area.height * 0.085);
        const bottomReserve = chart.LegendPosition != null ? 0.2 : 0.12;

        return {
            left: area.left + leftPadding,
            top: area.top + topPadding,
            right: area.right - (area.right - area.left) * 0.07,
            bottom: area.bottom - area.height * bottomReserve
        };
    }

    private get area() { return this._doc; }

    // ========== 标题 ==========

    private DrawTitle(area: Rect, chart: ParsedChart, centerY: number): void {
        if (!chart.Title) return;
        const doc = this._doc;
        const font = ChartTextStyleResolver.ResolveTitleFont(chart);
        const text = chart.Title;

        doc.save();
        doc.fontSize(font.size);
        doc.font("Helvetica");
        doc.fillColor(chart.TitleColor.toPdfKitColor());

        const textWidth = doc.widthOfString(text);
        const centerX = (area.left + area.right) / 2;

        // 标题边框
        if (chart.HasTitleBorder) {
            const pad = 3.5;
            doc.lineWidth(Math.max(0.5, chart.TitleBorderWidthPt));
            doc.strokeColor(chart.TitleBorderColor.toPdfKitColor());
            doc.rect(centerX - textWidth / 2 - pad, centerY - font.size / 2 - pad, textWidth + pad * 2, font.size + pad * 2);
            doc.stroke();
        }

        doc.text(text, centerX - textWidth / 2, centerY - font.size / 2);
        doc.restore();
    }

    // ========== 值轴 ==========

    private ComputeValueAxis(chart: ParsedChart): { min: number; max: number; major: number } {
        let min = Infinity, max = -Infinity;

        for (const series of chart.Series) {
            for (const v of series.Values) {
                if (v.hasValue) {
                    min = Math.min(min, v.value);
                    max = Math.max(max, v.value);
                }
            }
        }

        if (!isFinite(min) || max <= min) { min = 0; max = 1; }

        // 显式轴设置优先
        if ((chart as any).ValueAxisMin != null && (chart as any).ValueAxisMin.hasValue) {
            min = (chart as any).ValueAxisMin.value;
        }
        if ((chart as any).ValueAxisMax != null && (chart as any).ValueAxisMax.hasValue) {
            max = (chart as any).ValueAxisMax.value;
        }

        let major = chart.MajorUnit ?? ChartValueAxisIntervals.ComputeMajorUnit(min, max);
        if (major <= 0) major = 1;

        return { min, max, major };
    }

    private DrawValueAxis(
        area: Rect,
        plot: Rect,
        chart: ParsedChart,
        min: number,
        max: number,
        major: number
    ): void {
        if (max <= min) return;
        const doc = this._doc;

        doc.save();
        doc.lineWidth(Math.max(0.5, chart.ValueGridlineWidthPt));
        doc.strokeColor(chart.ValueGridlineColor.toPdfKitColor());

        // 网格线
        if (chart.HasValueGridlines) {
            const steps = Math.round((max - min) / major);
            for (let i = 0; i <= steps; i++) {
                const value = min + i * major;
                const y = this.ValueToY(value, min, max, plot);
                doc.moveTo(plot.left, y);
                doc.lineTo(plot.right, y);
                doc.stroke();
            }
        }

        // 值轴线
        ChartValueAxisLineRenderer.Draw(doc, plot, chart, min, max, major);

        // 值轴标签
        doc.font("Helvetica");
        doc.fontSize(this.labelSizePt);
        doc.fillColor(chart.AxisLabelColor.toPdfKitColor());

        const steps2 = Math.round((max - min) / major);
        for (let i = 0; i <= steps2; i++) {
            const value = min + i * major;
            const y = this.ValueToY(value, min, max, plot);
            const label = this.FormatValue(value / chart.ValueDisplayUnitDivisor, chart.ValueFormatCode);
            const tw = doc.widthOfString(label);
            doc.text(label, plot.left - tw - 4, y - this.labelSizePt / 2);
        }

        doc.restore();
    }

    // ========== 分类轴 ==========

    private DrawCategoryAxis(
        area: Rect,
        plot: Rect,
        chart: ParsedChart
    ): void {
        const doc = this._doc;

        if (chart.Kind === ChartKind.Line && !chart.IsLine3D) {
            LineChartCategoryAxisRenderer.Draw(doc, plot, chart, this._dateSystem);
        } else {
            doc.save();
            doc.lineWidth(1);
            doc.strokeColor("#D9D9D9");
            doc.moveTo(plot.left, plot.bottom);
            doc.lineTo(plot.right, plot.bottom);
            doc.stroke();
            doc.restore();
        }

        if (this._suppressCategoryAxisLabels) return;

        const count = Math.max(chart.Categories.length, chart.CategoryValues.length);
        if (count === 0) return;

        doc.save();
        doc.font("Helvetica");
        doc.fontSize(this.labelSizePt);
        doc.fillColor(chart.AxisLabelColor.toPdfKitColor());

        const rotationDeg = Math.abs(chart.CategoryAxisTextRotationDeg) > 0.1 ? chart.CategoryAxisTextRotationDeg : 0;

        for (let i = 0; i < count; i++) {
            const x = this.CategoryX(i, count, plot, chart.CrossBetween);
            const label = chart.Categories[i] ?? "";

            if (Math.abs(rotationDeg) > 0.1) {
                doc.save();
                doc.translate(x, plot.bottom + 6);
                doc.rotate(rotationDeg);
                doc.text(label, -doc.widthOfString(label), 0);
                doc.restore();
            } else {
                const tw = doc.widthOfString(label);
                doc.text(label, x - tw / 2, plot.bottom + 4);
            }
        }

        doc.restore();
    }

    // ========== 序列（折线/面积） ==========

    private DrawSeries(
        plot: Rect,
        chart: ParsedChart,
        min: number,
        max: number
    ): void {
        const doc = this._doc;
        const pointCount = this.MaxPointCount(chart);
        const baselineY = this.ValueToY(Math.max(min, 0), min, max, plot);

        for (let s = 0; s < chart.Series.length; s++) {
            const series = chart.Series[s];
            const displayValues = ChartStackingMath.BuildDisplayValues(chart, s);
            const renderPoints: Array<{ x: number; y: number }> = [];

            // 日期轴点
            let datePoints = ChartSeriesPointBuilder.BuildDateSeriesPoints(
                chart, displayValues, plot, min, max, this._dateSystem
            );
            if (!datePoints) {
                datePoints = [];
                for (let i = 0; i < displayValues.length; i++) {
                    if (!displayValues[i].hasValue) continue;
                    const x = this.PointX(chart, i, pointCount, plot);
                    const y = this.ValueToY(displayValues[i].value, min, max, plot);
                    datePoints.push({ x, y });
                }
            }
            renderPoints.push(...datePoints);

            // 面积图填充
            if (chart.Kind === ChartKind.Area) {
                doc.save();
                doc.fillColor(series.Color.toPdfKitColor());
                if (datePoints.length > 0) {
                    doc.moveTo(datePoints[0].x, baselineY);
                    for (const p of datePoints) doc.lineTo(p.x, p.y);
                    for (let i = datePoints.length - 1; i >= 0; i--) {
                        doc.lineTo(datePoints[i].x, baselineY);
                    }
                    doc.closePath();
                    doc.fill();
                }
                doc.restore();
            }

            // 线条
            if (series.HasVisibleLine && datePoints.length > 1) {
                ChartSeriesStrokeRenderer.Draw(doc, datePoints, series);
            }

            // 标记
            this.DrawSeriesMarkers(plot, series, renderPoints);
        }
    }

    // ========== 柱体 ==========

    private DrawColumns(
        plot: Rect,
        chart: ParsedChart,
        min: number,
        max: number
    ): void {
        if (chart.Series.length === 0) return;
        const doc = this._doc;
        const pointCount = this.MaxPointCount(chart);
        if (pointCount === 0) return;

        const slot = (plot.right - plot.left) / pointCount;
        const { groupWidth, barWidth, seriesPitch } = this.ResolveColumnGeometry(chart);
        const baselineY = this.ValueToY(Math.max(min, 0), min, max, plot);

        for (let i = 0; i < pointCount; i++) {
            const slotLeft = plot.left + i * slot + (slot - groupWidth) / 2;
            for (let s = 0; s < chart.Series.length; s++) {
                const series = chart.Series[s];
                if (i >= series.Values.length || !series.Values[i].hasValue) continue;

                const y = this.ValueToY(series.Values[i].value, min, max, plot);
                const left = slotLeft + s * seriesPitch;
                const top = Math.min(y, baselineY);
                const bottom = Math.max(y, baselineY);

                doc.save();
                const pointColor = ChartExplicitFormatting.PointColor(series, i);
                doc.fillColor(ToPdfKitColor(pointColor));
                doc.rect(left, top, barWidth, bottom - top);
                doc.fill();
                doc.restore();
            }
        }
    }

    // ========== 条形 ==========

    private DrawBars(
        plot: Rect,
        chart: ParsedChart,
        min: number,
        max: number
    ): void {
        if (chart.Series.length === 0) return;
        const doc = this._doc;
        const pointCount = this.MaxPointCount(chart);
        if (pointCount === 0) return;

        const slot = (plot.bottom - plot.top) / pointCount;
        const { groupWidth, barWidth, seriesPitch } = this.ResolveColumnGeometry(chart);
        const baselineX = this.ValueToX(Math.max(min, 0), min, max, plot);

        for (let i = 0; i < pointCount; i++) {
            const slotTop = plot.top + i * slot + (slot - groupWidth) / 2;
            for (let s = 0; s < chart.Series.length; s++) {
                const series = chart.Series[s];
                if (i >= series.Values.length || !series.Values[i].hasValue) continue;

                const x = this.ValueToX(series.Values[i].value, min, max, plot);
                const top = slotTop + s * seriesPitch;
                const left = Math.min(x, baselineX);
                const right = Math.max(x, baselineX);

                doc.save();
                doc.fillColor(series.Color.toPdfKitColor());
                doc.rect(left, top, right - left, barWidth);
                doc.fill();
                doc.restore();
            }
        }
    }

    // ========== 饼图 ==========

    private DrawPie(
        area: Rect,
        chart: ParsedChart
    ): void {
        const doc = this._doc;

        // Doughnut
        if (chart.IsDoughnut) {
            const doughnutRect = DoughnutChartRenderer.Draw(doc, area, chart);
            this.DrawLegend(area, doughnutRect, chart);
            return;
        }

        if (chart.IsPie3D) {
            const values = chart.Series[0]?.Values ?? [];
            let sum = 0;
            for (const v of values) if (v.hasValue) sum += Math.abs(v.value);
            if (sum > 0) {
                const colors: Array<{ red: number; green: number; blue: number; alpha: number }> = [];
                for (let i = 0; i < values.length; i++) {
                    colors.push(this.PiePointColor(chart, i));
                }
                Pie3DChartRenderer.Draw(doc, area, chart, values.map(v => ({ hasValue: v.hasValue, value: v.value })), colors, sum);
            }
            return;
        }

        // 标准饼图
        const size = Math.min(area.right - area.left, area.bottom - area.top) * 0.635;
        const centerX = (area.left + area.right) / 2;
        const centerY = area.top + (area.bottom - area.top) * 0.19 + size / 2;
        const rect = {
            left: centerX - size / 2, top: centerY - size / 2,
            right: centerX + size / 2, bottom: centerY + size / 2
        };

        if (chart.Series.length === 0) return;
        const values = chart.Series[0].Values;
        let sum = 0;
        for (const v of values) if (v.hasValue) sum += Math.abs(v.value);
        if (sum <= 0) return;

        let startAngle = -90 + (chart.PieRotationDeg ?? 0);
        for (let i = 0; i < values.length; i++) {
            if (!values[i].hasValue) continue;
            const sweep = (Math.abs(values[i].value) / sum) * 360;
            const color = this.PiePointColor(chart, i);
            const h = (v: number) => Math.round(v).toString(16).padStart(2, "0");
            const hex = `#${h(color.red)}${h(color.green)}${h(color.blue)}`;

            doc.save();
            doc.fillColor(hex);
            doc.moveTo(centerX, centerY);
            const startRad = startAngle * Math.PI / 180;
            const endRad = (startAngle + sweep) * Math.PI / 180;
            const steps = Math.max(16, Math.ceil(sweep / 3));
            for (let j = 0; j <= steps; j++) {
                const angle = startRad + (endRad - startRad) * j / steps;
                doc.lineTo(centerX + size / 2 * Math.cos(angle), centerY + size / 2 * Math.sin(angle));
            }
            doc.closePath();
            doc.fill();
            doc.restore();

            startAngle += sweep;
        }

        this.DrawLegend(area, rect, chart);
    }

    private PiePointColor(chart: ParsedChart, index: number): { red: number; green: number; blue: number; alpha: number } {
        if (chart.Series[0]?.PointColors && index < chart.Series[0].PointColors.length
            && chart.Series[0].PointColors[index]) {
            return chart.Series[0].PointColors[index]!;
        }
        const palette = [
            { red: 0x44, green: 0x72, blue: 0xC4, alpha: 255 },
            { red: 0xED, green: 0x7D, blue: 0x31, alpha: 255 },
            { red: 0xA5, green: 0xA5, blue: 0xA5, alpha: 255 },
            { red: 0xFF, green: 0xC0, blue: 0x00, alpha: 255 },
            { red: 0x5B, green: 0x9B, blue: 0xD5, alpha: 255 },
            { red: 0x70, green: 0xAD, blue: 0x47, alpha: 255 }
        ];
        return palette[index % palette.length];
    }

    // ========== 散点图 ==========

    private DrawScatter(
        area: Rect,
        plot: Rect,
        chart: ParsedChart
    ): void {
        const { min, max, major } = this.ComputeValueAxis(chart);

        // 值轴
        this.DrawValueAxis(area, plot, chart, min, max, major);

        // X 轴
        const xRange = ScatterChartAxisMath.ResolveXRange(chart);
        if (xRange.major > 0) {
            const doc = this._doc;
            doc.save();
            doc.font("Helvetica");
            doc.fontSize(this.labelSizePt);
            doc.fillColor(chart.AxisLabelColor.toPdfKitColor());

            const steps = Math.round((xRange.max - xRange.min) / xRange.major);
            for (let i = 0; i <= steps; i++) {
                const value = xRange.min + i * xRange.major;
                const x = ScatterChartAxisMath.ToScreen(value, xRange.min, xRange.max, plot.left, plot.right);
                doc.text(this.FormatValue(value, chart.ScatterXFormatCode), x - 12, plot.bottom + 4);
            }
            doc.restore();
        }

        // 序列
        for (let s = 0; s < chart.Series.length; s++) {
            const series = chart.Series[s];
            const count = Math.min(series.XValues.length, series.Values.length);
            const points: Array<{ x: number; y: number }> = [];

            for (let i = 0; i < count; i++) {
                if (!series.XValues[i].hasValue || !series.Values[i].hasValue) continue;
                points.push({
                    x: ScatterChartAxisMath.ToScreen(series.XValues[i].value, xRange.min, xRange.max, plot.left, plot.right),
                    y: this.ValueToY(series.Values[i].value, min, max, plot)
                });
            }

            // 连线
            if (series.HasVisibleLine && points.length > 1) {
                const doc = this._doc;
                doc.save();
                doc.strokeColor(series.LineColor.toPdfKitColor());
                doc.lineWidth(Math.max(0.5, series.LineWidthPt));
                doc.lineCap("round");
                doc.lineJoin("round");
                doc.moveTo(points[0].x, points[0].y);
                if (series.SmoothLine) {
                    ScatterSmoothPathBuilder.Append(doc, points);
                } else {
                    for (let i = 1; i < points.length; i++) doc.lineTo(points[i].x, points[i].y);
                }
                doc.stroke();
                doc.restore();
            }

            // 气泡
            if (chart.IsBubble) {
                BubbleChartMath.DrawSeries(this._doc, plot, chart, min, max);
            } else {
                this.DrawSeriesMarkers(plot, series, points);
            }
        }

        this.DrawPlotAreaBorder(area, plot, chart);
        this.DrawAxisTitles(area, plot, chart);
        this.DrawLegend(area, plot, chart);
    }

    // ========== 雷达图 ==========

    private DrawRadar(
        area: Rect,
        plot: Rect,
        chart: ParsedChart
    ): void {
        const { min, max, major } = this.ComputeValueAxis(chart);
        if (chart.Series.length === 0 || max <= min) return;

        const centerX = (plot.left + plot.right) / 2;
        const centerY = (plot.top + plot.bottom) / 2 + 1;
        const radius = Math.min(plot.right - plot.left, plot.bottom - plot.top) * 0.462;
        const categories = chart.Categories.length;
        if (categories <= 2) return;

        const doc = this._doc;

        // 网格
        doc.save();
        doc.strokeColor("#D9D9D9");
        doc.lineWidth(1);

        for (let c = 0; c < categories; c++) {
            const angle = -Math.PI / 2 + Math.PI * 2 * c / categories;
            doc.moveTo(centerX, centerY);
            doc.lineTo(centerX + radius * Math.cos(angle), centerY + radius * Math.sin(angle));
            doc.stroke();
        }

        const steps = major > 0 ? Math.round((max - min) / major) : 5;
        const ringMajor = major > 0 ? major : (max - min) / 5;
        const rings = Math.max(1, Math.round((max - min) / ringMajor));
        for (let r = 1; r <= rings; r++) {
            const rr = radius * r / rings;
            doc.moveTo(centerX + rr * Math.cos(-Math.PI / 2), centerY + rr * Math.sin(-Math.PI / 2));
            for (let c = 1; c <= categories; c++) {
                const angle = -Math.PI / 2 + Math.PI * 2 * c / categories;
                doc.lineTo(centerX + rr * Math.cos(angle), centerY + rr * Math.sin(angle));
            }
            doc.closePath();
            doc.stroke();
        }
        doc.restore();

        // 序列
        for (const series of chart.Series) {
            const pts: Array<{ x: number; y: number }> = [];
            for (let c = 0; c < categories; c++) {
                if (c >= series.Values.length || !series.Values[c].hasValue) continue;
                const angle = -Math.PI / 2 + Math.PI * 2 * c / categories;
                const r = Math.max(0, Math.min(radius,
                    (series.Values[c].value - min) / (max - min) * radius));
                pts.push({
                    x: centerX + r * Math.cos(angle),
                    y: centerY + r * Math.sin(angle)
                });
            }

            if (pts.length < 3) continue;

            doc.save();
            if (chart.RadarStyle === "filled") {
                doc.fillColor(series.Color.toPdfKitColor());
                doc.globalAlpha(0.43); // 110/255
                doc.moveTo(pts[0].x, pts[0].y);
                for (let i = 1; i < pts.length; i++) doc.lineTo(pts[i].x, pts[i].y);
                doc.closePath();
                doc.fill();
                doc.globalAlpha(1);
            }
            if (series.HasVisibleLine) {
                doc.strokeColor(series.LineColor.toPdfKitColor());
                doc.lineWidth(Math.max(0.5, series.LineWidthPt));
                doc.lineCap("round");
                doc.lineJoin("round");
                doc.moveTo(pts[0].x, pts[0].y);
                for (let i = 1; i < pts.length; i++) doc.lineTo(pts[i].x, pts[i].y);
                doc.closePath();
                doc.stroke();
            }
            doc.restore();
        }

        this.DrawLegend(area, plot, chart);
    }

    // ========== 图例 ==========

    private DrawLegend(
        area: Rect,
        plot: Rect,
        chart: ParsedChart
    ): void {
        if (!chart.LegendPosition) return;
        const doc = this._doc;
        const itemCount = this.LegendItemCount(chart);
        if (itemCount === 0) return;

        const swatchWidth = this.ResolveLegendSwatchWidth(chart);
        const gap = 4;
        const itemGap = 8;
        const lineHeight = this.labelSizePt * 1.7;
        const textSize = this.labelSizePt;

        // 估算总宽度
        let totalWidth = 0;
        const widths: number[] = [];
        for (let i = 0; i < itemCount; i++) {
            const label = this.LegendItemLabel(chart, i);
            const tw = doc.widthOfString(label || "Series");
            widths.push(tw);
            totalWidth += swatchWidth + gap + tw + (i > 0 ? itemGap : 0);
        }

        // 位置
        let legendLeft: number, legendTop: number;
        const pos = chart.LegendPosition;
        if (pos === "b" || pos === null) {
            legendLeft = (area.left + area.right - totalWidth) / 2;
            legendTop = Math.max(plot.bottom + 12, area.bottom - area.height * 0.08);
        } else if (pos === "t") {
            legendLeft = (area.left + area.right - totalWidth) / 2;
            legendTop = area.top + area.height * 0.02;
        } else if (pos === "l") {
            legendLeft = area.left + this.labelSizePt;
            legendTop = (area.top + area.bottom - itemCount * lineHeight) / 2;
        } else {
            legendLeft = area.right - totalWidth - this.labelSizePt;
            legendTop = (area.top + area.bottom - itemCount * lineHeight) / 2;
        }

        doc.save();
        doc.font("Helvetica");
        doc.fontSize(textSize);

        let x = legendLeft;
        let y = legendTop;

        for (let i = 0; i < itemCount; i++) {
            const label = this.LegendItemLabel(chart, i);
            const color = this.LegendItemColor(chart, i);
            const h = (v: number) => Math.round(v).toString(16).padStart(2, "0");
            const hex = `#${h(color.red)}${h(color.green)}${h(color.blue)}`;

            // swatch
            if (chart.Kind === ChartKind.Line || chart.Kind === ChartKind.Scatter || chart.Kind === ChartKind.Radar) {
                doc.strokeColor(hex);
                doc.lineWidth(2.5);
                doc.moveTo(x, y + textSize / 2);
                doc.lineTo(x + swatchWidth, y + textSize / 2);
                doc.stroke();
            } else {
                doc.fillColor(hex);
                doc.rect(x, y + 2, swatchWidth, Math.min(textSize - 4, 7));
                doc.fill();
            }

            // label
            doc.fillColor(chart.LegendTextColor.toPdfKitColor());
            doc.text(label, x + swatchWidth + gap, y);

            x += swatchWidth + gap + widths[i] + itemGap;
            if (pos === "l" || pos === "r") {
                x = legendLeft;
                y += lineHeight;
            }
        }

        doc.restore();
    }

    private LegendItemCount(chart: ParsedChart): number {
        if (chart == null) return 0;
        if (chart.Kind === ChartKind.Pie) return chart.Categories.length;
        return chart.Series.length;
    }

    private LegendItemLabel(chart: ParsedChart, legendIndex: number): string {
        if (!chart) return "";
        if (chart.Kind === ChartKind.Pie) return chart.Categories[legendIndex] ?? "";
        const seriesIndex = this.ResolveLegendSeriesIndex(chart, legendIndex);
        return chart.Series[seriesIndex]?.Name ?? "";
    }

    private LegendItemColor(chart: ParsedChart, legendIndex: number): { red: number; green: number; blue: number; alpha: number } {
        if (!chart) return { red: 128, green: 128, blue: 128, alpha: 255 };
        if (chart.Kind === ChartKind.Pie) return this.PiePointColor(chart, legendIndex);
        const seriesIndex = this.ResolveLegendSeriesIndex(chart, legendIndex);
        return chart.Series[seriesIndex]?.Color ?? { red: 128, green: 128, blue: 128, alpha: 255 };
    }

    private ResolveLegendSeriesIndex(chart: ParsedChart, legendIndex: number): number {
        if (!chart.Series || legendIndex >= chart.Series.length) return legendIndex;
        return legendIndex;
    }

    private ResolveLegendSwatchWidth(chart: ParsedChart): number {
        if (chart.IsDoughnut) return 7;
        if (chart.Kind === ChartKind.Line || chart.Kind === ChartKind.Scatter) return 19.2;
        return 9;
    }

    // ========== 辅助绘制 ==========

    private DrawSeriesMarkers(
        plot: Rect,
        series: ChartSeries,
        points: Array<{ x: number; y: number }>
    ): void {
        if (!series.MarkerVisible || points.length === 0) return;
        const doc = this._doc;
        const diameter = Math.max(3, series.MarkerSize);
        const radius = diameter / 2;

        doc.save();
        doc.fillColor(series.MarkerFillColor.toPdfKitColor());

        for (const p of points) {
            doc.circle(p.x, p.y, radius);
            doc.fill();
        }
        doc.restore();
    }

    private DrawAxisTitles(
        area: Rect,
        plot: Rect,
        chart: ParsedChart
    ): void {
        const doc = this._doc;
        doc.save();
        doc.font("Helvetica");
        doc.fontSize(10);
        doc.fillColor(chart.AxisLabelColor.toPdfKitColor());

        if (chart.CategoryAxisTitle) {
            const tw = doc.widthOfString(chart.CategoryAxisTitle);
            doc.text(chart.CategoryAxisTitle, (plot.left + plot.right - tw) / 2, plot.bottom + 25);
        }

        if (chart.ValueAxisTitle) {
            const tw = doc.widthOfString(chart.ValueAxisTitle);
            doc.save();
            doc.translate(plot.left - 40, (plot.top + plot.bottom) / 2 + tw / 2);
            doc.rotate(-90);
            doc.text(chart.ValueAxisTitle, 0, 0);
            doc.restore();
        }

        doc.restore();
    }

    private DrawPlotAreaBorder(
        area: Rect,
        plot: Rect,
        chart: ParsedChart
    ): void {
        const doc = this._doc;
        doc.save();
        doc.lineWidth(Math.max(0.5, chart.PlotAreaBorderWidthPt || 0.5));
        doc.strokeColor(chart.PlotAreaBorderColor?.toPdfKitColor() ?? "#D9D9D9");
        doc.rect(plot.left, plot.top, plot.right - plot.left, plot.bottom - plot.top);
        doc.stroke();
        doc.restore();
    }

    private DrawDataTable(
        area: Rect,
        plot: Rect,
        chart: ParsedChart
    ): void {
        if (!chart.HasDataTable || chart.Series.length === 0) return;
        // 简化：跳过详细的数据表绘制
    }

    // ========== 几何辅助 ==========

    private ResolveColumnGeometry(chart: ParsedChart): { groupWidth: number; barWidth: number; seriesPitch: number } {
        const seriesCount = chart?.Series?.length ?? 0;
        if (seriesCount <= 0) return { groupWidth: 0, barWidth: 0, seriesPitch: 0 };

        const gapRatio = (chart.GapWidthPercent ?? 150) / 100;
        const overlapRatio = (chart.OverlapPercent ?? 0) / 100;
        const pitchFactor = Math.max(0, 1 - overlapRatio);
        const spanFactor = 1 + (seriesCount - 1) * pitchFactor;
        const denom = spanFactor + gapRatio;

        if (denom <= 0) {
            const fallback = 0.8 / seriesCount;
            return { groupWidth: 0.8, barWidth: fallback, seriesPitch: fallback };
        }

        const barWidthRatio = 1 / denom;
        const seriesPitchRatio = barWidthRatio * pitchFactor;
        const groupWidthRatio = barWidthRatio * spanFactor;
        return { groupWidth: groupWidthRatio, barWidth: barWidthRatio, seriesPitch: seriesPitchRatio };
    }

    private MaxPointCount(chart: ParsedChart): number {
        let count = Math.max(chart.Categories.length, chart.CategoryValues.length);
        for (const s of chart.Series) count = Math.max(count, s.Values.length);
        return count;
    }

    private ValueToY(value: number, min: number, max: number, plot: { top: number; bottom: number }): number {
        if (max <= min) return plot.bottom;
        const ratio = (value - min) / (max - min);
        return plot.bottom - ratio * (plot.bottom - plot.top);
    }

    private ValueToX(value: number, min: number, max: number, plot: { left: number; right: number }): number {
        if (max <= min) return plot.left;
        const ratio = (value - min) / (max - min);
        return plot.left + ratio * (plot.right - plot.left);
    }

    private PointX(chart: ParsedChart, i: number, count: number, plot: { left: number; right: number }): number {
        const w = plot.right - plot.left;
        const crossBetween = chart.CrossBetween;
        const slot = count > 0 ? w / count : 0;
        return plot.left + i * slot + (crossBetween ? slot / 2 : 0);
    }

    private CategoryX(i: number, count: number, plot: { left: number; right: number }, crossBetween: boolean): number {
        const w = plot.right - plot.left;
        const slot = count > 0 ? w / count : 0;
        return plot.left + i * slot + (crossBetween ? slot / 2 : slot / 2);
    }

    // ========== 文字测量 ==========

    private MeasureText(_text: string, _size: number): number {
        return _text.length * _size * 0.6;
    }

    private MeasureMaxValueAxisLabelWidth(chart: ParsedChart): number {
        let maxWidth = 0;
        const { min, max, major } = this.ComputeValueAxis(chart);
        const steps = major > 0 ? Math.round((max - min) / major) : 5;
        for (let i = 0; i <= steps; i++) {
            const value = min + i * major;
            const label = this.FormatValue(value, chart.ValueFormatCode);
            maxWidth = Math.max(maxWidth, this.MeasureText(label, this.labelSizePt));
        }
        return maxWidth;
    }

    private FormatValue(value: number, formatCode: string): string {
        if (!formatCode) return String(value);

        // 简化的数值格式化
        if (/0\.0+/.test(formatCode)) {
            const match = formatCode.match(/0\.0+/);
            if (match) {
                const decimals = match[0].length - 2;
                return value.toFixed(decimals);
            }
        }

        if (formatCode === "0%" || /0\.0+%/.test(formatCode)) {
            const decimals = /0\.0+/.test(formatCode) ? (formatCode.match(/0\.0+/)![0].length - 2) : 0;
            return (value * 100).toFixed(decimals) + "%";
        }

        if (formatCode.includes("$")) {
            return "$" + value.toFixed(2);
        }

        return String(value);
    }
}
