/**
 * 图表种类枚举。
 */
export enum ChartKind {
    Column = "column",
    Bar = "bar",
    Line = "line",
    Area = "area",
    Pie = "pie",
    Scatter = "scatter",
    Radar = "radar",
    Doughnut = "doughnut",
    Surface = "surface",
    Bubble = "bubble"
}

/**
 * 矩形类型别名（PDF points）。
 * 兼容内联对象字面量和额外的 width/height 计算属性访问。
 */
export type Rect = {
    left: number;
    top: number;
    right: number;
    bottom: number;
    [key: string]: any;
};

/**
 * 颜色值结构。
 */
export interface RgbaColor {
    red: number;
    green: number;
    blue: number;
    alpha: number;
    toPdfKitColor(): string;
}

/**
 * 将 RgbaColor 转为 PDFKit 颜色字符串的辅助函数。
 */
export function ToPdfKitColor(c: { red: number; green: number; blue: number; alpha?: number }): string {
    const h = (v: number) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0");
    if (c.alpha != null && c.alpha < 255) {
        return `#${h(c.red)}${h(c.green)}${h(c.blue)}${h(c.alpha)}`;
    }
    return `#${h(c.red)}${h(c.green)}${h(c.blue)}`;
}

/**
 * 图表数据值类型（可空 double）。
 */
export interface ChartValue {
    hasValue: boolean;
    value: number;
}

/**
 * 图表序列。
 */
export class ChartSeries {
    Name: string = "";
    Values: ChartValue[] = [];
    XValues: ChartValue[] = [];
    BubbleSizes: ChartValue[] | null = null;

    // 颜色和填充
    Color: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    PointColors: (RgbaColor | null)[] | null = null;

    // 线条
    HasVisibleLine: boolean = true;
    LineColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    LineWidthPt: number = 1;
    LineDashStyle: string = "solid";
    SmoothLine: boolean = false;

    // 标记
    MarkerVisible: boolean = false;
    MarkerSize: number = 7;
    MarkerSymbol: string = "circle";
    MarkerFillColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    MarkerStrokeColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    MarkerStrokeWidthPt: number = 0.75;

    // 数据标签
    ShowValueDataLabels: boolean = false;
    DataLabelPosition: string = "ctr"; // ctr, inEnd, outEnd
    DataLabelColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    DataLabelSizePt: number | null = null;
    ValueFormatCode: string = "";

    // 3D related
    HasPatternFill: boolean = false;
    PatternPreset: string = "";
    PatternForegroundColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 0, toPdfKitColor() { return ToPdfKitColor(this); } };
    PatternBackgroundColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 0, toPdfKitColor() { return ToPdfKitColor(this); } };
    Bubble3D: boolean = false;

    // 饼图
    ExplodePoints: number[] | null = null;
    ExplodeDistancePt: number = 0;

    // 误差线
    HasErrorBars: boolean = false;
    ErrorValueType: string = "stdErr";
    ErrorBarColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    ErrorBarWidthPt: number = 1;
    ErrorBarsHaveEndCaps: boolean = false;
}

/**
 * 图表字体。
 */
export interface ChartFont {
    Name: string;
    Size: number;
    Bold: boolean;
    Italic: boolean;
    Color: RgbaColor;
}

/**
 * 解析后的图表。
 * 包含从 Excel XML 中提取的所有图表属性和数据。
 */
export class ParsedChart {
    // ===== 基础属性 =====
    Kind: ChartKind = ChartKind.Column;
    Title: string | null = null;
    TitleAllCaps: boolean = false;
    HasManualTitle: boolean = false;
    TitleX: number = 0;
    TitleY: number = 0;
    TitleFont: ChartFont | null = null;
    TitleColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    TitleTextOutlineColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 0, toPdfKitColor() { return ToPdfKitColor(this); } };
    HasTitleBorder: boolean = false;
    TitleBorderColor: RgbaColor = { red: 0x88, green: 0x88, blue: 0x88, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    TitleBorderWidthPt: number = 1;

    Categories: string[] = [];
    CategoryValues: ChartValue[] = [];
    CategoryFormatCode: string = "";

    Series: ChartSeries[] = [];

    // ===== 绘图区域 =====
    PlotX: number = 0;
    PlotY: number = 0;
    PlotW: number = 0.8;
    PlotH: number = 0.6;
    HasManualPlot: boolean = false;

    // ===== 图例 =====
    LegendPosition: string | null = "b"; // b, t, l, r, tr, br
    HasManualLegend: boolean = false;
    LegendX: number = 0;
    LegendY: number = 0;
    LegendW: number = 1;
    LegendH: number = 0.2;
    LegendTextColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    LegendTextOutlineColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 0, toPdfKitColor() { return ToPdfKitColor(this); } };

    // ===== 堆叠 =====
    IsStacked: boolean = false;
    IsPercentStacked: boolean = false;
    CrossBetween: boolean = false;

    // ===== 尺寸和间距 =====
    GapWidthPercent: number = 150;
    OverlapPercent: number = 0;
    TextSizePt: number | null = null;

    // ===== 值轴 =====
    ValueFormatCode: string = "";
    ValueUnitDisplayLabel: string = "";
    ValueDisplayUnitDivisor: number = 1;
    ValueDisplayUnitLabel: string = "";
    ValueDisplayUnitLabelX: number = 0;
    ValueDisplayUnitLabelY: number = 0;
    HasManualValueDisplayUnitLabel: boolean = false;
    MajorUnit: number | null = null;
    MinorUnit: number | null = null;

    HasValueGridlines: boolean = true;
    HasMinorValueGridlines: boolean = false;
    ValueGridlineColor: RgbaColor = { red: 0xD9, green: 0xD9, blue: 0xD9, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    ValueGridlineWidthPt: number = 0.5;
    ValueGridlineDashStyle: string = "solid";
    MinorValueGridlineColor: RgbaColor = { red: 0xD9, green: 0xD9, blue: 0xD9, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    MinorValueGridlineWidthPt: number = 0.5;
    MinorValueGridlineDashStyle: string = "solid";

    HasValueAxisLine: boolean = true;
    ValueAxisLineColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    ValueAxisLineWidthPt: number = 1;
    ValueAxisMajorTickMark: string = "out"; // none, in, out, cross

    ValueAxisTitle: string = "";

    // ===== 分类轴 =====
    HasCategoryAxisLine: boolean = true;
    CategoryAxisLineColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    CategoryAxisLineWidthPt: number = 1;
    HasCategoryGridlines: boolean = false;
    CategoryGridlineColor: RgbaColor = { red: 0xD9, green: 0xD9, blue: 0xD9, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    CategoryGridlineWidthPt: number = 0.5;
    CategoryAxisTextRotationDeg: number = 0;
    CategoryAxisMajorTickMark: string = "out";
    CategoryAxisTitle: string = "";
    CategoryBaseTimeUnit: string = "days";
    HasDateCategoryAxis: boolean = false;
    DateAxisMin: ChartValue | null = null;
    DateAxisMax: ChartValue | null = null;
    DateAxisMajorUnit: number | null = null;

    // ===== 轴标签 =====
    AxisLabelColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    AxisTextOutlineColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 0, toPdfKitColor() { return ToPdfKitColor(this); } };
    HasVisibleSeriesAxis: boolean = false;
    SeriesAxisLabelColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };

    // ===== 图表区域 =====
    HasChartAreaFill: boolean = false;
    ChartAreaFillColor: RgbaColor = { red: 0xFF, green: 0xFF, blue: 0xFF, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    HasChartAreaBorder: boolean = false;
    ChartAreaBorderColor: RgbaColor = { red: 0x88, green: 0x88, blue: 0x88, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    ChartAreaBorderWidthPt: number = 1;
    HasChartAreaGradient: boolean = false;
    ChartAreaGradientColors: RgbaColor[] = [];
    ChartAreaGradientPositions: number[] = [];
    ChartAreaGradientAngleDeg: number = 0;
    HasChartAreaPattern: boolean = false;
    ChartAreaPatternPreset: string = "";
    ChartAreaPatternForegroundColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 0, toPdfKitColor() { return ToPdfKitColor(this); } };
    ChartAreaPatternBackgroundColor: RgbaColor = { red: 0xFF, green: 0xFF, blue: 0xFF, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    ChartAreaSoftEdgeRadiusPt: number = 0;
    HasChartAreaGlow: boolean = false;
    ChartAreaGlowColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 0, toPdfKitColor() { return ToPdfKitColor(this); } };
    ChartAreaGlowRadiusPt: number = 0;
    RoundedCorners: boolean = false;

    // ===== 绘图区域边框 =====
    HasPlotAreaBorder: boolean = false;
    PlotAreaBorderColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    PlotAreaBorderWidthPt: number = 1;

    // ===== 饼图 =====
    IsPie3D: boolean = false;
    PieRotationDeg: number | null = null;
    PiePointLineColors: RgbaColor[] = [];
    PiePointLineWidthsPt: number[] = [];
    IsDoughnut: boolean = false;
    DoughnutHoleSizeRatio: number = 0.5;
    IsOfPie: boolean = false;
    OfPieType: string = "pie"; // pie, bar
    HasOfPieSeparatorLines: boolean = false;
    OfPieSeparatorLineColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    OfPieSeparatorLineWidthPt: number = 1;
    OfPieSplitType: string = "auto"; // auto, pos, val, percent
    OfPieSplitPosition: { hasValue: boolean; value: number } | null = null;
    OfPieSecondPieSizePercent: number = 75;
    OfPieGapWidthPercent: number = 2;

    // ===== 3D =====
    IsArea3D: boolean = false;
    IsLine3D: boolean = false;
    IsBar3D: boolean = false;
    DepthPercent: number = 100;
    GapDepthPercent: number = 150;
    PerspectiveDeg: number = 30;
    RightAngleAxes: boolean = true;
    RotationXDeg: number = 15;
    RotationYDeg: number = 20;
    BarShape: string = "box"; // box, cylinder, cone, pyramid, tube
    SideWallImageData: Uint8Array | null = null;
    BackWallImageData: Uint8Array | null = null;

    // ===== 散点图 =====
    ScatterVaryColors: boolean = false;
    ScatterCategoryAxisDeleted: boolean = false;
    ScatterValueAxisDeleted: boolean = false;
    ScatterTitleOverlay: boolean = false;
    ScatterXFormatCode: string = "";
    ScatterXAxisMin: ChartValue | null = null;
    ScatterXAxisMax: ChartValue | null = null;
    ScatterXAxisMajorUnit: number | null = null;
    IsBubble: boolean = false;

    // ===== 雷达图 =====
    RadarStyle: string = "filled"; // filled, marker

    // ===== 数据表 =====
    HasDataTable: boolean = false;
    DataTableLineColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    DataTableLineWidthPt: number = 0.5;
    DataTableTextColor: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };
    DataTableShowKeys: boolean = false;
    DataTableShowHorizontalBorders: boolean = true;
    DataTableShowVerticalBorders: boolean = true;
    DataTableShowOutline: boolean = false;

    // ===== 序列轴 =====
    SeriesAxisLabelColor2D: RgbaColor = { red: 0, green: 0, blue: 0, alpha: 255, toPdfKitColor() { return ToPdfKitColor(this); } };

    // ===== 其他 =====
    Culture: string | null = null;
}

/**
 * 文字对齐枚举。
 */
export enum TextAlign {
    Left = "left",
    Center = "center",
    Right = "right"
}

/**
 * 垂直对齐枚举。
 */
export enum VAlign {
    Top = "top",
    Middle = "middle",
    Bottom = "bottom",
    Baseline = "baseline"
}
