import * as path from "path";
import { TestCase, TestRunner } from "../Shared/TestInfrastructure";
import { ChartAxisAndColumnTests } from "./ChartAxisAndColumnTests";
import { ChartDetailRenderingTests } from "./ChartDetailRenderingTests";
import { Line3DCategoryChartTests } from "./Line3DCategoryChartTests";
import { Line3DChartTextTests } from "./Line3DChartTextTests";
import { PrintAreaRenderingTests } from "./PrintAreaRenderingTests";
import { PrinterSettingsTests } from "./PrinterSettingsTests";
import {
  ASYNC_PDF_REASON,
  CHART_INTERNALS_REASON,
  guardedTest,
  inputRootDirectory,
  PDF_INSPECTION_REASON,
  RASTER_REASON,
  SHAPE_INTERNALS_REASON,
  TEXT_LAYOUT_REASON,
} from "./Support";

/**
 * Port of tests/Aspose.Cells_FOSS.PdfRenderingTests/Program.cs.
 *
 * Test names, ordering and suite name match the .NET suite. Fixtures are
 * resolved from the repository Input/ folder. Missing fixtures are reported as
 * SKIP. The upstream assertions rely on SkiaSharp rasterization, .NET-only
 * chart parsing and PDF inspection helpers; tests needing those facilities are
 * reported as SKIP with an explicit reason after the fixture check.
 */

function PrintAreaRangesAndScalingModes(): void {
  PrintAreaRenderingTests.RangesAndScalingModes();
}

function PrintAreaFixtureLayout(): void {
  PrintAreaRenderingTests.FixtureLayout();
}

function PrinterSettingsPublicHeaderAndRecovery(): void {
  PrinterSettingsTests.PublicHeaderAndRecovery();
}

function PrinterSettingsFixtureRoundTrip(): void {
  PrinterSettingsTests.FixtureRoundTrip();
}

function PrinterSettingsReferenceOutputGeometry(): void {
  PrinterSettingsTests.ReferenceOutputGeometry();
}

function ChartDetailStackedZeroAndTinySegments(): void {
  ChartDetailRenderingTests.StackedZeroAndTinySegments();
}

function ChartDetailMultilineCategoryLabels(): void {
  ChartDetailRenderingTests.MultilineCategoryLabels();
}

function ChartDetailAxisHeadroomAndGridlines(): void {
  ChartDetailRenderingTests.AxisHeadroomAndGridlines();
}

function ChartAxisLegacyValueAxes(): void {
  ChartAxisAndColumnTests.LegacyValueAxes();
}

function ChartAxisColumnBorders(): void {
  ChartAxisAndColumnTests.ColumnBorders();
}

function ChartAxisLeagueAxisGap(): void {
  ChartAxisAndColumnTests.LeagueAxisGap();
}

function PrintAreaChartFormatting(): void {
  PrintAreaRenderingTests.ChartFormatting();
}

function EachFixtureConvertsToValidPdf(): void {
  guardedTest(["sales1.xlsx"], PDF_INSPECTION_REASON);
}

function AllInputXlsxConvertToOutputPdfDirectory(): void {
  guardedTest([], ASYNC_PDF_REASON);
}

function NumberformatExcelFixtureDisplaysLocalizedValues(): void {
  guardedTest(["display/numberformat.xlsx"], RASTER_REASON);
}

function NumberformatExcelFixtureUsesSelectedSectionFontColors(): void {
  guardedTest(["display/numberformat.xlsx"], RASTER_REASON);
}

function Mapchart1ParsesVisibleRegionMap(): void {
  guardedTest(["mapchart/mapchart1.xlsx"], CHART_INTERNALS_REASON);
}

function Mapchart2PreservesGroupedGuidanceAndDivergingScale(): void {
  guardedTest(["mapchart/mapchart2.xlsx"], CHART_INTERNALS_REASON);
}

function LinechartCategoryLabelsFollowChartLanguage(): void {
  guardedTest(["linechart1.xlsx"], CHART_INTERNALS_REASON);
}

function LinechartManualPlotKeepsAutomaticDecorationsInChartBands(): void {
  guardedTest(["linechart/linechart1.xlsx"], CHART_INTERNALS_REASON);
}

function LinechartDateAxisPreservesIrregularCategoryValues(): void {
  guardedTest(["linechart/linechart1-1.xlsx"], CHART_INTERNALS_REASON);
}

function LinechartDateAxisTitleAndRotationFollowExcelReference(): void {
  guardedTest(["linechart/linechart1-1.xlsx"], CHART_INTERNALS_REASON);
}

function LinechartOutlierDateAxisSuppressesRotatedLabelsLikeExcelReference(): void {
  guardedTest(["linechart/linechart1.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart12RendersHorizontalCategoryAxisLabels(): void {
  guardedTest(["linechart/linechart1-2.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart12CategoryLayoutMatchesExcelPlotAndLegend(): void {
  guardedTest(["linechart/linechart1-2.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart52StackedCategoryLayoutMatchesExcel(): void {
  guardedTest(["linechart/linechart5-2.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart11AutomaticLayoutMatchesExcelBands(): void {
  guardedTest(["linechart/linechart1-1.xlsx"], CHART_INTERNALS_REASON);
}

function LinechartMonthAxisPointsSnapToMonthSlots(): void {
  guardedTest(["linechart/linechart1.xlsx"], CHART_INTERNALS_REASON);
}

function LinechartDateAxisCrossBetweenInsetsFirstPointFromValueAxis(): void {
  guardedTest(["linechart/linechart1.xlsx"], CHART_INTERNALS_REASON);
}

function Columnchart1ParsesGapWidthFromChartXml(): void {
  guardedTest(["columnchart/columnchart1.xlsx"], CHART_INTERNALS_REASON);
}

function ColumnchartClusteredGeometryUsesGapWidthToShrinkBars(): void {
  guardedTest(["columnchart/columnchart1.xlsx"], CHART_INTERNALS_REASON);
}

function Columnchart11StackedColumnsShareSingleSlotAndStackVertically(): void {
  guardedTest(["columnchart/columnchart1-1.xlsx"], CHART_INTERNALS_REASON);
}

function Columnchart2RendersAxisTitlesAndManualLegendLayout(): void {
  guardedTest(["columnchart/columnchart2.xlsx"], CHART_INTERNALS_REASON);
}

function Columnchart21RendersCenteredValueDataLabels(): void {
  guardedTest(["columnchart/columnchart2-1.xlsx"], CHART_INTERNALS_REASON);
}

function Columnchart3RendersDataTableAndStandardErrorBars(): void {
  guardedTest(["columnchart/columnchart3.xlsx"], CHART_INTERNALS_REASON);
}

function Columnchart31UsesTallCrossPagePercentPlot(): void {
  guardedTest(["columnchart/columnchart3-1.xlsx"], CHART_INTERNALS_REASON);
}

function Columnchart4RendersTrueThreeDColumns(): void {
  guardedTest(["columnchart/columnchart4.xlsx"], CHART_INTERNALS_REASON);
}

function Columnchart41RendersLabelsDenseGridAndAxisLine(): void {
  guardedTest(["columnchart/columnchart4-1.xlsx"], CHART_INTERNALS_REASON);
}

function Columnchart5RendersStackedThreeDColumns(): void {
  guardedTest(["columnchart/columnchart5.xlsx"], CHART_INTERNALS_REASON);
}

function Columnchart6RendersPercentStackedThreeDCylinders(): void {
  guardedTest(["columnchart/columnchart6.xlsx"], CHART_INTERNALS_REASON);
}

function Columnchart61UsesBoxTitleColorAndDensePercentTicks(): void {
  guardedTest(["columnchart/columnchart6-1.xlsx"], CHART_INTERNALS_REASON);
}

function Columnchart7RendersTexturedWallsAndSeriesDepth(): void {
  guardedTest(["columnchart/columnchart7.xlsx"], CHART_INTERNALS_REASON);
}

function ColumnchartStandardThreeDProjectionRespectsViewAndGaps(): void {
  guardedTest([], CHART_INTERNALS_REASON);
}

function Columnchart71PreservesMonthGapsAndHighPerspective(): void {
  guardedTest(["columnchart/columnchart7-1.xlsx"], CHART_INTERNALS_REASON);
}

function ColumnchartThreeDMonthSlotsSupportBothDateSystems(): void {
  guardedTest([], CHART_INTERNALS_REASON);
}

function Barchart4RendersTrueThreeDHorizontalBars(): void {
  guardedTest(["BarChart/barchart4.xlsx"], CHART_INTERNALS_REASON);
}

function Barchart41PreservesDenseGridAndValueAxisLine(): void {
  guardedTest(["BarChart/barchart4-1.xlsx"], CHART_INTERNALS_REASON);
}

function Barchart5PreservesStackedThreeDBarGrouping(): void {
  guardedTest(["BarChart/barchart5.xlsx"], CHART_INTERNALS_REASON);
}

function Barchart51PreservesSaturatedBlueChartGradient(): void {
  guardedTest(["BarChart/barchart5-1.xlsx"], CHART_INTERNALS_REASON);
}

function Barchart6NormalizesPercentStackedThreeDBars(): void {
  guardedTest(["BarChart/barchart6.xlsx"], CHART_INTERNALS_REASON);
}

function Barchart61PreservesStandardPercentStackedPerspective(): void {
  guardedTest(["BarChart/barchart6-1.xlsx"], CHART_INTERNALS_REASON);
}

function Barchart7RendersTexturedDateAxisThreeDBars(): void {
  guardedTest(["BarChart/barchart7.xlsx"], CHART_INTERNALS_REASON);
}

function Barchart71UsesExcelValueAxisHeadroom(): void {
  guardedTest(["BarChart/barchart7-1.xlsx"], CHART_INTERNALS_REASON);
}

function Barchart1RendersHorizontalBars(): void {
  guardedTest(["BarChart/barchart1.xlsx"], CHART_INTERNALS_REASON);
}

function Barchart2StackedBarKeepsExcelLegendOrder(): void {
  guardedTest(["BarChart/barchart2.xlsx"], CHART_INTERNALS_REASON);
}

function Barchart31PercentStackedBarRendersDenseValueTicksAndHorizontalCategoryGridlines(): void {
  guardedTest(["BarChart/barchart3-1.xlsx"], CHART_INTERNALS_REASON);
}

function Piechart1RendersSeriesNameFallbackTitleAndCategoryLegend(): void {
  guardedTest(["PieChart/piechart1.xlsx"], CHART_INTERNALS_REASON);
}

function Piechart11RendersChartAreaFill(): void {
  guardedTest(["PieChart/piechart1-1.xlsx"], CHART_INTERNALS_REASON);
}

function Piechart2RendersPerspectiveThreeDPie(): void {
  guardedTest(["PieChart/piechart2.xlsx"], CHART_INTERNALS_REASON);
}

function Piechart21UsesRotationAndChartGlow(): void {
  guardedTest(["PieChart/piechart2-1.xlsx"], CHART_INTERNALS_REASON);
}

function Piechart3RendersTruePieOfPieLayout(): void {
  guardedTest(["PieChart/piechart3.xlsx"], CHART_INTERNALS_REASON);
}

function Piechart31ParsesGradientChartTitleColor(): void {
  guardedTest(["PieChart/piechart3-1.xlsx"], CHART_INTERNALS_REASON);
}

function Piechart31RendersPlotAreaBorder(): void {
  guardedTest(["PieChart/piechart3-1.xlsx"], CHART_INTERNALS_REASON);
}

function Piechart4RendersBarOfPieSecondaryStack(): void {
  guardedTest(["PieChart/piechart4.xlsx"], CHART_INTERNALS_REASON);
}

function Piechart41RendersWeaveChartAreaAndWideBarOfPie(): void {
  guardedTest(["PieChart/piechart4-1.xlsx"], CHART_INTERNALS_REASON);
}

function Piechart5RendersThreeDoughnutRingsInExcelOrder(): void {
  guardedTest(["PieChart/piechart5.xlsx"], CHART_INTERNALS_REASON);
}

function Piechart51DoughnutLegendMatchesExcelSpacing(): void {
  guardedTest(["PieChart/piechart5-1.xlsx"], CHART_INTERNALS_REASON);
}

function Scatterchart1UsesNumericDateXValuesAndMarkers(): void {
  guardedTest(["ScatterChart/scatter1.xlsx"], CHART_INTERNALS_REASON);
}

function Scatterchart11PadsAutomaticDateAxis(): void {
  guardedTest(["ScatterChart/scatter1-1.xlsx"], CHART_INTERNALS_REASON);
}

function Scatterchart12StartsNumericXAxisAtZero(): void {
  guardedTest(["ScatterChart/scatter1-2.xlsx"], CHART_INTERNALS_REASON);
}

function Scatterchart13UsesDisplayUnitsAndXHeadroom(): void {
  guardedTest(["ScatterChart/scatter1-3.xlsx"], CHART_INTERNALS_REASON);
}

function Scatterchart14PlacesManualDisplayUnitAndLegend(): void {
  guardedTest(["ScatterChart/scatter1-4.xlsx"], CHART_INTERNALS_REASON);
}

function Scatterchart15PreservesHiddenAxesAndPageOrder(): void {
  guardedTest(["ScatterChart/scatter1-5.xlsx"], CHART_INTERNALS_REASON);
}

function Scatterchart15MatchesExcelChartTextGeometry(): void {
  guardedTest(["ScatterChart/scatter1-5.xlsx"], CHART_INTERNALS_REASON);
}

function Scatterchart15MatchesExcelCellTextGeometry(): void {
  guardedTest(["ScatterChart/scatter1-5.xlsx"], CHART_INTERNALS_REASON);
}

function Scatterchart2UsesPlaceholderTitleAndLineLegendSpacing(): void {
  guardedTest(["ScatterChart/scatter2.xlsx"], CHART_INTERNALS_REASON);
}

function Scatterchart21ReservesAxisTitlesAndUsesCompactValueTicks(): void {
  guardedTest(["ScatterChart/scatter2-1.xlsx"], CHART_INTERNALS_REASON);
}

function Scatterchart22UsesSmoothSeriesCurves(): void {
  guardedTest(["ScatterChart/scatter2-2.xlsx"], CHART_INTERNALS_REASON);
}

function Scatterchart31RendersSoftEdgeAndTallPlot(): void {
  guardedTest(["ScatterChart/scatter3-1.xlsx"], CHART_INTERNALS_REASON);
}

function Scatterchart32UsesFixedNumericAxisBand(): void {
  guardedTest(["ScatterChart/scatter3-2.xlsx"], CHART_INTERNALS_REASON);
}

function Scatterchart33DrawsCompletePlotBorder(): void {
  guardedTest(["ScatterChart/scatter3-3.xlsx"], CHART_INTERNALS_REASON);
}

function Scatterchart6RendersBubblesOnZeroCrossingAxis(): void {
  guardedTest(["ScatterChart/scatter6.xlsx"], CHART_INTERNALS_REASON);
}

function Scatterchart61UsesZeroBasedCurrencyAxisAndScaledBubbles(): void {
  guardedTest(["ScatterChart/scatter6-1.xlsx"], CHART_INTERNALS_REASON);
}

function Scatterchart7RendersSeriesLevelThreeDBubbleLighting(): void {
  guardedTest(["ScatterChart/scatter7.xlsx"], CHART_INTERNALS_REASON);
}

function Scatterchart71UsesReusableSoftThreeDBubbleTexture(): void {
  guardedTest(["ScatterChart/scatter7-1.xlsx"], CHART_INTERNALS_REASON);
}

function ColumnchartAutoPlotRectUsesTighterExcelLikeMargins(): void {
  guardedTest(["ColumnChart/columnchart1.xlsx"], CHART_INTERNALS_REASON);
}

function Areachart1KeepsExcelPlaceholderTitleAndUsesFilledAreasWithoutOutlines(): void {
  guardedTest(["areachart/areachart1.xlsx"], CHART_INTERNALS_REASON);
}

function Areachart1DefaultRowsKeepExcelChartAnchorHeight(): void {
  guardedTest(["areachart/areachart1.xlsx"], CHART_INTERNALS_REASON);
}

function Areachart2ResolvesSeriesNamesFromWorkbookForStackedDateAxis(): void {
  guardedTest(["areachart/areachart2.xlsx"], CHART_INTERNALS_REASON);
}

function Areachart21RendersMultipleStackedAreaBands(): void {
  guardedTest(["areachart/areachart2-1.xlsx"], CHART_INTERNALS_REASON);
}

function Areachart31PercentStackedPlotAreaMatchesExcelReference(): void {
  guardedTest(["areachart/areachart3-1.xlsx"], CHART_INTERNALS_REASON);
}

function Areachart4RendersTrueThreeDAreaPlanes(): void {
  guardedTest(["areachart/areachart4.xlsx"], CHART_INTERNALS_REASON);
}

function Areachart41RendersCompactMultiSeriesThreeDLayout(): void {
  guardedTest(["areachart/areachart4-1.xlsx"], CHART_INTERNALS_REASON);
}

function Areachart5RendersStackedSeriesOnOneThreeDPlane(): void {
  guardedTest(["areachart/areachart5.xlsx"], CHART_INTERNALS_REASON);
}

function Areachart51RendersGradientStackedThreeDArea(): void {
  guardedTest(["areachart/areachart5-1.xlsx"], CHART_INTERNALS_REASON);
}

function Areachart6UsesShallowPercentStackedThreeDDepth(): void {
  guardedTest(["areachart/areachart6.xlsx"], CHART_INTERNALS_REASON);
}

function Areachart7ExpandsAndReversesMonthlyDateAxis(): void {
  guardedTest(["areachart/areachart7.xlsx"], CHART_INTERNALS_REASON);
}

function Areachart71UsesPerspectiveAxisScale(): void {
  guardedTest(["areachart/areachart7-1.xlsx"], CHART_INTERNALS_REASON);
}

function Radarchart1RendersRadarGridAndSeries(): void {
  guardedTest(["radarchart/radarchart1.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart4DateAxisSortsSeriesPointsByDatePosition(): void {
  guardedTest(["linechart/linechart4.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart4RendersSeriesMarkers(): void {
  guardedTest(["linechart/linechart4.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart41ParsesMinorValueGridlines(): void {
  guardedTest(["linechart/linechart4-1.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart41RendersBlueValueAxisAndCrossTicks(): void {
  guardedTest(["linechart/linechart4-1.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart21ParsesRedChartTitle(): void {
  guardedTest(["linechart/linechart2-1.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart21RendersChartTitleBorder(): void {
  guardedTest(["linechart/linechart2-1.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart21RespectsAxisStrokesAndMonthTicks(): void {
  guardedTest(["linechart/linechart2-1.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart61ParsesTextColorsFromBlipFill(): void {
  guardedTest(["linechart/linechart6-1.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart6RendersChartAreaBorderStyle(): void {
  guardedTest(["linechart/linechart6.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart7ParsesLine3dChartMetadata(): void {
  guardedTest(["linechart/linechart7.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart7RendersProjectedFilledSeries(): void {
  guardedTest(["linechart/linechart7.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart7ReusesAxisTextAndPreservesRotation(): void {
  Line3DChartTextTests.Verify(path.join(inputRootDirectory(), "linechart", "linechart7.xlsx"), null, null);
}

function Linechart71PreservesPerspectiveAndCategoryLayout(): void {
  Line3DCategoryChartTests.Verify(path.join(inputRootDirectory(), "linechart", "linechart7-1.xlsx"));
}

function Linechart2StackedLineUsesCumulativeSeriesAndAxisRange(): void {
  guardedTest(["linechart/linechart2.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart2StackedDateAxisPlotAreaIsWider(): void {
  guardedTest(["linechart/linechart2.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart2StackedDateAxisGridlinesFollowPlot(): void {
  guardedTest(["linechart/linechart2.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart3PercentStackedLineUsesNormalizedSeriesAndAxisRange(): void {
  guardedTest(["linechart/linechart3.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart3AutomaticLayoutMatchesExcelReference(): void {
  guardedTest(["linechart/linechart3.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart3PercentageIntervalsRespectPlotHeightAndExplicitUnits(): void {
  guardedTest(["linechart/linechart3.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart3TopStrokeStaysWithinPlot(): void {
  guardedTest(["linechart/linechart3.xlsx"], CHART_INTERNALS_REASON);
}

function Linechart31SplitPagesKeepPercentTicksOnMainSliceOnly(): void {
  guardedTest(["linechart/linechart3-1.xlsx"], CHART_INTERNALS_REASON);
}

function ChartHorizontalTextUsesPdfTextWithoutPositionRegression(): void {
  guardedTest(["linechart/linechart1-1.xlsx"], PDF_INSPECTION_REASON);
}

function ChartRotatedCategoryAxisLabelsUsePdfText(): void {
  guardedTest(["linechart/linechart1-1.xlsx"], PDF_INSPECTION_REASON);
}

function TrueTypeSubsetPrototypeCanProbeHorizontalChartFonts(): void {
  guardedTest(["linechart/linechart1-1.xlsx"], PDF_INSPECTION_REASON);
}

function PdfTextIsNotEmbeddedAsFullFontPrograms(): void {
  guardedTest(["Fill.xlsx"], PDF_INSPECTION_REASON);
}

function Type3TextOptimizationKeepsTextUprightAndCompact(): void {
  guardedTest(["sales1.xlsx"], PDF_INSPECTION_REASON);
}

function LinechartTextReuseReducesSizeAndPreservesPageClips(): void {
  guardedTest(["linechart/linechart1.xlsx"], PDF_INSPECTION_REASON);
}

function PaginationAndPageSizeMatchExcelReference(): void {
  guardedTest(["sales1.xlsx"], PDF_INSPECTION_REASON);
}

function DefaultPaperSizeIsA4(): void {
  guardedTest(["sales1.xlsx"], PDF_INSPECTION_REASON);
}

function OnePagePerSheetProducesSinglePage(): void {
  guardedTest(["sales1.xlsx"], PDF_INSPECTION_REASON);
}

function AllColumnsInOnePageWidth(): void {
  guardedTest(["sales1.xlsx"], PDF_INSPECTION_REASON);
}

function DrawingShapesWorkbookRendersToA4(): void {
  guardedTest(["drawing.xlsx"], RASTER_REASON);
}

function Shape1MatchesExcelShapePrimitives(): void {
  guardedTest(["shape/shape1.xlsx"], SHAPE_INTERNALS_REASON);
}

function Shape2PreservesFreeformCalloutAndSCurve(): void {
  guardedTest(["shape/shape2.xlsx"], SHAPE_INTERNALS_REASON);
}

function Shape3PreservesAdjustedConnectorsAndEffects(): void {
  guardedTest(["shape/shape3.xlsx"], SHAPE_INTERNALS_REASON);
}

function Shape4RendersPresetDetailsAndTransparency(): void {
  guardedTest(["shape/shape4.xlsx"], SHAPE_INTERNALS_REASON);
}

function Shape4GeometryAdjustmentsAndAlphaAreBounded(): void {
  guardedTest(["shape/shape4.xlsx"], SHAPE_INTERNALS_REASON);
}

function ShapeReflectionAndPatternPaintWithClipping(): void {
  guardedTest([], SHAPE_INTERNALS_REASON);
}

function ShapeCustomGeometryValidatesNumericCommands(): void {
  guardedTest(["shape1.xlsx"], SHAPE_INTERNALS_REASON);
}

function ShapeTextboxesBypassType3TextRewrite(): void {
  guardedTest(["shape1.xlsx"], RASTER_REASON);
}

function RotatedTextWorkbookRendersToA4(): void {
  guardedTest(["rotate.xlsx"], RASTER_REASON);
}

function PicutureWorkbookMatchesExcelHorizontalPageSlices(): void {
  guardedTest(["picture.xlsx"], RASTER_REASON);
}

function PicuturePngPictureIsEmbeddedLosslessly(): void {
  guardedTest(["picture.xlsx"], RASTER_REASON);
}

function RichTextWorkbookRendersMultipleRunStyles(): void {
  guardedTest(["RichText.xlsx"], RASTER_REASON);
}

function ConditionalFormattingWorkbookRendersToA4(): void {
  guardedTest(["conditional-formatting.xlsx"], RASTER_REASON);
}

function CellValueConditionalFillIsContinuous(): void {
  guardedTest(["conditional_formatting/cell_value.xlsx"], RASTER_REASON);
}

function CellValueConditionalFontIsBold(): void {
  guardedTest(["conditional_formatting/cell_value.xlsx"], RASTER_REASON);
}

function DataBarUsesExcelDefaultLengthsWithoutBorder(): void {
  guardedTest(["conditional_formatting/data_bar.xlsx"], RASTER_REASON);
}

function DataBar1ReadsSolidFullLengthExtension(): void {
  guardedTest(["conditional_formatting/data_bar1.xlsx"], RASTER_REASON);
}

function DateRulesUseExplicitEvaluationDate(): void {
  guardedTest(["conditional_formatting/date_rules.xlsx"], RASTER_REASON);
}

function DuplicateUniqueValuesReadsBackgroundOnlyDxfFill(): void {
  guardedTest(["conditional_formatting/duplicate_unique_values.xlsx"], RASTER_REASON);
}

function FormulaRuleEvaluatesRelativeNumericComparisons(): void {
  guardedTest(["conditional_formatting/formula_rule.xlsx"], RASTER_REASON);
}

function FormulaRuleRendersItalicDifferentialFont(): void {
  guardedTest(["conditional_formatting/formula_rule.xlsx"], RASTER_REASON);
}

function IconSetUsesExcelPercentCutoffsAndTrafficLightStyle(): void {
  guardedTest(["conditional_formatting/icon_set.xlsx"], RASTER_REASON);
}

function AboveBelowAverageMatchesExcelCells(): void {
  guardedTest(["conditional_formatting/test_above_below_average.xlsx"], RASTER_REASON);
}

function CellValueRulesRenderDifferentialUnderline(): void {
  guardedTest(["conditional_formatting/test_cell_value_rules.xlsx"], RASTER_REASON);
}

function ColorScaleUsesExcelChannelInterpolation(): void {
  guardedTest(["conditional_formatting/test_color_scale.xlsx"], RASTER_REASON);
}

function FillWorkbookRendersPatternFills(): void {
  guardedTest(["Fill.xlsx"], RASTER_REASON);
}

function FillHstripeWorkbookRendersHorizontalStripes(): void {
  guardedTest(["FillHstripe.xlsx"], RASTER_REASON);
}

function SampleWorkbookKeepsShowcaseLabelFill(): void {
  guardedTest(["sample.xlsx"], RASTER_REASON);
}

function SampleWorkbookPreservesExplicitRotatedRowHeight(): void {
  guardedTest(["sample.xlsx"], RASTER_REASON);
}

function SampleWorkbookShrinkToFitKeepsTextInsideE13(): void {
  guardedTest(["sample.xlsx"], RASTER_REASON);
}

function SampleWorkbookDrawsDiagonalUpBorderInH7(): void {
  guardedTest(["sample.xlsx"], RASTER_REASON);
}

function SampleWorkbookKeepsRot180TextCenteredInH8(): void {
  guardedTest(["sample.xlsx"], RASTER_REASON);
}

function SampleWorkbookCentersRot135InG8(): void {
  guardedTest(["sample.xlsx"], RASTER_REASON);
}

function SampleWorkbookKeepsDistributedIndentVisibleInG13(): void {
  guardedTest(["sample.xlsx"], RASTER_REASON);
}

function SampleWorkbookDistributesWrappedTextInF13(): void {
  guardedTest(["sample.xlsx"], RASTER_REASON);
}

function SampleWorkbookDistributesReadingOrderCellsInRow9(): void {
  guardedTest(["sample.xlsx"], RASTER_REASON);
}

function WrappedDistributedLinePreservesSeparatorSlots(): void {
  guardedTest([], TEXT_LAYOUT_REASON);
}

function WrappedTextDoesNotCreateLinesForWordSeparators(): void {
  guardedTest([], TEXT_LAYOUT_REASON);
}

function SampleWorkbookRot255UsesTwoRightToLeftColumnsInI8(): void {
  guardedTest(["sample.xlsx"], RASTER_REASON);
}

function TableWorkbookAppliesBuiltinTableStyle(): void {
  guardedTest(["table.xlsx"], RASTER_REASON);
}

function main(): number {
  return TestRunner.run(
    "PdfRendering.ConversionTests",
    [
    new TestCase("print_area_ranges_and_scaling_modes", PrintAreaRangesAndScalingModes),
    new TestCase("0011_print_layout", PrintAreaFixtureLayout),
    new TestCase("printer_settings_public_header_and_recovery", PrinterSettingsPublicHeaderAndRecovery),
    new TestCase("0011_printer_settings_roundtrip", PrinterSettingsFixtureRoundTrip),
    new TestCase("0011_reference_output_geometry", PrinterSettingsReferenceOutputGeometry),
    new TestCase("0011_details_stacked_zero_and_tiny_segments", ChartDetailStackedZeroAndTinySegments),
    new TestCase("0011_details_multiline_category_labels", ChartDetailMultilineCategoryLabels),
    new TestCase("0011_details_axis_headroom_and_gridlines", ChartDetailAxisHeadroomAndGridlines),
    new TestCase("0011_value_axes_and_explicit_overrides", ChartAxisLegacyValueAxes),
    new TestCase("0011_column_borders_separate_touching_series", ChartAxisColumnBorders),
    new TestCase("0011_league_columns_keep_value_axis_gap", ChartAxisLeagueAxisGap),
    new TestCase("0011_chart_point_colors_labels_and_side_legends", PrintAreaChartFormatting),
    new TestCase("each_fixture_converts_to_valid_pdf", EachFixtureConvertsToValidPdf),
    new TestCase("all_input_xlsx_convert_to_output_pdf_directory", AllInputXlsxConvertToOutputPdfDirectory),
    new TestCase("numberformat_excel_fixture_displays_localized_values", NumberformatExcelFixtureDisplaysLocalizedValues),
    new TestCase("numberformat_excel_fixture_uses_selected_section_font_colors", NumberformatExcelFixtureUsesSelectedSectionFontColors),
    new TestCase("mapchart1_parses_visible_region_map", Mapchart1ParsesVisibleRegionMap),
    new TestCase("mapchart2_preserves_grouped_guidance_and_diverging_scale", Mapchart2PreservesGroupedGuidanceAndDivergingScale),
    new TestCase("linechart_category_labels_follow_chart_language", LinechartCategoryLabelsFollowChartLanguage),
    new TestCase("linechart_manual_plot_keeps_automatic_title_and_legend_in_chart_bands", LinechartManualPlotKeepsAutomaticDecorationsInChartBands),
    new TestCase("linechart_date_axis_preserves_irregular_category_values", LinechartDateAxisPreservesIrregularCategoryValues),
    new TestCase("linechart_date_axis_title_and_rotation_follow_excel_reference", LinechartDateAxisTitleAndRotationFollowExcelReference),
    new TestCase("linechart_outlier_date_axis_suppresses_rotated_labels_like_excel_reference", LinechartOutlierDateAxisSuppressesRotatedLabelsLikeExcelReference),
    new TestCase("linechart1_2_renders_horizontal_category_axis_labels", Linechart12RendersHorizontalCategoryAxisLabels),
    new TestCase("linechart1_2_category_layout_matches_excel_plot_and_legend", Linechart12CategoryLayoutMatchesExcelPlotAndLegend),
    new TestCase("linechart5_2_stacked_category_layout_matches_excel", Linechart52StackedCategoryLayoutMatchesExcel),
    new TestCase("linechart1_1_automatic_layout_matches_excel_bands", Linechart11AutomaticLayoutMatchesExcelBands),
    new TestCase("linechart_month_axis_points_snap_to_month_slots", LinechartMonthAxisPointsSnapToMonthSlots),
    new TestCase("linechart_date_axis_cross_between_insets_first_point_from_value_axis", LinechartDateAxisCrossBetweenInsetsFirstPointFromValueAxis),
    new TestCase("columnchart1_parses_gap_width_from_chart_xml", Columnchart1ParsesGapWidthFromChartXml),
    new TestCase("columnchart_clustered_geometry_uses_gap_width_to_shrink_bars", ColumnchartClusteredGeometryUsesGapWidthToShrinkBars),
    new TestCase("columnchart1_1_stacked_columns_share_single_slot_and_stack_vertically", Columnchart11StackedColumnsShareSingleSlotAndStackVertically),
    new TestCase("columnchart2_renders_axis_titles_and_manual_legend_layout", Columnchart2RendersAxisTitlesAndManualLegendLayout),
    new TestCase("columnchart2_1_renders_centered_value_data_labels", Columnchart21RendersCenteredValueDataLabels),
    new TestCase("columnchart3_renders_data_table_and_standard_error_bars", Columnchart3RendersDataTableAndStandardErrorBars),
    new TestCase("columnchart3_1_uses_tall_cross_page_percent_plot", Columnchart31UsesTallCrossPagePercentPlot),
    new TestCase("columnchart4_renders_true_3d_columns", Columnchart4RendersTrueThreeDColumns),
    new TestCase("columnchart4_1_renders_labels_dense_grid_and_axis_line", Columnchart41RendersLabelsDenseGridAndAxisLine),
    new TestCase("columnchart5_renders_stacked_3d_columns", Columnchart5RendersStackedThreeDColumns),
    new TestCase("columnchart6_renders_percent_stacked_3d_cylinders", Columnchart6RendersPercentStackedThreeDCylinders),
    new TestCase("columnchart6_1_uses_box_title_color_and_dense_percent_ticks", Columnchart61UsesBoxTitleColorAndDensePercentTicks),
    new TestCase("columnchart7_renders_textured_walls_and_series_depth", Columnchart7RendersTexturedWallsAndSeriesDepth),
    new TestCase("columnchart_standard_3d_projection_respects_view_and_gaps", ColumnchartStandardThreeDProjectionRespectsViewAndGaps),
    new TestCase("columnchart7_1_preserves_month_gaps_and_high_perspective", Columnchart71PreservesMonthGapsAndHighPerspective),
    new TestCase("columnchart_3d_month_slots_support_both_date_systems", ColumnchartThreeDMonthSlotsSupportBothDateSystems),
    new TestCase("barchart4_renders_true_3d_horizontal_bars", Barchart4RendersTrueThreeDHorizontalBars),
    new TestCase("barchart4_1_preserves_dense_grid_and_value_axis_line", Barchart41PreservesDenseGridAndValueAxisLine),
    new TestCase("barchart5_preserves_stacked_3d_bar_grouping", Barchart5PreservesStackedThreeDBarGrouping),
    new TestCase("barchart5_1_preserves_saturated_blue_chart_gradient", Barchart51PreservesSaturatedBlueChartGradient),
    new TestCase("barchart6_normalizes_percent_stacked_3d_bars", Barchart6NormalizesPercentStackedThreeDBars),
    new TestCase("barchart6_1_preserves_standard_percent_stacked_perspective", Barchart61PreservesStandardPercentStackedPerspective),
    new TestCase("barchart7_renders_textured_date_axis_3d_bars", Barchart7RendersTexturedDateAxisThreeDBars),
    new TestCase("barchart7_1_uses_excel_value_axis_headroom", Barchart71UsesExcelValueAxisHeadroom),
    new TestCase("barchart1_renders_horizontal_bars", Barchart1RendersHorizontalBars),
    new TestCase("barchart2_stacked_bar_keeps_excel_legend_order", Barchart2StackedBarKeepsExcelLegendOrder),
    new TestCase("barchart3_1_percent_stacked_bar_renders_dense_value_ticks_and_horizontal_category_gridlines", Barchart31PercentStackedBarRendersDenseValueTicksAndHorizontalCategoryGridlines),
    new TestCase("piechart1_renders_series_name_fallback_title_and_category_legend", Piechart1RendersSeriesNameFallbackTitleAndCategoryLegend),
    new TestCase("piechart1_1_renders_chart_area_fill", Piechart11RendersChartAreaFill),
    new TestCase("piechart2_renders_perspective_3d_pie", Piechart2RendersPerspectiveThreeDPie),
    new TestCase("piechart2_1_uses_rotation_and_chart_glow", Piechart21UsesRotationAndChartGlow),
    new TestCase("piechart3_renders_true_pie_of_pie_layout", Piechart3RendersTruePieOfPieLayout),
    new TestCase("piechart3_1_parses_gradient_chart_title_color", Piechart31ParsesGradientChartTitleColor),
    new TestCase("piechart3_1_renders_plot_area_border", Piechart31RendersPlotAreaBorder),
    new TestCase("piechart4_renders_bar_of_pie_secondary_stack", Piechart4RendersBarOfPieSecondaryStack),
    new TestCase("piechart4_1_renders_weave_chart_area_and_wide_bar_of_pie", Piechart41RendersWeaveChartAreaAndWideBarOfPie),
    new TestCase("piechart5_renders_three_doughnut_rings_in_excel_order", Piechart5RendersThreeDoughnutRingsInExcelOrder),
    new TestCase("piechart5_1_doughnut_legend_matches_excel_spacing", Piechart51DoughnutLegendMatchesExcelSpacing),
    new TestCase("scatterchart1_uses_numeric_date_x_values_and_markers", Scatterchart1UsesNumericDateXValuesAndMarkers),
    new TestCase("scatterchart1_1_pads_automatic_date_axis", Scatterchart11PadsAutomaticDateAxis),
    new TestCase("scatterchart1_2_starts_numeric_x_axis_at_zero", Scatterchart12StartsNumericXAxisAtZero),
    new TestCase("scatterchart1_3_uses_display_units_and_x_headroom", Scatterchart13UsesDisplayUnitsAndXHeadroom),
    new TestCase("scatterchart1_4_places_manual_display_unit_and_legend", Scatterchart14PlacesManualDisplayUnitAndLegend),
    new TestCase("scatterchart1_5_preserves_hidden_axes_and_page_order", Scatterchart15PreservesHiddenAxesAndPageOrder),
    new TestCase("scatterchart1_5_matches_excel_chart_text_geometry", Scatterchart15MatchesExcelChartTextGeometry),
    new TestCase("scatterchart1_5_matches_excel_cell_text_geometry", Scatterchart15MatchesExcelCellTextGeometry),
    new TestCase("scatterchart2_uses_placeholder_title_and_line_legend_spacing", Scatterchart2UsesPlaceholderTitleAndLineLegendSpacing),
    new TestCase("scatterchart2_1_reserves_axis_titles_and_uses_compact_value_ticks", Scatterchart21ReservesAxisTitlesAndUsesCompactValueTicks),
    new TestCase("scatterchart2_2_uses_smooth_series_curves", Scatterchart22UsesSmoothSeriesCurves),
    new TestCase("scatterchart3_1_renders_soft_edge_and_tall_plot", Scatterchart31RendersSoftEdgeAndTallPlot),
    new TestCase("scatterchart3_2_uses_fixed_numeric_axis_band", Scatterchart32UsesFixedNumericAxisBand),
    new TestCase("scatterchart3_3_draws_complete_plot_border", Scatterchart33DrawsCompletePlotBorder),
    new TestCase("scatterchart6_renders_bubbles_on_zero_crossing_axis", Scatterchart6RendersBubblesOnZeroCrossingAxis),
    new TestCase("scatterchart6_1_uses_zero_based_currency_axis_and_scaled_bubbles", Scatterchart61UsesZeroBasedCurrencyAxisAndScaledBubbles),
    new TestCase("scatterchart7_renders_series_level_3d_bubble_lighting", Scatterchart7RendersSeriesLevelThreeDBubbleLighting),
    new TestCase("scatterchart7_1_uses_reusable_soft_3d_bubble_texture", Scatterchart71UsesReusableSoftThreeDBubbleTexture),
    new TestCase("columnchart_auto_plot_rect_uses_tighter_excel_like_margins", ColumnchartAutoPlotRectUsesTighterExcelLikeMargins),
    new TestCase("areachart1_keeps_excel_placeholder_title_and_uses_filled_areas_without_outlines", Areachart1KeepsExcelPlaceholderTitleAndUsesFilledAreasWithoutOutlines),
    new TestCase("areachart1_default_rows_keep_excel_chart_anchor_height", Areachart1DefaultRowsKeepExcelChartAnchorHeight),
    new TestCase("areachart2_resolves_series_names_from_workbook_for_stacked_date_axis", Areachart2ResolvesSeriesNamesFromWorkbookForStackedDateAxis),
    new TestCase("areachart2_1_renders_multiple_stacked_area_bands", Areachart21RendersMultipleStackedAreaBands),
    new TestCase("areachart3_1_percent_stacked_plot_area_matches_excel_reference", Areachart31PercentStackedPlotAreaMatchesExcelReference),
    new TestCase("areachart4_renders_true_3d_area_planes", Areachart4RendersTrueThreeDAreaPlanes),
    new TestCase("areachart4_1_renders_compact_multi_series_3d_layout", Areachart41RendersCompactMultiSeriesThreeDLayout),
    new TestCase("areachart5_renders_stacked_series_on_one_3d_plane", Areachart5RendersStackedSeriesOnOneThreeDPlane),
    new TestCase("areachart5_1_renders_gradient_stacked_3d_area", Areachart51RendersGradientStackedThreeDArea),
    new TestCase("areachart6_uses_shallow_percent_stacked_3d_depth", Areachart6UsesShallowPercentStackedThreeDDepth),
    new TestCase("areachart7_expands_and_reverses_monthly_date_axis", Areachart7ExpandsAndReversesMonthlyDateAxis),
    new TestCase("areachart7_1_uses_perspective_axis_scale", Areachart71UsesPerspectiveAxisScale),
    new TestCase("radarchart1_renders_radar_grid_and_series", Radarchart1RendersRadarGridAndSeries),
    new TestCase("linechart4_date_axis_sorts_series_points_by_date_position", Linechart4DateAxisSortsSeriesPointsByDatePosition),
    new TestCase("linechart4_renders_series_markers", Linechart4RendersSeriesMarkers),
    new TestCase("linechart4_1_parses_minor_value_gridlines", Linechart41ParsesMinorValueGridlines),
    new TestCase("linechart4_1_renders_blue_value_axis_and_cross_ticks", Linechart41RendersBlueValueAxisAndCrossTicks),
    new TestCase("linechart2_1_parses_red_chart_title", Linechart21ParsesRedChartTitle),
    new TestCase("linechart2_1_renders_chart_title_border", Linechart21RendersChartTitleBorder),
    new TestCase("linechart2_1_respects_axis_strokes_and_month_ticks", Linechart21RespectsAxisStrokesAndMonthTicks),
    new TestCase("linechart6_1_parses_text_colors_from_blip_fill", Linechart61ParsesTextColorsFromBlipFill),
    new TestCase("linechart6_renders_chart_area_border_style", Linechart6RendersChartAreaBorderStyle),
    new TestCase("linechart7_parses_line3d_chart_metadata", Linechart7ParsesLine3dChartMetadata),
    new TestCase("linechart7_renders_projected_filled_series", Linechart7RendersProjectedFilledSeries),
    new TestCase("linechart7_reuses_axis_text_and_preserves_rotation", Linechart7ReusesAxisTextAndPreservesRotation),
    new TestCase("linechart7_1_preserves_perspective_and_category_layout", Linechart71PreservesPerspectiveAndCategoryLayout),
    new TestCase("linechart2_stacked_line_uses_cumulative_series_and_axis_range", Linechart2StackedLineUsesCumulativeSeriesAndAxisRange),
    new TestCase("linechart2_stacked_date_axis_plot_area_is_wider", Linechart2StackedDateAxisPlotAreaIsWider),
    new TestCase("linechart2_stacked_date_axis_gridlines_follow_plot", Linechart2StackedDateAxisGridlinesFollowPlot),
    new TestCase("linechart3_percent_stacked_line_uses_normalized_series_and_axis_range", Linechart3PercentStackedLineUsesNormalizedSeriesAndAxisRange),
    new TestCase("linechart3_automatic_layout_matches_excel_reference", Linechart3AutomaticLayoutMatchesExcelReference),
    new TestCase("linechart3_percentage_intervals_respect_plot_height_and_explicit_units", Linechart3PercentageIntervalsRespectPlotHeightAndExplicitUnits),
    new TestCase("linechart3_top_stroke_stays_within_plot", Linechart3TopStrokeStaysWithinPlot),
    new TestCase("linechart3_1_split_pages_keep_percent_ticks_on_main_slice_only", Linechart31SplitPagesKeepPercentTicksOnMainSliceOnly),
    new TestCase("chart_horizontal_text_uses_pdf_text_without_position_regression", ChartHorizontalTextUsesPdfTextWithoutPositionRegression),
    new TestCase("chart_rotated_category_axis_labels_use_pdf_text", ChartRotatedCategoryAxisLabelsUsePdfText),
    new TestCase("truetype_subset_prototype_can_probe_horizontal_chart_fonts", TrueTypeSubsetPrototypeCanProbeHorizontalChartFonts),
    new TestCase("pdf_text_is_not_embedded_as_full_font_programs", PdfTextIsNotEmbeddedAsFullFontPrograms),
    new TestCase("type3_text_optimization_keeps_text_upright_and_compact", Type3TextOptimizationKeepsTextUprightAndCompact),
    new TestCase("linechart_text_reuse_reduces_size_and_preserves_page_clips", LinechartTextReuseReducesSizeAndPreservesPageClips),
    new TestCase("pagination_and_page_size_match_excel_reference", PaginationAndPageSizeMatchExcelReference),
    new TestCase("default_paper_size_is_a4", DefaultPaperSizeIsA4),
    new TestCase("one_page_per_sheet_produces_single_page", OnePagePerSheetProducesSinglePage),
    new TestCase("all_columns_in_one_page_width", AllColumnsInOnePageWidth),
    new TestCase("drawing_shapes_workbook_renders_to_a4", DrawingShapesWorkbookRendersToA4),
    new TestCase("shape1_matches_excel_shape_primitives", Shape1MatchesExcelShapePrimitives),
    new TestCase("shape2_preserves_freeform_callout_and_s_curve", Shape2PreservesFreeformCalloutAndSCurve),
    new TestCase("shape3_preserves_adjusted_connectors_and_effects", Shape3PreservesAdjustedConnectorsAndEffects),
    new TestCase("shape4_renders_preset_details_and_transparency", Shape4RendersPresetDetailsAndTransparency),
    new TestCase("shape4_geometry_adjustments_and_alpha_are_bounded", Shape4GeometryAdjustmentsAndAlphaAreBounded),
    new TestCase("shape_reflection_and_pattern_paint_with_clipping", ShapeReflectionAndPatternPaintWithClipping),
    new TestCase("shape_custom_geometry_validates_numeric_commands", ShapeCustomGeometryValidatesNumericCommands),
    new TestCase("shape_textboxes_bypass_type3_text_rewrite", ShapeTextboxesBypassType3TextRewrite),
    new TestCase("rotated_text_workbook_renders_to_a4", RotatedTextWorkbookRendersToA4),
    new TestCase("picuture_workbook_matches_excel_horizontal_page_slices", PicutureWorkbookMatchesExcelHorizontalPageSlices),
    new TestCase("picuture_png_picture_is_embedded_losslessly", PicuturePngPictureIsEmbeddedLosslessly),
    new TestCase("rich_text_workbook_renders_multiple_run_styles", RichTextWorkbookRendersMultipleRunStyles),
    new TestCase("conditional_formatting_workbook_renders_to_a4", ConditionalFormattingWorkbookRendersToA4),
    new TestCase("cell_value_conditional_fill_is_continuous", CellValueConditionalFillIsContinuous),
    new TestCase("cell_value_conditional_font_is_bold", CellValueConditionalFontIsBold),
    new TestCase("data_bar_uses_excel_default_lengths_without_border", DataBarUsesExcelDefaultLengthsWithoutBorder),
    new TestCase("data_bar1_reads_solid_full_length_extension", DataBar1ReadsSolidFullLengthExtension),
    new TestCase("date_rules_use_explicit_evaluation_date", DateRulesUseExplicitEvaluationDate),
    new TestCase("duplicate_unique_values_reads_background_only_dxf_fill", DuplicateUniqueValuesReadsBackgroundOnlyDxfFill),
    new TestCase("formula_rule_evaluates_relative_numeric_comparisons", FormulaRuleEvaluatesRelativeNumericComparisons),
    new TestCase("formula_rule_renders_italic_differential_font", FormulaRuleRendersItalicDifferentialFont),
    new TestCase("icon_set_uses_excel_percent_cutoffs_and_traffic_light_style", IconSetUsesExcelPercentCutoffsAndTrafficLightStyle),
    new TestCase("above_below_average_matches_excel_cells", AboveBelowAverageMatchesExcelCells),
    new TestCase("cell_value_rules_render_differential_underline", CellValueRulesRenderDifferentialUnderline),
    new TestCase("color_scale_uses_excel_channel_interpolation", ColorScaleUsesExcelChannelInterpolation),
    new TestCase("fill_workbook_renders_pattern_fills", FillWorkbookRendersPatternFills),
    new TestCase("fill_hstripe_workbook_renders_horizontal_stripes", FillHstripeWorkbookRendersHorizontalStripes),
    new TestCase("sample_workbook_keeps_showcase_label_fill", SampleWorkbookKeepsShowcaseLabelFill),
    new TestCase("sample_workbook_preserves_explicit_rotated_row_height", SampleWorkbookPreservesExplicitRotatedRowHeight),
    new TestCase("sample_workbook_shrink_to_fit_keeps_text_inside_e13", SampleWorkbookShrinkToFitKeepsTextInsideE13),
    new TestCase("sample_workbook_draws_diagonal_up_border_in_h7", SampleWorkbookDrawsDiagonalUpBorderInH7),
    new TestCase("sample_workbook_keeps_rot180_text_centered_in_h8", SampleWorkbookKeepsRot180TextCenteredInH8),
    new TestCase("sample_workbook_centers_rot135_in_g8", SampleWorkbookCentersRot135InG8),
    new TestCase("sample_workbook_keeps_distributed_indent_visible_in_g13", SampleWorkbookKeepsDistributedIndentVisibleInG13),
    new TestCase("sample_workbook_distributes_wrapped_text_in_f13", SampleWorkbookDistributesWrappedTextInF13),
    new TestCase("sample_workbook_distributes_reading_order_cells_in_row9", SampleWorkbookDistributesReadingOrderCellsInRow9),
    new TestCase("wrapped_distributed_line_preserves_separator_slots", WrappedDistributedLinePreservesSeparatorSlots),
    new TestCase("wrapped_text_does_not_create_lines_for_word_separators", WrappedTextDoesNotCreateLinesForWordSeparators),
    new TestCase("sample_workbook_rot255_uses_two_right_to_left_columns_in_i8", SampleWorkbookRot255UsesTwoRightToLeftColumnsInI8),
    new TestCase("table_workbook_applies_builtin_table_style", TableWorkbookAppliesBuiltinTableStyle),
    ],
  );
}

process.exitCode = main();
