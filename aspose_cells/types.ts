export type CellValue = string | number | boolean | Date | null;

export interface CellCoordinates {
  row: number;
  col: number;
}

export interface CellRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export interface Style {
  font?: Font;
  fill?: Fill;
  border?: Border;
  alignment?: Alignment;
  numberFormat?: string;
  protection?: Protection;
}

export interface Font {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  strikeout?: boolean;
}

export interface Fill {
  patternType?:
    | "none"
    | "solid"
    | "gray125"
    | "darkGray"
    | "mediumGray"
    | "lightGray";
  fgColor?: string;
  bgColor?: string;
}

export interface Border {
  left?: BorderLine;
  right?: BorderLine;
  top?: BorderLine;
  bottom?: BorderLine;
}

export interface BorderLine {
  style?: "thin" | "medium" | "thick" | "dashed" | "dotted";
  color?: string;
}

export interface Alignment {
  horizontal?:
    | "general"
    | "left"
    | "center"
    | "right"
    | "fill"
    | "justify"
    | "centerContinuous";
  vertical?: "top" | "center" | "bottom" | "justify";
  wrapText?: boolean;
}

export interface Protection {
  locked?: boolean;
  hidden?: boolean;
}

export interface DataValidation {
  type:
    | "any"
    | "whole"
    | "decimal"
    | "list"
    | "date"
    | "time"
    | "textLength"
    | "custom";
  operator?:
    | "between"
    | "notBetween"
    | "equal"
    | "notEqual"
    | "greaterThan"
    | "lessThan"
    | "greaterThanOrEqual"
    | "lessThanOrEqual";
  formula1?: string;
  formula2?: string;
  allowBlank?: boolean;
  showDropDown?: boolean;
  showInputMessage?: boolean;
  inputTitle?: string;
  inputMessage?: string;
  errorTitle?: string;
  errorMessage?: string;
  errorStyle?: "stop" | "warning" | "information";
}

export interface Comment {
  author?: string;
  text: string;
  row: number;
  col: number;
}

export interface Hyperlink {
  address?: string;
  display?: string;
  tooltip?: string;
  type: "url" | "file" | "email" | "internal";
  target?: string;
}

export interface FilterColumn {
  col: number;
  filters: (string | number)[];
  blank?: boolean;
}

export interface ConditionalFormatRule {
  type: "cellValue" | "expression" | "colorScale" | "iconSet" | "dataBar";
  formula1?: string;
  formula2?: string;
  operator?:
    | "between"
    | "notBetween"
    | "equal"
    | "notEqual"
    | "greaterThan"
    | "lessThan"
    | "greaterThanOrEqual"
    | "lessThanOrEqual";
  style?: Style;
  colorScale?: ColorScaleRule;
  iconSet?: IconSetRule;
  dataBar?: DataBarRule;
}

export interface ColorScaleRule {
  minColor?: string;
  midColor?: string;
  maxColor?: string;
}

export interface IconSetRule {
  iconSet?:
    | "arrows3"
    | "arrows4"
    | "arrows5"
    | "flags3"
    | "trafficLights3"
    | "trafficLights4"
    | "redToGreen"
    | "symbols3"
    | "symbols4"
    | "rating3"
    | "rating4"
    | "rating5";
  reverse?: boolean;
}

export interface DataBarRule {
  color?: string;
  minValue?: number;
  maxValue?: number;
  direction?: "leftToRight" | "rightToLeft";
}

export enum SaveFormat {
  XLSX = "xlsx",
  CSV = "csv",
  JSON = "json",
  MARKDOWN = "markdown",
  HTML = "html",
}

export enum EncryptionType {
  AES = "aes",
}

export interface ShapeFill {
  type: "solid" | "gradient" | "none";
  color?: string;
  isSchemeColor?: boolean;
  gradientStops?: Array<{
    position: number;
    color: string;
    isSchemeColor?: boolean;
    lumMod?: number;
    lumOff?: number;
  }>;
  gradientAngle?: number;
}

export interface ShapeInfo {
  name: string;
  type: string;
  fromCol: number;
  fromColOff?: number;
  fromRow: number;
  fromRowOff?: number;
  toCol: number;
  toColOff?: number;
  toRow: number;
  toRowOff?: number;
  fill?: ShapeFill;
  lineColor?: string;
  lineWidth?: number;
  flipV?: boolean;
  flipH?: boolean;
  rotation?: number;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  cx?: number;
  cy?: number;
  xfrmX?: number;
  xfrmY?: number;
  picCx?: number;
  picCy?: number;
}

export type StraightConnectorShapeType = "straightConnector1" | "straightConnector2" | "straightConnector3" | "straightConnector";
export type BentConnectorShapeType = "bentConnector1" | "bentConnector2" | "bentConnector3" | "bentConnector4" | "bentConnector5" | "elbowConnector";
export type CurvedConnectorShapeType = "curvedConnector1" | "curvedConnector2" | "curvedConnector3" | "curvedConnector4" | "curvedConnector5";
export type LineShapeType = "line";

export interface StraightConnectorShapeInfo extends ShapeInfo {
  type: StraightConnectorShapeType;
  hasArrowStart?: boolean;
  hasArrowEnd?: boolean;
}

export interface BentConnectorShapeInfo extends ShapeInfo {
  type: BentConnectorShapeType;
  hasArrowStart?: boolean;
  hasArrowEnd?: boolean;
}

export interface CurvedConnectorShapeInfo extends ShapeInfo {
  type: CurvedConnectorShapeType;
  hasArrowStart?: boolean;
  hasArrowEnd?: boolean;
}

export interface LineShapeInfo extends ShapeInfo {
  type: LineShapeType;
  hasArrowStart?: boolean;
  hasArrowEnd?: boolean;
}

export type AllConnectorShapeInfo = StraightConnectorShapeInfo | BentConnectorShapeInfo | CurvedConnectorShapeInfo | LineShapeInfo;

export type RectangleShapeType = "rect" | "rectangle";
export interface RectangleShapeInfo extends ShapeInfo {
  type: RectangleShapeType;
}

export type EllipseShapeType = "ellipse" | "oval";
export interface EllipseShapeInfo extends ShapeInfo {
  type: EllipseShapeType;
}

export type TriangleShapeType = "triangle" | "rightTriangle";
export type QuadrilateralShapeType = "parallelogram" | "trapezoid" | "diamond" | "pentagon" | "hexagon" | "heptagon" | "octagon" | "decagon" | "dodecagon" | "nonIsoscelesTrapezoid";
export type PolygonShapeType = TriangleShapeType | QuadrilateralShapeType;

export interface TriangleShapeInfo extends ShapeInfo {
  type: TriangleShapeType;
}

export interface QuadrilateralShapeInfo extends ShapeInfo {
  type: QuadrilateralShapeType;
}

export interface PolygonShapeInfo extends ShapeInfo {
  type: PolygonShapeType;
}

export type ArrowShapeType = "arrow" | "rightArrow" | "leftArrow" | "upArrow" | "downArrow" | "leftRightArrow" | "upDownArrow" | "quadrilateralArrow" | "pentagonArrow";
export interface ArrowShapeInfo extends ShapeInfo {
  type: ArrowShapeType;
}

export type StarShapeType = "star4" | "star5" | "star6" | "star7" | "star8" | "star10" | "star12" | "star16" | "star24" | "star32" | "starBurst";
export interface StarShapeInfo extends ShapeInfo {
  type: StarShapeType;
}

export type RectangleVariantShapeType = "snip1Rect" | "singleCornerSnipped" | "snip2Same" | "snip2Diag" | "snip4Same" | "snip4Diag" | "round1Rect" | "round2Same" | "round2Diag" | "round4Same" | "round4Diag";
export interface RectangleVariantShapeInfo extends ShapeInfo {
  type: RectangleVariantShapeType;
}

export type CircleVariantShapeType = "pie" | "pieWedge" | "quarterCircle" | "partialCircle" | "halfCircle" | "threeQuarterCircle" | "donut";
export interface CircleVariantShapeInfo extends ShapeInfo {
  type: CircleVariantShapeType;
}

export type BraceShapeType = "brace" | "bracket";
export interface BraceShapeInfo extends ShapeInfo {
  type: BraceShapeType;
}

export type MathSymbolShapeType = "plus" | "mathPlus" | "mathMinus" | "mathMultiply" | "mathDivide" | "mathEqual" | "mathNotEqual";
export interface MathSymbolShapeInfo extends ShapeInfo {
  type: MathSymbolShapeType;
}

export type SpecialShapeType = 
  | "ribbon" | "ribbon2" | "chevron" | "chord" | "cloud" | "cloudCallout" 
  | "sun" | "moon" | "doubleWave" | "explosion" | "lightningBolt" 
  | "heart" | "pictureFrame" | "tetris" | "smileyFace" | "flowChartProcess";
export interface SpecialShapeInfo extends ShapeInfo {
  type: SpecialShapeType;
}

export type SealShapeType = "irregularSeal1" | "irregularSeal2" | "irregularSeal3" | "irregularSeal4" | "irregularSeal5" | "irregularSeal6" | "irregularSeal7" | "irregularSeal8" | "irregularSeal9" | "irregularSeal10";
export interface SealShapeInfo extends ShapeInfo {
  type: SealShapeType;
}

export type Shape3DType = "cube" | "cylinder" | "cylinderSolid" | "prism" | "prismRight";
export interface Shape3DInfo extends ShapeInfo {
  type: Shape3DType;
}

export type BasicShapeSubType = 
  | RectangleShapeType 
  | EllipseShapeType 
  | PolygonShapeType 
  | ArrowShapeType 
  | StarShapeType
  | RectangleVariantShapeType
  | CircleVariantShapeType
  | BraceShapeType
  | MathSymbolShapeType
  | SpecialShapeType
  | SealShapeType
  | Shape3DType;

export type BasicShapeInfoType = { type: BasicShapeSubType } & ShapeInfo;

export type ChartType =
  | "bar"
  | "barStacked"
  | "barStacked100"
  | "column"
  | "columnStacked"
  | "columnStacked100"
  | "line"
  | "lineStacked"
  | "lineStacked100"
  | "pie"
  | "doughnut"
  | "area"
  | "areaStacked"
  | "areaStacked100"
  | "scatter"
  | "radar";

export interface ChartSeries {
  name?: string;
  values: number[];
  categories?: string[];
}

export interface ChartInfo extends ShapeInfo {
  chartType: ChartType;
  title?: string;
  series: ChartSeries[];
  legendPosition?: "left" | "right" | "top" | "bottom";
  categoryAxisPosition?: "left" | "right" | "top" | "bottom";
  valueAxisPosition?: "left" | "right" | "top" | "bottom";
  plotArea?: {
    x: number;
    y: number;
    width: number;
    height: number;
  };
  fromRowOff?: number;
  fromColOff?: number;
  toRowOff?: number;
  toColOff?: number;
  originalX?: number;
  originalY?: number;
  sheetIndex: number;
  originalChartXml?: string;
  originalStyleXml?: string;
  originalColorsXml?: string;
}

export type AllShapeInfo = AllConnectorShapeInfo | BasicShapeInfoType | ChartInfo;

export interface ImageInfo {
  name: string;
  data: Buffer;
  width: number;
  height: number;
  ext: string;
}

export interface PictureShapeInfo extends ShapeInfo {
  type: "picture";
  imageIndex: number;
}
