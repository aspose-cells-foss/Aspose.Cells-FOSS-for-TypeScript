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
  color?: string; // hex color or scheme color name (e.g., "accent1" or "#4472c4")
  isSchemeColor?: boolean; // if true, color is a scheme color name
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
  type: string; // 'line','rect','ellipse','triangle','straightConnector1','bentConnector3', etc.
  fromCol: number;
  fromColOff?: number; // EMU
  fromRow: number;
  fromRowOff?: number; // EMU
  toCol: number;
  toColOff?: number;
  toRow: number;
  toRowOff?: number;
  fill?: ShapeFill;
  fillColor?: string; // resolved hex color (legacy)
  lineColor?: string;
  lineWidth?: number; // EMU
  flipV?: boolean;
  flipH?: boolean;
  rotation?: number; // rotation in degrees (1/60000 degree units stored in Excel)
  hasArrowEnd?: boolean;
  isConnector?: boolean;
  x?: number; // pre-calculated EMU position
  y?: number;
  width?: number;
  height?: number;
}

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
  sheetIndex: number;
}
