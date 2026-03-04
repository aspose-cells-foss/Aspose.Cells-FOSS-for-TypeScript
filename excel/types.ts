export type CellValue = string | number | boolean | Date | null

export interface CellCoordinates {
  row: number
  col: number
}

export interface CellRange {
  startRow: number
  startCol: number
  endRow: number
  endCol: number
}

export interface Style {
  font?: Font
  fill?: Fill
  border?: Border
  alignment?: Alignment
  numberFormat?: string
  protection?: Protection
}

export interface Font {
  name?: string
  size?: number
  bold?: boolean
  italic?: boolean
  underline?: boolean
  color?: string
  strikeout?: boolean
}

export interface Fill {
  patternType?: "none" | "solid" | "gray125" | "darkGray" | "mediumGray" | "lightGray"
  fgColor?: string
  bgColor?: string
}

export interface Border {
  left?: BorderLine
  right?: BorderLine
  top?: BorderLine
  bottom?: BorderLine
}

export interface BorderLine {
  style?: "thin" | "medium" | "thick" | "dashed" | "dotted"
  color?: string
}

export interface Alignment {
  horizontal?: "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous"
  vertical?: "top" | "center" | "bottom" | "justify"
  wrapText?: boolean
}

export interface Protection {
  locked?: boolean
  hidden?: boolean
}

export interface DataValidation {
  type: "any" | "whole" | "decimal" | "list" | "date" | "time" | "textLength" | "custom"
  operator?:
    | "between"
    | "notBetween"
    | "equal"
    | "notEqual"
    | "greaterThan"
    | "lessThan"
    | "greaterThanOrEqual"
    | "lessThanOrEqual"
  formula1?: string
  formula2?: string
  allowBlank?: boolean
  showDropDown?: boolean
  showInputMessage?: boolean
  inputTitle?: string
  inputMessage?: string
  errorTitle?: string
  errorMessage?: string
  errorStyle?: "stop" | "warning" | "information"
}

export interface Comment {
  author?: string
  text: string
  row: number
  col: number
}

export interface Hyperlink {
  address?: string
  display?: string
  tooltip?: string
  type: "url" | "file" | "email" | "internal"
  target?: string
}

export interface FilterColumn {
  col: number
  filters: (string | number)[]
  blank?: boolean
}

export interface ConditionalFormatRule {
  type: "cellValue" | "expression" | "colorScale" | "iconSet" | "dataBar"
  formula1?: string
  formula2?: string
  operator?:
    | "between"
    | "notBetween"
    | "equal"
    | "notEqual"
    | "greaterThan"
    | "lessThan"
    | "greaterThanOrEqual"
    | "lessThanOrEqual"
  style?: Style
  colorScale?: ColorScaleRule
  iconSet?: IconSetRule
  dataBar?: DataBarRule
}

export interface ColorScaleRule {
  minColor?: string
  midColor?: string
  maxColor?: string
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
    | "rating5"
  reverse?: boolean
}

export interface DataBarRule {
  color?: string
  minValue?: number
  maxValue?: number
  direction?: "leftToRight" | "rightToLeft"
}

export enum SaveFormat {
  XLSX = "xlsx",
  CSV = "csv",
  JSON = "json",
  MARKDOWN = "markdown",
}

export enum EncryptionType {
  AES = "aes",
}
