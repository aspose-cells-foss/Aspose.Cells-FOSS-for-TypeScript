/**
 * Public facade API for the Aspose.Cells_FOSS TypeScript port.
 */

export * from "./enums";
export * from "./Color";
export * from "./CellArea";

export * from "./CellsException";
export * from "./WorkbookLoadException";
export * from "./WorkbookSaveException";
export * from "./UnsupportedFeatureException";
export * from "./InvalidFileFormatException";
export * from "./FormulaException";
export * from "./StyleException";

export * from "./SaveOptions";
export * from "./PdfSaveOptions";
export * from "./LoadOptions";
export * from "./LoadIssue";
export * from "./LoadDiagnostics";
export * from "./WarningInfo";
export * from "./IWarningCallback";

export * from "./NumberFormat";
export * from "./FontSetting";
export * from "./Font";
export * from "./Border";
export * from "./Borders";
export * from "./style";
export * from "./StyleFlag";

export * from "./cell";
export * from "./Cells";
export * from "./Row";
export * from "./Column";
export * from "./RowCollection";
export * from "./ColumnCollection";

export * from "./worksheet";
export * from "./worksheetCollection";
export * from "./workbook";
export * from "./WorkbookSettings";
export * from "./WorkbookProperties";
export * from "./WorkbookProtection";
export * from "./WorkbookView";
export * from "./CalculationProperties";
export * from "./DocumentProperties";
export * from "./CoreDocumentProperties";
export * from "./ExtendedDocumentProperties";
export * from "./DefinedName";
export * from "./DefinedNameCollection";

export * from "./PageSetup";
export * from "./WorksheetProtection";
export * from "./autofilter";
export * from "./hyperlink";
export * from "./HyperlinkCollection";
export * from "./validation";
export * from "./ValidationCollection";
export * from "./FormatCondition";
export * from "./FormatConditionCollection";
export * from "./ConditionalFormattingCollection";
export * from "./ListObject";
export * from "./ListObjectCollection";
export * from "./Picture";
export * from "./PictureCollection";
export * from "./Shape";
export * from "./ShapeCollection";
export * from "./chart";
export * from "./chartCollection";
export * from "./comment";
export * from "./CommentCollection";

// Legacy API (pre-restructure classes, kept for backward compatibility).
export { ConditionalFormat, ConditionalFormatCollection } from "./conditionalformat";
export {
  HtmlSaveOptions,
  EncryptionType,
} from "./types";
export type {
  CellValue,
  CellCoordinates,
  CellRange,
  ColorScaleRule,
  IconSetRule,
  DataBarRule,
  ShapeFill,
  ShapeInfo,
  StraightConnectorShapeType,
  BentConnectorShapeType,
  CurvedConnectorShapeType,
  LineShapeType,
  StraightConnectorShapeInfo,
  BentConnectorShapeInfo,
  CurvedConnectorShapeInfo,
  LineShapeInfo,
  AllConnectorShapeInfo,
  RectangleShapeType,
  RectangleShapeInfo,
  EllipseShapeType,
  EllipseShapeInfo,
  TriangleShapeType,
  QuadrilateralShapeType,
  PolygonShapeType,
  TriangleShapeInfo,
  QuadrilateralShapeInfo,
  PolygonShapeInfo,
  ArrowShapeType,
  ArrowShapeInfo,
  StarShapeType,
  StarShapeInfo,
  RectangleVariantShapeType,
  RectangleVariantShapeInfo,
  CircleVariantShapeType,
  CircleVariantShapeInfo,
  BraceShapeType,
  BraceShapeInfo,
  MathSymbolShapeType,
  MathSymbolShapeInfo,
  SpecialShapeType,
  SpecialShapeInfo,
  SealShapeType,
  SealShapeInfo,
  Shape3DType,
  Shape3DInfo,
  BasicShapeSubType,
  BasicShapeInfoType,
  ChartInfo,
  AllShapeInfo,
  ImageInfo,
  PictureShapeInfo,
  PictureInfo,
} from "./types";
export type {
  Style as StyleType,
  DataValidation as DataValidationType,
  Comment as CommentType,
  Hyperlink as HyperlinkType,
} from "./types";
export {
  colToIndex,
  indexToCol,
  cellRef,
  parseCellRef,
  parseRange,
  escapeXml,
  unescapeXml,
  generateUuid,
} from "./util";
