import { OpenXmlWorkbookReader } from "../Shared/OpenXmlPackage";
import { XmlNode } from "../../aspose_cells/Xml/XmlReader";
import { DefaultColorHex, normalizeDouble } from "./ComparisonValueHelpers";

/** Represents the workbook date system used by the Open XML reader. */
export enum OpenXmlDateSystem {
  Windows1900,
  Mac1904,
}

/** Shared read context passed to the Open XML snapshot factories. */
export class OpenXmlReadContext {
  public constructor(
    public readonly Reader: OpenXmlWorkbookReader,
    public readonly SharedStrings: readonly string[],
    public readonly DateSystem: OpenXmlDateSystem,
    public readonly Styles: OpenXmlStylesContext,
  ) {}
}

/** A worksheet cell discovered in an Open XML package. */
export class OpenXmlCellDescriptor {
  public constructor(
    public readonly SheetName: string,
    public readonly CellReference: string,
    public readonly Cell: XmlNode,
  ) {}
}

/** Font data captured from an Open XML style sheet. */
export class OpenXmlFontDescriptor {
  private static readonly s_default = new OpenXmlFontDescriptor(
    "Calibri",
    normalizeDouble(11),
    false,
    false,
    false,
    false,
    DefaultColorHex,
  );

  public constructor(
    public readonly Name: string,
    public readonly Size: string,
    public readonly Bold: boolean,
    public readonly Italic: boolean,
    public readonly Underline: boolean,
    public readonly StrikeThrough: boolean,
    public readonly Color: string,
  ) {}

  public static get Default(): OpenXmlFontDescriptor {
    return OpenXmlFontDescriptor.s_default;
  }
}

/** Fill data captured from an Open XML style sheet. */
export class OpenXmlFillDescriptor {
  private static readonly s_default = new OpenXmlFillDescriptor("None", DefaultColorHex, DefaultColorHex);

  public constructor(
    public readonly Pattern: string,
    public readonly ForegroundColor: string,
    public readonly BackgroundColor: string,
  ) {}

  public static get Default(): OpenXmlFillDescriptor {
    return OpenXmlFillDescriptor.s_default;
  }
}

/** Border side data captured from an Open XML style sheet. */
export class OpenXmlBorderSideDescriptor {
  private static readonly s_default = new OpenXmlBorderSideDescriptor("None", DefaultColorHex);

  public constructor(
    public readonly Style: string,
    public readonly Color: string,
  ) {}

  public static get Default(): OpenXmlBorderSideDescriptor {
    return OpenXmlBorderSideDescriptor.s_default;
  }
}

/** Border data captured from an Open XML style sheet. */
export class OpenXmlBorderDescriptor {
  private static readonly s_default = new OpenXmlBorderDescriptor(
    OpenXmlBorderSideDescriptor.Default,
    OpenXmlBorderSideDescriptor.Default,
    OpenXmlBorderSideDescriptor.Default,
    OpenXmlBorderSideDescriptor.Default,
    OpenXmlBorderSideDescriptor.Default,
    false,
    false,
  );

  public constructor(
    public readonly Left: OpenXmlBorderSideDescriptor,
    public readonly Right: OpenXmlBorderSideDescriptor,
    public readonly Top: OpenXmlBorderSideDescriptor,
    public readonly Bottom: OpenXmlBorderSideDescriptor,
    public readonly Diagonal: OpenXmlBorderSideDescriptor,
    public readonly DiagonalUp: boolean,
    public readonly DiagonalDown: boolean,
  ) {}

  public static get Default(): OpenXmlBorderDescriptor {
    return OpenXmlBorderDescriptor.s_default;
  }
}

/** Immutable data snapshot for a single cell. */
export class CellSnapshot {
  public constructor(
    public readonly SheetName: string,
    public readonly CellReference: string,
    public readonly CellType: string,
    public readonly Value: string,
    public readonly Formula: string,
  ) {}

  public ToString(): string {
    return (
      "Sheet=" +
      this.SheetName +
      "; Cell=" +
      this.CellReference +
      "; Type=" +
      this.CellType +
      "; Formula=" +
      this.Formula +
      "; Value=" +
      this.Value
    );
  }
}

/** Immutable style snapshot for a single cell. */
export class StyleSnapshot {
  private static readonly s_default = new StyleSnapshot(
    "",
    "",
    "Calibri",
    normalizeDouble(11),
    false,
    false,
    false,
    false,
    DefaultColorHex,
    "None",
    DefaultColorHex,
    DefaultColorHex,
    "None",
    DefaultColorHex,
    "None",
    DefaultColorHex,
    "None",
    DefaultColorHex,
    "None",
    DefaultColorHex,
    "None",
    DefaultColorHex,
    false,
    false,
    "General",
    "Bottom",
    false,
    0,
    0,
    false,
    0,
    0,
    0,
    "",
    true,
    false,
  );

  public constructor(
    public readonly SheetName: string,
    public readonly CellReference: string,
    public readonly FontName: string,
    public readonly FontSize: string,
    public readonly FontBold: boolean,
    public readonly FontItalic: boolean,
    public readonly FontUnderline: boolean,
    public readonly FontStrikeThrough: boolean,
    public readonly FontColor: string,
    public readonly FillPattern: string,
    public readonly FillForegroundColor: string,
    public readonly FillBackgroundColor: string,
    public readonly LeftBorderStyle: string,
    public readonly LeftBorderColor: string,
    public readonly RightBorderStyle: string,
    public readonly RightBorderColor: string,
    public readonly TopBorderStyle: string,
    public readonly TopBorderColor: string,
    public readonly BottomBorderStyle: string,
    public readonly BottomBorderColor: string,
    public readonly DiagonalBorderStyle: string,
    public readonly DiagonalBorderColor: string,
    public readonly DiagonalUp: boolean,
    public readonly DiagonalDown: boolean,
    public readonly HorizontalAlignment: string,
    public readonly VerticalAlignment: string,
    public readonly WrapText: boolean,
    public readonly IndentLevel: number,
    public readonly TextRotation: number,
    public readonly ShrinkToFit: boolean,
    public readonly ReadingOrder: number,
    public readonly RelativeIndent: number,
    public readonly NumberFormatId: number,
    public readonly NumberFormatCode: string,
    public readonly IsLocked: boolean,
    public readonly IsHidden: boolean,
  ) {}

  public static get Default(): StyleSnapshot {
    return StyleSnapshot.s_default;
  }

  public CopyWithLocation(sheetName: string, cellReference: string): StyleSnapshot {
    return new StyleSnapshot(
      sheetName,
      cellReference,
      this.FontName,
      this.FontSize,
      this.FontBold,
      this.FontItalic,
      this.FontUnderline,
      this.FontStrikeThrough,
      this.FontColor,
      this.FillPattern,
      this.FillForegroundColor,
      this.FillBackgroundColor,
      this.LeftBorderStyle,
      this.LeftBorderColor,
      this.RightBorderStyle,
      this.RightBorderColor,
      this.TopBorderStyle,
      this.TopBorderColor,
      this.BottomBorderStyle,
      this.BottomBorderColor,
      this.DiagonalBorderStyle,
      this.DiagonalBorderColor,
      this.DiagonalUp,
      this.DiagonalDown,
      this.HorizontalAlignment,
      this.VerticalAlignment,
      this.WrapText,
      this.IndentLevel,
      this.TextRotation,
      this.ShrinkToFit,
      this.ReadingOrder,
      this.RelativeIndent,
      this.NumberFormatId,
      this.NumberFormatCode,
      this.IsLocked,
      this.IsHidden,
    );
  }

  public Equals(other: StyleSnapshot | null | undefined): boolean {
    if (other === null || other === undefined) {
      return false;
    }

    return (
      this.SheetName === other.SheetName &&
      this.CellReference === other.CellReference &&
      this.FontName === other.FontName &&
      this.FontSize === other.FontSize &&
      this.FontBold === other.FontBold &&
      this.FontItalic === other.FontItalic &&
      this.FontUnderline === other.FontUnderline &&
      this.FontStrikeThrough === other.FontStrikeThrough &&
      this.FontColor === other.FontColor &&
      this.FillPattern === other.FillPattern &&
      this.FillForegroundColor === other.FillForegroundColor &&
      this.FillBackgroundColor === other.FillBackgroundColor &&
      this.LeftBorderStyle === other.LeftBorderStyle &&
      this.LeftBorderColor === other.LeftBorderColor &&
      this.RightBorderStyle === other.RightBorderStyle &&
      this.RightBorderColor === other.RightBorderColor &&
      this.TopBorderStyle === other.TopBorderStyle &&
      this.TopBorderColor === other.TopBorderColor &&
      this.BottomBorderStyle === other.BottomBorderStyle &&
      this.BottomBorderColor === other.BottomBorderColor &&
      this.DiagonalBorderStyle === other.DiagonalBorderStyle &&
      this.DiagonalBorderColor === other.DiagonalBorderColor &&
      this.DiagonalUp === other.DiagonalUp &&
      this.DiagonalDown === other.DiagonalDown &&
      this.HorizontalAlignment === other.HorizontalAlignment &&
      this.VerticalAlignment === other.VerticalAlignment &&
      this.WrapText === other.WrapText &&
      this.IndentLevel === other.IndentLevel &&
      this.TextRotation === other.TextRotation &&
      this.ShrinkToFit === other.ShrinkToFit &&
      this.ReadingOrder === other.ReadingOrder &&
      this.RelativeIndent === other.RelativeIndent &&
      this.NumberFormatId === other.NumberFormatId &&
      this.NumberFormatCode === other.NumberFormatCode &&
      this.IsLocked === other.IsLocked &&
      this.IsHidden === other.IsHidden
    );
  }

  public ToString(): string {
    return (
      "Sheet=" +
      this.SheetName +
      "; Cell=" +
      this.CellReference +
      "; Font=" +
      this.FontName +
      "/" +
      this.FontSize +
      "/B=" +
      this.FontBold +
      "/I=" +
      this.FontItalic +
      "/U=" +
      this.FontUnderline +
      "/S=" +
      this.FontStrikeThrough +
      "/Color=" +
      this.FontColor +
      "; Borders=L(" +
      this.LeftBorderStyle +
      "," +
      this.LeftBorderColor +
      ")/R(" +
      this.RightBorderStyle +
      "," +
      this.RightBorderColor +
      ")/T(" +
      this.TopBorderStyle +
      "," +
      this.TopBorderColor +
      ")/B(" +
      this.BottomBorderStyle +
      "," +
      this.BottomBorderColor +
      ")/D(" +
      this.DiagonalBorderStyle +
      "," +
      this.DiagonalBorderColor +
      ",Up=" +
      this.DiagonalUp +
      ",Down=" +
      this.DiagonalDown +
      "); Align=H:" +
      this.HorizontalAlignment +
      "/V:" +
      this.VerticalAlignment +
      "/Wrap=" +
      this.WrapText +
      "/Indent=" +
      this.IndentLevel +
      "/Rotation=" +
      this.TextRotation +
      "/Shrink=" +
      this.ShrinkToFit +
      "/Reading=" +
      this.ReadingOrder +
      "/RelativeIndent=" +
      this.RelativeIndent +
      "; Fill=" +
      this.FillPattern +
      "/" +
      this.FillForegroundColor +
      "/" +
      this.FillBackgroundColor +
      "; Number=" +
      this.NumberFormatId +
      "/" +
      this.NumberFormatCode +
      "; Protection=Locked:" +
      this.IsLocked +
      "/Hidden:" +
      this.IsHidden
    );
  }
}

/** Style data parsed from an Open XML style sheet. */
export class OpenXmlStylesContext {
  private static readonly s_default = new OpenXmlStylesContext([StyleSnapshot.Default], new Set<number>());

  public constructor(
    public readonly CellFormats: readonly StyleSnapshot[],
    public readonly DateStyleIndexes: ReadonlySet<number>,
  ) {}

  public static get Default(): OpenXmlStylesContext {
    return OpenXmlStylesContext.s_default;
  }
}

/** A data or style mismatch between the library and the Open XML reader. */
export class SnapshotMismatch {
  public constructor(
    public readonly CellKey: string,
    public readonly LibrarySnapshot: string,
    public readonly OpenXmlSnapshot: string,
  ) {}
}

/**
 * Dictionary keyed case-insensitively by cell key, mirroring the .NET
 * Dictionary<string, T> with StringComparer.OrdinalIgnoreCase.
 */
export class CaseInsensitiveStringMap<T> {
  private readonly items = new Map<string, { key: string; value: T }>();

  public set(key: string, value: T): void {
    this.items.set(key.toLowerCase(), { key, value });
  }

  public get(key: string): T | undefined {
    const item = this.items.get(key.toLowerCase());
    return item === undefined ? undefined : item.value;
  }

  public has(key: string): boolean {
    return this.items.has(key.toLowerCase());
  }

  public keys(): string[] {
    return Array.from(this.items.values(), (item) => item.key);
  }

  public get Count(): number {
    return this.items.size;
  }
}
