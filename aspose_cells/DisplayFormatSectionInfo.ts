/**
 * Describes a single section of an Excel number format code.
 *
 * Ported from Aspose.Cells_FOSS.DisplayFormatSectionInfo.
 */
export class DisplayFormatSectionInfo {
  /** Raw section text as it appears between semicolon separators. */
  public Raw: string = "";

  /** Whether the section starts with a conditional token such as "[>100]". */
  public HasCondition: boolean = false;

  /** Comparison operator parsed from the conditional token. */
  public ConditionOperator: string = "";

  /** Comparison value parsed from the conditional token. */
  public ConditionValue: number = 0;
}
