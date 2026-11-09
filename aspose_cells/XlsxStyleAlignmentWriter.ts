import { AlignmentValue } from "./Core/AlignmentValue";
import { XmlElement } from "./Xml/XmlWriter";
import { getHorizontalAlignmentName, getVerticalAlignmentName } from "./XlsxWorkbookStylesValueHelpers";

/** Builds the optional alignment element used by cell formats and differential formats. */
export function buildStyleAlignment(alignment: AlignmentValue): XmlElement | null {
  const element = XmlElement.create("alignment");
  let hasValue = false;

  const horizontal = getHorizontalAlignmentName(alignment.Horizontal);
  if (horizontal.length > 0) {
    element.attr("horizontal", horizontal);
    hasValue = true;
  }

  const vertical = getVerticalAlignmentName(alignment.Vertical);
  if (vertical.length > 0) {
    element.attr("vertical", vertical);
    hasValue = true;
  }

  if (alignment.WrapText) {
    element.attr("wrapText", 1);
    hasValue = true;
  }

  if (alignment.IndentLevel > 0) {
    element.attr("indent", alignment.IndentLevel);
    hasValue = true;
  }

  if (alignment.TextRotation !== 0) {
    element.attr("textRotation", alignment.TextRotation);
    hasValue = true;
  }

  if (alignment.ShrinkToFit) {
    element.attr("shrinkToFit", 1);
    hasValue = true;
  }

  if (alignment.ReadingOrder !== 0) {
    element.attr("readingOrder", alignment.ReadingOrder);
    hasValue = true;
  }

  if (alignment.RelativeIndent !== 0) {
    element.attr("relativeIndent", alignment.RelativeIndent);
    hasValue = true;
  }

  return hasValue ? element : null;
}
