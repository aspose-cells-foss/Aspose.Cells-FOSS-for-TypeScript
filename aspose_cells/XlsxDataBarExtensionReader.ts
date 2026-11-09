import { FormatConditionModel } from "./Core/FormatConditionModel";
import { FormatConditionType } from "./FormatConditionType";
import { attribute, elementChildren, descendants, firstElement, XmlNode } from "./Xml/XmlReader";

/** Applies the x14 data bar extension to a data bar condition, mirroring XlsxDataBarExtensionReader.Apply. */
export function applyDataBarExtension(
  worksheetRoot: XmlNode,
  ruleElement: XmlNode,
  condition: FormatConditionModel,
): void {
  if (condition.Type !== FormatConditionType.DataBar) {
    return;
  }

  let ruleId: string | null = null;
  const idElements = descendants(ruleElement, "id");
  if (idElements.length > 0) {
    ruleId = idElements[0].textContent ?? "";
  }

  const extensions = firstElement(worksheetRoot, "extLst");
  if (ruleId === null || ruleId.length === 0 || extensions === null) {
    return;
  }

  for (const extendedRule of descendants(extensions, "cfRule")) {
    if (!equalsIgnoreCase(attribute(extendedRule, "id"), ruleId)) {
      continue;
    }

    const dataBar = firstElement(extendedRule, "dataBar");
    if (dataBar === null) {
      return;
    }

    const gradient = attribute(dataBar, "gradient");
    if (gradient !== null) {
      condition.DataBarGradient = gradient !== "0" && !equalsIgnoreCase(gradient, "false");
    }

    const minLength = tryReadLength(attribute(dataBar, "minLength"));
    if (minLength !== null) {
      condition.DataBarMinLength = minLength;
    }

    const maxLength = tryReadLength(attribute(dataBar, "maxLength"));
    if (maxLength !== null) {
      condition.DataBarMaxLength = maxLength;
    }

    if (condition.DataBarMinLength > condition.DataBarMaxLength) {
      condition.DataBarMinLength = 0.1;
      condition.DataBarMaxLength = 0.9;
    }

    for (const threshold of elementChildren(dataBar, "cfvo")) {
      condition.DataBarAutoMinimum = equalsIgnoreCase(attribute(threshold, "type"), "autoMin");
      break;
    }

    return;
  }
}

function tryReadLength(value: string | null): number | null {
  if (value === null) {
    return null;
  }

  const percent = Number(value);
  if (!Number.isFinite(percent) || percent < 0 || percent > 100) {
    return null;
  }

  return percent / 100;
}

function equalsIgnoreCase(left: string | null, right: string): boolean {
  return left !== null && left.toLowerCase() === right.toLowerCase();
}
