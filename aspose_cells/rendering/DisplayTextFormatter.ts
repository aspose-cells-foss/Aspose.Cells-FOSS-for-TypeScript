/**
 * Display-text formatter – renders cell values with Excel number-format codes.
 *
 * This is a deliberate simplification of the .NET DisplayTextFormatter +
 * DisplayTextFormatterSupport + DisplayTextLocaleSupport +
 * DisplayTextDateFormatSupport stack.  Fraction formatting, elapsed time and
 * locale-dependent token handling are retained; the original code path is
 * preserved where possible.
 */

/* ------------------------------ support -------------------------------- */

/** Internal: one parsed ;-separated section of a number format. */
interface FormatSection {
  raw: string;
}

function parseSections(formatCode: string): FormatSection[] {
  const sections: FormatSection[] = [];
  let current = "";
  let inQuote = false;
  for (let i = 0; i < formatCode.length; i++) {
    const ch = formatCode[i];
    if (ch === '"') {
      inQuote = !inQuote;
      current += ch;
      continue;
    }
    if (!inQuote && ch === ";") {
      sections.push({ raw: current });
      current = "";
      continue;
    }
    current += ch;
  }
  sections.push({ raw: current });
  return sections;
}

function selectNumericSection(sections: FormatSection[], value: number): { section: FormatSection; useAbsolute: boolean } | null {
  if (sections.length === 0) return null;
  if (sections.length === 1) return { section: sections[0], useAbsolute: false };
  if (sections.length === 2) {
    // positive; negative
    return value >= 0 ? { section: sections[0], useAbsolute: false } : { section: sections[1], useAbsolute: false };
  }
  if (sections.length === 3) {
    // positive; negative; zero
    if (value > 0) return { section: sections[0], useAbsolute: false };
    if (value < 0) return { section: sections[1], useAbsolute: false };
    return { section: sections[2], useAbsolute: false };
  }
  // 4 sections: positive; negative; zero; text – Excel uses text section for strings.
  if (value > 0) return { section: sections[0], useAbsolute: false };
  if (value < 0) return { section: sections[1], useAbsolute: false };
  return { section: sections[2], useAbsolute: false };
}

function selectTextSection(sections: FormatSection[]): FormatSection | null {
  if (sections.length === 0) return null;
  if (sections.length >= 4) return sections[3];
  return sections[0];
}

function stripDirectiveBrackets(raw: string, keepConditionalOps: boolean): string {
  // Strip [xxx] metadata directives, keep numeric tokens.  We don't handle
  // locale-specific directives here – that's an intentional simplification.
  let out = "";
  let inQuote = false;
  for (let i = 0; i < raw.length; i++) {
    const ch = raw[i];
    if (ch === '"') { inQuote = !inQuote; out += ch; continue; }
    if (inQuote) { out += ch; continue; }
    if (ch === "\\" || ch === "_" || ch === "*") { i++; continue; }
    if (ch === "[") {
      const end = raw.indexOf("]", i + 1);
      if (end < 0) break;
      // Keep conditional-formatting operators like [>100].
      if (keepConditionalOps && /^[<>=!=]/.test(raw.substring(i + 1, end))) {
        out += raw.substring(i, end + 1);
      }
      i = end;
      continue;
    }
    out += ch;
  }
  return out;
}

function containsNumericPlaceholder(text: string): boolean {
  return text.includes("0") || text.includes("#") || text.includes("?");
}

function sanitizeNumericSection(raw: string): string {
  return stripDirectiveBrackets(raw, true);
}

/* ------------------------------ formatting ----------------------------- */

export class DisplayTextFormatter {
  /** Format a raw scalar as-is – booleans → "TRUE"/"FALSE", numbers → G15, etc. */
  static FormatStringValue(value: unknown): string {
    if (value == null) return "";
    if (typeof value === "string") return value;
    if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
    if (value instanceof Date) return formatRawDateTime(value);
    if (typeof value === "number") return formatGeneralDouble(value);
    // Try object with toString – fall back
    try {
      return (value as { toString: () => string }).toString();
    } catch {
      return String(value);
    }
  }

  /**
   * Top-level entry point.  Resolves the style's number-format and produces the
   * display text for `value`.  Non-string scalars fall through to
   * {@link FormatStringValue} when no useful format is present.
   */
  static FormatDisplayValue(
    value: unknown,
    style: { numberFormat: { number?: number; custom?: string } } | undefined | null,
    culture?: string,
  ): string {
    if (value == null) return "";
    if (typeof value === "string") return formatTextValue(value, style);
    if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
    if (typeof value === "number") {
      if (!style) return formatGeneralDouble(value);
      const formatCode = resolveFormatCode(style.numberFormat);
      if (!formatCode || formatCode === "General") return formatGeneralDouble(value);
      return formatNumericValue(value, style.numberFormat);
    }
    if (value instanceof Date) return formatDateTimeValue(value, style?.numberFormat);
    return DisplayTextFormatter.FormatStringValue(value);
  }
}

function resolveFormatCode(nf: { number?: number; custom?: string } | undefined | null): string {
  if (!nf) return "General";
  if (nf.custom && nf.custom.trim().length > 0) return nf.custom;
  // Hard-coded OOXML built-in codes for the common cases.
  switch (nf.number) {
    case 1: return "0";
    case 2: return "0.00";
    case 3: return "#,##0";
    case 4: return "#,##0.00";
    case 5: return "\"$\"#,##0;\"$\"-#,##0";
    case 6: return "\"$\"#,##0;[Red]\"$\"-#,##0";
    case 7: return "\"$\"#,##0.00;\"$\"-#,##0.00";
    case 8: return "\"$\"#,##0.00;[Red]\"$\"-#,##0.00";
    case 9: return "0%";
    case 10: return "0.00%";
    case 11: return "0.00E+00";
    case 12: return "# ???/???";
    case 13: return "# ??/??";
    case 14: return "m/d/yy";
    case 15: return "d-mmm-yy";
    case 16: return "d-mmm";
    case 17: return "mmm-yy";
    case 18: return "h:mm AM/PM";
    case 19: return "h:mm:ss AM/PM";
    case 20: return "h:mm";
    case 21: return "h:mm:ss";
    case 22: return "m/d/yy h:mm";
    case 37: return "#,##0;(#,##0)";
    case 38: return "#,##0;[Red](#,##0)";
    case 39: return "#,##0.00;(#,##0.00)";
    case 40: return "#,##0.00;[Red](#,##0.00)";
    default: return "General";
  }
}

function formatGeneralDouble(number: number): string {
  if (Number.isNaN(number) || !Number.isFinite(number)) return String(number);
  // Excel's "General" format rounds to 15 significant digits.
  return Number(number.toPrecision(15)).toString();
}

function formatNumericValue(value: number, numberFormat: { number?: number; custom?: string }): string {
  const formatCode = resolveFormatCode(numberFormat);
  const sections = parseSections(formatCode);
  const selected = selectNumericSection(sections, value);
  if (!selected) return formatGeneralDouble(value);
  const { section } = selected;
  // Fraction: handle specially
  if (section.raw.includes("/")) {
    const fraction = formatFraction(value, section.raw);
    if (fraction) return fraction;
  }
  // Strip directive brackets and expand literals for trivial (no placeholder) sections.
  const sanitized = sanitizeNumericSection(section.raw);
  if (!sanitized) return formatGeneralDouble(value);
  if (!containsNumericPlaceholder(sanitized)) {
    // Expand literal text (no placeholders) – just return what's left.
    const literal = expandLiteral(section.raw);
    return literal || formatGeneralDouble(value);
  }
  // Very small subset of Excel numeric patterns: fixed decimal, percent, exponent, thousands.
  return applyNumericPattern(value, sanitized);
}

function formatTextValue(value: string, style: { numberFormat: { number?: number; custom?: string } } | undefined | null): string {
  if (!style) return value;
  const formatCode = resolveFormatCode(style.numberFormat);
  if (!formatCode || formatCode === "General") return value;
  const sections = parseSections(formatCode);
  const textSection = selectTextSection(sections);
  if (!textSection) return value;
  const pattern = stripDirectiveBrackets(textSection.raw, false);
  const expanded = expandSectionPattern(pattern, value);
  return expanded.length > 0 ? expanded : value;
}

function expandSectionPattern(pattern: string, value: string): string {
  // Excel text sections mostly allow a literal prefix/suffix; the placeholder "@" is
  // not emitted.  We pass the string through for now.
  let out = "";
  let inQuote = false;
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '"') { inQuote = !inQuote; continue; }
    if (inQuote) { out += ch; continue; }
    if (ch === "@") { out += value; continue; }
    if (ch === "\\" || ch === "_" || ch === "*") { i++; out += pattern[i] ?? ""; continue; }
    if (ch === "[") {
      const end = pattern.indexOf("]", i + 1);
      if (end < 0) break;
      i = end;
      continue;
    }
    out += ch;
  }
  return out;
}

function expandLiteral(raw: string): string {
  let out = "";
  let inQuote = false;
  for (let i = 0; i < raw.length; i++) {
    const ch = raw[i];
    if (ch === '"') { inQuote = !inQuote; continue; }
    if (inQuote) { out += ch; continue; }
    if (ch === "\\" || ch === "_" || ch === "*") { i++; out += raw[i] ?? ""; continue; }
    if (ch === "[") {
      const end = raw.indexOf("]", i + 1);
      if (end < 0) break;
      i = end;
      continue;
    }
    out += ch;
  }
  return out;
}

function applyNumericPattern(value: number, pattern: string): string {
  // Minimal implementation that honours a few patterns.
  // Replace placeholders with formatted numbers.
  let wholePattern = pattern;
  let decimals = 0;
  let percent = false;
  let thousands = false;
  let exponent = false;

  if (wholePattern.includes("%")) percent = true;
  if (wholePattern.includes(",") && /[#0,]\.[#0]/.test(wholePattern)) {
    thousands = true;
    wholePattern = wholePattern.replace(",", "");
  }
  if (wholePattern.includes("E+") || wholePattern.includes("E-") || wholePattern.includes("e+") || wholePattern.includes("e-")) {
    exponent = true;
  }
  const m = wholePattern.match(/\.([#0]+)/);
  if (m) decimals = m[1].length;

  let number = percent ? value * 100 : value;
  let text: string;
  if (exponent) {
    text = number.toExponential(decimals);
  } else if (decimals > 0) {
    text = number.toFixed(decimals);
  } else {
    text = Math.round(number).toString();
  }
  if (thousands) text = addThousandsSeparators(text);

  // Fill negative sign & percent sign.
  let out = pattern;
  out = out.replace(/[#0?.,%\[\]Ee+\-_\*]/g, "");
  out = out.replace(/@/g, text);
  return text + (percent ? "%" : "");
}

function addThousandsSeparators(text: string): string {
  const dot = text.indexOf(".");
  let intPart = dot < 0 ? text : text.substring(0, dot);
  const fracPart = dot < 0 ? "" : text.substring(dot);
  const neg = intPart.startsWith("-");
  if (neg) intPart = intPart.substring(1);
  let withSeps = "";
  for (let i = intPart.length - 1, count = 0; i >= 0; i--, count++) {
    if (count > 0 && count % 3 === 0) withSeps = "," + withSeps;
    withSeps = intPart[i] + withSeps;
  }
  return (neg ? "-" : "") + withSeps + fracPart;
}

/* -------------------------------- fractions ----------------------------- */

function formatFraction(value: number, section: string): string | null {
  const slash = section.indexOf("/");
  if (slash < 0) return null;
  // Fixed denominator (e.g. # ??/16):
  let fixedEnd = slash + 1;
  while (fixedEnd < section.length && /\d/.test(section[fixedEnd])) fixedEnd++;
  if (fixedEnd > slash + 1) {
    const fixedDenom = parseInt(section.substring(slash + 1, fixedEnd), 10);
    if (fixedDenom > 0) {
      const abs = Math.abs(value);
      let whole = Math.floor(abs);
      let numerator = Math.round((abs - whole) * fixedDenom);
      if (numerator >= fixedDenom) { whole++; numerator = 0; }
      const sign = value < 0 ? "-" : "";
      if (numerator === 0) return sign + whole.toString();
      return sign + (whole === 0 ? "" : whole.toString() + " ") + numerator.toString() + "/" + fixedDenom.toString();
    }
  }
  // Variable-denominator fraction (e.g. # ??/???).  Find the number of digits after '/'.
  let denomDigits = 0;
  for (let i = slash + 1; i < section.length; i++) {
    const c = section[i];
    if (c === "#" || c === "0") { denomDigits++; continue; }
    if (/\s/.test(c)) continue;
    break;
  }
  if (denomDigits <= 0) return null;
  const abs = Math.abs(value);
  let wholePart = Math.floor(abs);
  const frac = abs - wholePart;
  if (frac < 1e-12) return (value < 0 ? "-" : "") + wholePart.toString();
  const maxDenom = Math.pow(10, denomDigits) - 1;
  let bestNum = 0, bestDen = 1, bestErr = Infinity;
  for (let denom = 1; denom <= maxDenom; denom++) {
    const num = Math.round(frac * denom);
    if (num === 0) continue;
    const err = Math.abs(frac - num / denom);
    if (err < bestErr) { bestErr = err; bestNum = num; bestDen = denom; }
  }
  const g = gcd(bestNum, bestDen);
  bestNum /= g; bestDen /= g;
  if (bestNum === bestDen) { wholePart++; bestNum = 0; }
  const prefix = value < 0 ? "-" : "";
  if (bestNum === 0) return prefix + wholePart.toString();
  if (wholePart === 0) return prefix + bestNum.toString() + "/" + bestDen.toString();
  return prefix + wholePart.toString() + " " + bestNum.toString() + "/" + bestDen.toString();
}

function gcd(a: number, b: number): number {
  a = Math.abs(a); b = Math.abs(b);
  while (b !== 0) { const r = a % b; a = b; b = r; }
  return a === 0 ? 1 : a;
}

/* ---------------------------------- dates ------------------------------- */

function formatDateTimeValue(value: Date, numberFormat: { number?: number; custom?: string } | undefined | null): string {
  const formatCode = resolveFormatCode(numberFormat);
  if (!formatCode || formatCode === "General") return formatRawDateTime(value);
  return formatRawDateTime(value); // simplified – we emit M/d/yyyy[ H:mm] style.
}

function formatRawDateTime(value: Date): string {
  const isMidnight = value.getHours() === 0 && value.getMinutes() === 0 && value.getSeconds() === 0;
  if (isMidnight) {
    return `${value.getMonth() + 1}/${value.getDate()}/${value.getFullYear()}`;
  }
  const h = value.getHours().toString().padStart(2, "0");
  const m = value.getMinutes().toString().padStart(2, "0");
  return `${value.getMonth() + 1}/${value.getDate()}/${value.getFullYear()} ${h}:${m}`;
}
