/**
 * Reproduces the culture-specific expected strings that the .NET fixtures derive
 * from CultureInfo. The helpers mirror the custom-format outputs used by the suite.
 */

function pad2(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}

/** Formats a timestamp using the invariant "M/d/yyyy H:mm" pattern. */
export function invariantMonthDayYearTime(value: Date, useUtc: boolean): string {
  const month = useUtc ? value.getUTCMonth() + 1 : value.getMonth() + 1;
  const day = useUtc ? value.getUTCDate() : value.getDate();
  const year = useUtc ? value.getUTCFullYear() : value.getFullYear();
  const hour = useUtc ? value.getUTCHours() : value.getHours();
  const minute = useUtc ? value.getUTCMinutes() : value.getMinutes();
  return `${month}/${day}/${year} ${hour}:${pad2(minute)}`;
}

/** Formats a number with the French "#,##0.00" pattern. */
export function frenchGroupedNumber(value: number): string {
  return new Intl.NumberFormat("fr-FR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value);
}

/** Formats a timestamp using the "dddd, MMMM d, yyyy" pattern for the supplied locale. */
export function weekdayMonthDayYear(value: Date, locale: string): string {
  const weekday = new Intl.DateTimeFormat(locale, { weekday: "long" }).format(value);
  const month = new Intl.DateTimeFormat(locale, { month: "long" }).format(value);
  return `${weekday}, ${month} ${value.getDate()}, ${value.getFullYear()}`;
}

/** Formats a number with a quoted yen prefix and the Japanese "#,##0.00" pattern. */
export function yenGroupedNumber(value: number): string {
  return "¥" + new Intl.NumberFormat("ja-JP", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value);
}

/** Formats a timestamp using the German long-date pattern "dddd, d. MMMM yyyy". */
export function germanLongDate(value: Date): string {
  const weekday = new Intl.DateTimeFormat("de-DE", { weekday: "long" }).format(value);
  const month = new Intl.DateTimeFormat("de-DE", { month: "long" }).format(value);
  return `${weekday}, ${value.getDate()}. ${month} ${value.getFullYear()}`;
}

/** Formats a timestamp using the German long-time pattern "HH:mm:ss". */
export function germanLongTime(value: Date): string {
  return `${pad2(value.getHours())}:${pad2(value.getMinutes())}:${pad2(value.getSeconds())}`;
}

/** Gets the German number decimal separator. */
export function germanDecimalSeparator(): string {
  const parts = new Intl.NumberFormat("de-DE").formatToParts(1.5);
  const decimal = parts.find((part) => part.type === "decimal");
  return decimal === undefined ? "," : decimal.value;
}
