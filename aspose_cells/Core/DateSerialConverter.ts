import { DateSystem } from "./DateSystem";

/**
 * Converts between Excel date serials and JavaScript Date values.
 *
 * Persisted serials use the Date instant so that save/load round-trips keep the
 * same Date value. Display formatting separately uses ToWallClockSerial because
 * Excel formats wall-clock components, mirroring .NET DateTime semantics.
 */
export class DateSerialConverter {
  private static readonly Windows1900EpochMs: number = Date.UTC(1899, 11, 31);
  private static readonly Mac1904EpochMs: number = Date.UTC(1904, 0, 1);
  private static readonly TicksPerDay: number = 864000000000;

  public static ToSerial(value: Date, dateSystem: DateSystem): number {
    const baseDate = dateSystem === DateSystem.Mac1904 ? DateSerialConverter.Mac1904EpochMs : DateSerialConverter.Windows1900EpochMs;
    let serial = (value.getTime() - baseDate) / 86400000;
    // Excel's 1900 date system preserves the historic leap-year bug, so serials
    // on or after the phantom 1900-02-29 are shifted by one day.
    if (dateSystem === DateSystem.Windows1900 && serial >= 60) {
      serial += 1;
    }

    return serial;
  }

  public static FromSerial(serial: number, dateSystem: DateSystem): Date {
    // Reverse the 1900-system compatibility adjustment before converting back.
    if (dateSystem === DateSystem.Windows1900 && serial >= 60) {
      serial -= 1;
    }

    const baseDate = dateSystem === DateSystem.Mac1904 ? DateSerialConverter.Mac1904EpochMs : DateSerialConverter.Windows1900EpochMs;
    const scaled = serial * DateSerialConverter.TicksPerDay;
    const ticks = scaled >= 0 ? Math.floor(scaled + 0.5) : -Math.floor(-scaled + 0.5);
    return new Date(baseDate + ticks / 10000);
  }

  /**
   * Converts a Date to an Excel serial using its local wall-clock components.
   *
   * Display formatting (for example elapsed-time tokens) works on wall-clock
   * date/time values, while the persisted serial keeps instant semantics.
   */
  public static ToWallClockSerial(value: Date, dateSystem: DateSystem): number {
    const baseDate = dateSystem === DateSystem.Mac1904 ? DateSerialConverter.Mac1904EpochMs : DateSerialConverter.Windows1900EpochMs;
    const wallClockMs = Date.UTC(
      value.getFullYear(),
      value.getMonth(),
      value.getDate(),
      value.getHours(),
      value.getMinutes(),
      value.getSeconds(),
      value.getMilliseconds(),
    );
    let serial = (wallClockMs - baseDate) / 86400000;
    if (dateSystem === DateSystem.Windows1900 && serial >= 60) {
      serial += 1;
    }

    return serial;
  }
}
