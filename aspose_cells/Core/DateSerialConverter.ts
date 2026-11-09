import { DateSystem } from "./DateSystem";

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
}
