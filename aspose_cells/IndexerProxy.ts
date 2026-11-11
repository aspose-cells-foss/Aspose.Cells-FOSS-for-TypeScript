/**
 * Wraps an object so that bracket indexers used by the C# API keep working in TS.
 *
 * Missing properties are routed to the provided resolver:
 *   collection[0]      -> resolve(0)
 *   collection["Name"] -> resolve("Name")
 *   cells["A1"]        -> resolve("A1")
 *
 * Properties and methods that already exist on the target keep normal semantics.
 */

const ReservedNames: ReadonlySet<string> = new Set<string>([
  "then",
  "toJSON",
  "inspect",
  "$$typeof",
  "asymmetricMatch",
  "nodeType",
]);

export function createIndexerProxy<T extends object>(
  target: T,
  resolve: (key: string | number) => unknown,
  assign?: (key: string | number, value: unknown) => void,
): T {
  return new Proxy(target, {
    get(obj, property, receiver) {
      if (typeof property === "symbol") {
        return Reflect.get(obj, property, receiver);
      }

      if (property in obj) {
        return Reflect.get(obj, property, receiver);
      }

      if (ReservedNames.has(property)) {
        return undefined;
      }

      if (/^-?\d+$/.test(property)) {
        return resolve(Number(property));
      }

      return resolve(property);
    },
    set(obj, property, value, receiver) {
      if (assign !== undefined && typeof property === "string" && /^-?\d+$/.test(property)) {
        assign(Number(property), value);
        return true;
      }

      return Reflect.set(obj, property, value, receiver);
    },
  });
}
