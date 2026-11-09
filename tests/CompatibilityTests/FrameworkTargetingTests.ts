import * as fs from "fs";
import * as path from "path";
import { AssertEx, resolveRepositoryRoot } from "../Shared/TestInfrastructure";

interface PackageManifest {
  readonly scripts?: Record<string, string>;
  readonly dependencies?: Record<string, string>;
  readonly devDependencies?: Record<string, string>;
}

interface TypeScriptConfiguration {
  readonly compilerOptions?: Record<string, unknown>;
  readonly references?: unknown;
}

export function libraryTargetsNet8Only(): void {
  const root = resolveRepositoryRoot();

  AssertEx.true(fs.existsSync(path.join(root, "package.json")), "The library must declare its package manifest.");
  AssertEx.true(
    fs.existsSync(path.join(root, "aspose_cells", "api", "index.ts")),
    "The library must declare its TypeScript API facade.",
  );

  const manifest = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8")) as PackageManifest;
  AssertEx.equal("tsx index.ts", manifest.scripts?.start ?? "", "The library must expose the tsx start script.");

  const dependencies = manifest.dependencies ?? {};
  for (const dependency of ["@xmldom/xmldom", "adm-zip", "@zip.js/zip.js"]) {
    AssertEx.true(
      dependencies[dependency] !== undefined,
      "The library must declare the " + dependency + " runtime dependency.",
    );
  }

  const devDependencies = manifest.devDependencies ?? {};
  for (const dependency of ["typescript", "tsx", "@types/node", "@types/adm-zip"]) {
    AssertEx.true(
      devDependencies[dependency] !== undefined,
      "The library must declare the " + dependency + " development dependency.",
    );
  }

  const externalSdks = ["exceljs", "xlsx", "node-xlsx", "excel4node", "sheetjs"];
  for (const sdk of externalSdks) {
    AssertEx.false(
      dependencies[sdk] !== undefined || devDependencies[sdk] !== undefined,
      "The library must not depend on the external " + sdk + " SDK.",
    );
  }

  for (const dependency of Object.keys({ ...dependencies, ...devDependencies })) {
    AssertEx.false(
      dependency.toLowerCase().startsWith("aspose"),
      "The library must not depend on an external Aspose SDK package: " + dependency,
    );
  }

  const configurationFiles = fs.readdirSync(root).filter((name) => /^tsconfig.*\.json$/i.test(name));
  AssertEx.equal(1, configurationFiles.length, "The library must have one TypeScript target configuration.");

  const configuration = JSON.parse(
    fs.readFileSync(path.join(root, "tsconfig.json"), "utf8"),
  ) as TypeScriptConfiguration;
  const options = configuration.compilerOptions ?? {};

  AssertEx.equal("ES2020", options.target as string | undefined, "The library must target ES2020.");
  AssertEx.equal("esnext", options.module as string | undefined, "The library must use esnext modules.");
  AssertEx.equal(
    "bundler",
    options.moduleResolution as string | undefined,
    "The library must use bundler module resolution.",
  );
  AssertEx.equal(true, options.strict as boolean | undefined, "The library must enable TypeScript strict mode.");
  AssertEx.true(
    configuration.references === undefined,
    "The library must not restore a multi-target configuration.",
  );
}
