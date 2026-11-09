import * as fs from "fs";
import * as path from "path";

/** Walks up from the current directory until the repository root (package.json) is found. */
function findRepositoryRoot(): string {
  let directory = path.resolve(process.cwd());
  while (true) {
    if (fs.existsSync(path.join(directory, "package.json"))) {
      return directory;
    }

    const parent = path.dirname(directory);
    if (parent === directory) {
      return process.cwd();
    }

    directory = parent;
  }
}

/**
 * Returns an output path under <repo>/output/samples, creating the directory when needed.
 * Replaces the .NET samples' AppContext.BaseDirectory convention.
 */
export function sampleOutputPath(fileName: string): string {
  const directory = path.join(findRepositoryRoot(), "output", "samples");
  fs.mkdirSync(directory, { recursive: true });
  return path.join(directory, fileName);
}

/** Repository-relative path to the shared sample input folder. */
export function sampleInputPath(fileName: string): string {
  return path.join(findRepositoryRoot(), "samples", "files", fileName);
}
