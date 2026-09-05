import { mkdir, readdir } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { Extractor, ExtractorConfig } from "@microsoft/api-extractor";

const packageRoot = fileURLToPath(new URL("..", import.meta.url));
const configDir = join(packageRoot, "api-extractor");
const localBuild = process.argv.includes("--local");

const configFilePaths = (await readdir(configDir))
  .filter((name) => name.endsWith(".json"))
  .map((name) => join(configDir, name));

/** api-extractor writes a first report only into a folder that already exists */
await mkdir(join(packageRoot, "etc"), { recursive: true });

let succeeded = true;
for (const configFilePath of configFilePaths) {
  const config = ExtractorConfig.loadFileAndPrepare(configFilePath);
  const result = Extractor.invoke(config, { localBuild });
  succeeded &&= result.succeeded;
}

if (!succeeded) process.exitCode = 1;
