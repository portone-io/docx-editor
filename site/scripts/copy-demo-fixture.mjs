import { copyFile, mkdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";

// The site fetches the fixture as a static asset, and the repository keeps one
// copy of it under __fixtures__, so it is copied in rather than committed twice.
const source = fileURLToPath(
  new URL("../../__fixtures__/demo.docx", import.meta.url)
);
const target = fileURLToPath(new URL("../public/demo.docx", import.meta.url));

await mkdir(fileURLToPath(new URL("../public", import.meta.url)), {
  recursive: true,
});
await copyFile(source, target);
