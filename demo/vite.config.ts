import { fileURLToPath } from "node:url";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";
import libraryManifest from "../package.json";

const packageRootUrl = new URL("..", import.meta.url);
const packageRoot = fileURLToPath(packageRootUrl);
const demoRoot = fileURLToPath(new URL(".", import.meta.url));

/**
 * `demo/package.json` pins the published library, because the documentation site serves this
 * component and has to demonstrate the version the badge names. `pnpm dev` is the library's own
 * development loop, so here every entry point is redirected to the sources it is built from.
 *
 * The subpaths come from the library's development `exports` map, so an entry point that moves
 * stays aliased. Longer specifiers are matched first: Vite also applies an alias to everything
 * below it, so a bare `.` entry would otherwise swallow `./styles.css`.
 */
const librarySources = Object.entries(libraryManifest.exports)
  .map(([subpath, target]) => ({
    find: subpath.replace(/^\./, libraryManifest.name),
    replacement: fileURLToPath(new URL(target, packageRootUrl)),
  }))
  .sort((left, right) => right.find.length - left.find.length);

export default defineConfig({
  root: demoRoot,
  resolve: { alias: librarySources },
  server: {
    host: "127.0.0.1",
    // The shell reads the demo fixture and the aliases above resolve the library to its
    // sources; both sit above this root
    fs: { allow: [packageRoot] },
  },
  plugins: [react()],
});
