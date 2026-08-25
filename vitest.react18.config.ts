import { createRequire } from "node:module";
import { defineConfig, mergeConfig } from "vitest/config";
import baseConfig from "./vitest.config.ts";

/**
 * The React 18 workspace package is the only place where a package with a `react` peer
 * resolves it to React 18 instead of the React 19 the repository develops on, so every
 * entry below is resolved from there rather than from the root installation.
 */
const requireReact18 = createRequire(
  new URL("./react18/package.json", import.meta.url)
);

/**
 * React itself, through every entry `src/` reaches directly or through the automatic JSX
 * transform, plus each runtime dependency that carries a React copy of its own. A
 * dependency left out here would hand React 18 elements React 19 built, which the React 18
 * reconciler refuses.
 */
const REACT_18_ENTRIES = [
  "react",
  "react/jsx-runtime",
  "react/jsx-dev-runtime",
  "react-dom",
  "react-dom/client",
  "lucide-react",
];

export default mergeConfig(
  baseConfig,
  defineConfig({
    resolve: {
      alias: REACT_18_ENTRIES.map((entry) => ({
        find: new RegExp(`^${entry}$`),
        replacement: requireReact18.resolve(entry),
      })),
    },
  })
);
