import { fileURLToPath } from "node:url";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

const packageRoot = fileURLToPath(new URL("..", import.meta.url));
const harnessRoot = fileURLToPath(new URL("./harness", import.meta.url));

export default defineConfig({
  root: harnessRoot,
  server: {
    // The address the Playwright config waits on. `localhost` resolves either way round
    host: "127.0.0.1",
    // The harness reads the package's own sources and fixtures, which sit above its root
    fs: { allow: [packageRoot] },
  },
  plugins: [react()],
});
