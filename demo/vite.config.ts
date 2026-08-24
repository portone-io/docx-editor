import { fileURLToPath } from "node:url";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

const packageRoot = fileURLToPath(new URL("..", import.meta.url));
const demoRoot = fileURLToPath(new URL(".", import.meta.url));

export default defineConfig({
  root: demoRoot,
  server: {
    host: "127.0.0.1",
    // The shell reads the demo fixture, and the workspace link resolves the library to its
    // sources; both sit above this root
    fs: { allow: [packageRoot] },
  },
  plugins: [react()],
});
