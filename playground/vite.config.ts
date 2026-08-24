import { fileURLToPath } from "node:url";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

const packageRoot = fileURLToPath(new URL("..", import.meta.url));
const playgroundRoot = fileURLToPath(new URL(".", import.meta.url));

export default defineConfig({
  root: playgroundRoot,
  server: {
    host: "127.0.0.1",
    fs: { allow: [packageRoot] },
  },
  plugins: [react()],
});
