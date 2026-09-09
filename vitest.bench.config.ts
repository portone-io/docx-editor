import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    include: ["scripts/bench/**/*.test.ts"],
    pool: "threads",
    isolate: false,
    testTimeout: 900_000,
  },
});
