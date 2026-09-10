import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    include: ["scripts/bench/**/*.test.ts"],
    pool: "threads",
    isolate: false,
    fileParallelism: false,
    testTimeout: 900_000,
  },
});
