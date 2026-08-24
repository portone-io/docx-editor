import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    dir: "src",
    pool: "threads",
    isolate: false,
    passWithNoTests: true,
    /**
     * Well past the default five seconds, because a handful of tests here spend that on one
     * expectation and would otherwise flake whenever the suites run alongside each other:
     * libxml2 takes seconds to build the particle set of the OOXML schemas for a document it
     * turns down (`docx/exportSchemaValidation.test.ts`), and the zip caps are proven by
     * deflating tens of megabytes (`docx/container.test.ts`, `docx/imageRoundtrip.test.ts`).
     */
    testTimeout: 30_000,
  },
});
