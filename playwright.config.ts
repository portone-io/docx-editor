/**
 * The real browser suite, which is where an IME can be driven at all.
 *
 * It runs on the Chrome installed on the machine (`channel: "chrome"`) rather than on a browser
 * Playwright downloads, so nothing has to be fetched before `pnpm test:e2e` works. Composition is
 * driven through CDP (`Input.imeSetComposition`), which needs a real renderer and is the one thing
 * jsdom cannot stand in for.
 *
 * The harness server takes a port of its own: 5173 belongs to the apps' dev servers.
 */

import { defineConfig } from "@playwright/test";

const PORT = 5188;
const ORIGIN = `http://127.0.0.1:${PORT}`;

export default defineConfig({
  testDir: "./e2e",
  // One composition at a time: every test drives the one IME the browser has
  workers: 1,
  fullyParallel: false,
  forbidOnly: !!process.env.CI,
  reporter: "list",
  use: { baseURL: ORIGIN },
  projects: [{ name: "chromium", use: { channel: "chrome", headless: true } }],
  webServer: {
    command: `pnpm exec vite --config e2e/vite.config.ts --port ${PORT} --strictPort`,
    url: ORIGIN,
    reuseExistingServer: !process.env.CI,
    stdout: "pipe",
    stderr: "pipe",
  },
});
