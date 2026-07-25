import { defineConfig } from "vitest/config";
import type { Plugin } from "vite";
import react from "@vitejs/plugin-react";

/**
 * Office add-ins are loaded into an iframe by the host and MUST be served over
 * HTTPS with a certificate the OS trusts — an untrusted/self-signed cert makes
 * PowerPoint silently refuse to load the task pane. office-addin-dev-certs
 * installs a trusted local CA; run `npm run dev:certs` once per machine.
 *
 * Resolved lazily in a plugin `config` hook so only `vite` (serve) needs the
 * cert — `vite build` and `vitest` never touch it.
 */
function officeHttps(): Plugin {
  return {
    name: "stylesmith:office-https",
    async config(_config, { command }) {
      // Only the dev server needs certs — never `vite build` or `vitest`.
      if (command !== "serve" || process.env.VITEST) return undefined;
      try {
        const { getHttpsServerOptions } = await import("office-addin-dev-certs");
        const { cert, key, ca } = await getHttpsServerOptions();
        return { server: { https: { cert, key, ca } } };
      } catch {
        // eslint-disable-next-line no-console
        console.warn(
          "[stylesmith] HTTPS dev certs not found. Run `npm run dev:certs` to " +
            "enable sideloading over https://localhost:3000 (serving over HTTP for now).",
        );
        return undefined;
      }
    },
  };
}

export default defineConfig(({ command }) => ({
  // GitHub Pages serves the project at /stylesmith/, so built asset URLs need
  // that prefix. Dev (serve) and vitest stay at "/" so the localhost manifest
  // and the dev server keep working unchanged.
  base: command === "build" ? "/stylesmith/" : "/",
  root: ".",
  plugins: [react(), officeHttps()],
  server: {
    host: "localhost",
    port: 3000,
    strictPort: true,
  },
  build: {
    outDir: "dist",
    sourcemap: true,
    // Rare surfaces (Tokens, adoption, debug) are React.lazy()-split at their
    // import sites, which keeps them (and the Select component they pull in) out
    // of the first-paint chunk WITHOUT defeating tree-shaking the way forced
    // manualChunks does. The remaining main chunk (~200kB gzip) is Fluent v9's
    // floor — reducing it further means swapping Fluent primitives for custom
    // ones, a tradeoff against the "native look" mandate (deferred, see PRD risk).
    chunkSizeWarningLimit: 750,
    rollupOptions: {
      input: { taskpane: "taskpane.html" },
    },
  },
  test: {
    // core/ is pure logic — no DOM needed for the unit + architecture tests.
    environment: "node",
    include: ["tests/**/*.test.ts", "src/**/*.test.ts"],
  },
}));
