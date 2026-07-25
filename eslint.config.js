import tseslint from "typescript-eslint";
import importPlugin from "eslint-plugin-import";

/**
 * The load-bearing rule (CLAUDE.md · IMPLEMENTATION-PLAN §2):
 *   core/ must NEVER import from office/ (or ui/). core/ is pure and
 *   host-independent; office/ implements interfaces core/ declares.
 *
 * Enforced structurally here from the first commit, and again by
 * tests/architecture.test.ts so a violation fails `npm test` too.
 */
export default tseslint.config(
  { ignores: ["dist/**", "node_modules/**"] },
  ...tseslint.configs.recommended,
  {
    plugins: { import: importPlugin },
    rules: {
      "import/no-restricted-paths": [
        "error",
        {
          zones: [
            {
              target: "./src/core",
              from: "./src/office",
              message:
                "core/ must not import from office/ — core is pure and host-independent.",
            },
            {
              target: "./src/core",
              from: "./src/ui",
              message: "core/ must not import from ui/ — core is UI-independent.",
            },
          ],
        },
      ],
    },
  },
);
