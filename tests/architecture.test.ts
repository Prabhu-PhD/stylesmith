import { describe, it, expect } from "vitest";
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";

/**
 * The load-bearing architectural boundary (CLAUDE.md · IMPLEMENTATION-PLAN §2):
 *
 *   core/ must NEVER import from office/ or ui/.
 *
 * ESLint (import/no-restricted-paths) enforces this in the editor/CI, but this
 * test makes it fail `npm test` too — so the boundary holds from commit one,
 * before any lint step is wired into a pipeline.
 */

function tsFiles(dir: string): string[] {
  const out: string[] = [];
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const p = join(dir, entry.name);
    if (entry.isDirectory()) out.push(...tsFiles(p));
    else if (/\.tsx?$/.test(entry.name)) out.push(p);
  }
  return out;
}

/** Import specifiers referenced by a source file. */
function importSpecifiers(src: string): string[] {
  const specs: string[] = [];
  const re = /(?:import|export)[^"']*?from\s*["']([^"']+)["']|import\s*\(\s*["']([^"']+)["']\s*\)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(src)) !== null) {
    const spec = m[1] ?? m[2];
    if (spec) specs.push(spec);
  }
  return specs;
}

// A specifier that reaches into a forbidden zone, e.g. "../office/x", "../../ui/y".
const FORBIDDEN = /(^|\/)(office|ui)(\/|$)/;

describe("architecture boundary", () => {
  it("core/ imports nothing from office/ or ui/", () => {
    const offenders: string[] = [];
    for (const file of tsFiles("src/core")) {
      const src = readFileSync(file, "utf8");
      for (const spec of importSpecifiers(src)) {
        if (FORBIDDEN.test(spec)) offenders.push(`${file} -> ${spec}`);
      }
    }
    expect(offenders).toEqual([]);
  });
});
