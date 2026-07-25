/**
 * Schema migration harness, keyed on `schemaVersion`. The document may have been
 * written by an older build (or hand-edited), so every read goes through here:
 * peek the version, step forward through registered migrations, then validate
 * against the current schema.
 *
 * v1 has no migrations yet — but the harness and its tests exist from day one so
 * v2 can add one safely (CLAUDE.md / PRD §10.7).
 *
 * PURE — no Office JS.
 */
import { z } from "zod";
import { StyleSmithDocumentSchema, CURRENT_SCHEMA_VERSION } from "./schemas";
import { ok, err, type Result, type StyleSmithDocument } from "./types";

export { CURRENT_SCHEMA_VERSION };

export type MigrationError =
  | { readonly code: "missing-version" }
  | { readonly code: "from-newer"; readonly version: number }
  | { readonly code: "no-migration"; readonly version: number }
  | { readonly code: "invalid"; readonly issues: string[] };

/** A migration transforms a vN document into a v(N+1) document. */
type Migration = (doc: unknown) => unknown;

/** Keyed by source version: MIGRATIONS[n] migrates v(n) → v(n+1). Empty in v1. */
const MIGRATIONS: Readonly<Record<number, Migration>> = {};

const VersionPeekSchema = z.object({ schemaVersion: z.number().int() });

/**
 * Validate + upgrade an unknown value to a current-version document, or return a
 * typed error. Never throws on bad input.
 */
export function migrateToLatest(raw: unknown): Result<StyleSmithDocument, MigrationError> {
  const peek = VersionPeekSchema.safeParse(raw);
  if (!peek.success) return err({ code: "missing-version" });

  let version = peek.data.schemaVersion;
  if (version > CURRENT_SCHEMA_VERSION) {
    return err({ code: "from-newer", version });
  }

  let doc: unknown = raw;
  while (version < CURRENT_SCHEMA_VERSION) {
    const migrate = MIGRATIONS[version];
    if (!migrate) return err({ code: "no-migration", version });
    doc = migrate(doc);
    version += 1;
  }

  const parsed = StyleSmithDocumentSchema.safeParse(doc);
  if (!parsed.success) {
    return err({
      code: "invalid",
      issues: parsed.error.issues.map((i) => `${i.path.join(".") || "(root)"}: ${i.message}`),
    });
  }
  return ok(parsed.data);
}
