import { describe, it, expect } from "vitest";
import { migrateToLatest, CURRENT_SCHEMA_VERSION } from "../src/core/schema/migrations";
import { createEmptyDocument } from "../src/core/schema/defaults";
import { StyleSmithDocumentSchema } from "../src/core/schema/schemas";

describe("migrateToLatest", () => {
  it("accepts and returns a valid current-version document", () => {
    const empty = createEmptyDocument(() => "fixed-id");
    const result = migrateToLatest(empty);
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value.schemaVersion).toBe(CURRENT_SCHEMA_VERSION);
  });

  it("round-trips a hand-authored v1 document", () => {
    const raw = {
      schemaVersion: 1,
      tokenSets: [
        {
          id: "set",
          name: "Default",
          origin: null,
          tokens: [
            { id: "t1", name: "brand", type: "color", value: { kind: "themeColor", slot: "accent1" } },
          ],
        },
      ],
      styles: [
        {
          id: "s1",
          name: "Takeaway",
          origin: null,
          basedOn: null,
          layers: {
            text: { color: { kind: "literal", value: "#FFFFFF" } },
            shape: { fill: { kind: "tokenRef", tokenId: "t1", cached: "#0057B8" } },
            geometry: { adjustments: [{ kind: "literal", value: 0.15 }] },
            table: null,
          },
        },
      ],
    };
    const result = migrateToLatest(raw);
    expect(result.ok).toBe(true);
    if (result.ok) expect(StyleSmithDocumentSchema.safeParse(result.value).success).toBe(true);
  });

  it("rejects a document with no schemaVersion", () => {
    const result = migrateToLatest({ tokenSets: [], styles: [] });
    expect(result).toEqual({ ok: false, error: { code: "missing-version" } });
  });

  it("rejects a document from a newer schema version", () => {
    const result = migrateToLatest({ schemaVersion: 99, tokenSets: [], styles: [] });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe("from-newer");
      if (result.error.code === "from-newer") expect(result.error.version).toBe(99);
    }
  });

  it("reports validation issues for a structurally invalid document", () => {
    const result = migrateToLatest({
      schemaVersion: 1,
      tokenSets: [{ id: "set", name: "Default", origin: null, tokens: [{ id: "t", name: "x", type: "NOT_A_TYPE", value: { kind: "literal", value: 1 } }] }],
      styles: [],
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe("invalid");
      if (result.error.code === "invalid") expect(result.error.issues.length).toBeGreaterThan(0);
    }
  });

  it("rejects a completely non-object input without throwing", () => {
    expect(migrateToLatest(null).ok).toBe(false);
    expect(migrateToLatest("nope").ok).toBe(false);
    expect(migrateToLatest(42).ok).toBe(false);
  });
});
