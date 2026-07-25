/**
 * Constants, id/clone utilities, and the empty-document factory.
 *
 * PURE — no Office JS. `crypto.randomUUID` is a platform primitive (Node 17+,
 * all target browsers), not a host API.
 */
import type { LayerName, StyleSmithDocument, TokenType } from "./types";
import { CURRENT_SCHEMA_VERSION } from "./schemas";

/**
 * Maximum alias-chain depth (tokenRef → tokenRef → …). Chains longer than this
 * are rejected at write time and treated as unresolvable at read time. A guard,
 * not a feature — real two-tier token systems are 2–3 deep.
 */
export const MAX_ALIAS_DEPTH = 16;

/**
 * Which token `type` may bind to each style property (AC26.3), keyed by
 * "layer.property". `null` means the property is literal-only (not tokenisable),
 * e.g. text alignment is an enum. Table properties are deferred (Phase 5).
 *
 * A `themeColor` binding is additionally valid only where the type is "color".
 */
export const PROPERTY_TOKEN_TYPE: Readonly<Record<string, TokenType | null>> = {
  "text.fontFamily": "fontFamily",
  "text.fontSize": "fontSize",
  "text.fontWeight": "fontWeight",
  "text.color": "color",
  "text.lineSpacing": "lineSpacing",
  "text.spaceBefore": "spacing",
  "text.spaceAfter": "spacing",
  "text.alignment": null,
  "shape.fill": "color",
  "shape.borderColor": "color",
  "shape.borderWeight": "strokeWeight",
  "shape.borderDashStyle": "dashStyle",
  "geometry.adjustments": "radius",
};

/** The token type a property accepts, or null (literal-only) / undefined (unknown property). */
export function tokenTypeForProperty(layer: LayerName, property: string): TokenType | null | undefined {
  return PROPERTY_TOKEN_TYPE[`${layer}.${property}`];
}

/** A new GUID. Injected as a factory into model ops so tests can be deterministic. */
export function newId(): string {
  return crypto.randomUUID();
}

/** Deep clone of plain-JSON model data. Model ops are pure — they never mutate input. */
export function clone<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

/** A fresh document with a single empty "Default" token set and no styles. */
export function createEmptyDocument(makeId: () => string = newId): StyleSmithDocument {
  return {
    schemaVersion: CURRENT_SCHEMA_VERSION,
    tokenSets: [{ id: makeId(), name: "Default", origin: null, tokens: [] }],
    styles: [],
  };
}
