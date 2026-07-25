/**
 * Zod schemas — the single source of truth for the StyleSmith data model.
 * TS types are inferred from these (see ./types.ts); nothing hand-declares the
 * data shapes separately. Runtime validation and static types stay in lockstep.
 *
 * Model reference: PRD §10. Schema rules: CLAUDE.md "Data model".
 *
 * PURE — no Office JS.
 */
import { z } from "zod";

/* ── Enumerations ────────────────────────────────────────────────────────── */

/**
 * PowerPoint theme colour slots (PRD §10.4). `themeColor` bindings reference
 * one of these; they cascade with the deck theme and survive without the add-in.
 * Colour only — theme *font* (major/minor) binding is deferred (additive later).
 */
export const THEME_SLOTS = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
] as const;
export const ThemeSlotSchema = z.enum(THEME_SLOTS);

/** Token categories (PRD §10.2). Theme slots cover colour + fonts; the rest are native. */
export const TOKEN_TYPES = [
  "color",
  "fontFamily",
  "fontSize",
  "fontWeight",
  "lineSpacing",
  "radius",
  "strokeWeight",
  "dashStyle",
  "spacing",
] as const;
export const TokenTypeSchema = z.enum(TOKEN_TYPES);

/* ── Value kinds — the three-kind union at the core of the model (PRD §10.3) ─ */

/**
 * A concrete literal value. Colours/fonts/dash styles are strings; sizes,
 * weights, radii, spacing and adjustment values are numbers.
 */
export const LiteralPrimitiveSchema = z.union([z.string(), z.number()]);

export const LiteralValueSchema = z.object({
  kind: z.literal("literal"),
  value: LiteralPrimitiveSchema,
});

/**
 * A reference to a token. `cached` is the last-resolved literal — the fallback
 * used when the token is deleted or unresolvable, so a style never blanks a
 * shape (CLAUDE.md; AC27.5). Cache lives inline on the value (decision, Phase 1).
 */
export const TokenRefValueSchema = z.object({
  kind: z.literal("tokenRef"),
  tokenId: z.string().min(1),
  cached: LiteralPrimitiveSchema.optional(),
});

/** A binding to a PowerPoint theme colour slot. Also carries an inline cache. */
export const ThemeColorValueSchema = z.object({
  kind: z.literal("themeColor"),
  slot: ThemeSlotSchema,
  cached: LiteralPrimitiveSchema.optional(),
});

export const ValueKindSchema = z.discriminatedUnion("kind", [
  LiteralValueSchema,
  TokenRefValueSchema,
  ThemeColorValueSchema,
]);

/* ── Tokens ──────────────────────────────────────────────────────────────── */

export const TokenSchema = z.object({
  /** GUID — the reference key. NEVER the name (renaming must not sever links). */
  id: z.string().min(1),
  name: z.string().min(1),
  type: TokenTypeSchema,
  /** May itself be a `tokenRef` (alias chains — PRD §10.6). */
  value: ValueKindSchema,
  description: z.string().optional(),
});

export const TokenSetSchema = z.object({
  id: z.string().min(1),
  name: z.string().min(1),
  /** Reserved: brand-library id (v2). Always present, null in v1. */
  origin: z.string().nullable(),
  tokens: z.array(TokenSchema),
});

/* ── Style layers — each independently nullable (PRD §6, §10.2) ──────────── */

export const TEXT_PROPERTIES = [
  "fontFamily",
  "fontSize",
  "fontWeight",
  "color",
  "lineSpacing",
  "spaceBefore",
  "spaceAfter",
  "alignment",
] as const;

export const SHAPE_PROPERTIES = ["fill", "borderColor", "borderWeight", "borderDashStyle"] as const;

export const TextLayerSchema = z
  .object({
    fontFamily: ValueKindSchema.optional(),
    fontSize: ValueKindSchema.optional(),
    fontWeight: ValueKindSchema.optional(),
    color: ValueKindSchema.optional(),
    lineSpacing: ValueKindSchema.optional(),
    spaceBefore: ValueKindSchema.optional(),
    spaceAfter: ValueKindSchema.optional(),
    alignment: ValueKindSchema.optional(),
  })
  .strict();

export const ShapeLayerSchema = z
  .object({
    fill: ValueKindSchema.optional(),
    borderColor: ValueKindSchema.optional(),
    borderWeight: ValueKindSchema.optional(),
    borderDashStyle: ValueKindSchema.optional(),
  })
  .strict();

/** Geometry adjustments — index-significant (index i == adjustment handle i). */
export const GeometryLayerSchema = z
  .object({
    adjustments: z.array(ValueKindSchema),
  })
  .strict();

/**
 * Table layer — deferred (Table API granularity unverified). Kept as a
 * forward-compatible opaque object so Phase 5 can populate it without a
 * migration; the layer stays nullable (PRD risk table).
 */
export const TableLayerSchema = z.record(z.string(), z.unknown());

export const LayersSchema = z.object({
  text: TextLayerSchema.nullable(),
  shape: ShapeLayerSchema.nullable(),
  geometry: GeometryLayerSchema.nullable(),
  table: TableLayerSchema.nullable(),
});

export const StyleSchema = z.object({
  /** GUID — the linkage key written to shape tags. NEVER the name. */
  id: z.string().min(1),
  name: z.string().min(1),
  /** Reserved: component library (v2). Always present, null in v1. */
  origin: z.string().nullable(),
  /** Reserved: style inheritance (v2). Always present, null in v1. */
  basedOn: z.string().nullable(),
  layers: LayersSchema,
});

/* ── Document ────────────────────────────────────────────────────────────── */

export const CURRENT_SCHEMA_VERSION = 1;

export const StyleSmithDocumentSchema = z.object({
  schemaVersion: z.literal(CURRENT_SCHEMA_VERSION),
  /** v1 has exactly one set; the array shape lets v2 add a switcher w/o migration. */
  tokenSets: z.array(TokenSetSchema),
  styles: z.array(StyleSchema),
});
