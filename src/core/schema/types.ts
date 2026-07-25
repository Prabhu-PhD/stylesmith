/**
 * TS types for the data model — inferred from the Zod schemas so types and
 * runtime validation never drift — plus the small non-data helper types the
 * core logic returns (Result, resolution outcomes).
 *
 * PURE — no Office JS.
 */
import type { z } from "zod";
import type {
  LiteralPrimitiveSchema,
  LiteralValueSchema,
  TokenRefValueSchema,
  ThemeColorValueSchema,
  ValueKindSchema,
  ThemeSlotSchema,
  TokenTypeSchema,
  TokenSchema,
  TokenSetSchema,
  TextLayerSchema,
  ShapeLayerSchema,
  GeometryLayerSchema,
  TableLayerSchema,
  LayersSchema,
  StyleSchema,
  StyleSmithDocumentSchema,
} from "./schemas";

/* ── Inferred data types ─────────────────────────────────────────────────── */

export type LiteralPrimitive = z.infer<typeof LiteralPrimitiveSchema>;
export type LiteralValue = z.infer<typeof LiteralValueSchema>;
export type TokenRefValue = z.infer<typeof TokenRefValueSchema>;
export type ThemeColorValue = z.infer<typeof ThemeColorValueSchema>;
export type ValueKind = z.infer<typeof ValueKindSchema>;
export type ThemeSlot = z.infer<typeof ThemeSlotSchema>;
export type TokenType = z.infer<typeof TokenTypeSchema>;
export type Token = z.infer<typeof TokenSchema>;
export type TokenSet = z.infer<typeof TokenSetSchema>;
export type TextLayer = z.infer<typeof TextLayerSchema>;
export type ShapeLayer = z.infer<typeof ShapeLayerSchema>;
export type GeometryLayer = z.infer<typeof GeometryLayerSchema>;
export type TableLayer = z.infer<typeof TableLayerSchema>;
export type Layers = z.infer<typeof LayersSchema>;
export type Style = z.infer<typeof StyleSchema>;
export type StyleSmithDocument = z.infer<typeof StyleSmithDocumentSchema>;

/** The four style layers. */
export type LayerName = "text" | "shape" | "geometry" | "table";

/* ── Result — errors are typed values in core/, not thrown (CLAUDE.md) ────── */

export type Ok<T> = { readonly ok: true; readonly value: T };
export type Err<E> = { readonly ok: false; readonly error: E };
export type Result<T, E> = Ok<T> | Err<E>;

export const ok = <T>(value: T): Ok<T> => ({ ok: true, value });
export const err = <E>(error: E): Err<E> => ({ ok: false, error });

/* ── Token resolution outcomes ───────────────────────────────────────────── */

/** Where a resolved value ultimately came from. */
export type ResolveProvenance = "literal" | "token" | "themeColor" | "cache";

/** Why a resolution failed, or (when degraded) why it fell back to the cache. */
export type ResolveFailureReason =
  | "missing-token"
  | "cycle"
  | "depth-exceeded"
  | "theme-unresolved";

export interface ResolveSuccess {
  readonly ok: true;
  readonly value: LiteralPrimitive;
  readonly provenance: ResolveProvenance;
  /** True when the concrete value came from the cached fallback (token gone/unresolvable). */
  readonly degraded: boolean;
  /** Present only when `degraded` — the reason the primary path failed. */
  readonly reason?: ResolveFailureReason;
}

export interface ResolveFailure {
  readonly ok: false;
  readonly reason: ResolveFailureReason;
  /** The offending token id, where applicable. */
  readonly tokenId?: string;
}

export type ResolveResult = ResolveSuccess | ResolveFailure;
