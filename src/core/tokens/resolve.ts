/**
 * Token resolution — turn a ValueKind into a concrete literal.
 *
 * Rules (PRD §10.5):
 *   1. tokenRef → token → its value (which may itself be a tokenRef: alias chain)
 *   2. cycles and over-deep chains are caught here as a runtime safety net
 *      (they are also rejected at write time — see ./cycles.ts)
 *   3. a missing/unresolvable token falls back to the value's inline `cached`
 *      literal; a style must never blank a shape (AC27.5)
 *
 * The cache fallback applies at the TOP-LEVEL value only: if any point in the
 * chain fails, we fall back to the cache of the style property that started it —
 * exactly the "cached on the style" guarantee.
 *
 * PURE — no Office JS. A theme colour is resolved via an injected resolver
 * (the office/ layer reads the live deck theme; core stays host-independent).
 */
import type {
  ResolveResult,
  ThemeSlot,
  Token,
  ValueKind,
  StyleSmithDocument,
} from "../schema/types";
import { MAX_ALIAS_DEPTH } from "../schema/defaults";

export interface ResolveContext {
  /** Token id → token, across all token sets. Build with {@link indexTokens}. */
  readonly tokens: ReadonlyMap<string, Token>;
  /** Resolve a theme slot to its current colour, or undefined if unavailable. */
  readonly theme?: (slot: ThemeSlot) => string | undefined;
  /** Override the alias-chain depth limit (defaults to MAX_ALIAS_DEPTH). */
  readonly maxDepth?: number;
}

/** Flatten every token set into a single id → token lookup. */
export function indexTokens(doc: StyleSmithDocument): Map<string, Token> {
  const map = new Map<string, Token>();
  for (const set of doc.tokenSets) {
    for (const token of set.tokens) map.set(token.id, token);
  }
  return map;
}

/** Resolve a value to a concrete literal, applying the cached fallback on failure. */
export function resolveValue(value: ValueKind, ctx: ResolveContext): ResolveResult {
  const maxDepth = ctx.maxDepth ?? MAX_ALIAS_DEPTH;
  const primary = resolveInner(value, ctx, new Set<string>(), 0, maxDepth);
  if (primary.ok) return primary;

  // Primary path failed — fall back to this value's inline cache, if any.
  if ((value.kind === "tokenRef" || value.kind === "themeColor") && value.cached !== undefined) {
    return { ok: true, value: value.cached, provenance: "cache", degraded: true, reason: primary.reason };
  }
  return primary;
}

function resolveInner(
  value: ValueKind,
  ctx: ResolveContext,
  visited: ReadonlySet<string>,
  depth: number,
  maxDepth: number,
): ResolveResult {
  if (depth > maxDepth) return { ok: false, reason: "depth-exceeded" };

  switch (value.kind) {
    case "literal":
      return { ok: true, value: value.value, provenance: "literal", degraded: false };

    case "themeColor": {
      const hex = ctx.theme?.(value.slot);
      if (hex !== undefined) return { ok: true, value: hex, provenance: "themeColor", degraded: false };
      return { ok: false, reason: "theme-unresolved" };
    }

    case "tokenRef": {
      if (visited.has(value.tokenId)) return { ok: false, reason: "cycle", tokenId: value.tokenId };
      const token = ctx.tokens.get(value.tokenId);
      if (!token) return { ok: false, reason: "missing-token", tokenId: value.tokenId };

      const next = new Set(visited);
      next.add(value.tokenId);
      const inner = resolveInner(token.value, ctx, next, depth + 1, maxDepth);
      // A tokenRef that ultimately resolves is reported as token-sourced, unless
      // it terminated at a theme colour (more specific and useful to the UI).
      if (inner.ok && inner.provenance === "literal") {
        return { ok: true, value: inner.value, provenance: "token", degraded: false };
      }
      return inner;
    }
  }
}
